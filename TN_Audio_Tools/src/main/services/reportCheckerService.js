const fs = require('fs/promises');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const JSON5 = require('json5');

if (typeof globalThis.File === 'undefined') {
  globalThis.File = class File {};
}

const cheerio = require('cheerio');
const WordExtractor = require('word-extractor');

const SUPPORTED_REPORT_EXTENSIONS = new Set(['.doc', '.docx']);
const SUPPORTED_CHECKLIST_EXTENSIONS = new Set(['.xlsx', '.xls']);
const DEFAULT_RULES_RELATIVE_PATH = path.join(
  'src',
  'renderer',
  'modules',
  'reportChecker',
  'config',
  'moto_rules_for_analysis.json5'
);
const wordExtractor = new WordExtractor();
const LIBRE_OFFICE_CANDIDATE_PATHS = [
  'C:/Program Files/LibreOffice/program/soffice.exe',
  'C:/Program Files (x86)/LibreOffice/program/soffice.exe'
];

function runProcess(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      windowsHide: true,
      ...options
    });

    let stdout = '';
    let stderr = '';

    child.stdout?.on('data', (chunk) => {
      stdout += String(chunk);
    });

    child.stderr?.on('data', (chunk) => {
      stderr += String(chunk);
    });

    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }

      reject(new Error(stderr || stdout || `Command failed: ${command}`));
    });
  });
}

async function pathExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function queryWindowsAppPath(executableName) {
  const registryKeys = [
    `HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${executableName}`,
    `HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${executableName}`
  ];

  for (const registryKey of registryKeys) {
    try {
      const { stdout } = await runProcess('reg', ['QUERY', registryKey, '/ve']);
      const match = stdout.match(/REG_SZ\s+(.+)$/m);
      if (match && match[1]) {
        const executablePath = match[1].trim();
        if (await pathExists(executablePath)) {
          return executablePath;
        }
      }
    } catch {
      // Continue probing other registry keys.
    }
  }

  return null;
}

async function findLibreOfficeExecutable() {
  const appPath = await queryWindowsAppPath('soffice.exe');
  if (appPath) {
    return appPath;
  }

  for (const candidatePath of LIBRE_OFFICE_CANDIDATE_PATHS) {
    if (await pathExists(candidatePath)) {
      return candidatePath;
    }
  }

  return 'soffice';
}

async function convertDocWithLibreOffice(reportPath, outputDir) {
  const candidates = [await findLibreOfficeExecutable(), 'soffice.exe'];

  for (const candidate of candidates) {
    try {
      await runProcess(candidate, ['--headless', '--convert-to', 'docx', '--outdir', outputDir, reportPath]);
      const convertedPath = path.join(outputDir, `${path.parse(reportPath).name}.docx`);
      if (await pathExists(convertedPath)) {
        return convertedPath;
      }
    } catch {
      // Try the next converter candidate.
    }
  }

  return null;
}

async function convertDocWithCom(reportPath, outputDir, progId, formatCode) {
  const convertedPath = path.join(outputDir, `${path.parse(reportPath).name}.docx`);
  const escapedInput = reportPath.replace(/'/g, "''");
  const escapedOutput = convertedPath.replace(/'/g, "''");
  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$application = New-Object -ComObject '${progId}'`,
    '$application.Visible = $false',
    '$application.DisplayAlerts = 0',
    `$document = $application.Documents.Open('${escapedInput}')`,
    'if ($document.PSObject.Methods.Name -contains "SaveAs2") {',
    `  $document.SaveAs2('${escapedOutput}', ${formatCode})`,
    '} else {',
    `  $document.SaveAs('${escapedOutput}', ${formatCode})`,
    '}',
    '$document.Close()',
    '$application.Quit()'
  ].join('; ');

  try {
    await runProcess('powershell', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script]);
    return (await pathExists(convertedPath)) ? convertedPath : null;
  } catch {
    return null;
  }
}

async function convertDocWithWord(reportPath, outputDir) {
  return convertDocWithCom(reportPath, outputDir, 'Word.Application', 16);
}

async function convertDocWithWps(reportPath, outputDir) {
  const wpsProgIds = ['kwps.Application', 'wps.Application'];

  for (const progId of wpsProgIds) {
    const convertedPath = await convertDocWithCom(reportPath, outputDir, progId, 12);
    if (convertedPath) {
      return convertedPath;
    }
  }

  return null;
}

async function convertDocToTemporaryDocx(reportPath) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'tn-audio-report-'));

  try {
    const wordPath = await convertDocWithWord(reportPath, tempDir);
    if (wordPath) {
      return { tempDir, convertedPath: wordPath };
    }

    const wpsPath = await convertDocWithWps(reportPath, tempDir);
    if (wpsPath) {
      return { tempDir, convertedPath: wpsPath };
    }

    const libreOfficePath = await convertDocWithLibreOffice(reportPath, tempDir);
    if (libreOfficePath) {
      return { tempDir, convertedPath: libreOfficePath };
    }

    await fs.rm(tempDir, { recursive: true, force: true });
    return null;
  } catch (error) {
    await fs.rm(tempDir, { recursive: true, force: true });
    throw error;
  }
}

function normalizeText(value, textNormalizeConfig = {}) {
  if (typeof value !== 'string') {
    return '';
  }

  let normalized = value.replace(/\r/g, ' ').replace(/\n/g, ' ');

  if (textNormalizeConfig.removeNumberPrefix) {
    normalized = normalized.replace(/^\s*\d+(?:\.\d+)*\s*/, '');
  }

  if (textNormalizeConfig.trimSpecialChar) {
    normalized = normalized
      .replace(/\b(?:ok|not ok)\b\s*$/i, '')
      .replace(/[|]+/g, ' ')
      .replace(/[\u00a0]+/g, ' ')
      .trim();
  }

  if (textNormalizeConfig.multiSpaceToSingle) {
    normalized = normalized.replace(/\s+/g, ' ');
  }

  normalized = normalized.trim();

  if (textNormalizeConfig.caseInsensitive) {
    normalized = normalized.toLowerCase();
  }

  return normalized;
}

function pickForbiddenSuffixes(item, globalMatchConfig) {
  if (Array.isArray(item.forbiddenSuffix)) {
    return item.forbiddenSuffix;
  }

  return Array.isArray(globalMatchConfig.globalForbiddenSuffix)
    ? globalMatchConfig.globalForbiddenSuffix
    : [];
}

function hasAllKeywords(text, keywords, textNormalizeConfig) {
  if (!Array.isArray(keywords) || keywords.length === 0) {
    return true;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);
  return keywords.every((keyword) => normalizedText.includes(normalizeText(keyword, textNormalizeConfig)));
}

function hasRequiredSuffixes(text, suffixes, textNormalizeConfig) {
  if (!Array.isArray(suffixes) || suffixes.length === 0) {
    return true;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);

  return suffixes.every((suffix) => {
    const normalizedSuffix = normalizeText(suffix, textNormalizeConfig);

    if (normalizedSuffix === 'max') {
      return /(^|[^a-z0-9])max([^a-z0-9]|$)/i.test(normalizedText) && !/max-\d/i.test(normalizedText);
    }

    if (normalizedSuffix === 'min') {
      return /(^|[^a-z0-9])min([^a-z0-9]|$)/i.test(normalizedText) && !/min-\d/i.test(normalizedText);
    }

    return normalizedText.includes(normalizedSuffix);
  });
}

function hasForbiddenSuffix(text, suffixes, textNormalizeConfig) {
  if (!Array.isArray(suffixes) || suffixes.length === 0) {
    return false;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);
  return suffixes.some((suffix) => normalizedText.includes(normalizeText(suffix, textNormalizeConfig)));
}

function extractStatus(text) {
  if (!text) {
    return null;
  }

  if (/\bnot\s+ok\b/i.test(text)) {
    return 'Fail';
  }

  if (/\bok\b/i.test(text)) {
    return 'Pass';
  }

  return null;
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function extractMeasurementObject(rawText) {
  const match = rawText.match(/Measurement Object\s+([A-Za-z0-9_+\-]+)/i);
  return match ? match[1] : null;
}

function getRowNameCandidates(rowName) {
  const aliasMap = {
    crossroad: ['Crossroad', 'Crossroads'],
    tstation: ['TStation', 'Train Station', 'Train'],
    fullsizecar: ['FullsizeCar', 'Fullsize Car', 'Car'],
    callcenter: ['Callcenter', 'Call Center']
  };

  const normalized = String(rowName || '').replace(/\s+/g, '').toLowerCase();
  return aliasMap[normalized] || [rowName];
}

function extractNumberTokens(text) {
  if (!text) {
    return [];
  }

  const matches = text.match(/-?\d+(?:\.\d+)?(?:\s*(?:%|dB(?:m0|Pa)?|ms|Hz|MOS|s))?/gi);
  return matches || [];
}

function parseNumericValue(text) {
  if (!text) {
    return null;
  }

  const match = text.match(/-?\d+(?:\.\d+)?/);
  if (!match) {
    return null;
  }

  const parsed = Number.parseFloat(match[0]);
  return Number.isNaN(parsed) ? null : parsed;
}

function formatComputedNumber(value) {
  if (!Number.isFinite(value)) {
    return null;
  }

  const rounded = Math.round(value * 100) / 100;
  return String(rounded);
}

function buildTimestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');
  return [date.getFullYear(), pad(date.getMonth() + 1), pad(date.getDate())].join('') +
    '_' +
    [pad(date.getHours()), pad(date.getMinutes()), pad(date.getSeconds())].join('');
}

function makeSourcePreview(text) {
  if (!text) {
    return '';
  }

  return text.replace(/\s+/g, ' ').trim().slice(0, 160);
}

function getHeaderBasedCell(rowContext, headerPatterns, textNormalizeConfig) {
  if (!rowContext || !Array.isArray(rowContext.headers) || !Array.isArray(rowContext.cells)) {
    return null;
  }

  const normalizedPatterns = headerPatterns.map((pattern) => normalizeText(pattern, textNormalizeConfig));

  for (let index = 0; index < rowContext.headers.length; index += 1) {
    const normalizedHeader = normalizeText(rowContext.headers[index] || '', textNormalizeConfig);
    if (!normalizedHeader) {
      continue;
    }

    if (normalizedPatterns.some((pattern) => normalizedHeader.includes(pattern))) {
      return rowContext.cells[index] || null;
    }
  }

  return null;
}

function extractSummaryValue(rowContext, textNormalizeConfig) {
  if (!rowContext) {
    return null;
  }

  const directCell = getHeaderBasedCell(
    rowContext,
    ['single value', 'measured value', 'value', 'actual value', 'result'],
    textNormalizeConfig
  );

  const preferredSource = directCell || rowContext.cells.slice().reverse().find((cell) => extractNumberTokens(cell).length > 0) || rowContext.text;
  const tokens = extractNumberTokens(preferredSource);
  if (tokens.length > 0) {
    return tokens[tokens.length - 1].trim();
  }

  const status = extractStatus(directCell || rowContext.text);
  if (status) {
    return status;
  }

  return directCell || rowContext.cells[rowContext.cells.length - 1] || null;
}

function extractFormulaValue(rowContext, item, textNormalizeConfig) {
  if (!rowContext || !item.formula) {
    return null;
  }

  let rawValue = null;

  if (item.formula.targetField === 'marginValue') {
    rawValue = getHeaderBasedCell(rowContext, ['margin'], textNormalizeConfig);
  }

  if (!rawValue) {
    rawValue = rowContext.cells.slice().reverse().find((cell) => extractNumberTokens(cell).length > 0) || rowContext.text;
  }

  const numericValue = parseNumericValue(rawValue);
  if (numericValue === null) {
    return null;
  }

  return formatComputedNumber(Number(item.formula.standardValue) - numericValue);
}

function findLineWindow(lines, startIndex, size = 8) {
  if (!Array.isArray(lines) || lines.length === 0) {
    return [];
  }

  const safeIndex = Math.max(0, startIndex);
  return lines.slice(safeIndex, Math.min(lines.length, safeIndex + size));
}

function resolveAnchorValue(reportData, item, textNormalizeConfig) {
  const anchorText = item.anchorConfig?.anchorText;
  if (!anchorText) {
    return { matched: false, reason: '缺少 anchorText 配置' };
  }

  const normalizedAnchor = normalizeText(anchorText, textNormalizeConfig);
  const lineIndex = reportData.lines.findIndex((line) => normalizeText(line, textNormalizeConfig).includes(normalizedAnchor));
  if (lineIndex === -1) {
    return { matched: false, reason: '未找到锚点文本' };
  }

  const windowLines = findLineWindow(reportData.lines, lineIndex, 10);
  const joinedWindow = windowLines.join(' ');
  const dbAtMatch = joinedWindow.match(/(-?\d+(?:\.\d+)?)\s*dB\s+at/i);
  let value = dbAtMatch ? dbAtMatch[1] : null;

  if (!value) {
    const numericToken = extractNumberTokens(joinedWindow)[0];
    value = numericToken ? numericToken.match(/-?\d+(?:\.\d+)?/)?.[0] : null;
  }

  if (!value) {
    return { matched: false, reason: '锚点附近未找到数值' };
  }

  const numericValue = Number.parseFloat(value);
  const range = item.anchorConfig?.valueRange;
  if (Array.isArray(range) && range.length === 2) {
    if (numericValue < range[0] || numericValue > range[1]) {
      return { matched: false, reason: '锚点数值超出配置范围' };
    }
  }

  return {
    matched: true,
    value,
    sourcePreview: makeSourcePreview(joinedWindow),
    sourceType: 'anchor-window'
  };
}

function resolveRegexValue(reportData, item, textNormalizeConfig) {
  if (!item.regexConfig?.matchRegex) {
    return { matched: false, reason: '缺少正则配置' };
  }

  const regex = new RegExp(item.regexConfig.matchRegex, 'i');

  for (let lineIndex = 0; lineIndex < reportData.lines.length; lineIndex += 1) {
    const contextWindow = findLineWindow(reportData.lines, Math.max(0, lineIndex - 2), 12);
    const joinedWindow = contextWindow.join('\n');

    if (!hasAllKeywords(joinedWindow, item.coreKeywords, textNormalizeConfig)) {
      continue;
    }

    for (const candidateLine of contextWindow) {
      const match = candidateLine.match(regex);
      if (!match) {
        continue;
      }

      const groupIndex = Number(item.regexConfig.valueGroupIndex || 1);
      const value = match[groupIndex] || match[0];
      const numericValue = parseNumericValue(value);
      const range = item.regexConfig.valueRange;

      if (Array.isArray(range) && range.length === 2 && numericValue !== null) {
        if (numericValue < range[0] || numericValue > range[1]) {
          continue;
        }
      }

      return {
        matched: true,
        value: candidateLine.includes('%') && numericValue !== null ? String(numericValue / 100) : value,
        sourcePreview: makeSourcePreview(candidateLine),
        sourceType: 'regex-line'
      };
    }
  }

  return { matched: false, reason: '未匹配到正则目标文本' };
}

function resolveTableValue(reportData, item, textNormalizeConfig) {
  const tableConfig = item.tableConfig;
  if (!tableConfig) {
    return { matched: false, reason: '缺少 tableConfig 配置' };
  }

  const requiredHeaders = Array.isArray(tableConfig.tableHeaderMatch) ? tableConfig.tableHeaderMatch : [];
  const rowNameCandidates = getRowNameCandidates(tableConfig.rowNameMatch);

  for (const table of reportData.tables) {
    const headerRow = table.rows[0] || [];
    const normalizedHeaders = headerRow.map((header) => normalizeText(header, textNormalizeConfig));
    const hasHeaders = requiredHeaders.every((header) => {
      const normalizedHeader = normalizeText(header, textNormalizeConfig);
      return normalizedHeaders.some((cell) => cell.includes(normalizedHeader));
    });

    if (!hasHeaders) {
      continue;
    }

    const targetColumnIndex = headerRow.findIndex((header) => {
      return normalizeText(header, textNormalizeConfig).includes(normalizeText(tableConfig.targetColumnName, textNormalizeConfig));
    });

    if (targetColumnIndex === -1) {
      continue;
    }

    const targetRow = table.rows.slice(1).find((row) => {
      return row.some((cell) => rowNameCandidates.some((candidate) => normalizeText(cell, textNormalizeConfig).includes(normalizeText(candidate, textNormalizeConfig))));
    });

    if (!targetRow || !targetRow[targetColumnIndex]) {
      continue;
    }

    return {
      matched: true,
      value: targetRow[targetColumnIndex],
      sourcePreview: makeSourcePreview(targetRow.join(' | ')),
      sourceType: 'table-cell'
    };
  }

  for (const rowContext of reportData.tableRows) {
    const normalizedText = normalizeText(rowContext.text, textNormalizeConfig);

    if (!rowNameCandidates.some((candidate) => normalizedText.includes(normalizeText(candidate, textNormalizeConfig)))) {
      continue;
    }

    if (!normalizedText.includes(normalizeText(tableConfig.targetColumnName, textNormalizeConfig))) {
      continue;
    }

    const tokens = extractNumberTokens(rowContext.text);
    if (tokens.length === 0) {
      continue;
    }

    return {
      matched: true,
      value: tokens[tokens.length - 1].trim(),
      sourcePreview: makeSourcePreview(rowContext.text),
      sourceType: 'text-row-table-fallback'
    };
  }

  return { matched: false, reason: '未找到匹配表格或目标单元格' };
}

function selectCandidateRow(reportData, item, textNormalizeConfig, globalMatchConfig) {
  const forbiddenSuffixes = pickForbiddenSuffixes(item, globalMatchConfig);
  const candidates = reportData.tableRows.filter((rowContext) => {
    if (!hasAllKeywords(rowContext.text, item.coreKeywords, textNormalizeConfig)) {
      return false;
    }

    if (!hasRequiredSuffixes(rowContext.text, item.requiredSuffix, textNormalizeConfig)) {
      return false;
    }

    if (hasForbiddenSuffix(rowContext.text, forbiddenSuffixes, textNormalizeConfig)) {
      return false;
    }

    return true;
  });

  if (candidates.length === 0) {
    return null;
  }

  return candidates.sort((left, right) => right.text.length - left.text.length)[0];
}

function resolveRowBasedValue(reportData, item, textNormalizeConfig, globalMatchConfig) {
  const rowContext = selectCandidateRow(reportData, item, textNormalizeConfig, globalMatchConfig);
  if (!rowContext) {
    return { matched: false, reason: '未找到匹配的表格行' };
  }

  if (item.extractType === 'summary_table_match') {
    const value = extractSummaryValue(rowContext, textNormalizeConfig);
    return value
      ? { matched: true, value, sourcePreview: makeSourcePreview(rowContext.text), sourceType: 'table-row' }
      : { matched: false, reason: '匹配到行但未提取到值' };
  }

  if (item.extractType === 'formula_calc') {
    const value = extractFormulaValue(rowContext, item, textNormalizeConfig);
    return value
      ? { matched: true, value, sourcePreview: makeSourcePreview(rowContext.text), sourceType: 'table-row-formula' }
      : { matched: false, reason: '匹配到行但公式计算失败' };
  }

  if (item.extractType === 'status_judge') {
    const value = extractStatus(rowContext.text);
    return value
      ? { matched: true, value, sourcePreview: makeSourcePreview(rowContext.text), sourceType: 'table-row-status' }
      : { matched: false, reason: '匹配到行但未找到状态值' };
  }

  return { matched: false, reason: '未知的行匹配提取类型' };
}

function createSearchData(rawText, html) {
  const lines = rawText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const flattenedText = rawText
    .replace(/\r/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const measurementObject = extractMeasurementObject(rawText);
  let derivedRows = [];

  if (measurementObject) {
    const rowPattern = new RegExp(`(9\\.\\d+(?:\\.\\d+)*[\\s\\S]*?${escapeRegExp(measurementObject)})`, 'g');
    derivedRows = Array.from(flattenedText.matchAll(rowPattern)).map((match, index) => ({
      tableIndex: -1,
      rowIndex: index + 1,
      headers: [],
      cells: [],
      text: match[1].replace(new RegExp(`${escapeRegExp(measurementObject)}$`), '').trim()
    }));
  }

  if (derivedRows.length === 0) {
    const rowStartMatches = Array.from(flattenedText.matchAll(/\b9\.\d+(?:\.\d+)*\b/g));
    derivedRows = rowStartMatches
      .map((match, index) => {
        const start = match.index;
        const end = index + 1 < rowStartMatches.length ? rowStartMatches[index + 1].index : flattenedText.length;
        const text = flattenedText.slice(start, end).trim();

        return {
          tableIndex: -1,
          rowIndex: index + 1,
          headers: [],
          cells: [],
          text
        };
      })
      .filter((row) => row.text.length > 8);
  }

  const $ = cheerio.load(html);
  const tables = $('table')
    .toArray()
    .map((tableElement, tableIndex) => {
      const rows = $(tableElement)
        .find('tr')
        .toArray()
        .map((rowElement) => {
          return $(rowElement)
            .find('th, td')
            .toArray()
            .map((cellElement) => $(cellElement).text().replace(/\s+/g, ' ').trim())
            .filter(Boolean);
        })
        .filter((row) => row.length > 0);

      return { tableIndex, rows };
    })
    .filter((table) => table.rows.length > 0);

  const tableRows = tables.flatMap((table) => {
    const headers = table.rows[0] || [];
    return table.rows.slice(1).map((cells, rowOffset) => ({
      tableIndex: table.tableIndex,
      rowIndex: rowOffset + 1,
      headers,
      cells,
      text: cells.join(' | ')
    }));
  });

  return {
    rawText,
    html,
    lines,
    tables,
    tableRows: [...tableRows, ...derivedRows]
  };
}

async function loadRules(rulePath) {
  const content = await fs.readFile(rulePath, 'utf8');
  const rules = JSON5.parse(content);

  if (!Array.isArray(rules.extractItemList)) {
    throw new Error('规则文件缺少 extractItemList 配置');
  }

  return rules;
}

async function parseDocxReport(reportPath) {
  const [rawTextResult, htmlResult] = await Promise.all([
    mammoth.extractRawText({ path: reportPath }),
    mammoth.convertToHtml({ path: reportPath })
  ]);

  return createSearchData(rawTextResult.value || '', htmlResult.value || '');
}

async function parseReport(reportPath) {
  const reportExtension = path.extname(reportPath).toLowerCase();
  if (!SUPPORTED_REPORT_EXTENSIONS.has(reportExtension)) {
    throw new Error('当前仅支持 .doc 或 .docx 测试报告');
  }

  if (reportExtension === '.doc') {
    const converted = await convertDocToTemporaryDocx(reportPath);

    if (converted?.convertedPath) {
      try {
        return await parseDocxReport(converted.convertedPath);
      } finally {
        await fs.rm(converted.tempDir, { recursive: true, force: true });
      }
    }

    const extracted = await wordExtractor.extract(reportPath);
    const rawText = [
      extracted.getHeaders?.() || '',
      extracted.getBody?.() || '',
      extracted.getFootnotes?.() || '',
      extracted.getEndnotes?.() || '',
      extracted.getTextboxes?.() || ''
    ].filter(Boolean).join('\n');

    return createSearchData(rawText, '');
  }

  return parseDocxReport(reportPath);
}

function applyResultsToChecklist(checklistPath, reportPath, extractedItems) {
  const workbook = XLSX.readFile(checklistPath, { cellStyles: true, cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  extractedItems.forEach((itemResult) => {
    if (!itemResult.matched || !itemResult.value) {
      return;
    }

    XLSX.utils.sheet_add_aoa(worksheet, [[itemResult.value]], { origin: itemResult.outputCell });
  });

  const checklistName = path.parse(checklistPath).name;
  const reportName = path.parse(reportPath).name;
  const outputPath = path.join(path.dirname(reportPath), `${checklistName}_${reportName}_${buildTimestamp()}.xlsx`);

  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx' });
  return outputPath;
}

async function processSingleReport({ reportPath, checklistPath, rules }) {
  const reportData = await parseReport(reportPath);
  const textNormalizeConfig = rules.globalMatchConfig?.textNormalize || {};
  const globalMatchConfig = rules.globalMatchConfig || {};

  const extractedItems = rules.extractItemList.map((item) => {
    let extractionResult;

    if (['summary_table_match', 'formula_calc', 'status_judge'].includes(item.extractType)) {
      extractionResult = resolveRowBasedValue(reportData, item, textNormalizeConfig, globalMatchConfig);
    } else if (item.extractType === 'detail_anchor_extract') {
      extractionResult = resolveAnchorValue(reportData, item, textNormalizeConfig);
    } else if (item.extractType === 'table_struct_extract') {
      extractionResult = resolveTableValue(reportData, item, textNormalizeConfig);
    } else if (item.extractType === 'regex_extract') {
      extractionResult = resolveRegexValue(reportData, item, textNormalizeConfig);
    } else {
      extractionResult = { matched: false, reason: `暂不支持的提取类型: ${item.extractType}` };
    }

    return {
      itemId: item.itemId,
      checklistDesc: item.checklistDesc,
      outputCell: item.outputCell,
      extractType: item.extractType,
      ...extractionResult
    };
  });

  const outputPath = applyResultsToChecklist(checklistPath, reportPath, extractedItems);
  const matchedItems = extractedItems.filter((item) => item.matched).length;
  const unmatchedItems = extractedItems.filter((item) => !item.matched).map((item) => ({
    itemId: item.itemId,
    checklistDesc: item.checklistDesc,
    outputCell: item.outputCell,
    reason: item.reason || '未命中'
  }));

  return {
    reportPath,
    outputPath,
    totalItems: extractedItems.length,
    matchedItems,
    unmatchedItems,
    extractedItems
  };
}

async function validatePaths({ reportPaths, checklistPath, rulePath }) {
  if (!Array.isArray(reportPaths) || reportPaths.length === 0) {
    throw new Error('请先选择至少一个测试报告');
  }

  const checklistExtension = path.extname(checklistPath || '').toLowerCase();
  if (!SUPPORTED_CHECKLIST_EXTENSIONS.has(checklistExtension)) {
    throw new Error('checklist 仅支持 .xlsx 或 .xls 文件');
  }

  await fs.access(checklistPath);
  await Promise.all(reportPaths.map((reportPath) => fs.access(reportPath)));

  if (rulePath) {
    await fs.access(rulePath);
  }
}

function resolveRulePath(appPath, customRulePath) {
  if (customRulePath) {
    return customRulePath;
  }

  return path.join(appPath, DEFAULT_RULES_RELATIVE_PATH);
}

async function processReports({ reportPaths, checklistPath, rulePath, appPath }) {
  const resolvedRulePath = resolveRulePath(appPath, rulePath);
  await validatePaths({ reportPaths, checklistPath, rulePath: resolvedRulePath });

  const rules = await loadRules(resolvedRulePath);
  const results = [];

  for (const reportPath of reportPaths) {
    try {
      const result = await processSingleReport({ reportPath, checklistPath, rules });
      results.push({ status: 'success', ...result });
    } catch (error) {
      results.push({
        status: 'error',
        reportPath,
        error: error.message || '报告处理失败'
      });
    }
  }

  return {
    rulePath: resolvedRulePath,
    results
  };
}

module.exports = {
  processReports,
  DEFAULT_RULES_RELATIVE_PATH
};