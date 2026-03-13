const fs = require('fs/promises');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const mammoth = require('mammoth');
const JSZip = require('jszip');
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
const CHILD_PROCESS_TIMEOUT_MS = 90000;
const PERCENT_OUTPUT_CELLS = new Set(['I41', 'I42', 'I43', 'I44', 'I45', 'I46']);
const CHECKLIST_BORDER_RANGE = {
  startCol: 'A',
  endCol: 'K',
  startRow: 3,
  endRow: 75
};

function runProcess(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const { timeoutMs = 0, ...spawnOptions } = options;
    const child = spawn(command, args, {
      windowsHide: true,
      ...spawnOptions
    });

    let stdout = '';
    let stderr = '';
    let finished = false;

    const timeoutHandle = timeoutMs > 0
      ? setTimeout(() => {
          if (finished) {
            return;
          }

          finished = true;
          child.kill();
          reject(new Error(`Command timed out after ${timeoutMs}ms: ${command}`));
        }, timeoutMs)
      : null;

    const finalize = (handler) => {
      if (finished) {
        return;
      }

      finished = true;
      if (timeoutHandle) {
        clearTimeout(timeoutHandle);
      }
      handler();
    };

    child.stdout?.on('data', (chunk) => {
      stdout += String(chunk);
    });

    child.stderr?.on('data', (chunk) => {
      stderr += String(chunk);
    });

    child.on('error', (error) => {
      finalize(() => reject(error));
    });

    child.on('close', (code) => {
      finalize(() => {
        if (code === 0) {
          resolve({ stdout, stderr });
          return;
        }

        reject(new Error(stderr || stdout || `Command failed: ${command}`));
      });
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
      await runProcess(candidate, ['--headless', '--convert-to', 'docx', '--outdir', outputDir, reportPath], {
        timeoutMs: CHILD_PROCESS_TIMEOUT_MS
      });
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
    await runProcess('powershell', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script], {
      timeoutMs: CHILD_PROCESS_TIMEOUT_MS
    });
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

  normalized = normalized
    .replace(/\bha(?:nb|wb|sb|swb)\b/gi, ' HABAND ')
    .replace(/\bmxx-(\d+)/gi, ' MAX-$1 ')
    .replace(/\breceiv(?:e|ing|eing)\b/gi, ' RCV ')
    .replace(/\bsending\b/gi, ' SND ')
    .replace(/\bfrequency\b/gi, ' FREQ ')
    .replace(/\bcharact(?:er)?\.?\b/gi, ' CHAR ')
    .replace(/\baverage\b/gi, ' AVG ')
    .replace(/\bnominal\b/gi, ' NOM ')
    .replace(/\bvolume\b/gi, ' VOL ')
    .replace(/\bmandatory\b/gi, ' MAND ');

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

function itemMatchesKeywords(reportData, item, text, textNormalizeConfig, allowFallback = false) {
  if (hasAllKeywords(text, item.coreKeywords, textNormalizeConfig)) {
    return true;
  }

  const isLegacyTextOnlyReport = !Array.isArray(reportData?.tables) || reportData.tables.length === 0;
  if (!isLegacyTextOnlyReport && !allowFallback) {
    return false;
  }

  return hasAllKeywords(text, item.fallbackKeywords, textNormalizeConfig);
}

function getRowDescriptorText(rowContext) {
  if (rowContext && Array.isArray(rowContext.cells) && rowContext.cells.length > 0) {
    if (rowContext.cells.length === 1) {
      const singleCellText = rowContext.cells[0] || rowContext.text || '';
      const statusSplit = singleCellText.split(/\b(?:not\s+ok|ok|done)\b/i)[0]?.trim();
      return statusSplit || singleCellText;
    }

    return rowContext.cells[0] || rowContext.text || '';
  }

  const rawText = String(rowContext?.text || '');
  const pipeSplitText = rawText.split('|')[0].trim();
  const statusSplitText = pipeSplitText.split(/\b(?:not\s+ok|ok|done)\b/i)[0]?.trim();
  return statusSplitText || pipeSplitText;
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
      return /(^|[^a-z0-9])min([^a-z0-9]|$)/i.test(normalizedText)
        && !/min-\d/i.test(normalizedText)
        && !/min\.?\s+dist/i.test(normalizedText)
        && !/minimal/i.test(normalizedText);
    }

    return normalizedText.includes(normalizedSuffix);
  });
}

function hasExactRowKeywords(text, exactKeywords, textNormalizeConfig) {
  if (!Array.isArray(exactKeywords) || exactKeywords.length === 0) {
    return true;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);
  return exactKeywords.every((keyword) => normalizedText.includes(normalizeText(keyword, textNormalizeConfig)));
}

function hasForbiddenSuffix(text, suffixes, textNormalizeConfig) {
  if (!Array.isArray(suffixes) || suffixes.length === 0) {
    return false;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);
  return suffixes.some((suffix) => normalizedText.includes(normalizeText(suffix, textNormalizeConfig)));
}

function getCandidateScore(rowContext) {
  const structuredScore = Array.isArray(rowContext.cells) && rowContext.cells.length > 0 ? 1000 : 0;
  const numericCellScore = Array.isArray(rowContext.cells)
    ? rowContext.cells.filter((cell) => !isMeasurementObjectCell(cell) && extractNumberTokens(cell).length > 0).length * 100
    : 0;
  const headerScore = Array.isArray(rowContext.headers) && rowContext.headers.length > 0 ? 50 : 0;
  const statusScore = extractStatus(rowContext.text) ? 500 : 0;
  return structuredScore + numericCellScore + headerScore + statusScore + rowContext.text.length;
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

function getAmbientNoiseSceneCandidates(rowName) {
  const aliasMap = {
    pub: ['Pub', 'Pub Noise'],
    road: ['Road', 'Outside Traffic Road'],
    crossroad: ['Crossroad', 'Crossroads', 'Outside Traffic Crossroads'],
    tstation: ['Train Station', 'TStation', 'Train'],
    fullsizecar: ['Fullsize Car 130 km/h', 'Fullsize Car', 'FullsizeCar', 'Car'],
    cafeteria: ['Cafeteria', 'Cafeteria Noise'],
    mensa: ['Mensa'],
    callcenter: ['Callcenter', 'Call Center', 'Work Noise Office Callcenter']
  };

  const normalized = String(rowName || '').replace(/\s+/g, '').toLowerCase();
  return aliasMap[normalized] || getRowNameCandidates(rowName);
}

function sanitizeAmbientNoiseSceneLabel(value) {
  return String(value || '')
    .replace(/\t+/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\bDone\b.*$/i, '')
    .trim();
}

function getAmbientNoiseSceneKey(value) {
  const normalized = String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');

  if (!normalized) {
    return '';
  }

  if (normalized.includes('crossroad')) {
    return 'crossroad';
  }

  if (normalized.includes('trainstation') || normalized === 'tstation' || normalized === 'train') {
    return 'tstation';
  }

  if (normalized.includes('fullsizecar') || normalized === 'car') {
    return 'fullsizecar';
  }

  if (normalized.includes('callcenter')) {
    return 'callcenter';
  }

  if (normalized.includes('cafeteria')) {
    return 'cafeteria';
  }

  if (normalized.includes('mensa')) {
    return 'mensa';
  }

  if (normalized.includes('pub')) {
    return 'pub';
  }

  if (normalized === 'road' || normalized.includes('outsidetrafficroad')) {
    return 'road';
  }

  return normalized;
}

function isMeasurementObjectCell(text) {
  const normalized = String(text || '').trim();
  return /^[A-Za-z0-9]+(?:_[A-Za-z0-9+ --]+)+$/.test(normalized);
}

function isStatusCell(text) {
  const normalized = normalizeText(text, {
    removeNumberPrefix: true,
    multiSpaceToSingle: true,
    caseInsensitive: true,
    trimSpecialChar: true
  });

  return ['ok', 'not ok', 'done', 'pass', 'fail'].includes(normalized);
}

function isLikelyReportMetadataCell(text) {
  const normalized = String(text || '').trim();
  if (!normalized) {
    return false;
  }

  if (/^[A-Za-z0-9]+(?:_[A-Za-z0-9+-]+){2,}$/.test(normalized)) {
    return true;
  }

  return /\.(?:doc|docx|xls|xlsx)$/i.test(normalized);
}

function getLastMeaningfulNumericToken(text) {
  const cleanedText = String(text || '')
    .replace(/^[0-9.]+\s+/, '')
    .replace(/[A-Za-z0-9]+(?:_[A-Za-z0-9+-]+)+$/g, '')
    .trim();

  const tokens = extractNumberTokens(cleanedText);
  return tokens.length > 0 ? tokens[tokens.length - 1].trim() : null;
}

function extractRowNumericCell(rowContext) {
  if (!rowContext || !Array.isArray(rowContext.cells) || rowContext.cells.length === 0) {
    return null;
  }

  const numericCell = rowContext.cells
    .slice()
    .reverse()
    .find((cell) => {
      if (!cell || isMeasurementObjectCell(cell) || isStatusCell(cell) || isLikelyReportMetadataCell(cell)) {
        return false;
      }

      return extractNumberTokens(cell).length > 0;
    });

  return numericCell ? getLastMeaningfulNumericToken(numericCell) : null;
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

function extractAmbientNoiseAverageBlocks(rawText) {
  if (!rawText) {
    return [];
  }

  const lines = rawText.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const blockMap = new Map();
  let currentSceneKey = null;

  for (const line of lines) {
    const sceneMatch = /(?:7|9)\.12\.1\s+Quality in ambient noise,\s*([^\t\n\r]+?)(?:\t|\s{2,}|$)/i.exec(line);
    if (sceneMatch) {
      const sceneLabel = sanitizeAmbientNoiseSceneLabel(sceneMatch[1] || '');
      const sceneKey = getAmbientNoiseSceneKey(sceneLabel);

      if (sceneKey) {
        currentSceneKey = sceneKey;
        if (!blockMap.has(sceneKey)) {
          blockMap.set(sceneKey, {
            sceneLabel,
            gmos: null,
            nmos: null,
            smos: null,
            sourcePreview: ''
          });
        }
      }
    }

    if (!currentSceneKey || !blockMap.has(currentSceneKey)) {
      continue;
    }

    const block = blockMap.get(currentSceneKey);
    const metricPatterns = [
      ['gmos', /G-MOS \(Average(?:,\s*TS\s*103\s*(?:106|281)(?:\s*\(Model A\))?)?\)[:\t ]+([0-9.]+)/i],
      ['nmos', /^(?!Corrected\b).*N-MOS \(Average(?:,\s*TS\s*103\s*(?:106|281)(?:\s*\(Model A\))?)?\)[:\t ]+([0-9.]+)/i],
      ['smos', /S-MOS \(Average(?:,\s*TS\s*103\s*(?:106|281)(?:\s*\(Model A\))?)?\)[:\t ]+([0-9.]+)/i]
    ];

    metricPatterns.forEach(([metricKey, pattern]) => {
      const value = pattern.exec(line)?.[1] || null;
      if (value) {
        block[metricKey] = value;
        block.sourcePreview = makeSourcePreview(`${block.sourcePreview} ${line}`);
      }
    });
  }

  return Array.from(blockMap.values()).filter((block) => block.gmos || block.nmos || block.smos);
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

  const numericCellValue = extractRowNumericCell(rowContext);
  if (numericCellValue) {
    return numericCellValue;
  }

  const directCellValue = getLastMeaningfulNumericToken(directCell);
  if (directCellValue) {
    return directCellValue;
  }

  const fallbackValue = getLastMeaningfulNumericToken(rowContext.text);
  if (fallbackValue) {
    return fallbackValue;
  }

  const status = extractStatus(directCell || rowContext.text);
  if (status) {
    return status;
  }

  return directCell || rowContext.cells[rowContext.cells.length - 1] || null;
}

function extractFormulaValue(reportData, rowContext, item, textNormalizeConfig, extractedResultsByItemId) {
  if (!rowContext || !item.formula) {
    return null;
  }

  let rawValue = null;

  if (item.formula.targetField === 'marginValue') {
    rawValue = getHeaderBasedCell(rowContext, ['margin'], textNormalizeConfig);
  }

  if (!rawValue) {
    rawValue = extractRowNumericCell(rowContext) || getLastMeaningfulNumericToken(rowContext.text);
  }

  const numericValue = parseNumericValue(rawValue);
  if (numericValue === null) {
    return null;
  }

  const shouldUseBaseItemValue = rowContext?.sourceKind && rowContext.sourceKind !== 'html-table';

  if (shouldUseBaseItemValue && item.formula.baseItemId && extractedResultsByItemId instanceof Map) {
    const baseResult = extractedResultsByItemId.get(Number(item.formula.baseItemId));
    const baseNumericValue = parseNumericValue(baseResult?.value);

    if (baseNumericValue !== null) {
      return formatComputedNumber(baseNumericValue - numericValue);
    }
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

function resolveAnchorValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId) {
  const anchorText = item.anchorConfig?.anchorText;
  if (!anchorText) {
    return { matched: false, reason: '缺少 anchorText 配置' };
  }

  const normalizedAnchor = normalizeText(anchorText, textNormalizeConfig);
  const anchorIndexes = reportData.lines
    .map((line, index) => ({ line, index }))
    .filter(({ line }) => normalizeText(line, textNormalizeConfig).includes(normalizedAnchor));

  if (anchorIndexes.length === 0) {
    return { matched: false, reason: '未找到锚点文本' };
  }

  const rankedAnchors = anchorIndexes
    .map(({ index }) => {
      const windowLines = findLineWindow(reportData.lines, Math.max(0, index - 8), 18);
      const joinedWindow = windowLines.join(' ');
      const activeKeywords = itemMatchesKeywords(reportData, item, joinedWindow, textNormalizeConfig, true)
        ? (hasAllKeywords(joinedWindow, item.coreKeywords, textNormalizeConfig) ? item.coreKeywords : item.fallbackKeywords)
        : [];
      const keywordScore = Array.isArray(activeKeywords)
        ? activeKeywords.filter((keyword) => normalizeText(joinedWindow, textNormalizeConfig).includes(normalizeText(keyword, textNormalizeConfig))).length
        : 0;

      return { index, windowLines, joinedWindow, keywordScore };
    })
    .sort((left, right) => right.keywordScore - left.keywordScore);

  for (const candidate of rankedAnchors) {
    if (!itemMatchesKeywords(reportData, item, candidate.joinedWindow, textNormalizeConfig, true)) {
      continue;
    }

    const dbAtMatch = candidate.joinedWindow.match(/(-?\d+(?:\.\d+)?)\s*dB\s+at/i);
    let value = dbAtMatch ? dbAtMatch[1] : null;

    if (!value) {
      const numericToken = getLastMeaningfulNumericToken(candidate.joinedWindow);
      value = numericToken ? numericToken.match(/-?\d+(?:\.\d+)?/)?.[0] : null;
    }

    if (!value) {
      continue;
    }

    const numericValue = Number.parseFloat(value);
    const range = item.anchorConfig?.valueRange;
    if (Array.isArray(range) && range.length === 2) {
      if (numericValue < range[0] || numericValue > range[1]) {
        continue;
      }
    }

    return {
      matched: true,
      value,
      sourcePreview: makeSourcePreview(candidate.joinedWindow),
      sourceType: 'anchor-window'
    };
  }

  const fallbackRow = resolveRowBasedValue(reportData, {
    ...item,
    extractType: 'summary_table_match'
  }, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);

  if (fallbackRow?.matched) {
    return {
      ...fallbackRow,
      sourceType: 'anchor-fallback-row'
    };
  }

  return { matched: false, reason: '锚点附近未找到符合规则的数值' };
}

function resolveRegexValue(reportData, item, textNormalizeConfig, globalMatchConfig) {
  if (!item.regexConfig?.matchRegex) {
    return { matched: false, reason: '缺少正则配置' };
  }

  const regex = new RegExp(item.regexConfig.matchRegex, 'i');
  const forbiddenSuffixes = pickForbiddenSuffixes(item, globalMatchConfig);
  const range = item.regexConfig.valueRange;
  const candidates = [];

  for (const rowContext of reportData.tableRows) {
    if (!itemMatchesKeywords(reportData, item, rowContext.text, textNormalizeConfig, rowContext.sourceKind !== 'html-table')) {
      continue;
    }

    const descriptorText = getRowDescriptorText(rowContext);
    if (!hasRequiredSuffixes(descriptorText, item.requiredSuffix, textNormalizeConfig)) {
      continue;
    }

    if (hasForbiddenSuffix(descriptorText, forbiddenSuffixes, textNormalizeConfig)) {
      continue;
    }

    if (!hasExactRowKeywords(rowContext.text, item.exactRowKeywords, textNormalizeConfig)) {
      continue;
    }

    const match = rowContext.text.match(regex);
    if (!match) {
      continue;
    }

    const groupIndex = Number(item.regexConfig.valueGroupIndex || 1);
    const value = match[groupIndex] || match[0];
    const numericValue = parseNumericValue(value);

    if (Array.isArray(range) && range.length === 2 && numericValue !== null) {
      if (numericValue < range[0] || numericValue > range[1]) {
        continue;
      }
    }

    candidates.push({
      matched: true,
      value: rowContext.text.includes('%') && numericValue !== null ? String(numericValue / 100) : value,
      sourcePreview: makeSourcePreview(rowContext.text),
      sourceType: 'regex-row',
      score: 2000 + getCandidateScore(rowContext)
    });
  }

  for (let lineIndex = 0; lineIndex < reportData.lines.length; lineIndex += 1) {
    const contextWindow = findLineWindow(reportData.lines, Math.max(0, lineIndex - 2), 12);
    const joinedWindow = contextWindow.join('\n');

    if (!itemMatchesKeywords(reportData, item, joinedWindow, textNormalizeConfig, true)) {
      continue;
    }

    if (!hasRequiredSuffixes(joinedWindow, item.requiredSuffix, textNormalizeConfig)) {
      continue;
    }

    if (hasForbiddenSuffix(joinedWindow, forbiddenSuffixes, textNormalizeConfig)) {
      continue;
    }

    for (const candidateLine of contextWindow) {
      const match = candidateLine.match(regex);
      if (!match) {
        continue;
      }

      const candidateIndex = contextWindow.indexOf(candidateLine);

      const groupIndex = Number(item.regexConfig.valueGroupIndex || 1);
      const value = match[groupIndex] || match[0];
      const numericValue = parseNumericValue(value);

      if (Array.isArray(range) && range.length === 2 && numericValue !== null) {
        if (numericValue < range[0] || numericValue > range[1]) {
          continue;
        }
      }

      const requiredSuffixes = Array.isArray(item.requiredSuffix) ? item.requiredSuffix : [];
      const lineHasRequiredSuffix = hasRequiredSuffixes(candidateLine, requiredSuffixes, textNormalizeConfig);
      const sameLineKeywordScore = itemMatchesKeywords(reportData, item, candidateLine, textNormalizeConfig, true) ? 500 : 0;
      const nearestSuffixDistance = requiredSuffixes.length === 0
        ? 0
        : contextWindow.reduce((bestDistance, line, windowIndex) => {
          if (windowIndex > candidateIndex) {
            return bestDistance;
          }

          if (!hasRequiredSuffixes(line, requiredSuffixes, textNormalizeConfig)) {
            return bestDistance;
          }

          const distance = candidateIndex - windowIndex;
          return Math.min(bestDistance, distance);
        }, Number.POSITIVE_INFINITY);

      const proximityScore = Number.isFinite(nearestSuffixDistance)
        ? Math.max(0, 200 - nearestSuffixDistance * 60)
        : 0;

      candidates.push({
        matched: true,
        value: candidateLine.includes('%') && numericValue !== null ? String(numericValue / 100) : value,
        sourcePreview: makeSourcePreview(candidateLine),
        sourceType: 'regex-line',
        score: (lineHasRequiredSuffix ? 1000 : 0) + sameLineKeywordScore + proximityScore + candidateLine.length
      });
    }
  }

  if (candidates.length > 0) {
    return candidates.sort((left, right) => right.score - left.score)[0];
  }

  return { matched: false, reason: '未匹配到正则目标文本' };
}

function resolveAmbientNoiseAverageValue(reportData, item, textNormalizeConfig) {
  const blocks = Array.isArray(reportData.ambientNoiseBlocks) ? reportData.ambientNoiseBlocks : [];
  const tableConfig = item.tableConfig;

  if (blocks.length === 0 || !tableConfig?.rowNameMatch || !tableConfig?.targetColumnName) {
    return null;
  }

  const normalizedChecklistDesc = normalizeText(item.checklistDesc || '', textNormalizeConfig);
  if (!normalizedChecklistDesc.includes('ambient noise')) {
    return null;
  }

  const normalizedMetric = normalizeText(tableConfig.targetColumnName, textNormalizeConfig);
  const metricKey = normalizedMetric.includes('s-mos')
    ? 'smos'
    : normalizedMetric.includes('n-mos')
      ? 'nmos'
      : normalizedMetric.includes('g-mos')
        ? 'gmos'
        : null;

  if (!metricKey) {
    return null;
  }

  const candidateSceneKeys = getAmbientNoiseSceneCandidates(tableConfig.rowNameMatch)
    .map((candidate) => getAmbientNoiseSceneKey(candidate))
    .filter(Boolean);

  const targetBlock = blocks.find((block) => {
    const sceneKey = getAmbientNoiseSceneKey(block.sceneLabel);
    return candidateSceneKeys.includes(sceneKey);
  });

  if (!targetBlock || !targetBlock[metricKey]) {
    return null;
  }

  return {
    matched: true,
    value: targetBlock[metricKey],
    sourcePreview: targetBlock.sourcePreview,
    sourceType: 'ambient-average-block'
  };
}

function resolveTableValue(reportData, item, textNormalizeConfig) {
  const tableConfig = item.tableConfig;
  if (!tableConfig) {
    return { matched: false, reason: '缺少 tableConfig 配置' };
  }

  const ambientNoiseValue = resolveAmbientNoiseAverageValue(reportData, item, textNormalizeConfig);
  if (ambientNoiseValue) {
    return ambientNoiseValue;
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
    if (!itemMatchesKeywords(reportData, item, rowContext.text, textNormalizeConfig, rowContext.sourceKind !== 'html-table')) {
      return false;
    }

    const descriptorText = getRowDescriptorText(rowContext);

    if (!hasRequiredSuffixes(descriptorText, item.requiredSuffix, textNormalizeConfig)) {
      return false;
    }

    if (hasForbiddenSuffix(descriptorText, forbiddenSuffixes, textNormalizeConfig)) {
      return false;
    }

    if (!hasExactRowKeywords(rowContext.text, item.exactRowKeywords, textNormalizeConfig)) {
      return false;
    }

    return true;
  });

  if (candidates.length === 0) {
    return null;
  }

  return candidates.sort((left, right) => getCandidateScore(right) - getCandidateScore(left))[0];
}

function resolveRowBasedValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId) {
  let rowContext = selectCandidateRow(reportData, item, textNormalizeConfig, globalMatchConfig);

  if (item.extractType === 'status_judge') {
    const forbiddenSuffixes = pickForbiddenSuffixes(item, globalMatchConfig);
    const explicitStatusRows = reportData.tableRows.filter((candidate) => {
      if (!itemMatchesKeywords(reportData, item, candidate.text, textNormalizeConfig, candidate.sourceKind !== 'html-table')) {
        return false;
      }

      const descriptorText = getRowDescriptorText(candidate);
      if (!hasRequiredSuffixes(descriptorText, item.requiredSuffix, textNormalizeConfig)) {
        return false;
      }

      if (hasForbiddenSuffix(descriptorText, forbiddenSuffixes, textNormalizeConfig)) {
        return false;
      }

      if (!hasExactRowKeywords(candidate.text, item.exactRowKeywords, textNormalizeConfig)) {
        return false;
      }

      return Boolean(extractStatus(candidate.text));
    });

    const explicitStatusRow = explicitStatusRows.sort((left, right) => getCandidateScore(right) - getCandidateScore(left))[0];

    if (explicitStatusRow) {
      rowContext = explicitStatusRow;
    }
  }

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
    const value = extractFormulaValue(reportData, rowContext, item, textNormalizeConfig, extractedResultsByItemId);
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
  const lineRows = lines
    .map((line, index) => {
      const cells = line.split(/\t+/).map((cell) => cell.trim()).filter(Boolean);

      return {
        tableIndex: -1,
        rowIndex: index + 1,
        headers: [],
        cells,
        text: line,
        sourceKind: 'raw-line'
      };
    })
    .filter((row) => {
      if (row.text.length <= 8) {
        return false;
      }

      const hasExplicitStatus = Boolean(extractStatus(row.text)) || row.cells.some((cell) => isStatusCell(cell));
      const hasStructuredCells = row.cells.length >= 4;

      return hasExplicitStatus && hasStructuredCells;
    });
  let derivedRows = [];

  if (measurementObject) {
    const splitRows = flattenedText
      .split(measurementObject)
      .map((segment) => segment.trim())
      .filter((segment) => {
        return segment.length > 8
          && /\b(?:ok|done|not ok)\b/i.test(segment)
          && /-?\d+(?:\.\d+)?\s*$/.test(segment);
      })
      .map((text, index) => ({
        tableIndex: -1,
        rowIndex: index + 1,
        headers: [],
        cells: [],
        text,
        sourceKind: 'measurement-split'
      }));

    if (splitRows.length > 0) {
      derivedRows = splitRows;
    }
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
          text,
          sourceKind: 'row-start'
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
      text: cells.join(' | '),
      sourceKind: 'html-table'
    }));
  });

  return {
    rawText,
    html,
    lines,
    ambientNoiseBlocks: extractAmbientNoiseAverageBlocks(rawText),
    tables,
    tableRows: [...tableRows, ...lineRows, ...derivedRows]
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

    if (!rawText.trim()) {
      throw new Error('.doc 报告转换超时或未读取到内容。请优先另存为 .docx 后重试，或关闭可能弹出的 Word/WPS 隐藏窗口。');
    }

    return createSearchData(rawText, '');
  }

  return parseDocxReport(reportPath);
}

async function applyResultsToChecklist(checklistPath, reportPath, extractedItems) {
  const workbook = XLSX.readFile(checklistPath, { cellStyles: true, cellDates: true });
  const sheetName = resolveChecklistSheetName(workbook, reportPath);
  const worksheet = workbook.Sheets[sheetName];
  const numericIColumnCells = [];
  const percentIColumnCells = [];
  const preservedMerges = collectWorksheetMerges(worksheet, {
    columns: ['A', 'B', 'C'],
    startRow: 6,
    endRow: CHECKLIST_BORDER_RANGE.endRow
  });

  splitWorksheetMergesInRange(worksheet, CHECKLIST_BORDER_RANGE);
  ensureWorksheetRangeCells(worksheet, CHECKLIST_BORDER_RANGE);

  extractedItems.forEach((itemResult) => {
    if (!itemResult.matched || !itemResult.value) {
      return;
    }

    const isNumericCell = writeChecklistCell(worksheet, itemResult.outputCell, itemResult.value);
    if (isNumericCell && /^I\d+$/i.test(itemResult.outputCell)) {
      if (PERCENT_OUTPUT_CELLS.has(itemResult.outputCell.toUpperCase())) {
        percentIColumnCells.push(itemResult.outputCell);
      } else {
        numericIColumnCells.push(itemResult.outputCell);
      }
    }
  });

  const reportName = path.parse(reportPath).name;
  const outputPath = path.join(path.dirname(reportPath), buildOutputFileName(reportName));
  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx' });

  if (path.extname(checklistPath).toLowerCase() === '.xlsx') {
    await patchChecklistSheetStyles(checklistPath, outputPath, sheetName, {
      decimalCellAddresses: numericIColumnCells,
      percentCellAddresses: percentIColumnCells,
      borderedRange: CHECKLIST_BORDER_RANGE
    });

    await patchChecklistSheetMerges(outputPath, sheetName, preservedMerges);
  }

  return outputPath;
}

function collectWorksheetMerges(worksheet, mergeConfig) {
  if (!Array.isArray(worksheet['!merges']) || worksheet['!merges'].length === 0) {
    return [];
  }

  const allowedColumns = new Set((mergeConfig.columns || []).map((column) => XLSX.utils.decode_col(column)));
  const startRow = (mergeConfig.startRow || 1) - 1;
  const endRow = (mergeConfig.endRow || Number.MAX_SAFE_INTEGER) - 1;

  return worksheet['!merges']
    .filter((mergeRange) => {
      const isSingleColumn = mergeRange.s.c === mergeRange.e.c;
      const inAllowedColumn = allowedColumns.has(mergeRange.s.c);
      const inTargetRows = mergeRange.s.r >= startRow && mergeRange.e.r <= endRow;
      const isVerticalMerge = mergeRange.e.r > mergeRange.s.r;

      return isSingleColumn && inAllowedColumn && inTargetRows && isVerticalMerge;
    })
    .map((mergeRange) => ({
      s: { ...mergeRange.s },
      e: { ...mergeRange.e }
    }));
}

function splitWorksheetMergesInRange(worksheet, rangeConfig) {
  if (!Array.isArray(worksheet['!merges']) || worksheet['!merges'].length === 0) {
    return;
  }

  const targetStartCol = XLSX.utils.decode_col(rangeConfig.startCol);
  const targetEndCol = XLSX.utils.decode_col(rangeConfig.endCol);
  const targetStartRow = rangeConfig.startRow - 1;
  const targetEndRow = rangeConfig.endRow - 1;

  worksheet['!merges'] = worksheet['!merges'].filter((mergeRange) => {
    const intersects = !(
      mergeRange.e.c < targetStartCol
      || mergeRange.s.c > targetEndCol
      || mergeRange.e.r < targetStartRow
      || mergeRange.s.r > targetEndRow
    );

    return !intersects;
  });

  if (worksheet['!merges'].length === 0) {
    delete worksheet['!merges'];
  }
}

function ensureWorksheetRangeCells(worksheet, rangeConfig) {
  const startColIndex = XLSX.utils.decode_col(rangeConfig.startCol);
  const endColIndex = XLSX.utils.decode_col(rangeConfig.endCol);

  for (let row = rangeConfig.startRow; row <= rangeConfig.endRow; row += 1) {
    for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex += 1) {
      const cellAddress = XLSX.utils.encode_cell({ c: colIndex, r: row - 1 });
      if (!worksheet[cellAddress]) {
        worksheet[cellAddress] = {
          t: 's',
          v: ''
        };
      }
    }
  }
}

function writeChecklistCell(worksheet, cellAddress, rawValue) {
  const worksheetValue = toWorksheetValue(rawValue);
  const templateCell = worksheet[cellAddress] || {};
  const style = {
    ...(templateCell.s || {}),
    alignment: {
      ...((templateCell.s && templateCell.s.alignment) || {}),
      horizontal: 'center',
      vertical: 'center'
    }
  };

  if (typeof worksheetValue === 'number' && Number.isFinite(worksheetValue)) {
    worksheet[cellAddress] = {
      ...(templateCell || {}),
      t: 'n',
      v: worksheetValue,
      z: '0.00',
      s: style
    };
    delete worksheet[cellAddress].w;
    return true;
  }

  worksheet[cellAddress] = {
    ...(templateCell || {}),
    t: 's',
    v: String(worksheetValue ?? ''),
    s: style
  };
  delete worksheet[cellAddress].w;
  return false;
}

function toWorksheetValue(value) {
  if (typeof value !== 'string') {
    return value;
  }

  const trimmedValue = value.trim();
  if (/^-?\d+(?:\.\d+)?$/.test(trimmedValue)) {
    const numericValue = Number.parseFloat(trimmedValue);
    if (!Number.isNaN(numericValue)) {
      return numericValue;
    }
  }

  return value;
}

function normalizeSheetToken(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function resolveChecklistSheetName(workbook, reportPath) {
  const reportName = path.parse(reportPath).name;
  const reportTokens = reportName.split(/[_\s-]+/).map(normalizeSheetToken).filter(Boolean);
  const sheetNames = getWorkbookSheetNames(workbook);

  const preferredModeSheet = [
    {
      sheetName: 'Handset',
      reportTokens: new Set(['ha', 'handset'])
    },
    {
      sheetName: 'Handsfree',
      reportTokens: new Set(['hf', 'handsfree'])
    },
    {
      sheetName: 'Headset',
      reportTokens: new Set(['hs', 'headset'])
    }
  ].find((candidate) => {
    const hasSheet = sheetNames.some((sheetName) => normalizeSheetToken(sheetName) === normalizeSheetToken(candidate.sheetName));
    if (!hasSheet) {
      return false;
    }

    return reportTokens.some((token) => candidate.reportTokens.has(token));
  });

  if (preferredModeSheet) {
    return preferredModeSheet.sheetName;
  }

  const preferredSheet = sheetNames.find((candidate) => {
    const normalizedCandidate = normalizeSheetToken(candidate);
    if (!normalizedCandidate || ['report', 'help', 'changelog'].includes(normalizedCandidate)) {
      return false;
    }

    return reportTokens.some((token) => token && (token.includes(normalizedCandidate) || normalizedCandidate.includes(token)));
  });

  if (preferredSheet) {
    return preferredSheet;
  }

  const handsetFallbackSheet = sheetNames.find((candidate) => normalizeSheetToken(candidate) === 'handset');
  if (handsetFallbackSheet) {
    return handsetFallbackSheet;
  }

  return sheetNames[0];
}

function getWorkbookSheetNames(workbook) {
  if (Array.isArray(workbook?.SheetNames)) {
    return workbook.SheetNames;
  }

  if (Array.isArray(workbook?.worksheets)) {
    return workbook.worksheets.map((worksheet) => worksheet.name);
  }

  return [];
}

async function patchChecklistSheetStyles(templatePath, outputPath, sheetName, styleOptions = {}) {
  const [templateBuffer, outputBuffer] = await Promise.all([
    fs.readFile(templatePath),
    fs.readFile(outputPath)
  ]);

  const [templateZip, outputZip] = await Promise.all([
    JSZip.loadAsync(templateBuffer),
    JSZip.loadAsync(outputBuffer)
  ]);

  const worksheetPath = await resolveWorksheetPath(outputZip, sheetName);
  if (!worksheetPath) {
    return;
  }

  const templateWorksheetPath = await resolveWorksheetPath(templateZip, sheetName);
  if (!templateWorksheetPath) {
    return;
  }

  let outputSheetXml = await outputZip.file(worksheetPath).async('string');
  const templateSheetXml = await templateZip.file(templateWorksheetPath).async('string');
  const templateStylesXml = await templateZip.file('xl/styles.xml')?.async('string');
  const styleMap = extractSheetStyleMap(templateSheetXml);
  let patchedStylesXml = templateStylesXml || null;
  const cellStyleMap = new Map(styleMap);
  let borderAssignments = [];

  if (patchedStylesXml && Array.isArray(styleOptions.decimalCellAddresses) && styleOptions.decimalCellAddresses.length > 0) {
    patchedStylesXml = createNumberFormatStyleOverrides(patchedStylesXml, cellStyleMap, styleOptions.decimalCellAddresses, '2');
  }

  if (patchedStylesXml && Array.isArray(styleOptions.percentCellAddresses) && styleOptions.percentCellAddresses.length > 0) {
    patchedStylesXml = createNumberFormatStyleOverrides(patchedStylesXml, cellStyleMap, styleOptions.percentCellAddresses, '10');
  }

  if (patchedStylesXml && styleOptions.borderedRange) {
    borderAssignments = buildBorderAssignments(styleOptions.borderedRange);
    patchedStylesXml = createBorderStyleOverrides(patchedStylesXml, cellStyleMap, borderAssignments);
  }

  outputSheetXml = ensureSheetCellsExist(outputSheetXml, [
    ...cellStyleMap.keys(),
    ...borderAssignments.map((item) => item.cellAddress)
  ]);

  cellStyleMap.forEach((styleId, cellAddress) => {
    if (styleId === undefined || styleId === null) {
      return;
    }

    outputSheetXml = applyStyleIdToCell(outputSheetXml, cellAddress, styleId);
  });

  outputZip.file(worksheetPath, outputSheetXml);

  if (patchedStylesXml) {
    outputZip.file('xl/styles.xml', patchedStylesXml);
  }

  const patchedBuffer = await outputZip.generateAsync({ type: 'nodebuffer' });
  await fs.writeFile(outputPath, patchedBuffer);
}

async function resolveWorksheetPath(zip, sheetName) {
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');

  if (!workbookXml || !relsXml) {
    return null;
  }

  const escapedSheetName = escapeRegExp(sheetName);
  const sheetMatch = new RegExp(`<sheet[^>]*name="${escapedSheetName}"[^>]*r:id="([^"]+)"`, 'i').exec(workbookXml);
  if (!sheetMatch) {
    return null;
  }

  const relationId = sheetMatch[1];
  const relMatch = new RegExp(`<Relationship[^>]*Id="${escapeRegExp(relationId)}"[^>]*Target="([^"]+)"`, 'i').exec(relsXml);
  if (!relMatch) {
    return null;
  }

  return `xl/${relMatch[1].replace(/^\//, '')}`;
}

async function patchChecklistSheetMerges(outputPath, sheetName, mergeRanges) {
  if (!Array.isArray(mergeRanges) || mergeRanges.length === 0) {
    return;
  }

  const outputBuffer = await fs.readFile(outputPath);
  const outputZip = await JSZip.loadAsync(outputBuffer);
  const worksheetPath = await resolveWorksheetPath(outputZip, sheetName);
  if (!worksheetPath) {
    return;
  }

  let outputSheetXml = await outputZip.file(worksheetPath).async('string');
  const existingRefs = Array.from(outputSheetXml.matchAll(/<mergeCell\s+ref="([^"]+)"\s*\/>/gi)).map((match) => match[1]);
  const deferredRefs = mergeRanges.map((mergeRange) => XLSX.utils.encode_range(mergeRange));
  const mergeRefs = Array.from(new Set([...existingRefs, ...deferredRefs]));

  if (mergeRefs.length === 0) {
    return;
  }

  const mergeCellsXml = `<mergeCells count="${mergeRefs.length}">${mergeRefs.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`;

  if (/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/i.test(outputSheetXml)) {
    outputSheetXml = outputSheetXml.replace(/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/i, mergeCellsXml);
  } else if (/<\/sheetData>/i.test(outputSheetXml)) {
    outputSheetXml = outputSheetXml.replace(/<\/sheetData>/i, `</sheetData>${mergeCellsXml}`);
  } else {
    return;
  }

  outputZip.file(worksheetPath, outputSheetXml);
  const patchedBuffer = await outputZip.generateAsync({ type: 'nodebuffer' });
  await fs.writeFile(outputPath, patchedBuffer);
}

function extractSheetStyleMap(sheetXml) {
  const styleMap = new Map();

  const cellPattern = /<c\b[^>]*r="([^"]+)"[^>]*\bs="([^"]+)"[^>]*(?:\/?>)/gi;
  let match;
  while ((match = cellPattern.exec(sheetXml)) !== null) {
    styleMap.set(match[1], match[2]);
  }

  return styleMap;
}

function applyStyleIdToCell(sheetXml, cellAddress, styleId) {
  const pattern = new RegExp(`<c\\b[^>]*r="${escapeRegExp(cellAddress)}"[^>]*\/?>`, 'i');
  return sheetXml.replace(pattern, (match) => {
    const normalizedMatch = match.replace(/\bs="[^"]+"/i, '').replace(/\s*\/?>$/, '');
    const suffix = /\/\s*>$/.test(match) || /\/>$/.test(match) ? '/>' : '>';
    return `${normalizedMatch} s="${styleId}"${suffix}`;
  });
}

function ensureSheetCellsExist(sheetXml, cellAddresses) {
  if (!Array.isArray(cellAddresses) || cellAddresses.length === 0) {
    return sheetXml;
  }

  const sheetDataMatch = /<sheetData>([\s\S]*?)<\/sheetData>/i.exec(sheetXml);
  if (!sheetDataMatch) {
    return sheetXml;
  }

  const rowMap = new Map();
  const rowPattern = /<row\b[^>]*r="(\d+)"[^>]*>[\s\S]*?<\/row>/gi;
  let rowMatch;

  while ((rowMatch = rowPattern.exec(sheetDataMatch[1])) !== null) {
    rowMap.set(Number(rowMatch[1]), rowMatch[0]);
  }

  const uniqueAddresses = Array.from(new Set(cellAddresses)).sort((left, right) => {
    const leftRef = XLSX.utils.decode_cell(left);
    const rightRef = XLSX.utils.decode_cell(right);

    if (leftRef.r !== rightRef.r) {
      return leftRef.r - rightRef.r;
    }

    return leftRef.c - rightRef.c;
  });

  uniqueAddresses.forEach((cellAddress) => {
    const { r } = XLSX.utils.decode_cell(cellAddress);
    const rowNumber = r + 1;

    if (!rowMap.has(rowNumber)) {
      rowMap.set(rowNumber, `<row r="${rowNumber}"><c r="${cellAddress}"/></row>`);
      return;
    }

    rowMap.set(rowNumber, ensureCellExistsInRow(rowMap.get(rowNumber), cellAddress));
  });

  const rebuiltSheetData = Array.from(rowMap.entries())
    .sort((left, right) => left[0] - right[0])
    .map(([, rowXml]) => rowXml)
    .join('');

  return sheetXml.replace(sheetDataMatch[0], `<sheetData>${rebuiltSheetData}</sheetData>`);
}

function ensureCellExistsInRow(rowXml, cellAddress) {
  if (new RegExp(`<c\\b[^>]*r="${escapeRegExp(cellAddress)}"`, 'i').test(rowXml)) {
    return rowXml;
  }

  const rowStartMatch = /^<row\b[^>]*>/i.exec(rowXml);
  if (!rowStartMatch) {
    return rowXml;
  }

  const rowStart = rowStartMatch[0];
  const rowBody = rowXml.slice(rowStart.length, -'</row>'.length);
  const cellPattern = /<c\b[^>]*r="([A-Z]+\d+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/gi;
  const cells = [];
  let cellMatch;

  while ((cellMatch = cellPattern.exec(rowBody)) !== null) {
    cells.push({
      ref: cellMatch[1],
      xml: cellMatch[0]
    });
  }

  cells.push({
    ref: cellAddress,
    xml: `<c r="${cellAddress}"/>`
  });

  cells.sort((left, right) => XLSX.utils.decode_cell(left.ref).c - XLSX.utils.decode_cell(right.ref).c);

  return `${rowStart}${cells.map((cell) => cell.xml).join('')}</row>`;
}

function getCellXfList(stylesXml) {
  const cellXfsMatch = /(<cellXfs\b[^>]*count=")([^"]+)("[^>]*>)([\s\S]*?)(<\/cellXfs>)/i.exec(stylesXml);
  if (!cellXfsMatch) {
    return null;
  }

  return {
    match: cellXfsMatch,
    xfList: cellXfsMatch[4].match(/<xf\b[\s\S]*?<\/xf>|<xf\b[^>]*\/>/gi) || []
  };
}

function resolveCellStyleId(cellStyleMap, cellAddress) {
  const styleId = cellStyleMap.get(cellAddress);
  if (styleId === undefined || styleId === null || styleId === '') {
    return '0';
  }

  return String(styleId);
}

function createNumberFormatStyleOverrides(stylesXml, cellStyleMap, cellAddresses, numFmtId) {
  const styleIds = Array.from(new Set(
    cellAddresses
      .map((cellAddress) => resolveCellStyleId(cellStyleMap, cellAddress))
      .map((styleId) => Number(styleId))
      .filter((styleId) => Number.isInteger(styleId) && styleId >= 0)
  ));

  if (styleIds.length === 0) {
    return stylesXml;
  }

  const cellXfData = getCellXfList(stylesXml);
  if (!cellXfData) {
    return stylesXml;
  }

  const { match: cellXfsMatch, xfList } = cellXfData;
  const overrideStyleMap = new Map();

  styleIds.forEach((styleId) => {
    const sourceXf = xfList[styleId];
    if (!sourceXf) {
      return;
    }

    const decimalXf = forceNumberFormatXf(sourceXf, numFmtId);
    let overrideStyleId = xfList.findIndex((xf) => xf === decimalXf);
    if (overrideStyleId === -1) {
      xfList.push(decimalXf);
      overrideStyleId = xfList.length - 1;
    }

    overrideStyleMap.set(String(styleId), String(overrideStyleId));
  });

  const updatedStylesXml = stylesXml.replace(
    cellXfsMatch[0],
    `${cellXfsMatch[1]}${xfList.length}${cellXfsMatch[3]}${xfList.join('')}${cellXfsMatch[5]}`
  );

  cellAddresses.forEach((cellAddress) => {
    const sourceStyleId = resolveCellStyleId(cellStyleMap, cellAddress);
    const overrideStyleId = sourceStyleId !== undefined && sourceStyleId !== null
      ? overrideStyleMap.get(String(sourceStyleId))
      : null;

    if (overrideStyleId) {
      cellStyleMap.set(cellAddress, overrideStyleId);
    }
  });

  return updatedStylesXml;
}

function forceNumberFormatXf(xfXml, numFmtId) {
  const xfMatch = /^<xf\b([^>]*?)(\/?>[\s\S]*)$/i.exec(xfXml);
  if (!xfMatch) {
    return xfXml;
  }

  let attributes = xfMatch[1];
  const suffix = xfMatch[2];

  attributes = upsertXmlAttribute(attributes, 'numFmtId', String(numFmtId));
  attributes = upsertXmlAttribute(attributes, 'applyNumberFormat', '1');

  return `<xf${attributes}${suffix}`;
}

function buildBorderAssignments(rangeConfig) {
  const startColIndex = XLSX.utils.decode_col(rangeConfig.startCol);
  const endColIndex = XLSX.utils.decode_col(rangeConfig.endCol);
  const assignments = [];

  for (let row = rangeConfig.startRow; row <= rangeConfig.endRow; row += 1) {
    for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex += 1) {
      assignments.push({
        cellAddress: XLSX.utils.encode_cell({ c: colIndex, r: row - 1 }),
        borderSpec: {
          left: colIndex === startColIndex ? 'thick' : 'thin',
          right: colIndex === endColIndex ? 'thick' : 'thin',
          top: row === rangeConfig.startRow ? 'thick' : 'thin',
          bottom: row === rangeConfig.endRow ? 'thick' : 'thin'
        }
      });
    }
  }

  return assignments;
}

function createBorderStyleOverrides(stylesXml, cellStyleMap, borderAssignments) {
  if (!Array.isArray(borderAssignments) || borderAssignments.length === 0) {
    return stylesXml;
  }

  const bordersMatch = /(<borders\b[^>]*count=")([^"]+)("[^>]*>)([\s\S]*?)(<\/borders>)/i.exec(stylesXml);
  const cellXfData = getCellXfList(stylesXml);

  if (!bordersMatch || !cellXfData) {
    return stylesXml;
  }

  const borderList = bordersMatch[4].match(/<border\b[\s\S]*?<\/border>|<border\b[^>]*\/>/gi) || [];
  const { match: cellXfsMatch, xfList } = cellXfData;
  const overrideStyleMap = new Map();

  borderAssignments.forEach(({ cellAddress, borderSpec }) => {
    const sourceStyleId = resolveCellStyleId(cellStyleMap, cellAddress);
    const cacheKey = `${sourceStyleId}|${JSON.stringify(borderSpec)}`;

    if (!overrideStyleMap.has(cacheKey)) {
      const sourceXf = xfList[Number(sourceStyleId)] || xfList[0];
      if (!sourceXf) {
        return;
      }

      const borderXml = buildBorderXml(borderSpec);
      let borderId = borderList.findIndex((border) => border === borderXml);
      if (borderId === -1) {
        borderList.push(borderXml);
        borderId = borderList.length - 1;
      }

      const borderXf = forceBorderXf(sourceXf, borderId);
      let overrideStyleId = xfList.findIndex((xf) => xf === borderXf);
      if (overrideStyleId === -1) {
        xfList.push(borderXf);
        overrideStyleId = xfList.length - 1;
      }

      overrideStyleMap.set(cacheKey, String(overrideStyleId));
    }

    const overrideStyleId = overrideStyleMap.get(cacheKey);
    if (overrideStyleId) {
      cellStyleMap.set(cellAddress, overrideStyleId);
    }
  });

  let updatedStylesXml = stylesXml.replace(
    bordersMatch[0],
    `${bordersMatch[1]}${borderList.length}${bordersMatch[3]}${borderList.join('')}${bordersMatch[5]}`
  );

  updatedStylesXml = updatedStylesXml.replace(
    cellXfsMatch[0],
    `${cellXfsMatch[1]}${xfList.length}${cellXfsMatch[3]}${xfList.join('')}${cellXfsMatch[5]}`
  );

  return updatedStylesXml;
}

function buildBorderXml(borderSpec) {
  return `<border>${buildBorderSideXml('left', borderSpec.left)}${buildBorderSideXml('right', borderSpec.right)}${buildBorderSideXml('top', borderSpec.top)}${buildBorderSideXml('bottom', borderSpec.bottom)}<diagonal/></border>`;
}

function buildBorderSideXml(sideName, styleValue) {
  if (!styleValue) {
    return `<${sideName}/>`;
  }

  return `<${sideName} style="${styleValue}"><color auto="1"/></${sideName}>`;
}

function forceBorderXf(xfXml, borderId) {
  const xfMatch = /^<xf\b([^>]*?)(\/?>[\s\S]*)$/i.exec(xfXml);
  if (!xfMatch) {
    return xfXml;
  }

  let attributes = xfMatch[1];
  const suffix = xfMatch[2];

  attributes = upsertXmlAttribute(attributes, 'borderId', String(borderId));
  attributes = upsertXmlAttribute(attributes, 'applyBorder', '1');

  return `<xf${attributes}${suffix}`;
}

function upsertXmlAttribute(attributeText, attributeName, attributeValue) {
  const pattern = new RegExp(`\\b${escapeRegExp(attributeName)}="[^"]*"`, 'i');
  if (pattern.test(attributeText)) {
    return attributeText.replace(pattern, `${attributeName}="${attributeValue}"`);
  }

  return `${attributeText} ${attributeName}="${attributeValue}"`;
}

function buildOutputFileName(reportName) {
  const normalizedReportName = String(reportName || '').trim();
  const timestamp = buildTimestamp();

  if (!normalizedReportName) {
    return `checklist_${timestamp}.xlsx`;
  }

  return `${normalizedReportName}_checklist_${timestamp}.xlsx`;
}

async function processSingleReport({ reportPath, checklistPath, rules }) {
  const reportData = await parseReport(reportPath);
  const textNormalizeConfig = rules.globalMatchConfig?.textNormalize || {};
  const globalMatchConfig = rules.globalMatchConfig || {};

  const extractedResultsByItemId = new Map();
  const extractedItems = rules.extractItemList.map((item) => {
    let extractionResult;

    if (['summary_table_match', 'formula_calc', 'status_judge'].includes(item.extractType)) {
      extractionResult = resolveRowBasedValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);
    } else if (item.extractType === 'detail_anchor_extract') {
      extractionResult = resolveAnchorValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);
    } else if (item.extractType === 'table_struct_extract') {
      extractionResult = resolveTableValue(reportData, item, textNormalizeConfig);
    } else if (item.extractType === 'regex_extract') {
      extractionResult = resolveRegexValue(reportData, item, textNormalizeConfig, globalMatchConfig);
    } else {
      extractionResult = { matched: false, reason: `暂不支持的提取类型: ${item.extractType}` };
    }

    const itemResult = {
      itemId: item.itemId,
      checklistDesc: item.checklistDesc,
      outputCell: item.outputCell,
      extractType: item.extractType,
      ...extractionResult
    };

    extractedResultsByItemId.set(item.itemId, itemResult);
    return itemResult;
  });

  const outputPath = await applyResultsToChecklist(checklistPath, reportPath, extractedItems);
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