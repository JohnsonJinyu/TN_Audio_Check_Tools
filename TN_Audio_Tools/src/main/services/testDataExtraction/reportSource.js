const fs = require('fs/promises');
const path = require('path');
const mammoth = require('mammoth');
const JSON5 = require('json5');
const { parseDocxStructuredData } = require('./docxStructuredParser');

function normalizeReportBandwidth(value) {
  const normalized = String(value || '').trim().toUpperCase();
  if (!normalized) {
    return '';
  }

  if (['SB', 'SWB'].includes(normalized)) {
    return 'SWB';
  }

  if (normalized.endsWith('SWB')) {
    return 'SWB';
  }

  if (normalized.endsWith('WB')) {
    return 'WB';
  }

  if (normalized.endsWith('NB')) {
    return 'NB';
  }

  return '';
}

function deriveBandwidthFromPath(reportPath) {
  const normalizedPath = String(reportPath || '').toUpperCase();
  if (/([_\-\s]|^)SWB([_\-\s.]|$)/.test(normalizedPath) || /\bSB\b/.test(normalizedPath)) {
    return 'SWB';
  }

  if (/([_\-\s]|^)WB([_\-\s.]|$)/.test(normalizedPath)) {
    return 'WB';
  }

  if (/([_\-\s]|^)NB([_\-\s.]|$)/.test(normalizedPath)) {
    return 'NB';
  }

  return '';
}

function deriveTerminalModeFromPath(reportPath) {
  const normalizedPath = String(reportPath || '').toUpperCase();
  if (/\bEID\b/.test(normalizedPath) || /\bDIGITAL\b/.test(normalizedPath)) {
    return 'EID';
  }
  if (/\bVENICE\b/.test(normalizedPath) || /\bANALOG\b/.test(normalizedPath) || /\bELECTRICAL\b/.test(normalizedPath)) {
    return 'EI';
  }
  return '';
}

function deriveBandwidthFromText(rawText) {
  const normalizedText = String(rawText || '').toUpperCase();
  const directMatches = [
    normalizedText.match(/\b(?:AMR|EVS)[_\-\s]?SWB\b/),
    normalizedText.match(/\bSWB\b/),
    normalizedText.match(/\b(?:AMR|EVS)[_\-\s]?WB\b/),
    normalizedText.match(/\bWB\b/),
    normalizedText.match(/\b(?:AMR|EVS)[_\-\s]?NB\b/),
    normalizedText.match(/\bNB\b/)
  ].filter(Boolean);

  for (const match of directMatches) {
    const bandwidth = normalizeReportBandwidth(match[0]);
    if (bandwidth) {
      return bandwidth;
    }
  }

  return '';
}

function deriveTokenByCandidates(sourceText, candidates) {
  const normalizedText = String(sourceText || '').toUpperCase();
  if (!normalizedText) {
    return '';
  }

  return candidates.find((candidate) => new RegExp(`(^|[_\\-\\s])${candidate}([_\\-\\s.]|$)`, 'i').test(normalizedText)) || '';
}

function deriveMeasurementObject(rawText) {
  const normalizedText = String(rawText || '');
  if (!normalizedText.trim()) {
    return '';
  }

  const patterns = [
    /Measurement Object\s*[:：-]?\s*([^\r\n]+)/i,
    /Object\s*[:：-]?\s*([^\r\n]+)/i
  ];

  for (const pattern of patterns) {
    const matched = normalizedText.match(pattern);
    const candidate = matched?.[1]?.trim();
    if (candidate) {
      return candidate;
    }
  }

  return '';
}

function deriveProjectNameAndPhase(reportPath) {
  const reportName = path.parse(reportPath || '').name;
  const normalizedName = String(reportName || '').trim();

  // 阶段: EVB / EVT / DVT1 / DVT2 / PVT
  const phaseMatch = normalizedName.match(/(EVB|EVT|DVT[12]|PVT)/i);
  const projectPhase = phaseMatch ? phaseMatch[1].toUpperCase() : '';

  // 项目名: 文件名中第一个下划线前的部分，或第一个非阶段/非带宽/非codec的单词
  let projectName = '';
  if (normalizedName) {
    const firstUnderscore = normalizedName.indexOf('_');
    if (firstUnderscore > 0) {
      projectName = normalizedName.substring(0, firstUnderscore).toUpperCase();
    } else {
      // 没有下划线，取第一个单词
      const firstWord = normalizedName.split(/[\s_-]+/)[0];
      if (firstWord) projectName = firstWord.toUpperCase();
    }
  }

  return { projectName, projectPhase };
}

function deriveReportMetadata(reportPath, rawText, reportData) {
  const reportName = path.parse(reportPath || '').name;
  const reportContext = reportData?.reportContext || {};
  const measurementObject = String(reportContext.measurementObject || '').trim() || deriveMeasurementObject(rawText) || reportName;
  const combinedSource = `${reportName} ${measurementObject} ${rawText || ''}`;
  const { projectName, projectPhase } = deriveProjectNameAndPhase(reportPath);

  return {
    reportName,
    measurementObject,
    bandwidth: normalizeReportBandwidth(reportContext.bandwidth) || deriveBandwidthFromPath(reportPath) || deriveBandwidthFromText(rawText) || '',
    codec: String(reportContext.codec || '').trim().toUpperCase() || deriveTokenByCandidates(combinedSource, ['EVS', 'AMR']),
    network: String(reportContext.network || '').trim().toUpperCase() || deriveTokenByCandidates(combinedSource, ['VOLTE', 'VOWIFI', 'VONR', 'VOIP', 'WCDMA', 'GSM']),
    terminalMode: String(reportContext.terminalMode || '').trim().toUpperCase() || deriveTerminalModeFromPath(reportPath) || deriveTokenByCandidates(combinedSource, ['HA', 'HF', 'HS', 'HE', 'HH', 'EID', 'EIA', 'EI']),
    projectName,
    projectPhase
  };
}

function attachReportContext(reportData, reportPath, rawText = '') {
  const metadata = deriveReportMetadata(reportPath, rawText, reportData);

  return {
    ...reportData,
    reportContext: {
      ...(reportData?.reportContext || {}),
      ...metadata
    }
  };
}

function isSingleRulesConfig(rules) {
  return Array.isArray(rules?.extractItemList);
}

function isRuleBundleConfig(rules) {
  return rules && typeof rules === 'object' && rules.ruleProfiles && typeof rules.ruleProfiles === 'object';
}

function serializeRulesForExport(rules) {
  if (isSingleRulesConfig(rules)) {
    return rules;
  }

  if (isRuleBundleConfig(rules)) {
    return {
      ruleBaseInfo: rules.ruleBaseInfo || {},
      defaultProfileKey: rules.defaultProfileKey || '',
      ruleProfiles: Object.fromEntries(Object.entries(rules.ruleProfiles).map(([profileKey, profileRules]) => [
        profileKey,
        serializeRulesForExport(profileRules)
      ]))
    };
  }

  return rules;
}

function createReportSource({
  supportedReportExtensions,
  convertDocToTemporaryDocx,
  wordExtractor,
  createSearchData,
  parseXlsxReport
}) {
  async function normalizeRulesConfig(rulePath, rules) {
    if (isSingleRulesConfig(rules)) {
      return rules;
    }

    if (!isRuleBundleConfig(rules)) {
      throw new Error('规则文件缺少 extractItemList 配置');
    }

    const profileEntries = Object.entries(rules.ruleProfiles || {});
    if (profileEntries.length === 0) {
      throw new Error('规则文件缺少 ruleProfiles 配置');
    }

    const normalizedProfiles = {};
    for (const [profileKey, profileConfig] of profileEntries) {
      if (typeof profileConfig === 'string') {
        normalizedProfiles[profileKey] = await loadRules(path.resolve(path.dirname(rulePath), profileConfig));
        continue;
      }

      if (profileConfig && typeof profileConfig.rulePath === 'string') {
        const subRules = await loadRules(path.resolve(path.dirname(rulePath), profileConfig.rulePath));
        // 保留 bundle 层级的元数据（如 defaultChecklistPath）
        if (profileConfig.defaultChecklistPath) {
          subRules._defaultChecklistPath = profileConfig.defaultChecklistPath;
        }
        normalizedProfiles[profileKey] = subRules;
        continue;
      }

      normalizedProfiles[profileKey] = await normalizeRulesConfig(rulePath, profileConfig);
    }

    return {
      ruleBaseInfo: rules.ruleBaseInfo || {},
      defaultProfileKey: String(rules.defaultProfileKey || profileEntries[0][0]).trim() || profileEntries[0][0],
      ruleProfiles: normalizedProfiles,
      _ruleDir: path.dirname(rulePath)
    };
  }

  async function loadRules(rulePath) {
    const content = await fs.readFile(rulePath, 'utf8');
    const rules = JSON5.parse(content);

    return normalizeRulesConfig(rulePath, rules);
  }

  async function buildExportableRulesContent(rulePath) {
    const content = await fs.readFile(rulePath, 'utf8');
    const rules = JSON5.parse(content);

    if (isSingleRulesConfig(rules)) {
      return content;
    }

    return `${JSON.stringify(serializeRulesForExport(await normalizeRulesConfig(rulePath, rules)), null, 2)}\n`;
  }

  async function parseDocxReport(reportPath) {
    const [rawTextResult, htmlResult, structuredData] = await Promise.all([
      mammoth.extractRawText({ path: reportPath }),
      mammoth.convertToHtml({ path: reportPath }),
      parseDocxStructuredData(reportPath).catch(() => ({ lines: [], tables: [], headers: [], footers: [], pageCount: null }))
    ]);

    const searchData = createSearchData(rawTextResult.value || '', htmlResult.value || '', structuredData);

    return attachReportContext(
      {
        ...searchData,
        reportFormat: 'docx',
        structuredData,
        pageCount: structuredData.pageCount || null
      },
      reportPath,
      rawTextResult.value || ''
    );
  }

  // 解析入口只负责拿到标准化的搜索数据，不参与后续提取规则判断。
  async function parseReport(reportPath, options = {}) {
    const onProgress = typeof options.onProgress === 'function' ? options.onProgress : null;
    const reportExtension = path.extname(reportPath).toLowerCase();
    if (!supportedReportExtensions.has(reportExtension)) {
      throw new Error('当前仅支持 .xlsx / .xls / .doc / .docx 测试报告');
    }

    if (reportExtension === '.xlsx' || reportExtension === '.xls') {
      if (onProgress) onProgress({ status: 'running', detail: '正在解析 xlsx 测试数据' });
      let xlsxData = await parseXlsxReport(reportPath);

      return attachReportContext(xlsxData, reportPath);
    }

    if (reportExtension === '.doc') {
      const converted = await convertDocToTemporaryDocx(reportPath, {
        onProgress: onProgress
      });

      if (converted?.convertedPath) {
        if (onProgress) onProgress({ status: 'running', detail: '格式转换完成，正在解析转换后的 Word 报告' });
        const docxData = await parseDocxReport(converted.convertedPath);
        docxData._convertedDocxPath = converted.convertedPath;
        return docxData;
      }

      if (onProgress) onProgress({ status: 'running', detail: '未能完成保真转换，正在回退到文本提取模式' });
      const extracted = await wordExtractor.extract(reportPath);
      const rawText = [
        extracted.getHeaders?.() || '',
        extracted.getBody?.() || '',
        extracted.getFootnotes?.() || '',
        extracted.getEndnotes?.() || '',
        extracted.getTextboxes?.() || ''
      ].filter(Boolean).join('\n');

      if (!rawText.trim()) {
        throw new Error('.doc 报告未读取到有效文本内容。请优先另存为 .docx 后重试。');
      }

      return attachReportContext({
        ...createSearchData(rawText, ''),
        reportFormat: 'doc',
        structuredData: {
          lines: [],
          tables: [],
          headers: extracted.getHeaders?.() ? [extracted.getHeaders()] : [],
          footers: [],
          pageCount: null
        }
      }, reportPath, rawText);
    }

    if (onProgress) onProgress({ status: 'running', detail: '正在解析 Word 报告结构与内容' });
    return parseDocxReport(reportPath);
  }

  return {
    loadRules,
    buildExportableRulesContent,
    parseReport
  };
}

module.exports = {
  createReportSource
};
