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

function deriveReportMetadata(reportPath, rawText, reportData) {
  const reportName = path.parse(reportPath || '').name;
  const combinedSource = `${reportName} ${rawText || ''}`;
  const reportContext = reportData?.reportContext || {};

  return {
    reportName,
    measurementObject: String(reportContext.measurementObject || '').trim() || deriveMeasurementObject(rawText) || reportName,
    bandwidth: normalizeReportBandwidth(reportContext.bandwidth) || deriveBandwidthFromPath(reportPath) || deriveBandwidthFromText(rawText) || '',
    codec: String(reportContext.codec || '').trim().toUpperCase() || deriveTokenByCandidates(combinedSource, ['EVS', 'AMR']),
    network: String(reportContext.network || '').trim().toUpperCase() || deriveTokenByCandidates(combinedSource, ['VOLTE', 'VOWIFI', 'VONR', 'VOIP', 'WCDMA', 'GSM']),
    terminalMode: String(reportContext.terminalMode || '').trim().toUpperCase() || deriveTokenByCandidates(combinedSource, ['HA', 'HF', 'HS', 'HE', 'HH'])
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
  async function findSiblingWordReport(reportPath) {
    const parsedPath = path.parse(reportPath);
    const candidates = [
      path.join(parsedPath.dir, `${parsedPath.name}.docx`),
      path.join(parsedPath.dir, `${parsedPath.name}.doc`)
    ];

    for (const candidatePath of candidates) {
      if (candidatePath.toLowerCase() === String(reportPath || '').toLowerCase()) {
        continue;
      }

      try {
        await fs.access(candidatePath);
        return candidatePath;
      } catch {
        // Ignore missing sibling files.
      }
    }

    return '';
  }

  function mergeXlsxAndWordReportData(xlsxData, wordData) {
    if (!wordData) {
      return xlsxData;
    }

    return {
      ...xlsxData,
      rawText: [xlsxData.rawText, wordData.rawText].filter(Boolean).join('\n'),
      html: wordData.html || xlsxData.html || '',
      lines: Array.from(new Set([...(xlsxData.lines || []), ...(wordData.lines || [])])),
      ambientNoiseBlocks: (xlsxData.ambientNoiseBlocks && xlsxData.ambientNoiseBlocks.length > 0)
        ? xlsxData.ambientNoiseBlocks
        : (wordData.ambientNoiseBlocks || []),
      tables: [...(xlsxData.tables || []), ...(wordData.tables || [])],
      tableRows: [...(xlsxData.tableRows || []), ...(wordData.tableRows || [])],
      fallbackRows: [...(xlsxData.fallbackRows || []), ...(wordData.fallbackRows || [])],
      structuredData: wordData.structuredData || xlsxData.structuredData,
      auxiliarySources: [
        ...(xlsxData.auxiliarySources || []),
        {
          type: 'word-report',
          reportFormat: wordData.reportFormat || 'docx',
          reportPath: wordData.reportContext?.reportName || ''
        }
      ]
    };
  }

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
        normalizedProfiles[profileKey] = await loadRules(path.resolve(path.dirname(rulePath), profileConfig.rulePath));
        continue;
      }

      normalizedProfiles[profileKey] = await normalizeRulesConfig(rulePath, profileConfig);
    }

    return {
      ruleBaseInfo: rules.ruleBaseInfo || {},
      defaultProfileKey: String(rules.defaultProfileKey || profileEntries[0][0]).trim() || profileEntries[0][0],
      ruleProfiles: normalizedProfiles
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
      parseDocxStructuredData(reportPath).catch(() => ({ lines: [], tables: [], headers: [], footers: [] }))
    ]);

    const searchData = createSearchData(rawTextResult.value || '', htmlResult.value || '', structuredData);

    return attachReportContext(
      {
        ...searchData,
        reportFormat: 'docx',
        structuredData
      },
      reportPath,
      rawTextResult.value || ''
    );
  }

  // 解析入口只负责拿到标准化的搜索数据，不参与后续提取规则判断。
  async function parseReport(reportPath) {
    const reportExtension = path.extname(reportPath).toLowerCase();
    if (!supportedReportExtensions.has(reportExtension)) {
      throw new Error('当前仅支持 .xlsx / .xls / .doc / .docx 测试报告');
    }

    if (reportExtension === '.xlsx' || reportExtension === '.xls') {
      let xlsxData = await parseXlsxReport(reportPath);

      if (!Array.isArray(xlsxData?.ambientNoiseBlocks) || xlsxData.ambientNoiseBlocks.length === 0) {
        const siblingWordReport = await findSiblingWordReport(reportPath);
        if (siblingWordReport) {
          try {
            const wordData = await parseReport(siblingWordReport);
            xlsxData = mergeXlsxAndWordReportData(xlsxData, wordData);
          } catch {
            // Keep xlsx-only parsing path when sibling Word parsing fails.
          }
        }
      }

      return attachReportContext(xlsxData, reportPath);
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
        throw new Error('.doc 报告未读取到有效文本内容。请优先另存为 .docx 后重试。');
      }

      return attachReportContext({
        ...createSearchData(rawText, ''),
        reportFormat: 'doc',
        structuredData: {
          lines: [],
          tables: [],
          headers: extracted.getHeaders?.() ? [extracted.getHeaders()] : [],
          footers: []
        }
      }, reportPath, rawText);
    }

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
