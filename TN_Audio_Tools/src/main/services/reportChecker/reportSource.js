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

function attachReportContext(reportData, reportPath, rawText = '') {
  const existingBandwidth = normalizeReportBandwidth(reportData?.reportContext?.bandwidth);
  const bandwidth = existingBandwidth || deriveBandwidthFromPath(reportPath) || deriveBandwidthFromText(rawText) || '';

  return {
    ...reportData,
    reportContext: {
      ...(reportData?.reportContext || {}),
      bandwidth
    }
  };
}

function createReportSource({
  supportedReportExtensions,
  convertDocToTemporaryDocx,
  wordExtractor,
  createSearchData,
  parseXlsxReport
}) {
  async function loadRules(rulePath) {
    const content = await fs.readFile(rulePath, 'utf8');
    const rules = JSON5.parse(content);

    if (!Array.isArray(rules.extractItemList)) {
      throw new Error('规则文件缺少 extractItemList 配置');
    }

    return rules;
  }

  async function parseDocxReport(reportPath) {
    const [rawTextResult, htmlResult, structuredData] = await Promise.all([
      mammoth.extractRawText({ path: reportPath }),
      mammoth.convertToHtml({ path: reportPath }),
      parseDocxStructuredData(reportPath).catch(() => ({ lines: [], tables: [] }))
    ]);

    return attachReportContext(
      createSearchData(rawTextResult.value || '', htmlResult.value || '', structuredData),
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
      return attachReportContext(await parseXlsxReport(reportPath), reportPath);
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

      return attachReportContext(createSearchData(rawText, ''), reportPath, rawText);
    }

    return parseDocxReport(reportPath);
  }

  return {
    loadRules,
    parseReport
  };
}

module.exports = {
  createReportSource
};
