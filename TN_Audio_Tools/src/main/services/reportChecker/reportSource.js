const fs = require('fs/promises');
const path = require('path');
const mammoth = require('mammoth');
const JSON5 = require('json5');
const { parseDocxStructuredData } = require('./docxStructuredParser');

function createReportSource({
  supportedReportExtensions,
  convertDocToTemporaryDocx,
  wordExtractor,
  createSearchData
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

    return createSearchData(rawTextResult.value || '', htmlResult.value || '', structuredData);
  }

  // 解析入口只负责拿到标准化的搜索数据，不参与后续提取规则判断。
  async function parseReport(reportPath) {
    const reportExtension = path.extname(reportPath).toLowerCase();
    if (!supportedReportExtensions.has(reportExtension)) {
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

  return {
    loadRules,
    parseReport
  };
}

module.exports = {
  createReportSource
};
