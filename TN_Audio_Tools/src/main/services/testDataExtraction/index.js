require('./runtimePolyfills');

const fs = require('fs/promises');
const path = require('path');
const WordExtractor = require('word-extractor');
const { applyResultsToChecklist } = require('./checklistWriter');
const { createReportRunner } = require('./reportRunner');
const { createReportSource } = require('./reportSource');
const { createReportExtractor } = require('./reportExtractor');
const { createReportConverter } = require('./reportConverter');
const { createXlsxReportSource } = require('./xlsxReportSource');
const { parseChecklistReportOptions } = require('./checklistReportPanel');
const {
  analyzeExcelReport,
  analyzeWordReport,
  buildBatchConclusion
} = require('./reportConclusion');
const {
  createSearchData,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
} = require('./reportAnalysis');

const SUPPORTED_REPORT_EXTENSIONS = new Set(['.doc', '.docx', '.xlsx', '.xls']);
const SUPPORTED_CHECKLIST_EXTENSIONS = new Set(['.xlsx', '.xls']);
const RULES_CONFIG_RELATIVE_DIR = path.join(
  'src',
  'renderer',
  'modules',
  'testDataExtraction',
  'config'
);
const DEFAULT_RULES_RELATIVE_PATH = path.join(
  RULES_CONFIG_RELATIVE_DIR,
  'moto_rules_for_analysis.json5'
);
const wordExtractor = new WordExtractor();

async function resolveBundledRulesPath(appPath) {
  const candidatePaths = [
    appPath ? path.join(appPath, DEFAULT_RULES_RELATIVE_PATH) : '',
    process.resourcesPath
      ? path.join(process.resourcesPath, 'app.asar.unpacked', DEFAULT_RULES_RELATIVE_PATH)
      : '',
    process.resourcesPath
      ? path.join(process.resourcesPath, DEFAULT_RULES_RELATIVE_PATH)
      : '',
    path.resolve(__dirname, '..', '..', '..', 'renderer', 'modules', 'testDataExtraction', 'config', 'moto_rules_for_analysis.json5')
  ].filter(Boolean);

  for (const candidatePath of candidatePaths) {
    try {
      await fs.access(candidatePath);
      return candidatePath;
    } catch (error) {
      // Try next candidate path.
    }
  }

  throw new Error(`内置规则文件不存在：${candidatePaths.join(' | ')}`);
}

const { convertDocToTemporaryDocx } = createReportConverter({
  wordExtractor
});

const { parseXlsxReport } = createXlsxReportSource();

const { loadRules, buildExportableRulesContent, parseReport } = createReportSource({
  supportedReportExtensions: SUPPORTED_REPORT_EXTENSIONS,
  convertDocToTemporaryDocx,
  wordExtractor,
  createSearchData,
  parseXlsxReport
});

const { processSingleReport, inspectReportContext } = createReportExtractor({
  parseReport,
  applyResultsToChecklist,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue,
  analyzeExcelReport,
  analyzeWordReport
});

const { processReports } = createReportRunner({
  supportedChecklistExtensions: SUPPORTED_CHECKLIST_EXTENSIONS,
  resolveDefaultRulePath: resolveBundledRulesPath,
  loadRules,
  processSingleReport,
  buildBatchConclusion
});

async function inspectReport(reportPath, options = {}) {
  if (!reportPath) {
    throw new Error('缺少报告路径');
  }

  const reportData = await parseReport(reportPath);
  const mergedReportContext = inspectReportContext(reportData?.reportContext || {}, options);

  return {
    reportPath,
    reportFormat: reportData?.reportFormat || '',
    reportContext: mergedReportContext,
    suggestedReportPanelSelections: mergedReportContext.reportPanelSelections || null
  };
}

module.exports = {
  processReports,
  DEFAULT_RULES_RELATIVE_PATH,
  resolveBundledRulesPath,
  buildExportableRulesContent,
  parseChecklistReportOptions,
  inspectReport,
  parseReport
};
