require('./runtimePolyfills');

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
const DEFAULT_RULES_RELATIVE_PATH = path.join(
  'src',
  'renderer',
  'modules',
  'testDataExtraction',
  'config',
  'moto_rules_for_analysis.json5'
);
const wordExtractor = new WordExtractor();

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
  defaultRulesRelativePath: DEFAULT_RULES_RELATIVE_PATH,
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
  buildExportableRulesContent,
  parseChecklistReportOptions,
  inspectReport,
  parseReport
};
