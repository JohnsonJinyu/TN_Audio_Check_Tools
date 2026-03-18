require('./reportChecker/runtimePolyfills');

const path = require('path');
const { applyResultsToChecklist } = require('./reportChecker/checklistWriter');
const { createReportRunner } = require('./reportChecker/reportRunner');
const { createReportSource } = require('./reportChecker/reportSource');
const { createReportExtractor } = require('./reportChecker/reportExtractor');
const { createReportConverter } = require('./reportChecker/reportConverter');
const { createXlsxReportSource } = require('./reportChecker/xlsxReportSource');
const {
  analyzeExcelReport,
  analyzeWordReport,
  buildBatchConclusion
} = require('./reportChecker/reportConclusion');
const {
  createSearchData,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
} = require('./reportChecker/reportAnalysis');
const WordExtractor = require('word-extractor');

const SUPPORTED_REPORT_EXTENSIONS = new Set(['.doc', '.docx', '.xlsx', '.xls']);
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

const { convertDocToTemporaryDocx } = createReportConverter({
  childProcessTimeoutMs: CHILD_PROCESS_TIMEOUT_MS,
  libreOfficeCandidatePaths: LIBRE_OFFICE_CANDIDATE_PATHS
});

const { parseXlsxReport } = createXlsxReportSource();

const { loadRules, buildExportableRulesContent, parseReport } = createReportSource({
  supportedReportExtensions: SUPPORTED_REPORT_EXTENSIONS,
  convertDocToTemporaryDocx,
  wordExtractor,
  createSearchData,
  parseXlsxReport
});

const { processSingleReport } = createReportExtractor({
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

module.exports = {
  processReports,
  DEFAULT_RULES_RELATIVE_PATH,
  buildExportableRulesContent
};