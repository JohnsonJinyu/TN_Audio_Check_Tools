const path = require('path');
const { applyResultsToChecklist } = require('./reportChecker/checklistWriter');
const { createReportRunner } = require('./reportChecker/reportRunner');
const { createReportSource } = require('./reportChecker/reportSource');
const { createReportExtractor } = require('./reportChecker/reportExtractor');
const { createReportConverter } = require('./reportChecker/reportConverter');
const {
  createSearchData,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
} = require('./reportChecker/reportAnalysis');

if (typeof globalThis.File === 'undefined') {
  globalThis.File = class File {};
}
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

const { convertDocToTemporaryDocx } = createReportConverter({
  childProcessTimeoutMs: CHILD_PROCESS_TIMEOUT_MS,
  libreOfficeCandidatePaths: LIBRE_OFFICE_CANDIDATE_PATHS
});

const { loadRules, parseReport } = createReportSource({
  supportedReportExtensions: SUPPORTED_REPORT_EXTENSIONS,
  convertDocToTemporaryDocx,
  wordExtractor,
  createSearchData
});

const { processSingleReport } = createReportExtractor({
  parseReport,
  applyResultsToChecklist,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
});

const { processReports } = createReportRunner({
  supportedChecklistExtensions: SUPPORTED_CHECKLIST_EXTENSIONS,
  defaultRulesRelativePath: DEFAULT_RULES_RELATIVE_PATH,
  loadRules,
  processSingleReport
});

module.exports = {
  processReports,
  DEFAULT_RULES_RELATIVE_PATH
};