const {
  buildWordData,
  updateSummary,
  determineOverallStatus
} = require('./utils');
const { buildReviewFacts } = require('./reportFacts');
const {
  extractTableOfContents,
  checkTableOfContentsPages,
  checkChaptersAlignment
} = require('./checks/documentStructure');
const {
  checkReportBasicInfo,
  checkTestItemConsistency,
  checkNamePollution,
  findEngineerNames
} = require('./checks/metadata');
const { checkPolqaConfiguration } = require('./checks/polqa');
const { generateReviewReport } = require('./reportBuilder');

/**
 * 综合审查 Word 报告
 */
async function reviewWordReport(reportPath, reportData) {
  if (!reportPath) {
    throw new Error('缺少报告路径');
  }

  if (!reportData) {
    throw new Error('缺少报告数据');
  }

  const wordData = buildWordData(reportData);
  const reviewFacts = buildReviewFacts(reportPath, wordData);

  const allResults = {};
  const summary = {
    totalChecks: 8,
    passedChecks: 0,
    warningChecks: 0,
    reviewChecks: 0,
    errorChecks: 0
  };

  // 1. 提取目录
  const tocInfo = extractTableOfContents(wordData);
  allResults.tableOfContents = tocInfo;
  updateSummary(summary, tocInfo.chapters.length > 0 || tocInfo.tocLines?.length > 0 ? 'pass' : 'review');

  // 2. 检查目录页数
  const tocPagesResult = checkTableOfContentsPages(wordData, tocInfo);
  allResults.tableOfContentsPages = tocPagesResult;
  updateSummary(summary, tocPagesResult.status);

  // 3. 检查章节对应
  const chaptersAlignmentResult = checkChaptersAlignment(wordData, tocInfo);
  allResults.chaptersAlignment = chaptersAlignmentResult;
  updateSummary(summary, chaptersAlignmentResult.status);

  // 4. 检查基本信息
  const basicInfoResult = checkReportBasicInfo(reportPath, wordData, reviewFacts);
  allResults.basicInfo = basicInfoResult;
  updateSummary(summary, basicInfoResult.status);

  // 5. 检查测试项一致性
  const testItemResult = checkTestItemConsistency(reportPath, wordData, reviewFacts);
  allResults.testItemConsistency = testItemResult;
  updateSummary(summary, testItemResult.status);

  // 6. 检查名称污染
  const pollutionResult = checkNamePollution(reportPath);
  allResults.namePollution = pollutionResult;
  updateSummary(summary, pollutionResult.status);

  // 7. 查找人员信息
  const engineersResult = findEngineerNames(wordData, reviewFacts);
  allResults.engineers = engineersResult;
  updateSummary(summary, engineersResult.status);

  // 8. 检查 POLQA 配置
  const polqaResult = checkPolqaConfiguration(wordData, reviewFacts);
  allResults.polqa = polqaResult;
  updateSummary(summary, polqaResult.status);

  return {
    reportPath,
    reviewTimestamp: new Date().toISOString(),
    summary,
    checks: allResults,
    overallStatus: determineOverallStatus(summary)
  };
}

function createWordReviewService() {
  return {
    reviewWordReport,
    generateReviewReport,
    extractTableOfContents,
    checkTableOfContentsPages,
    checkChaptersAlignment,
    checkReportBasicInfo,
    checkTestItemConsistency,
    checkNamePollution,
    findEngineerNames,
    checkPolqaConfiguration
  };
}

module.exports = createWordReviewService;
