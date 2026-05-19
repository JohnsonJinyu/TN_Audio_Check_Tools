const {
  buildWordData,
  updateSummary,
  determineOverallStatus
} = require('./utils');
const { buildReviewFacts } = require('./reportFacts');
const { buildTestDataFacts } = require('./reportTestDataFacts');
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
const {
  checkAdjacentTestItemInterval,
  checkTotalTestSpan,
  checkDelayTestTiming,
  checkSidetoneDelayTiming,
  checkBgnConnectionTiming
} = require('./checks/timing');
const {
  checkLoudnessFrequencyResponseTrendConsistency,
  checkCurveValueCorroboration
} = require('./checks/contentConsistency');
const { generateReviewReport } = require('./reportBuilder');
const { getSettings } = require('../settingsService');

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
  const testDataFacts = buildTestDataFacts(reportData);

  const isXlsxOnly = wordData.reportFormat === 'xlsx';
  const isWordFormat = !isXlsxOnly;

  const allResults = {};
  const summary = {
    totalChecks: isXlsxOnly ? 9 : 17,
    passedChecks: 0,
    warningChecks: 0,
    reviewChecks: 0,
    errorChecks: 0
  };

  // === 文档结构检查（仅Word格式） ===
  if (isWordFormat) {
    // 1. 提取目录
    const tocInfo = extractTableOfContents(wordData);
    tocInfo.status = tocInfo.chapters.length > 0 || tocInfo.tocLines?.length > 0 ? 'pass' : 'review';
    allResults.tableOfContents = tocInfo;
    updateSummary(summary, tocInfo.status);

    // 2. 检查目录页数
    const tocPagesResult = checkTableOfContentsPages(wordData, tocInfo);
    allResults.tableOfContentsPages = tocPagesResult;
    updateSummary(summary, tocPagesResult.status);

    // 3. 检查章节对应
    const chaptersAlignmentResult = checkChaptersAlignment(wordData, tocInfo);
    allResults.chaptersAlignment = chaptersAlignmentResult;
    updateSummary(summary, chaptersAlignmentResult.status);
  } else {
    // xlsx格式不适用文档结构检查
    const docStructureSkipped = { issues: [], evidence: ['xlsx格式不适用文档结构检查'], status: 'pass' };
    allResults.tableOfContents = { chapters: [], tocLines: [], evidence: ['xlsx格式无目录结构'], status: 'pass' };
    allResults.tableOfContentsPages = docStructureSkipped;
    allResults.chaptersAlignment = docStructureSkipped;
  }

  // === 元数据与命名检查（仅Word格式） ===
  if (isWordFormat) {
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
  } else {
    const metadataSkipped = { issues: [], evidence: ['xlsx格式不适用元数据与命名检查'], status: 'pass' };
    allResults.basicInfo = metadataSkipped;
    allResults.testItemConsistency = metadataSkipped;
    allResults.namePollution = metadataSkipped;
    allResults.engineers = { issues: [], evidence: ['xlsx格式不含人员信息'], status: 'pass' };
    allResults.polqa = metadataSkipped;
  }

  // === 测试时间检查 (2.3) ===
  // 9. 相邻测试项间隔
  const timingAdjInterval = checkAdjacentTestItemInterval(testDataFacts);
  allResults.timingAdjacentInterval = timingAdjInterval;
  updateSummary(summary, timingAdjInterval.status);

  // 10. 全部测试项跨度
  const timingTotalSpan = checkTotalTestSpan(testDataFacts);
  allResults.timingTotalSpan = timingTotalSpan;
  updateSummary(summary, timingTotalSpan.status);

  // 11. 时延测试时序
  const timingDelayOrder = checkDelayTestTiming(testDataFacts);
  allResults.timingDelayOrder = timingDelayOrder;
  updateSummary(summary, timingDelayOrder.status);

  // 12. Sidetone Delay时序
  const timingSidetoneDelay = checkSidetoneDelayTiming(testDataFacts);
  allResults.timingSidetoneDelayOrder = timingSidetoneDelay;
  updateSummary(summary, timingSidetoneDelay.status);

  // 13. BGN Connection时序
  const timingBgnConnection = checkBgnConnectionTiming(testDataFacts);
  allResults.timingBgnConnectionOrder = timingBgnConnection;
  updateSummary(summary, timingBgnConnection.status);

  // === 内容合理性检查 (2.2) ===
  // 14. 响度与频响趋势一致性（单报告内）
  const contentLoudnessFR = await checkLoudnessFrequencyResponseTrendConsistency(testDataFacts, reportPath, getSettings().llm);
  allResults.contentLoudnessFRTrend = contentLoudnessFR;
  updateSummary(summary, contentLoudnessFR.status);

  // 15. 曲线与数值互相印证（单报告内）
  // 将LLM视觉分析结果传递给数值互证检查，实现交叉验证
  const llmContextForCorroboration = {
    rawFindings: contentLoudnessFR.rawFindings,
    monotonicityViolations: contentLoudnessFR.monotonicityViolations,
  };
  const contentCurveCorroboration = checkCurveValueCorroboration(testDataFacts, llmContextForCorroboration);
  allResults.contentCurveValueCorroboration = contentCurveCorroboration;
  updateSummary(summary, contentCurveCorroboration.status);

  // 16-17. 跨报告对比（单报告模式下标记为批量对比待执行）
  allResults.contentSameCodecDiffNetwork = {
    issues: [{ severity: 'review', message: '需要至少2份不同网络的报告进行跨报告对比（批量模式下自动执行）' }],
    evidence: ['跨报告检查在批量审查时自动触发'],
    status: 'review',
  };
  updateSummary(summary, 'review');

  allResults.contentSameNetworkDiffCodec = {
    issues: [{ severity: 'review', message: '需要至少2份不同codec的报告进行跨报告对比（批量模式下自动执行）' }],
    evidence: ['跨报告检查在批量审查时自动触发'],
    status: 'review',
  };
  updateSummary(summary, 'review');

  return {
    reportPath,
    reviewTimestamp: new Date().toISOString(),
    summary,
    checks: allResults,
    overallStatus: determineOverallStatus(summary),
    reportFacts: reviewFacts,
    testDataFacts,
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
    checkPolqaConfiguration,
    checkAdjacentTestItemInterval,
    checkTotalTestSpan,
    checkDelayTestTiming,
    checkSidetoneDelayTiming,
    checkBgnConnectionTiming,
    checkLoudnessFrequencyResponseTrendConsistency,
    checkCurveValueCorroboration,
  };
}

module.exports = createWordReviewService;
