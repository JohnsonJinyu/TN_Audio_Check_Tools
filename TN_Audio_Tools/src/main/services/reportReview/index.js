const createWordReviewService = require('./wordReviewService');
const { parseReport } = require('../testDataExtraction');

async function reviewWordReport(reportPath) {
  if (!reportPath) {
    throw new Error('缺少报告路径');
  }

  const reportData = await parseReport(reportPath);

  if (!reportData || !reportData.reportFormat) {
    throw new Error('无法解析报告文件');
  }

  if (reportData.reportFormat !== 'docx' && reportData.reportFormat !== 'doc') {
    throw new Error('只支持 Word 报告审查，请提供 .doc 或 .docx 文件');
  }

  const wordReviewService = createWordReviewService();
  const reviewResult = await wordReviewService.reviewWordReport(reportPath, reportData);
  const report = wordReviewService.generateReviewReport(reviewResult);

  return {
    reportPath,
    reviewResult,
    report,
    timestamp: new Date().toISOString()
  };
}

module.exports = {
  createWordReviewService,
  reviewWordReport
};
