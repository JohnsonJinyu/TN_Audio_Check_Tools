const path = require('path');
const fs = require('fs/promises');
const createWordReviewService = require('./wordReviewService');
const { parseReport } = require('../testDataExtraction');
const {
  checkSameCodecDifferentNetworkLoudness,
  checkSameNetworkDifferentCodecLoudness,
} = require('./checks/contentConsistency');
const { buildTestDataFacts } = require('./reportTestDataFacts');
const { determineOverallStatus } = require('./utils');

async function resolveEffectivePath(reportPath, reportData) {
  // 优先使用 COM 转换后的临时 .docx 路径
  if (reportData?._convertedDocxPath) {
    try { await fs.access(reportData._convertedDocxPath); return reportData._convertedDocxPath; } catch (_) {}
  }
  if (path.extname(reportPath).toLowerCase() !== '.doc') return reportPath;
  if (reportData.reportFormat !== 'docx') return reportPath;
  var docxPath = reportPath.replace(/\.doc$/i, '.docx');
  try { await fs.access(docxPath); return docxPath; } catch (_) { return reportPath; }
}

async function reviewWordReport(reportPath) {
  if (!reportPath) {
    throw new Error('缺少报告路径');
  }

  const reportData = await parseReport(reportPath);

  if (!reportData || !reportData.reportFormat) {
    throw new Error('无法解析报告文件');
  }

  const effectivePath = await resolveEffectivePath(reportPath, reportData);

  const wordReviewService = createWordReviewService();
  const reviewResult = await wordReviewService.reviewWordReport(effectivePath, reportData);
  const report = wordReviewService.generateReviewReport(reviewResult);

  return {
    reportPath,
    reviewResult,
    report,
    timestamp: new Date().toISOString()
  };
}

/**
 * 跨报告对比检查
 * 在批量审查后调用，对多份报告进行跨报告一致性检查（2.2.1, 2.2.2）
 */
function runCrossReportChecks(reviewResults) {
  if (!Array.isArray(reviewResults) || reviewResults.length < 2) {
    return reviewResults || [];
  }

  // 收集所有报告的响度数据
  const allMetrics = reviewResults.map((r) => {
    const tf = r.reviewResult?.testDataFacts;
    const ctx = r.reviewResult?.reportFacts?.metadata || {};
    return {
      reportPath: r.reportPath || r.docxPath || '',
      metrics: tf?.loudnessMetrics || [],
      codec: ctx.codec || '',
      network: ctx.bandwidth || ctx.network || '',
    };
  });

  // 执行跨报告检查
  const codecDiffResult = checkSameCodecDifferentNetworkLoudness(allMetrics);
  const networkDiffResult = checkSameNetworkDifferentCodecLoudness(allMetrics);

  // 合并结果到每份报告
  return reviewResults.map((r) => {
    r.reviewResult.checks.contentSameCodecDiffNetwork = codecDiffResult;
    r.reviewResult.checks.contentSameNetworkDiffCodec = networkDiffResult;

    // 重算 summary
    const summary = { ...r.reviewResult.summary };
    const oldCodec = r.reviewResult.checks.contentSameCodecDiffNetwork?.status || 'review';
    const oldNetwork = r.reviewResult.checks.contentSameNetworkDiffCodec?.status || 'review';
    // 用新结果替换旧的review占位
    if (oldCodec === 'review' && codecDiffResult.status !== 'review') {
      summary.reviewChecks -= 1;
      if (codecDiffResult.status === 'pass') summary.passedChecks += 1;
      else if (codecDiffResult.status === 'warning') summary.warningChecks += 1;
      else if (codecDiffResult.status === 'error') summary.errorChecks += 1;
    }
    if (oldNetwork === 'review' && networkDiffResult.status !== 'review') {
      summary.reviewChecks -= 1;
      if (networkDiffResult.status === 'pass') summary.passedChecks += 1;
      else if (networkDiffResult.status === 'warning') summary.warningChecks += 1;
      else if (networkDiffResult.status === 'error') summary.errorChecks += 1;
    }
    r.reviewResult.summary = summary;
    r.reviewResult.overallStatus = determineOverallStatus(summary);
    r.report = createWordReviewService().generateReviewReport(r.reviewResult);

    return r;
  });
}

/**
 * 配对审查：合并 .docx（文档结构）和 .xlsx（测试数据）进行全文审查
 */
async function reviewPairedReport(docxPath, xlsxPath) {
  if (!docxPath || !xlsxPath) {
    throw new Error('配对审查需要同时提供 .docx 和 .xlsx 文件路径');
  }

  const docxData = await parseReport(docxPath);
  const xlsxData = await parseReport(xlsxPath);

  const effectiveDocxPath = await resolveEffectivePath(docxPath, docxData);

  const mergedData = {
    ...docxData,
    reportFormat: 'paired',
    detailedRows: xlsxData.detailedRows || [],
    xlsxReportContext: xlsxData.reportContext || {},
    pairedDocxPath: effectiveDocxPath,
    pairedXlsxPath: xlsxPath,
  };

  const wordReviewService = createWordReviewService();
  const reviewResult = await wordReviewService.reviewWordReport(effectiveDocxPath, mergedData);
  const report = wordReviewService.generateReviewReport(reviewResult);

  return {
    docxPath,
    xlsxPath,
    reviewResult,
    report,
    timestamp: new Date().toISOString(),
  };
}

module.exports = {
  createWordReviewService,
  reviewWordReport,
  reviewPairedReport,
  runCrossReportChecks,
};
