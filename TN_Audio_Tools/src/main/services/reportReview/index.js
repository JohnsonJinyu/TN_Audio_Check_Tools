const path = require('path');
const fs = require('fs/promises');
const createWordReviewService = require('./wordReviewService');
const { parseReport } = require('../testDataExtraction');
const progressBus = require('./progressBus');
const {
  checkSameCodecDifferentNetworkLoudness,
  checkSameNetworkDifferentCodecLoudness,
} = require('./checks/contentConsistency');
const { buildTestDataFacts } = require('./reportTestDataFacts');
const { determineOverallStatus } = require('./utils');

function isLegacyDocPath(reportPath) {
  const lowerPath = String(reportPath || '').toLowerCase();
  return lowerPath.endsWith('.doc') && !lowerPath.endsWith('.docx');
}

function isXlsxPath(reportPath) {
  return /\.(xlsx|xls)$/i.test(String(reportPath || ''));
}

function buildReviewSteps(options) {
  const reportPath = options?.reportPath || '';
  const xlsxPath = options?.xlsxPath || '';
  const isPaired = Boolean(xlsxPath);
  const isXlsxOnly = !isPaired && isXlsxPath(reportPath);
  const steps = [
    { id: 'identify', label: '识别文件与审查模式' },
  ];

  if (isLegacyDocPath(reportPath)) {
    steps.push({ id: 'convert', label: '格式转换' });
  }

  if (isPaired || !isXlsxOnly) {
    steps.push({ id: 'parse-word', label: '解析 Word 报告' });
  }

  if (isPaired || isXlsxOnly) {
    steps.push({ id: 'parse-xlsx', label: '解析 xlsx 测试数据' });
  }

  steps.push({ id: 'extract-facts', label: '提取审查事实与上下文' });

  if (!isXlsxOnly) {
    steps.push({ id: 'structure-metadata', label: '执行文档结构与元数据检查' });
  }

  steps.push({ id: 'timing', label: '执行时序检查' });
  steps.push({ id: 'curve-values', label: '执行曲线与数值检查' });

  if (!isXlsxOnly) {
    steps.push({ id: 'chart-prepare', label: '提取并整理图表' });
    steps.push({ id: 'chart-analyze', label: '上传并分析图表批次' });
  }

  steps.push({ id: 'summarize', label: '汇总结果并生成报告' });
  return steps;
}

function createReviewProgressController(options) {
  const groupLabel = options?.groupLabel || '';
  const mode = options?.mode || 'single';
  const steps = buildReviewSteps(options);

  function emitStep(stepId, detail, status = 'running', extra = {}) {
    const stepIndex = steps.findIndex((step) => step.id === stepId);
    if (stepIndex < 0) return;

    const step = steps[stepIndex];
    progressBus.emitReviewProgress({
      mode,
      groupLabel,
      stepId,
      stepLabel: step.label,
      stepIndex: stepIndex + 1,
      totalSteps: steps.length,
      percent: status === 'done' && stepIndex === steps.length - 1
        ? 100
        : Math.round(((stepIndex + 1) / steps.length) * 100),
      detail: detail || step.label,
      status,
      ...extra,
    });
  }

  return {
    steps,
    emitStep,
  };
}

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

  const progressController = createReviewProgressController({
    reportPath,
    groupLabel: path.parse(reportPath).name,
    mode: 'single',
  });

  const parseStepId = isXlsxPath(reportPath) ? 'parse-xlsx' : 'parse-word';
  progressController.emitStep('identify', '正在识别文件格式与审查模式');
  progressController.emitStep(parseStepId, '正在准备解析报告');

  const reportData = await parseReport(reportPath, {
    onProgress(progress) {
      const detail = progress?.detail || '';
      if (isLegacyDocPath(reportPath) && /转换|重建/.test(detail)) {
        progressController.emitStep('convert', detail, progress?.status || 'running');
        return;
      }
      progressController.emitStep(parseStepId, detail || '正在解析报告', progress?.status || 'running');
    }
  });

  if (!reportData || !reportData.reportFormat) {
    throw new Error('无法解析报告文件');
  }

  const effectivePath = await resolveEffectivePath(reportPath, reportData);

  const wordReviewService = createWordReviewService();
  const reviewResult = await wordReviewService.reviewWordReport(effectivePath, reportData, {
    progressController
  });
  progressController.emitStep('summarize', '正在汇总检查结果并生成报告');
  const report = wordReviewService.generateReviewReport(reviewResult);
  progressController.emitStep('summarize', '当前报告组审查完成', 'done');

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

  const progressController = createReviewProgressController({
    reportPath: docxPath,
    xlsxPath,
    groupLabel: path.parse(docxPath).name,
    mode: 'paired',
  });

  progressController.emitStep('identify', '正在识别配对报告与审查模式');
  progressController.emitStep('parse-word', '正在准备解析 Word 报告');
  const docxData = await parseReport(docxPath, {
    onProgress(progress) {
      const detail = progress?.detail || '';
      if (isLegacyDocPath(docxPath) && /转换|重建/.test(detail)) {
        progressController.emitStep('convert', detail, progress?.status || 'running');
        return;
      }
      progressController.emitStep('parse-word', detail || '正在解析 Word 报告', progress?.status || 'running');
    }
  });
  progressController.emitStep('parse-xlsx', '正在解析配对 xlsx 测试数据');
  const xlsxData = await parseReport(xlsxPath, {
    onProgress(progress) {
      progressController.emitStep('parse-xlsx', progress?.detail || '正在解析 xlsx 测试数据', progress?.status || 'running');
    }
  });

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
  const reviewResult = await wordReviewService.reviewWordReport(effectiveDocxPath, mergedData, {
    progressController
  });
  progressController.emitStep('summarize', '正在汇总配对审查结果并生成报告');
  const report = wordReviewService.generateReviewReport(reviewResult);
  progressController.emitStep('summarize', '当前报告组审查完成', 'done');

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
