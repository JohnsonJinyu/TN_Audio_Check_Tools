const fs = require('fs/promises');
const path = require('path');
const { clampMaxConcurrentTasks } = require('./concurrency');

function buildTimestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');
  return [date.getFullYear(), pad(date.getMonth() + 1), pad(date.getDate())].join('') +
    '_' +
    [pad(date.getHours()), pad(date.getMinutes())].join('');
}

function buildSharedChecklistOutputPath(checklistPath, outputDirectory, reportPaths) {
  const checklistName = path.parse(checklistPath || '').name;
  const parentFolderName = path.basename(path.dirname(checklistPath || ''));
  const normalizedFolderName = String(parentFolderName || '').trim();
  const fileNameBase = [checklistName, normalizedFolderName, 'merged']
    .filter(Boolean)
    .join('_');

  return path.join(
    outputDirectory,
    `${fileNameBase || 'checklist'}_${buildTimestamp()}.xlsx`
  );
}

function emitProgress(onProgress, payload) {
  if (typeof onProgress !== 'function') {
    return;
  }

  onProgress(payload);
}

function createReportRunner({
  supportedChecklistExtensions,
  resolveDefaultRulePath,
  loadRules,
  processSingleReport,
  buildBatchConclusion
}) {
  async function validatePaths({ reportPaths, checklistPath, rulePath }) {
    if (!Array.isArray(reportPaths) || reportPaths.length === 0) {
      throw new Error('请先选择至少一个测试报告');
    }

    const hasExcelReports = reportPaths.some((reportPath) => ['.xlsx', '.xls'].includes(path.extname(reportPath || '').toLowerCase()));
    if (hasExcelReports && !checklistPath) {
      throw new Error('存在 Excel 报告时，必须提供 checklist 文件。');
    }

    if (checklistPath) {
      const checklistExtension = path.extname(checklistPath || '').toLowerCase();
      if (!supportedChecklistExtensions.has(checklistExtension)) {
        throw new Error('checklist 仅支持 .xlsx 或 .xls 文件');
      }

      await fs.access(checklistPath);
    }

    await Promise.all(reportPaths.map((reportPath) => fs.access(reportPath)));

    if (rulePath) {
      await fs.access(rulePath);
    }
  }

  async function resolveRulePath(appPath, customRulePath) {
    if (customRulePath) {
      return customRulePath;
    }

    if (typeof resolveDefaultRulePath === 'function') {
      return resolveDefaultRulePath(appPath);
    }

    throw new Error('缺少默认规则路径解析器');
  }

  // 这一层只做流程编排，不关心具体的报告解析和提取细节。
  async function processReports({
    reportPaths,
    checklistPath,
    rulePath,
    customer,
    reportPanelSelections,
    reportPanelSelectionsByPath,
    maxConcurrentTasks,
    outputDirectory,
    appPath,
    onProgress
  }) {
    const resolvedRulePath = await resolveRulePath(appPath, rulePath);
    await validatePaths({ reportPaths, checklistPath, rulePath: resolvedRulePath });

    const rules = await loadRules(resolvedRulePath);
    const results = new Array(reportPaths.length);
    const total = reportPaths.length;
    const resolvedOutputDirectory = String(outputDirectory || '').trim() || path.dirname(reportPaths[0]);
    const sharedChecklistOutputPath = checklistPath
      ? buildSharedChecklistOutputPath(checklistPath, resolvedOutputDirectory, reportPaths)
      : '';
    const shouldMergeIntoSingleChecklist = Boolean(checklistPath && reportPaths.length > 0);
    const resolvedMaxConcurrentTasks = shouldMergeIntoSingleChecklist
      ? 1
      : Math.min(total, clampMaxConcurrentTasks(maxConcurrentTasks, 1));
    let completed = 0;
    let successCount = 0;

    emitProgress(onProgress, {
      type: 'batch-start',
      total,
      completed: 0,
      successCount: 0,
      errorCount: 0
    });

    let nextIndex = 0;

    async function processReportAtIndex(reportIndex) {
      const reportPath = reportPaths[reportIndex];
      let resultEntry;

      try {
        const result = await processSingleReport({
          reportPath,
          checklistPath,
          rules,
          customer,
          reportPanelSelections,
          reportPanelSelectionsOverride: reportPanelSelectionsByPath?.[reportPath] || null,
          outputDirectory: resolvedOutputDirectory,
          checklistWriteOptions: shouldMergeIntoSingleChecklist
            ? {
              outputPath: sharedChecklistOutputPath,
              reuseExistingOutput: completed > 0,
              skipReportSheetUpdates: completed > 0
            }
            : null
        });
        resultEntry = { status: 'success', ...result };
        successCount += 1;
      } catch (error) {
        resultEntry = {
          status: 'error',
          reportPath,
          error: error.message || '报告处理失败'
        };
      }

      results[reportIndex] = resultEntry;
      completed += 1;

      emitProgress(onProgress, {
        type: 'report-complete',
        total,
        completed,
        successCount,
        errorCount: completed - successCount,
        result: resultEntry
      });
    }

    async function workerLoop() {
      while (true) {
        const currentIndex = nextIndex;
        nextIndex += 1;

        if (currentIndex >= total) {
          return;
        }

        await processReportAtIndex(currentIndex);
      }
    }

    await Promise.all(
      Array.from({ length: resolvedMaxConcurrentTasks }, () => workerLoop())
    );

    emitProgress(onProgress, {
      type: 'batch-complete',
      total,
      completed,
      successCount,
      errorCount: completed - successCount
    });

    return {
      rulePath: resolvedRulePath,
      results,
      conclusion: typeof buildBatchConclusion === 'function'
        ? buildBatchConclusion({ results, checklistPath })
        : null
    };
  }

  return {
    processReports
  };
}

module.exports = {
  createReportRunner
};
