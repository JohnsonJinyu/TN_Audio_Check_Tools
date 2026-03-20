const fs = require('fs/promises');
const path = require('path');

function emitProgress(onProgress, payload) {
  if (typeof onProgress !== 'function') {
    return;
  }

  onProgress(payload);
}

function createReportRunner({
  supportedChecklistExtensions,
  defaultRulesRelativePath,
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

  function resolveRulePath(appPath, customRulePath) {
    if (customRulePath) {
      return customRulePath;
    }

    return path.join(appPath, defaultRulesRelativePath);
  }

  // 这一层只做流程编排，不关心具体的报告解析和提取细节。
  async function processReports({
    reportPaths,
    checklistPath,
    rulePath,
    customer,
    reportPanelSelections,
    reportPanelSelectionsByPath,
    appPath,
    onProgress
  }) {
    const resolvedRulePath = resolveRulePath(appPath, rulePath);
    await validatePaths({ reportPaths, checklistPath, rulePath: resolvedRulePath });

    const rules = await loadRules(resolvedRulePath);
    const results = [];
    const total = reportPaths.length;

    emitProgress(onProgress, {
      type: 'batch-start',
      total,
      completed: 0,
      successCount: 0,
      errorCount: 0
    });

    for (const reportPath of reportPaths) {
      let resultEntry;

      try {
        const result = await processSingleReport({
          reportPath,
          checklistPath,
          rules,
          customer,
          reportPanelSelections,
          reportPanelSelectionsOverride: reportPanelSelectionsByPath?.[reportPath] || null
        });
        resultEntry = { status: 'success', ...result };
      } catch (error) {
        resultEntry = {
          status: 'error',
          reportPath,
          error: error.message || '报告处理失败'
        };
      }

      results.push(resultEntry);
      const completed = results.length;
      const successCount = results.filter((item) => item.status === 'success').length;
      const errorCount = completed - successCount;

      emitProgress(onProgress, {
        type: 'report-complete',
        total,
        completed,
        successCount,
        errorCount,
        result: resultEntry
      });
    }

    emitProgress(onProgress, {
      type: 'batch-complete',
      total,
      completed: results.length,
      successCount: results.filter((item) => item.status === 'success').length,
      errorCount: results.filter((item) => item.status === 'error').length
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
