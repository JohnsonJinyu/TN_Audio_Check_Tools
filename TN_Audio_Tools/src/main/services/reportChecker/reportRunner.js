const fs = require('fs/promises');
const path = require('path');

function createReportRunner({
  supportedChecklistExtensions,
  defaultRulesRelativePath,
  loadRules,
  processSingleReport
}) {
  async function validatePaths({ reportPaths, checklistPath, rulePath }) {
    if (!Array.isArray(reportPaths) || reportPaths.length === 0) {
      throw new Error('请先选择至少一个测试报告');
    }

    const checklistExtension = path.extname(checklistPath || '').toLowerCase();
    if (!supportedChecklistExtensions.has(checklistExtension)) {
      throw new Error('checklist 仅支持 .xlsx 或 .xls 文件');
    }

    await fs.access(checklistPath);
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
  async function processReports({ reportPaths, checklistPath, rulePath, appPath }) {
    const resolvedRulePath = resolveRulePath(appPath, rulePath);
    await validatePaths({ reportPaths, checklistPath, rulePath: resolvedRulePath });

    const rules = await loadRules(resolvedRulePath);
    const results = [];

    for (const reportPath of reportPaths) {
      try {
        const result = await processSingleReport({ reportPath, checklistPath, rules });
        results.push({ status: 'success', ...result });
      } catch (error) {
        results.push({
          status: 'error',
          reportPath,
          error: error.message || '报告处理失败'
        });
      }
    }

    return {
      rulePath: resolvedRulePath,
      results
    };
  }

  return {
    processReports
  };
}

module.exports = {
  createReportRunner
};
