const fs = require('fs/promises');
const path = require('path');
const { clampMaxConcurrentTasks } = require('./concurrency');

function buildTimestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');
  return [date.getFullYear(), pad(date.getMonth() + 1), pad(date.getDate())].join('') +
    '_' +
    [pad(date.getHours()), pad(date.getMinutes()), pad(date.getSeconds())].join('');
}

function sanitizeFileNameSegment(value) {
  return String(value || '')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}

function buildVocoderSegment(panelSelections = {}) {
  const explicitVocoder = sanitizeFileNameSegment(panelSelections?.C15 || '');
  if (explicitVocoder) {
    return explicitVocoder;
  }

  return '';
}

function buildPlannedOutputBaseName(reportPath, projectMeta = {}, panelSelections = {}) {
  const timestamp = buildTimestamp();
  const projectName = sanitizeFileNameSegment(projectMeta?.projectName || '');
  const projectPhase = sanitizeFileNameSegment(projectMeta?.projectPhase || '');
  const network = sanitizeFileNameSegment(panelSelections?.B15 || '');
  const vocoder = buildVocoderSegment(panelSelections);

  if (projectName && projectPhase) {
    const parts = [projectName, projectPhase];
    if (network) parts.push(network);
    if (vocoder) parts.push(vocoder);
    parts.push('checklist', timestamp);
    return `${parts.join('_')}.xlsx`;
  }

  const reportName = sanitizeFileNameSegment(path.parse(reportPath || '').name);
  return `${reportName || 'checklist'}_checklist_${timestamp}.xlsx`;
}

function buildOutputPlan(reportPaths, outputDirectory, reportProjectMetaByPath = {}, reportPanelSelectionsByPath = {}) {
  const baseNameGroups = new Map();

  reportPaths.forEach((reportPath) => {
    const baseName = buildPlannedOutputBaseName(
      reportPath,
      reportProjectMetaByPath?.[reportPath] || {},
      reportPanelSelectionsByPath?.[reportPath] || {}
    );

    const group = baseNameGroups.get(baseName) || [];
    group.push(reportPath);
    baseNameGroups.set(baseName, group);
  });

  const outputPathByReport = {};
  baseNameGroups.forEach((group, baseName) => {
    if (group.length === 1) {
      outputPathByReport[group[0]] = path.join(outputDirectory, baseName);
      return;
    }

    const usedNames = new Set();
    group.forEach((reportPath, index) => {
      const parsedBaseName = path.parse(baseName);
      const reportSuffix = sanitizeFileNameSegment(path.parse(reportPath).name) || `report_${index + 1}`;
      let candidateName = `${parsedBaseName.name}_${reportSuffix}${parsedBaseName.ext || '.xlsx'}`;
      let dedupeIndex = 2;
      while (usedNames.has(candidateName)) {
        candidateName = `${parsedBaseName.name}_${reportSuffix}_${dedupeIndex}${parsedBaseName.ext || '.xlsx'}`;
        dedupeIndex += 1;
      }
      usedNames.add(candidateName);
      outputPathByReport[reportPath] = path.join(outputDirectory, candidateName);
    });
  });

  return outputPathByReport;
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

function buildPerReportOutputPath(checklistPath, outputDirectory, reportPath, reportContext = {}) {
  const reportName = path.parse(reportPath || '').name;
  const timestamp = buildTimestamp();

  // 尝试使用项目名_阶段_vocoder_网络_checklist_日期格式
  const projectName = String(reportContext?.projectName || '').trim();
  const projectPhase = String(reportContext?.projectPhase || '').trim();
  const vocoder = String(reportContext?.codec || '').trim();
  const network = String(reportContext?.network || '').trim();
  const bandwidth = String(reportContext?.bandwidth || '').trim();

  if (projectName && projectPhase) {
    const parts = [projectName, projectPhase];
    if (vocoder) parts.push(vocoder);
    if (network) parts.push(network);
    if (bandwidth) parts.push(bandwidth);
    parts.push('checklist', timestamp);
    return path.join(outputDirectory, `${parts.join('_')}.xlsx`);
  }

  // 回退: 用报告名 + checklist 名
  const checklistName = path.parse(checklistPath || '').name;
  return path.join(outputDirectory, `${reportName}_${checklistName}_${timestamp}.xlsx`);
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
      // checklistPath 为空时，后续会尝试从规则配置中获取内置模板，这里不抛出错误
      console.log('[reportRunner] 未提供 checklist 文件，将尝试使用内置模板');
    }

    if (checklistPath) {
      const resolvedChecklistPath = String(checklistPath || '').trim();
      if (!resolvedChecklistPath) {
        return;
      }
      const checklistExtension = path.extname(resolvedChecklistPath).toLowerCase();
      if (!supportedChecklistExtensions.has(checklistExtension)) {
        throw new Error('checklist 仅支持 .xlsx 或 .xls 文件');
      }

      await fs.access(resolvedChecklistPath);
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
    ruleProfileOverridesByPath,
    reportProjectMetaByPath,
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
    const outputPlanByReport = buildOutputPlan(
      reportPaths,
      resolvedOutputDirectory,
      reportProjectMetaByPath,
      reportPanelSelectionsByPath
    );
    const resolvedMaxConcurrentTasks = Math.min(total, clampMaxConcurrentTasks(maxConcurrentTasks, 1));
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
        const projectMeta = reportProjectMetaByPath?.[reportPath] || {};
        const result = await processSingleReport({
          reportPath,
          checklistPath,
          rules,
          customer,
          reportPanelSelections,
          reportPanelSelectionsOverride: reportPanelSelectionsByPath?.[reportPath] || null,
          ruleProfileOverride: ruleProfileOverridesByPath?.[reportPath] || '',
          outputDirectory: resolvedOutputDirectory,
          checklistWriteOptions: {
            outputPath: outputPlanByReport[reportPath]
          },
          projectMeta: {
            projectName: String(projectMeta.projectName || '').trim(),
            projectPhase: String(projectMeta.projectPhase || '').trim()
          }
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
