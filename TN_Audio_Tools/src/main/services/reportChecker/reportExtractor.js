const path = require('path');
const fs = require('fs/promises');

function createReportExtractor({
  parseReport,
  applyResultsToChecklist,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue,
  analyzeExcelReport,
  analyzeWordReport
}) {
  function resolveTerminalModeKey(value) {
    const terminalMode = String(value || '').trim().toUpperCase();
    if (terminalMode === 'HA') {
      return 'handset';
    }

    if (terminalMode === 'HH' || terminalMode === 'HF') {
      return 'handsfree';
    }

    if (terminalMode === 'HE' || terminalMode === 'HS') {
      return 'headset';
    }

    return '';
  }

  function resolveChecklistModeKeyFromName(checklistPath) {
    const checklistName = path.basename(checklistPath || '').toLowerCase();
    if (checklistName.includes('handsfree')) {
      return 'handsfree';
    }

    if (checklistName.includes('headset')) {
      return 'headset';
    }

    if (checklistName.includes('handset')) {
      return 'handset';
    }

    return '';
  }

  async function resolveChecklistPathForReport(checklistPath, reportContext = {}) {
    if (!checklistPath) {
      return checklistPath;
    }

    const targetModeKey = resolveTerminalModeKey(reportContext?.terminalMode);
    if (!targetModeKey) {
      return checklistPath;
    }

    const currentModeKey = resolveChecklistModeKeyFromName(checklistPath);
    if (!currentModeKey || currentModeKey === targetModeKey) {
      return checklistPath;
    }

    const modeNameMap = {
      handset: 'Handset',
      handsfree: 'Handsfree',
      headset: 'Headset'
    };
    const nextModeName = modeNameMap[targetModeKey];
    const checklistDir = path.dirname(checklistPath);
    const checklistName = path.basename(checklistPath);
    const candidateName = checklistName.replace(/handset|handsfree|headset/i, nextModeName);
    const candidatePath = path.join(checklistDir, candidateName);

    try {
      await fs.access(candidatePath);
      return candidatePath;
    } catch {
      return checklistPath;
    }
  }

  function resolveRuleProfileKey(reportData, checklistPath) {
    const modeKey = resolveTerminalModeKey(reportData?.reportContext?.terminalMode);
    if (modeKey === 'handset') {
      return 'handset';
    }

    if (modeKey === 'handsfree') {
      return 'handsfree';
    }

    if (modeKey === 'headset') {
      return 'headset';
    }

    const checklistModeKey = resolveChecklistModeKeyFromName(checklistPath);
    if (checklistModeKey) {
      return checklistModeKey;
    }

    return 'handset';
  }

  function resolveRulesForReport(rules, reportData, checklistPath) {
    if (!rules?.ruleProfiles) {
      return {
        activeRules: rules,
        profileKey: 'default'
      };
    }

    const profileKey = resolveRuleProfileKey(reportData, checklistPath);
    const activeRules = rules.ruleProfiles[profileKey]
      || rules.ruleProfiles[rules.defaultProfileKey]
      || Object.values(rules.ruleProfiles)[0];

    if (!activeRules?.extractItemList) {
      throw new Error(`未找到可用的规则 profile: ${profileKey}`);
    }

    return {
      activeRules,
      profileKey
    };
  }

  function normalizeBandwidth(value) {
    const normalized = String(value || '').trim().toUpperCase();
    if (!normalized) {
      return '';
    }

    if (['SB', 'SWB'].includes(normalized)) {
      return 'SWB';
    }

    if (normalized.endsWith('SWB')) {
      return 'SWB';
    }

    if (normalized.endsWith('WB')) {
      return 'WB';
    }

    if (normalized.endsWith('NB')) {
      return 'NB';
    }

    return normalized;
  }

  function isItemApplicable(reportData, item) {
    const reportTerminalMode = String(reportData?.reportContext?.terminalMode || '').trim().toUpperCase();
    const applicableTerminalModes = Array.isArray(item?.applicableTerminalModes)
      ? item.applicableTerminalModes.map((mode) => String(mode || '').trim().toUpperCase()).filter(Boolean)
      : [];
    if (applicableTerminalModes.length > 0 && !applicableTerminalModes.includes(reportTerminalMode)) {
      return {
        applicable: false,
        reason: `当前模式 ${reportTerminalMode || 'unknown'} 不要求测试该项`
      };
    }

    const excludedTerminalModes = Array.isArray(item?.excludedTerminalModes)
      ? item.excludedTerminalModes.map((mode) => String(mode || '').trim().toUpperCase()).filter(Boolean)
      : [];
    if (excludedTerminalModes.includes(reportTerminalMode)) {
      return {
        applicable: false,
        reason: `当前模式 ${reportTerminalMode} 不要求测试该项`
      };
    }

    const reportBandwidth = normalizeBandwidth(reportData?.reportContext?.bandwidth);
    if (!reportBandwidth) {
      return { applicable: true };
    }

    const applicableBandwidths = Array.isArray(item?.applicableBandwidths)
      ? item.applicableBandwidths.map(normalizeBandwidth).filter(Boolean)
      : [];
    if (applicableBandwidths.length > 0 && !applicableBandwidths.includes(reportBandwidth)) {
      return {
        applicable: false,
        reason: `当前带宽 ${reportBandwidth} 不要求测试该项`
      };
    }

    const excludedBandwidths = Array.isArray(item?.excludedBandwidths)
      ? item.excludedBandwidths.map(normalizeBandwidth).filter(Boolean)
      : [];
    if (excludedBandwidths.includes(reportBandwidth)) {
      return {
        applicable: false,
        reason: `当前带宽 ${reportBandwidth} 不要求测试该项`
      };
    }

    return { applicable: true };
  }

  async function processSingleReport({ reportPath, checklistPath, rules }) {
    const reportData = await parseReport(reportPath);
    const reportKind = reportData?.reportFormat === 'xlsx' ? 'excel' : 'word';
    const checklistPathForReport = await resolveChecklistPathForReport(checklistPath, reportData?.reportContext || {});
    const { activeRules, profileKey } = resolveRulesForReport(rules, reportData, checklistPathForReport);

    if (reportKind === 'word') {
      return {
        reportPath,
        reportKind,
        reportFormat: reportData?.reportFormat || 'docx',
        bundleKey: path.parse(reportPath).name,
        reportContext: reportData?.reportContext || {},
        ruleProfileKey: profileKey,
        outputPath: '',
        totalItems: 0,
        matchedItems: 0,
        skippedItems: [],
        unmatchedItems: [],
        extractedItems: [],
        audit: analyzeWordReport({ reportPath, reportData })
      };
    }

    if (!checklistPathForReport) {
      throw new Error('Excel 报告处理需要 checklist 文件。');
    }

    const textNormalizeConfig = activeRules.globalMatchConfig?.textNormalize || {};
    const globalMatchConfig = activeRules.globalMatchConfig || {};

    const extractedResultsByItemId = new Map();
    const extractedItems = activeRules.extractItemList.map((item) => {
      const applicability = isItemApplicable(reportData, item);
      if (!applicability.applicable) {
        const skippedResult = {
          itemId: item.itemId,
          checklistDesc: item.checklistDesc,
          outputCell: item.outputCell,
          extractType: item.extractType,
          matched: false,
          skipped: true,
          reason: applicability.reason
        };

        extractedResultsByItemId.set(item.itemId, skippedResult);
        return skippedResult;
      }

      let extractionResult;

      if (['summary_table_match', 'formula_calc', 'status_judge'].includes(item.extractType)) {
        extractionResult = resolveRowBasedValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);
      } else if (item.extractType === 'detail_anchor_extract') {
        extractionResult = resolveAnchorValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);
      } else if (item.extractType === 'table_struct_extract') {
        extractionResult = resolveTableValue(reportData, item, textNormalizeConfig);
      } else if (item.extractType === 'regex_extract') {
        extractionResult = resolveRegexValue(reportData, item, textNormalizeConfig, globalMatchConfig);
      } else {
        extractionResult = { matched: false, reason: `暂不支持的提取类型: ${item.extractType}` };
      }

      const itemResult = {
        itemId: item.itemId,
        checklistDesc: item.checklistDesc,
        outputCell: item.outputCell,
        extractType: item.extractType,
        ...extractionResult
      };

      extractedResultsByItemId.set(item.itemId, itemResult);
      return itemResult;
    });

    const outputPath = await applyResultsToChecklist(checklistPathForReport, reportPath, extractedItems, reportData?.reportContext || {});
    const matchedItems = extractedItems.filter((item) => item.matched).length;
    const skippedItems = extractedItems.filter((item) => item.skipped).map((item) => ({
      itemId: item.itemId,
      checklistDesc: item.checklistDesc,
      outputCell: item.outputCell,
      reason: item.reason || '当前场景不要求测试'
    }));
    const unmatchedItems = extractedItems.filter((item) => !item.matched && !item.skipped).map((item) => ({
      itemId: item.itemId,
      checklistDesc: item.checklistDesc,
      outputCell: item.outputCell,
      reason: item.reason || '未命中'
    }));

    return {
      reportPath,
      reportKind,
      reportFormat: reportData?.reportFormat || 'xlsx',
      bundleKey: path.parse(reportPath).name,
      reportContext: reportData?.reportContext || {},
      checklistPathUsed: checklistPathForReport,
      ruleProfileKey: profileKey,
      outputPath,
      totalItems: extractedItems.length,
      matchedItems,
      skippedItems,
      unmatchedItems,
      extractedItems,
      audit: analyzeExcelReport({ reportPath, reportData, extractedItems })
    };
  }

  return {
    processSingleReport
  };
}

module.exports = {
  createReportExtractor
};
