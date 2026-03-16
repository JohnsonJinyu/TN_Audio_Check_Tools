function createReportExtractor({
  parseReport,
  applyResultsToChecklist,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
}) {
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
    const textNormalizeConfig = rules.globalMatchConfig?.textNormalize || {};
    const globalMatchConfig = rules.globalMatchConfig || {};

    const extractedResultsByItemId = new Map();
    const extractedItems = rules.extractItemList.map((item) => {
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

    const outputPath = await applyResultsToChecklist(checklistPath, reportPath, extractedItems);
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
      outputPath,
      totalItems: extractedItems.length,
      matchedItems,
      skippedItems,
      unmatchedItems,
      extractedItems
    };
  }

  return {
    processSingleReport
  };
}

module.exports = {
  createReportExtractor
};
