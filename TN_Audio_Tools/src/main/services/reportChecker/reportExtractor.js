function createReportExtractor({
  parseReport,
  applyResultsToChecklist,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
}) {
  async function processSingleReport({ reportPath, checklistPath, rules }) {
    const reportData = await parseReport(reportPath);
    const textNormalizeConfig = rules.globalMatchConfig?.textNormalize || {};
    const globalMatchConfig = rules.globalMatchConfig || {};

    const extractedResultsByItemId = new Map();
    const extractedItems = rules.extractItemList.map((item) => {
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
    const unmatchedItems = extractedItems.filter((item) => !item.matched).map((item) => ({
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
