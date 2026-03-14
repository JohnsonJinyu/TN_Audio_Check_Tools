require('../src/main/services/reportChecker/runtimePolyfills');

const fs = require('fs');
const path = require('path');
const JSON5 = require('json5');
const XLSX = require('xlsx');
const WordExtractor = require('word-extractor');
const { createReportConverter } = require('../src/main/services/reportChecker/reportConverter');
const { createReportSource } = require('../src/main/services/reportChecker/reportSource');
const { resolveOutputCell } = require('../src/main/services/reportChecker/checklistLayout');
const {
  createSearchData,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
} = require('../src/main/services/reportChecker/reportAnalysis');

const CORPUS_PAIRS = [
  ['Kansas NA_DVT1_VOLTE AMR_NB_2024729.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_VOLTE_AMR_NB_v5.0.0_0729.xlsx'],
  ['Kansas NA_DVT1_VOLTE AMR_WB_2024729.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_VOLTE_AMR_WB_v5.0.0_0729.xlsx'],
  ['Kansas NA_Handset_VOIP_DVT1_20240730.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_VoIP_v5.0.0_0730.xlsx'],
  ['Kansas_DVT1_HA_VOLTE_EVS_NB_0803.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_VOLTE_EVS_NB_v5.0.0_0810.xlsx'],
  ['Kansas_DVT1_HA_VOLTE_EVS_WB_0803.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_VOLTE_EVS_WB_v5.0.0_0810.xlsx'],
  ['kansasNA_DVT1_handset_VONR_EVS_NB_0814.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_VONR_EVS_NB_v1.0_0814.xlsx'],
  ['kansasNA_DVT1_handset_VONR_EVS_WB_0814.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_VONR_EVS_WB_v1.0_0814.xlsx'],
  ['KANSAS_DVT1_HA_WCDMA_NB_0725.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_WCDMA_AMR_NB_v5.0.0_0725.xlsx'],
  ['KANSAS_DVT1_HA_WCDMA_WB_0725.doc', 'Kansas5G NA_DVT1_Source1st_3GPP_Tuning_report_for_Handset_mode_WCDMA_AMR_WB_v5.0.0_0724.xlsx']
];

function cellValue(worksheet, cellAddress) {
  return worksheet[cellAddress] ? worksheet[cellAddress].v : undefined;
}

function loadRules(appPath) {
  const rulePath = path.join(appPath, 'src', 'renderer', 'modules', 'reportChecker', 'config', 'moto_rules_for_analysis.json5');
  return JSON5.parse(fs.readFileSync(rulePath, 'utf8'));
}

function createParser() {
  const wordExtractor = new WordExtractor();
  const { convertDocToTemporaryDocx } = createReportConverter({
    childProcessTimeoutMs: 90000,
    libreOfficeCandidatePaths: [
      'C:/Program Files/LibreOffice/program/soffice.exe',
      'C:/Program Files (x86)/LibreOffice/program/soffice.exe'
    ]
  });

  return createReportSource({
    supportedReportExtensions: new Set(['.doc', '.docx']),
    convertDocToTemporaryDocx,
    wordExtractor,
    createSearchData
  }).parseReport;
}

function extractItems(reportData, rules) {
  const textNormalizeConfig = rules.globalMatchConfig?.textNormalize || {};
  const globalMatchConfig = rules.globalMatchConfig || {};
  const extractedResultsByItemId = new Map();

  return rules.extractItemList.map((item) => {
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
}

function diffItems(extractedItems, checklistPath) {
  const targetWorksheet = XLSX.readFile(checklistPath).Sheets.Handset;

  return extractedItems.filter((item) => {
    const resolvedOutputCell = resolveOutputCell(targetWorksheet, item);
    if (!resolvedOutputCell) {
      return false;
    }

    const targetValue = cellValue(targetWorksheet, resolvedOutputCell);
    const outputValue = item.matched ? item.value : undefined;

    if (targetValue === undefined && outputValue === undefined) {
      return false;
    }

    if (typeof targetValue === 'number' && outputValue !== undefined) {
      const numeric = Number(outputValue);
      if (!Number.isNaN(numeric)) {
        return Math.abs(targetValue - numeric) > 0.011;
      }
    }

    return String(targetValue ?? '') !== String(outputValue ?? '');
  }).map((item) => ({
    outputCell: resolveOutputCell(targetWorksheet, item),
    itemId: item.itemId,
    checklistDesc: item.checklistDesc,
    target: cellValue(targetWorksheet, resolveOutputCell(targetWorksheet, item)),
    output: item.matched ? item.value : undefined,
    matched: item.matched,
    sourceType: item.sourceType,
    reason: item.reason || null,
    sourcePreview: item.sourcePreview
  }));
}

async function analyzeCorpus() {
  const appPath = path.resolve(__dirname, '..');
  const corpusDir = process.argv[2]
    ? path.resolve(process.argv[2])
    : path.resolve(appPath, '..', '..', '参考文件', '3GPP Handset report');
  const parseReport = createParser();
  const rules = loadRules(appPath);
  const byItem = new Map();
  const byReport = [];

  for (const [reportName, checklistName] of CORPUS_PAIRS) {
    const reportPath = path.join(corpusDir, reportName);
    const checklistPath = path.join(corpusDir, checklistName);

    try {
      const reportData = await parseReport(reportPath);
      const extractedItems = extractItems(reportData, rules);
      const diffs = diffItems(extractedItems, checklistPath);

      byReport.push({
        reportName,
        checklistName,
        matchedItems: extractedItems.filter((item) => item.matched).length,
        totalItems: extractedItems.length,
        diffCount: diffs.length,
        diffItemIds: diffs.map((item) => item.itemId)
      });

      for (const diff of diffs) {
        if (!byItem.has(diff.itemId)) {
          byItem.set(diff.itemId, []);
        }

        byItem.get(diff.itemId).push({ reportName, ...diff });
      }
    } catch (error) {
      byReport.push({
        reportName,
        checklistName,
        error: error.message || String(error)
      });
    }
  }

  const itemSummary = Array.from(byItem.entries())
    .map(([itemId, diffs]) => ({
      itemId,
      reports: diffs.length,
      checklistDesc: diffs[0].checklistDesc,
      sample: diffs.slice(0, 5)
    }))
    .sort((left, right) => right.reports - left.reports || left.itemId - right.itemId);

  return { byReport, itemSummary };
}

module.exports = {
  analyzeCorpus
};

if (require.main === module) {
  analyzeCorpus()
    .then((summary) => {
      console.log(JSON.stringify(summary, null, 2));
    })
    .catch((error) => {
      console.error(error.stack || error);
      process.exit(1);
    });
}