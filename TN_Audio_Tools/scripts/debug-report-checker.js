const { processReports } = require('../src/main/services/reportCheckerService');
const XLSX = require('xlsx');

function cellValue(ws, cell) {
  return ws[cell] ? ws[cell].v : undefined;
}

async function main() {
  const mode = process.argv[2];
  const appPath = process.cwd();

  if (mode === 'items') {
    const base = 'c:/Users/Lenovo/Desktop/Audio_AI_TOOL/参考文件/lamu_HA';
    const result = (await processReports({
      reportPaths: [`${base}/LAMU26_DVT2_VOWIFI_AMR_NB_HA_1105.doc`],
      checklistPath: `${base}/lamu26_DVT2_1st_Voice_Tuning_Checklist_for_Handset_VOWIFI_AMR_NB_1106.xlsx`,
      appPath
    })).results[0];

    const items = result.extractedItems.filter((item) => [20, 29, 30, 31, 32, 33, 34, 38, 41].includes(item.itemId));
    console.log(JSON.stringify(items, null, 2));
    return;
  }

  if (mode === 'utah-diff') {
    const report = 'c:/Users/Lenovo/Desktop/Audio_AI_TOOL/TN_Audio/TN_Audio_Tools/docs/Utah_Handset_mode_VOLTE_EVS_SWB_0521.doc';
    const checklist = 'c:/Users/Lenovo/Desktop/Audio_AI_TOOL/TN_Audio/TN_Audio_Tools/docs/moto_checlist_空.xlsx';
    const target = 'c:/Users/Lenovo/Desktop/Audio_AI_TOOL/TN_Audio/TN_Audio_Tools/docs/Utah_Voice_Tuning_Checklist_v5.0.2_Handset_mode_VOLTE_EVS_SWB_0521_人工版本.xlsx';
    const targetWs = XLSX.readFile(target).Sheets.Handset;
    const result = (await processReports({ reportPaths: [report], checklistPath: checklist, appPath })).results[0];
    const diffs = result.extractedItems.filter((item) => {
      const targetValue = cellValue(targetWs, item.outputCell);
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
      itemId: item.itemId,
      cell: item.outputCell,
      target: cellValue(targetWs, item.outputCell),
      output: item.matched ? item.value : undefined,
      matched: item.matched,
      sourceType: item.sourceType,
      sourcePreview: item.sourcePreview,
      reason: item.reason || null
    }));

    console.log(JSON.stringify({ matchedItems: result.matchedItems, totalItems: result.totalItems, diffCount: diffs.length, diffs }, null, 2));
    return;
  }

  if (mode === 'lamu-diff') {
    const base = 'c:/Users/Lenovo/Desktop/Audio_AI_TOOL/参考文件/lamu_HA';
    const target = `${base}/lamu26_DVT2_1st_Voice_Tuning_Checklist_for_Handset_VOWIFI_AMR_NB_1106.xlsx`;
    const targetWs = XLSX.readFile(target).Sheets.Handset;
    const result = (await processReports({
      reportPaths: [`${base}/LAMU26_DVT2_VOWIFI_AMR_NB_HA_1105.doc`],
      checklistPath: target,
      appPath
    })).results[0];

    const diffs = result.extractedItems.filter((item) => {
      const targetValue = cellValue(targetWs, item.outputCell);
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
      itemId: item.itemId,
      cell: item.outputCell,
      target: cellValue(targetWs, item.outputCell),
      output: item.matched ? item.value : undefined,
      matched: item.matched,
      sourceType: item.sourceType,
      sourcePreview: item.sourcePreview,
      reason: item.reason || null
    }));

    console.log(JSON.stringify({ matchedItems: result.matchedItems, totalItems: result.totalItems, diffCount: diffs.length, diffs }, null, 2));
    return;
  }

  if (mode === 'formula-check') {
    const report = process.argv[3];
    const checklist = process.argv[4];
    const result = (await processReports({ reportPaths: [report], checklistPath: checklist, appPath })).results[0];
    const items = result.extractedItems.filter((item) => [11, 12, 14, 15].includes(item.itemId));
    console.log(JSON.stringify(items, null, 2));
  }
}

main().catch((error) => {
  console.error(error.stack || error);
  process.exit(1);
});
