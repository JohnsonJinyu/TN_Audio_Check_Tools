const fs = require('fs/promises');
const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { styleChecklistWithCom } = require('./excelComStyler');
const { resolveOutputCell, usesPercentNumberFormat } = require('./checklistLayout');
const { resolveChecklistStyleProfile } = require('./checklistStyleProfiles');

async function applyResultsToChecklist(checklistPath, reportPath, extractedItems, reportContext = {}) {
  const writePlan = buildChecklistWritePlan(
    checklistPath,
    reportPath,
    extractedItems,
    reportContext
  );

  try {
    await copyChecklistTemplate(checklistPath, writePlan.outputPath);

    const styleEngine = await styleChecklistWithCom({
      outputPath: writePlan.outputPath,
      sheetName: writePlan.sheetName,
      decimalCells: writePlan.decimalCells,
      percentCells: writePlan.percentCells,
      skippedCells: writePlan.skippedCells,
      valueUpdates: writePlan.cellUpdates,
      reportUpdates: writePlan.reportUpdates
    });

    console.log(`[reportChecker] checklist values and styles applied via ${styleEngine}`);
    return writePlan.outputPath;
  } catch (error) {
    console.warn(`[reportChecker] COM checklist update unavailable, falling back to xlsx rewrite: ${error.message}`);
  }

  const { outputPath, sheetName, decimalCells, percentCells, skippedCells, styleProfile } = writeChecklistDataToOutput(
    checklistPath,
    reportPath,
    extractedItems,
    reportContext,
    writePlan.outputPath
  );

  await applyChecklistStylesToOutput(outputPath, sheetName, decimalCells, percentCells, skippedCells, styleProfile);
  return outputPath;
}

function buildChecklistWritePlan(checklistPath, reportPath, extractedItems, reportContext = {}) {
  const workbook = XLSX.readFile(checklistPath, { cellStyles: true, cellDates: true });
  const sheetName = resolveChecklistSheetName(getWorkbookSheetNames(workbook), reportPath, reportContext);
  const worksheet = workbook.Sheets[sheetName];
  const styleProfile = resolveChecklistStyleProfile({ reportContext, sheetName });
  const decimalCells = [];
  const percentCells = [];
  const skippedCells = [];
  const cellUpdates = [];

  for (const itemResult of extractedItems) {
    const resolvedOutputCell = resolveOutputCell(worksheet, itemResult);
    if (itemResult?.skipped && resolvedOutputCell) {
      skippedCells.push(resolvedOutputCell);
    }

    if (!shouldWriteItem(itemResult)) {
      continue;
    }

    if (!resolvedOutputCell) {
      continue;
    }

    const normalizedValue = normalizeChecklistValue(worksheet, resolvedOutputCell, itemResult.value);
    cellUpdates.push({
      sheetName,
      cellAddress: resolvedOutputCell,
      value: normalizedValue
    });

    if (typeof normalizedValue === 'number' && Number.isFinite(normalizedValue) && /^I\d+$/i.test(resolvedOutputCell)) {
      if (usesPercentNumberFormat(worksheet, resolvedOutputCell)) {
        percentCells.push(resolvedOutputCell);
      } else {
        decimalCells.push(resolvedOutputCell);
      }
    }
  }

  const outputPath = path.join(
    path.dirname(reportPath),
    buildOutputFileName(path.parse(reportPath).name)
  );

  return {
    outputPath,
    sheetName,
    decimalCells: [...new Set(decimalCells)],
    percentCells: [...new Set(percentCells)],
    skippedCells: [...new Set(skippedCells)],
    styleProfile,
    cellUpdates,
    reportUpdates: resolveReportSheetUpdates(workbook, reportContext)
  };
}

async function copyChecklistTemplate(checklistPath, outputPath) {
  await fs.copyFile(checklistPath, outputPath);
}

// 兜底链路仍保留 xlsx 重写，避免 COM 不可用时整个流程不可用。
function writeChecklistDataToOutput(checklistPath, reportPath, extractedItems, reportContext = {}, outputPathOverride = '') {
  const workbook = XLSX.readFile(checklistPath, { cellStyles: true, cellDates: true });
  const sheetName = resolveChecklistSheetName(getWorkbookSheetNames(workbook), reportPath, reportContext);
  const worksheet = workbook.Sheets[sheetName];
  const styleProfile = resolveChecklistStyleProfile({ reportContext, sheetName });
  const decimalCells = [];
  const percentCells = [];
  const skippedCells = [];

  for (const itemResult of extractedItems) {
    const resolvedOutputCell = resolveOutputCell(worksheet, itemResult);
    if (itemResult?.skipped && resolvedOutputCell) {
      skippedCells.push(resolvedOutputCell);
    }

    if (!shouldWriteItem(itemResult)) {
      continue;
    }

    if (!resolvedOutputCell) {
      continue;
    }

    const normalizedValue = normalizeChecklistValue(worksheet, resolvedOutputCell, itemResult.value);

    writeLegacyChecklistCell(worksheet, resolvedOutputCell, normalizedValue);

    if (typeof normalizedValue === 'number' && Number.isFinite(normalizedValue) && /^I\d+$/i.test(resolvedOutputCell)) {
      if (usesPercentNumberFormat(worksheet, resolvedOutputCell)) {
        percentCells.push(resolvedOutputCell);
      } else {
        decimalCells.push(resolvedOutputCell);
      }
    }
  }

  applyReportSheetUpdates(workbook, reportContext);

  const outputPath = outputPathOverride || path.join(
    path.dirname(reportPath),
    buildOutputFileName(path.parse(reportPath).name)
  );

  stripWorksheetFormulas(workbook);
  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx' });

  return {
    outputPath,
    sheetName,
    decimalCells,
    percentCells,
    skippedCells: [...new Set(skippedCells)],
    styleProfile
  };
}

function resolveReportSheetUpdates(workbook, reportContext = {}) {
  const reportSheet = workbook?.Sheets?.Report;
  if (!reportSheet) {
    return [];
  }

  const updates = [];
  const selectedPanelValues = reportContext?.reportPanelSelections || {};

  const headsetInterfaceValue = normalizeReportFieldValue(selectedPanelValues.B13 || reportContext.headsetInterface);
  if (headsetInterfaceValue && reportSheet.B13) {
    updates.push({ sheetName: 'Report', cellAddress: 'B13', value: headsetInterfaceValue });
  }

  const networkValue = normalizeReportNetworkValue(selectedPanelValues.B15 || reportContext.network);
  if (networkValue && reportSheet.B15) {
    updates.push({ sheetName: 'Report', cellAddress: 'B15', value: networkValue });
  }

  const vocoderValue = normalizeReportFieldValue(selectedPanelValues.C15) || buildReportBandwidthValue(reportContext);
  if (vocoderValue && reportSheet.C15) {
    updates.push({ sheetName: 'Report', cellAddress: 'C15', value: vocoderValue });
  }

  const bitrateValue = normalizeReportFieldValue(selectedPanelValues.D15 || reportContext.bitrate);
  if (bitrateValue && reportSheet.D15) {
    updates.push({ sheetName: 'Report', cellAddress: 'D15', value: bitrateValue });
  }

  return updates;
}

function applyReportSheetUpdates(workbook, reportContext = {}) {
  const reportSheet = workbook?.Sheets?.Report;
  if (!reportSheet) {
    return;
  }

  for (const update of resolveReportSheetUpdates(workbook, reportContext)) {
    writeLegacyChecklistCell(reportSheet, update.cellAddress, update.value);
  }
}

function normalizeReportNetworkValue(value) {
  const normalized = String(value || '').trim().toUpperCase();
  const networkMap = {
    VOLTE: 'VoLTE',
    VOWIFI: 'VoWiFi',
    VONR: 'VoNR',
    VOIP: 'VoIP',
    WCDMA: 'WCDMA',
    GSM: 'GSM'
  };

  return networkMap[normalized] || '';
}

function normalizeReportFieldValue(value) {
  const normalized = String(value || '').trim();
  return normalized || '';
}

function buildReportBandwidthValue(reportContext = {}) {
  const codec = String(reportContext.codec || '').trim().toUpperCase();
  const bandwidth = normalizeBandwidthValue(reportContext.bandwidth);
  if (!codec || !bandwidth) {
    return '';
  }

  return `${codec}_${bandwidth}`;
}

function normalizeBandwidthValue(value) {
  const normalized = String(value || '').trim().toUpperCase();
  if (!normalized) {
    return '';
  }

  if (normalized === 'SB') {
    return 'SWB';
  }

  return normalized;
}

function stripWorksheetFormulas(workbook) {
  for (const sheetName of getWorkbookSheetNames(workbook)) {
    const worksheet = workbook.Sheets?.[sheetName];
    if (!worksheet) {
      continue;
    }

    for (const [cellAddress, cell] of Object.entries(worksheet)) {
      if (cellAddress.startsWith('!') || !cell || typeof cell !== 'object') {
        continue;
      }

      if (Object.prototype.hasOwnProperty.call(cell, 'f')) {
        delete cell.f;
      }

      if (Object.prototype.hasOwnProperty.call(cell, 'F')) {
        delete cell.F;
      }

      if (Object.prototype.hasOwnProperty.call(cell, 'D')) {
        delete cell.D;
      }
    }
  }
}

// 样式统一在第二阶段完成，边框、对齐和数字格式都由 ExcelJS 在内存里处理。
async function applyChecklistStylesToOutput(outputPath, sheetName, decimalCells, percentCells, skippedCells, styleProfile) {
  try {
    const styleEngine = await styleChecklistWithCom({
      outputPath,
      sheetName,
      decimalCells,
      percentCells,
      skippedCells
    });
    console.log(`[reportChecker] checklist styles applied via ${styleEngine}`);
    return;
  } catch (error) {
    // COM/WPS 不可用时再回退到纯库方案，确保功能至少可用。
    console.warn(`[reportChecker] COM checklist styling unavailable, falling back to ExcelJS: ${error.message}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputPath);

  const worksheet = workbook.getWorksheet(sheetName);

  applySkippedCellStyles(worksheet, skippedCells, styleProfile);
  applyNumericFormats(worksheet, decimalCells, '0.00');
  applyNumericFormats(worksheet, percentCells, '0.00%');

  await workbook.xlsx.writeFile(outputPath);
}

function shouldWriteItem(itemResult) {
  return Boolean(itemResult?.matched) && itemResult.value !== undefined && itemResult.value !== null && itemResult.value !== '';
}

function applySkippedCellStyles(worksheet, cellAddresses, styleProfile) {
  if (!styleProfile?.skippedFill) {
    return;
  }

  for (const cellAddress of new Set(cellAddresses || [])) {
    if (!/^I\d+$/i.test(cellAddress)) {
      continue;
    }

    const cell = worksheet.getCell(cellAddress);
    cell.fill = styleProfile.skippedFill;
  }
}

function applyNumericFormats(worksheet, cellAddresses, numFmt) {
  for (const cellAddress of new Set(cellAddresses)) {
    const cell = worksheet.getCell(cellAddress);
    cell.numFmt = numFmt;
  }
}

function resolveChecklistSheetName(sheetNames, reportPath, reportContext = {}) {
  const preferredModeFromContext = resolveModeSheetByTerminalMode(sheetNames, reportContext?.terminalMode);
  if (preferredModeFromContext) {
    return preferredModeFromContext;
  }

  const reportName = path.parse(reportPath).name;
  const reportTokens = reportName.split(/[\s_-]+/).map(normalizeSheetToken).filter(Boolean);

  const preferredModeSheet = [
    {
      sheetName: 'Handset',
      reportTokens: new Set(['ha', 'handset'])
    },
    {
      sheetName: 'Handsfree',
      reportTokens: new Set(['hf', 'hh', 'handsfree'])
    },
    {
      sheetName: 'Headset',
      reportTokens: new Set(['hs', 'he', 'headset'])
    }
  ].find((candidate) => {
    const hasSheet = sheetNames.some((sheetName) => normalizeSheetToken(sheetName) === normalizeSheetToken(candidate.sheetName));
    if (!hasSheet) {
      return false;
    }

    return reportTokens.some((token) => candidate.reportTokens.has(token));
  });

  if (preferredModeSheet) {
    return preferredModeSheet.sheetName;
  }

  const preferredSheet = sheetNames.find((candidate) => {
    const normalizedCandidate = normalizeSheetToken(candidate);
    if (!normalizedCandidate || ['report', 'help', 'changelog'].includes(normalizedCandidate)) {
      return false;
    }

    return reportTokens.some((token) => token && (token.includes(normalizedCandidate) || normalizedCandidate.includes(token)));
  });

  if (preferredSheet) {
    return preferredSheet;
  }

  const handsetFallbackSheet = sheetNames.find((candidate) => normalizeSheetToken(candidate) === 'handset');
  if (handsetFallbackSheet) {
    return handsetFallbackSheet;
  }

  return sheetNames[0];
}

function resolveModeSheetByTerminalMode(sheetNames, terminalMode) {
  const normalizedMode = String(terminalMode || '').trim().toUpperCase();
  if (!normalizedMode) {
    return '';
  }

  const modeToSheet = {
    HA: 'Handset',
    HE: 'Headset',
    HS: 'Headset',
    HH: 'Handsfree',
    HF: 'Handsfree'
  };

  const targetSheet = modeToSheet[normalizedMode];
  if (!targetSheet) {
    return '';
  }

  return sheetNames.find((sheetName) => normalizeSheetToken(sheetName) === normalizeSheetToken(targetSheet)) || '';
}

function getWorkbookSheetNames(workbook) {
  if (Array.isArray(workbook?.SheetNames)) {
    return workbook.SheetNames;
  }

  if (Array.isArray(workbook?.worksheets)) {
    return workbook.worksheets.map((worksheet) => worksheet.name);
  }

  return [];
}

function normalizeSheetToken(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function writeLegacyChecklistCell(worksheet, cellAddress, rawValue) {
  const worksheetValue = toWorksheetValue(rawValue);
  const templateCell = worksheet[cellAddress] || {};

  if (typeof worksheetValue === 'number' && Number.isFinite(worksheetValue)) {
    worksheet[cellAddress] = {
      ...(templateCell || {}),
      t: 'n',
      v: worksheetValue
    };
    delete worksheet[cellAddress].w;
    return;
  }

  worksheet[cellAddress] = {
    ...(templateCell || {}),
    t: 's',
    v: String(worksheetValue ?? '')
  };
  delete worksheet[cellAddress].w;
}

function toWorksheetValue(value) {
  if (typeof value !== 'string') {
    return value;
  }

  const trimmedValue = value.trim();
  if (/^-?\d+(?:\.\d+)?$/.test(trimmedValue)) {
    const numericValue = Number.parseFloat(trimmedValue);
    if (!Number.isNaN(numericValue)) {
      return numericValue;
    }
  }

  return value;
}

function normalizeChecklistValue(worksheet, cellAddress, value) {
  const worksheetValue = toWorksheetValue(value);

  if (!usesPercentNumberFormat(worksheet, cellAddress)) {
    return worksheetValue;
  }

  if (typeof worksheetValue === 'number' && Number.isFinite(worksheetValue) && Math.abs(worksheetValue) > 1) {
    return worksheetValue / 100;
  }

  return worksheetValue;
}

function buildOutputFileName(reportName) {
  const normalizedReportName = String(reportName || '').trim();
  const timestamp = buildTimestamp();

  if (!normalizedReportName) {
    return `checklist_${timestamp}.xlsx`;
  }

  return `${normalizedReportName}_checklist_${timestamp}.xlsx`;
}

function buildTimestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');
  return [date.getFullYear(), pad(date.getMonth() + 1), pad(date.getDate())].join('') +
    '_' +
    [pad(date.getHours()), pad(date.getMinutes())].join('');
}

module.exports = {
  applyResultsToChecklist
};
