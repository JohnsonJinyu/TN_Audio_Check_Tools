const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { styleChecklistWithCom } = require('./excelComStyler');
const { resolveOutputCell, usesPercentNumberFormat } = require('./checklistLayout');

const CHECKLIST_BORDER_RANGE = {
  startCol: 'A',
  endCol: 'K',
  startRow: 3,
  endRow: 75
};
const CHECKLIST_LEFT_ALIGN_ROWS = new Set([5, 66, 71]);
const CHECKLIST_DEFAULT_ALIGNMENT = {
  horizontal: 'center',
  vertical: 'middle'
};
const CHECKLIST_LEFT_ROW_ALIGNMENT = {
  horizontal: 'left',
  vertical: 'middle'
};

async function applyResultsToChecklist(checklistPath, reportPath, extractedItems) {
  const { outputPath, sheetName, decimalCells, percentCells } = writeChecklistDataToOutput(
    checklistPath,
    reportPath,
    extractedItems
  );

  await applyChecklistStylesToOutput(outputPath, sheetName, decimalCells, percentCells);
  return outputPath;
}

// 先用 xlsx 写出一份标准化的新工作簿，避免直接读取模板时被 shared formula 卡住。
function writeChecklistDataToOutput(checklistPath, reportPath, extractedItems) {
  const workbook = XLSX.readFile(checklistPath, { cellStyles: true, cellDates: true });
  const sheetName = resolveChecklistSheetName(getWorkbookSheetNames(workbook), reportPath);
  const worksheet = workbook.Sheets[sheetName];
  const decimalCells = [];
  const percentCells = [];

  for (const itemResult of extractedItems) {
    if (!shouldWriteItem(itemResult)) {
      continue;
    }

    const resolvedOutputCell = resolveOutputCell(worksheet, itemResult);
    if (!resolvedOutputCell) {
      continue;
    }

    const normalizedValue = toWorksheetValue(itemResult.value);

    writeLegacyChecklistCell(worksheet, resolvedOutputCell, normalizedValue);

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

  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx' });

  return {
    outputPath,
    sheetName,
    decimalCells,
    percentCells
  };
}

// 样式统一在第二阶段完成，边框、对齐和数字格式都由 ExcelJS 在内存里处理。
async function applyChecklistStylesToOutput(outputPath, sheetName, decimalCells, percentCells) {
  try {
    const styleEngine = await styleChecklistWithCom({
      outputPath,
      sheetName,
      decimalCells,
      percentCells
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

  // 如果目标是整个 A3:K75 都形成完整网格，就不能把区域内 merge 恢复回来。
  unmergeCellsInRange(worksheet, CHECKLIST_BORDER_RANGE);
  applyChecklistRegionStyles(worksheet, CHECKLIST_BORDER_RANGE);
  applyNumericFormats(worksheet, decimalCells, '0.00');
  applyNumericFormats(worksheet, percentCells, '0.00%');

  await workbook.xlsx.writeFile(outputPath);
}

function shouldWriteItem(itemResult) {
  return Boolean(itemResult?.matched) && itemResult.value !== undefined && itemResult.value !== null && itemResult.value !== '';
}

function unmergeCellsInRange(worksheet, rangeConfig) {
  const mergeRefs = [...(worksheet.model?.merges || [])];

  for (const mergeRef of mergeRefs) {
    const mergeBounds = parseRangeRef(mergeRef);
    if (!mergeBounds || !doesRangeIntersect(mergeBounds, rangeConfig)) {
      continue;
    }

    worksheet.unMergeCells(mergeRef);
  }
}

function applyChecklistRegionStyles(worksheet, rangeConfig) {
  const startColumnNumber = columnLabelToNumber(rangeConfig.startCol);
  const endColumnNumber = columnLabelToNumber(rangeConfig.endCol);

  for (let rowNumber = rangeConfig.startRow; rowNumber <= rangeConfig.endRow; rowNumber += 1) {
    for (let columnNumber = startColumnNumber; columnNumber <= endColumnNumber; columnNumber += 1) {
      const cellAddress = `${columnNumberToLabel(columnNumber)}${rowNumber}`;
      const cell = worksheet.getCell(cellAddress);

      // 这里按物理单元格逐格设置边框，才能保证整个 A3:K75 都满足“外围粗、内部细”。
      cell.alignment = {
        ...(cell.alignment || {}),
        ...resolveChecklistAlignment(rowNumber)
      };
      cell.border = buildCellBorderSpec(rowNumber, columnNumber, rangeConfig, startColumnNumber, endColumnNumber);
    }
  }
}

function applyNumericFormats(worksheet, cellAddresses, numFmt) {
  for (const cellAddress of new Set(cellAddresses)) {
    const cell = worksheet.getCell(cellAddress);
    cell.numFmt = numFmt;
  }
}

function buildCellBorderSpec(rowNumber, columnNumber, rangeConfig, startColumnNumber, endColumnNumber) {
  return {
    left: buildBorderEdge(columnNumber === startColumnNumber ? 'thick' : 'thin'),
    right: buildBorderEdge(columnNumber === endColumnNumber ? 'thick' : 'thin'),
    top: buildBorderEdge(rowNumber === rangeConfig.startRow ? 'thick' : 'thin'),
    bottom: buildBorderEdge(rowNumber === rangeConfig.endRow ? 'thick' : 'thin')
  };
}

function buildBorderEdge(style) {
  return {
    style,
    color: { argb: 'FF000000' }
  };
}

function resolveChecklistAlignment(rowNumber) {
  if (CHECKLIST_LEFT_ALIGN_ROWS.has(rowNumber)) {
    return CHECKLIST_LEFT_ROW_ALIGNMENT;
  }

  return CHECKLIST_DEFAULT_ALIGNMENT;
}

function resolveChecklistSheetName(sheetNames, reportPath) {
  const reportName = path.parse(reportPath).name;
  const reportTokens = reportName.split(/[\s_-]+/).map(normalizeSheetToken).filter(Boolean);

  const preferredModeSheet = [
    {
      sheetName: 'Handset',
      reportTokens: new Set(['ha', 'handset'])
    },
    {
      sheetName: 'Handsfree',
      reportTokens: new Set(['hf', 'handsfree'])
    },
    {
      sheetName: 'Headset',
      reportTokens: new Set(['hs', 'headset'])
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

function parseRangeRef(rangeRef) {
  const [startRef, endRef] = String(rangeRef || '').split(':');
  if (!startRef || !endRef) {
    return null;
  }

  const startCell = parseCellRef(startRef);
  const endCell = parseCellRef(endRef);
  if (!startCell || !endCell) {
    return null;
  }

  return {
    startCol: startCell.col,
    endCol: endCell.col,
    startRow: startCell.row,
    endRow: endCell.row
  };
}

function parseCellRef(cellRef) {
  const match = /^([A-Z]+)(\d+)$/i.exec(String(cellRef || '').trim());
  if (!match) {
    return null;
  }

  return {
    col: columnLabelToNumber(match[1]),
    row: Number.parseInt(match[2], 10)
  };
}

function doesRangeIntersect(mergeBounds, rangeConfig) {
  const rangeStartCol = columnLabelToNumber(rangeConfig.startCol);
  const rangeEndCol = columnLabelToNumber(rangeConfig.endCol);

  return !(
    mergeBounds.endCol < rangeStartCol ||
    mergeBounds.startCol > rangeEndCol ||
    mergeBounds.endRow < rangeConfig.startRow ||
    mergeBounds.startRow > rangeConfig.endRow
  );
}

function columnLabelToNumber(columnLabel) {
  let result = 0;
  const normalized = String(columnLabel || '').toUpperCase();

  for (const char of normalized) {
    result = result * 26 + (char.charCodeAt(0) - 64);
  }

  return result;
}

function columnNumberToLabel(columnNumber) {
  let current = columnNumber;
  let label = '';

  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }

  return label;
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
    [pad(date.getHours()), pad(date.getMinutes()), pad(date.getSeconds())].join('');
}

module.exports = {
  applyResultsToChecklist
};
