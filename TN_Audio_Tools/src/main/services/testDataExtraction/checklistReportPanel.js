const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

const REPORT_PANEL_FIELDS = [
  { cell: 'B13', key: 'headsetInterface', label: 'Headset Interface' },
  { cell: 'B15', key: 'network', label: 'Network' },
  { cell: 'C15', key: 'vocoder', label: 'Vocoder' },
  { cell: 'D15', key: 'bitrate', label: 'Bitrate' }
];

async function parseChecklistReportOptions(checklistPath) {
  if (!checklistPath) {
    throw new Error('缺少 checklist 路径');
  }

  const extension = path.extname(checklistPath).toLowerCase();
  if (extension === '.xls') {
    return parseChecklistReportOptionsFromXls(checklistPath);
  }

  return parseChecklistReportOptionsFromXlsx(checklistPath);
}

async function parseChecklistReportOptionsFromXlsx(checklistPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(checklistPath);

  const reportSheet = workbook.getWorksheet('Report') || workbook.worksheets[0];
  if (!reportSheet) {
    throw new Error('未找到可读取的 checklist 工作表');
  }

  const fields = REPORT_PANEL_FIELDS.map((field) => {
    const cell = reportSheet.getCell(field.cell);
    const currentValue = normalizeOptionValue(cell.value);
    const options = extractOptionsFromValidation(reportSheet, cell.dataValidation, field.cell);
    return {
      ...field,
      currentValue,
      options,
      dataValidation: cell.dataValidation
    };
  });

  const cascadeMaps = buildCascadeMaps(reportSheet, fields);
  const fieldsWithCascade = fields.map((field) => {
    const { dataValidation: _dv, ...rest } = field;
    return cascadeMaps[field.cell] ? { ...rest, cascadeMap: cascadeMaps[field.cell] } : rest;
  });

  return {
    checklistPath,
    reportSheetName: reportSheet.name,
    fields: fieldsWithCascade
  };
}

function parseChecklistReportOptionsFromXls(checklistPath) {
  const workbook = XLSX.readFile(checklistPath, { cellStyles: true, cellDates: true });
  const sheetName = workbook.SheetNames.includes('Report') ? 'Report' : workbook.SheetNames[0];
  const reportSheet = workbook.Sheets[sheetName];
  if (!reportSheet) {
    throw new Error('未找到可读取的 checklist 工作表');
  }

  const fields = REPORT_PANEL_FIELDS.map((field) => {
    const cell = reportSheet[field.cell];
    return {
      ...field,
      currentValue: normalizeOptionValue(cell?.v),
      options: []
    };
  });

  return {
    checklistPath,
    reportSheetName: sheetName,
    fields,
    note: '.xls 模板暂不支持读取下拉候选，当前仅回传单元格现值。'
  };
}

function buildCascadeMaps(reportSheet, fields) {
  const cascadeMaps = {};

  const b15Field = fields.find((f) => f.cell === 'B15');
  const c15Field = fields.find((f) => f.cell === 'C15');
  const d15Field = fields.find((f) => f.cell === 'D15');

  if (b15Field && c15Field?.dataValidation) {
    const c15Map = {};
    const b15Options = b15Field.options || [];
    for (const b15Value of b15Options) {
      c15Map[b15Value] = extractOptionsFromValidation(
        reportSheet,
        c15Field.dataValidation,
        'C15',
        { B15: b15Value }
      );
    }
    cascadeMaps['C15'] = c15Map;
  }

  if (c15Field && d15Field?.dataValidation) {
    const allC15Values = new Set();
    if (cascadeMaps['C15']) {
      for (const opts of Object.values(cascadeMaps['C15'])) {
        for (const v of opts) allC15Values.add(v);
      }
    } else {
      for (const v of c15Field.options || []) allC15Values.add(v);
    }
    const d15Map = {};
    for (const c15Value of allC15Values) {
      d15Map[c15Value] = extractOptionsFromValidation(
        reportSheet,
        d15Field.dataValidation,
        'D15',
        { C15: c15Value }
      );
    }
    cascadeMaps['D15'] = d15Map;
  }

  return cascadeMaps;
}

function extractOptionsFromValidation(reportSheet, dataValidation, cellAddress, cellOverrides) {
  if (!dataValidation || dataValidation.type !== 'list') {
    return [];
  }

  const formula = Array.isArray(dataValidation.formulae) ? dataValidation.formulae[0] : '';
  if (!formula) {
    return [];
  }

  if (formula.startsWith('"') && formula.endsWith('"')) {
    return formula
      .slice(1, -1)
      .split(',')
      .map((item) => item.trim())
      .filter(Boolean);
  }

  const formulaText = formula.startsWith('=') ? formula.slice(1) : formula;
  return resolveValidationFormulaOptions(reportSheet, formulaText, cellAddress, new Set(), cellOverrides || {});
}

function expandRangeAddresses(startCell, endCell) {
  const start = splitCellAddress(startCell);
  const end = splitCellAddress(endCell);

  if (!start || !end) {
    return [];
  }

  const addresses = [];
  for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row += 1) {
    for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col += 1) {
      addresses.push(`${columnNumberToName(col)}${row}`);
    }
  }

  return addresses;
}

function splitCellAddress(address) {
  const match = String(address || '').toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return null;
  }

  return {
    col: columnNameToNumber(match[1]),
    row: Number.parseInt(match[2], 10)
  };
}

function resolveValidationFormulaOptions(reportSheet, formulaText, cellAddress, visited, cellOverrides) {
  const normalizedFormula = String(formulaText || '').trim();
  if (!normalizedFormula) {
    return [];
  }

  const visitKey = `${reportSheet?.name || 'sheet'}::${cellAddress || ''}::${normalizedFormula}`;
  if (visited.has(visitKey)) {
    return [];
  }
  visited.add(visitKey);

  const indirectMatch = normalizedFormula.match(/^INDIRECT\((.+)\)$/i);
  if (indirectMatch) {
    const evaluatedReference = evaluateIndirectExpression(reportSheet, indirectMatch[1], cellOverrides);
    const referencedNameOptions = resolveDefinedNameOptions(reportSheet.workbook, evaluatedReference);
    if (referencedNameOptions.length > 0) {
      return referencedNameOptions;
    }

    if (!evaluatedReference) {
      return [];
    }

    return resolveValidationFormulaOptions(reportSheet, evaluatedReference, cellAddress, visited, cellOverrides);
  }

  const namedRangeOptions = resolveDefinedNameOptions(reportSheet.workbook, normalizedFormula);
  if (namedRangeOptions.length > 0) {
    return namedRangeOptions;
  }

  return resolveSheetReferenceOptions(reportSheet.workbook, normalizedFormula, reportSheet.name);
}

function resolveDefinedNameOptions(workbook, definedName) {
  const normalizedName = String(definedName || '').trim();
  if (!normalizedName) {
    return [];
  }

  const nameModel = Array.isArray(workbook?.definedNames?.model) ? workbook.definedNames.model : [];
  const matchedName = nameModel.find((item) => String(item?.name || '').trim().toUpperCase() === normalizedName.toUpperCase());
  if (!matchedName) {
    return [];
  }

  const ranges = Array.isArray(matchedName.ranges) ? matchedName.ranges : [];
  const options = ranges.flatMap((rangeRef) => resolveSheetReferenceOptions(workbook, rangeRef, ''));
  return [...new Set(options.filter(Boolean))];
}

function resolveSheetReferenceOptions(workbook, formulaText, currentSheetName) {
  const normalizedFormula = String(formulaText || '').trim();

  const rangeMatch = normalizedFormula.match(/^(?:(?:'([^']+)')|([^!]+))!([$]?[A-Z]+[$]?\d+)(?::([$]?[A-Z]+[$]?\d+))?$/);
  if (rangeMatch) {
    const targetSheetName = (rangeMatch[1] || rangeMatch[2] || '').trim();
    const startCell = rangeMatch[3].replace(/\$/g, '');
    const endCell = (rangeMatch[4] || rangeMatch[3]).replace(/\$/g, '');
    const targetSheet = workbook.getWorksheet(targetSheetName);
    if (!targetSheet) {
      return [];
    }

    const addresses = expandRangeAddresses(startCell, endCell);
    const options = addresses
      .map((address) => normalizeOptionValue(targetSheet.getCell(address).value))
      .filter(Boolean);

    return [...new Set(options)];
  }

  const singleCellMatch = normalizedFormula.match(/^([$]?[A-Z]+[$]?\d+)$/);
  if (singleCellMatch) {
    const targetSheet = workbook.getWorksheet(currentSheetName);
    if (!targetSheet) {
      return [];
    }

    const value = normalizeOptionValue(targetSheet.getCell(singleCellMatch[1].replace(/\$/g, '')).value);
    return value ? [value] : [];
  }

  return [];
}

function evaluateIndirectExpression(reportSheet, expression, cellOverrides) {
  const parts = String(expression || '')
    .split('&')
    .map((item) => item.trim())
    .filter(Boolean);

  const resolved = parts.map((part) => resolveIndirectPart(reportSheet, part, cellOverrides));
  if (resolved.some((item) => item === null)) {
    return '';
  }

  return resolved.join('');
}

function resolveIndirectPart(reportSheet, part, cellOverrides) {
  if (!part) {
    return '';
  }

  const stringMatch = part.match(/^"(.*)"$/);
  if (stringMatch) {
    return stringMatch[1];
  }

  const cellRefMatch = part.match(/^([$]?[A-Z]+[$]?\d+)$/i);
  if (cellRefMatch) {
    const normalizedRef = cellRefMatch[1].replace(/\$/g, '').toUpperCase();
    if (cellOverrides && Object.prototype.hasOwnProperty.call(cellOverrides, normalizedRef)) {
      return String(cellOverrides[normalizedRef] || '');
    }

    const cellValue = normalizeOptionValue(reportSheet.getCell(normalizedRef).value);
    return cellValue || '';
  }

  const definedNameOptions = resolveDefinedNameOptions(reportSheet.workbook, part);
  if (definedNameOptions.length === 1) {
    return definedNameOptions[0];
  }

  return part;
}

function columnNameToNumber(columnName) {
  let columnNumber = 0;
  const normalized = String(columnName || '').toUpperCase();

  for (let index = 0; index < normalized.length; index += 1) {
    columnNumber = (columnNumber * 26) + (normalized.charCodeAt(index) - 64);
  }

  return columnNumber;
}

function columnNumberToName(columnNumber) {
  let value = Number(columnNumber);
  let result = '';

  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }

  return result;
}

function normalizeOptionValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'object') {
    if (typeof value.text === 'string') {
      return value.text.trim();
    }

    if (value.richText && Array.isArray(value.richText)) {
      return value.richText.map((item) => item?.text || '').join('').trim();
    }

    if (value.result !== undefined && value.result !== null) {
      return String(value.result).trim();
    }
  }

  return String(value).trim();
}

module.exports = {
  parseChecklistReportOptions
};
