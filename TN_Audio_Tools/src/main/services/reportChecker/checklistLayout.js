function normalizeCellText(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function getWorksheetCellValue(worksheet, cellAddress) {
  return worksheet?.[cellAddress]?.v;
}

function parseOutputCell(cellAddress) {
  const match = /^([A-Z]+)(\d+)$/i.exec(String(cellAddress || '').trim());
  if (!match) {
    return null;
  }

  return {
    column: match[1].toUpperCase(),
    row: Number(match[2])
  };
}

function buildCellAddress(column, row) {
  return `${column}${row}`;
}

function isCompact3gppHandsetLayout(worksheet) {
  const row41Label = normalizeCellText(getWorksheetCellValue(worksheet, 'C41'));
  const row42Label = normalizeCellText(getWorksheetCellValue(worksheet, 'C42'));
  const row45Metric = normalizeCellText(getWorksheetCellValue(worksheet, 'C45'));
  const row45Variant = normalizeCellText(getWorksheetCellValue(worksheet, 'D45'));

  return row41Label === 'short dt class'
    && row42Label === 'short st class'
    && row45Metric === 's-mos'
    && row45Variant === 'average';
}

function isHandsfreeChecklistLayout(worksheet) {
  const row28Section = normalizeCellText(getWorksheetCellValue(worksheet, 'A28'));
  const row28Metric = normalizeCellText(getWorksheetCellValue(worksheet, 'C28'));
  const row34Section = normalizeCellText(getWorksheetCellValue(worksheet, 'A34'));
  const row53Section = normalizeCellText(getWorksheetCellValue(worksheet, 'A53'));

  return row28Section === 'echo control characteristics'
    && row28Metric === 'short dt class'
    && row34Section === 'quality in the presence of ambient noise'
    && row53Section === 'polqa mos';
}

function resolveHandsfreeOutputCell(itemId, configuredCell) {
  const explicitMap = {
    26: 'I19',
    27: 'I20',
    28: 'I21',
    29: 'I22',
    30: 'I23',
    31: 'I24',
    32: 'I25',
    33: 'I26',
    34: 'I27',
    36: 'I28',
    37: 'I29',
    38: 'I30',
    39: 'I31',
    40: 'I32',
    41: 'I33',
    42: 'I34',
    43: 'I35',
    44: 'I36',
    45: 'I37',
    46: 'I38',
    47: 'I39',
    48: 'I40',
    49: 'I41',
    50: 'I42',
    51: 'I43',
    52: 'I44',
    53: 'I45',
    54: 'I46',
    55: 'I47',
    56: 'I48',
    57: 'I49',
    58: 'I50',
    59: 'I51',
    60: 'I52',
    65: 'I54',
    66: 'I55',
    67: 'I56',
    68: 'I57'
  };

  return explicitMap[itemId] || configuredCell;
}

function resolveOutputCell(worksheet, itemResult) {
  const configuredCell = String(itemResult?.outputCell || '').trim();
  if (!configuredCell || !worksheet) {
    return configuredCell || null;
  }

  const itemId = Number(itemResult?.itemId);

  if (isHandsfreeChecklistLayout(worksheet)) {
    return resolveHandsfreeOutputCell(itemId, configuredCell);
  }

  if (!isCompact3gppHandsetLayout(worksheet)) {
    return configuredCell;
  }

  if (itemId === 36 || itemId === 39) {
    return null;
  }

  if (itemId === 37) {
    return 'I41';
  }

  if (itemId === 38) {
    return 'I42';
  }

  if (itemId === 40) {
    return 'I43';
  }

  if (itemId === 41) {
    return 'I44';
  }

  if (itemId >= 42) {
    const parsed = parseOutputCell(configuredCell);
    if (parsed?.column === 'I' && parsed.row >= 47) {
      return buildCellAddress('I', parsed.row - 2);
    }
  }

  return configuredCell;
}

function usesPercentNumberFormat(worksheet, cellAddress) {
  const numberFormat = String(worksheet?.[cellAddress]?.z || '').trim();
  return numberFormat.includes('%');
}

module.exports = {
  resolveOutputCell,
  usesPercentNumberFormat,
  isCompact3gppHandsetLayout,
  isHandsfreeChecklistLayout
};