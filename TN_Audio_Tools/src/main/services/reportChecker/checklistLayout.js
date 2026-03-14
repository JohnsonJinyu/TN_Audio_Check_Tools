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

function resolveOutputCell(worksheet, itemResult) {
  const configuredCell = String(itemResult?.outputCell || '').trim();
  if (!configuredCell || !worksheet) {
    return configuredCell || null;
  }

  if (!isCompact3gppHandsetLayout(worksheet)) {
    return configuredCell;
  }

  const itemId = Number(itemResult?.itemId);
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
  isCompact3gppHandsetLayout
};