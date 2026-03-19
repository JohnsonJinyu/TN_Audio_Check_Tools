const fs = require('fs/promises');
const path = require('path');
const XLSX = require('xlsx');
const AdmZip = require('adm-zip');
const { resolveOutputCell, usesPercentNumberFormat } = require('./checklistLayout');
const { styleChecklistWithLibraries } = require('./checklistLibraryStyler');
const { resolveChecklistStyleProfile } = require('./checklistStyleProfiles');

async function applyResultsToChecklist(checklistPath, reportPath, extractedItems, reportContext = {}) {
  const resolvedChecklistPath = path.resolve(checklistPath);
  const resolvedReportPath = path.resolve(reportPath);
  const writePlan = buildChecklistWritePlan(
    resolvedChecklistPath,
    resolvedReportPath,
    extractedItems,
    reportContext
  );

  await fs.copyFile(resolvedChecklistPath, writePlan.outputPath);
  await applyChecklistWritePlan(writePlan);
  await applyChecklistStyles(writePlan, extractedItems, reportContext);

  console.log('[reportChecker] checklist values applied via library-only template-preserving pipeline');
  return writePlan.outputPath;
}

function buildChecklistWritePlan(checklistPath, reportPath, extractedItems, reportContext = {}) {
  const workbook = XLSX.readFile(checklistPath, { cellStyles: true, cellDates: true });
  const sheetName = resolveChecklistSheetName(getWorkbookSheetNames(workbook), reportPath, reportContext);
  const worksheet = workbook.Sheets[sheetName];
  const valueUpdates = [];

  for (const itemResult of extractedItems) {
    const resolvedOutputCell = resolveOutputCell(worksheet, itemResult);
    if (!shouldWriteItem(itemResult)) {
      continue;
    }

    if (!resolvedOutputCell) {
      continue;
    }

    const normalizedValue = normalizeChecklistValue(worksheet, resolvedOutputCell, itemResult.value);

    valueUpdates.push({
      sheetName,
      cellAddress: resolvedOutputCell,
      value: normalizedValue
    });
  }

  const outputPath = path.join(
    path.dirname(reportPath),
    buildOutputFileName(path.parse(reportPath).name)
  );

  return {
    outputPath,
    templatePath: checklistPath,
    sheetName,
    valueUpdates,
    reportUpdates: resolveReportSheetUpdates(workbook, reportContext)
  };
}

async function applyChecklistWritePlan(writePlan) {
  const zip = new AdmZip(writePlan.outputPath);

  const workbookXmlPath = 'xl/workbook.xml';
  const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
  const workbookXml = readZipText(zip, workbookXmlPath);
  const workbookRelsXml = readZipText(zip, workbookRelsPath);

  const worksheetPathByName = resolveWorksheetPathByName(workbookXml, workbookRelsXml);
  const groupedUpdates = groupUpdatesBySheet([...(writePlan.valueUpdates || []), ...(writePlan.reportUpdates || [])]);

  for (const [sheetName, updates] of Object.entries(groupedUpdates)) {
    const worksheetPath = worksheetPathByName[sheetName];
    if (!worksheetPath) {
      continue;
    }

    let worksheetXml = readZipText(zip, worksheetPath);
    for (const update of updates) {
      worksheetXml = replaceCellValueInWorksheetXml(worksheetXml, update.cellAddress, update.value);
    }
    zip.updateFile(worksheetPath, Buffer.from(worksheetXml, 'utf8'));
  }

  const updatedWorkbookXml = enableWorkbookRecalculation(workbookXml);
  const finalizedWorkbookXml = ensureFirstSheetActive(updatedWorkbookXml);
  if (finalizedWorkbookXml !== workbookXml) {
    zip.updateFile(workbookXmlPath, Buffer.from(finalizedWorkbookXml, 'utf8'));
  }

  zip.writeZip(writePlan.outputPath);
}

async function applyChecklistStyles(writePlan, extractedItems, reportContext = {}) {
  const styleProfile = resolveChecklistStyleProfile({
    reportContext,
    sheetName: writePlan.sheetName
  });

  const skippedCells = extractedItems
    .filter((item) => item?.skipped && item?.outputCell)
    .map((item) => String(item.outputCell).trim().toUpperCase())
    .filter(Boolean);

  await styleChecklistWithLibraries({
    outputPath: writePlan.outputPath,
    sheetName: writePlan.sheetName,
    templatePath: writePlan.templatePath,
    templateSheetName: writePlan.sheetName,
    decimalCells: [],
    percentCells: [],
    skippedCells: styleProfile?.skippedFill ? skippedCells : [],
    valueUpdates: [],
    reportUpdates: []
  });
}

function groupUpdatesBySheet(updates) {
  const grouped = {};

  for (const update of updates) {
    const sheetName = String(update?.sheetName || '').trim();
    const cellAddress = String(update?.cellAddress || '').trim().toUpperCase();
    if (!sheetName || !cellAddress) {
      continue;
    }

    if (!grouped[sheetName]) {
      grouped[sheetName] = [];
    }

    grouped[sheetName].push({
      sheetName,
      cellAddress,
      value: update.value
    });
  }
  return grouped;
}

function resolveWorksheetPathByName(workbookXml, workbookRelsXml) {
  const relationById = {};
  const relRegex = /<Relationship\b([^>]+?)\/>/g;
  let relMatch;
  while ((relMatch = relRegex.exec(workbookRelsXml)) !== null) {
    const attrs = parseXmlAttributes(relMatch[1]);
    if (!attrs.Id || !attrs.Target) {
      continue;
    }

    relationById[attrs.Id] = attrs.Target.replace(/^\/?/, '');
  }

  const worksheetPathByName = {};
  const sheetRegex = /<sheet\b([^>]+?)\/>/g;
  let sheetMatch;
  while ((sheetMatch = sheetRegex.exec(workbookXml)) !== null) {
    const attrs = parseXmlAttributes(sheetMatch[1]);
    const sheetName = attrs.name;
    const relationId = attrs['r:id'];
    const target = relationId ? relationById[relationId] : '';
    if (!sheetName || !target) {
      continue;
    }

    worksheetPathByName[sheetName] = target.startsWith('xl/') ? target : `xl/${target}`;
  }

  return worksheetPathByName;
}

function enableWorkbookRecalculation(workbookXml) {
  const calcPrMatch = /<calcPr\b([^>]*)\/>/i;
  const recalcAttrs = ' fullCalcOnLoad="1" calcOnSave="1" forceFullCalc="1"';

  if (calcPrMatch.test(workbookXml)) {
    return workbookXml.replace(calcPrMatch, (full, attrs) => {
      const cleanedAttrs = String(attrs || '')
        .replace(/\s+fullCalcOnLoad="[^"]*"/i, '')
        .replace(/\s+calcOnSave="[^"]*"/i, '')
        .replace(/\s+forceFullCalc="[^"]*"/i, '');
      return `<calcPr${cleanedAttrs}${recalcAttrs}/>`;
    });
  }

  return workbookXml.replace('</workbook>', `<calcPr${recalcAttrs}/></workbook>`);
}

function ensureFirstSheetActive(workbookXml) {
  const bookViewsRegex = /<bookViews>([\s\S]*?)<\/bookViews>/i;
  if (!bookViewsRegex.test(workbookXml)) {
    return workbookXml;
  }

  return workbookXml.replace(bookViewsRegex, (full) => {
    return full.replace(/<workbookView\b([^>]*)\/>/i, (_, attrs) => {
      const cleanedAttrs = String(attrs || '')
        .replace(/\s+activeTab="[^"]*"/i, '')
        .replace(/\s+firstSheet="[^"]*"/i, '');
      return `<workbookView${cleanedAttrs} activeTab="0" firstSheet="0"/>`;
    });
  });
}

function parseXmlAttributes(attributeText) {
  const attributes = {};
  const attrRegex = /(\S+)=['"]([^'"]*)['"]/g;
  let match;
  while ((match = attrRegex.exec(attributeText)) !== null) {
    attributes[match[1]] = match[2];
  }

  return attributes;
}

function replaceCellValueInWorksheetXml(worksheetXml, cellAddress, value) {
  const escapedAddress = escapeRegExp(cellAddress);
  const selfClosingRegex = new RegExp(`<c\\b([^>]*?\\br=["']${escapedAddress}["'][^>]*)\\/>`, 'i');
  const openCloseRegex = new RegExp(`<c\\b([^>]*?\\br=["']${escapedAddress}["'][^>]*?)>([\\s\\S]*?)<\\/c>`, 'i');

  const buildAttributes = (attributeText, forceInlineStr) => {
    let attrs = String(attributeText || '')
      .replace(/\s+t=["'][^"']*["']/i, '')
      .replace(/\/\s*$/, '')
      .trim();

    if (forceInlineStr) {
      attrs = `${attrs} t="inlineStr"`.trim();
    }

    return attrs ? ` ${attrs}` : '';
  };

  const normalizedValue = toWorksheetValue(value);
  const isNumeric = typeof normalizedValue === 'number' && Number.isFinite(normalizedValue);
  const innerXml = isNumeric
    ? `<v>${normalizedValue}</v>`
    : `<is><t xml:space="preserve">${escapeXml(String(normalizedValue ?? ''))}</t></is>`;

  if (selfClosingRegex.test(worksheetXml)) {
    return worksheetXml.replace(selfClosingRegex, (_, attrs) => `<c${buildAttributes(attrs, !isNumeric)}>${innerXml}</c>`);
  }

  if (openCloseRegex.test(worksheetXml)) {
    return worksheetXml.replace(openCloseRegex, (_, attrs) => `<c${buildAttributes(attrs, !isNumeric)}>${innerXml}</c>`);
  }

  return worksheetXml;
}

function escapeXml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function escapeRegExp(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function readZipText(zip, filePath) {
  const file = zip.getEntry(filePath);
  if (!file) {
    throw new Error(`未找到工作簿结构文件: ${filePath}`);
  }

  return file.getData().toString('utf8');
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

function shouldWriteItem(itemResult) {
  return Boolean(itemResult?.matched) && itemResult.value !== undefined && itemResult.value !== null && itemResult.value !== '';
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
