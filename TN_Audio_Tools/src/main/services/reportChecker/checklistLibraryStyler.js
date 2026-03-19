const AdmZip = require('adm-zip');

const GRAY_SKIP_FILL_XML = '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>';

async function styleChecklistWithLibraries({ outputPath, sheetName, decimalCells, percentCells, skippedCells }) {
  const zip = new AdmZip(outputPath);
  const workbookXml = readZipText(zip, 'xl/workbook.xml');
  const workbookRelsXml = readZipText(zip, 'xl/_rels/workbook.xml.rels');
  const worksheetPathByName = resolveWorksheetPathByName(workbookXml, workbookRelsXml);
  const worksheetPath = worksheetPathByName[sheetName];

  if (!worksheetPath) {
    zip.writeZip(outputPath);
    return 'LIBRARY';
  }

  const requestedSkippedCells = normalizeCellAddresses(skippedCells);
  const requestedDecimalCells = normalizeCellAddresses(decimalCells);
  const requestedPercentCells = normalizeCellAddresses(percentCells);

  if (requestedSkippedCells.length === 0 && requestedDecimalCells.length === 0 && requestedPercentCells.length === 0) {
    zip.writeZip(outputPath);
    return 'LIBRARY';
  }

  let worksheetXml = readZipText(zip, worksheetPath);
  const stylesState = parseStylesState(readZipText(zip, 'xl/styles.xml'));

  const styleCache = new Map();

  for (const cellAddress of requestedDecimalCells) {
    worksheetXml = applyCellStyleMutation({
      worksheetXml,
      stylesState,
      cellAddress,
      styleCache,
      cacheKey: `decimal:${cellAddress}`,
      mutateAttributes: (attrs) => ({
        ...attrs,
        numFmtId: '2',
        applyNumberFormat: '1'
      })
    });
  }

  for (const cellAddress of requestedPercentCells) {
    worksheetXml = applyCellStyleMutation({
      worksheetXml,
      stylesState,
      cellAddress,
      styleCache,
      cacheKey: `percent:${cellAddress}`,
      mutateAttributes: (attrs) => ({
        ...attrs,
        numFmtId: '10',
        applyNumberFormat: '1'
      })
    });
  }

  const skippedFillId = requestedSkippedCells.length > 0 ? ensureFill(stylesState, GRAY_SKIP_FILL_XML) : -1;
  for (const cellAddress of requestedSkippedCells) {
    worksheetXml = applyCellStyleMutation({
      worksheetXml,
      stylesState,
      cellAddress,
      styleCache,
      cacheKey: `skip:${cellAddress}`,
      mutateAttributes: (attrs) => ({
        ...attrs,
        fillId: String(skippedFillId),
        applyFill: '1'
      })
    });
  }

  zip.updateFile(worksheetPath, Buffer.from(worksheetXml, 'utf8'));
  zip.updateFile('xl/styles.xml', Buffer.from(serializeStylesState(stylesState), 'utf8'));
  zip.writeZip(outputPath);
  return 'LIBRARY';
}

function normalizeCellAddresses(cellAddresses) {
  return [...new Set((Array.isArray(cellAddresses) ? cellAddresses : [])
    .map((value) => String(value || '').trim().toUpperCase())
    .filter(Boolean))];
}

function applyCellStyleMutation({ worksheetXml, stylesState, cellAddress, styleCache, cacheKey, mutateAttributes }) {
  const currentStyleId = getCellStyleId(worksheetXml, cellAddress);
  const resolvedCacheKey = `${cacheKey}:${currentStyleId}`;

  let nextStyleId = styleCache.get(resolvedCacheKey);
  if (nextStyleId === undefined) {
    const baseXfXml = stylesState.cellXfs[currentStyleId] || '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    const baseAttributes = parseRootTagAttributes(baseXfXml, 'xf');
    const nextAttributes = mutateAttributes(baseAttributes);
    nextStyleId = ensureCellXf(stylesState, updateRootTagAttributes(baseXfXml, 'xf', nextAttributes));
    styleCache.set(resolvedCacheKey, nextStyleId);
  }

  return setCellStyleId(worksheetXml, cellAddress, nextStyleId);
}

function getCellStyleId(worksheetXml, cellAddress) {
  const cellTag = findCellTag(worksheetXml, cellAddress);
  if (!cellTag) {
    return 0;
  }

  const styleMatch = cellTag.match(/\bs="(\d+)"/i);
  return styleMatch ? Number.parseInt(styleMatch[1], 10) : 0;
}

function setCellStyleId(worksheetXml, cellAddress, styleId) {
  const escapedAddress = escapeRegExp(cellAddress);
  const selfClosingRegex = new RegExp(`<c\\b([^>]*?\\br=["']${escapedAddress}["'][^>]*)\\/>`, 'i');
  const openCloseRegex = new RegExp(`<c\\b([^>]*?\\br=["']${escapedAddress}["'][^>]*?)>([\\s\\S]*?)<\\/c>`, 'i');

  const replaceStyleAttribute = (attributeText) => {
    const currentAttributes = String(attributeText || '').replace(/\s+s=["'][^"']*["']/i, '').trim();
    return currentAttributes ? ` ${currentAttributes} s="${styleId}"` : ` s="${styleId}"`;
  };

  if (selfClosingRegex.test(worksheetXml)) {
    return worksheetXml.replace(selfClosingRegex, (_, attrs) => `<c${replaceStyleAttribute(attrs)}/>`);
  }

  if (openCloseRegex.test(worksheetXml)) {
    return worksheetXml.replace(openCloseRegex, (_, attrs, innerXml) => `<c${replaceStyleAttribute(attrs)}>${innerXml}</c>`);
  }

  return worksheetXml;
}

function findCellTag(worksheetXml, cellAddress) {
  const escapedAddress = escapeRegExp(cellAddress);
  const selfClosingRegex = new RegExp(`<c\\b([^>]*?\\br=["']${escapedAddress}["'][^>]*)\\/>`, 'i');
  const openCloseRegex = new RegExp(`<c\\b([^>]*?\\br=["']${escapedAddress}["'][^>]*?)>([\\s\\S]*?)<\\/c>`, 'i');

  const selfClosingMatch = worksheetXml.match(selfClosingRegex);
  if (selfClosingMatch) {
    return `<c${selfClosingMatch[1]}/>`;
  }

  const openCloseMatch = worksheetXml.match(openCloseRegex);
  if (openCloseMatch) {
    return `<c${openCloseMatch[1]}>`;
  }

  return '';
}

function parseStylesState(stylesXml) {
  const fillsMatch = stylesXml.match(/<fills\b[^>]*count="(\d+)"[^>]*>([\s\S]*?)<\/fills>/i);
  const cellXfsMatch = stylesXml.match(/<cellXfs\b[^>]*count="(\d+)"[^>]*>([\s\S]*?)<\/cellXfs>/i);

  if (!fillsMatch || !cellXfsMatch) {
    throw new Error('styles.xml 缺少 fills 或 cellXfs 节点，无法应用库样式补丁');
  }

  return {
    stylesXml,
    fillsSectionXml: fillsMatch[0],
    fills: collectXmlNodes(fillsMatch[2], /<fill\b[\s\S]*?<\/fill>/gi),
    cellXfsSectionXml: cellXfsMatch[0],
    cellXfs: collectXmlNodes(cellXfsMatch[2], /<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/gi)
  };
}

function serializeStylesState(stylesState) {
  const fillsSectionXml = `<fills count="${stylesState.fills.length}">${stylesState.fills.join('')}</fills>`;
  const cellXfsSectionXml = `<cellXfs count="${stylesState.cellXfs.length}">${stylesState.cellXfs.join('')}</cellXfs>`;

  return stylesState.stylesXml
    .replace(stylesState.fillsSectionXml, fillsSectionXml)
    .replace(stylesState.cellXfsSectionXml, cellXfsSectionXml);
}

function ensureFill(stylesState, fillXml) {
  const normalizedTarget = normalizeXml(fillXml);
  const existingIndex = stylesState.fills.findIndex((item) => normalizeXml(item) === normalizedTarget);
  if (existingIndex >= 0) {
    return existingIndex;
  }

  stylesState.fills.push(fillXml);
  return stylesState.fills.length - 1;
}

function ensureCellXf(stylesState, xfXml) {
  const normalizedTarget = normalizeXml(xfXml);
  const existingIndex = stylesState.cellXfs.findIndex((item) => normalizeXml(item) === normalizedTarget);
  if (existingIndex >= 0) {
    return existingIndex;
  }

  stylesState.cellXfs.push(xfXml);
  return stylesState.cellXfs.length - 1;
}

function collectXmlNodes(source, regex) {
  const matches = source.match(regex);
  return Array.isArray(matches) ? matches : [];
}

function parseAttributes(xmlSnippet) {
  const attributes = {};
  const attributeRegex = /(\S+)="([^"]*)"/g;
  let match;
  while ((match = attributeRegex.exec(xmlSnippet)) !== null) {
    attributes[match[1]] = match[2];
  }
  return attributes;
}

function parseRootTagAttributes(xmlSnippet, tagName) {
  const rootTagMatch = String(xmlSnippet || '').match(new RegExp(`^<${tagName}\\b([^>]*)\/?>`, 'i'));
  if (!rootTagMatch) {
    return {};
  }

  return parseAttributes(rootTagMatch[1]);
}

function updateRootTagAttributes(xmlSnippet, tagName, nextAttributes) {
  const rootTagRegex = new RegExp(`^<${tagName}\\b([^>]*?)(\/?)>`, 'i');
  const rootTagMatch = String(xmlSnippet || '').match(rootTagRegex);
  if (!rootTagMatch) {
    return xmlSnippet;
  }

  const mergedAttributes = {
    numFmtId: '0',
    fontId: '0',
    fillId: '0',
    borderId: '0',
    xfId: '0',
    ...parseAttributes(rootTagMatch[1]),
    ...nextAttributes
  };

  const orderedKeys = ['numFmtId', 'fontId', 'fillId', 'borderId', 'xfId', 'applyNumberFormat', 'applyFill', 'applyFont', 'applyBorder', 'applyAlignment', 'applyProtection', 'quotePrefix', 'pivotButton'];
  const seen = new Set();
  const attributeParts = [];

  for (const key of orderedKeys) {
    if (mergedAttributes[key] === undefined) {
      continue;
    }

    attributeParts.push(`${key}="${mergedAttributes[key]}"`);
    seen.add(key);
  }

  for (const [key, value] of Object.entries(mergedAttributes)) {
    if (seen.has(key) || value === undefined) {
      continue;
    }

    attributeParts.push(`${key}="${value}"`);
  }

  const suffix = rootTagMatch[2] === '/' ? '/>' : '>';
  const rebuiltRootTag = `<${tagName}${attributeParts.length ? ` ${attributeParts.join(' ')}` : ''}${suffix}`;
  return rebuiltRootTag + String(xmlSnippet || '').slice(rootTagMatch[0].length);
}

function resolveWorksheetPathByName(workbookXml, workbookRelsXml) {
  const relationById = {};
  const relRegex = /<Relationship\b([^>]+?)\/>/g;
  let relMatch;
  while ((relMatch = relRegex.exec(workbookRelsXml)) !== null) {
    const attrs = parseAttributes(relMatch[1]);
    if (attrs.Id && attrs.Target) {
      relationById[attrs.Id] = attrs.Target.replace(/^\/?/, '');
    }
  }

  const worksheetPathByName = {};
  const sheetRegex = /<sheet\b([^>]+?)\/>/g;
  let sheetMatch;
  while ((sheetMatch = sheetRegex.exec(workbookXml)) !== null) {
    const attrs = parseAttributes(sheetMatch[1]);
    if (!attrs.name || !attrs['r:id']) {
      continue;
    }

    const target = relationById[attrs['r:id']];
    if (!target) {
      continue;
    }

    worksheetPathByName[attrs.name] = target.startsWith('xl/') ? target : `xl/${target}`;
  }

  return worksheetPathByName;
}

function readZipText(zip, filePath) {
  const entry = zip.getEntry(filePath);
  if (!entry) {
    throw new Error(`未找到 xlsx 结构文件: ${filePath}`);
  }

  return entry.getData().toString('utf8');
}

function normalizeXml(value) {
  return String(value || '').replace(/\s+/g, '');
}

function escapeRegExp(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

module.exports = {
  styleChecklistWithLibraries
};