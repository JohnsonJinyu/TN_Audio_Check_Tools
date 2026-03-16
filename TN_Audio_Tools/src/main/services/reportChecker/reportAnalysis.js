require('./runtimePolyfills');

const cheerio = require('cheerio');

const optionalKeywordSet = new Set(['avg', 'average', 'with', 'check', 'mandatory', 'mand', 'informative', 'calculated', 'value', 'result', 'max', 'min', 'nom', 'vol', 'volume']);
const variantSuffixTokens = ['ecrp', '20up', '25down', '10out', 'max', 'max-1', 'max-2', 'max-3', 'max-4', 'max-5', 'max-6', 'max-7', 'nom', 'min', '1o2', '2o2'];

function normalizeText(value, textNormalizeConfig = {}) {
  if (typeof value !== 'string') {
    return '';
  }

  let normalized = value.replace(/\r/g, ' ').replace(/\n/g, ' ');

  normalized = normalized
    .replace(/\bhasb\b/gi, ' HABAND ')
    .replace(/\bha(?:nb|wb|sb|swb)(max(?:-\d+)?|nom|min)\b/gi, ' HABAND $1 ')
    .replace(/\bha(?:nb|wb|sb|swb)\b/gi, ' HABAND ')
    .replace(/\bbgn\b/gi, ' NOISE ')
    .replace(/\bbackground\s+noise\b/gi, ' NOISE ')
    .replace(/\bin\s+noise\b/gi, ' NOISE ')
    .replace(/\bmxx-(\d+)/gi, ' MAX-$1 ')
    .replace(/\breceiv(?:e|ing|eing)\b/gi, ' RCV ')
    .replace(/\bsending\b/gi, ' SND ')
    .replace(/\bfrequency\b/gi, ' FREQ ')
    .replace(/\bcharact(?:er)?\.?\b/gi, ' CHAR ')
    .replace(/\baverage\b/gi, ' AVG ')
    .replace(/\bnominal\b/gi, ' NOM ')
    .replace(/\bsinglevalue\b/gi, ' SINGLE VALUE ')
    .replace(/\bvolume\b/gi, ' VOL ')
    .replace(/\bmandatory\b/gi, ' MAND ');

  if (textNormalizeConfig.removeNumberPrefix) {
    normalized = normalized.replace(/^\s*\d+(?:\.\d+)*\s*/, '');
  }

  if (textNormalizeConfig.trimSpecialChar) {
    normalized = normalized
      .replace(/\b(?:ok|not ok)\b\s*$/i, '')
      .replace(/[|]+/g, ' ')
      .replace(/[\u00a0]+/g, ' ')
      .trim();
  }

  if (textNormalizeConfig.multiSpaceToSingle) {
    normalized = normalized.replace(/\s+/g, ' ');
  }

  normalized = normalized.trim();

  if (textNormalizeConfig.caseInsensitive) {
    normalized = normalized.toLowerCase();
  }

  return normalized;
}

function pickForbiddenSuffixes(item, globalMatchConfig) {
  if (Array.isArray(item.forbiddenSuffix)) {
    return item.forbiddenSuffix;
  }

  return Array.isArray(globalMatchConfig.globalForbiddenSuffix)
    ? globalMatchConfig.globalForbiddenSuffix
    : [];
}

function hasAllKeywords(text, keywords, textNormalizeConfig) {
  if (!Array.isArray(keywords) || keywords.length === 0) {
    return true;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);
  return keywords.every((keyword) => normalizedText.includes(normalizeText(keyword, textNormalizeConfig)));
}

function stripOptionalKeywords(keywords, textNormalizeConfig) {
  if (!Array.isArray(keywords) || keywords.length === 0) {
    return [];
  }

  return keywords.filter((keyword) => {
    const normalizedKeyword = normalizeText(keyword, textNormalizeConfig);
    return normalizedKeyword && !optionalKeywordSet.has(normalizedKeyword);
  });
}

function getKeywordMatchTier(reportData, item, text, textNormalizeConfig, allowFallback = false) {
  if (hasAllKeywords(text, item.coreKeywords, textNormalizeConfig)) {
    return 4;
  }

  const isLegacyTextOnlyReport = !Array.isArray(reportData?.tables) || reportData.tables.length === 0;
  if ((isLegacyTextOnlyReport || allowFallback) && hasAllKeywords(text, item.fallbackKeywords, textNormalizeConfig)) {
    return 3;
  }

  const relaxedCoreKeywords = stripOptionalKeywords(item.coreKeywords, textNormalizeConfig);
  if (relaxedCoreKeywords.length >= 3 && hasAllKeywords(text, relaxedCoreKeywords, textNormalizeConfig)) {
    return 2;
  }

  const relaxedFallbackKeywords = stripOptionalKeywords(item.fallbackKeywords, textNormalizeConfig);
  if ((isLegacyTextOnlyReport || allowFallback) && relaxedFallbackKeywords.length >= 3 && hasAllKeywords(text, relaxedFallbackKeywords, textNormalizeConfig)) {
    return 1;
  }

  return 0;
}

function itemMatchesKeywords(reportData, item, text, textNormalizeConfig, allowFallback = false) {
  return getKeywordMatchTier(reportData, item, text, textNormalizeConfig, allowFallback) > 0;
}

function getRowDescriptorText(rowContext) {
  if (rowContext && Array.isArray(rowContext.cells) && rowContext.cells.length > 0) {
    if (rowContext.cells.length === 1) {
      const singleCellText = rowContext.cells[0] || rowContext.text || '';
      const statusSplit = singleCellText.split(/\b(?:not\s+ok|ok|done)\b/i)[0]?.trim();
      return statusSplit || singleCellText;
    }

    return rowContext.cells[0] || rowContext.text || '';
  }

  const rawText = String(rowContext?.text || '');
  const pipeSplitText = rawText.split('|')[0].trim();
  const statusSplitText = pipeSplitText.split(/\b(?:not\s+ok|ok|done)\b/i)[0]?.trim();
  return statusSplitText || pipeSplitText;
}

function hasRequiredSuffixes(text, suffixes, textNormalizeConfig) {
  if (!Array.isArray(suffixes) || suffixes.length === 0) {
    return true;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);

  return suffixes.every((suffix) => {
    const normalizedSuffix = normalizeText(suffix, textNormalizeConfig);

    if (normalizedSuffix === 'max') {
      return /(^|[^a-z0-9])max([^a-z0-9]|$)/i.test(normalizedText) && !/max-\d/i.test(normalizedText);
    }

    if (normalizedSuffix === 'min') {
      return /(^|[^a-z0-9])min([^a-z0-9]|$)/i.test(normalizedText)
        && !/min-\d/i.test(normalizedText)
        && !/min\.?\s+dist/i.test(normalizedText)
        && !/minimal/i.test(normalizedText);
    }

    return normalizedText.includes(normalizedSuffix);
  });
}

function hasExactRowKeywords(text, exactKeywords, textNormalizeConfig) {
  if (!Array.isArray(exactKeywords) || exactKeywords.length === 0) {
    return true;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);
  return exactKeywords.every((keyword) => normalizedText.includes(normalizeText(keyword, textNormalizeConfig)));
}

function hasForbiddenSuffix(text, suffixes, textNormalizeConfig) {
  if (!Array.isArray(suffixes) || suffixes.length === 0) {
    return false;
  }

  const normalizedText = normalizeText(text, textNormalizeConfig);
  return suffixes.some((suffix) => normalizedText.includes(normalizeText(suffix, textNormalizeConfig)));
}

function extractStatus(text) {
  if (!text) {
    return null;
  }

  if (/\bnot\s+ok\b/i.test(text)) {
    return 'Fail';
  }

  if (/\bok\b/i.test(text)) {
    return 'Pass';
  }

  return null;
}

function isMeasurementObjectCell(text) {
  const normalized = String(text || '').trim();
  return /^[A-Za-z0-9]+(?:_[A-Za-z0-9+-]+)+$/.test(normalized);
}

function extractNumberTokens(text) {
  if (!text) {
    return [];
  }

  const normalizedText = String(text)
    .replace(/(\d)\s*\.\s*(\d)/g, '$1.$2')
    .replace(/(\d\.\d+)\s+(\d)/g, '$1$2');

  return normalizedText.match(/-?\d+(?:\.\d+)?(?:\s*(?:%|dB(?:m0|Pa)?|ms|Hz|MOS|s))?/gi) || [];
}

function getCandidateScore(rowContext) {
  const structuredScore = Array.isArray(rowContext.cells) && rowContext.cells.length > 0 ? 1000 : 0;
  const numericCellScore = Array.isArray(rowContext.cells)
    ? rowContext.cells.filter((cell) => !isMeasurementObjectCell(cell) && extractNumberTokens(cell).length > 0).length * 100
    : 0;
  const headerScore = Array.isArray(rowContext.headers) && rowContext.headers.length > 0 ? 50 : 0;
  const statusScore = extractStatus(rowContext.text) ? 500 : 0;
  return structuredScore + numericCellScore + headerScore + statusScore + rowContext.text.length;
}

function getPrimaryRows(reportData) {
  return Array.isArray(reportData?.tableRows) ? reportData.tableRows : [];
}

function getFallbackRows(reportData) {
  return Array.isArray(reportData?.fallbackRows) ? reportData.fallbackRows : [];
}

function getRowsForMatching(reportData, includeFallbackRows = false) {
  return includeFallbackRows
    ? [...getPrimaryRows(reportData), ...getFallbackRows(reportData)]
    : getPrimaryRows(reportData);
}

function getVariantPenalty(item, descriptorText, textNormalizeConfig) {
  if (!item || !descriptorText) {
    return 0;
  }

  const hasExplicitRequiredSuffix = Array.isArray(item.requiredSuffix) && item.requiredSuffix.length > 0;
  if (hasExplicitRequiredSuffix) {
    return 0;
  }

  const normalizedDescriptor = normalizeText(descriptorText, textNormalizeConfig);
  const configuredKeywords = [
    ...(Array.isArray(item.coreKeywords) ? item.coreKeywords : []),
    ...(Array.isArray(item.fallbackKeywords) ? item.fallbackKeywords : []),
    ...(Array.isArray(item.exactRowKeywords) ? item.exactRowKeywords : [])
  ].map((keyword) => normalizeText(keyword, textNormalizeConfig));

  const matchedVariantTokens = variantSuffixTokens.filter((token) => normalizedDescriptor.includes(normalizeText(token, textNormalizeConfig)));
  if (matchedVariantTokens.length === 0) {
    return 0;
  }

  const hasConfiguredVariant = matchedVariantTokens.some((token) => configuredKeywords.includes(normalizeText(token, textNormalizeConfig)));
  return hasConfiguredVariant ? 0 : -4000;
}

function extractMeasurementObject(rawText) {
  const match = rawText.match(/Measurement Object\s+([A-Za-z0-9_+\-]+)/i);
  return match ? match[1] : null;
}

function getRowNameCandidates(rowName) {
  const aliasMap = {
    crossroad: ['Crossroad', 'Crossroads'],
    tstation: ['TStation', 'Train Station', 'Train'],
    fullsizecar: ['FullsizeCar', 'Fullsize Car', 'Car'],
    callcenter: ['Callcenter', 'Call Center']
  };

  const normalized = String(rowName || '').replace(/\s+/g, '').toLowerCase();
  return aliasMap[normalized] || [rowName];
}

function getAmbientNoiseSceneCandidates(rowName) {
  const aliasMap = {
    pub: ['Pub', 'Pub Noise'],
    road: ['Road', 'Outside Traffic Road'],
    crossroad: ['Crossroad', 'Crossroads', 'Xroad', 'Outside Traffic Crossroads'],
    tstation: ['Train Station', 'TStation', 'Train'],
    fullsizecar: ['Fullsize Car 130 km/h', 'Fullsize Car', 'FullsizeCar', 'Car'],
    cafeteria: ['Cafeteria', 'Cafeteria Noise', 'Count'],
    mensa: ['Mensa'],
    callcenter: ['Callcenter', 'Call Center', 'CallC', 'CallCenter', 'Work Noise Office Callcenter']
  };

  const normalized = String(rowName || '').replace(/\s+/g, '').toLowerCase();
  return aliasMap[normalized] || getRowNameCandidates(rowName);
}

function sanitizeAmbientNoiseSceneLabel(value) {
  return String(value || '')
    .replace(/\t+/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\bDone\b.*$/i, '')
    .trim();
}

function getAmbientNoiseSceneKey(value) {
  const normalized = String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');

  if (!normalized) {
    return '';
  }

  if (normalized.includes('crossroad') || normalized.includes('xroad')) {
    return 'crossroad';
  }

  if (normalized.includes('trainstation') || normalized === 'tstation' || normalized === 'train') {
    return 'tstation';
  }

  if (normalized.includes('fullsizecar') || normalized === 'car') {
    return 'fullsizecar';
  }

  if (normalized.includes('callcenter')) {
    return 'callcenter';
  }

  if (normalized === 'callc') {
    return 'callcenter';
  }

  if (normalized.includes('cafeteria')) {
    return 'cafeteria';
  }

  if (normalized === 'count') {
    return 'cafeteria';
  }

  if (normalized.includes('mensa')) {
    return 'mensa';
  }

  if (normalized.includes('pub')) {
    return 'pub';
  }

  if (normalized === 'road' || normalized.includes('outsidetrafficroad')) {
    return 'road';
  }

  return normalized;
}

function getAmbientNoiseMetricKey(value) {
  const normalized = normalizeText(value, {
    removeNumberPrefix: true,
    multiSpaceToSingle: true,
    caseInsensitive: true,
    trimSpecialChar: true
  });

  if (normalized.includes('s-mos')) {
    return 'smos';
  }

  if (normalized.includes('n-mos')) {
    return 'nmos';
  }

  if (normalized.includes('g-mos')) {
    return 'gmos';
  }

  return null;
}

function extractMetricRowValue(cells) {
  if (!Array.isArray(cells) || cells.length === 0) {
    return null;
  }

  for (const cell of cells.slice(1)) {
    if (!cell || /\b\d{1,2}\/\d{1,2}\/\d{4}\b/.test(cell) || isStatusCell(cell) || isLikelyReportMetadataCell(cell)) {
      continue;
    }

    const collapsedNumericCell = String(cell).replace(/\s+/g, '');
    if (/^-?\d+(?:\.\d+)?$/.test(collapsedNumericCell)) {
      return collapsedNumericCell;
    }

    const tokens = extractNumberTokens(cell);
    const value = tokens.length > 0 ? tokens[0].trim() : null;
    if (value) {
      return value;
    }
  }

  return null;
}

function isStatusCell(text) {
  const normalized = normalizeText(text, {
    removeNumberPrefix: true,
    multiSpaceToSingle: true,
    caseInsensitive: true,
    trimSpecialChar: true
  });

  return ['ok', 'not ok', 'done', 'pass', 'fail'].includes(normalized);
}

function isLikelyReportMetadataCell(text) {
  const normalized = String(text || '').trim();
  if (!normalized) {
    return false;
  }

  const collapsed = normalized.replace(/\s+/g, '');

  if (/^[A-Za-z0-9]+(?:_[A-Za-z0-9+-]+){2,}$/.test(normalized)) {
    return true;
  }

  if (/^[A-Za-z0-9]+(?:_[A-Za-z0-9+-]+){2,}$/.test(collapsed)) {
    return true;
  }

  return /\.(?:doc|docx|xls|xlsx)$/i.test(normalized);
}

function getLastMeaningfulNumericToken(text) {
  const cleanedText = String(text || '')
    .replace(/^\d+(?:\.\d+)+(?:[a-z])?\s+/i, '')
    .replace(/[A-Za-z0-9]+(?:_[A-Za-z0-9+-]+)+$/g, '')
    .replace(/[A-Za-z0-9]+(?:\s+[A-Za-z0-9]+)*(?:_[A-Za-z0-9+-]+)+$/g, '')
    .trim();

  const normalizedNumericText = cleanedText
    .replace(/(\d)\s*\.\s*(\d)/g, '$1.$2')
    .replace(/(\d\.\d+)\s+(\d)/g, '$1$2');

  const tokens = extractNumberTokens(normalizedNumericText);
  return tokens.length > 0 ? tokens[tokens.length - 1].trim() : null;
}

function extractRowNumericCell(rowContext) {
  if (!rowContext || !Array.isArray(rowContext.cells) || rowContext.cells.length === 0) {
    return null;
  }

  const numericCell = rowContext.cells
    .slice()
    .reverse()
    .find((cell) => {
      if (!cell || isMeasurementObjectCell(cell) || isStatusCell(cell) || isLikelyReportMetadataCell(cell)) {
        return false;
      }

      return extractNumberTokens(cell).length > 0;
    });

  return numericCell ? getLastMeaningfulNumericToken(numericCell) : null;
}

function parseNumericValue(text) {
  if (!text) {
    return null;
  }

  const normalizedText = String(text)
    .replace(/(\d)\s*\.\s*(\d)/g, '$1.$2')
    .replace(/(\d\.\d+)\s+(\d)/g, '$1$2');

  const match = normalizedText.match(/-?\d+(?:\.\d+)?/);
  if (!match) {
    return null;
  }

  const parsed = Number.parseFloat(match[0]);
  return Number.isNaN(parsed) ? null : parsed;
}

function formatComputedNumber(value) {
  if (!Number.isFinite(value)) {
    return null;
  }

  const rounded = Math.round(value * 100) / 100;
  return String(rounded);
}

function makeSourcePreview(text) {
  if (!text) {
    return '';
  }

  return text.replace(/\s+/g, ' ').trim().slice(0, 160);
}

function scoreHeaderPatternMatch(normalizedHeader, normalizedPattern) {
  if (!normalizedHeader || !normalizedPattern || !normalizedHeader.includes(normalizedPattern)) {
    return 0;
  }

  if (normalizedHeader === normalizedPattern) {
    return 10000 + normalizedPattern.length;
  }

  let score = 1000 + normalizedPattern.length;

  if (normalizedHeader.startsWith(`${normalizedPattern} `) || normalizedHeader.endsWith(` ${normalizedPattern}`)) {
    score += 250;
  }

  if (normalizedHeader.startsWith(normalizedPattern)) {
    score += 100;
  }

  if (normalizedHeader.includes('description') && !normalizedPattern.includes('description')) {
    score -= 1000;
  }

  return score;
}

function extractAmbientNoiseAverageBlocks(rawText) {
  if (!rawText) {
    return [];
  }

  const lines = rawText.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const blockMap = new Map();
  let currentSceneKey = null;

  for (const line of lines) {
    const sceneMatch = /(?:7|9)\.12\.1\s+Quality in ambient noise,\s*([^\t\n\r]+?)(?:\t|\s{2,}|$)/i.exec(line);
    if (sceneMatch) {
      const sceneLabel = sanitizeAmbientNoiseSceneLabel(sceneMatch[1] || '');
      const sceneKey = getAmbientNoiseSceneKey(sceneLabel);

      if (sceneKey) {
        currentSceneKey = sceneKey;
        if (!blockMap.has(sceneKey)) {
          blockMap.set(sceneKey, {
            sceneLabel,
            gmos: null,
            nmos: null,
            smos: null,
            sourcePreview: ''
          });
        }
      }
    }

    if (!currentSceneKey || !blockMap.has(currentSceneKey)) {
      continue;
    }

    const block = blockMap.get(currentSceneKey);
    const metricPatterns = [
      ['gmos', /G-MOS \(Average(?:,\s*TS\s*103\s*(?:106|281)(?:\s*\(Model A\))?)?\)[:\t ]+([0-9.]+)/i],
      ['nmos', /^(?!Corrected\b).*N-MOS \(Average(?:,\s*TS\s*103\s*(?:106|281)(?:\s*\(Model A\))?)?\)[:\t ]+([0-9.]+)/i],
      ['smos', /S-MOS \(Average(?:,\s*TS\s*103\s*(?:106|281)(?:\s*\(Model A\))?)?\)[:\t ]+([0-9.]+)/i]
    ];

    metricPatterns.forEach(([metricKey, pattern]) => {
      const value = pattern.exec(line)?.[1] || null;
      if (value) {
        block[metricKey] = value;
        block.sourcePreview = makeSourcePreview(`${block.sourcePreview} ${line}`);
      }
    });
  }

  return Array.from(blockMap.values()).filter((block) => block.gmos || block.nmos || block.smos);
}

function getHeaderBasedCell(rowContext, headerPatterns, textNormalizeConfig) {
  if (!rowContext || !Array.isArray(rowContext.headers) || !Array.isArray(rowContext.cells)) {
    return null;
  }

  const normalizedPatterns = headerPatterns.map((pattern) => normalizeText(pattern, textNormalizeConfig));
  let bestMatch = null;

  for (let index = 0; index < rowContext.headers.length; index += 1) {
    const normalizedHeader = normalizeText(rowContext.headers[index] || '', textNormalizeConfig);
    if (!normalizedHeader) {
      continue;
    }

    normalizedPatterns.forEach((pattern) => {
      const score = scoreHeaderPatternMatch(normalizedHeader, pattern);
      if (!score) {
        return;
      }

      if (!bestMatch || score > bestMatch.score) {
        bestMatch = {
          score,
          cell: rowContext.cells[index] || null
        };
      }
    });
  }

  if (bestMatch) {
    return bestMatch.cell;
  }

  return null;
}

function extractSummaryValue(rowContext, textNormalizeConfig) {
  if (!rowContext) {
    return null;
  }

  const directCell = getHeaderBasedCell(
    rowContext,
    ['single value', 'measured value', 'value', 'actual value', 'result'],
    textNormalizeConfig
  );

  const directCellValue = getLastMeaningfulNumericToken(directCell);
  if (directCellValue) {
    return directCellValue;
  }

  const numericCellValue = extractRowNumericCell(rowContext);
  if (numericCellValue) {
    return numericCellValue;
  }

  const fallbackValue = getLastMeaningfulNumericToken(rowContext.text);
  if (fallbackValue) {
    return fallbackValue;
  }

  const status = extractStatus(directCell || rowContext.text);
  if (status) {
    return status;
  }

  return directCell || rowContext.cells[rowContext.cells.length - 1] || null;
}

function extractFormulaValue(reportData, rowContext, item, textNormalizeConfig, extractedResultsByItemId) {
  if (!rowContext || !item.formula) {
    return null;
  }

  let rawValue = null;

  if (item.formula.targetField === 'marginValue') {
    rawValue = getHeaderBasedCell(rowContext, ['margin'], textNormalizeConfig);
  }

  if (!rawValue) {
    rawValue = extractRowNumericCell(rowContext) || getLastMeaningfulNumericToken(rowContext.text);
  }

  const numericValue = parseNumericValue(rawValue);
  if (numericValue === null) {
    return null;
  }

  const shouldUseBaseItemValue = item.formula?.useBaseItemValue === true
    && rowContext?.sourceKind
    && rowContext.sourceKind !== 'html-table';

  if (shouldUseBaseItemValue && item.formula.baseItemId && extractedResultsByItemId instanceof Map) {
    const baseResult = extractedResultsByItemId.get(Number(item.formula.baseItemId));
    const baseNumericValue = parseNumericValue(baseResult?.value);

    if (baseNumericValue !== null) {
      return formatComputedNumber(baseNumericValue - numericValue);
    }
  }

  return formatComputedNumber(Number(item.formula.standardValue) - numericValue);
}

function findLineWindow(lines, startIndex, size = 8) {
  if (!Array.isArray(lines) || lines.length === 0) {
    return [];
  }

  const safeIndex = Math.max(0, startIndex);
  return lines.slice(safeIndex, Math.min(lines.length, safeIndex + size));
}

function resolveAnchorValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId) {
  const anchorText = item.anchorConfig?.anchorText;
  if (!anchorText) {
    return { matched: false, reason: '缺少 anchorText 配置' };
  }

  const normalizedAnchor = normalizeText(anchorText, textNormalizeConfig);
  const anchorIndexes = reportData.lines
    .map((line, index) => ({ line, index }))
    .filter(({ line }) => normalizeText(line, textNormalizeConfig).includes(normalizedAnchor));

  if (anchorIndexes.length === 0) {
    const fallbackRow = resolveRowBasedValue(reportData, {
      ...item,
      extractType: 'summary_table_match'
    }, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);

    if (fallbackRow?.matched) {
      return {
        ...fallbackRow,
        sourceType: 'anchor-missing-fallback-row'
      };
    }

    return { matched: false, reason: '未找到锚点文本' };
  }

  const rankedAnchors = anchorIndexes
    .map(({ index }) => {
      const windowLines = findLineWindow(reportData.lines, Math.max(0, index - 8), 18);
      const joinedWindow = windowLines.join(' ');
      const activeKeywords = itemMatchesKeywords(reportData, item, joinedWindow, textNormalizeConfig, true)
        ? (hasAllKeywords(joinedWindow, item.coreKeywords, textNormalizeConfig) ? item.coreKeywords : item.fallbackKeywords)
        : [];
      const keywordScore = Array.isArray(activeKeywords)
        ? activeKeywords.filter((keyword) => normalizeText(joinedWindow, textNormalizeConfig).includes(normalizeText(keyword, textNormalizeConfig))).length
        : 0;

      return { index, windowLines, joinedWindow, keywordScore };
    })
    .sort((left, right) => right.keywordScore - left.keywordScore);

  for (const candidate of rankedAnchors) {
    if (!itemMatchesKeywords(reportData, item, candidate.joinedWindow, textNormalizeConfig, true)) {
      continue;
    }

    const dbAtMatch = candidate.joinedWindow.match(/(-?\d+(?:\.\d+)?)\s*dB\s+at/i);
    let value = dbAtMatch ? dbAtMatch[1] : null;

    if (!value) {
      const numericToken = getLastMeaningfulNumericToken(candidate.joinedWindow);
      value = numericToken ? numericToken.match(/-?\d+(?:\.\d+)?/)?.[0] : null;
    }

    if (!value) {
      continue;
    }

    const numericValue = Number.parseFloat(value);
    const range = item.anchorConfig?.valueRange;
    if (Array.isArray(range) && range.length === 2) {
      if (numericValue < range[0] || numericValue > range[1]) {
        continue;
      }
    }

    return {
      matched: true,
      value,
      sourcePreview: makeSourcePreview(candidate.joinedWindow),
      sourceType: 'anchor-window'
    };
  }

  const fallbackRow = resolveRowBasedValue(reportData, {
    ...item,
    extractType: 'summary_table_match'
  }, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);

  if (fallbackRow?.matched) {
    return {
      ...fallbackRow,
      sourceType: 'anchor-fallback-row'
    };
  }

  return { matched: false, reason: '锚点附近未找到符合规则的数值' };
}

function resolveRegexValue(reportData, item, textNormalizeConfig, globalMatchConfig) {
  if (!item.regexConfig?.matchRegex) {
    return { matched: false, reason: '缺少正则配置' };
  }

  const regex = new RegExp(item.regexConfig.matchRegex, 'i');
  const forbiddenSuffixes = pickForbiddenSuffixes(item, globalMatchConfig);
  const range = item.regexConfig.valueRange;
  const candidates = [];

  for (const rowContext of getRowsForMatching(reportData, true)) {
    const keywordTier = getKeywordMatchTier(reportData, item, rowContext.text, textNormalizeConfig, rowContext.sourceKind !== 'html-table');
    if (keywordTier === 0) {
      continue;
    }

    const descriptorText = getRowDescriptorText(rowContext);
    if (!hasRequiredSuffixes(descriptorText, item.requiredSuffix, textNormalizeConfig)) {
      continue;
    }

    if (hasForbiddenSuffix(descriptorText, forbiddenSuffixes, textNormalizeConfig)) {
      continue;
    }

    if (!hasExactRowKeywords(rowContext.text, item.exactRowKeywords, textNormalizeConfig)) {
      continue;
    }

    const match = rowContext.text.match(regex);
    if (!match) {
      continue;
    }

    const groupIndex = Number(item.regexConfig.valueGroupIndex || 1);
    const value = match[groupIndex] || match[0];
    const numericValue = parseNumericValue(value);

    if (Array.isArray(range) && range.length === 2 && numericValue !== null) {
      if (numericValue < range[0] || numericValue > range[1]) {
        continue;
      }
    }

    candidates.push({
      matched: true,
        value,
      sourcePreview: makeSourcePreview(rowContext.text),
      sourceType: 'regex-row',
        score: keywordTier * 10000 + 2000 + getCandidateScore(rowContext)
    });
  }

  const hasAnchorConstraints = (Array.isArray(item.exactRowKeywords) && item.exactRowKeywords.length > 0)
    || (Array.isArray(item.requiredSuffix) && item.requiredSuffix.length > 0);

  const anchorIndexes = hasAnchorConstraints
    ? reportData.lines
      .map((line, index) => ({ line, index }))
      .filter(({ line }) => {
        if (!itemMatchesKeywords(reportData, item, line, textNormalizeConfig, true)) {
          return false;
        }

        if (!hasRequiredSuffixes(line, item.requiredSuffix, textNormalizeConfig)) {
          return false;
        }

        if (!hasExactRowKeywords(line, item.exactRowKeywords, textNormalizeConfig)) {
          return false;
        }

        return true;
      })
      .map(({ index }) => index)
    : [];

  const lineSearchIndexes = anchorIndexes.length > 0
    ? anchorIndexes
    : reportData.lines.map((_, index) => index);

  for (const lineIndex of lineSearchIndexes) {
    const contextWindow = anchorIndexes.length > 0
      ? findLineWindow(reportData.lines, lineIndex, 8)
      : findLineWindow(reportData.lines, Math.max(0, lineIndex - 2), 12);
    const joinedWindow = contextWindow.join('\n');

    if (!itemMatchesKeywords(reportData, item, joinedWindow, textNormalizeConfig, true)) {
      continue;
    }

    if (!hasRequiredSuffixes(joinedWindow, item.requiredSuffix, textNormalizeConfig)) {
      continue;
    }

    if (hasForbiddenSuffix(joinedWindow, forbiddenSuffixes, textNormalizeConfig)) {
      continue;
    }

    for (const candidateLine of contextWindow) {
      const match = candidateLine.match(regex);
      if (!match) {
        continue;
      }

      const candidateIndex = contextWindow.indexOf(candidateLine);
      const groupIndex = Number(item.regexConfig.valueGroupIndex || 1);
      const value = match[groupIndex] || match[0];
      const numericValue = parseNumericValue(value);

      if (Array.isArray(range) && range.length === 2 && numericValue !== null) {
        if (numericValue < range[0] || numericValue > range[1]) {
          continue;
        }
      }

      const requiredSuffixes = Array.isArray(item.requiredSuffix) ? item.requiredSuffix : [];
      const lineHasRequiredSuffix = hasRequiredSuffixes(candidateLine, requiredSuffixes, textNormalizeConfig);
      const sameLineKeywordScore = itemMatchesKeywords(reportData, item, candidateLine, textNormalizeConfig, true) ? 500 : 0;
      const distanceFromAnchor = anchorIndexes.length > 0 ? candidateIndex : 0;
      const nearestSuffixDistance = requiredSuffixes.length === 0
        ? 0
        : contextWindow.reduce((bestDistance, line, windowIndex) => {
          if (windowIndex > candidateIndex) {
            return bestDistance;
          }

          if (!hasRequiredSuffixes(line, requiredSuffixes, textNormalizeConfig)) {
            return bestDistance;
          }

          const distance = candidateIndex - windowIndex;
          return Math.min(bestDistance, distance);
        }, Number.POSITIVE_INFINITY);

      const proximityScore = Number.isFinite(nearestSuffixDistance)
        ? Math.max(0, 200 - nearestSuffixDistance * 60)
        : 0;

      candidates.push({
        matched: true,
        value,
        sourcePreview: makeSourcePreview(candidateLine),
        sourceType: 'regex-line',
        score: (lineHasRequiredSuffix ? 1000 : 0) + sameLineKeywordScore + proximityScore + candidateLine.length + (anchorIndexes.length > 0 ? Math.max(0, 300 - distanceFromAnchor * 80) : 0)
      });
    }
  }

  if (candidates.length > 0) {
    return candidates.sort((left, right) => right.score - left.score)[0];
  }

  if (reportData?.reportFormat === 'xlsx') {
    const rowNameHint = getRegexRowNameHint(item);
    const fallbackRow = getRowsForMatching(reportData, true)
      .filter((rowContext) => {
        const keywordTier = getKeywordMatchTier(reportData, item, rowContext.text, textNormalizeConfig, rowContext.sourceKind !== 'html-table');
        if (keywordTier === 0) {
          return false;
        }

        const descriptorText = getRowDescriptorText(rowContext);
        if (!hasRequiredSuffixes(descriptorText, item.requiredSuffix, textNormalizeConfig)) {
          return false;
        }

        if (hasForbiddenSuffix(descriptorText, forbiddenSuffixes, textNormalizeConfig)) {
          return false;
        }

        if (!hasExactRowKeywords(rowContext.text, item.exactRowKeywords, textNormalizeConfig)) {
          return false;
        }

        if (!rowNameHint) {
          return true;
        }

        return normalizeText(rowContext.text, textNormalizeConfig).includes(normalizeText(rowNameHint, textNormalizeConfig));
      })
      .sort((left, right) => getCandidateScore(right) - getCandidateScore(left))[0];

    if (fallbackRow) {
      const value = extractSummaryValue(fallbackRow, textNormalizeConfig);
      if (value) {
        return {
          matched: true,
          value,
          sourcePreview: makeSourcePreview(fallbackRow.text),
          sourceType: 'regex-row-fallback'
        };
      }
    }
  }

  return { matched: false, reason: '未匹配到正则目标文本' };
}

function getRegexRowNameHint(item) {
  const matchRegex = String(item?.regexConfig?.matchRegex || '');
  if (!matchRegex) {
    return null;
  }

  if (/ST Class A1\\\+A2/i.test(matchRegex)) {
    return 'ST Class A1+A2';
  }

  if (/DT Class A1\\\+A2/i.test(matchRegex)) {
    return 'DT Class A1+A2';
  }

  if (/DT Class A1/i.test(matchRegex)) {
    return 'DT Class A1';
  }

  return null;
}

function resolveAmbientNoiseAverageValue(reportData, item, textNormalizeConfig) {
  const blocks = Array.isArray(reportData.ambientNoiseBlocks) ? reportData.ambientNoiseBlocks : [];
  const tableConfig = item.tableConfig;

  if (blocks.length === 0 || !tableConfig?.rowNameMatch || !tableConfig?.targetColumnName) {
    return null;
  }

  const normalizedChecklistDesc = normalizeText(item.checklistDesc || '', textNormalizeConfig);
  if (!normalizedChecklistDesc.includes('ambient noise')) {
    return null;
  }

  const normalizedMetric = normalizeText(tableConfig.targetColumnName, textNormalizeConfig);
  const metricKey = normalizedMetric.includes('s-mos')
    ? 'smos'
    : normalizedMetric.includes('n-mos')
      ? 'nmos'
      : normalizedMetric.includes('g-mos')
        ? 'gmos'
        : null;

  if (!metricKey) {
    return null;
  }

  const candidateSceneKeys = getAmbientNoiseSceneCandidates(tableConfig.rowNameMatch)
    .map((candidate) => getAmbientNoiseSceneKey(candidate))
    .filter(Boolean);

  const targetBlock = blocks.find((block) => {
    const sceneKey = getAmbientNoiseSceneKey(block.sceneLabel);
    return candidateSceneKeys.includes(sceneKey);
  });

  if (!targetBlock || !targetBlock[metricKey]) {
    return null;
  }

  return {
    matched: true,
    value: targetBlock[metricKey],
    sourcePreview: targetBlock.sourcePreview,
    sourceType: 'ambient-average-block'
  };
}

function resolveAmbientNoiseOverallAverageValue(reportData, item, textNormalizeConfig) {
  const normalizedChecklistDesc = normalizeText(item.checklistDesc || '', textNormalizeConfig);
  if (!normalizedChecklistDesc.includes('ambient noise') || !normalizedChecklistDesc.includes('average')) {
    return null;
  }

  const metricKey = normalizedChecklistDesc.includes('s-mos')
    ? 'smos'
    : normalizedChecklistDesc.includes('n-mos')
      ? 'nmos'
      : normalizedChecklistDesc.includes('g-mos')
        ? 'gmos'
        : null;

  if (!metricKey) {
    return null;
  }

  const metricPattern = metricKey === 'smos'
    ? /average\s+s-mos/i
    : metricKey === 'nmos'
      ? /average\s+n-mos/i
      : /average\s+g-mos/i;
  const calculatedValuePattern = /calculated\s+value[^0-9-]*(-?\d+(?:\.\d+)?)/i;

  for (const rowContext of getRowsForMatching(reportData, true)) {
    const rowText = String(rowContext?.text || '').trim();
    if (!rowText || !metricPattern.test(rowText)) {
      continue;
    }

    if (!/calculated\s+value|mos\s*\(avg\)/i.test(rowText)) {
      continue;
    }

    const value = calculatedValuePattern.exec(rowText)?.[1]
      || extractSummaryValue(rowContext, textNormalizeConfig)
      || getLastMeaningfulNumericToken(rowText);
    const numericValue = parseNumericValue(value);
    if (numericValue === null || numericValue > 5) {
      continue;
    }

    return {
      matched: true,
      value: String(numericValue),
      sourcePreview: makeSourcePreview(rowText),
      sourceType: 'ambient-average-row'
    };
  }

  for (let lineIndex = 0; lineIndex < reportData.lines.length; lineIndex += 1) {
    const line = String(reportData.lines[lineIndex] || '').trim();
    if (!metricPattern.test(line)) {
      continue;
    }

    const inlineValue = calculatedValuePattern.exec(line)?.[1] || getLastMeaningfulNumericToken(line);
    const inlineNumericValue = parseNumericValue(inlineValue);
    if (inlineNumericValue !== null && inlineNumericValue <= 5) {
      return {
        matched: true,
        value: String(inlineNumericValue),
        sourcePreview: makeSourcePreview(line),
        sourceType: 'ambient-average-line'
      };
    }

    const windowLines = reportData.lines.slice(lineIndex + 1, lineIndex + 6);
    for (const candidateLine of windowLines) {
      const trimmedCandidate = String(candidateLine || '').trim();
      if (!trimmedCandidate) {
        continue;
      }

      if (!/calculated\s+value|mos/i.test(trimmedCandidate)) {
        continue;
      }

      const numericValue = parseNumericValue(calculatedValuePattern.exec(trimmedCandidate)?.[1] || getLastMeaningfulNumericToken(trimmedCandidate));
      if (numericValue === null) {
        continue;
      }

      return {
        matched: true,
        value: String(numericValue),
        sourcePreview: makeSourcePreview(`${line} ${trimmedCandidate}`),
        sourceType: 'ambient-average-line'
      };
    }
  }

  return null;
}

function resolveAmbientNoiseMetricRowValue(reportData, item, textNormalizeConfig) {
  const tableConfig = item.tableConfig;
  if (!tableConfig?.rowNameMatch || !tableConfig?.targetColumnName) {
    return null;
  }

  const normalizedChecklistDesc = normalizeText(item.checklistDesc || '', textNormalizeConfig);
  if (!normalizedChecklistDesc.includes('ambient noise')) {
    return null;
  }

  const metricKey = getAmbientNoiseMetricKey(tableConfig.targetColumnName);
  if (!metricKey) {
    return null;
  }

  const sceneKeys = new Set(
    getAmbientNoiseSceneCandidates(tableConfig.rowNameMatch)
      .map((candidate) => getAmbientNoiseSceneKey(candidate))
      .filter(Boolean)
  );

  const targetRow = reportData.tableRows.find((rowContext) => {
    const identifier = String(rowContext?.cells?.[0] || '').trim();
    if (!identifier) {
      return false;
    }

    const match = identifier.match(/^([A-Z]+)_MOS_([A-Za-z0-9]+)_(?:HASB|HA(?:NB|WB|SB|SWB))$/i);
    if (!match) {
      return false;
    }

    const rowMetricKey = match[1].toLowerCase() === 's'
      ? 'smos'
      : match[1].toLowerCase() === 'n'
        ? 'nmos'
        : match[1].toLowerCase() === 'g'
          ? 'gmos'
          : null;

    if (rowMetricKey !== metricKey) {
      return false;
    }

    const sceneKey = getAmbientNoiseSceneKey(match[2]);
    return sceneKeys.has(sceneKey);
  });

  if (!targetRow) {
    return null;
  }

  const value = extractMetricRowValue(targetRow.cells);
  if (!value) {
    return null;
  }

  return {
    matched: true,
    value,
    sourcePreview: makeSourcePreview(targetRow.text),
    sourceType: 'ambient-metric-row'
  };
}

function resolveAmbientNoiseMetricLineValue(reportData, item, textNormalizeConfig) {
  const tableConfig = item.tableConfig;
  if (!tableConfig?.rowNameMatch || !tableConfig?.targetColumnName) {
    return null;
  }

  const normalizedChecklistDesc = normalizeText(item.checklistDesc || '', textNormalizeConfig);
  if (!normalizedChecklistDesc.includes('ambient noise')) {
    return null;
  }

  const metricKey = getAmbientNoiseMetricKey(tableConfig.targetColumnName);
  if (!metricKey) {
    return null;
  }

  const sceneKeys = new Set(
    getAmbientNoiseSceneCandidates(tableConfig.rowNameMatch)
      .map((candidate) => getAmbientNoiseSceneKey(candidate))
      .filter(Boolean)
  );

  for (let lineIndex = 0; lineIndex < reportData.lines.length; lineIndex += 1) {
    const line = String(reportData.lines[lineIndex] || '').trim();
    if (!line) {
      continue;
    }

    const identifierMatch = line.match(/^([A-Z]+)_MOS_([A-Za-z0-9]+)_(?:HASB|HA(?:NB|WB|SB|SWB))(.*)$/i);
    if (!identifierMatch) {
      continue;
    }

    const rowMetricKey = identifierMatch[1].toLowerCase() === 's'
      ? 'smos'
      : identifierMatch[1].toLowerCase() === 'n'
        ? 'nmos'
        : identifierMatch[1].toLowerCase() === 'g'
          ? 'gmos'
          : null;

    if (rowMetricKey !== metricKey) {
      continue;
    }

    const sceneKey = getAmbientNoiseSceneKey(identifierMatch[2]);
    if (!sceneKeys.has(sceneKey)) {
      continue;
    }

    const inlineValue = parseNumericValue(identifierMatch[3]);
    if (inlineValue !== null) {
      return {
        matched: true,
        value: String(inlineValue),
        sourcePreview: makeSourcePreview(line),
        sourceType: 'ambient-metric-line'
      };
    }

    const windowLines = reportData.lines.slice(lineIndex + 1, lineIndex + 6);
    for (const candidateLine of windowLines) {
      const trimmedCandidate = String(candidateLine || '').trim();
      if (!trimmedCandidate || /^(?:Measured|Calculated Value|G-MOS|N-MOS|S-MOS)$/i.test(trimmedCandidate)) {
        continue;
      }

      if (/\b\d{1,2}\/\d{1,2}\/\d{4}\b/.test(trimmedCandidate)) {
        continue;
      }

      const numericValue = parseNumericValue(trimmedCandidate);
      if (numericValue === null) {
        continue;
      }

      return {
        matched: true,
        value: String(numericValue),
        sourcePreview: makeSourcePreview(`${line} ${trimmedCandidate}`),
        sourceType: 'ambient-metric-line'
      };
    }
  }

  return null;
}

function resolveTableValue(reportData, item, textNormalizeConfig) {
  const tableConfig = item.tableConfig;
  if (!tableConfig) {
    return { matched: false, reason: '缺少 tableConfig 配置' };
  }

  const ambientNoiseValue = resolveAmbientNoiseAverageValue(reportData, item, textNormalizeConfig);
  if (ambientNoiseValue) {
    return ambientNoiseValue;
  }

  const ambientMetricRowValue = resolveAmbientNoiseMetricRowValue(reportData, item, textNormalizeConfig);
  if (ambientMetricRowValue) {
    return ambientMetricRowValue;
  }

  const ambientMetricLineValue = resolveAmbientNoiseMetricLineValue(reportData, item, textNormalizeConfig);
  if (ambientMetricLineValue) {
    return ambientMetricLineValue;
  }

  const requiredHeaders = Array.isArray(tableConfig.tableHeaderMatch) ? tableConfig.tableHeaderMatch : [];
  const rowNameCandidates = getRowNameCandidates(tableConfig.rowNameMatch);

  for (const table of reportData.tables) {
    const headerRow = table.rows[0] || [];
    const normalizedHeaders = headerRow.map((header) => normalizeText(header, textNormalizeConfig));
    const hasHeaders = requiredHeaders.every((header) => {
      const normalizedHeader = normalizeText(header, textNormalizeConfig);
      return normalizedHeaders.some((cell) => cell.includes(normalizedHeader));
    });

    if (!hasHeaders) {
      continue;
    }

    const targetColumnIndex = headerRow.findIndex((header) => {
      return normalizeText(header, textNormalizeConfig).includes(normalizeText(tableConfig.targetColumnName, textNormalizeConfig));
    });

    if (targetColumnIndex === -1) {
      continue;
    }

    const targetRow = table.rows.slice(1).find((row) => {
      return row.some((cell) => rowNameCandidates.some((candidate) => normalizeText(cell, textNormalizeConfig).includes(normalizeText(candidate, textNormalizeConfig))));
    });

    if (!targetRow || !targetRow[targetColumnIndex]) {
      continue;
    }

    return {
      matched: true,
      value: targetRow[targetColumnIndex],
      sourcePreview: makeSourcePreview(targetRow.join(' | ')),
      sourceType: 'table-cell'
    };
  }

  for (const rowContext of getRowsForMatching(reportData, true)) {
    const normalizedText = normalizeText(rowContext.text, textNormalizeConfig);

    if (!rowNameCandidates.some((candidate) => normalizedText.includes(normalizeText(candidate, textNormalizeConfig)))) {
      continue;
    }

    if (!normalizedText.includes(normalizeText(tableConfig.targetColumnName, textNormalizeConfig))) {
      continue;
    }

    const value = extractMetricRowValue(rowContext.cells) || getLastMeaningfulNumericToken(rowContext.text);
    if (!value) {
      continue;
    }

    return {
      matched: true,
      value,
      sourcePreview: makeSourcePreview(rowContext.text),
      sourceType: 'text-row-table-fallback'
    };
  }

  return { matched: false, reason: '未找到匹配表格或目标单元格' };
}

function selectCandidateRow(reportData, item, textNormalizeConfig, globalMatchConfig) {
  const forbiddenSuffixes = pickForbiddenSuffixes(item, globalMatchConfig);
  const collectCandidates = (rows) => rows
    .map((rowContext) => ({
      rowContext,
      keywordTier: getKeywordMatchTier(reportData, item, rowContext.text, textNormalizeConfig, rowContext.sourceKind !== 'html-table')
    }))
    .filter(({ rowContext, keywordTier }) => {
      if (keywordTier === 0) {
        return false;
      }

      const descriptorText = getRowDescriptorText(rowContext);
      if (!hasRequiredSuffixes(descriptorText, item.requiredSuffix, textNormalizeConfig)) {
        return false;
      }

      if (hasForbiddenSuffix(descriptorText, forbiddenSuffixes, textNormalizeConfig)) {
        return false;
      }

      if (!hasExactRowKeywords(rowContext.text, item.exactRowKeywords, textNormalizeConfig)) {
        return false;
      }

      return true;
    })
    .sort((left, right) => {
      const rightDescriptor = getRowDescriptorText(right.rowContext);
      const leftDescriptor = getRowDescriptorText(left.rowContext);
      const rightScore = right.keywordTier * 10000 + getCandidateScore(right.rowContext) + getVariantPenalty(item, rightDescriptor, textNormalizeConfig);
      const leftScore = left.keywordTier * 10000 + getCandidateScore(left.rowContext) + getVariantPenalty(item, leftDescriptor, textNormalizeConfig);
      return rightScore - leftScore;
    });

  const primaryCandidates = collectCandidates(getPrimaryRows(reportData));
  if (primaryCandidates.length > 0) {
    return primaryCandidates[0].rowContext;
  }

  const fallbackCandidates = collectCandidates(getFallbackRows(reportData));
  return fallbackCandidates.length > 0 ? fallbackCandidates[0].rowContext : null;
}

function resolveRowBasedValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId) {
  if (item.extractType === 'summary_table_match') {
    const ambientOverallAverageValue = resolveAmbientNoiseOverallAverageValue(reportData, item, textNormalizeConfig);
    if (ambientOverallAverageValue) {
      return ambientOverallAverageValue;
    }
  }

  let rowContext = selectCandidateRow(reportData, item, textNormalizeConfig, globalMatchConfig);

  if (item.extractType === 'status_judge') {
    const forbiddenSuffixes = pickForbiddenSuffixes(item, globalMatchConfig);
    const explicitStatusRows = getRowsForMatching(reportData, true).filter((candidate) => {
      if (!itemMatchesKeywords(reportData, item, candidate.text, textNormalizeConfig, candidate.sourceKind !== 'html-table')) {
        return false;
      }

      const descriptorText = getRowDescriptorText(candidate);
      if (!hasRequiredSuffixes(descriptorText, item.requiredSuffix, textNormalizeConfig)) {
        return false;
      }

      if (hasForbiddenSuffix(descriptorText, forbiddenSuffixes, textNormalizeConfig)) {
        return false;
      }

      if (!hasExactRowKeywords(candidate.text, item.exactRowKeywords, textNormalizeConfig)) {
        return false;
      }

      return Boolean(extractStatus(candidate.text));
    });

    const explicitStatusRow = explicitStatusRows.sort((left, right) => getCandidateScore(right) - getCandidateScore(left))[0];

    if (explicitStatusRow) {
      rowContext = explicitStatusRow;
    }
  }

  if (!rowContext) {
    return { matched: false, reason: '未找到匹配的表格行' };
  }

  if (item.extractType === 'summary_table_match') {
    const value = extractSummaryValue(rowContext, textNormalizeConfig);
    return value
      ? { matched: true, value, sourcePreview: makeSourcePreview(rowContext.text), sourceType: 'table-row' }
      : { matched: false, reason: '匹配到行但未提取到值' };
  }

  if (item.extractType === 'formula_calc') {
    const value = extractFormulaValue(reportData, rowContext, item, textNormalizeConfig, extractedResultsByItemId);
    return value
      ? { matched: true, value, sourcePreview: makeSourcePreview(rowContext.text), sourceType: 'table-row-formula' }
      : { matched: false, reason: '匹配到行但公式计算失败' };
  }

  if (item.extractType === 'status_judge') {
    const value = extractStatus(rowContext.text);
    return value
      ? { matched: true, value, sourcePreview: makeSourcePreview(rowContext.text), sourceType: 'table-row-status' }
      : { matched: false, reason: '匹配到行但未找到状态值' };
  }

  return { matched: false, reason: '未知的行匹配提取类型' };
}

function createSearchData(rawText, html, structuredData = {}) {
  const rawLines = rawText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const structuredLines = Array.isArray(structuredData?.lines)
    ? structuredData.lines.map((line) => String(line || '').trim()).filter(Boolean)
    : [];

  const lines = structuredLines.length > 0
    ? Array.from(new Set([...structuredLines, ...rawLines]))
    : rawLines;

  const flattenedText = rawText
    .replace(/\r/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const measurementObject = extractMeasurementObject(rawText);
  const lineRows = lines
    .map((line, index) => {
      const cells = line.split(/\t+/).map((cell) => cell.trim()).filter(Boolean);

      return {
        tableIndex: -1,
        rowIndex: index + 1,
        headers: [],
        cells,
        text: line,
        sourceKind: 'raw-line'
      };
    })
    .filter((row) => {
      if (row.text.length <= 8) {
        return false;
      }

      const hasExplicitStatus = Boolean(extractStatus(row.text)) || row.cells.some((cell) => isStatusCell(cell));
      const hasStructuredCells = row.cells.length >= 4;

      return hasExplicitStatus && hasStructuredCells;
    });

  let derivedRows = [];

  if (measurementObject) {
    const splitRows = flattenedText
      .split(measurementObject)
      .map((segment) => segment.trim())
      .filter((segment) => segment.length > 8 && /\b(?:ok|done|not ok)\b/i.test(segment) && /-?\d+(?:\.\d+)?\s*$/.test(segment))
      .map((text, index) => ({
        tableIndex: -1,
        rowIndex: index + 1,
        headers: [],
        cells: [],
        text,
        sourceKind: 'measurement-split'
      }));

    if (splitRows.length > 0) {
      derivedRows = splitRows;
    }
  }

  if (derivedRows.length === 0) {
    const rowStartMatches = Array.from(flattenedText.matchAll(/\b9\.\d+(?:\.\d+)*\b/g));
    derivedRows = rowStartMatches
      .map((match, index) => {
        const start = match.index;
        const end = index + 1 < rowStartMatches.length ? rowStartMatches[index + 1].index : flattenedText.length;
        const text = flattenedText.slice(start, end).trim();

        return {
          tableIndex: -1,
          rowIndex: index + 1,
          headers: [],
          cells: [],
          text,
          sourceKind: 'row-start'
        };
      })
      .filter((row) => row.text.length > 8);
  }

  const structuredTables = Array.isArray(structuredData?.tables)
    ? structuredData.tables
      .map((table, tableIndex) => ({
        tableIndex,
        rows: Array.isArray(table?.rows)
          ? table.rows.map((row) => Array.isArray(row) ? row.map((cell) => String(cell || '').trim()).filter(Boolean) : []).filter((row) => row.length > 0)
          : []
      }))
      .filter((table) => table.rows.length > 0)
    : [];

  const $ = cheerio.load(html);
  const htmlTables = $('table')
    .toArray()
    .map((tableElement, tableIndex) => {
      const rows = $(tableElement)
        .find('tr')
        .toArray()
        .map((rowElement) => {
          return $(rowElement)
            .find('th, td')
            .toArray()
            .map((cellElement) => $(cellElement).text().replace(/\s+/g, ' ').trim())
            .filter(Boolean);
        })
        .filter((row) => row.length > 0);

      return { tableIndex, rows };
    })
    .filter((table) => table.rows.length > 0);

  const tables = structuredTables.length > 0 ? structuredTables : htmlTables;

  const tableRows = tables.flatMap((table) => {
    const headers = table.rows[0] || [];
    return table.rows.slice(1).map((cells, rowOffset) => ({
      tableIndex: table.tableIndex,
      rowIndex: rowOffset + 1,
      headers,
      cells,
      text: cells.join(' | '),
      sourceKind: 'html-table'
    }));
  });

  return {
    rawText,
    html,
    lines,
    ambientNoiseBlocks: extractAmbientNoiseAverageBlocks(rawText),
    tables,
    tableRows,
    fallbackRows: [...lineRows, ...derivedRows]
  };
}

module.exports = {
  createSearchData,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
};
