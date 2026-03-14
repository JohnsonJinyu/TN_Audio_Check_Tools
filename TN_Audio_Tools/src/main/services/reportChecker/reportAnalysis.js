const cheerio = require('cheerio');

function normalizeText(value, textNormalizeConfig = {}) {
  if (typeof value !== 'string') {
    return '';
  }

  let normalized = value.replace(/\r/g, ' ').replace(/\n/g, ' ');

  normalized = normalized
    .replace(/\bha(?:nb|wb|sb|swb)\b/gi, ' HABAND ')
    .replace(/\bmxx-(\d+)/gi, ' MAX-$1 ')
    .replace(/\breceiv(?:e|ing|eing)\b/gi, ' RCV ')
    .replace(/\bsending\b/gi, ' SND ')
    .replace(/\bfrequency\b/gi, ' FREQ ')
    .replace(/\bcharact(?:er)?\.?\b/gi, ' CHAR ')
    .replace(/\baverage\b/gi, ' AVG ')
    .replace(/\bnominal\b/gi, ' NOM ')
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

function itemMatchesKeywords(reportData, item, text, textNormalizeConfig, allowFallback = false) {
  if (hasAllKeywords(text, item.coreKeywords, textNormalizeConfig)) {
    return true;
  }

  const isLegacyTextOnlyReport = !Array.isArray(reportData?.tables) || reportData.tables.length === 0;
  if (!isLegacyTextOnlyReport && !allowFallback) {
    return false;
  }

  return hasAllKeywords(text, item.fallbackKeywords, textNormalizeConfig);
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

  return text.match(/-?\d+(?:\.\d+)?(?:\s*(?:%|dB(?:m0|Pa)?|ms|Hz|MOS|s))?/gi) || [];
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
    crossroad: ['Crossroad', 'Crossroads', 'Outside Traffic Crossroads'],
    tstation: ['Train Station', 'TStation', 'Train'],
    fullsizecar: ['Fullsize Car 130 km/h', 'Fullsize Car', 'FullsizeCar', 'Car'],
    cafeteria: ['Cafeteria', 'Cafeteria Noise'],
    mensa: ['Mensa'],
    callcenter: ['Callcenter', 'Call Center', 'Work Noise Office Callcenter']
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

  if (normalized.includes('crossroad')) {
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

  if (normalized.includes('cafeteria')) {
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

  if (/^[A-Za-z0-9]+(?:_[A-Za-z0-9+-]+){2,}$/.test(normalized)) {
    return true;
  }

  return /\.(?:doc|docx|xls|xlsx)$/i.test(normalized);
}

function getLastMeaningfulNumericToken(text) {
  const cleanedText = String(text || '')
    .replace(/^[0-9.]+\s+/, '')
    .replace(/[A-Za-z0-9]+(?:_[A-Za-z0-9+-]+)+$/g, '')
    .trim();

  const tokens = extractNumberTokens(cleanedText);
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

  const match = text.match(/-?\d+(?:\.\d+)?/);
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

  for (let index = 0; index < rowContext.headers.length; index += 1) {
    const normalizedHeader = normalizeText(rowContext.headers[index] || '', textNormalizeConfig);
    if (!normalizedHeader) {
      continue;
    }

    if (normalizedPatterns.some((pattern) => normalizedHeader.includes(pattern))) {
      return rowContext.cells[index] || null;
    }
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

  const numericCellValue = extractRowNumericCell(rowContext);
  if (numericCellValue) {
    return numericCellValue;
  }

  const directCellValue = getLastMeaningfulNumericToken(directCell);
  if (directCellValue) {
    return directCellValue;
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

  const shouldUseBaseItemValue = rowContext?.sourceKind && rowContext.sourceKind !== 'html-table';

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

  for (const rowContext of reportData.tableRows) {
    if (!itemMatchesKeywords(reportData, item, rowContext.text, textNormalizeConfig, rowContext.sourceKind !== 'html-table')) {
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
      value: rowContext.text.includes('%') && numericValue !== null ? String(numericValue / 100) : value,
      sourcePreview: makeSourcePreview(rowContext.text),
      sourceType: 'regex-row',
      score: 2000 + getCandidateScore(rowContext)
    });
  }

  for (let lineIndex = 0; lineIndex < reportData.lines.length; lineIndex += 1) {
    const contextWindow = findLineWindow(reportData.lines, Math.max(0, lineIndex - 2), 12);
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
        value: candidateLine.includes('%') && numericValue !== null ? String(numericValue / 100) : value,
        sourcePreview: makeSourcePreview(candidateLine),
        sourceType: 'regex-line',
        score: (lineHasRequiredSuffix ? 1000 : 0) + sameLineKeywordScore + proximityScore + candidateLine.length
      });
    }
  }

  if (candidates.length > 0) {
    return candidates.sort((left, right) => right.score - left.score)[0];
  }

  return { matched: false, reason: '未匹配到正则目标文本' };
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

function resolveTableValue(reportData, item, textNormalizeConfig) {
  const tableConfig = item.tableConfig;
  if (!tableConfig) {
    return { matched: false, reason: '缺少 tableConfig 配置' };
  }

  const ambientNoiseValue = resolveAmbientNoiseAverageValue(reportData, item, textNormalizeConfig);
  if (ambientNoiseValue) {
    return ambientNoiseValue;
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

  for (const rowContext of reportData.tableRows) {
    const normalizedText = normalizeText(rowContext.text, textNormalizeConfig);

    if (!rowNameCandidates.some((candidate) => normalizedText.includes(normalizeText(candidate, textNormalizeConfig)))) {
      continue;
    }

    if (!normalizedText.includes(normalizeText(tableConfig.targetColumnName, textNormalizeConfig))) {
      continue;
    }

    const tokens = extractNumberTokens(rowContext.text);
    if (tokens.length === 0) {
      continue;
    }

    return {
      matched: true,
      value: tokens[tokens.length - 1].trim(),
      sourcePreview: makeSourcePreview(rowContext.text),
      sourceType: 'text-row-table-fallback'
    };
  }

  return { matched: false, reason: '未找到匹配表格或目标单元格' };
}

function selectCandidateRow(reportData, item, textNormalizeConfig, globalMatchConfig) {
  const forbiddenSuffixes = pickForbiddenSuffixes(item, globalMatchConfig);
  const candidates = reportData.tableRows.filter((rowContext) => {
    if (!itemMatchesKeywords(reportData, item, rowContext.text, textNormalizeConfig, rowContext.sourceKind !== 'html-table')) {
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
  });

  if (candidates.length === 0) {
    return null;
  }

  return candidates.sort((left, right) => getCandidateScore(right) - getCandidateScore(left))[0];
}

function resolveRowBasedValue(reportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId) {
  let rowContext = selectCandidateRow(reportData, item, textNormalizeConfig, globalMatchConfig);

  if (item.extractType === 'status_judge') {
    const forbiddenSuffixes = pickForbiddenSuffixes(item, globalMatchConfig);
    const explicitStatusRows = reportData.tableRows.filter((candidate) => {
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

function createSearchData(rawText, html) {
  const lines = rawText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

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

  const $ = cheerio.load(html);
  const tables = $('table')
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
    tableRows: [...tableRows, ...lineRows, ...derivedRows]
  };
}

module.exports = {
  createSearchData,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue
};
