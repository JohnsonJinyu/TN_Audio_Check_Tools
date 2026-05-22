const { normalizeText, normalizeUpperText, buildWordData } = require('./utils');
const { normalizeVolumeLevel, extractVolumeLevelFromTitle } = require('./volumeLevelUtils');

// 列索引仅作为 row.raw 数组格式的回退，优先使用对象属性名
var DETAILED_COL = {
  NAME: 3,
  SMD: 7,
  DIRECTION: 19,
  VOLUME_CTRL: 22,
  VALUE: 1,
  UNIT: 2,
  BGN_SCENARIO: 18,
  DATE: 10,
  TIME: 11,
};

/**
 * 从行数据读取字段值 — 优先用对象属性名，回退到 raw[] 数组索引
 */
function _field(row, propName, colIndex) {
  if (!row) return '';
  // 对象属性优先
  if (row[propName] !== undefined && row[propName] !== null) return row[propName];
  // 回退: raw 数组
  var raw = Array.isArray(row) ? row : (row.raw || []);
  if (raw[colIndex] !== undefined && raw[colIndex] !== null) return raw[colIndex];
  return '';
}

function classifyTestCategory(name, smd, bgnScenario) {
  // 仅用 name 做分类匹配（smd为SMD标识符，可能含与测试类型无关的通用关键字）
  const nameText = normalizeUpperText(String(name || ''));
  const fullText = normalizeUpperText([name, smd].filter(Boolean).join(' '));

  if (fullText.includes('BGN') && fullText.includes('CONNECT')) return 'bgn_connection';
  if (bgnScenario && /\b(?:PUB|ROAD|CROSSROAD|TRAIN|STATION|CAR|CAFE|STREET|OFFICE|HOTEL)\b/i.test(String(bgnScenario))) {
    return '3quest';
  }
  if (fullText.includes('SIDETONE') && fullText.includes('DELAY')) return 'sidetone_delay';
  if (fullText.includes('SIDETONE') && !fullText.includes('DELAY')) return 'sidetone';
  // 仅在 name 中匹配时延关键词，排除校准/补偿/锁定等无关项
  var isDelayName = (
    /\b(?:ROUND\s*TRIP|ONE\s*WAY|GROUP)\s*DELAY\b/i.test(nameText) ||
    /\bDELAY\s*TIME\b/i.test(nameText) ||
    /\bDELAY\b/i.test(nameText) ||
    /时延/.test(nameText)
  );
  if (fullText.includes('ECHO') && fullText.includes('DELAY')) return 'echo_delay';
  if (isDelayName && !/CALIBRAT|COMPENSAT|OFFSET|LOCK/i.test(nameText)) return 'delay';
  if (/LOUDNESS|RLR|SLR|STMR/.test(fullText)) return 'loudness';
  if (/FREQUENCY\s*RESPONSE|频响|FREQ\b|SENSITIVITY[\s,]*FREQUENCY/i.test(fullText)) return 'frequency_response';
  if (/MOS-LQO|POLQA|P\.863/.test(fullText)) return 'polqa';
  return 'other';
}

function parseDateTime(dateVal, timeVal) {
  if (!dateVal && !timeVal) return null;

  try {
    if (dateVal instanceof Date) {
      if (timeVal instanceof Date) {
        return new Date(
          dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate(),
          timeVal.getHours(), timeVal.getMinutes(), timeVal.getSeconds()
        );
      }
      return dateVal;
    }

    if (typeof dateVal === 'number' && dateVal > 40000) {
      const excelEpoch = new Date(1899, 11, 30);
      const msPerDay = 86400000;
      return new Date(excelEpoch.getTime() + dateVal * msPerDay);
    }

    if (typeof dateVal === 'string' && typeof timeVal === 'string') {
      const combined = `${dateVal} ${timeVal}`;
      const parsed = new Date(combined);
      if (!isNaN(parsed.getTime())) return parsed;
    }
  } catch (_) {
    return null;
  }

  return null;
}

function extractSectionNumber(text) {
  var normalized = normalizeText(text);
  if (!normalized) return '';
  var match = normalized.match(/^(\d+(?:\.\d+){1,4}[a-z]?)\b/i);
  return match ? match[1] : '';
}

function cleanTimingTitle(text) {
  return normalizeText(text)
    .replace(/\bPAGEREF\b\s+_Toc\d+\s+\d+$/i, '')
    .replace(/\b_?Toc\d+\b/gi, '')
    .replace(/\s+\d+$/g, '')
    .trim();
}

function normalizeTimingTitleKey(text) {
  return cleanTimingTitle(text)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractWordTimingTitles(reportData) {
  var wordData = buildWordData(reportData);
  if (!wordData || (!Array.isArray(wordData.paragraphs) && !Array.isArray(wordData.tables))) return [];

  var headingPatterns = [
    { regex: /^(\d{1,2}(?:\.\d{1,2}){0,4}[a-z]?)\s+(.{3,})$/i },
    { regex: /^(?:table|表|figure|图)\s*(\d+(?:\.\d+)*)\s*[:.-]?\s+(.{3,})$/i }
  ];

  var allLines = [
    ...(wordData.paragraphs || []).map(function(text, index) {
      return { text: normalizeText(text), index: index, source: 'paragraph' };
    }),
    ...((wordData.tables || []).flatMap(function(table, tableIndex) {
      return (table.rows || []).map(function(row, rowIndex) {
        return {
          text: row.map(function(cell) { return normalizeText(cell); }).filter(Boolean).join(' | '),
          index: rowIndex,
          source: 'table-' + tableIndex
        };
      });
    }))
  ].filter(function(line) { return line.text; });

  var titles = [];
  allLines.forEach(function(line) {
    var text = normalizeText(line.text);
    if (!text || text.length > 240) return;
    headingPatterns.forEach(function(pattern) {
      var match = text.match(pattern.regex);
      if (!match) return;
      var title = cleanTimingTitle(match[2]);
      if (!title) return;
      titles.push({
        chapterNumber: String(match[1] || '').trim(),
        title: title,
        titleKey: normalizeTimingTitleKey(title),
      });
    });
  });

  return Array.from(new Map(
    titles
      .filter(function(item) { return item.chapterNumber && item.titleKey; })
      .map(function(item) { return [item.chapterNumber + '|' + item.titleKey, item]; })
  ).values());
}

function scoreTimingTitleMatch(itemText, titleKey) {
  if (!itemText || !titleKey) return 0;
  if (itemText.includes(titleKey) || titleKey.includes(itemText)) return 999;
  var tokens = titleKey.split(' ').filter(function(token) { return token.length >= 3; });
  var score = 0;
  tokens.forEach(function(token) {
    if (itemText.includes(token)) score += 1;
  });
  return score;
}

function buildGroupedTimingItems(reportData, rawItems) {
  var evidence = [];
  if (!Array.isArray(rawItems) || rawItems.length === 0) {
    return { items: [], evidence: evidence };
  }

  var titleEntries = extractWordTimingTitles(reportData);
  if (titleEntries.length === 0) {
    return { items: rawItems, evidence: evidence };
  }

  var groupedMap = new Map();
  rawItems.forEach(function(item) {
    var sectionNumber = extractSectionNumber(item.descriptor)
      || extractSectionNumber(item.smd)
      || extractSectionNumber(item.name);
    if (!sectionNumber) return;

    var candidates = titleEntries.filter(function(entry) { return entry.chapterNumber === sectionNumber; });
    if (candidates.length === 0) return;

    var itemText = normalizeTimingTitleKey([item.descriptor, item.name, item.smd].filter(Boolean).join(' '));
    var best = candidates[0];
    var bestScore = scoreTimingTitleMatch(itemText, candidates[0].titleKey);
    for (var i = 1; i < candidates.length; i++) {
      var score = scoreTimingTitleMatch(itemText, candidates[i].titleKey);
      if (score > bestScore) {
        best = candidates[i];
        bestScore = score;
      }
    }

    if (bestScore <= 0) return;

    var key = sectionNumber + '|' + best.titleKey;
    if (!groupedMap.has(key)) {
      groupedMap.set(key, {
        chapterNumber: sectionNumber,
        descriptor: sectionNumber + ' ' + best.title,
        name: best.title,
        smd: sectionNumber,
        direction: item.direction || null,
        timestamp: item.timestamp || null,
        startTimestamp: item.timestamp || null,
        endTimestamp: item.timestamp || null,
        bgnScenario: item.bgnScenario || '',
        testCategory: item.testCategory || 'other',
        rowIndex: item.rowIndex,
        sourceRowCount: 0,
      });
    }

    var current = groupedMap.get(key);
    current.sourceRowCount += 1;

    if (!current.direction && item.direction) current.direction = item.direction;
    if (!current.bgnScenario && item.bgnScenario) current.bgnScenario = item.bgnScenario;

    if (!current.timestamp || (item.timestamp && item.timestamp < current.timestamp)) {
      current.timestamp = item.timestamp;
    }
    if (!current.startTimestamp || (item.timestamp && item.timestamp < current.startTimestamp)) {
      current.startTimestamp = item.timestamp;
    }
    if (!current.endTimestamp || (item.timestamp && item.timestamp > current.endTimestamp)) {
      current.endTimestamp = item.timestamp;
    }
    if (item.rowIndex < current.rowIndex) current.rowIndex = item.rowIndex;

    if (current.testCategory === 'other' && item.testCategory && item.testCategory !== 'other') {
      current.testCategory = item.testCategory;
    }
    if ((current.testCategory !== 'delay' && current.testCategory !== 'echo_delay')
      && (item.testCategory === 'delay' || item.testCategory === 'echo_delay')) {
      current.testCategory = item.testCategory;
    }
  });

  var groupedItems = Array.from(groupedMap.values())
    .map(function(item) {
      item.testCategory = classifyTestCategory(item.name || item.descriptor, item.descriptor || item.smd, item.bgnScenario);
      return item;
    })
    .filter(function(item) { return !!item.timestamp; })
    .sort(function(a, b) {
      if (a.timestamp && b.timestamp) return a.timestamp - b.timestamp;
      return a.rowIndex - b.rowIndex;
    });

  if (groupedItems.length === 0) {
    return { items: rawItems, evidence: evidence };
  }

  evidence.push('Detailed 行级时间戳 ' + rawItems.length + ' 条，按Word标题聚合为 ' + groupedItems.length + ' 个测试项');
  return { items: groupedItems, evidence: evidence };
}

/**
 * 从 ACQUA xlsx 的 Detailed sheet 提取测试项时间戳
 */
function extractTimestamps(reportData) {
  const evidence = [];
  const rawTestItemTimestamps = [];
  let hasAbsoluteTimestamps = false;

  const detailedRows = reportData?.detailedRows || reportData?.detailedRowContexts || [];
  if (!detailedRows.length) {
    evidence.push('未找到ACQUA Detailed数据，无法提取时间戳');
    return { testItemTimestamps: [], hasAbsoluteTimestamps: false, evidence };
  }

  detailedRows.forEach((row, rowIndex) => {
    var name = normalizeText(_field(row, 'Name', DETAILED_COL.NAME));
    var smd = normalizeText(_field(row, 'SMD', DETAILED_COL.SMD));
    var bgnScenario = _field(row, 'BGNScenario', DETAILED_COL.BGN_SCENARIO);
    var direction = normalizeText(_field(row, 'Direction', DETAILED_COL.DIRECTION));
    var dateVal = _field(row, 'Date', DETAILED_COL.DATE) || null;
    var timeVal = _field(row, 'Time', DETAILED_COL.TIME) || null;

    var timestamp = parseDateTime(dateVal, timeVal);
    if (timestamp) hasAbsoluteTimestamps = true;

    rawTestItemTimestamps.push({
      rowIndex,
      descriptor: [smd, name].filter(Boolean).join(' - ') || `Row ${rowIndex + 1}`,
      name,
      smd,
      direction: direction || null,
      timestamp,
      bgnScenario: String(bgnScenario || '').trim(),
      testCategory: classifyTestCategory(name, smd, bgnScenario),
    });
  });

  if (hasAbsoluteTimestamps) {
    const validTimestamps = rawTestItemTimestamps.filter((t) => t.timestamp);
    if (validTimestamps.length > 0) {
      const sorted = [...validTimestamps].sort((a, b) => a.timestamp - b.timestamp);
      evidence.push(`Detailed 行级共 ${rawTestItemTimestamps.length} 条，${validTimestamps.length} 条有时间戳`);
      evidence.push(`时间范围: ${sorted[0].timestamp.toISOString()} 至 ${sorted[sorted.length - 1].timestamp.toISOString()}`);
    }
  } else {
    evidence.push('ACQUA Detailed数据中未找到Date/Time时间戳，时序检查将以行序为准');
  }

  var grouped = buildGroupedTimingItems(reportData, rawTestItemTimestamps);
  evidence.push.apply(evidence, grouped.evidence || []);

  return {
    testItemTimestamps: grouped.items,
    rawTestItemTimestamps,
    hasAbsoluteTimestamps,
    evidence
  };
}

/**
 * 从 ACQUA xlsx 提取响度相关数值
 */
function extractLoudnessMetrics(reportData) {
  const evidence = [];
  const metrics = [];

  const detailedRows = reportData?.detailedRows || reportData?.detailedRowContexts || [];
  if (!detailedRows.length) {
    evidence.push('未找到ACQUA Detailed数据');
    return { loudnessMetrics: [], evidence };
  }

  const reportContext = reportData?.reportContext || {};
  const codec = reportContext.codec || '';
  const network = reportContext.network || '';
  const bandwidth = reportContext.bandwidth || '';

  detailedRows.forEach((row, rowIndex) => {
    var name = normalizeText(_field(row, 'Name', DETAILED_COL.NAME));
    var smd = normalizeText(_field(row, 'SMD', DETAILED_COL.SMD));
    var unit = normalizeText(_field(row, 'Unit', DETAILED_COL.UNIT));
    var direction = normalizeText(_field(row, 'Direction', DETAILED_COL.DIRECTION));
    var volumeCtrl = normalizeText(_field(row, 'VolumeCTRL', DETAILED_COL.VOLUME_CTRL));
    var rawValue = _field(row, 'Value', DETAILED_COL.VALUE) || null;

    var category = classifyTestCategory(name, smd, '');
    if (category !== 'loudness') return;

    var numericValue = Number(rawValue);
    if (isNaN(numericValue)) return;

    metrics.push({
      descriptor: [smd, name].filter(Boolean).join(' - '),
      direction: direction || null,
      network,
      codec,
      bandwidth,
      value: numericValue,
      unit: unit || 'dB',
      volumeLevel: normalizeVolumeLevel(volumeCtrl)
        || extractVolumeLevelFromTitle(name)
        || extractVolumeLevelFromTitle(smd)
        || extractVolumeLevelFromTitle([smd, name].filter(Boolean).join(' ')),
      category: name.toUpperCase().includes('RLR') ? 'RLR'
        : name.toUpperCase().includes('SLR') ? 'SLR'
        : name.toUpperCase().includes('STMR') ? 'STMR'
        : 'loudness',
      rowIndex,
    });
  });

  var missingLevel = metrics.filter(function(m) { return !m.volumeLevel; }).length;
  if (missingLevel > 0) evidence.push('其中 ' + missingLevel + ' 个响度测点未能提取到音量等级');
  evidence.push(`提取到 ${metrics.length} 个响度相关数值`);

  return { loudnessMetrics: metrics, evidence };
}

/**
 * 从 ACQUA xlsx 提取频响相关数据
 */
function extractFrequencyResponseData(reportData) {
  const evidence = [];
  const metrics = [];

  const detailedRows = reportData?.detailedRows || reportData?.detailedRowContexts || [];
  if (!detailedRows.length) {
    evidence.push('未找到ACQUA Detailed数据');
    return { frequencyResponseMetrics: [], evidence };
  }

  detailedRows.forEach((row, rowIndex) => {
    var name = normalizeText(_field(row, 'Name', DETAILED_COL.NAME));
    var smd = normalizeText(_field(row, 'SMD', DETAILED_COL.SMD));
    var unit = normalizeText(_field(row, 'Unit', DETAILED_COL.UNIT));
    var direction = normalizeText(_field(row, 'Direction', DETAILED_COL.DIRECTION));
    var volumeCtrl = normalizeText(_field(row, 'VolumeCTRL', DETAILED_COL.VOLUME_CTRL));
    var rawValue = _field(row, 'Value', DETAILED_COL.VALUE) || null;

    var category = classifyTestCategory(name, smd, '');
    if (category !== 'frequency_response') return;

    var numericValue = Number(rawValue);
    if (isNaN(numericValue)) return;

    metrics.push({
      descriptor: [smd, name].filter(Boolean).join(' - '),
      direction: direction || null,
      amplitude: numericValue,
      unit: unit || 'dB',
      volumeLevel: normalizeVolumeLevel(volumeCtrl)
        || extractVolumeLevelFromTitle(name)
        || extractVolumeLevelFromTitle(smd)
        || extractVolumeLevelFromTitle([smd, name].filter(Boolean).join(' ')),
      frequencyBin: null,
      rowIndex,
    });
  });

  var missingFRLevel = metrics.filter(function(m) { return !m.volumeLevel; }).length;
  if (missingFRLevel > 0) evidence.push('其中 ' + missingFRLevel + ' 个频响测点未能提取到音量等级');
  evidence.push(`提取到 ${metrics.length} 个频响相关数值`);

  return { frequencyResponseMetrics: metrics, evidence };
}

/**
 * 主入口：构建测试数据事实
 * 对Word报告返回null（优雅降级）
 */
function buildTestDataFacts(reportData) {
  if (!reportData) return null;

  const format = (reportData.reportFormat || '').toLowerCase();
  if (format !== 'xlsx' && format !== 'xls' && format !== 'paired') return null;

  const timestamps = extractTimestamps(reportData);
  const loudness = extractLoudnessMetrics(reportData);
  const frequencyResponse = extractFrequencyResponseData(reportData);

  return {
    testItemTimestamps: timestamps.testItemTimestamps,
    rawTestItemTimestamps: timestamps.rawTestItemTimestamps,
    hasAbsoluteTimestamps: timestamps.hasAbsoluteTimestamps,
    loudnessMetrics: loudness.loudnessMetrics,
    frequencyResponseMetrics: frequencyResponse.frequencyResponseMetrics,
    timingEvidence: timestamps.evidence,
    loudnessEvidence: loudness.evidence,
    frequencyResponseEvidence: frequencyResponse.evidence,
    evidence: [...timestamps.evidence, ...loudness.evidence, ...frequencyResponse.evidence],
    sourceFormat: format,
  };
}

module.exports = {
  DETAILED_COL,
  classifyTestCategory,
  parseDateTime,
  extractTimestamps,
  extractLoudnessMetrics,
  extractFrequencyResponseData,
  buildTestDataFacts,
};
