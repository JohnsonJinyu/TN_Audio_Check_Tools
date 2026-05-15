const { normalizeText, normalizeUpperText } = require('./utils');

const DETAILED_COL = {
  MEASUREMENT_OBJECT: 0,
  VALUE: 1,
  UNIT: 2,
  NAME: 3,
  COMMENT: 4,
  STATUS: 5,
  SMD_NO: 6,
  SMD: 7,
  INDEX: 8,
  RUN: 9,
  DATE: 10,
  TIME: 11,
  LOWER_LIMIT: 12,
  UPPER_LIMIT: 13,
  DESCRIPTION: 14,
  BANDWIDTH: 17,
  BGN_SCENARIO: 18,
  DIRECTION: 19,
  USE_CASE: 20,
  VAR_NAME: 21,
  VOLUME_CTRL: 22,
};

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
    /时延/.test(nameText)
  );
  if (isDelayName && !/CALIBRAT|COMPENSAT|OFFSET|LOCK/i.test(nameText)) return 'delay';
  if (fullText.includes('ECHO') && fullText.includes('DELAY')) return 'echo_delay';
  if (/LOUDNESS|RLR|SLR|STMR/.test(fullText)) return 'loudness';
  if (/FREQUENCY\s*RESPONSE|频响|FREQ\b/.test(fullText)) return 'frequency_response';
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

/**
 * 从 ACQUA xlsx 的 Detailed sheet 提取测试项时间戳
 */
function extractTimestamps(reportData) {
  const evidence = [];
  const testItemTimestamps = [];
  let hasAbsoluteTimestamps = false;

  const detailedRows = reportData?.detailedRows || reportData?.detailedRowContexts || [];
  if (!detailedRows.length) {
    evidence.push('未找到ACQUA Detailed数据，无法提取时间戳');
    return { testItemTimestamps: [], hasAbsoluteTimestamps: false, evidence };
  }

  detailedRows.forEach((row, rowIndex) => {
    const rawRow = Array.isArray(row) ? row : row.raw || [];
    const name = normalizeText(rawRow[DETAILED_COL.NAME] || row.Name || '');
    const smd = normalizeText(rawRow[DETAILED_COL.SMD] || row.SMD || '');
    const bgnScenario = rawRow[DETAILED_COL.BGN_SCENARIO] || row.BGNScenario || '';
    const direction = normalizeText(rawRow[DETAILED_COL.DIRECTION] || row.Direction || '');
    const dateVal = rawRow[DETAILED_COL.DATE] || row.Date || null;
    const timeVal = rawRow[DETAILED_COL.TIME] || row.Time || null;

    const timestamp = parseDateTime(dateVal, timeVal);
    if (timestamp) hasAbsoluteTimestamps = true;

    testItemTimestamps.push({
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
    const validTimestamps = testItemTimestamps.filter((t) => t.timestamp);
    if (validTimestamps.length > 0) {
      const sorted = [...validTimestamps].sort((a, b) => a.timestamp - b.timestamp);
      evidence.push(`共 ${testItemTimestamps.length} 个测试项，${validTimestamps.length} 个有时间戳`);
      evidence.push(`时间范围: ${sorted[0].timestamp.toISOString()} 至 ${sorted[sorted.length - 1].timestamp.toISOString()}`);
    }
  } else {
    evidence.push('ACQUA Detailed数据中未找到Date/Time时间戳，时序检查将以行序为准');
  }

  return { testItemTimestamps, hasAbsoluteTimestamps, evidence };
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
    const rawRow = Array.isArray(row) ? row : row.raw || [];
    const name = normalizeText(rawRow[DETAILED_COL.NAME] || row.Name || '');
    const smd = normalizeText(rawRow[DETAILED_COL.SMD] || row.SMD || '');
    const unit = normalizeText(rawRow[DETAILED_COL.UNIT] || row.Unit || '');
    const direction = normalizeText(rawRow[DETAILED_COL.DIRECTION] || row.Direction || '');
    const volumeCtrl = normalizeText(rawRow[DETAILED_COL.VOLUME_CTRL] || row.VolumeCTRL || '');
    const rawValue = rawRow[DETAILED_COL.VALUE] || row.Value || null;

    const category = classifyTestCategory(name, smd, '');
    if (category !== 'loudness') return;

    const numericValue = Number(rawValue);
    if (isNaN(numericValue)) return;

    metrics.push({
      descriptor: [smd, name].filter(Boolean).join(' - '),
      direction: direction || null,
      network,
      codec,
      bandwidth,
      value: numericValue,
      unit: unit || 'dB',
      volumeLevel: volumeCtrl || null,
      category: name.toUpperCase().includes('RLR') ? 'RLR'
        : name.toUpperCase().includes('SLR') ? 'SLR'
        : name.toUpperCase().includes('STMR') ? 'STMR'
        : 'loudness',
      rowIndex,
    });
  });

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
    const rawRow = Array.isArray(row) ? row : row.raw || [];
    const name = normalizeText(rawRow[DETAILED_COL.NAME] || row.Name || '');
    const smd = normalizeText(rawRow[DETAILED_COL.SMD] || row.SMD || '');
    const unit = normalizeText(rawRow[DETAILED_COL.UNIT] || row.Unit || '');
    const direction = normalizeText(rawRow[DETAILED_COL.DIRECTION] || row.Direction || '');
    const volumeCtrl = normalizeText(rawRow[DETAILED_COL.VOLUME_CTRL] || row.VolumeCTRL || '');
    const rawValue = rawRow[DETAILED_COL.VALUE] || row.Value || null;

    const category = classifyTestCategory(name, smd, '');
    if (category !== 'frequency_response') return;

    const numericValue = Number(rawValue);
    if (isNaN(numericValue)) return;

    metrics.push({
      descriptor: [smd, name].filter(Boolean).join(' - '),
      direction: direction || null,
      amplitude: numericValue,
      unit: unit || 'dB',
      volumeLevel: volumeCtrl || null,
      frequencyBin: null,
      rowIndex,
    });
  });

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
