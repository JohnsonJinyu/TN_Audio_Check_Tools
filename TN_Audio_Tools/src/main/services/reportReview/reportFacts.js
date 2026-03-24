const path = require('path');

const { normalizeText, normalizeUpperText } = require('./utils');

const MODE_ALIASES = {
  HA: ['HA', 'HANDSET'],
  HE: ['HE', 'HS', 'HEADSET'],
  HH: ['HH', 'HF', 'HANDSFREE', 'HANDS-FREE', 'HANDS FREE']
};

const BANDWIDTH_ALIASES = {
  NB: ['NB'],
  WB: ['WB'],
  SWB: ['SWB', 'SB'],
  FB: ['FB']
};

const CODEC_ALIASES = {
  AMR: ['AMR'],
  'AMR-WB': ['AMR-WB', 'AMR_WB', 'AMR WB'],
  EVS: ['EVS'],
  OPUS: ['OPUS'],
  SILK: ['SILK'],
  G729: ['G729', 'G.729'],
  G711: ['G711', 'G.711']
};

function tokenizeMeaningfulText(value) {
  return normalizeUpperText(value)
    .split(/[^A-Z0-9]+/)
    .map((token) => token.trim())
    .filter(Boolean)
    .filter((token) => token.length > 1)
    .filter((token) => !/^20\d{6}$/.test(token))
    .filter((token) => !/^V?\d+(?:\.\d+)+$/.test(token));
}

function normalizeTerminalMode(value) {
  const upperValue = normalizeUpperText(value);

  if (!upperValue) {
    return '';
  }

  for (const [mode, aliases] of Object.entries(MODE_ALIASES)) {
    if (aliases.some((alias) => upperValue.includes(alias))) {
      return mode;
    }
  }

  return '';
}

function normalizeBandwidth(value) {
  const upperValue = normalizeUpperText(value);

  if (!upperValue) {
    return '';
  }

  for (const [bandwidth, aliases] of Object.entries(BANDWIDTH_ALIASES)) {
    if (aliases.some((alias) => upperValue.includes(alias))) {
      return bandwidth;
    }
  }

  return '';
}

function normalizeCodec(value) {
  const upperValue = normalizeUpperText(value);

  if (!upperValue) {
    return '';
  }

  for (const [codec, aliases] of Object.entries(CODEC_ALIASES)) {
    if (aliases.some((alias) => upperValue.includes(alias))) {
      return codec;
    }
  }

  return '';
}

function pushUniqueEvidence(target, value) {
  const normalized = normalizeText(value);
  if (normalized && !target.includes(normalized)) {
    target.push(normalized);
  }
}

function collectAllLines(wordData) {
  const lines = [];

  (wordData.headers || []).forEach((header, index) => {
    const text = normalizeText(header);
    if (text) {
      lines.push({ source: 'header', index, text });
    }
  });

  (wordData.paragraphs || []).forEach((paragraph, index) => {
    const text = normalizeText(paragraph);
    if (text) {
      lines.push({ source: 'paragraph', index, text });
    }
  });

  (wordData.tables || []).forEach((table, tableIndex) => {
    (table.rows || []).forEach((row, rowIndex) => {
      const cells = Array.isArray(row) ? row.map((cell) => normalizeText(cell)).filter(Boolean) : [];
      if (cells.length > 0) {
        lines.push({
          source: 'table',
          index: rowIndex,
          tableIndex,
          cells,
          text: cells.join(' | ')
        });
      }
    });
  });

  (wordData.footers || []).forEach((footer, index) => {
    const text = normalizeText(footer);
    if (text) {
      lines.push({ source: 'footer', index, text });
    }
  });

  return lines;
}

function deriveMetadataFromFileName(fileName) {
  const normalizedFileName = normalizeUpperText(fileName);
  const codec = normalizeCodec(normalizedFileName);
  const bandwidth = normalizeBandwidth(normalizedFileName);
  const terminalMode = normalizeTerminalMode(normalizedFileName);

  return {
    codec,
    bandwidth,
    terminalMode
  };
}

function extractMeasurementObject(lines, reportContext = {}) {
  const evidences = [];
  const candidates = [];
  const contextValue = normalizeText(reportContext.measurementObject);

  if (contextValue) {
    candidates.push({ value: contextValue, source: 'reportContext' });
    evidences.push(`reportContext: ${contextValue}`);
  }

  const directPatterns = [
    /measurement\s*object\s*[:：-]\s*(.+)$/i,
    /(?:^|\|)\s*measurement\s*object\s*\|\s*(.+)$/i,
    /(?:^|\b)object\s*[:：-]\s*(.+)$/i,
    /对象\s*[:：-]\s*(.+)$/
  ];

  const rowLabelPatterns = [
    /^measurement\s*object$/i,
    /^object$/i,
    /^测试对象$/,
    /^对象$/
  ];

  lines.forEach((line) => {
    directPatterns.forEach((pattern) => {
      const match = line.text.match(pattern);
      if (match?.[1]) {
        const value = normalizeText(match[1].split('|')[0]);
        if (value) {
          candidates.push({ value, source: line.source });
          evidences.push(`${line.source}: ${value}`);
        }
      }
    });

    if (line.source === 'table' && Array.isArray(line.cells) && line.cells.length >= 2) {
      const label = normalizeText(line.cells[0]);
      if (rowLabelPatterns.some((pattern) => pattern.test(label))) {
        const value = normalizeText(line.cells.slice(1).join(' '));
        if (value) {
          candidates.push({ value, source: 'table' });
          evidences.push(`table: ${value}`);
        }
      }
    }
  });

  const uniqueCandidates = Array.from(new Map(
    candidates.map((candidate) => [normalizeUpperText(candidate.value), candidate])
  ).values());

  return {
    value: uniqueCandidates[0]?.value || '',
    candidates: uniqueCandidates,
    evidences
  };
}

function collectMentionEvidence(lines, aliases) {
  const evidences = [];

  lines.forEach((line) => {
    if (aliases.some((alias) => normalizeUpperText(line.text).includes(alias))) {
      pushUniqueEvidence(evidences, line.text);
    }
  });

  return evidences;
}

function buildExpectedMetadata(reportPath, wordData, lines) {
  const fileName = path.parse(reportPath).name;
  const fileNameUpper = normalizeUpperText(fileName);
  const fileTokens = tokenizeMeaningfulText(fileName);
  const reportContext = wordData.reportContext || {};
  const measurementObject = extractMeasurementObject(lines, reportContext);

  const fileNameDerived = deriveMetadataFromFileName(fileName);

  const codec = normalizeCodec(reportContext.codec) || fileNameDerived.codec;
  const bandwidth = normalizeBandwidth(reportContext.bandwidth) || fileNameDerived.bandwidth;
  const terminalMode = normalizeTerminalMode(reportContext.terminalMode) || fileNameDerived.terminalMode;

  return {
    fileName,
    fileNameUpper,
    fileTokens,
    reportContext,
    measurementObject,
    codec,
    bandwidth,
    terminalMode,
    headerText: (wordData.headers || []).map((item) => normalizeText(item)).filter(Boolean).join(' | '),
    footerText: (wordData.footers || []).map((item) => normalizeText(item)).filter(Boolean).join(' | ')
  };
}

function extractEngineerFacts(lines) {
  const candidates = [];
  const patterns = [
    /(?:tested by|tester|engineer|prepared by|author|reviewed by|reviewer|approved by|signed by|测试人员|工程师|编写者|作者|审核人|复核者|批准人|签署人)\s*[:：]\s*([A-Za-z][A-Za-z .'-]{1,60}|[\u4e00-\u9fff]{2,8})/i,
    /(?:tested by|tester|engineer|prepared by|author|reviewed by|reviewer|approved by|signed by|测试人员|工程师|编写者|作者|审核人|复核者|批准人|签署人)\s+([A-Za-z][A-Za-z .'-]{1,60}|[\u4e00-\u9fff]{2,8})/i
  ];

  const rowLabelPatterns = [
    /^(?:tested by|tester|engineer|prepared by|author|reviewed by|reviewer|approved by|signed by)$/i,
    /^(?:测试人员|工程师|编写者|作者|审核人|复核者|批准人|签署人)$/
  ];

  lines.forEach((line) => {
    patterns.forEach((pattern) => {
      const match = line.text.match(pattern);
      if (match?.[1]) {
        candidates.push({ name: normalizeText(match[1]), source: line.source, evidence: line.text });
      }
    });

    if (line.source === 'table' && Array.isArray(line.cells) && line.cells.length >= 2) {
      const label = normalizeText(line.cells[0]);
      if (rowLabelPatterns.some((pattern) => pattern.test(label))) {
        const name = normalizeText(line.cells[1]);
        if (name) {
          candidates.push({ name, source: 'table', evidence: line.text });
        }
      }
    }
  });

  return Array.from(new Map(
    candidates
      .filter((candidate) => candidate.name)
      .map((candidate) => [normalizeUpperText(candidate.name), candidate])
  ).values());
}

function extractPolqaFacts(lines) {
  const evidences = {
    algorithm: [],
    version: [],
    reference: []
  };

  const algorithmPattern = /(POLQA|MOS-LQO|P\.863|MOS-LQO\s+P\.863)/i;
  const versionPattern = /(POLQA|MOS-LQO|P\.863).{0,100}?(?:algorithm\s+version|version|ver\.?|版本)\s*[:：]?\s*([A-Za-z0-9._-]+)/i;
  const referencePattern = /(reference\s+(?:signal|speech|source|file)|ref\.?\s*(?:signal|speech|source|file)|参考(?:音源|语音|信号|文件))\s*[:：]?\s*([^|,.;。，；]{2,80})/i;

  lines.forEach((line) => {
    if (algorithmPattern.test(line.text)) {
      pushUniqueEvidence(evidences.algorithm, line.text);
    }

    const versionMatch = line.text.match(versionPattern);
    if (versionMatch) {
      pushUniqueEvidence(evidences.version, `${versionMatch[1]} ${versionMatch[2]}`);
    }

    const referenceMatch = line.text.match(referencePattern);
    if (referenceMatch) {
      pushUniqueEvidence(evidences.reference, `${referenceMatch[1]} ${normalizeText(referenceMatch[2])}`);
    }
  });

  return evidences;
}

function buildReviewFacts(reportPath, wordData) {
  const lines = collectAllLines(wordData);
  const metadata = buildExpectedMetadata(reportPath, wordData, lines);

  return {
    lines,
    metadata,
    engineers: extractEngineerFacts(lines),
    polqa: extractPolqaFacts(lines)
  };
}

module.exports = {
  MODE_ALIASES,
  BANDWIDTH_ALIASES,
  CODEC_ALIASES,
  tokenizeMeaningfulText,
  normalizeTerminalMode,
  normalizeBandwidth,
  normalizeCodec,
  collectAllLines,
  collectMentionEvidence,
  buildReviewFacts
};