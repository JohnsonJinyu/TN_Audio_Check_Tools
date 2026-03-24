const path = require('path');

const { normalizeText, normalizeUpperText } = require('../utils');
const {
  MODE_ALIASES,
  CODEC_ALIASES,
  BANDWIDTH_ALIASES,
  collectMentionEvidence,
  buildReviewFacts,
  tokenizeMeaningfulText,
  normalizeTerminalMode
} = require('../reportFacts');

function computeOverlapRatio(leftValue, rightValue) {
  const leftTokens = tokenizeMeaningfulText(leftValue);
  const rightTokens = tokenizeMeaningfulText(rightValue);

  if (leftTokens.length === 0 || rightTokens.length === 0) {
    return 0;
  }

  const commonTokens = leftTokens.filter((token) => rightTokens.includes(token));
  return commonTokens.length / Math.max(leftTokens.length, rightTokens.length);
}

function checkReportBasicInfo(reportPath, wordData, reviewFacts = buildReviewFacts(reportPath, wordData)) {
  const issues = [];
  const evidence = [];

  const { metadata } = reviewFacts;
  const fileName = metadata.fileNameUpper;

  evidence.push(`报告文件名：${fileName}`);

  const headers = wordData.headers || [];
  const footers = wordData.footers || [];

  evidence.push(`检测到页眉段落数：${headers.length}，页脚段落数：${footers.length}`);

  const measurementObject = metadata.measurementObject.value;

  if (!measurementObject) {
    issues.push({
      severity: 'warning',
      message: '未在报告中找到明确的 Measurement Object 信息'
    });
    evidence.push('Measurement Object：未找到');
  } else {
    evidence.push(`Measurement Object：${measurementObject}`);
    metadata.measurementObject.evidences.slice(0, 3).forEach((item) => evidence.push(`识别证据：${item}`));

    const objectFileOverlap = computeOverlapRatio(measurementObject, metadata.fileName);
    if (objectFileOverlap < 0.35) {
      issues.push({
        severity: 'warning',
        message: `Measurement Object 与报告文件名的关键信息不够吻合：${measurementObject} vs ${fileName}`
      });
      evidence.push(`Measurement Object 与文件名重合度偏低：${objectFileOverlap.toFixed(2)}`);
    }
  }

  const headerSource = metadata.headerText || metadata.footerText;
  if (!headerSource) {
    issues.push({
      severity: 'review',
      message: '未提取到页眉或页脚文本，无法自动完成命名一致性校验'
    });
  } else {
    evidence.push(`页眉页脚摘要：${headerSource.slice(0, 120)}`);
    const headerOverlap = computeOverlapRatio(metadata.fileName, headerSource);
    if (headerOverlap < 0.25) {
      issues.push({
        severity: 'warning',
        message: '报告文件名与页眉页脚文本重合度偏低，建议人工确认命名一致性'
      });
      evidence.push(`文件名与页眉页脚重合度：${headerOverlap.toFixed(2)}`);
    }
  }

  if (issues.length === 0) {
    evidence.push('✓ 报告名称与 Measurement Object 信息基本一致');
    return { issues, evidence, status: 'pass' };
  }

  return { issues, evidence, status: issues.some((item) => item.severity === 'review') ? 'review' : 'warning' };
}

function checkTestItemConsistency(reportPath, wordData, reviewFacts = buildReviewFacts(reportPath, wordData)) {
  const issues = [];
  const evidence = [];

  const { metadata, lines } = reviewFacts;

  const extractedCodec = metadata.codec || null;
  const extractedBandwidth = metadata.bandwidth || null;
  const extractedMode = metadata.terminalMode || null;

  evidence.push(`文件名提取 - Codec: ${extractedCodec}, Bandwidth: ${extractedBandwidth}, Mode: ${extractedMode}`);

  const codecMentions = extractedCodec ? collectMentionEvidence(lines, CODEC_ALIASES[extractedCodec] || [extractedCodec]) : [];
  const bandwidthMentions = extractedBandwidth ? collectMentionEvidence(lines, BANDWIDTH_ALIASES[extractedBandwidth] || [extractedBandwidth]) : [];
  const modeMentions = extractedMode ? collectMentionEvidence(lines, MODE_ALIASES[normalizeTerminalMode(extractedMode)] || [extractedMode]) : [];

  evidence.push(
    `报告正文中找到 Codec 提及：${codecMentions.length > 0 ? '✓' : '✗'}, Bandwidth：${
      bandwidthMentions.length > 0 ? '✓' : '✗'
    }, Mode：${modeMentions.length > 0 ? '✓' : '✗'}`
  );

  codecMentions.slice(0, 2).forEach((item) => evidence.push(`Codec 证据：${item}`));
  bandwidthMentions.slice(0, 2).forEach((item) => evidence.push(`Bandwidth 证据：${item}`));
  modeMentions.slice(0, 2).forEach((item) => evidence.push(`Mode 证据：${item}`));

  if (codecMentions.length === 0 && extractedCodec) {
    issues.push({
      severity: 'warning',
      message: `文件名表明使用 ${extractedCodec}，但报告正文中未明确提及该编码方式`
    });
  }

  if (bandwidthMentions.length === 0 && extractedBandwidth) {
    issues.push({
      severity: 'warning',
      message: `文件名表明使用 ${extractedBandwidth}，但报告正文中未明确提及该带宽`
    });
  }

  if (modeMentions.length === 0 && extractedMode) {
    issues.push({
      severity: 'warning',
      message: `文件名表明是 ${extractedMode} 模式，但报告正文中未明确说明`
    });
  }

  if (issues.length === 0) {
    evidence.push('✓ 测试项信息与文件名标注一致');
    return { issues, evidence, status: 'pass' };
  }

  return { issues, evidence, status: 'warning' };
}

function checkNamePollution(reportPath) {
  const issues = [];
  const evidence = [];

  const fileName = path.parse(reportPath).name;
  evidence.push(`报告文件名：${fileName}`);

  const pollutionPatterns = [
    { regex: /\d{8}/, label: '日期格式 (YYYYMMDD)', severity: 'warning' },
    { regex: /\d{4}-\d{2}-\d{2}/, label: '日期格式 (YYYY-MM-DD)', severity: 'warning' },
    { regex: /v\d+\.\d+/i, label: '版本号 (v1.0 等)', severity: 'warning' },
    { regex: /DVT|PVT|EVT|ALPHA|BETA/i, label: '测试版本标识', severity: 'error' },
    { regex: /draft|草稿/i, label: '草稿标记', severity: 'warning' },
    { regex: /\(|_\(|\[\(/i, label: '非标准特殊标记', severity: 'review' }
  ];

  pollutionPatterns.forEach((pattern) => {
    if (pattern.regex.test(fileName)) {
      issues.push({
        severity: pattern.severity,
        message: `报告名称包含 ${pattern.label}：${fileName.match(pattern.regex)[0]}`,
        pollutionType: pattern.label
      });
      evidence.push(`检测到污染信息：${pattern.label}`);
    }
  });

  if (issues.length === 0) {
    evidence.push('✓ 报告名称清洁，无污染信息');
    return { issues, evidence, status: 'pass' };
  }

  return { issues, evidence, status: issues.some((i) => i.severity === 'error') ? 'error' : 'warning' };
}

function findEngineerNames(wordData, reviewFacts = buildReviewFacts('', wordData)) {
  const evidence = [];
  const foundNames = reviewFacts.engineers || [];

  if (foundNames.length === 0) {
    evidence.push('未在报告中找到明确的测试人员/工程师信息');
    return { engineers: [], evidence, status: 'review' };
  }

  foundNames.forEach((item) => {
    evidence.push(`${item.source}: ${item.name}`);
    if (item.evidence) {
      evidence.push(`证据：${item.evidence}`);
    }
  });

  evidence.push(`✓ 识别到 ${foundNames.length} 个测试人员信息`);

  return { engineers: foundNames, evidence, status: 'pass' };
}

module.exports = {
  checkReportBasicInfo,
  checkTestItemConsistency,
  checkNamePollution,
  findEngineerNames
};