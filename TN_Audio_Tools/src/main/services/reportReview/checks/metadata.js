const path = require('path');

const { normalizeText, normalizeUpperText } = require('../utils');

function checkReportBasicInfo(reportPath, wordData) {
  const issues = [];
  const evidence = [];

  const fileName = path.parse(reportPath).name.toUpperCase();
  const fileNameTokens = fileName.split(/[_\-\s]+/).filter(Boolean);

  evidence.push(`报告文件名：${fileName}`);

  const headers = wordData.headers || [];
  const footers = wordData.footers || [];

  evidence.push(`检测到页眉段落数：${headers.length}，页脚段落数：${footers.length}`);

  let measurementObject = '';

  wordData.paragraphs.forEach((para, index) => {
    const upperText = normalizeUpperText(para);

    if (
      upperText.includes('MEASUREMENT OBJECT') ||
      (upperText.includes('OBJECT') && (upperText.includes('MEAS') || upperText.includes('MEASUREMENT')))
    ) {
      const valuePart = para.split(/[:：]/)[1];
      if (valuePart) {
        measurementObject = normalizeText(valuePart);
        evidence.push(`第 ${index + 1} 行：识别到 Measurement Object = "${measurementObject}"`);
      }
    }

    if (!measurementObject && /^[^:：]*(?:object|对象)[^:：]*[:：]/i.test(para)) {
      const match = para.match(/[:：]\s*(.+)$/);
      if (match) {
        measurementObject = normalizeText(match[1]);
      }
    }
  });

  if (!measurementObject) {
    issues.push({
      severity: 'warning',
      message: '未在报告中找到明确的 Measurement Object 信息'
    });
    evidence.push('Measurement Object：未找到');
  } else {
    evidence.push('Measurement Object 与文件名匹配度分析中...');

    const moTokens = normalizeUpperText(measurementObject).split(/[_\-\s]+/).filter(Boolean);
    const commonTokens = moTokens.filter((t) => fileNameTokens.includes(t));

    if (commonTokens.length < Math.min(moTokens.length, fileNameTokens.length) * 0.5) {
      issues.push({
        severity: 'warning',
        message: `Measurement Object 与报告文件名的关键信息不够吻合：${measurementObject} vs ${fileName}`
      });
      evidence.push(`关键词重合度较低：${commonTokens.length}/${Math.max(moTokens.length, fileNameTokens.length)}`);
    }
  }

  if (issues.length === 0) {
    evidence.push('✓ 报告名称与 Measurement Object 信息基本一致');
    return { issues, evidence, status: 'pass' };
  }

  return { issues, evidence, status: 'warning' };
}

function checkTestItemConsistency(reportPath, wordData) {
  const issues = [];
  const evidence = [];

  const fileName = normalizeUpperText(path.parse(reportPath).name);
  evidence.push(`报告文件名：${fileName}`);

  const codecMatch = fileName.match(/\b(AMR|AMR-WB|EVS|OPUS|SILK|G729|G711)\b/i);
  const bandwidthMatch = fileName.match(/\b(NB|WB|SWB|FB)\b/);
  const modeMatch = fileName.match(/\b(HA|HE|HH|HANDSET|HEADSET|HANDSFREE)\b/i);

  const extractedCodec = codecMatch ? codecMatch[1].toUpperCase() : null;
  const extractedBandwidth = bandwidthMatch ? bandwidthMatch[1].toUpperCase() : null;
  const extractedMode = modeMatch ? normalizeUpperText(modeMatch[1]) : null;

  evidence.push(`文件名提取 - Codec: ${extractedCodec}, Bandwidth: ${extractedBandwidth}, Mode: ${extractedMode}`);

  const reportText = (wordData.paragraphs || []).join(' ').toUpperCase();

  const codecMentions = [];
  const bandwidthMentions = [];
  const modeMentions = [];

  if (extractedCodec && reportText.includes(extractedCodec)) {
    codecMentions.push(extractedCodec);
  }

  if (extractedBandwidth && reportText.includes(extractedBandwidth)) {
    bandwidthMentions.push(extractedBandwidth);
  }

  const modeAliases = {
    HANDSET: ['HA', 'HANDSET'],
    HEADSET: ['HE', 'HEADSET'],
    HANDSFREE: ['HH', 'HANDSFREE', 'HANDS-FREE']
  };

  if (extractedMode) {
    for (const aliases of Object.values(modeAliases)) {
      if (aliases.some((a) => normalizeUpperText(extractedMode).includes(a))) {
        if (aliases.some((a) => reportText.includes(a))) {
          modeMentions.push(extractedMode);
        }
      }
    }
  }

  evidence.push(
    `报告正文中找到 Codec 提及：${codecMentions.length > 0 ? '✓' : '✗'}, Bandwidth：${
      bandwidthMentions.length > 0 ? '✓' : '✗'
    }, Mode：${modeMentions.length > 0 ? '✓' : '✗'}`
  );

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

function findEngineerNames(wordData) {
  const evidence = [];
  const foundNames = [];

  const namePatterns = [
    /(?:tested by|tester|engineer|测试人员|工程师)[\s:：]+([a-zA-Z\u4e00-\u9fff\s]+)/i,
    /(?:prepared by|author|作者|编写者)[\s:：]+([a-zA-Z\u4e00-\u9fff\s]+)/i,
    /(?:reviewed by|reviewer|审核人|复核者)[\s:：]+([a-zA-Z\u4e00-\u9fff\s]+)/i,
    /(?:signed by|signature|签署人)[\s:：]+([a-zA-Z\u4e00-\u9fff\s]+)/i
  ];

  if (wordData.paragraphs) {
    wordData.paragraphs.forEach((para, index) => {
      const text = normalizeText(para);
      const upperText = normalizeUpperText(para);

      namePatterns.forEach((pattern) => {
        const match = text.match(pattern);
        if (match && match[1]) {
          const name = normalizeText(match[1]);
          foundNames.push({ name, context: upperText.slice(0, 80), lineIndex: index + 1 });
        }
      });
    });
  }

  if (foundNames.length === 0) {
    evidence.push('未在报告中找到明确的测试人员/工程师信息');
    return { engineers: [], evidence, status: 'review' };
  }

  const uniqueNames = Array.from(new Map(foundNames.map((n) => [normalizeUpperText(n.name), n])).values());

  uniqueNames.forEach((n) => {
    evidence.push(`第 ${n.lineIndex} 行：${n.name}`);
  });

  evidence.push(`✓ 识别到 ${uniqueNames.length} 个测试人员信息`);

  return { engineers: uniqueNames, evidence, status: 'pass' };
}

module.exports = {
  checkReportBasicInfo,
  checkTestItemConsistency,
  checkNamePollution,
  findEngineerNames
};