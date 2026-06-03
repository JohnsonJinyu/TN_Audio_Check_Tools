const axios = require('axios');
const { PNG } = require('pngjs');
const { CANONICAL_LEVELS, compareVolumeLevel, normalizeVolumeLevel } = require('./volumeLevelUtils');

var PROMPT_TEMPLATE = [
  '你是一位专业的音频测试工程师。请分析以下ACQUA音频测试报告中的频率响应/响度曲线图，并与响度数值交叉验证。',
  '',
  '## 图片说明',
  '- 图片是X-Y折线图（频率-幅度），X轴是频率(Hz)，Y轴是幅度(dB)',
  '- Sending Dir(SND/TX) = 发送方向，Receiving Dir(RCV/RX) = 接收方向',
  '- 响度曲线图标题格式: "Loudness Rating RCV (RLR)" 或 "Loudness Rating SND (SLR)" + 测试等级',
  '- 频响曲线图标题格式: "Sensitivity, frequency RCV" 或 "Sensitivity, frequency SND" + 测试等级',
  '- 测试等级体系（8级，音量从大到小）: MAX > MAX-1 > MAX-2 > MAX-3(NOM) > MAX-4 > MAX-5 > MAX-6 > MAX-7(MIN)',
  '- 其中 NOM = MAX-3(NOM), MIN = MAX-7(MIN)；频响通常只在 MAX、NOM、MIN 三个等级测试',
  '- 对 Loudness Rating 而言，RLR / SLR 数值越小通常表示实际越响，对应曲线整体应更高',
  '',
  '## 核心检查任务',
  '',
  '### 任务1：同一方向-同一等级曲线趋势一致性',
  '对于同一方向（SND或RCV）的同一测试等级（如皆为MAX或皆为NOM），',
  '响度曲线图（RLR/SLR）和频响曲线图（Sensitivity, frequency）的曲线趋势应当一致或相似。',
  '比较方法：观察同一等级下曲线峰谷位置、整体形状、上升/下降趋势是否吻合。',
  '不一致（峰谷位置相反、趋势走向不同）→ 判定为异常。',
  '',
  '### 任务2：跨等级响度曲线单调性',
  '同一方向（SND或RCV）下，从MAX到MIN的响度曲线幅度应随等级变化而单调变化。',
  'MAX等级曲线幅度应最高，MIN等级应最低，中间等级应依次递变。',
  '如果出现波峰波谷（如MAX-2曲线反超MAX-1或低于MAX-4）→ 判定为异常。',
  '',
  '### 任务3：数值与曲线互相印证',
  '下方提供了从xlsx提取的响度数值（已按等级排序）。请验证：',
  '- 数值随等级变化是否单调（无波峰波谷）',
  '- 数值排序与对应曲线幅度排序是否一致',
  '- RCV方向：RLR值越小（如-9dB）→音量越大→曲线应越高；RLR值越大（如+8dB）→音量越小→曲线应越低',
  '',
  '## 响度数值（来自xlsx测试数据，已按等级排序）',
  '%LOUDNESS_DATA%',
  '',
  '## 频响数值（来自xlsx测试数据）',
  '%FREQ_RESPONSE_DATA%',
  '',
  '注意：如果频响数值显示"无频响数据"，频响信息仅来自图片，此时只使用图片进行趋势分析。',
  '',
  '## 输出格式',
  '请严格按以下JSON格式输出，不要输出任何其他内容：',
  '{',
  '  "findings": [',
  '    {',
  '      "direction": "SND或RCV",',
  '      "volumeLevel": "MAX或NOM或MIN或ALL",',
  '      "imageContext": "图片上下文摘要",',
  '      "trendConsistent": true,',
  '      "levelMonotonic": true,',
  '      "detail": "具体的曲线趋势观察和与数值的交叉验证结论（中文）",',
  '      "severity": "pass"',
  '    }',
  '  ],',
  '  "overallAssessment": "综合评估：(1)同等级FR-响度趋势是否一致 (2)跨等级单调性 (3)数值印证 (4)整体判定（中文，150字以内）",',
  '  "overallSeverity": "pass"',
  '}',
  '',
  'severity取值：pass(趋势一致且数值印证)、warning(有轻微差异但可接受)、error(趋势明显相反或严重异常)'
].join('\n');

function buildLoudnessSummary(testDataFacts) {
  if (!testDataFacts || !testDataFacts.loudnessMetrics || testDataFacts.loudnessMetrics.length === 0) {
    return '无响度数据';
  }
  var metrics = testDataFacts.loudnessMetrics;
  var byDir = {};
  metrics.forEach(function(m) {
    var dir = m.direction || 'unknown';
    if (!byDir[dir]) byDir[dir] = [];
    byDir[dir].push({ volume: m.volumeLevel, value: m.value, category: m.category });
  });
  var lines = [];
  Object.keys(byDir).forEach(function(dir) {
    // 按等级排序
    byDir[dir].sort(function(a, b) { return compareVolumeLevel(a.volume, b.volume); });
    lines.push(dir + '方向: ' + byDir[dir].length + '个测点');
    byDir[dir].slice(0, 12).forEach(function(m) {
      lines.push('  等级' + (m.volume || '?') + ': ' + (m.value != null ? Number(m.value).toFixed(2) : '?') + 'dB' + (m.category ? ' (' + m.category + ')' : ''));
    });
    if (byDir[dir].length > 12) lines.push('  ... 共' + byDir[dir].length + '个测点');
  });
  return lines.join('\n');
}

function buildFRSummary(testDataFacts) {
  if (!testDataFacts || !testDataFacts.frequencyResponseMetrics || testDataFacts.frequencyResponseMetrics.length === 0) {
    return '无频响数据';
  }
  var metrics = testDataFacts.frequencyResponseMetrics;
  var byDir = {};
  metrics.forEach(function(m) {
    var dir = m.direction || 'unknown';
    if (!byDir[dir]) byDir[dir] = [];
    byDir[dir].push({ volume: m.volumeLevel, amplitude: m.amplitude });
  });
  var lines = [];
  Object.keys(byDir).forEach(function(dir) {
    // 按等级排序
    byDir[dir].sort(function(a, b) { return compareVolumeLevel(a.volume, b.volume); });
    lines.push(dir + '方向: ' + byDir[dir].length + '个测点');
    byDir[dir].slice(0, 12).forEach(function(m) {
      lines.push('  等级' + (m.volume || '?') + ': 幅度' + (m.amplitude != null ? Number(m.amplitude).toFixed(2) + 'dB' : '?'));
    });
    if (byDir[dir].length > 12) lines.push('  ... 共' + byDir[dir].length + '个测点');
  });
  return lines.join('\n');
}

function parseLlmResponse(text) {
  var findings = [];
  var overallAssessment = '';
  var overallSeverity = 'review';
  var monotonicityViolations = [];

  try {
    var jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      var parsed = JSON.parse(jsonMatch[0]);
      if (Array.isArray(parsed.findings)) {
        findings = parsed.findings.map(function(f) {
          return {
            severity: f.severity === 'error' ? 'error' : (f.severity === 'warning' ? 'warning' : 'pass'),
            detail: f.detail || '',
            direction: f.direction || 'unknown',
            volumeLevel: normalizeVolumeLevel(f.volumeLevel || '') || f.volumeLevel || '',
            role: f.role || '',
            imageIndex: f.imageIndex,
            trendConsistent: f.trendConsistent,
            levelMonotonic: f.levelMonotonic,
            frequencyRange: f.frequencyRange || '',
            expectedBehavior: f.expectedBehavior || '',
            actualBehavior: f.actualBehavior || '',
          };
        });
      }
      overallAssessment = parsed.overallAssessment || '';
      if (Array.isArray(parsed.monotonicityViolations)) {
        monotonicityViolations = parsed.monotonicityViolations;
      }
      if (parsed.overallSeverity === 'error') overallSeverity = 'error';
      else if (parsed.overallSeverity === 'warning') overallSeverity = 'warning';
      else if (parsed.overallSeverity === 'pass') overallSeverity = 'pass';
    }
  } catch (_) {
    if (/不一致|相反|异常|错误|error/i.test(text)) overallSeverity = 'warning';
    overallAssessment = text.slice(0, 200);
  }

  return { findings, overallAssessment, overallSeverity, monotonicityViolations };
}

function isAliasOnlyLevelMismatchFinding(finding) {
  if (!finding || (finding.severity !== 'error' && finding.severity !== 'warning')) return false;

  var combined = [finding.detail, finding.expectedBehavior, finding.actualBehavior, finding.volumeLevel]
    .filter(Boolean)
    .join(' ')
    .toUpperCase();

  var looksLikeMismatch = /等级不匹配|无法进行同等级对比|不匹配/.test(combined);
  if (!looksLikeMismatch) return false;

  var nomAliasOnly = /\bNOM\b/.test(combined) && /MAX\s*[-–]?\s*3\s*\(\s*NOM\s*\)/.test(combined);
  var minAliasOnly = /\bMIN\b/.test(combined) && /MAX\s*[-–]?\s*7\s*\(\s*MIN\s*\)/.test(combined);
  return nomAliasOnly || minAliasOnly;
}

function cropPng(png, x, y, width, height) {
  var out = new PNG({ width: width, height: height });
  PNG.bitblt(png, out, x, y, width, height, 0, 0);
  return out;
}

function haveIdenticalPlotBody(leftImage, rightImage) {
  if (!leftImage || !rightImage || !leftImage.base64 || !rightImage.base64) return false;

  try {
    var left = PNG.sync.read(Buffer.from(leftImage.base64, 'base64'));
    var right = PNG.sync.read(Buffer.from(rightImage.base64, 'base64'));
    if (!left || !right || left.width !== right.width || left.height !== right.height) return false;

    // 只比对图表主体，排除标题与上下文文字；当前 ACQUA 导出图尺寸固定。
    var leftPlot = cropPng(left, 60, 35, 255, 205);
    var rightPlot = cropPng(right, 60, 35, 255, 205);

    for (var i = 0; i < leftPlot.data.length; i++) {
      if (leftPlot.data[i] !== rightPlot.data[i]) return false;
    }
    return true;
  } catch (_) {
    return false;
  }
}

function buildDuplicatePlotFinding(task, referenceImage, loudnessImage) {
  return {
    direction: task.direction,
    volumeLevel: task.level || '',
    role: loudnessImage.category || '',
    imageIndex: loudnessImage.index,
    severity: 'error',
    trendConsistent: false,
    levelMonotonic: true,
    frequencyRange: '',
    expectedBehavior: '响度图应为独立测试曲线，不应与同等级FR基准图的图表主体完全相同',
    actualBehavior: 'FR基准图与响度图的图表主体像素完全一致，仅标题文字不同',
    detail: '检测到同等级响度图与FR基准图疑似复用同一底图，请优先核查报告原图是否放错',
    duplicatePlot: true,
    referenceImageIndex: referenceImage.index,
  };
}

function taskMatchesSwbRule(task, testDataFacts) {
  var metrics = (testDataFacts && testDataFacts.loudnessMetrics) || [];
  for (var i = 0; i < metrics.length; i++) {
    if (String(metrics[i].bandwidth || '').toUpperCase() === 'SWB') return true;
  }

  var combined = ((task && task.label) || '') + ' ' + ((task && task.direction) || '');
  (task && task.images || []).forEach(function(img) {
    combined += ' ' + String(img && img.contextText || '');
    combined += ' ' + String(img && img.fileName || '');
  });

  return /SUPERWIDEBAND|\bSWB\b|\b(?:HA|HE|HH|HS)SB\b/i.test(combined);
}

function isActionableMonotonicityViolation(violation) {
  if (!violation) return false;
  var description = String(violation.description || '').trim();
  var combined = [description, violation.levelA, violation.levelB, violation.direction].filter(Boolean).join(' ');

  if (!combined) return false;

  // 这类返回并不是“检测到异常”，而是 LLM 表示样本不足或无法判定，
  // 不应在前端以 warning 形式展示成异常发现。
  if (/(无法验证|无法判断|无法确认|样本不足|数据不足|缺少按等级分组|缺少.*测试数据|未找到足够|仅有\s*1\s*个|只有\s*1\s*个)/i.test(combined)) {
    return false;
  }

  return true;
}

async function analyzeChartImages(params) {
  var images = params.images || [];
  var testDataFacts = params.testDataFacts || {};
  var llmSettings = params.settings || {};

  var apiUrl = String(llmSettings.apiUrl || '').replace(/\/+$/, '');
  var apiKey = String(llmSettings.apiKey || '');
  var model = String(llmSettings.model || 'claude-sonnet-4-20250514');

  if (!apiUrl || !apiKey) {
    return {
      issues: [{ severity: 'review', message: '未配置API地址或API Key，请在设置中填写' }],
      evidence: ['LLM图表分析未配置'],
      status: 'review'
    };
  }

  var prompt = PROMPT_TEMPLATE
    .replace('%LOUDNESS_DATA%', buildLoudnessSummary(testDataFacts))
    .replace('%FREQ_RESPONSE_DATA%', buildFRSummary(testDataFacts));

  // 构建消息内容（OpenAI/Anthropic 兼容的多模态格式）
  var content = [{ type: 'text', text: prompt }];
  images.forEach(function(img) {
    content.push({
      type: 'image_url',
      image_url: { url: 'data:' + img.contentType + ';base64,' + img.base64 }
    });
  });

  var evidence = [];

  try {
    var response = await axios.post(apiUrl + '/v1/chat/completions', {
      model: model,
      max_tokens: 2048,
      temperature: 0.1,
      messages: [{ role: 'user', content: content }],
      response_format: { type: 'json_object' }
    }, {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + apiKey
      },
      timeout: 180000
    });

    var responseText = '';
    var choice = (response.data && response.data.choices && response.data.choices[0]);
    if (choice && choice.message && choice.message.content) {
      responseText = choice.message.content;
    } else {
      responseText = JSON.stringify(response.data);
    }

    var parsed = parseLlmResponse(responseText);
    if (parsed.overallAssessment) {
      evidence.push('AI综合评估: ' + parsed.overallAssessment);
    }

    var passCount = 0;
    var warnCount = 0;
    var errCount = 0;
    parsed.findings.forEach(function(f) {
      if (f.severity === 'error') errCount++;
      else if (f.severity === 'warning') warnCount++;
      else passCount++;
    });
    evidence.push('分析结果: ' + passCount + ' 方向通过' + (warnCount > 0 ? ', ' + warnCount + ' 项警告' : '') + (errCount > 0 ? ', ' + errCount + ' 项异常' : ''));

    var issues = [];
    if (parsed.overallSeverity === 'warning') {
      issues.push({ severity: 'warning', message: 'AI图表分析: ' + (parsed.overallAssessment || '检测到轻微趋势差异') });
    } else if (parsed.overallSeverity === 'error') {
      issues.push({ severity: 'error', message: 'AI图表分析: ' + (parsed.overallAssessment || '检测到明显趋势不一致') });
    }

    return { issues: issues, evidence: evidence, status: parsed.overallSeverity === 'error' ? 'error' : (parsed.overallSeverity === 'warning' ? 'warning' : 'pass'), rawFindings: parsed.findings };
  } catch (error) {
    var errMsg = '';
    if (error.code === 'ECONNREFUSED' || error.code === 'ENOTFOUND') {
      errMsg = '无法连接API地址，请检查网络';
    } else if (error.response && error.response.status === 401) {
      errMsg = 'API Key无效或已过期(401)';
    } else if (error.response && error.response.status === 403) {
      errMsg = 'API Key无权限(403)';
    } else if (error.code === 'ETIMEDOUT' || error.code === 'ECONNABORTED') {
      errMsg = 'API响应超时，请重试';
    } else if (error.response && error.response.status === 429) {
      errMsg = '请求频率超限(429)，请稍后';
    } else {
      errMsg = (error.message || '未知错误');
    }
    return {
      issues: [{ severity: 'review', message: 'AI图表分析失败: ' + errMsg }],
      evidence: evidence,
      status: 'review'
    };
  }
}

var SINGLE_IMAGE_PROMPT = [
  '你是一位专业的音频测试工程师。请分析下面这张ACQUA音频测试报告中的曲线图。',
  '',
  '## 图片信息',
  '图片上下文: %CONTEXT%',
  '测试等级: %VOLUME_LEVEL%',
  '',
  '## 响度数值（来自xlsx，用于交叉验证）',
  '%LOUDNESS_DATA%',
  '',
  '## 分析要求',
  '1. 判断这张图的方向（SND发送/TX 或 RCV接收/RX）',
  '2. 确认音量等级（MAX/MAX-1/.../MAX-7(MIN)），与上下文中的测试等级一致',
  '3. 描述曲线的整体形态趋势（随频率升高是上升/下降/平坦）',
  '4. 判断这条曲线是否正常、有无异常特征',
  '5. 如果上下文中有dB值，判断该dB值对应的曲线幅度是否合理',
  '',
  '## 输出格式（严格JSON）',
  '{',
  '  "direction": "SND或RCV",',
  '  "volumeLevel": "MAX或NOM或MIN或unknown",',
  '  "trendDescription": "曲线趋势描述（中文，50字以内）",',
  '  "isNormal": true,',
  '  "detail": "详细观察（中文）",',
  '  "severity": "pass"',
  '}',
  'severity: pass(正常), warning(有轻微异常), error(明显异常)'
].join('\n');

var SUMMARY_PROMPT = [
  '你是一位专业的音频测试工程师。以下是逐张分析ACQUA报告的汇总结果。请综合判断。',
  '',
  '## 各图分析结果',
  '%FINDINGS%',
  '',
  '## 响度数值',
  '%LOUDNESS_DATA%',
  '',
  '## 综合检查要求',
  '1. 同一方向-同一等级下，响度曲线(RLR/SLR)与频响曲线(Sensitivity, frequency)趋势是否一致？',
  '2. 同一方向不同等级下，响度曲线幅度是否单调变化（从MAX到MIN依次递变）？',
  '3. 不同等级下，响度数值与曲线幅度是否互相印证？',
  '4. 是否存在波峰波谷等异常？',
  '',
  '## 输出格式（严格JSON）',
  '{',
  '  "overallAssessment": "综合评估（中文，150字以内）",',
  '  "overallSeverity": "pass"',
  '}',
  'severity: pass(一致), warning(轻微差异), error(明显异常)'
].join('\n');

async function analyzeChartImagesSequential(params, onProgress) {
  var images = params.images || [];
  var testDataFacts = params.testDataFacts || {};
  var llmSettings = params.settings || {};

  var apiUrl = String(llmSettings.apiUrl || '').replace(/\/+$/, '');
  var apiKey = String(llmSettings.apiKey || '');
  var model = String(llmSettings.model || 'claude-sonnet-4-20250514');

  if (!apiUrl || !apiKey) {
    return { issues: [{ severity: 'review', message: '未配置API凭据' }], evidence: [], status: 'review' };
  }

  var loudnessSummary = buildLoudnessSummary(testDataFacts);
  var allFindings = [];
  var evidence = [];

  for (var i = 0; i < images.length; i++) {
    var img = images[i];
    var ctx = (img.contextText || img.fileName || '?').slice(0, 200);
    var prompt = SINGLE_IMAGE_PROMPT
      .replace('%CONTEXT%', ctx)
      .replace('%VOLUME_LEVEL%', img.volumeLevel || '未知')
      .replace('%LOUDNESS_DATA%', loudnessSummary);

    if (onProgress) {
      try { onProgress({ current: i + 1, total: images.length, fileName: img.fileName, status: 'analyzing' }); } catch (_) {}
    }

    try {
      var response = await axios.post(apiUrl + '/v1/chat/completions', {
        model: model,
        max_tokens: 512,
        temperature: 0.1,
        messages: [{ role: 'user', content: [
          { type: 'text', text: prompt },
          { type: 'image_url', image_url: { url: 'data:' + img.contentType + ';base64,' + img.base64 } }
        ]}],
        response_format: { type: 'json_object' }
      }, {
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
        timeout: 60000
      });

      var text = '';
      var choice = (response.data && response.data.choices && response.data.choices[0]);
      if (choice && choice.message && choice.message.content) {
        text = choice.message.content;
      }
      var parsed = parseLlmResponse(text);
      var finding = parsed.findings[0] || {};
      finding.imageContext = ctx;
      finding.fileName = img.fileName;
      allFindings.push(finding);
      evidence.push('[' + (i + 1) + '/' + images.length + '] ' + img.fileName + ': ' + (finding.detail || finding.trendDescription || 'OK'));

      if (onProgress) {
        try { onProgress({ current: i + 1, total: images.length, fileName: img.fileName, status: finding.severity || 'pass', detail: finding.detail || '' }); } catch (_) {}
      }
    } catch (e) {
      evidence.push('[' + (i + 1) + '/' + images.length + '] ' + img.fileName + ': 分析失败 - ' + (e.message || '未知错误'));
      allFindings.push({ severity: 'review', imageContext: ctx, fileName: img.fileName, detail: '分析失败: ' + (e.message || '') });
      if (onProgress) {
        try { onProgress({ current: i + 1, total: images.length, fileName: img.fileName, status: 'error', detail: e.message || '' }); } catch (_) {}
      }
    }

    // 请求间隔，避免触发 API 限流
    if (i < images.length - 1) {
      await new Promise(function(r) { setTimeout(r, 500); });
    }
  }

  // 汇总请求
  var passCount = 0, warnCount = 0, errCount = 0;
  allFindings.forEach(function(f) {
    if (f.severity === 'error') errCount++;
    else if (f.severity === 'warning') warnCount++;
    else passCount++;
  });

  var findingsText = allFindings.map(function(f, idx) {
    return (idx + 1) + '. [' + (f.direction || '?') + '] ' + (f.imageContext || f.fileName || '') + ' - ' + (f.detail || '') + ' (' + (f.severity || 'pass') + ')';
  }).join('\n');

  var summaryPrompt = SUMMARY_PROMPT
    .replace('%FINDINGS%', findingsText)
    .replace('%LOUDNESS_DATA%', loudnessSummary);

  var overallSeverity = 'pass';
  var overallAssessment = '';
  try {
    var sumResp = await axios.post(apiUrl + '/v1/chat/completions', {
      model: model,
      max_tokens: 512,
      temperature: 0.1,
      messages: [{ role: 'user', content: summaryPrompt }],
      response_format: { type: 'json_object' }
    }, {
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
      timeout: 30000
    });

    var sumText = '';
    var sumChoice = (sumResp.data && sumResp.data.choices && sumResp.data.choices[0]);
    if (sumChoice && sumChoice.message && sumChoice.message.content) {
      sumText = sumChoice.message.content;
    }
    var sumParsed = parseLlmResponse(sumText);
    overallSeverity = sumParsed.overallSeverity || 'pass';
    overallAssessment = sumParsed.overallAssessment || '';
  } catch (e) {
    overallAssessment = '汇总请求失败，基于单图分析结果判定';
    if (errCount > 0) overallSeverity = 'error';
    else if (warnCount > 0) overallSeverity = 'warning';
  }

  evidence.push('分析完成: ' + passCount + ' 通过' + (warnCount > 0 ? ', ' + warnCount + ' 警告' : '') + (errCount > 0 ? ', ' + errCount + ' 失败' : ''));
  if (overallAssessment) {
    evidence.push('综合评估: ' + overallAssessment);
  }

  var issues = [];
  if (overallSeverity === 'error' || overallSeverity === 'warning') {
    issues.push({ severity: overallSeverity, message: (overallSeverity === 'error' ? '图表趋势异常' : '图表趋势有轻微差异') + ': ' + (overallAssessment || '请查看详情') });
  }

  return { issues: issues, evidence: evidence, status: overallSeverity === 'error' ? 'error' : (overallSeverity === 'warning' ? 'warning' : 'pass'), rawFindings: allFindings };
}

var GROUP_COMPARE_PROMPT = [
  '你是一位专业的音频测试工程师。请对比分析以下按测试等级分组的曲线图。',
  '',
  '## 图片说明',
  '- 标注为"FR基准-等级"的是该等级的独立频响测试曲线（来自 Sensitivity, frequency RCV/SND）',
  '- 标注为"RLR响度-等级"的是该等级下的接收响度测试曲线（来自 Loudness Rating RCV (RLR)）',
  '- 标注为"SLR响度-等级"的是该等级下的发送响度测试曲线（来自 Loudness Rating SND (SLR)）',
  '- 测试等级体系（8级，音量从大到小）: MAX > MAX-1 > MAX-2 > MAX-3(NOM) > MAX-4 > MAX-5 > MAX-6 > MAX-7(MIN)',
  '- 每张图为X-Y折线图：X轴频率(Hz，低频在左高频在右)，Y轴幅度(dB)',
  '',
  '## 对比要求',
  '',
  '### 要求1：同等级FR-响度趋势对比',
  '对于每个测试等级（尤其是MAX, NOM, MIN），将该等级下的FR基准图与对应响度曲线图进行视觉对比：',
  '1. 观察FR基准的曲线整体形状：低频段/中频段/高频段的走势、峰谷位置',
  '2. 对比同等级响度图：曲线峰谷位置、上升/下降趋势、高频衰减特征是否与FR基准一致',
  '3. 不一致（峰谷相反、趋势走向不同）→ 判定为异常',
  '',
  '### 要求2：跨等级响度曲线单调性',
  '同一方向下从MAX到MIN的响度曲线幅度应单调递减：',
  '1. MAX等级曲线幅度应最高，MIN等级应最低',
  '2. 中间等级（MAX-1到MAX-6）曲线幅度应依次递变',
  '3. 如果出现波峰波谷（某中间等级幅度反超相邻等级）→ 判定为异常',
  '',
  '### 要求3：RCV方向逆相关确认',
  'RCV方向：RLR值越小（如-9dB）→音量越大→曲线应整体更高',
  '',
  '## 图片组', '%IMAGE_GROUP%',
  '',
  '## 响度数值参考', '%LOUDNESS_DATA%',
  '',
  '## 输出格式（严格JSON，只输出JSON）',
  '每个finding必须包含具体的异常定位信息，便于工程师直接定位问题：',
  '{',
  '  "direction": "RCV或SND",',
  '  "findings": [',
  '    {',
  '      "imageIndex": 1,',
  '      "volumeLevel": "MAX或NOM或MIN或unknown",',
  '      "role": "reference或loudness_rlr或loudness_slr",',
  '      "trendMatchesReference": true,',
  '      "levelMonotonic": true,',
  '      "severity": "pass或warning或error",',
  '      "frequencyRange": "异常所在频段，如200-500Hz或1-3kHz。无异常时留空",',
  '      "expectedBehavior": "期望趋势，如：1kHz以上应平坦衰减",',
  '      "actualBehavior": "实际趋势，如：2kHz处出现5dB谷值",',
  '      "detail": "具体观察结论（中文）：曲线趋势、与基准的对比、异常点说明"',
  '    }',
  '  ],',
  '  "monotonicityViolations": [',
  '    { "direction": "RCV", "levelA": "MAX-1", "levelB": "MAX-2", "description": "MAX-2曲线幅度反超MAX-1约3dB，违反单调递减" }',
  '  ],',
  '  注意：只有在明确发现跨等级曲线幅度异常时，才输出 monotonicityViolations。若样本不足、缺少等级、无法判断，请返回空数组 []，不要输出“无法验证/缺少数据”类占位描述。',
  '  "overallAssessment": "综合判定：(1)各等级下FR与响度曲线趋势是否一致 (2)跨等级单调性 (3)具体异常点总结 (4)建议（中文，150字以内）",',
  '  "overallSeverity": "pass"',
  '}',
  'severity: pass=所有曲线趋势一致且单调, warning=个别有轻微偏离, error=明显不一致或严重异常'
].join('\n');

var LEVEL_COMPARE_PROMPT = [
  '你是一位专业的音频测试工程师。请只分析下面这一组“同方向、同等级”的曲线图。',
  '',
  '## 图片说明',
  '- 这一组图片都属于同一方向、同一测试等级',
  '- FR基准图来自 Sensitivity, frequency',
  '- RLR/SLR 响度图来自 Loudness Rating RCV/SND',
  '- 你的任务是判断该等级下响度曲线与频响基准的走势是否一致或相似',
  '',
  '## 判定要求',
  '1. 只比较这一组内部的同等级图片，不要做跨等级判断',
  '2. 观察低频/中频/高频的整体走势、峰谷位置、衰减趋势是否一致或相似',
  '3. 如果缺少FR图或缺少响度图，请返回 review 语义，不要硬判通过',
  '',
  '## 图片组', '%IMAGE_GROUP%',
  '',
  '## 输出格式（严格JSON）',
  '{',
  '  "findings": [',
  '    {',
  '      "direction": "RCV或SND",',
  '      "volumeLevel": "MAX或MAX-3(NOM)或MAX-7(MIN)",',
  '      "role": "reference或loudness_rlr或loudness_slr",',
  '      "trendConsistent": true,',
  '      "severity": "pass或warning或error",',
  '      "frequencyRange": "异常频段，无异常留空",',
  '      "expectedBehavior": "期望走势",',
  '      "actualBehavior": "实际走势",',
  '      "detail": "结论（中文）"',
  '    }',
  '  ],',
  '  "overallAssessment": "该等级对比结论（中文，80字以内）",',
  '  "overallSeverity": "pass"',
  '}'
].join('\n');

var MONOTONICITY_PROMPT = [
  '你是一位专业的音频测试工程师。请只分析下面同一方向下不同等级的响度曲线。',
  '',
  '## 分析目标',
  '1. 判断从 MAX 到 MIN，响度曲线整体幅度是否单调递变',
  '2. 结合给出的响度数值，判断曲线高低排序是否与数值变化趋势一致',
  '3. 如果样本不足或缺失等级，请返回空的 monotonicityViolations，不要编造异常',
  '',
  '## 图片组', '%IMAGE_GROUP%',
  '',
  '## 响度数值', '%LOUDNESS_DATA%',
  '',
  '## 输出格式（严格JSON）',
  '{',
  '  "findings": [',
  '    {',
  '      "direction": "RCV或SND",',
  '      "volumeLevel": "ALL",',
  '      "trendConsistent": true,',
  '      "levelMonotonic": true,',
  '      "severity": "pass或warning或error",',
  '      "detail": "跨等级单调性与数值印证结论（中文）"',
  '    }',
  '  ],',
  '  "monotonicityViolations": [',
  '    { "direction": "RCV", "levelA": "MAX-1", "levelB": "MAX-2", "description": "异常说明" }',
  '  ],',
  '  "overallAssessment": "该方向跨等级结论（中文，80字以内）",',
  '  "overallSeverity": "pass"',
  '}'
].join('\n');

/**
 * 主要对比等级（频响通常只测MAX/NOM/MIN）
 */
var COMPARE_LEVELS = ['MAX', 'MAX-3(NOM)', 'MAX-7(MIN)'];

function _groupByLevel(imgs, guessDir) {
  var groups = { RCV: {}, SND: {}, unknown: {} };
  ['RCV', 'SND', 'unknown'].forEach(function(d) {
    COMPARE_LEVELS.forEach(function(l) { groups[d][l] = []; });
    groups[d]['other'] = [];
  });

  imgs.forEach(function(img) {
    var dir = guessDir(img.contextText || '', img.category || '');
    var lvl = img.volumeLevel || 'other';
    if (COMPARE_LEVELS.indexOf(lvl) === -1 && lvl !== 'other') lvl = 'other';
    if (!groups[dir]) dir = 'unknown';
    if (!groups[dir][lvl]) lvl = 'other';
    groups[dir][lvl].push(img);
  });
  return groups;
}

function _buildImgLabels(batchImgs) {
  return batchImgs.map(function(img, i) {
    var role = img.category === 'fr_reference' ? 'FR基准'
      : (img.category === 'loudness_rlr' ? 'RLR响度' : 'SLR响度');
    var lvlLabel = img.volumeLevel ? '[' + img.volumeLevel + ']' : '[等级未知]';
    return (i + 1) + '. ' + lvlLabel + ' ' + role + ' ' + (img.contextText || img.fileName).slice(0, 150);
  }).join('\n');
}

function _pickRepresentativeImages(imgs, limit) {
  if (!Array.isArray(imgs) || imgs.length === 0) return [];
  return imgs.slice(0, Math.max(1, limit || 1));
}

function _buildSequentialTaskPlans(dir, frGroups, loudGroups) {
  var plans = [];

  COMPARE_LEVELS.forEach(function(lvl) {
    var refs = _pickRepresentativeImages((frGroups[dir][lvl] || []).concat(frGroups.unknown[lvl] || []), 1);
    var louds = _pickRepresentativeImages((loudGroups[dir][lvl] || []).concat(loudGroups.unknown[lvl] || []), 2);
    if (refs.length === 0 || louds.length === 0) return;
    plans.push({
      type: 'level',
      direction: dir,
      level: lvl,
      label: dir + ' ' + lvl,
      images: refs.concat(louds)
    });
  });

  return plans;
}

async function _postChartTask(http, apiUrl, apiKey, model, prompt, batchImgs) {
  var content = [{ type: 'text', text: prompt }];
  batchImgs.forEach(function(img) {
    content.push({ type: 'image_url', image_url: { url: 'data:' + img.contentType + ';base64,' + img.base64 } });
  });

  var response = await http.post(apiUrl + '/v1/chat/completions', {
    model: model,
    max_tokens: 1200,
    temperature: 0.1,
    messages: [{ role: 'user', content: content }],
    response_format: { type: 'json_object' }
  }, {
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
    timeout: 90000
  });

  var text = '';
  var choice = (response.data && response.data.choices && response.data.choices[0]);
  if (choice && choice.message && choice.message.content) text = choice.message.content;
  return parseLlmResponse(text);
}

/**
 * 逐张发射图片准备进度（异步延迟，让前端能看到每张图片的序号变化）
 */
function _emitPerImageProgress(batchImgs, directionLabel, onProgress) {
  if (!onProgress) return;
  return new Promise(function(resolve) {
    var i = 0;
    function next() {
      if (i >= batchImgs.length) { resolve(); return; }
      try {
        onProgress({
          current: i + 1,
          total: batchImgs.length,
          fileName: batchImgs[i].fileName || '',
          status: 'preparing',
          imageCount: batchImgs.length,
          detail: '[' + directionLabel + '] 加载图片 ' + (i + 1) + '/' + batchImgs.length + ': ' + (batchImgs[i].fileName || '')
        });
      } catch (_) {}
      i++;
      setTimeout(next, 25);
    }
    next();
  });
}

async function analyzeGroupedCharts(params, onProgress) {
  var images = params.images || [];
  var testDataFacts = params.testDataFacts || {};
  var llmSettings = params.settings || {};

  var apiUrl = String(llmSettings.apiUrl || '').replace(/\/+$/, '');
  var apiKey = String(llmSettings.apiKey || '');
  var model = String(llmSettings.model || 'claude-sonnet-4-20250514');

  if (!apiUrl || !apiKey) {
    return { issues: [{ severity: 'review', message: '未配置API凭据' }], evidence: [], status: 'review' };
  }

  var frRefs = images.filter(function(img) { return img.category === 'fr_reference'; });
  var loudnessImgs = images.filter(function(img) { return img.category === 'loudness_rlr' || img.category === 'loudness_slr'; });

  var logs = [];
  logs.push('FR参考图 ' + frRefs.length + ' 张, 响度图 ' + loudnessImgs.length + ' 张');

  if (loudnessImgs.length === 0) {
    return { issues: [{ severity: 'review', message: '未找到RLR/SLR响度测试图片，无法进行趋势对比' }], evidence: [], logs: logs, checklist: [], status: 'review' };
  }

  function guessDirection(ctx, category) {
    if (/RCV|RX|Receiving|接收/i.test(ctx)) return 'RCV';
    if (/SND|TX|Sending|发送/i.test(ctx)) return 'SND';
    if (category === 'loudness_rlr') return 'RCV';
    if (category === 'loudness_slr') return 'SND';
    return 'unknown';
  }

  var frGroups = _groupByLevel(frRefs, guessDirection);
  var loudGroups = _groupByLevel(loudnessImgs, guessDirection);

  // 统计各方向各等级分布
  var dirLevels = {};
  ['RCV', 'SND'].forEach(function(dir) {
    dirLevels[dir] = [];
    COMPARE_LEVELS.concat(['other']).forEach(function(lvl) {
      var count = (frGroups[dir][lvl] || []).length + (loudGroups[dir][lvl] || []).length;
      if (count > 0) dirLevels[dir].push(lvl + '(' + count + '张)');
    });
  });
  logs.push('分组: RCV=' + (dirLevels['RCV'].join(',') || '(空)') + '; SND=' + (dirLevels['SND'].join(',') || '(空)'));

  var allFindings = [];
  var dirAssessments = [];
  var failedDirections = [];
  var incompleteDirections = [];
  var loudnessSummary = buildLoudnessSummary(testDataFacts);
  var http = require('axios');
  var taskPlans = [];

  ['RCV', 'SND'].forEach(function(dir) {
    var hasFrReference = COMPARE_LEVELS.some(function(lvl) {
      return (frGroups[dir][lvl] || []).length > 0 || (frGroups.unknown[lvl] || []).length > 0;
    });
    var hasLoudness = CANONICAL_LEVELS.some(function(lvl) {
      return (loudGroups[dir][lvl] || []).length > 0 || (loudGroups.unknown[lvl] || []).length > 0;
    });

    if (!hasLoudness) return;
    if (!hasFrReference) {
      incompleteDirections.push(dir);
      logs.push(dir + '方向缺少FR参考图，无法按规则完成同等级趋势对比');
      return;
    }

    var plans = _buildSequentialTaskPlans(dir, frGroups, loudGroups);
    if (plans.length === 0) {
      incompleteDirections.push(dir);
      logs.push(dir + '方向未形成有效的小组任务，已标记人工复核');
      return;
    }
    taskPlans = taskPlans.concat(plans);
  });

  var totalTasks = taskPlans.length;
  var taskDone = 0;

  for (var task of taskPlans) {
    await _emitPerImageProgress(task.images, task.label, onProgress);

    if (task.type === 'level') {
      var referenceImage = task.images.find(function(img) { return img.category === 'fr_reference'; });
      var loudnessImage = task.images.find(function(img) { return img.category === 'loudness_rlr' || img.category === 'loudness_slr'; });
      if (referenceImage && loudnessImage && haveIdenticalPlotBody(referenceImage, loudnessImage)) {
        if (taskMatchesSwbRule(task, testDataFacts)) {
          logs.push(task.label + ': 检测到FR与响度图图表主体一致，但当前为SWB报告，按已确认规则允许');
        } else {
        allFindings.push(buildDuplicatePlotFinding(task, referenceImage, loudnessImage));
        logs.push(task.label + ': 检测到FR与响度图图表主体完全一致，疑似复用同一底图');
        taskDone++;
        if (onProgress) {
          try { onProgress({ current: taskDone, total: totalTasks, fileName: task.label, status: 'error', imageCount: task.images.length, detail: task.label + ' 检测到疑似重复底图' }); } catch (_) {}
        }
        continue;
        }
      }
    }

    var imgDescs = _buildImgLabels(task.images);
    var prompt = LEVEL_COMPARE_PROMPT.replace('%IMAGE_GROUP%', imgDescs);

    if (onProgress) {
      try { onProgress({ current: taskDone + 1, total: totalTasks, fileName: task.label, status: 'analyzing', imageCount: task.images.length, detail: '正在发送 ' + task.label + ' 共 ' + task.images.length + ' 张曲线图给AI分析...' }); } catch (_) {}
    }

    try {
      var parsed = await _postChartTask(http, apiUrl, apiKey, model, prompt, task.images);
      parsed.findings.forEach(function(f) {
        f.direction = f.direction || task.direction;
        if (task.level && (!f.volumeLevel || f.volumeLevel === 'unknown')) f.volumeLevel = task.level;
      });
      var filteredFindings = parsed.findings.filter(function(f) {
        return !isAliasOnlyLevelMismatchFinding(f);
      });
      if (filteredFindings.length !== parsed.findings.length) {
        logs.push(task.label + ': 已忽略 ' + (parsed.findings.length - filteredFindings.length) + ' 条仅由等级别名差异触发的误报');
      }
      allFindings = allFindings.concat(filteredFindings);
      if (parsed.overallAssessment) {
        dirAssessments.push({ direction: task.direction, assessment: '[' + task.label + '] ' + parsed.overallAssessment });
      }
      logs.push(task.label + ': ' + (parsed.overallAssessment || 'OK'));

    } catch (e) {
      if (failedDirections.indexOf(task.direction) === -1) failedDirections.push(task.direction);
      logs.push(task.label + ' 分析失败: ' + (e.message || '未知错误'));
    }

    taskDone++;
    if (onProgress) {
      try { onProgress({ current: taskDone, total: totalTasks, fileName: task.label, status: 'done', imageCount: task.images.length, detail: task.label + ' 分析完成' }); } catch (_) {}
    }
  }

  var passCount = 0, warnCount = 0, errCount = 0;
  allFindings.forEach(function(f) {
    if (f.severity === 'error') errCount++;
    else if (f.severity === 'warning') warnCount++;
    else passCount++;
  });
  logs.push('趋势对比: ' + passCount + ' 通过' + (warnCount > 0 ? ', ' + warnCount + ' 警告' : '') + (errCount > 0 ? ', ' + errCount + ' 异常' : '') + (incompleteDirections.length > 0 ? ', ' + incompleteDirections.length + ' 方向缺少FR参考图' : '') + (failedDirections.length > 0 ? ', ' + failedDirections.length + ' 方向分析失败' : ''));

  var overallSeverity = errCount > 0 ? 'error' : (warnCount > 0 ? 'warning' : 'pass');
  var overallStatus = (frRefs.length === 0 || incompleteDirections.length > 0 || failedDirections.length > 0 || allFindings.length === 0)
    ? 'review'
    : overallSeverity;

  // 每个异常finding生成一条issue
  var issues = [];
  if (frRefs.length === 0) {
    issues.push({ severity: 'review', message: '未识别到可用于对比的频响参考图，无法按判定规则完成趋势一致性分析' });
  }
  incompleteDirections.forEach(function(dir) {
    issues.push({ severity: 'review', message: dir + '方向缺少频响参考图，无法完成同等级趋势对比' });
  });
  failedDirections.forEach(function(dir) {
    issues.push({ severity: 'review', message: dir + '方向AI分析失败，请人工复核' });
  });
  allFindings.forEach(function(f, i) {
    if (f.severity !== 'error' && f.severity !== 'warning') return;
    var parts = [];
    if (f.direction) parts.push('[' + f.direction + ']');
    if (f.volumeLevel && f.volumeLevel !== 'ALL') parts.push(f.volumeLevel);
    if (f.role === 'reference') parts.push('(FR基准)');
    else if (f.role === 'loudness_rlr') parts.push('(RLR响度)');
    else if (f.role === 'loudness_slr') parts.push('(SLR响度)');
    if (f.imageIndex != null) parts.push('图#' + f.imageIndex);
    if (f.detail) parts.push(f.detail);
    issues.push({
      severity: f.severity,
      message: parts.join(' '),
      meta: {
        direction: f.direction,
        volumeLevel: f.volumeLevel,
        imageIndex: f.imageIndex,
        frequencyRange: f.frequencyRange,
        role: f.role,
        detail: f.detail,
        expectedBehavior: f.expectedBehavior,
        actualBehavior: f.actualBehavior,
        trendConsistent: f.trendConsistent,
        levelMonotonic: f.levelMonotonic,
      }
    });
  });

  if (issues.length === 0 && overallStatus === 'review') {
    issues.push({ severity: 'review', message: '可比的频响/响度曲线不足，当前结果需人工复核' });
  }

  var conclusion = '';
  if (dirAssessments.length === 1) {
    conclusion = dirAssessments[0].assessment;
  } else if (dirAssessments.length > 1) {
    conclusion = dirAssessments.map(function(d) { return '[' + d.direction + '] ' + d.assessment; }).join('; ');
  }

  // 诊断相关evidence（不含操作日志）
  var diagnosticEvidence = [];
  // 完整逐项清单（包含所有finding的pass/warn/error）
  var checklist = allFindings.map(function(f) {
    return {
      direction: f.direction || '?',
      volumeLevel: f.volumeLevel || '',
      role: f.role || '',
      imageIndex: f.imageIndex,
      status: f.severity || 'pass',
      detail: f.detail || '',
      trendMatchesReference: f.trendConsistent,
      levelMonotonic: f.levelMonotonic,
      frequencyRange: f.frequencyRange || '',
      expectedBehavior: f.expectedBehavior || '',
      actualBehavior: f.actualBehavior || '',
    };
  });

  return {
    issues: issues,
    evidence: diagnosticEvidence,
    logs: logs,
    checklist: checklist,
    status: overallStatus,
    conclusion: conclusion,
    rawFindings: allFindings,
    monotonicityViolations: [],
  };
}

module.exports = { analyzeChartImages, analyzeChartImagesSequential, analyzeGroupedCharts };
