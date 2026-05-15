const axios = require('axios');

var PROMPT_TEMPLATE = [
  '你是一位专业的音频测试工程师。请分析以下ACQUA音频测试报告中的频率响应图表图片，与数值数据进行交叉验证。',
  '',
  '## 图片特征',
  '- 图片为WMF矢量格式的X-Y折线图，每条曲线代表不同测试条件下的频率响应或响度',
  '- Sending Dir = 发送方向(SND/TX)的频响/响度曲线',
  '- Receiving Dir = 接收方向(RCV/RX)的频响/响度曲线',
  '- MAX-1到MAX-7 = 不同音量级别下的响度曲线',
  '- NOM = 标称音量下的曲线',
  '',
  '## 检查要求',
  '1. 观察每张图中曲线的整体形态趋势（随频率升高是上升/下降/平坦）',
  '2. 对比同一方向（SND或RCV）上不同音量级别的曲线：音量增大时曲线是否同向偏移',
  '3. 判断响度测试项与频响测试项的曲线趋势是否一致',
  '4. 注意：RCV接收方向RLR数值越小代表声音越大（逆相关），这是正常的',
  '',
  '## 数值数据',
  '%LOUDNESS_DATA%',
  '',
  '%FREQ_RESPONSE_DATA%',
  '',
  '## 图片上下文',
  '每张图片附带了从报告中提取的上下文标签，请结合标签判断图片类型。',
  '',
  '## 输出格式',
  '请严格按以下JSON格式输出，不要输出任何其他内容：',
  '{',
  '  "findings": [',
  '    {',
  '      "direction": "SND或RCV或unknown",',
  '      "imageContext": "图片上下文摘要",',
  '      "trendConsistent": true,',
  '      "detail": "具体观察到的曲线趋势描述（中文）",',
  '      "severity": "pass"',
  '    }',
  '  ],',
  '  "overallAssessment": "综合所有图片的整体评估结论（中文，100字以内）",',
  '  "overallSeverity": "pass"',
  '}',
  '',
  'severity取值：pass(趋势一致)、warning(有轻微差异)、error(趋势明显相反或异常)'
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
    lines.push(dir + '方向: ' + byDir[dir].length + '个测点');
    byDir[dir].slice(0, 10).forEach(function(m) {
      lines.push('  音量' + (m.volume || '?') + ': ' + (m.value != null ? Number(m.value).toFixed(2) : '?') + 'dB' + (m.category ? ' (' + m.category + ')' : ''));
    });
    if (byDir[dir].length > 10) lines.push('  ... 共' + byDir[dir].length + '个测点');
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
    lines.push(dir + '方向: ' + byDir[dir].length + '个测点');
    byDir[dir].slice(0, 10).forEach(function(m) {
      lines.push('  音量' + (m.volume || '?') + ': 幅度' + (m.amplitude != null ? Number(m.amplitude).toFixed(2) + 'dB' : '?'));
    });
    if (byDir[dir].length > 10) lines.push('  ... 共' + byDir[dir].length + '个测点');
  });
  return lines.join('\n');
}

function parseLlmResponse(text) {
  var findings = [];
  var overallAssessment = '';
  var overallSeverity = 'review';

  try {
    // 尝试提取 JSON 块
    var jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      var parsed = JSON.parse(jsonMatch[0]);
      if (Array.isArray(parsed.findings)) {
        findings = parsed.findings.map(function(f) {
          return {
            severity: f.severity === 'error' ? 'error' : (f.severity === 'warning' ? 'warning' : 'pass'),
            detail: f.detail || '',
            direction: f.direction || 'unknown',
            trendConsistent: f.trendConsistent
          };
        });
      }
      overallAssessment = parsed.overallAssessment || '';
      if (parsed.overallSeverity === 'error') overallSeverity = 'error';
      else if (parsed.overallSeverity === 'warning') overallSeverity = 'warning';
      else if (parsed.overallSeverity === 'pass') overallSeverity = 'pass';
    }
  } catch (_) {
    // JSON解析失败，尝试从文本中提取关键信息
    if (/不一致|相反|异常|错误|error/i.test(text)) overallSeverity = 'warning';
    overallAssessment = text.slice(0, 200);
  }

  return { findings, overallAssessment, overallSeverity };
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

    if (parsed.overallAssessment) {
      evidence.push('AI综合评估: ' + parsed.overallAssessment);
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

module.exports = { analyzeChartImages };
