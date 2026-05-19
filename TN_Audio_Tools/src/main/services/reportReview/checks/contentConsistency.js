const progressBus = require('../progressBus');
const { normalizeVolumeLevel, compareVolumeLevel } = require('../volumeLevelUtils');

/**
 * 测试内容合理性检查 (2.2)
 *
 * 验证报告内及跨报告的测试数值合理性：
 * - 2.2.1 & 2.2.2: 跨报告响度一致性 (batch mode)
 * - 2.2.3: 单报告内 TX/RX 响度与频响趋势一致性
 * - 2.2.4: 单报告内曲线与数值互相印证
 */

function requireTestData(testDataFacts) {
  if (!testDataFacts) {
    return { ok: false, reason: '非ACQUA xlsx格式报告，无法进行内容合理性自动检查。请提供ACQUA导出的.xlsx测试报告。' };
  }
  if (!testDataFacts.loudnessMetrics || testDataFacts.loudnessMetrics.length === 0) {
    return { ok: false, reason: '未从报告中提取到响度相关数值，无法进行内容合理性检查。' };
  }
  return { ok: true, reason: '' };
}

/**
 * 2.2.1 同codec不同网络响度差异≤1dB（跨报告）
 * 需要在批量对比时调用，单报告模式下标记为review
 */
function checkSameCodecDifferentNetworkLoudness(allMetrics) {
  const issues = [];
  const evidence = [];

  if (!allMetrics || allMetrics.length < 2) {
    evidence.push('需要至少2份不同网络的报告才能进行此项检查');
    return { issues: [{ severity: 'review', message: '需要至少2份不同网络的报告进行对比' }], evidence, status: 'review' };
  }

  const flatMetrics = [];
  allMetrics.forEach(({ reportPath, metrics, codec, network }) => {
    (metrics || []).forEach((m) => {
      flatMetrics.push({ ...m, reportPath, codec: codec || m.codec, network: network || m.network });
    });
  });

  // 按 codec + direction + category (RLR/SLR/STMR) 分组
  const groups = {};
  flatMetrics.forEach((m) => {
    const key = `${m.codec || '?'}|${m.direction || '?'}|${m.category || '?'}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(m);
  });

  Object.entries(groups).forEach(([key, metrics]) => {
    if (metrics.length < 2) return;

    // 检查不同 network 之间的差异
    const byNetwork = {};
    metrics.forEach((m) => {
      const net = m.network || 'unknown';
      if (!byNetwork[net]) byNetwork[net] = [];
      byNetwork[net].push(m.value);
    });

    const networks = Object.keys(byNetwork);
    if (networks.length < 2) return;

    const [codec, direction, category] = key.split('|');
    evidence.push(`${codec} ${direction} ${category}: ${networks.length} 个网络 (${networks.join(', ')})`);

    for (let i = 0; i < networks.length; i++) {
      for (let j = i + 1; j < networks.length; j++) {
        const avgA = byNetwork[networks[i]].reduce((a, b) => a + b, 0) / byNetwork[networks[i]].length;
        const avgB = byNetwork[networks[j]].reduce((a, b) => a + b, 0) / byNetwork[networks[j]].length;
        const diff = Math.abs(avgA - avgB);

        if (diff > 3) {
          issues.push({
            severity: 'error',
            message: `${codec} ${direction} ${category}: ${networks[i]}(${avgA.toFixed(2)}dB) vs ${networks[j]}(${avgB.toFixed(2)}dB) 差异 ${diff.toFixed(2)}dB，严重超过1dB`,
          });
        } else if (diff > 1) {
          issues.push({
            severity: 'warning',
            message: `${codec} ${direction} ${category}: ${networks[i]}(${avgA.toFixed(2)}dB) vs ${networks[j]}(${avgB.toFixed(2)}dB) 差异 ${diff.toFixed(2)}dB，超过1dB`,
          });
        } else {
          evidence.push(`  ${networks[i]} vs ${networks[j]}: diff=${diff.toFixed(2)}dB ✓`);
        }
      }
    }
  });

  if (issues.length === 0) {
    evidence.push('✓ 同codec不同网络间响度差异均在1dB以内');
    return { issues: [], evidence, status: 'pass' };
  }

  return { issues, evidence, status: issues.some((i) => i.severity === 'error') ? 'error' : 'warning' };
}

/**
 * 2.2.2 同network不同codec响度差异≤1dB（跨报告）
 */
function checkSameNetworkDifferentCodecLoudness(allMetrics) {
  const issues = [];
  const evidence = [];

  if (!allMetrics || allMetrics.length < 2) {
    evidence.push('需要至少2份不同codec的报告才能进行此项检查');
    return { issues: [{ severity: 'review', message: '需要至少2份不同codec的报告进行对比' }], evidence, status: 'review' };
  }

  const flatMetrics = [];
  allMetrics.forEach(({ reportPath, metrics, codec, network }) => {
    (metrics || []).forEach((m) => {
      flatMetrics.push({ ...m, reportPath, codec: codec || m.codec, network: network || m.network });
    });
  });

  // 按 network + direction + category 分组
  const groups = {};
  flatMetrics.forEach((m) => {
    const key = `${m.network || '?'}|${m.direction || '?'}|${m.category || '?'}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(m);
  });

  Object.entries(groups).forEach(([key, metrics]) => {
    if (metrics.length < 2) return;

    const byCodec = {};
    metrics.forEach((m) => {
      const c = m.codec || 'unknown';
      if (!byCodec[c]) byCodec[c] = [];
      byCodec[c].push(m.value);
    });

    const codecs = Object.keys(byCodec);
    if (codecs.length < 2) return;

    const [network, direction, category] = key.split('|');
    evidence.push(`${network} ${direction} ${category}: ${codecs.length} 个codec (${codecs.join(', ')})`);

    for (let i = 0; i < codecs.length; i++) {
      for (let j = i + 1; j < codecs.length; j++) {
        const avgA = byCodec[codecs[i]].reduce((a, b) => a + b, 0) / byCodec[codecs[i]].length;
        const avgB = byCodec[codecs[j]].reduce((a, b) => a + b, 0) / byCodec[codecs[j]].length;
        const diff = Math.abs(avgA - avgB);

        if (diff > 3) {
          issues.push({
            severity: 'error',
            message: `${network} ${direction} ${category}: ${codecs[i]}(${avgA.toFixed(2)}dB) vs ${codecs[j]}(${avgB.toFixed(2)}dB) 差异 ${diff.toFixed(2)}dB，严重超过1dB`,
          });
        } else if (diff > 1) {
          issues.push({
            severity: 'warning',
            message: `${network} ${direction} ${category}: ${codecs[i]}(${avgA.toFixed(2)}dB) vs ${codecs[j]}(${avgB.toFixed(2)}dB) 差异 ${diff.toFixed(2)}dB，超过1dB`,
          });
        } else {
          evidence.push(`  ${codecs[i]} vs ${codecs[j]}: diff=${diff.toFixed(2)}dB ✓`);
        }
      }
    }
  });

  if (issues.length === 0) {
    evidence.push('✓ 同网络不同codec间响度差异均在1dB以内');
    return { issues: [], evidence, status: 'pass' };
  }

  return { issues, evidence, status: issues.some((i) => i.severity === 'error') ? 'error' : 'warning' };
}

/**
 * 2.2.3 单报告内 TX/RX 响度与频响趋势一致性
 *
 * 检查逻辑：同一份报告内，发送方向(SND/TX)和接收方向(RCV/RX)的响度测试
 * 如果有多条数据，其变化趋势应与对应方向的频响幅值趋势一致。
 */
async function checkLoudnessFrequencyResponseTrendConsistency(testDataFacts, reportPath, llmSettings) {
  var _ev = [];
  var _rd = requireTestData(testDataFacts);
  if (!_rd.ok) {
    _ev.push(_rd.reason);
    return { issues: [{ severity: 'review', message: _rd.reason }], evidence: _ev, status: 'review' };
  }

  var _m = testDataFacts.loudnessMetrics || [];
  var _f = testDataFacts.frequencyResponseMetrics || [];
  _ev.push('响度数据 ' + _m.length + ' 条，频响数据 ' + _f.length + ' 条');

  if (llmSettings && llmSettings.enabled && llmSettings.apiUrl && llmSettings.apiKey && reportPath) {
    try {
      var { extractReportImages } = require('../imageExtractor');
      var result = await extractReportImages(reportPath);
      // 兼容新旧返回格式
      var images = (result && result.images) ? result.images : (Array.isArray(result) ? result : []);
      var warnings = (result && result.warnings) ? result.warnings : [];
      if (warnings.length > 0) _ev = _ev.concat(warnings);
      if (images.length > 0) {
        var chartImages = images;
        _ev.push('AI分析 ' + chartImages.length + ' 张曲线图');
        var { analyzeGroupedCharts } = require('../llmService');
        var llmResult = await analyzeGroupedCharts({ images: chartImages, testDataFacts: testDataFacts, settings: llmSettings }, function(progress) {
          try { progressBus.emit('chart-progress', { imageCurrent: progress.current, imageTotal: progress.total, fileName: progress.fileName || '', imageCount: progress.imageCount || 0, status: progress.status || 'analyzing', detail: progress.detail || '' }); } catch (_) {}
        });
        if (llmResult.evidence) _ev = _ev.concat(llmResult.evidence);
        var frResult = {
          issues: llmResult.issues || [],
          evidence: _ev,
          logs: llmResult.logs || [],
          checklist: llmResult.checklist || [],
          status: llmResult.status || 'pass',
          conclusion: llmResult.conclusion || '',
          rawFindings: llmResult.rawFindings || [],
          monotonicityViolations: llmResult.monotonicityViolations || [],
        };
        return frResult;
      }
      _ev.push('未检测到响度/频响曲线图，需人工对比。报告中图片总数: ' + (images.length + (warnings ? warnings.length : 0)));
    } catch (e) {
      console.error('[checkLoudnessFR] AI分析异常:', e.message, e.stack);
      _ev.push('AI分析失败: ' + (e.message || '未知错误') + ' (步骤: ' + (e.stack ? e.stack.split('\n')[1] || '' : '').trim().slice(0, 80) + ')');
    }
  }

  if (!llmSettings) {
    _ev.push('AI 图表分析不可用：LLM 设置未传入');
  } else if (!llmSettings.enabled) {
    _ev.push('AI 图表分析未启用：请在设置 > AI 图表分析中开启并配置 API 凭据');
  } else if (!llmSettings.apiUrl || !llmSettings.apiKey) {
    _ev.push('AI 图表分析凭据不完整：请在设置中填写 API 地址和 Key');
  } else if (!reportPath) {
    _ev.push('无法提取图片：报告路径为空');
  }
  return { issues: [{ severity: 'review', message: '请人工对比报告中响度与频响测试项的曲线趋势是否一致' }], evidence: _ev, status: 'review' };
}
/**
 * 2.2.4 单报告内曲线与数值互相印证
 *
 * 检查逻辑：不同音量等级下，响度值越大，频响幅值也应越大。
 * 对相同方向、相同类型的测试，按音量等级排列后检查单调性。
 */
function checkCurveValueCorroboration(testDataFacts, llmContext) {
  const evidence = [];
  const issues = [];
  const checklist = [];
  const ready = requireTestData(testDataFacts);

  if (!ready.ok) {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, checklist, status: 'review' };
  }

  const metrics = testDataFacts.loudnessMetrics || [];
  const frMetrics = testDataFacts.frequencyResponseMetrics || [];

  // 按 volumeLevel 分组，检查单调性
  const volumeGroups = {};
  metrics.forEach((m) => {
    const vol = normalizeVolumeLevel(m.volumeLevel) || 'unknown';
    const dir = m.direction || 'unknown';
    const key = `${dir}|${vol}`;
    if (!volumeGroups[key]) volumeGroups[key] = [];
    volumeGroups[key].push(m.value);
  });

  const sortedVolumes = Object.keys(volumeGroups).sort((a, b) => {
    const volA = a.split('|')[1];
    const volB = b.split('|')[1];
    return compareVolumeLevel(volA, volB);
  });

  if (sortedVolumes.length < 2) {
    evidence.push('仅有1个音量等级的测试数据，无法进行跨音量印证');
    return { issues: [{ severity: 'review', message: '仅有1个音量水平，无法验证曲线与数值的跨音量一致性' }], evidence, checklist, chartData: {}, status: 'review' };
  }

  evidence.push(`共 ${sortedVolumes.length} 个音量分组`);

  // 按8级等级体系主动枚举匹配 — 三层回退确保覆盖各种xlsx格式
  var TARGET_LEVELS = require('../volumeLevelUtils').CANONICAL_LEVELS;
  var extractLevel = require('../volumeLevelUtils').extractVolumeLevelFromTitle;
  var byDir = {};
  var matchedCount = 0;

  // 为每个目标等级生成精确正则（防止MAX误匹配MAX-1等）
  function _buildLevelRE(targetLevel) {
    if (targetLevel === 'MAX')           return /\bMAX\b(?!\s*[-–]\s*\d|\s*\(?(?:NOM|MIN))/i;
    if (targetLevel === 'MAX-1')         return /\bMAX\s*[-–]\s*1\b/i;
    if (targetLevel === 'MAX-2')         return /\bMAX\s*[-–]\s*2\b/i;
    if (targetLevel === 'MAX-3(NOM)')    return /\b(?:MAX\s*[-–]\s*3|NOM)\b/i;
    if (targetLevel === 'MAX-4')         return /\bMAX\s*[-–]\s*4\b/i;
    if (targetLevel === 'MAX-5')         return /\bMAX\s*[-–]\s*5\b/i;
    if (targetLevel === 'MAX-6')         return /\bMAX\s*[-–]\s*6\b/i;
    if (targetLevel === 'MAX-7(MIN)')    return /\b(?:MAX\s*[-–]\s*7|MIN)\b/i;
    return new RegExp('\\b' + targetLevel.replace(/[-\s]/g, '[-\s]') + '\\b', 'i');
  }

  TARGET_LEVELS.forEach(function(targetLevel) {
    var levelRE = _buildLevelRE(targetLevel);
    metrics.forEach(function(m) {
      var lvl = normalizeVolumeLevel(m.volumeLevel);
      var desc = m.descriptor || '';

      // 优先从 descriptor（SMD标题）中提取精确等级：
      // xlsx 的 VolumeCTRL 列对中间等级（MAX-1~MAX-6）均写成 "MAX"（粗粒度），
      // 而 SMD 字段含有如 "BIN MAX-1 HHWB" 这样的精确等级信息。
      // 若 descriptor 能给出明确等级，以 descriptor 为准，忽略 VolumeCTRL 的粗粒度值。
      var descLevel = normalizeVolumeLevel(extractLevel(desc));
      var matched;
      if (descLevel) {
        matched = descLevel === targetLevel;
      } else {
        // descriptor 无法提取等级时，回退到 VolumeCTRL 和精确正则
        matched = (lvl && lvl === targetLevel) || levelRE.test(desc);
      }
      if (!matched) return;

      var dir = m.direction || 'unknown';
      if (!byDir[dir]) byDir[dir] = [];
      var alreadyAdded = byDir[dir].some(function(e) {
        return e.volume === targetLevel && e.descriptor === m.descriptor;
      });
      if (!alreadyAdded) {
        byDir[dir].push({ volume: targetLevel, value: m.value, descriptor: m.descriptor });
        matchedCount++;
      }
    });
  });
  evidence.push('主动匹配到 ' + matchedCount + ' 条响度数据 (' + Object.keys(byDir).length + ' 个方向)');

  // 统计缺失的等级
  var missingLevels = [];
  TARGET_LEVELS.forEach(function(targetLevel) {
    var found = false;
    Object.keys(byDir).forEach(function(dir) {
      if (byDir[dir].some(function(e) { return e.volume === targetLevel; })) found = true;
    });
    if (!found) missingLevels.push(targetLevel);
  });
  if (missingLevels.length > 0) {
    evidence.push('未找到数据的等级: ' + missingLevels.join(', '));
  }

  // 单调性容差：同方向相邻等级间变化≤0.5dB视为持平
  var MONO_TOLERANCE = 0.5;

  Object.entries(byDir).forEach(([dir, items]) => {
    const sorted = items.sort(function(a, b) {
      return compareVolumeLevel(a.volume, b.volume);
    });
    if (sorted.length < 2) {
      evidence.push(dir + '方向仅有 ' + sorted.length + ' 个可识别等级（需≥2），跳过单调性检查');
      return;
    }

    var increasing = 0;
    var decreasing = 0;
    var sameValue = 0;

    for (var i = 1; i < sorted.length; i++) {
      if (sorted[i].volume === sorted[i - 1].volume) continue;

      var diff = sorted[i].value - sorted[i - 1].value;
      if (diff > MONO_TOLERANCE) increasing += 1;
      else if (diff < -MONO_TOLERANCE) decreasing += 1;
      else sameValue += 1;

      // 每对过渡生成一条checklist项
      var fromVal = sorted[i - 1].value;
      var toVal = sorted[i].value;
      var isMonotonic = decreasing > increasing ? diff <= MONO_TOLERANCE : diff >= -MONO_TOLERANCE;
      checklist.push({
        direction: dir,
        fromLevel: sorted[i - 1].volume,
        toLevel: sorted[i].volume,
        fromValue: fromVal,
        toValue: toVal,
        status: isMonotonic ? 'pass' : (Math.abs(diff) < 1.0 ? 'warning' : 'error'),
        detail: fromVal.toFixed(2) + 'dB → ' + toVal.toFixed(2) + 'dB (' + (diff > 0 ? '+' : '') + diff.toFixed(2) + 'dB)',
      });
    }

    var total = increasing + decreasing + sameValue;
    if (total === 0) {
      evidence.push(dir + '方向仅1个音量级别，无法进行跨音量单调性对比');
      return;
    }

    var consistency = Math.max(increasing, decreasing) / total;
    evidence.push(dir + '方向跨音量单调性: ' + (consistency * 100).toFixed(0) + '% (' + increasing + '升/' + decreasing + '降/' + sameValue + '平)');

    if (consistency < 0.8) {
      issues.push({
        severity: 'error',
        message: dir + '方向不同音量下响度变化方向不一致（' + increasing + '次上升, ' + decreasing + '次下降, ' + sameValue + '次持平），跨音量单调性仅' + (consistency * 100).toFixed(0) + '%，需排查测试数据或曲线异常'
      });
    } else if (consistency < 1.0) {
      issues.push({
        severity: 'warning',
        message: dir + '方向跨音量单调性为' + (consistency * 100).toFixed(0) + '%（' + increasing + '升/' + decreasing + '降/' + sameValue + '平），存在少量不一致，建议复核'
      });
    }
  });

  // LLM交叉验证：将AI视觉发现与数值单调性结果对照
  if (llmContext && llmContext.monotonicityViolations && llmContext.monotonicityViolations.length > 0) {
    var llmViolations = llmContext.monotonicityViolations;
    evidence.push('AI视觉交叉验证: 发现 ' + llmViolations.length + ' 处跨等级曲线幅度异常');
    llmViolations.forEach(function(v) {
      var alreadyFlagged = issues.some(function(issue) {
        return issue.message.indexOf(v.direction) !== -1;
      });
      if (alreadyFlagged) {
        // LLM视觉确认了数值异常 → 提升置信度
        evidence.push('  [AI验证] ' + v.direction + ': ' + (v.description || v.levelA + '与' + v.levelB + '幅度异常') + ' — 与数值分析一致');
      } else {
        // LLM发现了数值分析未捕获的异常
        issues.push({
          severity: 'warning',
          message: '[AI视觉发现] ' + (v.direction || '?') + ': ' + (v.description || (v.levelA + '与' + v.levelB + '曲线幅度异常')),
          meta: { source: 'llm_verified', direction: v.direction, levelA: v.levelA, levelB: v.levelB, description: v.description }
        });
      }
    });
  }

  // 生成折线图数据（按方向，同等级取平均值，按等级排序）
  var chartData = {};
  Object.keys(byDir).forEach(function(dir) {
    // 先按等级分组，聚合求均值
    var levelGroups = {};
    byDir[dir].forEach(function(item) {
      var lvl = normalizeVolumeLevel(item.volume) || item.volume || 'unknown';
      if (!levelGroups[lvl]) levelGroups[lvl] = [];
      levelGroups[lvl].push(item.value);
    });
    var points = [];
    Object.keys(levelGroups).forEach(function(lvl) {
      var values = levelGroups[lvl];
      var avg = values.reduce(function(s, v) { return s + v; }, 0) / values.length;
      points.push({
        level: lvl,
        value: Number(avg.toFixed(2)),
        count: values.length,
      });
    });
    points.sort(function(a, b) {
      return compareVolumeLevel(a.level, b.level);
    });
    if (points.length >= 2) chartData[dir] = points;
  });

  if (frMetrics.length > 0) {
    evidence.push(`频响数据: ${frMetrics.length} 个测点`);
  }

  if (issues.length === 0) {
    evidence.push('✓ 不同音量下响度变化趋势一致');
    return { issues: [], evidence, checklist, chartData, status: 'pass' };
  }

  var hasError = issues.some(function(item) { return item.severity === 'error'; });
  var hasWarning = issues.some(function(item) { return item.severity === 'warning'; });
  if (hasError) return { issues, evidence, checklist, chartData, status: 'error' };
  if (hasWarning) return { issues, evidence, checklist, chartData, status: 'warning' };
  return { issues, evidence, checklist, chartData, status: 'review' };
}

module.exports = {
  checkSameCodecDifferentNetworkLoudness,
  checkSameNetworkDifferentCodecLoudness,
  checkLoudnessFrequencyResponseTrendConsistency,
  checkCurveValueCorroboration,
};
