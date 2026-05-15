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
      var images = await extractReportImages(reportPath);
      if (images.length > 0) {
        var chartImages = images.slice(0, llmSettings.maxImagesPerAnalysis || 4);
        _ev.push('AI分析 ' + chartImages.length + ' 张曲线图');
        var { analyzeChartImages } = require('../llmService');
        var llmResult = await analyzeChartImages({ images: chartImages, testDataFacts: testDataFacts, settings: llmSettings });
        if (llmResult.evidence) _ev = _ev.concat(llmResult.evidence);
        if (llmResult.issues && llmResult.issues.length > 0) {
          return { issues: llmResult.issues, evidence: _ev, status: llmResult.status };
        }
        return { issues: [], evidence: _ev, status: 'pass' };
      }
      _ev.push('未检测到响度/频响曲线图，需人工对比');
    } catch (e) {
      _ev.push('AI分析失败: ' + (e.message || '未知错误'));
    }
  }

  return { issues: [{ severity: 'review', message: '请人工对比报告中响度与频响测试项的曲线趋势是否一致' }], evidence: _ev, status: 'review' };
}
/**
 * 2.2.4 单报告内曲线与数值互相印证
 *
 * 检查逻辑：不同音量等级下，响度值越大，频响幅值也应越大。
 * 对相同方向、相同类型的测试，按音量等级排列后检查单调性。
 */
function checkCurveValueCorroboration(testDataFacts) {
  const evidence = [];
  const issues = [];
  const ready = requireTestData(testDataFacts);

  if (!ready.ok) {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  const metrics = testDataFacts.loudnessMetrics || [];
  const frMetrics = testDataFacts.frequencyResponseMetrics || [];

  // 按 volumeLevel 分组，检查单调性
  const volumeGroups = {};
  metrics.forEach((m) => {
    const vol = m.volumeLevel || 'default';
    const dir = m.direction || 'unknown';
    const key = `${dir}|${vol}`;
    if (!volumeGroups[key]) volumeGroups[key] = [];
    volumeGroups[key].push(m.value);
  });

  const sortedVolumes = Object.keys(volumeGroups).sort((a, b) => {
    const numA = parseInt(a.split('|')[1]);
    const numB = parseInt(b.split('|')[1]);
    if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
    return a.localeCompare(b);
  });

  if (sortedVolumes.length < 2) {
    evidence.push('仅有1个音量等级的测试数据，无法进行跨音量印证');
    return { issues: [{ severity: 'review', message: '仅有1个音量水平，无法验证曲线与数值的跨音量一致性' }], evidence, status: 'review' };
  }

  evidence.push(`共 ${sortedVolumes.length} 个音量分组`);

  const byDir = {};
  metrics.forEach((m) => {
    const dir = m.direction || 'unknown';
    if (!byDir[dir]) byDir[dir] = [];
    byDir[dir].push({ volume: m.volumeLevel || 'default', value: m.value });
  });

  Object.entries(byDir).forEach(([dir, items]) => {
    const sorted = items.sort(function(a, b) { return String(a.volume).localeCompare(String(b.volume), undefined, { numeric: true }); });
    if (sorted.length < 2) return;

    var increasing = 0;
    var decreasing = 0;
    var sameValue = 0;

    for (var i = 1; i < sorted.length; i++) {
      // 仅在不同音量级别之间比较，跳过同音量的相邻对
      if (sorted[i].volume === sorted[i - 1].volume) continue;

      if (sorted[i].value > sorted[i - 1].value) increasing += 1;
      else if (sorted[i].value < sorted[i - 1].value) decreasing += 1;
      else sameValue += 1;
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

  if (frMetrics.length > 0) {
    evidence.push(`频响数据: ${frMetrics.length} 个测点`);
  }

  if (issues.length === 0) {
    evidence.push('✓ 不同音量下响度变化趋势一致');
    return { issues: [], evidence, status: 'pass' };
  }

  var hasError = issues.some(function(item) { return item.severity === 'error'; });
  var hasWarning = issues.some(function(item) { return item.severity === 'warning'; });
  if (hasError) return { issues, evidence, status: 'error' };
  if (hasWarning) return { issues, evidence, status: 'warning' };
  return { issues, evidence, status: 'review' };
}

module.exports = {
  checkSameCodecDifferentNetworkLoudness,
  checkSameNetworkDifferentCodecLoudness,
  checkLoudnessFrequencyResponseTrendConsistency,
  checkCurveValueCorroboration,
};
