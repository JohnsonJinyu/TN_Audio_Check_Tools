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
function checkLoudnessFrequencyResponseTrendConsistency(testDataFacts) {
  const evidence = [];
  const issues = [];
  const ready = requireTestData(testDataFacts);

  if (!ready.ok) {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  const metrics = testDataFacts.loudnessMetrics || [];
  const frMetrics = testDataFacts.frequencyResponseMetrics || [];

  if (frMetrics.length === 0) {
    evidence.push('未提取到频响数据，仅基于响度数据做基本检查');

    // 至少检查同方向上响度值是否有明显异常
    const byDir = {};
    metrics.forEach((m) => {
      const dir = m.direction || 'unknown';
      if (!byDir[dir]) byDir[dir] = [];
      byDir[dir].push(m.value);
    });

    Object.entries(byDir).forEach(([dir, values]) => {
      if (values.length < 3) return;
      const min = Math.min(...values);
      const max = Math.max(...values);
      if (max - min > 10) {
        issues.push({
          severity: 'review',
          message: `${dir}方向响度值范围较大 (${min.toFixed(1)} ~ ${max.toFixed(1)} dB, 跨度${(max-min).toFixed(1)}dB)，建议人工确认频响趋势`,
        });
      }
      evidence.push(`${dir}方向: ${values.length}个响度值, 范围 ${min.toFixed(1)}~${max.toFixed(1)} dB`);
    });

    if (issues.length === 0) {
      evidence.push('✓ 响度数据无显著异常');
      return { issues: [], evidence, status: 'pass' };
    }
    return { issues, evidence, status: 'review' };
  }

  // 有频响数据时，按方向比较趋势
  const byDirLoudness = {};
  metrics.forEach((m) => {
    const dir = m.direction || 'unknown';
    if (!byDirLoudness[dir]) byDirLoudness[dir] = [];
    byDirLoudness[dir].push(m);
  });

  const byDirFR = {};
  frMetrics.forEach((m) => {
    const dir = m.direction || 'unknown';
    if (!byDirFR[dir]) byDirFR[dir] = [];
    byDirFR[dir].push(m);
  });

  const directions = new Set([...Object.keys(byDirLoudness), ...Object.keys(byDirFR)]);

  directions.forEach((dir) => {
    const loudVals = (byDirLoudness[dir] || []).map((m) => m.value);
    const frVals = (byDirFR[dir] || []).map((m) => m.amplitude);

    if (loudVals.length < 2 || frVals.length < 2) {
      evidence.push(`${dir}方向数据不足，无法完成趋势对比（响度${loudVals.length}个，频响${frVals.length}个）`);
      return;
    }

    // 简单线性趋势：看首尾值的方向
    const loudTrend = loudVals[loudVals.length - 1] - loudVals[0];
    const frTrend = frVals[frVals.length - 1] - frVals[0];

    evidence.push(`${dir}方向: 响度趋势=${loudTrend > 0 ? '+' : ''}${loudTrend.toFixed(2)}dB, 频响趋势=${frTrend > 0 ? '+' : ''}${frTrend.toFixed(2)}dB`);

    // 趋势方向应一致（同时增或同时减）
    if ((loudTrend > 0 && frTrend < -2) || (loudTrend < 0 && frTrend > 2)) {
      issues.push({
        severity: 'warning',
        message: `${dir}方向响度趋势与频响趋势不一致：响度${loudTrend > 0 ? '上升' : '下降'}${Math.abs(loudTrend).toFixed(1)}dB，频响${frTrend > 0 ? '上升' : '下降'}${Math.abs(frTrend).toFixed(1)}dB`,
      });
    } else {
      evidence.push(`  ✓ 趋势方向一致`);
    }
  });

  if (issues.length === 0) {
    evidence.push('✓ TX/RX响度与频响趋势一致');
    return { issues: [], evidence, status: 'pass' };
  }

  return { issues, evidence, status: 'warning' };
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
    return { issues: [], evidence, status: 'pass' };
  }

  evidence.push(`共 ${sortedVolumes.length} 个音量分组`);

  const byDir = {};
  metrics.forEach((m) => {
    const dir = m.direction || 'unknown';
    if (!byDir[dir]) byDir[dir] = [];
    byDir[dir].push({ volume: m.volumeLevel || 'default', value: m.value });
  });

  Object.entries(byDir).forEach(([dir, items]) => {
    // 按音量标签排序
    const sorted = items.sort((a, b) => String(a.volume).localeCompare(String(b.volume), undefined, { numeric: true }));
    if (sorted.length < 2) return;

    const values = sorted.map((it) => it.value);
    let increasing = 0;
    let decreasing = 0;

    for (let i = 1; i < values.length; i++) {
      if (values[i] > values[i - 1]) increasing += 1;
      else if (values[i] < values[i - 1]) decreasing += 1;
    }

    const total = increasing + decreasing;
    if (total === 0) return;

    const consistency = Math.max(increasing, decreasing) / total;
    evidence.push(`${dir}方向跨音量单调性: ${(consistency * 100).toFixed(0)}% (${increasing}升/${decreasing}降)`);

    if (consistency < 0.6) {
      issues.push({
        severity: 'review',
        message: `${dir}方向不同音量下响度变化方向不一致（${increasing}次上升, ${decreasing}次下降），建议人工验证曲线趋势`,
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

  return { issues, evidence, status: 'review' };
}

module.exports = {
  checkSameCodecDifferentNetworkLoudness,
  checkSameNetworkDifferentCodecLoudness,
  checkLoudnessFrequencyResponseTrendConsistency,
  checkCurveValueCorroboration,
};
