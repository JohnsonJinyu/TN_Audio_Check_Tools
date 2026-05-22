const { formatLocalDateTime } = require('../reportTestDataFacts');

/**
 * 测试时间检查 (2.3)
 *
 * 基于ACQUA xlsx Detailed sheet的Date/Time列，验证测试执行的时序合理性。
 * 对Word报告（无testDataFacts）所有检查返回review状态并提示需要ACQUA数据。
 */

function requireTimestamps(testDataFacts) {
  if (!testDataFacts || !testDataFacts.testItemTimestamps || testDataFacts.testItemTimestamps.length === 0) {
    return { ok: false, reason: '非ACQUA xlsx格式报告，无法进行时序检查。请提供ACQUA导出的.xlsx测试报告。' };
  }

  if (!testDataFacts.hasAbsoluteTimestamps) {
    return {
      ok: 'partial',
      reason: 'ACQUA数据中未找到Date/Time时间戳，部分时序检查将以行序为参考（准确性降低）。',
      items: testDataFacts.testItemTimestamps,
    };
  }

  const sorted = [...testDataFacts.testItemTimestamps]
    .filter((t) => t.timestamp)
    .sort((a, b) => a.timestamp - b.timestamp);

  return { ok: true, sorted, reason: '' };
}

/**
 * 2.3.1 相邻测试项间隔检查
 * 除3quest测试外，相邻测试项间隔应≤5分钟（300秒）
 */
function checkAdjacentTestItemInterval(testDataFacts) {
  const evidence = [];
  const issues = [];
  const ready = requireTimestamps(testDataFacts);

  if (ready.ok === false) {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  if (ready.ok === 'partial') {
    evidence.push(ready.reason);
  }

  const items = ready.ok === true ? ready.sorted : (ready.items || []);
  const non3questItems = items.filter((t) => t.testCategory !== '3quest');

  let violations = 0;
  const maxViolations = 10;

  for (let i = 1; i < items.length && violations < maxViolations; i++) {
    if (items[i].testCategory === '3quest' || items[i - 1].testCategory === '3quest') continue;
    if (!items[i].timestamp || !items[i - 1].timestamp) continue;

    const deltaSec = (items[i].timestamp - items[i - 1].timestamp) / 1000;
    if (deltaSec > 300) {
      violations += 1;
      const severity = deltaSec > 600 ? 'warning' : 'review';
      issues.push({
        severity,
        message: `测试项 "${items[i - 1].descriptor}" 与 "${items[i].descriptor}" 之间间隔 ${deltaSec.toFixed(0)} 秒（>5分钟）`,
      });
    }
  }

  if (violations > maxViolations) {
    evidence.push(`超过 ${maxViolations} 个间隔异常（共检查 ${items.length} 个测试项），仅展示前 ${maxViolations} 个`);
  }

  if (violations === 0) {
    evidence.push('✓ 相邻测试项间隔均在5分钟以内');
    return { issues: [], evidence, status: 'pass' };
  }

  evidence.push(`共发现 ${violations} 个间隔超过5分钟的相邻测试项对`);
  return { issues, evidence, status: violations > 5 ? 'warning' : 'review' };
}

/**
 * 2.3.2 全部测试项时间跨度检查
 * 一份报告内所有测试项的时间差应在6小时以内
 */
function checkTotalTestSpan(testDataFacts) {
  const evidence = [];
  const issues = [];
  const ready = requireTimestamps(testDataFacts);

  if (ready.ok === false) {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  if (ready.ok === 'partial') {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  // 使用原始 sheet 行序的首项/末项时间戳（非排序后的 min/max）
  const itemsInSheetOrder = (testDataFacts?.testItemTimestamps || [])
    .filter((t) => t.timestamp);
  if (itemsInSheetOrder.length < 2) {
    evidence.push('时间戳不足，无法计算测试总时长');
    return { issues: [{ severity: 'review', message: '时间戳数量不足' }], evidence, status: 'review' };
  }

  const firstTs = itemsInSheetOrder[0].timestamp;
  const lastTs = itemsInSheetOrder[itemsInSheetOrder.length - 1].timestamp;
  const totalHours = (lastTs - firstTs) / 3600000;

  evidence.push(`首个测试时间: ${itemsInSheetOrder[0].descriptor}`);
  evidence.push(`  └ ${formatLocalDateTime(firstTs)}`);
  evidence.push(`末个测试时间: ${itemsInSheetOrder[itemsInSheetOrder.length - 1].descriptor}`);
  evidence.push(`  └ ${formatLocalDateTime(lastTs)}`);
  evidence.push(`总时长: ${totalHours.toFixed(1)} 小时`);

  if (totalHours > 8) {
    issues.push({
      severity: 'error',
      message: `全部测试项跨度为 ${totalHours.toFixed(1)} 小时，超过6小时要求，且已达到严重异常阈值（>8小时）`,
    });
    return { issues, evidence, status: 'error' };
  }

  if (totalHours > 6) {
    issues.push({
      severity: 'warning',
      message: `全部测试项跨度为 ${totalHours.toFixed(1)} 小时，超过6小时上限`,
    });
    return { issues, evidence, status: 'warning' };
  }

  evidence.push('✓ 全部测试项在6小时内完成');
  return { issues: [], evidence, status: 'pass' };
}

/**
 * 2.3.3 时延测试时序检查
 * Delay测试应排在所有其他测试的最前面
 */
function checkDelayTestTiming(testDataFacts) {
  const evidence = [];
  const issues = [];
  const ready = requireTimestamps(testDataFacts);

  if (ready.ok === false) {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  if (ready.ok === 'partial') {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  const delayItems = ready.sorted.filter((t) => t.testCategory === 'delay' || t.testCategory === 'echo_delay');
  const nonDelayItems = ready.sorted.filter(
    (t) => t.testCategory !== 'delay' && t.testCategory !== 'echo_delay'
  );

  if (delayItems.length === 0) {
    evidence.push('未找到时延类测试项');
    return { issues: [], evidence, status: 'pass' };
  }

  evidence.push(`找到 ${delayItems.length} 个时延类测试项`);

  const latestDelay = delayItems[delayItems.length - 1].timestamp;
  const earliestNonDelay = nonDelayItems[0]?.timestamp;

  if (earliestNonDelay && latestDelay > earliestNonDelay) {
    const violatingNonDelay = nonDelayItems.filter((t) => t.timestamp < latestDelay);
    issues.push({
      severity: 'warning',
      message: `时延测试未完全排在所有测试之前：最新时延测试在 ${formatLocalDateTime(latestDelay)}，但有 ${violatingNonDelay.length} 个非时延测试在其之前执行`,
    });
    evidence.push(
      `违规示例: ${violatingNonDelay.slice(0, 3).map((t) => t.descriptor).join('; ')}`
    );
    return { issues, evidence, status: 'warning' };
  }

  evidence.push('✓ 时延测试排在所有测试之前');
  return { issues: [], evidence, status: 'pass' };
}

/**
 * 2.3.4 Sidetone Delay时序检查
 * Sidetone Delay测试应在Sidetone测试之前
 */
function checkSidetoneDelayTiming(testDataFacts) {
  const evidence = [];
  const issues = [];
  const ready = requireTimestamps(testDataFacts);

  if (ready.ok === false) {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  if (ready.ok === 'partial') {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  const sdItems = ready.sorted.filter((t) => t.testCategory === 'sidetone_delay');
  const stItems = ready.sorted.filter((t) => t.testCategory === 'sidetone');

  if (sdItems.length === 0) {
    evidence.push('未找到 Sidetone Delay 测试项，跳过检查');
    return { issues: [], evidence, status: 'pass' };
  }

  if (stItems.length === 0) {
    evidence.push('未找到 Sidetone 测试项，跳过检查');
    return { issues: [], evidence, status: 'pass' };
  }

  evidence.push(`找到 ${sdItems.length} 个Sidetone Delay项，${stItems.length} 个Sidetone项`);

  const latestSD = sdItems[sdItems.length - 1].timestamp;
  const earliestST = stItems[0].timestamp;

  if (latestSD > earliestST) {
    const violatingST = stItems.filter((t) => t.timestamp < latestSD);
    issues.push({
      severity: 'warning',
      message: `Sidetone Delay未在Sidetone之前执行：有 ${violatingST.length} 个Sidetone测试在Sidetone Delay之前`,
    });
    return { issues, evidence, status: 'warning' };
  }

  evidence.push('✓ Sidetone Delay在Sidetone之前执行');
  return { issues: [], evidence, status: 'pass' };
}

/**
 * 2.3.5 BGN Connection时序检查
 * BGN Connection（3quest连接测试）应在3quest测试之前
 */
function checkBgnConnectionTiming(testDataFacts) {
  const evidence = [];
  const issues = [];
  const ready = requireTimestamps(testDataFacts);

  if (ready.ok === false) {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  if (ready.ok === 'partial') {
    evidence.push(ready.reason);
    return { issues: [{ severity: 'review', message: ready.reason }], evidence, status: 'review' };
  }

  const bgnConnItems = ready.sorted.filter((t) => t.testCategory === 'bgn_connection');
  const threeQuestItems = ready.sorted.filter((t) => t.testCategory === '3quest');

  if (bgnConnItems.length === 0) {
    evidence.push('未找到 BGN Connection 测试项，跳过检查');
    return { issues: [], evidence, status: 'pass' };
  }

  if (threeQuestItems.length === 0) {
    evidence.push('未找到 3quest 测试项，跳过检查');
    return { issues: [], evidence, status: 'pass' };
  }

  evidence.push(`找到 ${bgnConnItems.length} 个BGN Connection项，${threeQuestItems.length} 个3quest项`);

  const latestBgn = bgnConnItems[bgnConnItems.length - 1].timestamp;
  const earliest3q = threeQuestItems[0].timestamp;

  if (latestBgn > earliest3q) {
    const violating3q = threeQuestItems.filter((t) => t.timestamp < latestBgn);
    issues.push({
      severity: 'warning',
      message: `BGN Connection未在3quest之前执行：有 ${violating3q.length} 个3quest测试在BGN Connection之前`,
    });
    return { issues, evidence, status: 'warning' };
  }

  evidence.push('✓ BGN Connection在3quest之前执行');
  return { issues: [], evidence, status: 'pass' };
}

module.exports = {
  checkAdjacentTestItemInterval,
  checkTotalTestSpan,
  checkDelayTestTiming,
  checkSidetoneDelayTiming,
  checkBgnConnectionTiming,
};
