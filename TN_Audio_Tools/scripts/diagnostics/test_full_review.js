/**
 * 全面回归测试：Word + xlsx 报告审查
 */
const path = require('path');
const { reviewWordReport, runCrossReportChecks } = require('../../src/main/services/reportReview');

const UTAH_DIR = path.resolve(__dirname, '../../../../参考文件/Utah_All');

async function testSingle(filePath, label) {
  console.log(`\n${'─'.repeat(60)}`);
  console.log(`[${label}] ${path.basename(filePath)}`);
  console.log(`${'─'.repeat(60)}`);

  try {
    const result = await reviewWordReport(filePath);
    const { summary, checks, overallStatus } = result.reviewResult;

    console.log(`  Overall: ${overallStatus}`);
    console.log(`  Summary: total=${summary.totalChecks} passed=${summary.passedChecks} warning=${summary.warningChecks} review=${summary.reviewChecks} error=${summary.errorChecks}`);

    // Show non-pass checks
    const nonPass = Object.entries(checks).filter(([, v]) => v && v.status !== 'pass');
    if (nonPass.length > 0) {
      console.log('  Non-pass checks:');
      nonPass.forEach(([key, result]) => {
        console.log(`    ${key}: ${result.status}`);
        (result.issues || []).slice(0, 2).forEach((i) => {
          console.log(`      [${i.severity}] ${i.message}`);
        });
      });
    } else {
      console.log('  All checks passed!');
    }

    return result;
  } catch (err) {
    console.error(`  FAILED: ${err.message}`);
    return null;
  }
}

async function main() {
  // Test 1: Word report (backward compatibility)
  console.log('=== Word报告向后兼容测试 ===');
  const wordResult = await testSingle(
    path.join(UTAH_DIR, 'HA/utah_HA_NB.doc'),
    'Word'
  );

  // Test 2: xlsx report
  console.log('\n=== xlsx报告审查测试 ===');
  const xlsxResult = await testSingle(
    path.join(UTAH_DIR, 'HA/utah_HA_NB.xlsx'),
    'xlsx'
  );

  // Test 3: Batch cross-report
  console.log('\n=== 跨报告批量对比测试 ===');
  const xlsxResults = [];
  for (const f of ['utah_HA_NB.xlsx', 'utah_HA_WB.xlsx', 'utah_HA_SWB.xlsx']) {
    const r = await testSingle(path.join(UTAH_DIR, 'HA', f), 'xlsx');
    if (r) xlsxResults.push(r);
  }

  if (xlsxResults.length >= 2) {
    console.log('\n--- Running cross-report checks ---');
    const crossResults = runCrossReportChecks(xlsxResults);
    crossResults.forEach((r) => {
      const codecCheck = r.reviewResult.checks.contentSameCodecDiffNetwork;
      const networkCheck = r.reviewResult.checks.contentSameNetworkDiffCodec;
      console.log(`  ${path.basename(r.reportPath)}:`);
      console.log(`    contentSameCodecDiffNetwork: ${codecCheck?.status}`);
      console.log(`    contentSameNetworkDiffCodec: ${networkCheck?.status}`);
      (codecCheck?.issues || []).slice(0, 2).forEach((i) => console.log(`      [${i.severity}] ${i.message}`));
      (networkCheck?.issues || []).slice(0, 2).forEach((i) => console.log(`      [${i.severity}] ${i.message}`));
    });
  }

  console.log('\n=== 全部测试完成 ===');
}

main().catch(console.error);
