/**
 * 快速测试：对 Utah_All 下的 Word 报告运行报告审查
 * 用法: node scripts/diagnostics/test_review.js
 */
const path = require('path');
const { reviewWordReport } = require('../../src/main/services/reportReview');

const UTAH_DIR = path.resolve(__dirname, '../../../../参考文件/Utah_All');
const TEST_FILES = [
  path.join(UTAH_DIR, 'HA/utah_HA_NB.doc'),
  path.join(UTAH_DIR, 'HA/utah_HA_WB.doc'),
  path.join(UTAH_DIR, 'HA/utah_HA_SWB.doc'),
];

async function main() {
  for (const filePath of TEST_FILES) {
    console.log(`\n${'='.repeat(70)}`);
    console.log(`Testing: ${path.basename(filePath)}`);
    console.log(`${'='.repeat(70)}`);

    try {
      const result = await reviewWordReport(filePath);
      const { summary, checks, overallStatus } = result.reviewResult;

      console.log(`Overall: ${overallStatus}`);
      console.log(`Summary: passed=${summary.passedChecks}, warning=${summary.warningChecks}, review=${summary.reviewChecks}, error=${summary.errorChecks}`);

      for (const [key, checkResult] of Object.entries(checks)) {
        const icon = checkResult.status === 'pass' ? '✓' : checkResult.status === 'warning' ? '⚠' : checkResult.status === 'error' ? '✗' : '?';
        console.log(`  ${icon} ${key}: ${checkResult.status}`);
        if (checkResult.issues && checkResult.issues.length > 0) {
          checkResult.issues.forEach((issue) => {
            console.log(`    → [${issue.severity}] ${issue.message}`);
          });
        }
      }
    } catch (err) {
      console.error(`FAILED: ${err.message}`);
      console.error(err.stack);
    }
  }
}

main().catch(console.error);
