require('../src/main/services/reportChecker/runtimePolyfills');

const fs = require('fs');
const os = require('os');
const path = require('path');
const { analyzeCorpus } = require('./analyze-3gpp-corpus');

async function main() {
  const summary = await analyzeCorpus();
  const summaryPath = path.join(os.tmpdir(), 'tn-audio-3gpp-summary.json');
  fs.writeFileSync(summaryPath, JSON.stringify(summary, null, 2), 'utf8');

  const compact = {
    summaryPath,
    byReport: summary.byReport.map((item) => ({
      reportName: item.reportName,
      diffCount: item.diffCount,
      diffItemIds: item.diffItemIds
    })),
    topItems: summary.itemSummary.slice(0, 15).map((item) => ({
      itemId: item.itemId,
      reports: item.reports,
      checklistDesc: item.checklistDesc
    }))
  };

  console.log(JSON.stringify(compact, null, 2));
}

main().catch((error) => {
  const errorPath = path.join(os.tmpdir(), 'tn-audio-3gpp-error.txt');
  fs.writeFileSync(errorPath, String(error.stack || error), 'utf8');
  console.error(error.stack || error);
  process.exit(1);
});