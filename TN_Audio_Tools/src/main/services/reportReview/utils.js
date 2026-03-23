function normalizeText(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function normalizeUpperText(value) {
  return normalizeText(value).toUpperCase();
}

function buildWordData(reportData) {
  return {
    paragraphs: reportData.rawLines || reportData.lines || [],
    headers: reportData.structuredData?.headers || [],
    footers: reportData.structuredData?.footers || [],
    pageCount: reportData.pageCount || null
  };
}

function updateSummary(summary, status) {
  if (status === 'pass') {
    summary.passedChecks += 1;
  } else if (status === 'warning') {
    summary.warningChecks += 1;
  } else if (status === 'review') {
    summary.reviewChecks += 1;
  } else if (status === 'error') {
    summary.errorChecks += 1;
  }
}

function determineOverallStatus(summary) {
  if (summary.errorChecks > 0) return 'error';
  if (summary.warningChecks > 0) return 'warning';
  if (summary.reviewChecks > 0) return 'review';
  return 'pass';
}

module.exports = {
  normalizeText,
  normalizeUpperText,
  buildWordData,
  updateSummary,
  determineOverallStatus
};