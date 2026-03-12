export const DASHBOARD_STATS_STORAGE_KEY = 'tn-audio-dashboard-stats';

const REPORT_HISTORY_LIMIT = 100;

function createDefaultDashboardData() {
  return {
    processedAudio: 0,
    checkedReports: 0,
    analysisSuccess: 0,
    conversionsCompleted: 0,
    reportHistory: []
  };
}

function getSafeLocalStorage() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return null;
  }

  return window.localStorage;
}

function normalizeDashboardData(rawValue) {
  return {
    processedAudio: Number(rawValue?.processedAudio) || 0,
    checkedReports: Number(rawValue?.checkedReports) || 0,
    analysisSuccess: Number(rawValue?.analysisSuccess) || 0,
    conversionsCompleted: Number(rawValue?.conversionsCompleted) || 0,
    reportHistory: Array.isArray(rawValue?.reportHistory) ? rawValue.reportHistory : []
  };
}

export function readDashboardData() {
  const storage = getSafeLocalStorage();
  if (!storage) {
    return createDefaultDashboardData();
  }

  try {
    const rawValue = storage.getItem(DASHBOARD_STATS_STORAGE_KEY);
    const parsedValue = rawValue ? JSON.parse(rawValue) : createDefaultDashboardData();
    return normalizeDashboardData(parsedValue);
  } catch (error) {
    return createDefaultDashboardData();
  }
}

export function writeDashboardData(nextValue) {
  const storage = getSafeLocalStorage();
  if (!storage) {
    return;
  }

  try {
    storage.setItem(DASHBOARD_STATS_STORAGE_KEY, JSON.stringify(normalizeDashboardData(nextValue)));
  } catch (error) {
    // Ignore local persistence failures.
  }
}

function getFileNameFromPath(filePath) {
  if (!filePath) {
    return '';
  }

  const normalizedPath = String(filePath).replace(/\\/g, '/');
  const segments = normalizedPath.split('/').filter(Boolean);
  return segments[segments.length - 1] || filePath;
}

export function recordReportCheckResults(resultList) {
  if (!Array.isArray(resultList) || resultList.length === 0) {
    return;
  }

  const currentValue = readDashboardData();
  const successCount = resultList.filter((item) => item.status === 'success').length;
  const checkedAt = new Date().toISOString();

  const historyEntries = resultList.map((item, index) => ({
    id: `${checkedAt}-${index}`,
    reportName: getFileNameFromPath(item.reportPath),
    outputName: getFileNameFromPath(item.outputPath),
    outputPath: item.outputPath || '',
    status: item.status || 'unknown',
    matchedItems: Number(item.matchedItems) || 0,
    error: item.error || '',
    checkedAt
  }));

  writeDashboardData({
    ...currentValue,
    checkedReports: currentValue.checkedReports + resultList.length,
    analysisSuccess: currentValue.analysisSuccess + successCount,
    reportHistory: [...historyEntries.reverse(), ...currentValue.reportHistory].slice(0, REPORT_HISTORY_LIMIT)
  });
}