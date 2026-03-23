export const WORD_REVIEW_HISTORY_STORAGE_KEY = 'tn-audio-word-review-history';
export const WORD_REVIEW_HISTORY_UPDATED_EVENT = 'tn-audio-word-review-history-updated';

const WORD_REVIEW_HISTORY_LIMIT = 100;

function createDefaultWordReviewHistory() {
  return [];
}

function getSafeLocalStorage() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return null;
  }

  return window.localStorage;
}

function getFileNameFromPath(filePath) {
  if (!filePath) {
    return '';
  }

  const normalizedPath = String(filePath).replace(/\\/g, '/');
  const segments = normalizedPath.split('/').filter(Boolean);
  return segments[segments.length - 1] || filePath;
}

export function readWordReviewHistory() {
  const storage = getSafeLocalStorage();
  if (!storage) {
    return createDefaultWordReviewHistory();
  }

  try {
    const rawValue = storage.getItem(WORD_REVIEW_HISTORY_STORAGE_KEY);
    if (!rawValue) {
      return createDefaultWordReviewHistory();
    }

    const parsed = JSON.parse(rawValue);
    return Array.isArray(parsed) ? parsed : createDefaultWordReviewHistory();
  } catch (error) {
    return createDefaultWordReviewHistory();
  }
}

export function writeWordReviewHistory(historyList) {
  const storage = getSafeLocalStorage();
  if (!storage) {
    return;
  }

  try {
    const safeHistoryList = Array.isArray(historyList) ? historyList : [];
    storage.setItem(
      WORD_REVIEW_HISTORY_STORAGE_KEY,
      JSON.stringify(safeHistoryList.slice(0, WORD_REVIEW_HISTORY_LIMIT))
    );

    if (typeof window !== 'undefined') {
      window.dispatchEvent(new CustomEvent(WORD_REVIEW_HISTORY_UPDATED_EVENT));
    }
  } catch (error) {
    // Ignore local persistence failures.
  }
}

export function recordWordReviewResult(reportPath, reviewResult) {
  if (!reportPath || !reviewResult) {
    return;
  }

  const currentHistory = readWordReviewHistory();
  const entry = {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 11)}`,
    reportName: getFileNameFromPath(reportPath),
    reportPath,
    status: reviewResult.reviewResult?.overallStatus || 'unknown',
    checkedAt: new Date().toISOString(),
    result: reviewResult
  };

  writeWordReviewHistory([entry, ...currentHistory]);
}

export function clearWordReviewHistory() {
  writeWordReviewHistory([]);
}
