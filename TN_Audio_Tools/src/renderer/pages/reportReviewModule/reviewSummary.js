export const reviewStatusText = {
  pass: '通过',
  warning: '有警告',
  review: '需复核',
  error: '有错误'
};

export const reviewStatusColor = {
  pass: 'green',
  warning: 'orange',
  review: 'blue',
  error: 'red'
};

export const reviewStatusTheme = {
  pass: {
    accent: 'var(--status-pass)',
    soft: 'color-mix(in srgb, var(--status-pass) 8%, transparent)',
    border: 'color-mix(in srgb, var(--status-pass) 40%, var(--border-color))',
    title: 'var(--status-pass)',
    muted: 'var(--status-pass)'
  },
  warning: {
    accent: 'var(--status-warn)',
    soft: 'color-mix(in srgb, var(--status-warn) 8%, transparent)',
    border: 'color-mix(in srgb, var(--status-warn) 40%, var(--border-color))',
    title: 'var(--status-warn)',
    muted: 'var(--status-warn)'
  },
  review: {
    accent: 'var(--status-info)',
    soft: 'color-mix(in srgb, var(--status-info) 8%, transparent)',
    border: 'color-mix(in srgb, var(--status-info) 40%, var(--border-color))',
    title: 'var(--status-info)',
    muted: 'var(--status-info)'
  },
  error: {
    accent: 'var(--status-error)',
    soft: 'color-mix(in srgb, var(--status-error) 8%, transparent)',
    border: 'color-mix(in srgb, var(--status-error) 40%, var(--border-color))',
    title: 'var(--status-error)',
    muted: 'var(--status-error)'
  }
};

const HIDDEN_SECTION_KEYS = new Set(['engineers']);

function getVisibleSections(resultData) {
  const sections = Array.isArray(resultData?.report?.sections) ? resultData.report.sections : [];
  return sections.filter((section) => section && !HIDDEN_SECTION_KEYS.has(section.key));
}

function getSafeSummary(reviewResult) {
  const summary = reviewResult?.summary || {};

  return {
    passedChecks: Number(summary.passedChecks) || 0,
    warningChecks: Number(summary.warningChecks) || 0,
    reviewChecks: Number(summary.reviewChecks) || 0,
    errorChecks: Number(summary.errorChecks) || 0
  };
}

function collectIssueMessages(resultData, limit = 2) {
  const sections = getVisibleSections(resultData);
  const messages = [];

  sections.forEach((section) => {
    if (!Array.isArray(section?.issues)) {
      return;
    }

    section.issues.forEach((issue) => {
      const message = String(issue?.message || '').trim();
      if (message) {
        messages.push(message);
      }
    });
  });

  return Array.from(new Set(messages)).slice(0, limit);
}

export function buildReviewDigest(resultData) {
  const reviewResult = resultData?.reviewResult || {};
  const overallStatus = reviewResult.overallStatus || 'review';
  const summary = getSafeSummary(reviewResult);
  const issueMessages = collectIssueMessages(resultData);

  let headline = '存在待确认结论，请结合详情继续复核';
  if (overallStatus === 'pass') {
    headline = summary.passedChecks > 0
      ? `初步可通过，${summary.passedChecks} 项检查已通过`
      : '初步可通过，未发现阻断项';
  } else if (overallStatus === 'warning') {
    headline = `存在 ${summary.warningChecks || 1} 项警告，建议复核后再确认`;
  } else if (overallStatus === 'review') {
    headline = `存在 ${summary.reviewChecks || 1} 项待人工复核，请优先查看详情`;
  } else if (overallStatus === 'error') {
    headline = `存在 ${summary.errorChecks || 1} 项错误，需修正后再复审`;
  }

  let detail = '';
  if (issueMessages.length > 0) {
    detail = issueMessages.join('；');
  } else if (overallStatus === 'pass') {
    detail = '当前未发现需要阻断处理的问题。';
  } else if (overallStatus === 'warning') {
    detail = '建议优先检查目录、章节定位和配置完整性。';
  } else if (overallStatus === 'review') {
    detail = '当前结论不足以直接放行，建议人工确认关键章节。';
  } else if (overallStatus === 'error') {
    detail = '当前报告存在明确错误，建议先修复再提交。';
  }

  return {
    overallStatus,
    statusText: reviewStatusText[overallStatus] || overallStatus || '-',
    statusColor: reviewStatusColor[overallStatus] || 'default',
    theme: reviewStatusTheme[overallStatus] || reviewStatusTheme.review,
    headline,
    detail,
    summary,
    statsText: `通过 ${summary.passedChecks} / 警告 ${summary.warningChecks} / 复核 ${summary.reviewChecks} / 错误 ${summary.errorChecks}`
  };
}

export function getReviewSectionsByStatus(resultData, status) {
  return getReviewSectionsByStatusMap(resultData)[status] || [];
}

export function getReviewSectionsByStatusMap(resultData) {
  const sections = getVisibleSections(resultData);

  return sections.reduce((groups, section) => {
    const status = section?.status || 'review';
    if (!groups[status]) {
      groups[status] = [];
    }

    groups[status].push(section);
    return groups;
  }, {
    pass: [],
    warning: [],
    review: [],
    error: []
  });
}