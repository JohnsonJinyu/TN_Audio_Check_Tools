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
    accent: '#389e0d',
    soft: '#f6ffed',
    border: '#b7eb8f',
    title: '#135200',
    muted: '#3f6600'
  },
  warning: {
    accent: '#d48806',
    soft: '#fffbe6',
    border: '#ffe58f',
    title: '#874d00',
    muted: '#ad6800'
  },
  review: {
    accent: '#1677ff',
    soft: '#f0f5ff',
    border: '#adc6ff',
    title: '#10239e',
    muted: '#1d39c4'
  },
  error: {
    accent: '#cf1322',
    soft: '#fff1f0',
    border: '#ffa39e',
    title: '#820014',
    muted: '#a8071a'
  }
};

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
  const sections = Array.isArray(resultData?.report?.sections) ? resultData.report.sections : [];
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
  const sections = Array.isArray(resultData?.report?.sections) ? resultData.report.sections : [];
  return sections.filter((section) => section?.status === status);
}