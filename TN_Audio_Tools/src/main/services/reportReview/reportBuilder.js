function normalizeConclusionText(text) {
  return String(text || '').replace(/^✓\s*/, '').trim();
}

function buildSectionConclusion(checkResult) {
  if (!checkResult) return '';

  // 检查函数显式提供的结论优先
  var explicitConclusion = normalizeConclusionText(checkResult.conclusion);
  if (explicitConclusion) return explicitConclusion;

  // pass 状态没有异常，不展示结论（避免与"诊断依据"重复）
  if (checkResult.status === 'pass') return '';

  // 非 pass：优先取 issue message 作为摘要
  var issues = Array.isArray(checkResult.issues) ? checkResult.issues : [];
  var evidence = Array.isArray(checkResult.evidence) ? checkResult.evidence : [];
  var firstIssue = issues.find(function(issue) {
    return issue && typeof issue.message === 'string' && issue.message.trim();
  });
  if (firstIssue) return normalizeConclusionText(firstIssue.message);

  // 没有 issue 但有 evidence（如 engineers 只返回 evidence 不返回 issues）
  for (var i = evidence.length - 1; i >= 0; i -= 1) {
    var text = normalizeConclusionText(evidence[i]);
    if (text) return text;
  }

  if (checkResult.status === 'warning') return '发现需关注的问题';
  if (checkResult.status === 'error') return '发现明确异常';
  return '需要人工复核';
}

function generateReviewReport(reviewResult) {
  const report = {
    title: 'Word 报告审查结果',
    timestamp: reviewResult.reviewTimestamp,
    overallStatus: reviewResult.overallStatus,
    summary: reviewResult.summary,
    sections: []
  };

  const sectionMap = {
    tableOfContents: {
      title: '目录提取',
      description: '已列出文档中识别到的章节结构'
    },
    tableOfContentsPages: {
      title: '目录页数检查',
      description: '验证目录记载的页数与文档总页数是否一致'
    },
    chaptersAlignment: {
      title: '章节与目录对应',
      description: '检查文档内容章节是否与目录结构对应'
    },
    basicInfo: {
      title: '报告基本信息',
      description: '检查报告名称、页眉、Measurement Object 等关键信息一致性'
    },
    testItemConsistency: {
      title: '测试项一致性',
      description: '检查 codec、bandwidth、terminal mode 等信息是否与文件名及正文一致'
    },
    namePollution: {
      title: '名称污染检查',
      description: '检查报告名称中是否混入日期、版本等污染信息'
    },
    engineers: {
      title: '人员信息',
      description: '查找报告中的测试人员、工程师名字'
    },
    polqa: {
      title: 'POLQA 配置检查',
      description: '检查 POLQA 算法版本和参考音源配置是否完整'
    },
    timingAdjacentInterval: {
      title: '相邻测试项间隔检查',
      description: '除3quest测试外，相邻测试项的间隔时间应≤5分钟'
    },
    timingTotalSpan: {
      title: '全部测试项跨度检查',
      description: '全部测试项应在6小时内完成；超过8小时视为严重异常'
    },
    timingDelayOrder: {
      title: '时延测试时序检查',
      description: '时延测试应在测试序列的最前面执行'
    },
    timingSidetoneDelayOrder: {
      title: 'Sidetone Delay时序检查',
      description: 'Sidetone Delay测试应在Sidetone测试之前完成'
    },
    timingBgnConnectionOrder: {
      title: 'BGN Connection时序检查',
      description: 'BGN Connection测试应在3quest测试之前完成'
    },
    contentLoudnessFRTrend: {
      title: '响度与频响趋势一致性',
      description: '对比响度曲线图与频响曲线图的视觉趋势是否一致（需AI视觉分析，当前需人工复核）'
    },
    contentCurveValueCorroboration: {
      title: '曲线与数值互相印证',
      description: '不同音量级别下，Loudness Rating (RLR/SLR) 数值应保持合理单调；对 rating 而言数值越小通常表示实际越响'
    },
    contentSameCodecDiffNetwork: {
      title: '同Codec不同网络响度差异（跨报告）',
      description: '相同codec类型在不同网络下的RX/TX响度差异应≤1dB'
    },
    contentSameNetworkDiffCodec: {
      title: '同网络不同Codec响度差异（跨报告）',
      description: '相同网络在不同codec类型下的RX/TX响度差异应≤1dB'
    }
  };

  Object.entries(reviewResult.checks).forEach(([key, checkResult]) => {
    const sectionDef = sectionMap[key];
    if (sectionDef) {
      report.sections.push({
        key,
        title: sectionDef.title,
        description: sectionDef.description,
        status: checkResult.status,
        issues: checkResult.issues || [],
        evidence: checkResult.evidence || [],
        logs: checkResult.logs || [],
        checklist: checkResult.checklist || [],
        chartData: checkResult.chartData || {},
        comparisonCards: checkResult.comparisonCards || [],
        conclusion: buildSectionConclusion(checkResult),
        rawFindings: checkResult.rawFindings || [],
        data: (() => {
          if (key === 'tableOfContents') return { chapters: checkResult.chapters };
          if (key === 'engineers') return { engineers: checkResult.engineers };
          return null;
        })()
      });
    }
  });

  return report;
}

module.exports = {
  generateReviewReport
};