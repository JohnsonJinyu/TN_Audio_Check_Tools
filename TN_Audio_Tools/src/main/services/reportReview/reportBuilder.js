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
      description: '除3quest测试外，相邻测试项的间隔时间应≤1分钟'
    },
    timingTotalSpan: {
      title: '全部测试项跨度检查',
      description: '全部测试项应在6小时内完成'
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
      description: '不同音量级别下，响度数值变化方向应与频响幅值变化方向一致'
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
        conclusion: checkResult.conclusion || '',
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