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