const { normalizeText, normalizeUpperText } = require('../utils');

function extractTableOfContents(wordData) {
  const evidence = [];
  const chapters = [];

  if (!wordData || !wordData.paragraphs) {
    return { chapters: [], evidence: ['未能提取段落数据'] };
  }

  const tocPatterns = [
    /目录|table\s+of\s+contents|contents/i,
    /^\s*1\s+/
  ];

  let inTocSection = false;
  const tocLines = [];

  wordData.paragraphs.forEach((para, index) => {
    const text = normalizeText(para);
    const upperText = normalizeUpperText(para);

    if (tocPatterns[0].test(upperText)) {
      inTocSection = true;
      evidence.push(`第 ${index + 1} 行：识别到目录标记`);
    }

    if (inTocSection && text && /^\d+[\s.]+.*\d+\s*$/.test(text)) {
      tocLines.push({
        rawText: text,
        lineIndex: index
      });
    }

    if (inTocSection && /^\s*1[\s.]+\w+/.test(text) && index > 5) {
      inTocSection = false;
    }
  });

  const headingPatterns = [
    { regex: /^(\d+)\s+([a-z ]+)$/i, level: 1 },
    { regex: /^(\d+\.\d+)\s+(.+)$/i, level: 2 }
  ];

  wordData.paragraphs.forEach((para, index) => {
    const text = normalizeText(para);

    headingPatterns.forEach((pattern) => {
      const match = text.match(pattern.regex);
      if (match) {
        chapters.push({
          number: match[1],
          title: match[2],
          level: pattern.level,
          paragraphIndex: index,
          lineNumber: index + 1
        });
      }
    });
  });

  if (tocLines.length === 0 && chapters.length === 0) {
    evidence.push('警告：未找到显式目录信息或标题结构');
  } else {
    evidence.push(`识别到 ${chapters.length} 个章节标题`);
  }

  return {
    chapters,
    tocLines,
    evidence,
    pageCount: wordData.pageCount || null
  };
}

function checkTableOfContentsPages(wordData, tocInfo) {
  const issues = [];
  const evidence = [];

  const totalPages = wordData.pageCount;

  if (!totalPages) {
    issues.push({
      severity: 'warning',
      message: '无法从 Word 文档中确定总页数，需要人工确认'
    });
    evidence.push('未能读取页数统计信息');
    return { issues, evidence, status: 'review' };
  }

  const lastTocEntry = tocInfo.tocLines[tocInfo.tocLines.length - 1];
  if (lastTocEntry) {
    const pageMatch = lastTocEntry.rawText.match(/(\d+)\s*$/);
    if (pageMatch) {
      const lastPageInToc = parseInt(pageMatch[1], 10);
      if (lastPageInToc > totalPages) {
        issues.push({
          severity: 'error',
          message: `目录中最后一章页码 ${lastPageInToc} 超过报告总页数 ${totalPages}，目录可能老旧或错误`
        });
        evidence.push(`目录页码：${lastPageInToc}，文档总页：${totalPages}`);
      } else if (Math.abs(lastPageInToc - totalPages) > 2) {
        issues.push({
          severity: 'warning',
          message: `目录末页 ${lastPageInToc} 与文档总页数 ${totalPages} 的差距较大，可能有附录或额外说明部分`
        });
        evidence.push(`目录末页与总页差距：${Math.abs(lastPageInToc - totalPages)} 页`);
      }
    }
  }

  if (issues.length === 0) {
    evidence.push(`✓ 目录页数与报告总页数相符：${totalPages} 页`);
    return { issues, evidence, status: 'pass' };
  }

  return { issues, evidence, status: 'warning' };
}

function checkChaptersAlignment(wordData, tocInfo) {
  const issues = [];
  const evidence = [];

  const documentChapters = new Set(
    tocInfo.chapters.map((ch) => normalizeUpperText(ch.title))
  );

  const actualChapters = [];
  wordData.paragraphs.forEach((para, index) => {
    const text = normalizeText(para);
    if (/^[0-9.]+\s+[a-z ]+$/i.test(text) && text.length < 80) {
      actualChapters.push({
        text: normalizeUpperText(text),
        originalText: text,
        lineIndex: index + 1
      });
    }
  });

  evidence.push(`目录章节数：${tocInfo.chapters.length}，文档中发现标题数：${actualChapters.length}`);

  const missingChapters = [];
  documentChapters.forEach((tocChapter) => {
    const found = actualChapters.some((actualCh) => actualCh.text.includes(tocChapter.split(/\s+/)[0]));
    if (!found) {
      missingChapters.push(tocChapter);
    }
  });

  if (missingChapters.length > 0) {
    issues.push({
      severity: 'warning',
      message: `目录中有 ${missingChapters.length} 个章节在文档中未找到：${missingChapters.slice(0, 3).join('、')}`
    });
    evidence.push(`缺失章节（部分）：${missingChapters.slice(0, 3).join('、')}`);
  }

  const unlistedChapters = actualChapters.filter((ch) => {
    const chapterNum = ch.text.split(/\s+/)[0];
    return !Array.from(documentChapters).some((tc) => tc.includes(chapterNum));
  });

  if (unlistedChapters.length > 0) {
    issues.push({
      severity: 'review',
      message: `文档中有 ${unlistedChapters.length} 个章节未列在目录中，可能是新增内容或目录未更新`
    });
    evidence.push(`文档中未在目录列出的章节：${unlistedChapters.slice(0, 3).map((c) => c.originalText).join('、')}`);
  }

  if (issues.length === 0) {
    evidence.push('✓ 目录与文档章节结构对应正常');
    return { issues, evidence, status: 'pass' };
  }

  return { issues, evidence, status: missingChapters.length > 2 ? 'warning' : 'review' };
}

module.exports = {
  extractTableOfContents,
  checkTableOfContentsPages,
  checkChaptersAlignment
};