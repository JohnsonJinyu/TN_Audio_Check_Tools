const { normalizeText, normalizeUpperText } = require('../utils');

function extractTableOfContents(wordData) {
  const evidence = [];
  const chapters = [];
  const tocLines = [];

  if (!wordData || !wordData.paragraphs) {
    return { chapters: [], evidence: ['未能提取段落数据'] };
  }

  const headingPatterns = [
    { regex: /^(\d+(?:\.\d+)*)[a-z]?\s+(.+)$/i },
    { regex: /^(?:table|表)\s*(\d+(?:\.\d+)*)\s*[:.-]?\s+(.+)$/i }
  ];

  const allLines = [
    ...(wordData.paragraphs || []).map((text, index) => ({ text: normalizeText(text), index, source: 'paragraph' })),
    ...((wordData.tables || []).flatMap((table, tableIndex) => (table.rows || []).map((row, rowIndex) => ({
      text: row.map((cell) => normalizeText(cell)).filter(Boolean).join(' | '),
      cells: row,
      index: rowIndex,
      source: `table-${tableIndex}`
    }))))
  ].filter((line) => line.text);

  let inTocSection = false;

  allLines.forEach((line) => {
    const text = normalizeText(line.text);
    const upperText = normalizeUpperText(text);

    if (/目录|table\s+of\s+contents|contents/i.test(upperText)) {
      inTocSection = true;
      evidence.push(`${line.source} 第 ${line.index + 1} 行：识别到目录标记`);
      return;
    }

    const rowWithPage = text.match(/^(\d+(?:\.\d+)*[a-z]?)\s+(.+?)(?:\.{2,}|\t|\s{2,}|\|)\s*(\d+)\s*$/i);
    if (inTocSection && rowWithPage) {
      tocLines.push({
        rawText: text,
        lineIndex: line.index,
        source: line.source,
        chapterNumber: rowWithPage[1],
        title: normalizeText(rowWithPage[2]),
        pageNumber: parseInt(rowWithPage[3], 10)
      });
    }

    if (line.source.startsWith('table-') && Array.isArray(line.cells) && line.cells.length >= 2) {
      const lastCell = normalizeText(line.cells[line.cells.length - 1]);
      const firstCell = normalizeText(line.cells[0]);
      if (/^\d+$/.test(lastCell) && /^(\d+(?:\.\d+)*[a-z]?)\s+.+/i.test(firstCell)) {
        const match = firstCell.match(/^(\d+(?:\.\d+)*[a-z]?)\s+(.+)$/i);
        tocLines.push({
          rawText: text,
          lineIndex: line.index,
          source: line.source,
          chapterNumber: match[1],
          title: normalizeText(match[2]),
          pageNumber: parseInt(lastCell, 10)
        });
      }
    }

    if (inTocSection && /^(?:appendix|annex|references|1\s+|1\.)/i.test(text) && tocLines.length > 0 && line.index > 5) {
      inTocSection = false;
    }

    headingPatterns.forEach((pattern) => {
      const match = text.match(pattern.regex);
      if (match && text.length < 200) {
        chapters.push({
          number: match[1],
          title: normalizeText(match[2]),
          paragraphIndex: line.index,
          lineNumber: line.index + 1,
          source: line.source
        });
      }
    });
  });

  const uniqueChapters = Array.from(new Map(
    chapters.map((chapter) => [`${chapter.number}|${normalizeUpperText(chapter.title)}`, chapter])
  ).values());

  if (tocLines.length === 0 && uniqueChapters.length === 0) {
    evidence.push('警告：未找到显式目录信息或标题结构');
  } else {
    evidence.push(`识别到 ${uniqueChapters.length} 个章节标题，${tocLines.length} 条目录项`);
  }

  return {
    chapters: uniqueChapters,
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
    const lastPageInToc = lastTocEntry.pageNumber;
    if (lastPageInToc) {
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

  const tocEntries = tocInfo.tocLines || [];
  const actualChapters = tocInfo.chapters || [];

  evidence.push(`目录章节数：${tocEntries.length || tocInfo.chapters.length}，文档中发现标题数：${actualChapters.length}`);

  const missingChapters = [];
  tocEntries.forEach((tocEntry) => {
    const found = actualChapters.some((chapter) => {
      if (chapter.number === tocEntry.chapterNumber) {
        return true;
      }

      const tocTitle = normalizeUpperText(tocEntry.title);
      const chapterTitle = normalizeUpperText(chapter.title);
      return tocTitle && chapterTitle && (chapterTitle.includes(tocTitle) || tocTitle.includes(chapterTitle));
    });

    if (!found) {
      missingChapters.push(`${tocEntry.chapterNumber} ${tocEntry.title}`.trim());
    }
  });

  if (missingChapters.length > 0) {
    issues.push({
      severity: 'warning',
      message: `目录中有 ${missingChapters.length} 个章节在文档中未找到：${missingChapters.slice(0, 3).join('、')}`
    });
    evidence.push(`缺失章节（部分）：${missingChapters.slice(0, 3).join('、')}`);
  }

  const unlistedChapters = actualChapters.filter((chapter) => {
    return !tocEntries.some((tocEntry) => {
      if (tocEntry.chapterNumber === chapter.number) {
        return true;
      }

      const tocTitle = normalizeUpperText(tocEntry.title);
      const chapterTitle = normalizeUpperText(chapter.title);
      return tocTitle && chapterTitle && (chapterTitle.includes(tocTitle) || tocTitle.includes(chapterTitle));
    });
  });

  if (unlistedChapters.length > 0) {
    issues.push({
      severity: 'review',
      message: `文档中有 ${unlistedChapters.length} 个章节未列在目录中，可能是新增内容或目录未更新`
    });
    evidence.push(`文档中未在目录列出的章节：${unlistedChapters.slice(0, 3).map((chapter) => `${chapter.number} ${chapter.title}`.trim()).join('、')}`);
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