const path = require('path');

function normalizeText(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function normalizeUpperText(value) {
  return normalizeText(value).toUpperCase();
}

function tokenize(value) {
  return Array.from(new Set(
    normalizeUpperText(value)
      .split(/[^A-Z0-9]+/)
      .filter((token) => token.length > 1 && !['REPORT', 'OBJECT', 'MEASUREMENT'].includes(token))
  ));
}

function getReportBaseName(reportPath = '') {
  return path.parse(reportPath || '').name || '';
}

function getBundleKey(reportPath = '', reportContext = {}) {
  return normalizeText(getReportBaseName(reportPath) || reportContext.measurementObject || 'unknown');
}

function aggregateStatus(statusList) {
  const priority = {
    error: 5,
    warning: 4,
    review: 3,
    missing: 2,
    pass: 1,
    not_applicable: 0
  };

  return [...statusList].sort((left, right) => (priority[right] || 0) - (priority[left] || 0))[0] || 'not_applicable';
}

function clampEvidence(evidenceList, limit = 5) {
  return Array.from(new Set((evidenceList || []).map(normalizeText).filter(Boolean))).slice(0, limit);
}

function collectEvidence(lines, patterns, limit = 5) {
  const matches = [];

  for (const line of lines || []) {
    const normalizedLine = normalizeText(line);
    if (!normalizedLine) {
      continue;
    }

    if (patterns.some((pattern) => pattern.test(normalizedLine))) {
      matches.push(normalizedLine);
    }
  }

  return clampEvidence(matches, limit);
}

function toComparableNumber(value) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const normalized = normalizeText(value).replace(/,/g, '');
  if (!normalized) {
    return null;
  }

  if (!/^[-+]?\d+(?:\.\d+)?$/.test(normalized)) {
    return null;
  }

  const parsedValue = Number(normalized);
  return Number.isFinite(parsedValue) ? parsedValue : null;
}

function buildSectionReview(title, lines, patterns) {
  const evidence = collectEvidence(lines, patterns, 5);
  if (evidence.length === 0) {
    return {
      title,
      status: 'warning',
      detected: false,
      requiresEngineerReview: false,
      message: `${title}相关章节未在 Word 中识别到，需要人工确认报告是否缺页或命名不一致。`,
      evidence: []
    };
  }

  return {
    title,
    status: 'review',
    detected: true,
    requiresEngineerReview: true,
    message: `${title}相关章节已识别，可作为自动证据抓取入口，但正确性仍需音频工程师复核。`,
    evidence
  };
}

function buildNameOverlapScore(leftValue, rightValue) {
  const leftTokens = tokenize(leftValue);
  const rightTokens = new Set(tokenize(rightValue));
  if (leftTokens.length === 0 || rightTokens.size === 0) {
    return 0;
  }

  const overlapCount = leftTokens.filter((token) => rightTokens.has(token)).length;
  return overlapCount / leftTokens.length;
}

function buildExcelDuplicateCandidates(tableRows) {
  const signatureMap = new Map();

  tableRows
    .filter((row) => row.sourceKind === 'xlsx-detailed')
    .forEach((row) => {
      const cells = Array.isArray(row.cells) ? row.cells : [];
      const signatureParts = [cells[0], cells[6], cells[7], cells[8], cells[9], cells[10], cells[11], cells[13]]
        .map(normalizeUpperText)
        .filter(Boolean);
      const signature = signatureParts.join('|');
      if (!signature) {
        return;
      }

      const descriptor = normalizeText(cells[0]);
      const current = signatureMap.get(signature) || {
        descriptor,
        count: 0,
        samples: []
      };

      current.count += 1;
      if (current.samples.length < 3) {
        current.samples.push(normalizeText(row.text));
      }
      signatureMap.set(signature, current);
    });

  const duplicates = Array.from(signatureMap.values())
    .filter((item) => item.count > 1)
    .sort((left, right) => right.count - left.count);

  return {
    count: duplicates.length,
    items: duplicates.slice(0, 20).map((item) => ({
      descriptor: item.descriptor,
      count: item.count,
      evidence: clampEvidence(item.samples, 3)
    }))
  };
}

function buildExcelExtraCandidates(tableRows, extractedItems) {
  const descriptorMap = new Map();
  const matchedCorpus = normalizeUpperText(
    (extractedItems || [])
      .filter((item) => item.matched)
      .map((item) => item.sourcePreview)
      .join(' ')
  );

  if (!matchedCorpus) {
    return { count: 0, items: [] };
  }

  tableRows
    .filter((row) => row.sourceKind === 'xlsx-detailed')
    .forEach((row) => {
      const descriptor = normalizeText(row.cells?.[0]);
      const normalizedDescriptor = normalizeUpperText(descriptor);
      if (!normalizedDescriptor || matchedCorpus.includes(normalizedDescriptor)) {
        return;
      }

      const current = descriptorMap.get(normalizedDescriptor) || {
        descriptor,
        count: 0,
        evidence: []
      };

      current.count += 1;
      if (current.evidence.length < 2) {
        current.evidence.push(normalizeText(row.text));
      }
      descriptorMap.set(normalizedDescriptor, current);
    });

  const items = Array.from(descriptorMap.values()).sort((left, right) => right.count - left.count);
  return {
    count: items.length,
    items: items.slice(0, 20).map((item) => ({
      descriptor: item.descriptor,
      count: item.count,
      evidence: clampEvidence(item.evidence, 2)
    }))
  };
}

function analyzeExcelReport({ reportData, extractedItems }) {
  const matchedItems = (extractedItems || []).filter((item) => item.matched);
  const skippedItems = (extractedItems || []).filter((item) => item.skipped);
  const unmatchedItems = (extractedItems || []).filter((item) => !item.matched && !item.skipped);
  const duplicateCandidates = buildExcelDuplicateCandidates(reportData?.tableRows || []);
  const extraCandidates = buildExcelExtraCandidates(reportData?.tableRows || [], extractedItems || []);
  const status = aggregateStatus([
    unmatchedItems.length > 0 ? 'warning' : 'pass',
    duplicateCandidates.count > 0 ? 'review' : 'pass',
    extraCandidates.count > 0 ? 'review' : 'pass'
  ]);

  return {
    category: 'excel',
    coverage: {
      status,
      matchedCount: matchedItems.length,
      missingCount: unmatchedItems.length,
      skippedCount: skippedItems.length,
      duplicateCount: duplicateCandidates.count,
      extraCandidateCount: extraCandidates.count,
      missingItems: unmatchedItems.slice(0, 20).map((item) => ({
        itemId: item.itemId,
        outputCell: item.outputCell,
        checklistDesc: item.checklistDesc,
        reason: item.reason || '未命中'
      })),
      duplicateItems: duplicateCandidates.items,
      extraCandidateItems: extraCandidates.items,
      notes: extraCandidates.count > 0
        ? ['多测项当前按“Excel 结构化行与已回填 checklist 差集”做候选提示，最终仍需结合音频测试计划复核。']
        : []
    }
  };
}

function createWordFinding(id, title, status, message, evidence = []) {
  return {
    id,
    title,
    status,
    message,
    evidence: clampEvidence(evidence, 5)
  };
}

function analyzeWordReport({ reportPath, reportData }) {
  const headerLines = reportData?.structuredData?.headers || [];
  const footerLines = reportData?.structuredData?.footers || [];
  const bodyLines = reportData?.lines || [];
  const combinedLines = clampEvidence([...headerLines, ...bodyLines, ...footerLines], 1000);
  const reportName = getReportBaseName(reportPath);
  const measurementObject = normalizeText(reportData?.reportContext?.measurementObject);
  const headerFooterText = `${headerLines.join(' ')} ${footerLines.join(' ')}`.trim();
  const findings = [];

  const tocEvidence = collectEvidence(combinedLines, [/table of contents/i, /^contents$/i, /目录/], 3);
  findings.push(
    tocEvidence.length > 0
      ? createWordFinding('toc', '目录识别', 'pass', '已在 Word 中识别到目录线索，可继续做目录与章节对应检查。', tocEvidence)
      : createWordFinding('toc', '目录识别', 'warning', '未在 Word 中识别到目录线索，需要人工确认目录是否缺失或未更新。')
  );

  const tableEvidence = collectEvidence(combinedLines, [/^(?:table|表)\s*\d+/i, /\btable\s+\d+\b/i], 5);
  findings.push(
    tableEvidence.length > 0
      ? createWordFinding('tables', 'Table 章节识别', 'pass', '已识别到 table 章节或标题，可继续做目录对应复核。', tableEvidence)
      : createWordFinding('tables', 'Table 章节识别', 'review', '未识别到稳定的 table 标题，目录与 table 对应关系需要人工复核。')
  );

  const sectionEvidence = collectEvidence(combinedLines, [/^\d+(?:\.\d+)+[a-z]?\b/i], 5);
  findings.push(
    sectionEvidence.length > 0
      ? createWordFinding('sections', '章节层级识别', 'pass', '已识别到编号章节，可作为后续文档完整性审查入口。', sectionEvidence)
      : createWordFinding('sections', '章节层级识别', 'review', '未识别到稳定的章节编号，章节层级与目录一致性需要人工确认。')
  );

  const headerFooterOverlap = buildNameOverlapScore(reportName, headerFooterText || measurementObject);
  findings.push(
    headerFooterText
      ? createWordFinding(
        'name-consistency',
        '报告名与页眉页脚一致性',
        headerFooterOverlap >= 0.5 ? 'pass' : 'warning',
        headerFooterOverlap >= 0.5
          ? '报告名与页眉页脚文本基本一致。'
          : '报告名与页眉页脚的关键词重合度偏低，建议人工确认命名一致性。',
        clampEvidence([...headerLines, ...footerLines], 4)
      )
      : createWordFinding('name-consistency', '报告名与页眉页脚一致性', 'review', '当前未提取到页眉页脚文本，无法自动完成命名一致性检查。')
  );

  const objectCandidate = measurementObject || reportName;
  const objectPolluted = /(20\d{6}|\d{8})|\b(?:VER|V)\d+(?:\.\d+)*\b|\bDVT\d+\b/i.test(objectCandidate);
  findings.push(
    createWordFinding(
      'object-cleanliness',
      'Measurement Object 清洁度',
      objectPolluted ? 'review' : objectCandidate ? 'pass' : 'warning',
      objectPolluted
        ? 'Measurement Object 或报告名中疑似混入日期或测试版本，需要人工确认是否污染命名。'
        : objectCandidate
          ? 'Measurement Object 未识别到明显的日期/版本污染。'
          : '未识别到 Measurement Object，需要人工确认。',
      objectCandidate ? [objectCandidate] : []
    )
  );

  const peopleEvidence = collectEvidence(
    combinedLines,
    [/(tested by|prepared by|author|engineer|tester|approved by)\s*[:：]/i, /(测试人员|编制|审核|批准|作者)\s*[:：]/],
    5
  );
  findings.push(
    peopleEvidence.length > 0
      ? createWordFinding('people', '测试人员信息', 'pass', '已识别到测试人员或文档责任人字段。', peopleEvidence)
      : createWordFinding('people', '测试人员信息', 'warning', '未识别到测试人员或责任人字段，建议人工检查。')
  );

  const polqaEvidence = collectEvidence(combinedLines, [/(POLQA|P\.863)/i], 6);
  if (polqaEvidence.length > 0) {
    const versionEvidence = collectEvidence(combinedLines, [/(POLQA|P\.863).{0,80}(version|ver\.?|v)\s*[:：]?\s*[A-Za-z0-9._-]+/i], 3);
    const referenceEvidence = collectEvidence(combinedLines, [/(reference (source|signal|speech|file)|参考音源|参考语音|source file)/i], 3);
    findings.push(
      createWordFinding(
        'polqa-config',
        'POLQA 配置检查',
        versionEvidence.length > 0 && referenceEvidence.length > 0 ? 'pass' : 'review',
        versionEvidence.length > 0 && referenceEvidence.length > 0
          ? '已识别到 POLQA/P.863 版本与参考音源配置。'
          : 'POLQA/P.863 已出现，但版本或参考音源信息不完整，需要人工复核。',
        clampEvidence([...versionEvidence, ...referenceEvidence, ...polqaEvidence], 5)
      )
    );
  }

  const loudnessReview = buildSectionReview('响度', combinedLines, [
    /loudness/i,
    /send loudness/i,
    /receive loudness/i,
    /\bSLR\b/i,
    /\bRLR\b/i,
    /响度/
  ]);
  const frequencyResponseReview = buildSectionReview('频响曲线', combinedLines, [
    /frequency response/i,
    /frequency characteristic/i,
    /frequency characteristics/i,
    /freq\.?\s*response/i,
    /频响/i,
    /频率响应/
  ]);

  return {
    category: 'word',
    documentCompleteness: {
      status: aggregateStatus(findings.map((item) => item.status)),
      findings
    },
    curveReview: {
      loudness: loudnessReview,
      frequencyResponse: frequencyResponseReview
    }
  };
}

function buildComparableItems(result) {
  const comparableMap = new Map();

  (result.extractedItems || [])
    .filter((item) => item.matched)
    .forEach((item) => {
      comparableMap.set(item.outputCell, {
        outputCell: item.outputCell,
        checklistDesc: item.checklistDesc,
        value: item.value
      });
    });

  return comparableMap;
}

function buildConsistencyGroups(excelResults, groupBy) {
  const groupMap = new Map();

  excelResults.forEach((result) => {
    const reportContext = result.reportContext || {};
    const groupKey = groupBy(reportContext);
    if (!groupKey || groupKey.includes('unknown')) {
      return;
    }

    const current = groupMap.get(groupKey) || [];
    current.push(result);
    groupMap.set(groupKey, current);
  });

  return Array.from(groupMap.values()).filter((group) => group.length > 1);
}

function analyzeConsistencyGroup(group, comparisonType) {
  const itemMap = new Map();

  group.forEach((result) => {
    buildComparableItems(result).forEach((item, outputCell) => {
      const current = itemMap.get(outputCell) || {
        outputCell,
        checklistDesc: item.checklistDesc,
        values: []
      };

      current.values.push({
        reportName: getReportBaseName(result.reportPath),
        value: item.value
      });
      itemMap.set(outputCell, current);
    });
  });

  const flaggedItems = [];
  let comparableItemCount = 0;

  itemMap.forEach((item) => {
    if (item.values.length < 2) {
      return;
    }

    comparableItemCount += 1;
    const numericValues = item.values
      .map((entry) => ({ ...entry, numericValue: toComparableNumber(entry.value) }))
      .filter((entry) => entry.numericValue !== null);

    if (numericValues.length === item.values.length) {
      const onlyValues = numericValues.map((entry) => entry.numericValue);
      const spread = Math.max(...onlyValues) - Math.min(...onlyValues);
      if (spread > 1.5 || spread > 0.3) {
        flaggedItems.push({
          outputCell: item.outputCell,
          checklistDesc: item.checklistDesc,
          severity: spread > 1.5 ? 'warning' : 'review',
          reason: spread > 1.5 ? '跨报告数值差异偏大' : '跨报告数值存在差异，建议复核',
          spread: Number(spread.toFixed(3)),
          values: item.values
        });
      }
      return;
    }

    const distinctValues = new Set(item.values.map((entry) => normalizeUpperText(entry.value)));
    if (distinctValues.size > 1) {
      flaggedItems.push({
        outputCell: item.outputCell,
        checklistDesc: item.checklistDesc,
        severity: 'review',
        reason: '跨报告状态或文本值不一致',
        values: item.values
      });
    }
  });

  const reportContext = group[0]?.reportContext || {};
  return {
    comparisonType,
    status: comparableItemCount === 0 ? 'not_applicable' : aggregateStatus(flaggedItems.map((item) => item.severity).concat('pass')),
    groupKey: comparisonType === 'same-codec-cross-network'
      ? `${reportContext.codec || 'unknown'} | ${reportContext.bandwidth || 'unknown'} | ${reportContext.terminalMode || 'unknown'}`
      : `${reportContext.network || 'unknown'} | ${reportContext.bandwidth || 'unknown'} | ${reportContext.terminalMode || 'unknown'}`,
    reports: group.map((result) => ({
      reportName: getReportBaseName(result.reportPath),
      network: result.reportContext?.network || '',
      codec: result.reportContext?.codec || ''
    })),
    comparableItemCount,
    flaggedItems: flaggedItems.slice(0, 30)
  };
}

function analyzeConsistency(excelResults) {
  const sameCodecGroups = buildConsistencyGroups(
    excelResults,
    (context) => `${context.codec || 'unknown'}|${context.bandwidth || 'unknown'}|${context.terminalMode || 'unknown'}`
  )
    .filter((group) => new Set(group.map((result) => result.reportContext?.network || 'unknown')).size > 1)
    .map((group) => analyzeConsistencyGroup(group, 'same-codec-cross-network'));

  const sameNetworkGroups = buildConsistencyGroups(
    excelResults,
    (context) => `${context.network || 'unknown'}|${context.bandwidth || 'unknown'}|${context.terminalMode || 'unknown'}`
  )
    .filter((group) => new Set(group.map((result) => result.reportContext?.codec || 'unknown')).size > 1)
    .map((group) => analyzeConsistencyGroup(group, 'same-network-cross-codec'));

  const groups = [...sameCodecGroups, ...sameNetworkGroups].filter((group) => group.comparableItemCount > 0);
  const flaggedCount = groups.reduce((sum, group) => sum + group.flaggedItems.length, 0);

  return {
    status: groups.length === 0 ? 'not_applicable' : aggregateStatus(groups.map((group) => group.status)),
    enabled: groups.length > 0,
    groupCount: groups.length,
    flaggedCount,
    groups
  };
}

function mergeContext(items) {
  return items.reduce((accumulator, item) => ({
    measurementObject: accumulator.measurementObject || item.reportContext?.measurementObject || '',
    customer: accumulator.customer || item.reportContext?.customer || '',
    codec: accumulator.codec || item.reportContext?.codec || '',
    network: accumulator.network || item.reportContext?.network || '',
    bandwidth: accumulator.bandwidth || item.reportContext?.bandwidth || '',
    terminalMode: accumulator.terminalMode || item.reportContext?.terminalMode || '',
    reportPanelSelections: accumulator.reportPanelSelections || item.reportContext?.reportPanelSelections || null
  }), {
    measurementObject: '',
    customer: '',
    codec: '',
    network: '',
    bandwidth: '',
    terminalMode: '',
    reportPanelSelections: null
  });
}

function buildRunConfigSummary(results) {
  const successResults = (results || []).filter((item) => item.status === 'success');
  const firstContext = successResults.find((item) => item.reportContext)?.reportContext || {};
  const profileKeys = Array.from(new Set(
    successResults
      .map((item) => normalizeText(item.ruleProfileKey || ''))
      .filter(Boolean)
  ));

  return {
    customer: normalizeText(firstContext.customer || ''),
    reportPanelSelections: firstContext.reportPanelSelections || null,
    ruleProfiles: profileKeys
  };
}

function buildSkipReasonStats(excelResults) {
  const dimensionMap = new Map();

  for (const result of excelResults || []) {
    for (const item of result.skippedItems || []) {
      const dimension = normalizeText(item?.skipContext?.dimension || 'unknown');
      const actual = normalizeText(item?.skipContext?.actual || 'unknown');
      const key = `${dimension}::${actual}`;
      const current = dimensionMap.get(key) || {
        dimension,
        actual,
        count: 0,
        examples: []
      };

      current.count += 1;
      if (current.examples.length < 3) {
        current.examples.push(`${item.outputCell} - ${item.checklistDesc}`);
      }

      dimensionMap.set(key, current);
    }
  }

  const stats = Array.from(dimensionMap.values()).sort((left, right) => right.count - left.count);
  return {
    totalGroups: stats.length,
    topGroups: stats.slice(0, 10)
  };
}

function buildBatchConclusion({ results, checklistPath }) {
  const successResults = (results || []).filter((item) => item.status === 'success');
  const excelResults = successResults.filter((item) => item.reportKind === 'excel');
  const wordResults = successResults.filter((item) => item.reportKind === 'word');
  const bundleMap = new Map();

  successResults.forEach((result) => {
    const bundleKey = result.bundleKey || getBundleKey(result.reportPath, result.reportContext);
    const current = bundleMap.get(bundleKey) || [];
    current.push(result);
    bundleMap.set(bundleKey, current);
  });

  const bundles = Array.from(bundleMap.entries()).map(([key, items]) => {
    const excelItems = items.filter((item) => item.reportKind === 'excel');
    const wordItems = items.filter((item) => item.reportKind === 'word');

    return {
      key,
      context: mergeContext(items),
      sourceMode: excelItems.length > 0 && wordItems.length > 0
        ? 'excel+word'
        : excelItems.length > 0
          ? 'excel'
          : 'word',
      hasChecklistOutput: excelItems.some((item) => item.outputPath),
      excelCount: excelItems.length,
      wordCount: wordItems.length,
      reportNames: items.map((item) => getReportBaseName(item.reportPath)),
      excelCoverage: {
        status: aggregateStatus(excelItems.map((item) => item.audit?.coverage?.status || 'not_applicable')),
        missingCount: excelItems.reduce((sum, item) => sum + (item.audit?.coverage?.missingCount || 0), 0),
        duplicateCount: excelItems.reduce((sum, item) => sum + (item.audit?.coverage?.duplicateCount || 0), 0),
        extraCandidateCount: excelItems.reduce((sum, item) => sum + (item.audit?.coverage?.extraCandidateCount || 0), 0)
      },
      wordAudit: {
        status: aggregateStatus(wordItems.flatMap((item) => [
          item.audit?.documentCompleteness?.status || 'not_applicable',
          item.audit?.curveReview?.loudness?.status || 'not_applicable',
          item.audit?.curveReview?.frequencyResponse?.status || 'not_applicable'
        ])),
        findingCount: wordItems.reduce((sum, item) => sum + (item.audit?.documentCompleteness?.findings?.length || 0), 0),
        loudnessDetected: wordItems.some((item) => item.audit?.curveReview?.loudness?.detected),
        frequencyDetected: wordItems.some((item) => item.audit?.curveReview?.frequencyResponse?.detected)
      },
      items
    };
  }).sort((left, right) => left.key.localeCompare(right.key));

  const consistency = analyzeConsistency(excelResults);
  const totalMissingCount = excelResults.reduce((sum, item) => sum + (item.audit?.coverage?.missingCount || 0), 0);
  const totalDuplicateCount = excelResults.reduce((sum, item) => sum + (item.audit?.coverage?.duplicateCount || 0), 0);
  const totalExtraCandidateCount = excelResults.reduce((sum, item) => sum + (item.audit?.coverage?.extraCandidateCount || 0), 0);
  const skipReasonStats = buildSkipReasonStats(excelResults);
  const wordFindingCount = wordResults.reduce((sum, item) => sum + (item.audit?.documentCompleteness?.findings?.length || 0), 0);
  const loudnessDetectedCount = wordResults.filter((item) => item.audit?.curveReview?.loudness?.detected).length;
  const frequencyDetectedCount = wordResults.filter((item) => item.audit?.curveReview?.frequencyResponse?.detected).length;

  const suggestedActions = [];
  if (!checklistPath && excelResults.length > 0) {
    suggestedActions.push('当前存在 Excel 报告但未提供 checklist，无法完成自动填表与覆盖性评估。');
  }
  if (totalMissingCount > 0) {
    suggestedActions.push(`Excel 覆盖性评估识别到 ${totalMissingCount} 个漏测项，建议先优先处理这些 checklist 缺口。`);
  }
  if (totalDuplicateCount > 0) {
    suggestedActions.push(`Excel 结构化结果中识别到 ${totalDuplicateCount} 组重测候选，需要结合测试计划做复核。`);
  }
  if (consistency.enabled && consistency.flaggedCount > 0) {
    suggestedActions.push(`跨网络/跨 codec 一致性检查已触发 ${consistency.flaggedCount} 个差异项，建议在结论窗口重点人工复核。`);
  }
  if (wordResults.length === 0) {
    suggestedActions.push('当前没有 Word 报告，曲线分析与文档完整性审查证据仍不完整。');
  }

  return {
    runConfig: buildRunConfigSummary(results || []),
    overview: {
      totalReports: results.length,
      successCount: successResults.length,
      errorCount: (results || []).filter((item) => item.status === 'error').length,
      excelCount: (results || []).filter((item) => item.reportKind === 'excel').length,
      wordCount: (results || []).filter((item) => item.reportKind === 'word').length,
      checklistCount: checklistPath ? 1 : 0,
      outputCount: excelResults.filter((item) => item.outputPath).length
    },
    excelCoverage: {
      status: aggregateStatus(excelResults.map((item) => item.audit?.coverage?.status || 'not_applicable')),
      reportCount: excelResults.length,
      matchedCount: excelResults.reduce((sum, item) => sum + (item.audit?.coverage?.matchedCount || 0), 0),
      missingCount: totalMissingCount,
      skippedCount: excelResults.reduce((sum, item) => sum + (item.audit?.coverage?.skippedCount || 0), 0),
      duplicateCount: totalDuplicateCount,
      extraCandidateCount: totalExtraCandidateCount,
      skipReasonStats,
      reportSummaries: excelResults.map((item) => ({
        reportName: getReportBaseName(item.reportPath),
        status: item.audit?.coverage?.status || 'not_applicable',
        missingCount: item.audit?.coverage?.missingCount || 0,
        duplicateCount: item.audit?.coverage?.duplicateCount || 0,
        extraCandidateCount: item.audit?.coverage?.extraCandidateCount || 0,
        missingItems: item.audit?.coverage?.missingItems || [],
        duplicateItems: item.audit?.coverage?.duplicateItems || [],
        extraCandidateItems: item.audit?.coverage?.extraCandidateItems || [],
        notes: item.audit?.coverage?.notes || []
      }))
    },
    wordAudit: {
      status: aggregateStatus(wordResults.flatMap((item) => [
        item.audit?.documentCompleteness?.status || 'not_applicable',
        item.audit?.curveReview?.loudness?.status || 'not_applicable',
        item.audit?.curveReview?.frequencyResponse?.status || 'not_applicable'
      ])),
      reportCount: wordResults.length,
      findingCount: wordFindingCount,
      loudnessDetectedCount,
      frequencyDetectedCount,
      reportSummaries: wordResults.map((item) => ({
        reportName: getReportBaseName(item.reportPath),
        documentStatus: item.audit?.documentCompleteness?.status || 'not_applicable',
        findings: item.audit?.documentCompleteness?.findings || [],
        loudness: item.audit?.curveReview?.loudness || null,
        frequencyResponse: item.audit?.curveReview?.frequencyResponse || null
      }))
    },
    consistency,
    bundles,
    suggestedActions
  };
}

module.exports = {
  analyzeExcelReport,
  analyzeWordReport,
  buildBatchConclusion
};