import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Card, Button, Upload, Table, Space, Tag, Modal, Progress, message, Typography, Select } from 'antd';
import { UploadOutlined, DeleteOutlined, CheckCircleOutlined, ExportOutlined } from '@ant-design/icons';
import { recordReportCheckResults } from '../modules/dashboard/storage';
import '../styles/pages.css';

const { Text, Paragraph } = Typography;
const compactUploadDraggerStyle = { padding: '10px 14px', minHeight: '96px' };
const CUSTOMER_OPTIONS = ['MOTOROLA', 'SAMSUNG', 'T-Mobile', 'ATT'];
const REPORT_PANEL_FIELDS = [
  { cell: 'B13', label: 'Headset Interface' },
  { cell: 'B15', label: 'Network' },
  { cell: 'C15', label: 'Vocoder' },
  { cell: 'D15', label: 'Bitrate' }
];

function getReportKind(fileName = '') {
  const normalized = String(fileName).toLowerCase();
  if (normalized.endsWith('.xlsx') || normalized.endsWith('.xls')) {
    return 'excel';
  }

  if (normalized.endsWith('.doc') || normalized.endsWith('.docx')) {
    return 'word';
  }

  return 'unknown';
}

function getBundleKey(fileName = '') {
  return String(fileName).replace(/\.(xlsx|xls|docx|doc)$/i, '');
}

function detectReportContext(fileName = '') {
  const normalizedName = getBundleKey(fileName).toUpperCase();
  const parts = normalizedName.split(/[_\-\s]+/).filter(Boolean);

  const network = ['VOLTE', 'VOWIFI', 'VONR', 'VOIP', 'WCDMA', 'GSM'].find((item) => parts.includes(item)) || '';
  const codec = ['EVS', 'AMR'].find((item) => parts.includes(item)) || '';
  const bandwidth = ['SWB', 'WB', 'NB', 'SB'].find((item) => parts.includes(item)) || '';
  const terminalMode = ['HA', 'HF', 'HS', 'HE', 'HH'].find((item) => parts.includes(item)) || '';

  return {
    measurementObject: normalizedName,
    network,
    codec,
    bandwidth,
    terminalMode
  };
}

function buildUploadSummary(files, checklistFile) {
  const excelFiles = files.filter((item) => item.reportKind === 'excel');
  const wordFiles = files.filter((item) => item.reportKind === 'word');
  return {
    totalReports: files.length,
    excelCount: excelFiles.length,
    wordCount: wordFiles.length,
    checklistCount: checklistFile?.path ? 1 : 0
  };
}

function getStatusMeta(status = 'not_applicable') {
  const statusMap = {
    pass: { color: 'green', label: '通过' },
    review: { color: 'gold', label: '待复核' },
    warning: { color: 'red', label: '异常' },
    missing: { color: 'orange', label: '缺失' },
    success: { color: 'green', label: '完成' },
    error: { color: 'red', label: '失败' },
    pending: { color: 'blue', label: '待处理' },
    processing: { color: 'gold', label: '处理中' },
    not_applicable: { color: 'default', label: '未触发' }
  };

  return statusMap[status] || { color: 'default', label: status || '未知' };
}

function buildConclusionData(uploadSummary, processedConclusion) {
  if (processedConclusion) {
    return {
      ...processedConclusion,
      overview: {
        ...processedConclusion.overview,
        totalReports: uploadSummary.totalReports,
        excelCount: uploadSummary.excelCount,
        wordCount: uploadSummary.wordCount,
        checklistCount: uploadSummary.checklistCount
      }
    };
  }

  return {
    runConfig: {
      customer: '',
      reportPanelSelections: null,
      ruleProfiles: []
    },
    overview: {
      totalReports: uploadSummary.totalReports,
      successCount: 0,
      errorCount: 0,
      excelCount: uploadSummary.excelCount,
      wordCount: uploadSummary.wordCount,
      checklistCount: uploadSummary.checklistCount,
      outputCount: 0
    },
    excelCoverage: {
      status: 'not_applicable',
      reportCount: uploadSummary.excelCount,
      matchedCount: 0,
      missingCount: 0,
      skippedCount: 0,
      duplicateCount: 0,
      extraCandidateCount: 0,
      reportSummaries: []
    },
    wordAudit: {
      status: 'not_applicable',
      reportCount: uploadSummary.wordCount,
      findingCount: 0,
      loudnessDetectedCount: 0,
      frequencyDetectedCount: 0,
      reportSummaries: []
    },
    consistency: {
      status: 'not_applicable',
      enabled: false,
      groupCount: 0,
      flaggedCount: 0,
      groups: []
    },
    bundles: [],
    suggestedActions: uploadSummary.totalReports === 0
      ? ['上传 Excel、Word 和 checklist 后，这里会输出填表结果、文档审查和一致性结论。']
      : []
  };
}

function getOutputFileName(outputPath) {
  if (!outputPath) {
    return '';
  }

  const normalizedPath = String(outputPath).replace(/\\/g, '/');
  const segments = normalizedPath.split('/');
  return segments[segments.length - 1] || outputPath;
}

function ReportChecker() {
  const [files, setFiles] = useState([]);
  const [ruleFile, setRuleFile] = useState(null);
  const [checklistFile, setChecklistFile] = useState(null);
  const [selectedCustomer, setSelectedCustomer] = useState('MOTOROLA');
  const [reportPanelMeta, setReportPanelMeta] = useState({
    reportSheetName: 'Report',
    fields: []
  });
  const [reportPanelSelections, setReportPanelSelections] = useState({
    B13: '',
    B15: '',
    C15: '',
    D15: ''
  });
  const [reportPanelDynamicOptions, setReportPanelDynamicOptions] = useState({});
  const [processedConclusion, setProcessedConclusion] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [exportingRules, setExportingRules] = useState(false);
  const [progressState, setProgressState] = useState({
    active: false,
    total: 0,
    completed: 0,
    successCount: 0,
    errorCount: 0,
    currentReportName: ''
  });
  const activeRunIdRef = useRef(null);

  useEffect(() => {
    if (!window.electron?.reportChecker?.onProgress) {
      return undefined;
    }

    const unsubscribe = window.electron.reportChecker.onProgress((payload) => {
      if (!payload || payload.runId !== activeRunIdRef.current) {
        return;
      }

      if (payload.type === 'batch-start') {
        setProgressState({
          active: true,
          total: payload.total || 0,
          completed: payload.completed || 0,
          successCount: payload.successCount || 0,
          errorCount: payload.errorCount || 0,
          currentReportName: ''
        });
        return;
      }

      if (payload.type === 'report-complete' && payload.result) {
        const result = payload.result;

        setFiles((prev) => prev.map((item) => {
          if (item.path !== result.reportPath) {
            return item;
          }

          if (result.status === 'error') {
            return {
              ...item,
              status: 'error',
              error: result.error,
              items: 0,
              outputPath: '',
              outputName: '',
              unmatchedItems: [],
              skippedItems: [],
              audit: null
            };
          }

          return {
            ...item,
            status: 'success',
            items: result.matchedItems,
            ruleProfileKey: result.ruleProfileKey || item.ruleProfileKey || '',
            outputPath: result.outputPath,
            outputName: getOutputFileName(result.outputPath),
            unmatchedItems: result.unmatchedItems || [],
            skippedItems: result.skippedItems || [],
            reportContext: result.reportContext || item.reportContext,
            audit: result.audit || null,
            error: ''
          };
        }));

        setProgressState({
          active: true,
          total: payload.total || 0,
          completed: payload.completed || 0,
          successCount: payload.successCount || 0,
          errorCount: payload.errorCount || 0,
          currentReportName: getOutputFileName(result.reportPath)
        });
        return;
      }

      if (payload.type === 'batch-complete') {
        setProgressState((prev) => ({
          ...prev,
          active: false,
          total: payload.total || prev.total,
          completed: payload.completed || prev.completed,
          successCount: payload.successCount || prev.successCount,
          errorCount: payload.errorCount || prev.errorCount
        }));
      }
    });

    return () => {
      unsubscribe();
    };
  }, []);

  const loadChecklistReportPanelOptions = async (checklistPath) => {
    if (!checklistPath || !window.electron?.reportChecker?.getChecklistReportOptions) {
      return;
    }

    try {
      const result = await window.electron.reportChecker.getChecklistReportOptions(checklistPath);
      const nextFields = Array.isArray(result?.fields) ? result.fields : [];
      const nextSelections = nextFields.reduce((acc, field) => {
        acc[field.cell] = field.currentValue || '';
        return acc;
      }, { B13: '', B15: '', C15: '', D15: '' });

      setReportPanelMeta({
        reportSheetName: result?.reportSheetName || 'Report',
        fields: nextFields,
        note: result?.note || ''
      });
      setReportPanelSelections(nextSelections);
      setReportPanelDynamicOptions({});

      if (result?.note) {
        message.info(result.note);
      }
    } catch (error) {
      setReportPanelMeta({ reportSheetName: 'Report', fields: [] });
      setReportPanelSelections({ B13: '', B15: '', C15: '', D15: '' });
      setReportPanelDynamicOptions({});
      message.warning(error?.message || '读取 checklist Report 参数失败，将使用报告自动识别值。');
    }
  };

  const updateReportPanelSelection = (cell, value) => {
    const newSelections = { ...reportPanelSelections, [cell]: value || '' };
    const newDynamicOptions = { ...reportPanelDynamicOptions };

    if (cell === 'B15') {
      const c15Field = reportPanelMeta.fields.find((f) => f.cell === 'C15');
      const c15Options = (c15Field?.cascadeMap || {})[value] || c15Field?.options || [];
      newDynamicOptions['C15'] = c15Options;
      const firstC15 = c15Options[0] || '';
      newSelections['C15'] = firstC15;

      const d15Field = reportPanelMeta.fields.find((f) => f.cell === 'D15');
      const d15Options = (d15Field?.cascadeMap || {})[firstC15] || d15Field?.options || [];
      newDynamicOptions['D15'] = d15Options;
      newSelections['D15'] = d15Options[0] || '';
    } else if (cell === 'C15') {
      const d15Field = reportPanelMeta.fields.find((f) => f.cell === 'D15');
      const d15Options = (d15Field?.cascadeMap || {})[value] || d15Field?.options || [];
      newDynamicOptions['D15'] = d15Options;
      newSelections['D15'] = d15Options[0] || '';
    }

    setReportPanelSelections(newSelections);
    setReportPanelDynamicOptions(newDynamicOptions);
  };

  const handleUpload = (file, target, onSuccess) => {
    if (!file.path) {
      message.error('当前环境未提供本地文件路径，无法执行桌面端文件处理。');
      return;
    }

    if (target === 'report') {
      const extension = file.name.toLowerCase();
      if (!extension.endsWith('.doc') && !extension.endsWith('.docx') && !extension.endsWith('.xlsx') && !extension.endsWith('.xls')) {
        message.error('当前后台仅支持 .xlsx / .xls / .doc / .docx 测试报告。');
        return;
      }

      const newItem = {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        name: file.name,
        path: file.path,
        bundleKey: getBundleKey(file.name),
        reportKind: getReportKind(file.name),
        reportContext: detectReportContext(file.name),
        status: 'pending',
        items: 0,
        ruleProfileKey: '',
        outputPath: '',
        outputName: '',
        error: '',
        skippedItems: [],
        unmatchedItems: [],
        audit: null
      };

      setProcessedConclusion(null);
      setFiles((prev) => {
        const exists = prev.some((item) => item.path === file.path);
        return exists ? prev : [newItem, ...prev];
      });
      message.success(`已添加报告: ${file.name}`);
    }

    if (target === 'rules') {
      setProcessedConclusion(null);
      setRuleFile({ name: file.name, path: file.path });
      message.success(`已上传规则: ${file.name}`);
    }

    if (target === 'checklist') {
      setProcessedConclusion(null);
      setChecklistFile({ name: file.name, path: file.path });
      message.success(`已上传 checklist: ${file.name}`);
      loadChecklistReportPanelOptions(file.path);
    }

    if (onSuccess) {
      setTimeout(() => onSuccess('ok'), 0);
    }
  };

  const removeReport = (reportId) => {
    setProcessedConclusion(null);
    setFiles((prev) => prev.filter((item) => item.id !== reportId));
  };

  const clearReports = () => {
    if (processing || files.length === 0) {
      return;
    }

  setProcessedConclusion(null);
    setFiles([]);
    setProgressState({
      active: false,
      total: 0,
      completed: 0,
      successCount: 0,
      errorCount: 0,
      currentReportName: ''
    });
    message.success('已清空测试报告列表');
  };

  const openOutputFolder = async (record) => {
    if (!record.outputPath) {
      message.warning('该报告还没有生成输出文件。');
      return;
    }

    try {
      await window.electron.reportChecker.showOutputInFolder(record.outputPath);
    } catch (error) {
      message.error(error?.message || '打开输出目录失败');
    }
  };

  const showDetails = (record) => {
    Modal.info({
      title: record.name,
      width: 860,
      content: (
        <div style={{ marginTop: 16 }}>
          <Paragraph>
            <Text strong>状态：</Text> {record.status}
          </Paragraph>
          <Paragraph>
            <Text strong>报告类型：</Text> {record.reportKind === 'word' ? 'Word 审查' : 'Excel 填表'}
          </Paragraph>
          <Paragraph>
            <Text strong>识别上下文：</Text>
            {' '}
            {[record.reportContext?.codec, record.reportContext?.network, record.reportContext?.bandwidth, record.reportContext?.terminalMode]
              .filter(Boolean)
              .join(' / ') || '未识别'}
          </Paragraph>
          <Paragraph>
            <Text strong>客户：</Text> {record.reportContext?.customer || selectedCustomer || '未指定'}
          </Paragraph>
          <Paragraph>
            <Text strong>Report 参数：</Text>
            {' '}
            {record.reportContext?.reportPanelSelections
              ? `B13=${record.reportContext.reportPanelSelections.B13 || '-'} / B15=${record.reportContext.reportPanelSelections.B15 || '-'} / C15=${record.reportContext.reportPanelSelections.C15 || '-'} / D15=${record.reportContext.reportPanelSelections.D15 || '-'}`
              : '未指定'}
          </Paragraph>
          <Paragraph>
            <Text strong>命中规则数：</Text> {record.items || 0}
          </Paragraph>
          <Paragraph>
            <Text strong>生效规则 Profile：</Text> {record.ruleProfileKey || '未标记'}
          </Paragraph>
          <Paragraph>
            <Text strong>输出文件：</Text> {record.outputPath || (record.reportKind === 'word' ? 'Word 审查不生成 checklist 输出' : '尚未生成')}
          </Paragraph>
          {record.error ? (
            <Paragraph type="danger">
              <Text strong>错误：</Text> {record.error}
            </Paragraph>
          ) : null}
          {record.unmatchedItems?.length ? (
            <div>
              <Text strong>未命中规则：</Text>
              <div style={{ maxHeight: 240, overflow: 'auto', marginTop: 8, paddingRight: 8 }}>
                {record.unmatchedItems.slice(0, 20).map((item) => (
                  <Paragraph key={`${record.id}-${item.itemId}`} style={{ marginBottom: 8 }}>
                    {item.outputCell} - {item.checklistDesc} ({item.reason})
                  </Paragraph>
                ))}
              </div>
            </div>
          ) : null}
          {record.skippedItems?.length ? (
            <div style={{ marginTop: 16 }}>
              <Text strong>按场景跳过：</Text>
              <div style={{ maxHeight: 180, overflow: 'auto', marginTop: 8, paddingRight: 8 }}>
                {record.skippedItems.map((item) => (
                  <Paragraph key={`${record.id}-skipped-${item.itemId}`} style={{ marginBottom: 8 }}>
                    {item.outputCell} - {item.checklistDesc} ({item.reason})
                    {item.skipContext
                      ? `；维度=${item.skipContext.dimension || '-'}，实际=${item.skipContext.actual || '-'}，允许=${(item.skipContext.include || []).join('/') || '-'}，排除=${(item.skipContext.exclude || []).join('/') || '-'}`
                      : ''}
                  </Paragraph>
                ))}
              </div>
            </div>
          ) : null}
          {record.audit?.coverage ? (
            <div style={{ marginTop: 16 }}>
              <Text strong>Excel 覆盖性评估：</Text>
              <div style={{ marginTop: 8 }}>
                <Paragraph style={{ marginBottom: 8 }}>
                  状态：<Tag color={getStatusMeta(record.audit.coverage.status).color}>{getStatusMeta(record.audit.coverage.status).label}</Tag>
                  漏测 {record.audit.coverage.missingCount}，重测候选 {record.audit.coverage.duplicateCount}，多测候选 {record.audit.coverage.extraCandidateCount}
                </Paragraph>
                {record.audit.coverage.notes?.map((note) => (
                  <Paragraph key={note} style={{ marginBottom: 8 }}>{note}</Paragraph>
                ))}
              </div>
            </div>
          ) : null}
          {record.audit?.documentCompleteness ? (
            <div style={{ marginTop: 16 }}>
              <Text strong>Word 文档审查：</Text>
              <div style={{ marginTop: 8, maxHeight: 220, overflow: 'auto', paddingRight: 8 }}>
                {record.audit.documentCompleteness.findings.map((finding) => (
                  <Paragraph key={finding.id} style={{ marginBottom: 10 }}>
                    <Tag color={getStatusMeta(finding.status).color}>{getStatusMeta(finding.status).label}</Tag>
                    {finding.title} - {finding.message}
                    {finding.evidence?.length ? `；证据：${finding.evidence.join('；')}` : ''}
                  </Paragraph>
                ))}
              </div>
            </div>
          ) : null}
          {record.audit?.curveReview ? (
            <div style={{ marginTop: 16 }}>
              <Text strong>曲线审查入口：</Text>
              <div style={{ marginTop: 8 }}>
                {['loudness', 'frequencyResponse'].map((key) => {
                  const review = record.audit.curveReview[key];
                  if (!review) {
                    return null;
                  }

                  return (
                    <Paragraph key={key} style={{ marginBottom: 10 }}>
                      <Tag color={getStatusMeta(review.status).color}>{getStatusMeta(review.status).label}</Tag>
                      {review.title} - {review.message}
                      {review.evidence?.length ? `；证据：${review.evidence.join('；')}` : ''}
                    </Paragraph>
                  );
                })}
              </div>
            </div>
          ) : null}
        </div>
      )
    });
  };

  const showExcelCoverageSummary = (conclusionData) => {
    Modal.info({
      title: 'Excel 填表与覆盖性',
      width: 940,
      content: (
        <div style={{ marginTop: 16 }}>
          <Paragraph>
            <Tag color={getStatusMeta(conclusionData.excelCoverage.status).color}>{getStatusMeta(conclusionData.excelCoverage.status).label}</Tag>
            共 {conclusionData.excelCoverage.reportCount} 份 Excel 报告，命中 {conclusionData.excelCoverage.matchedCount} 项，漏测 {conclusionData.excelCoverage.missingCount} 项，重测候选 {conclusionData.excelCoverage.duplicateCount} 组，多测候选 {conclusionData.excelCoverage.extraCandidateCount} 组。
          </Paragraph>
          {conclusionData.excelCoverage.skipReasonStats?.topGroups?.length ? (
            <div style={{ marginBottom: 16 }}>
              <Text strong>跳过原因统计：</Text>
              <div style={{ marginTop: 8 }}>
                {conclusionData.excelCoverage.skipReasonStats.topGroups.map((group) => (
                  <Paragraph key={`${group.dimension}-${group.actual}`} style={{ marginBottom: 6 }}>
                    {group.dimension} / {group.actual}：{group.count} 项；示例 {group.examples.join('；')}
                  </Paragraph>
                ))}
              </div>
            </div>
          ) : null}
          <div style={{ maxHeight: 420, overflow: 'auto', paddingRight: 8 }}>
            {conclusionData.excelCoverage.reportSummaries.map((summary) => (
              <div key={summary.reportName} style={{ marginBottom: 16, paddingBottom: 12, borderBottom: '1px solid #f0f0f0' }}>
                <Paragraph style={{ marginBottom: 8 }}>
                  <Text strong>{summary.reportName}</Text>
                  {' '}
                  <Tag color={getStatusMeta(summary.status).color}>{getStatusMeta(summary.status).label}</Tag>
                </Paragraph>
                <Paragraph style={{ marginBottom: 8 }}>
                  漏测 {summary.missingCount}，重测候选 {summary.duplicateCount}，多测候选 {summary.extraCandidateCount}
                </Paragraph>
                {summary.missingItems?.slice(0, 6).map((item) => (
                  <Paragraph key={`${summary.reportName}-${item.itemId}`} style={{ marginBottom: 6 }}>
                    漏测: {item.outputCell} - {item.checklistDesc} ({item.reason})
                  </Paragraph>
                ))}
                {summary.duplicateItems?.slice(0, 4).map((item) => (
                  <Paragraph key={`${summary.reportName}-dup-${item.descriptor}`} style={{ marginBottom: 6 }}>
                    重测候选: {item.descriptor}，出现 {item.count} 次
                  </Paragraph>
                ))}
                {summary.extraCandidateItems?.slice(0, 4).map((item) => (
                  <Paragraph key={`${summary.reportName}-extra-${item.descriptor}`} style={{ marginBottom: 6 }}>
                    多测候选: {item.descriptor}，出现 {item.count} 次
                  </Paragraph>
                ))}
                {summary.notes?.map((note) => (
                  <Paragraph key={`${summary.reportName}-${note}`} style={{ marginBottom: 6 }}>{note}</Paragraph>
                ))}
              </div>
            ))}
          </div>
        </div>
      )
    });
  };

  const showWordAuditSummary = (conclusionData) => {
    Modal.info({
      title: 'Word 曲线与文档审查',
      width: 940,
      content: (
        <div style={{ marginTop: 16 }}>
          <Paragraph>
            <Tag color={getStatusMeta(conclusionData.wordAudit.status).color}>{getStatusMeta(conclusionData.wordAudit.status).label}</Tag>
            共 {conclusionData.wordAudit.reportCount} 份 Word 报告，文档审查项 {conclusionData.wordAudit.findingCount} 条，识别到响度章节 {conclusionData.wordAudit.loudnessDetectedCount} 份，识别到频响章节 {conclusionData.wordAudit.frequencyDetectedCount} 份。
          </Paragraph>
          <div style={{ maxHeight: 420, overflow: 'auto', paddingRight: 8 }}>
            {conclusionData.wordAudit.reportSummaries.map((summary) => (
              <div key={summary.reportName} style={{ marginBottom: 16, paddingBottom: 12, borderBottom: '1px solid #f0f0f0' }}>
                <Paragraph style={{ marginBottom: 8 }}>
                  <Text strong>{summary.reportName}</Text>
                  {' '}
                  <Tag color={getStatusMeta(summary.documentStatus).color}>{getStatusMeta(summary.documentStatus).label}</Tag>
                </Paragraph>
                {summary.findings?.map((finding) => (
                  <Paragraph key={`${summary.reportName}-${finding.id}`} style={{ marginBottom: 6 }}>
                    <Tag color={getStatusMeta(finding.status).color}>{getStatusMeta(finding.status).label}</Tag>
                    {finding.title} - {finding.message}
                  </Paragraph>
                ))}
                {[summary.loudness, summary.frequencyResponse].filter(Boolean).map((review) => (
                  <Paragraph key={`${summary.reportName}-${review.title}`} style={{ marginBottom: 6 }}>
                    <Tag color={getStatusMeta(review.status).color}>{getStatusMeta(review.status).label}</Tag>
                    {review.title} - {review.message}
                  </Paragraph>
                ))}
              </div>
            ))}
          </div>
        </div>
      )
    });
  };

  const showConsistencySummary = (conclusionData) => {
    Modal.info({
      title: '跨报告一致性审查',
      width: 940,
      content: (
        <div style={{ marginTop: 16 }}>
          <Paragraph>
            <Tag color={getStatusMeta(conclusionData.consistency.status).color}>{getStatusMeta(conclusionData.consistency.status).label}</Tag>
            {conclusionData.consistency.enabled
              ? `共识别 ${conclusionData.consistency.groupCount} 组可比样本，存在 ${conclusionData.consistency.flaggedCount} 个差异项。`
              : '当前样本尚未形成可执行的一致性对比组。'}
          </Paragraph>
          <div style={{ maxHeight: 420, overflow: 'auto', paddingRight: 8 }}>
            {conclusionData.consistency.groups.map((group) => (
              <div key={`${group.comparisonType}-${group.groupKey}`} style={{ marginBottom: 16, paddingBottom: 12, borderBottom: '1px solid #f0f0f0' }}>
                <Paragraph style={{ marginBottom: 8 }}>
                  <Text strong>{group.comparisonType === 'same-codec-cross-network' ? '同 codec 跨 network' : '同 network 跨 codec'}</Text>
                  {' '}
                  {group.groupKey}
                  {' '}
                  <Tag color={getStatusMeta(group.status).color}>{getStatusMeta(group.status).label}</Tag>
                </Paragraph>
                <Paragraph style={{ marginBottom: 8 }}>
                  对比报告：{group.reports.map((item) => item.reportName).join(' / ')}
                </Paragraph>
                {group.flaggedItems?.length ? group.flaggedItems.map((item) => (
                  <Paragraph key={`${group.groupKey}-${item.outputCell}`} style={{ marginBottom: 6 }}>
                    <Tag color={getStatusMeta(item.severity).color}>{getStatusMeta(item.severity).label}</Tag>
                    {item.outputCell} - {item.checklistDesc}，{item.reason}
                    {typeof item.spread === 'number' ? `，差值 ${item.spread}` : ''}
                  </Paragraph>
                )) : (
                  <Paragraph style={{ marginBottom: 6 }}>当前组没有触发明显差异项。</Paragraph>
                )}
              </div>
            ))}
          </div>
        </div>
      )
    });
  };

  const showBundleSummary = (bundle) => {
    Modal.info({
      title: `报告包：${bundle.key}`,
      width: 940,
      content: (
        <div style={{ marginTop: 16 }}>
          <Paragraph>
            数据源：{bundle.sourceMode === 'excel+word' ? 'Excel + Word' : bundle.sourceMode === 'excel' ? '仅 Excel' : '仅 Word'}；
            输出文件：{bundle.hasChecklistOutput ? '已生成' : '未生成'}
          </Paragraph>
          <Paragraph>
            识别上下文：{[bundle.context.codec, bundle.context.network, bundle.context.bandwidth, bundle.context.terminalMode].filter(Boolean).join(' / ') || '未识别'}
          </Paragraph>
          <Paragraph>
            客户：{bundle.context.customer || selectedCustomer || '未指定'}
          </Paragraph>
          <Paragraph>
            Report 参数：
            {bundle.context.reportPanelSelections
              ? ` B13=${bundle.context.reportPanelSelections.B13 || '-'} / B15=${bundle.context.reportPanelSelections.B15 || '-'} / C15=${bundle.context.reportPanelSelections.C15 || '-'} / D15=${bundle.context.reportPanelSelections.D15 || '-'}`
              : ' 未指定'}
          </Paragraph>
          <Paragraph>
            Excel 覆盖性：{bundle.excelCoverage.missingCount} 个漏测，{bundle.excelCoverage.duplicateCount} 组重测候选，{bundle.excelCoverage.extraCandidateCount} 组多测候选。
          </Paragraph>
          <Paragraph>
            Word 审查：{bundle.wordAudit.findingCount} 条文档审查项，响度{bundle.wordAudit.loudnessDetected ? '已识别' : '未识别'}，频响{bundle.wordAudit.frequencyDetected ? '已识别' : '未识别'}。
          </Paragraph>
          <div style={{ maxHeight: 360, overflow: 'auto', paddingRight: 8 }}>
            {bundle.items.map((item) => (
              <div key={item.reportPath} style={{ marginBottom: 14, paddingBottom: 12, borderBottom: '1px solid #f0f0f0' }}>
                <Paragraph style={{ marginBottom: 8 }}>
                  <Text strong>{getOutputFileName(item.reportPath)}</Text>
                  {' '}
                  <Tag color={item.reportKind === 'word' ? 'purple' : 'blue'}>{item.reportKind === 'word' ? 'Word' : 'Excel'}</Tag>
                </Paragraph>
                {item.audit?.coverage ? (
                  <Paragraph style={{ marginBottom: 6 }}>
                    覆盖性状态：<Tag color={getStatusMeta(item.audit.coverage.status).color}>{getStatusMeta(item.audit.coverage.status).label}</Tag>
                    漏测 {item.audit.coverage.missingCount}，重测候选 {item.audit.coverage.duplicateCount}，多测候选 {item.audit.coverage.extraCandidateCount}
                  </Paragraph>
                ) : null}
                {item.audit?.documentCompleteness ? (
                  <Paragraph style={{ marginBottom: 6 }}>
                    文档状态：<Tag color={getStatusMeta(item.audit.documentCompleteness.status).color}>{getStatusMeta(item.audit.documentCompleteness.status).label}</Tag>
                    审查项 {item.audit.documentCompleteness.findings.length}
                  </Paragraph>
                ) : null}
              </div>
            ))}
          </div>
        </div>
      )
    });
  };

  const processReports = async () => {
    if (files.length === 0) {
      message.warning('请先上传至少一个测试报告。');
      return;
    }

    if (files.some((item) => item.reportKind === 'excel') && !checklistFile?.path) {
      message.warning('存在 Excel 报告时，请先上传 checklist Excel 文件。');
      return;
    }

    const runId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    activeRunIdRef.current = runId;
    setProcessedConclusion(null);
    setProcessing(true);
    setFiles((prev) => prev.map((item) => ({ ...item, status: 'processing', error: '' })));
    setProgressState({
      active: true,
      total: files.length,
      completed: 0,
      successCount: 0,
      errorCount: 0,
      currentReportName: ''
    });

    try {
      const response = await window.electron.reportChecker.processReports({
        runId,
        reportPaths: files.map((item) => item.path),
        checklistPath: checklistFile?.path || null,
        rulePath: ruleFile?.path || null,
        customer: selectedCustomer,
        reportPanelSelections
      });

      const resultMap = new Map(response.results.map((item) => [item.reportPath, item]));
      recordReportCheckResults(response.results);
      setProcessedConclusion(response.conclusion || null);

      setFiles((prev) => prev.map((item) => {
        const result = resultMap.get(item.path);
        if (!result) {
          return item;
        }

        if (result.status === 'error') {
          return {
            ...item,
            status: 'error',
            error: result.error,
            items: 0,
            outputPath: '',
            outputName: '',
            skippedItems: [],
            unmatchedItems: [],
            audit: null
          };
        }

        return {
          ...item,
          status: 'success',
          items: result.matchedItems,
          ruleProfileKey: result.ruleProfileKey || item.ruleProfileKey || '',
          outputPath: result.outputPath,
          outputName: getOutputFileName(result.outputPath),
          skippedItems: result.skippedItems || [],
          unmatchedItems: result.unmatchedItems || [],
          reportContext: result.reportContext || item.reportContext,
          audit: result.audit || null,
          error: ''
        };
      }));

      const successCount = response.results.filter((item) => item.status === 'success').length;
      const errorCount = response.results.length - successCount;
      setProgressState((prev) => ({
        ...prev,
        active: false,
        total: response.results.length,
        completed: response.results.length,
        successCount,
        errorCount
      }));
      message.success(`处理完成：成功 ${successCount} 份，失败 ${errorCount} 份。`);
    } catch (error) {
      const errorMessage = error?.message || '执行报告检查失败';
      setFiles((prev) => prev.map((item) => ({ ...item, status: 'error', error: errorMessage })));
      setProgressState((prev) => ({
        ...prev,
        active: false,
        errorCount: prev.total || prev.completed ? Math.max(prev.errorCount, prev.total - prev.completed) : prev.errorCount
      }));
      message.error(errorMessage);
    } finally {
      activeRunIdRef.current = null;
      setProcessing(false);
    }
  };

  const progressPercent = progressState.total > 0
    ? Math.min(100, Math.round((progressState.completed / progressState.total) * 100))
    : 0;
  const uploadSummary = useMemo(() => buildUploadSummary(files, checklistFile), [files, checklistFile]);
  const conclusionData = useMemo(() => buildConclusionData(uploadSummary, processedConclusion), [uploadSummary, processedConclusion]);

  const exportRules = async () => {
    setExportingRules(true);

    try {
      const result = await window.electron.reportChecker.exportRules(ruleFile?.path || null);
      if (result?.canceled) {
        return;
      }

      message.success(`规则已导出到: ${result.filePath}`);
    } catch (error) {
      message.error(error?.message || '导出规则失败');
    } finally {
      setExportingRules(false);
    }
  };

  const columns = [
    {
      title: '文件名',
      dataIndex: 'name',
      key: 'name',
      width: 320,
      ellipsis: true,
      render: (text) => (
        <Text className="report-checker-table-text" ellipsis={{ tooltip: text }}>
          {text}
        </Text>
      )
    },
    {
      title: '状态',
      dataIndex: 'status',
      key: 'status',
      width: 100,
      render: (status) => {
        const colors = {
          success: 'green',
          error: 'red',
          pending: 'blue',
          processing: 'gold'
        };
        return <Tag color={colors[status] || 'default'}>{getStatusMeta(status).label}</Tag>;
      }
    },
    {
      title: '检查项',
      dataIndex: 'items',
      key: 'items',
      width: 96,
      render: (items) => <span>{items || 0}</span>
    },
    {
      title: '输出文件名',
      dataIndex: 'outputName',
      key: 'outputName',
      width: 420,
      ellipsis: true,
      render: (_, record) => {
        if (!record.outputName) {
          return <Text type="secondary">{record.reportKind === 'word' ? 'Word 审查不生成输出' : '处理完成后显示'}</Text>;
        }

        return (
          <Text className="report-checker-table-text" ellipsis={{ tooltip: record.outputName }}>
            {record.outputName}
          </Text>
        );
      }
    },
    {
      title: '操作',
      key: 'action',
      width: 220,
      render: (_, record) => (
        <Space wrap={false}>
          <Button type="primary" size="small" onClick={() => showDetails(record)}>详情</Button>
          <Button size="small" disabled={!record.outputPath} onClick={() => openOutputFolder(record)}>打开目录</Button>
          <Button danger size="small" icon={<DeleteOutlined />} onClick={() => removeReport(record.id)}>删除</Button>
        </Space>
      )
    }
  ];

  return (
    <div className="page-container">
      <Card className="report-checker-card report-checker-section-card report-checker-guide-card" title="使用说明">
        <div className="report-checker-note-list">
          <Paragraph style={{ marginBottom: 10 }}>先上传报告、checklist 和可选规则，再执行检查，最后在结论窗口查看覆盖性、文档审查和一致性结果。</Paragraph>
          <Paragraph style={{ marginBottom: 10 }}>Excel 负责 checklist 填表与覆盖性评估，Word 不再生成 checklist 输出。</Paragraph>
          <Paragraph style={{ marginBottom: 10 }}>跨网络或跨 codec 的一致性审查只在存在可比样本时触发。</Paragraph>
          <Paragraph style={{ marginBottom: 0 }}>响度、频响和 Word 文档审查会输出证据入口，但最终正确性仍需音频工程师确认。</Paragraph>
        </div>
      </Card>

      <div className="report-checker-top-grid">
        <div className="report-checker-top-card">
          <span className="report-checker-top-label">报告总数</span>
          <strong>{uploadSummary.totalReports}</strong>
        </div>
        <div className="report-checker-top-card">
          <span className="report-checker-top-label">已上传 Excel</span>
          <strong>{uploadSummary.excelCount}</strong>
        </div>
        <div className="report-checker-top-card">
          <span className="report-checker-top-label">已上传 Word</span>
          <strong>{uploadSummary.wordCount}</strong>
        </div>
        <div className="report-checker-top-card">
          <span className="report-checker-top-label">已上传 checklist</span>
          <strong>{uploadSummary.checklistCount}</strong>
        </div>
      </div>

      <div className="report-checker-upload-stack">
        <Card 
          className="report-checker-card report-checker-main-card"
          title="上传测试报告"
          extra={
            <div className="report-checker-actions">
              <Button
                danger
                icon={<DeleteOutlined />}
                className="report-checker-clear-action report-checker-section-action"
                disabled={processing || files.length === 0}
                onClick={clearReports}
              >
                清空列表
              </Button>
              <Upload
                customRequest={({ file, onSuccess }) => handleUpload(file, 'report', onSuccess)}
                multiple
                accept=".xlsx,.xls,.doc,.docx"
                showUploadList={false}
              >
                <Button icon={<UploadOutlined />} className="report-checker-upload-action report-checker-section-action">
                  上传报告
                </Button>
              </Upload>
            </div>
          }
        >
          <div className="report-checker-upload-summary report-checker-upload-summary-expanded">
            <div className="report-checker-upload-summary-item">
              <span className="report-checker-upload-summary-label">测试报告总数</span>
              <span className="report-checker-upload-summary-value">{uploadSummary.totalReports}</span>
            </div>
            <div className="report-checker-upload-summary-item">
              <span className="report-checker-upload-summary-label">Excel / Word / checklist</span>
              <span className="report-checker-upload-summary-text">
                {uploadSummary.excelCount} / {uploadSummary.wordCount} / {uploadSummary.checklistCount}
              </span>
            </div>
            <div className="report-checker-upload-summary-item">
              <span className="report-checker-upload-summary-label">当前状态</span>
              <span className="report-checker-upload-summary-text">
                {uploadSummary.totalReports > 0
                  ? `已选择 ${uploadSummary.totalReports} 份报告，可直接执行 Excel 填表、Word 审查和结论汇总。`
                  : '还没有选择测试报告'}
              </span>
            </div>
          </div>

          <p style={{ marginBottom: '24px', color: '#8c8c8c' }}>
            Excel 报告仅用于 checklist 填表、漏测/重测/多测候选评估与跨报告一致性检查；Word 报告仅用于响度/频响章节识别和文档完整性审查。
          </p>
          <p style={{ marginTop: '-12px', marginBottom: '24px', color: '#8c8c8c' }}>
            .doc 报告会先在后台尝试转成 .docx；Word 路径不会生成 checklist 输出文件，结果统一进入下方结论窗口。
          </p>

          {files.length === 0 ? (
            <Upload.Dragger
              customRequest={({ file, onSuccess }) => handleUpload(file, 'report', onSuccess)}
              multiple
              accept=".xlsx,.xls,.doc,.docx"
              showUploadList={false}
              className="report-checker-upload report-checker-upload-report"
              style={{ padding: '16px 18px', minHeight: '108px' }}
            >
              <UploadOutlined style={{ fontSize: '36px', color: '#bfbfbf', marginBottom: '10px' }} />
              <p style={{ fontSize: '15px', color: '#262626', marginBottom: '6px' }}>
                拖拽报告文件到此处，或点击上方按钮上传
              </p>
              <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
                当前支持格式: Excel (.xlsx, .xls) 优先，兼容 Word (.doc, .docx)
              </p>
            </Upload.Dragger>
          ) : (
            <Table
              className="report-checker-table report-checker-report-table"
              columns={columns}
              dataSource={files}
              rowKey="id"
              scroll={{ x: 1280 }}
              pagination={{ pageSize: 10 }}
            />
          )}
        </Card>

        <div className="report-checker-aux-grid">
          <Card
            className="report-checker-card report-checker-section-card"
            title="上传 checklist"
            extra={
              <Upload
                customRequest={({ file, onSuccess }) => handleUpload(file, 'checklist', onSuccess)}
                accept=".xlsx,.xls"
                showUploadList={false}
              >
                <Button icon={<UploadOutlined />} className="report-checker-upload-action report-checker-section-action">
                  上传 checklist
                </Button>
              </Upload>
            }
          >
            <Upload.Dragger
              customRequest={({ file, onSuccess }) => handleUpload(file, 'checklist', onSuccess)}
              accept=".xlsx,.xls"
              showUploadList={false}
              className="report-checker-upload"
              style={compactUploadDraggerStyle}
            >
              {checklistFile ? (
                <div style={{ padding: '8px 0' }}>
                  <Text strong style={{ display: 'block', fontSize: '16px', color: '#262626', marginBottom: '6px' }}>
                    已选择 checklist
                  </Text>
                  <Text style={{ fontSize: '14px', color: '#595959', wordBreak: 'break-all' }}>
                    {checklistFile.name}
                  </Text>
                </div>
              ) : (
                <>
                  <UploadOutlined style={{ fontSize: '24px', color: '#bfbfbf', marginBottom: '6px' }} />
                  <p style={{ fontSize: '14px', color: '#262626', marginBottom: '3px' }}>
                    拖拽 checklist 文件到此处，或点击上传
                  </p>
                  <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
                    支持格式: Excel (.xlsx, .xls)
                  </p>
                </>
              )}
            </Upload.Dragger>
          </Card>

          <Card
            className="report-checker-card report-checker-section-card"
            title="上传规则"
            extra={
              <Space>
                <Button
                  icon={<ExportOutlined />}
                  className="report-checker-upload-action report-checker-section-action"
                  loading={exportingRules}
                  onClick={exportRules}
                >
                  导出规则
                </Button>
                <Upload
                  customRequest={({ file, onSuccess }) => handleUpload(file, 'rules', onSuccess)}
                  accept=".json,.json5"
                  showUploadList={false}
                >
                  <Button icon={<UploadOutlined />} className="report-checker-upload-action report-checker-section-action">
                    上传规则
                  </Button>
                </Upload>
              </Space>
            }
          >
            <Upload.Dragger
              customRequest={({ file, onSuccess }) => handleUpload(file, 'rules', onSuccess)}
              accept=".json,.json5"
              showUploadList={false}
              className="report-checker-upload"
              style={compactUploadDraggerStyle}
            >
              <UploadOutlined style={{ fontSize: '24px', color: '#bfbfbf', marginBottom: '6px' }} />
              <p style={{ fontSize: '14px', color: '#262626', marginBottom: '3px' }}>
                拖拽规则文件到此处，或点击上传
              </p>
              <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
                支持格式: JSON / JSON5；不上传时默认使用内置规则
              </p>
            </Upload.Dragger>
            {ruleFile && (
              <p style={{ marginTop: '12px', color: '#595959' }}>
                已选择: {ruleFile.name}
              </p>
            )}
          </Card>
        </div>
      </div>

      <Card className="report-checker-card report-checker-section-card" title="任务参数">
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(240px, 1fr))', gap: 16 }}>
          <div>
            <Text strong style={{ display: 'block', marginBottom: 8 }}>客户</Text>
            <Select
              value={selectedCustomer}
              style={{ width: '100%' }}
              options={CUSTOMER_OPTIONS.map((item) => ({ label: item, value: item }))}
              onChange={setSelectedCustomer}
            />
          </div>
          <div>
            <Text strong style={{ display: 'block', marginBottom: 8 }}>
              Report 参数来源
            </Text>
            <Text type="secondary">
              {checklistFile?.name
                ? `${checklistFile.name} / ${reportPanelMeta.reportSheetName || 'Report'} 页`
                : '先上传 checklist 后读取参数'}
            </Text>
          </div>
          {REPORT_PANEL_FIELDS.map((field) => {
            const panelField = reportPanelMeta.fields.find((item) => item.cell === field.cell);
            const options = reportPanelDynamicOptions[field.cell] || panelField?.options || [];
            const currentValue = reportPanelSelections[field.cell] || panelField?.currentValue || '';

            return (
              <div key={field.cell}>
                <Text strong style={{ display: 'block', marginBottom: 8 }}>
                  {field.label} ({field.cell})
                </Text>
                <Select
                  showSearch
                  allowClear
                  value={currentValue || undefined}
                  style={{ width: '100%' }}
                  placeholder="请选择参数值"
                  onChange={(value) => updateReportPanelSelection(field.cell, value)}
                  options={options.map((item) => ({ label: item, value: item }))}
                  notFoundContent={options.length === 0 ? '暂无候选值' : null}
                />
              </div>
            );
          })}
        </div>
        <Paragraph type="secondary" style={{ marginTop: 16, marginBottom: 0 }}>
          客户与 Report 参数会随本次任务提交后端，用于规则分发、动态工作表切换与 Report 页回写。
        </Paragraph>
      </Card>

      <Card className="report-checker-card report-checker-action-card">
        <div className="report-checker-action-panel">
          <div>
            <div className="report-checker-action-title">开始检查并输出结论</div>
            <div className="report-checker-action-text">上传完成后，从这里启动处理。结果会先更新进度，再在下方结论窗口集中展示。</div>
          </div>
          <Button
            type="primary"
            icon={<CheckCircleOutlined />}
            loading={processing}
            onClick={processReports}
            className="report-checker-primary-action"
          >
            开始检查并输出结论
          </Button>
        </div>
      </Card>

      {(processing || progressState.completed > 0) && (
        <Card className="report-checker-card report-checker-progress-card">
          <div className="report-checker-progress-header">
            <div>
              <div className="report-checker-progress-title">批量处理进度</div>
              <div className="report-checker-progress-subtitle">
                {progressState.active
                  ? `正在处理 ${progressState.completed + 1 <= progressState.total ? progressState.completed + 1 : progressState.total}/${progressState.total} 份报告`
                  : `已完成 ${progressState.completed}/${progressState.total} 份报告`}
              </div>
            </div>
            <div className="report-checker-progress-stats">
              <span>成功 {progressState.successCount}</span>
              <span>失败 {progressState.errorCount}</span>
            </div>
          </div>
          <Progress percent={progressPercent} status={progressState.active ? 'active' : progressState.errorCount > 0 ? 'exception' : 'success'} />
          <div className="report-checker-progress-footer">
            <span>
              {progressState.currentReportName ? `最近完成: ${progressState.currentReportName}` : '等待后台返回首个结果...'}
            </span>
            <span>{progressState.completed}/{progressState.total}</span>
          </div>
        </Card>
      )}

      <Card className="report-checker-card report-checker-conclusion-card" title="结论输出">
        <div className="report-checker-conclusion-actions" style={{ marginBottom: 16 }}>
          <Text strong>本次运行参数</Text>
          <div className="report-checker-conclusion-action-list">
            <Paragraph style={{ marginBottom: 8 }}>
              客户：{conclusionData.runConfig?.customer || selectedCustomer || '未指定'}
            </Paragraph>
            <Paragraph style={{ marginBottom: 8 }}>
              Report 参数：
              {conclusionData.runConfig?.reportPanelSelections
                ? ` B13=${conclusionData.runConfig.reportPanelSelections.B13 || '-'} / B15=${conclusionData.runConfig.reportPanelSelections.B15 || '-'} / C15=${conclusionData.runConfig.reportPanelSelections.C15 || '-'} / D15=${conclusionData.runConfig.reportPanelSelections.D15 || '-'}`
                : ' 未指定'}
            </Paragraph>
            <Paragraph style={{ marginBottom: 0 }}>
              生效规则 Profile：{Array.isArray(conclusionData.runConfig?.ruleProfiles) && conclusionData.runConfig.ruleProfiles.length > 0
                ? conclusionData.runConfig.ruleProfiles.join(' / ')
                : '未识别'}
            </Paragraph>
          </div>
        </div>

        <div className="report-checker-conclusion-grid">
          <div className="report-checker-conclusion-metric">
            <span className="report-checker-conclusion-label">报告总数</span>
            <strong>{conclusionData.overview.totalReports}</strong>
          </div>
          <div className="report-checker-conclusion-metric">
            <span className="report-checker-conclusion-label">已生成输出</span>
            <strong>{conclusionData.overview.outputCount}</strong>
          </div>
          <div className="report-checker-conclusion-metric">
            <span className="report-checker-conclusion-label">成功处理</span>
            <strong>{conclusionData.overview.successCount}</strong>
          </div>
          <div className="report-checker-conclusion-metric">
            <span className="report-checker-conclusion-label">失败报告</span>
            <strong>{conclusionData.overview.errorCount}</strong>
          </div>
        </div>

        <div className="report-checker-insight-grid">
          <button type="button" className="report-checker-insight-card" onClick={() => showExcelCoverageSummary(conclusionData)}>
            <div className="report-checker-insight-header">
              <span className="report-checker-insight-title">Excel 填表与覆盖性</span>
              <Tag color={getStatusMeta(conclusionData.excelCoverage.status).color}>{getStatusMeta(conclusionData.excelCoverage.status).label}</Tag>
            </div>
            <strong>{conclusionData.excelCoverage.reportCount}</strong>
            <span className="report-checker-insight-text">
              命中 {conclusionData.excelCoverage.matchedCount}，漏测 {conclusionData.excelCoverage.missingCount}，跳过 {conclusionData.excelCoverage.skippedCount}，重测候选 {conclusionData.excelCoverage.duplicateCount}
            </span>
          </button>

          <button type="button" className="report-checker-insight-card" onClick={() => showWordAuditSummary(conclusionData)}>
            <div className="report-checker-insight-header">
              <span className="report-checker-insight-title">Word 曲线与文档审查</span>
              <Tag color={getStatusMeta(conclusionData.wordAudit.status).color}>{getStatusMeta(conclusionData.wordAudit.status).label}</Tag>
            </div>
            <strong>{conclusionData.wordAudit.reportCount}</strong>
            <span className="report-checker-insight-text">
              审查项 {conclusionData.wordAudit.findingCount}，响度章节 {conclusionData.wordAudit.loudnessDetectedCount}，频响章节 {conclusionData.wordAudit.frequencyDetectedCount}
            </span>
          </button>

          <button type="button" className="report-checker-insight-card" onClick={() => showConsistencySummary(conclusionData)}>
            <div className="report-checker-insight-header">
              <span className="report-checker-insight-title">跨报告一致性</span>
              <Tag color={getStatusMeta(conclusionData.consistency.status).color}>{getStatusMeta(conclusionData.consistency.status).label}</Tag>
            </div>
            <strong>{conclusionData.consistency.enabled ? conclusionData.consistency.groupCount : 0}</strong>
            <span className="report-checker-insight-text">
              {conclusionData.consistency.enabled
                ? `可比组 ${conclusionData.consistency.groupCount}，差异项 ${conclusionData.consistency.flaggedCount}`
                : '当前样本尚未形成跨报告可比组'}
            </span>
          </button>

          <button type="button" className="report-checker-insight-card" onClick={() => Modal.info({ title: '当前建议', width: 760, content: (<div style={{ marginTop: 16 }}>{conclusionData.suggestedActions.length > 0 ? conclusionData.suggestedActions.map((item) => <Paragraph key={item} style={{ marginBottom: 8 }}>{item}</Paragraph>) : <Paragraph style={{ marginBottom: 0 }}>当前没有新增建议。</Paragraph>}</div>) })}>
            <div className="report-checker-insight-header">
              <span className="report-checker-insight-title">人工复核建议</span>
              <Tag color={conclusionData.suggestedActions.length > 0 ? 'gold' : 'green'}>{conclusionData.suggestedActions.length > 0 ? '关注' : '稳定'}</Tag>
            </div>
            <strong>{conclusionData.suggestedActions.length}</strong>
            <span className="report-checker-insight-text">
              点击查看当前批次最值得优先处理的结论与补充动作。
            </span>
          </button>
        </div>

        {conclusionData.bundles.length > 0 ? (
          <div className="report-checker-bundle-list">
            {conclusionData.bundles.map((bundle) => (
              <button type="button" key={bundle.key} className="report-checker-bundle-item" onClick={() => showBundleSummary(bundle)}>
                <div className="report-checker-bundle-header">
                  <div>
                    <div className="report-checker-bundle-title">{bundle.key}</div>
                    <div className="report-checker-bundle-meta">
                      <span>{bundle.context.customer || '未知客户'}</span>
                      <span>{bundle.context.codec || '未知 codec'}</span>
                      <span>{bundle.context.network || '未知 network'}</span>
                      <span>{bundle.context.bandwidth || '未知 bandwidth'}</span>
                      <span>{bundle.context.terminalMode || '未知 mode'}</span>
                    </div>
                  </div>
                  <div className="report-checker-bundle-tags">
                    <Tag color={bundle.sourceMode === 'excel+word' ? 'green' : bundle.sourceMode === 'excel' ? 'blue' : bundle.sourceMode === 'word' ? 'purple' : 'default'}>
                      {bundle.sourceMode === 'excel+word' ? 'Excel + Word 联合' : bundle.sourceMode === 'excel' ? '仅 Excel' : bundle.sourceMode === 'word' ? '仅 Word' : '来源未识别'}
                    </Tag>
                    <Tag color={getStatusMeta(bundle.excelCoverage.status).color}>覆盖性 {getStatusMeta(bundle.excelCoverage.status).label}</Tag>
                    <Tag color={getStatusMeta(bundle.wordAudit.status).color}>文档审查 {getStatusMeta(bundle.wordAudit.status).label}</Tag>
                  </div>
                </div>
                <div className="report-checker-bundle-stats">
                  <span>Excel {bundle.excelCount}</span>
                  <span>Word {bundle.wordCount}</span>
                  <span>漏测 {bundle.excelCoverage.missingCount}</span>
                  <span>文档项 {bundle.wordAudit.findingCount}</span>
                  <span>{bundle.hasChecklistOutput ? '已生成输出' : '未生成输出'}</span>
                </div>
              </button>
            ))}
          </div>
        ) : (
          <Paragraph type="secondary" style={{ marginTop: 16, marginBottom: 0 }}>
            执行检查后，这里会按 Excel 填表、Word 审查和跨报告一致性输出完整结论。
          </Paragraph>
        )}

        {conclusionData.suggestedActions.length > 0 ? (
          <div className="report-checker-conclusion-actions">
            <Text strong>当前建议</Text>
            <div className="report-checker-conclusion-action-list">
              {conclusionData.suggestedActions.map((item) => (
                <Paragraph key={item} style={{ marginBottom: 8 }}>
                  {item}
                </Paragraph>
              ))}
            </div>
          </div>
        ) : null}
      </Card>
    </div>
  );
}

export default ReportChecker;
