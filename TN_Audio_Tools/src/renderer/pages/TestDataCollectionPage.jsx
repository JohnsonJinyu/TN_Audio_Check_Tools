import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Alert, App as AntdApp, Button, Card, Progress, Select, Space, Table, Tag, Typography, Upload } from 'antd';
import { UploadOutlined, DeleteOutlined, CheckCircleOutlined, ExportOutlined } from '@ant-design/icons';
import { recordDataCollectionResults } from '../modules/dashboard/storage';
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
const REPORT_PANEL_FIELD_SHORT_LABELS = {
  B13: '接口',
  B15: '网络',
  C15: '编码',
  D15: '码率'
};

const PROJECT_PHASE_OPTIONS = ['EVB', 'EVT', 'DVT1', 'DVT2', 'PVT'];
const EMPTY_REPORT_PANEL_SELECTIONS = {
  B13: '',
  B15: '',
  C15: '',
  D15: ''
};

const KNOWN_RULE_PROFILE_LABELS = {
  handset: 'Handset',
  handsfree: 'Handsfree',
  headset: 'Headset',
  electrical_interface: 'Electrical Interface - Analogue',
  electrical_interface_digital: 'Electrical Interface - Digital'
};
const CHECKLIST_TEMPLATE_OPTIONS = [
  { label: 'Handset', value: 'handset' },
  { label: 'Handsfree', value: 'handsfree' },
  { label: 'Headset', value: 'headset' },
  { label: 'Electrical Interface', value: 'electrical_interface' }
];
const RULE_PROFILE_TO_CHECKLIST_TEMPLATE = {
  handset: 'handset',
  handsfree: 'handsfree',
  headset: 'headset',
  electrical_interface: 'electrical_interface',
  electrical_interface_digital: 'electrical_interface'
};

function createEmptyReportPanelSelections() {
  return { ...EMPTY_REPORT_PANEL_SELECTIONS };
}

function normalizePanelSelections(selections = {}) {
  return {
    B13: String(selections.B13 || '').trim(),
    B15: String(selections.B15 || '').trim(),
    C15: String(selections.C15 || '').trim(),
    D15: String(selections.D15 || '').trim()
  };
}

function arePanelSelectionsEmpty(selections = {}) {
  const normalized = normalizePanelSelections(selections);
  return !normalized.B13 && !normalized.B15 && !normalized.C15 && !normalized.D15;
}

function panelSelectionsEqual(left = {}, right = {}) {
  const normalizedLeft = normalizePanelSelections(left);
  const normalizedRight = normalizePanelSelections(right);
  return REPORT_PANEL_FIELDS.every((field) => normalizedLeft[field.cell] === normalizedRight[field.cell]);
}

function dynamicOptionsEqual(left = {}, right = {}) {
  const keys = Array.from(new Set([...Object.keys(left || {}), ...Object.keys(right || {})]));
  return keys.every((key) => {
    const leftValues = Array.isArray(left?.[key]) ? left[key] : [];
    const rightValues = Array.isArray(right?.[key]) ? right[key] : [];
    return leftValues.length === rightValues.length && leftValues.every((value, index) => value === rightValues[index]);
  });
}

function getPanelField(reportPanelMeta, cell) {
  return reportPanelMeta.fields.find((field) => field.cell === cell);
}

function getOptionsForCell(reportPanelMeta, cell, selections = {}) {
  const field = getPanelField(reportPanelMeta, cell);
  if (!field) {
    return [];
  }

  if (cell === 'C15') {
    return (field.cascadeMap || {})[String(selections.B15 || '').trim()] || field.options || [];
  }

  if (cell === 'D15') {
    return (field.cascadeMap || {})[String(selections.C15 || '').trim()] || field.options || [];
  }

  return field.options || [];
}

function buildPanelState(reportPanelMeta, requestedSelections = {}) {
  const nextSelections = normalizePanelSelections(requestedSelections);
  const nextDynamicOptions = {};

  const c15Options = getOptionsForCell(reportPanelMeta, 'C15', nextSelections);
  if (c15Options.length > 0) {
    nextDynamicOptions.C15 = c15Options;
  }
  if (nextSelections.C15 && c15Options.length > 0 && !c15Options.includes(nextSelections.C15)) {
    nextSelections.C15 = '';
  }

  const d15Options = getOptionsForCell(reportPanelMeta, 'D15', nextSelections);
  if (d15Options.length > 0) {
    nextDynamicOptions.D15 = d15Options;
  }
  if (nextSelections.D15 && d15Options.length > 0 && !d15Options.includes(nextSelections.D15)) {
    nextSelections.D15 = '';
  }

  return {
    selections: nextSelections,
    dynamicOptions: nextDynamicOptions
  };
}

function applyManualPanelSelectionChange(reportPanelMeta, currentSelections = {}, cell, value) {
  const nextSelections = normalizePanelSelections(currentSelections);
  const nextDynamicOptions = {};
  nextSelections[cell] = String(value || '').trim();

  if (cell === 'B15') {
    const c15Options = getOptionsForCell(reportPanelMeta, 'C15', nextSelections);
    nextDynamicOptions.C15 = c15Options;
    nextSelections.C15 = c15Options.includes(nextSelections.C15) ? nextSelections.C15 : (c15Options[0] || '');

    const d15Options = getOptionsForCell(reportPanelMeta, 'D15', nextSelections);
    nextDynamicOptions.D15 = d15Options;
    nextSelections.D15 = d15Options.includes(nextSelections.D15) ? nextSelections.D15 : (d15Options[0] || '');
    return {
      selections: nextSelections,
      dynamicOptions: nextDynamicOptions
    };
  }

  if (cell === 'C15') {
    const d15Options = getOptionsForCell(reportPanelMeta, 'D15', nextSelections);
    nextDynamicOptions.D15 = d15Options;
    nextSelections.D15 = d15Options.includes(nextSelections.D15) ? nextSelections.D15 : (d15Options[0] || '');
    return {
      selections: nextSelections,
      dynamicOptions: nextDynamicOptions
    };
  }

  return buildPanelState(reportPanelMeta, nextSelections);
}

function buildDetectedTags(reportContext = {}) {
  return [
    reportContext.network,
    reportContext.codec && reportContext.bandwidth ? `${reportContext.codec}_${reportContext.bandwidth}` : reportContext.codec,
    reportContext.terminalMode
  ].filter(Boolean);
}

function normalizeRuleProfileValue(value = '') {
  return String(value || '').trim().toLowerCase();
}

function formatRuleProfileLabel(value = '') {
  const normalized = normalizeRuleProfileValue(value);
  return KNOWN_RULE_PROFILE_LABELS[normalized] || normalized || '未指定';
}

function mapRuleProfileToChecklistTemplate(value = '') {
  const normalized = normalizeRuleProfileValue(value);
  return RULE_PROFILE_TO_CHECKLIST_TEMPLATE[normalized] || '';
}

function formatChecklistTemplateLabel(value = '') {
  const normalized = normalizeRuleProfileValue(value);
  const matchedOption = CHECKLIST_TEMPLATE_OPTIONS.find((item) => item.value === normalized);
  return matchedOption?.label || normalized || '未指定';
}

function getRuleProfileOptions(record) {
  const profileKeys = Array.from(new Set([
    ...(Array.isArray(record.availableRuleProfiles) ? record.availableRuleProfiles : []),
    normalizeRuleProfileValue(record.suggestedRuleProfile),
    normalizeRuleProfileValue(record.selectedRuleProfile),
    normalizeRuleProfileValue(record.ruleProfileKey)
  ].filter(Boolean)));

  return profileKeys.map((value) => ({
    label: formatRuleProfileLabel(value),
    value
  }));
}

function getAutoDetectionMeta(record, isMultiExcelMode = false) {
  if (record.contextInspectionStatus === 'pending') {
    return {
      color: 'gold',
      label: '识别中',
      detail: `正在读取${record.reportKind === 'excel' ? '报告内容和参数' : '报告内容'}并生成推荐规则`,
      missingFields: [],
      tags: []
    };
  }

  if (record.contextInspectionStatus === 'error') {
    return {
      color: 'red',
      label: '识别失败',
      detail: record.contextInspectionError || '未能读取稳定上下文，请人工确认',
      missingFields: record.reportKind === 'excel' ? REPORT_PANEL_FIELDS.map((field) => field.label) : [],
      tags: []
    };
  }

  const selections = normalizePanelSelections(record.reportPanelSelections || record.reportContext?.reportPanelSelections || {});
  const missingFields = REPORT_PANEL_FIELDS
    .filter((field) => !selections[field.cell])
    .map((field) => field.label);
  const detectedTags = buildDetectedTags(record.reportContext || {});

  if (missingFields.length === 0) {
    return {
      color: 'green',
      label: '已自动识别',
      detail: record.reportKind === 'excel' ? '已填入下方确认面板，可直接复核' : '已生成规则预选，可直接复核',
      missingFields,
      tags: detectedTags
    };
  }

  if (detectedTags.length > 0 || !arePanelSelectionsEmpty(selections)) {
    return {
      color: 'gold',
      label: '待人工确认',
      detail: record.reportKind === 'excel'
        ? `已识别稳定项，缺少 ${missingFields.join(' / ')} 的稳定来源`
        : '已识别部分上下文，但仍需人工确认规则模式',
      missingFields,
      tags: detectedTags
    };
  }

  return {
    color: 'orange',
    label: '未识别',
    detail: '当前未提取到稳定参数，请人工确认',
    missingFields,
    tags: detectedTags
  };
}

function getParameterConfirmationMeta(record) {
  return String(record?.parameterConfirmationStatus || '').trim().toLowerCase() === 'confirmed'
    ? { color: 'green', label: '已确认' }
    : { color: 'gold', label: '待确认' };
}

function getRuleSelectionMeta(record) {
  const selectedRuleProfile = normalizeRuleProfileValue(record.selectedRuleProfile || record.ruleProfileKey);
  const suggestedRuleProfile = normalizeRuleProfileValue(record.suggestedRuleProfile);

  if (!selectedRuleProfile) {
    return {
      color: 'red',
      label: '未选择规则',
      detail: '文件名未能稳定预选规则，请手动指定。'
    };
  }

  if (selectedRuleProfile === suggestedRuleProfile) {
    return {
      color: 'green',
      label: '使用预选规则',
      detail: record.suggestedRuleProfileReason ? `依据 ${record.suggestedRuleProfileReason}` : '已采用系统预选结果。'
    };
  }

  return {
    color: 'gold',
    label: '已手动切换',
    detail: '当前规则以人工选择为准，并会联动 checklist 模式。'
  };
}

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

  // 从文件名提取项目名和阶段
  const phaseMatch = normalizedName.match(/(EVB|EVT|DVT[12]|PVT)/);
  const projectPhase = phaseMatch ? phaseMatch[1] : '';
  let projectName = '';
  if (normalizedName) {
    const firstUnderscore = normalizedName.indexOf('_');
    if (firstUnderscore > 0) {
      projectName = normalizedName.substring(0, firstUnderscore);
    } else {
      const firstWord = normalizedName.split(/[\s_-]+/)[0];
      if (firstWord) projectName = firstWord;
    }
  }

  return {
    measurementObject: normalizedName,
    network,
    codec,
    bandwidth,
    terminalMode,
    projectName,
    projectPhase
  };
}

function buildUploadSummary(files, checklistFile, presetChecklistPath) {
  const excelFiles = files.filter((item) => item.reportKind === 'excel');
  const wordFiles = files.filter((item) => item.reportKind === 'word');
  return {
    totalReports: files.length,
    excelCount: excelFiles.length,
    wordCount: wordFiles.length,
    checklistCount: checklistFile?.path || presetChecklistPath ? 1 : 0
  };
}

function sanitizeOutputToken(value = '') {
  return String(value || '').trim().replace(/\s+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
}

function buildPredictedOutputBaseName(record) {
  const projectName = sanitizeOutputToken(record.projectName || '');
  const projectPhase = sanitizeOutputToken(record.projectPhase || '');
  const selections = normalizePanelSelections(record.reportPanelSelections || record.reportContext?.reportPanelSelections || {});
  const network = sanitizeOutputToken(selections.B15 || record.reportContext?.network || '');
  const vocoder = sanitizeOutputToken(selections.C15 || record.reportContext?.vocoder || [record.reportContext?.codec, record.reportContext?.bandwidth].filter(Boolean).join('_'));

  if (!projectName || !projectPhase) {
    return '';
  }

  const parts = [projectName, projectPhase];
  if (network) parts.push(network);
  if (vocoder) parts.push(vocoder);
  parts.push('checklist');
  return parts.join('_');
}

function findDuplicateOutputConfigs(files) {
  const groups = new Map();
  files.forEach((record) => {
    const baseName = buildPredictedOutputBaseName(record);
    if (!baseName) {
      return;
    }

    const group = groups.get(baseName) || [];
    group.push(record);
    groups.set(baseName, group);
  });

  return Array.from(groups.entries())
    .filter(([, group]) => group.length > 1)
    .map(([baseName, group]) => ({ baseName, group }));
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
    sourcePolicy: {
      status: 'not_applicable',
      preferredSource: 'excel',
      currentMode: uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0 ? 'word_only' : 'none',
      confidence: uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0 ? 'low' : 'not_applicable',
      confidenceLabel: uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0 ? '低' : '未触发',
      manualConfirmationLevel: uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0 ? 'required' : 'none',
      manualConfirmationRequired: uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0,
      summary: uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0
        ? '当前批次仅上传了 Word 报告，正式写入前必须人工确认。'
        : '请优先上传 Excel 报告作为主数据源。',
      detail: uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0
        ? 'Word 提取准确性和稳定性低于 Excel，仅建议作为兼容通道使用。'
        : 'Excel 结构化数据更稳定，适合作为 checklist 自动填表主通道。'
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

function TestDataCollectionPage() {
  const { message, modal } = AntdApp.useApp();
  const [files, setFiles] = useState([]);
  const [ruleFile, setRuleFile] = useState(null);
  const [checklistFile, setChecklistFile] = useState(null);
  const [presetChecklistTemplate, setPresetChecklistTemplate] = useState('');
  const [presetChecklistPath, setPresetChecklistPath] = useState('');
  const [suggestedChecklistTemplate, setSuggestedChecklistTemplate] = useState('');
  const [checklistSelectionSource, setChecklistSelectionSource] = useState('');
  const [checklistConfirmationStatus, setChecklistConfirmationStatus] = useState('pending');
  const [selectedCustomer, setSelectedCustomer] = useState('MOTOROLA');
  const [reportPanelMeta, setReportPanelMeta] = useState({
    reportSheetName: 'Report',
    fields: []
  });
  const [reportPanelSelections, setReportPanelSelections] = useState({
    ...EMPTY_REPORT_PANEL_SELECTIONS
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
    const excelReports = files.filter((item) => item.reportKind === 'excel');
    const suggestedSelections = excelReports.length === 1
      ? normalizePanelSelections(excelReports[0].reportPanelSelections || excelReports[0].reportContext?.reportPanelSelections || {})
      : null;
    const baseSelections = suggestedSelections && arePanelSelectionsEmpty(reportPanelSelections)
      ? suggestedSelections
      : reportPanelSelections;
    const nextPanelState = buildPanelState(reportPanelMeta, baseSelections);

    if (!panelSelectionsEqual(reportPanelSelections, nextPanelState.selections)) {
      setReportPanelSelections(nextPanelState.selections);
    }

    if (!dynamicOptionsEqual(reportPanelDynamicOptions, nextPanelState.dynamicOptions)) {
      setReportPanelDynamicOptions(nextPanelState.dynamicOptions);
    }
  }, [files, reportPanelMeta]);

  useEffect(() => {
    if (!window.electron?.testDataCollection?.onProgress) {
      return undefined;
    }

    const unsubscribe = window.electron.testDataCollection.onProgress((payload) => {
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

  useEffect(() => {
    if (checklistFile?.path) {
      return;
    }

    const candidateTemplates = Array.from(new Set(
      files
        .map((item) => mapRuleProfileToChecklistTemplate(item.selectedRuleProfile || item.suggestedRuleProfile || item.ruleProfileKey))
        .filter(Boolean)
    ));
    const nextSuggestedTemplate = candidateTemplates.length === 1 ? candidateTemplates[0] : '';

    setSuggestedChecklistTemplate((prev) => (prev === nextSuggestedTemplate ? prev : nextSuggestedTemplate));

    if (checklistSelectionSource === 'manual' || checklistSelectionSource === 'upload') {
      return;
    }

    if (!nextSuggestedTemplate) {
      if (checklistSelectionSource === 'auto') {
        setPresetChecklistTemplate('');
        setPresetChecklistPath('');
        setChecklistSelectionSource('');
        setChecklistConfirmationStatus('pending');
        resetChecklistReportPanelOptions();
      }
      return;
    }

    if (presetChecklistTemplate === nextSuggestedTemplate && checklistSelectionSource === 'auto') {
      return;
    }

    setPresetChecklistTemplate(nextSuggestedTemplate);
    setChecklistSelectionSource('auto');
    setChecklistConfirmationStatus('pending');
    resolveAndLoadPresetChecklistTemplate(nextSuggestedTemplate);
  }, [files, checklistFile?.path, checklistSelectionSource, presetChecklistTemplate]);

  const loadChecklistReportPanelOptions = async (checklistPath) => {
    if (!checklistPath || !window.electron?.testDataCollection?.getChecklistReportOptions) {
      return;
    }

    try {
      const result = await window.electron.testDataCollection.getChecklistReportOptions(checklistPath);
      const nextFields = Array.isArray(result?.fields) ? result.fields : [];
      const nextSelections = nextFields.reduce((acc, field) => {
        acc[field.cell] = field.currentValue || '';
        return acc;
      }, createEmptyReportPanelSelections());

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
      setReportPanelSelections(createEmptyReportPanelSelections());
      setReportPanelDynamicOptions({});
      message.warning(error?.message || '读取 checklist Report 参数失败，将使用报告自动识别值。');
    }
  };

  const resetChecklistReportPanelOptions = () => {
    setReportPanelMeta({ reportSheetName: 'Report', fields: [] });
    setReportPanelSelections(createEmptyReportPanelSelections());
    setReportPanelDynamicOptions({});
  };

  const resolveAndLoadPresetChecklistTemplate = async (profileKey, customRulePath = ruleFile?.path || null) => {
    const normalizedProfileKey = normalizeRuleProfileValue(profileKey);
    if (!normalizedProfileKey || !window.electron?.testDataCollection?.resolvePresetChecklistTemplate) {
      setPresetChecklistPath('');
      resetChecklistReportPanelOptions();
      return;
    }

    try {
      const resolved = await window.electron.testDataCollection.resolvePresetChecklistTemplate({
        profileKey: normalizedProfileKey,
        rulePath: customRulePath
      });

      if (!resolved?.templatePath) {
        setPresetChecklistPath('');
        resetChecklistReportPanelOptions();
        return;
      }

      setPresetChecklistPath(resolved.templatePath);
      await loadChecklistReportPanelOptions(resolved.templatePath);
    } catch (error) {
      setPresetChecklistPath('');
      resetChecklistReportPanelOptions();
      message.warning(error?.message || '读取预设模板参数失败。');
    }
  };

  const confirmChecklistSelection = () => {
    if (checklistFile?.path) {
      setChecklistConfirmationStatus('confirmed');
      message.success('已确认当前 checklist 文件。');
      return;
    }

    if (!presetChecklistTemplate || !presetChecklistPath) {
      message.warning('请先选择或等待自动预选 checklist 模板，再确认。');
      return;
    }

    setChecklistConfirmationStatus('confirmed');
    message.success(`已确认预设模板：${formatChecklistTemplateLabel(presetChecklistTemplate)}。`);
  };

  const inspectUploadedReport = async (filePath, customRulePath = ruleFile?.path || null) => {
    if (!window.electron?.testDataCollection?.inspectReportContext) {
      return null;
    }

    return window.electron.testDataCollection.inspectReportContext({
      reportPath: filePath,
      rulePath: customRulePath,
      customer: selectedCustomer
    });
  };

  const refreshRuleSuggestionsForReports = (nextRulePath = null) => {
    files.forEach((item) => {
      inspectUploadedReport(item.path, nextRulePath)
        .then((inspection) => {
          if (!inspection) {
            return;
          }

          setFiles((prev) => prev.map((currentItem) => {
            if (currentItem.path !== item.path) {
              return currentItem;
            }

            const currentSelectedRuleProfile = normalizeRuleProfileValue(currentItem.selectedRuleProfile || currentItem.ruleProfileKey);
            const previousSuggestedRuleProfile = normalizeRuleProfileValue(currentItem.suggestedRuleProfile);
            const nextSuggestedRuleProfile = normalizeRuleProfileValue(inspection.suggestedRuleProfile);
            const shouldAdoptSuggestedRule = !currentSelectedRuleProfile || currentSelectedRuleProfile === previousSuggestedRuleProfile;

            return {
              ...currentItem,
              reportContext: inspection.reportContext || currentItem.reportContext,
              contextInspectionStatus: 'success',
              contextInspectionError: '',
              suggestedRuleProfile: nextSuggestedRuleProfile,
              suggestedRuleProfileReason: inspection.suggestedRuleProfileReason || '',
              availableRuleProfiles: Array.isArray(inspection.availableRuleProfiles) ? inspection.availableRuleProfiles : currentItem.availableRuleProfiles,
              needsRuleConfirmation: Boolean(inspection.needsRuleConfirmation),
              selectedRuleProfile: shouldAdoptSuggestedRule ? nextSuggestedRuleProfile : currentSelectedRuleProfile,
              reportPanelSelections: normalizePanelSelections(
                inspection.suggestedReportPanelSelections
                || inspection.reportContext?.reportPanelSelections
                || currentItem.reportPanelSelections
              )
            };
          }));
        })
        .catch((error) => {
          setFiles((prev) => prev.map((currentItem) => {
            if (currentItem.path !== item.path) {
              return currentItem;
            }

            return {
              ...currentItem,
              contextInspectionStatus: 'error',
              contextInspectionError: error?.message || '读取规则预选失败'
            };
          }));
        });
    });
  };

  const updateReportPanelSelection = (cell, value) => {
    const nextPanelState = applyManualPanelSelectionChange(reportPanelMeta, reportPanelSelections, cell, value);
    setReportPanelSelections(nextPanelState.selections);
    setReportPanelDynamicOptions(nextPanelState.dynamicOptions);
  };

  const updateFileReportPanelSelection = (reportPath, cell, value) => {
    setFiles((prev) => prev.map((item) => {
      if (item.path !== reportPath) {
        return item;
      }

      const nextPanelState = applyManualPanelSelectionChange(
        reportPanelMeta,
        item.reportPanelSelections || item.reportContext?.reportPanelSelections || createEmptyReportPanelSelections(),
        cell,
        value
      );

      return {
        ...item,
        parameterConfirmationStatus: 'pending',
        reportPanelSelections: nextPanelState.selections,
        reportContext: {
          ...(item.reportContext || {}),
          reportPanelSelections: nextPanelState.selections
        }
      };
    }));
  };

  const confirmFileReportPanelSelection = (reportPath) => {
    setFiles((prev) => prev.map((item) => {
      if (item.path !== reportPath) {
        return item;
      }

      if (!normalizeRuleProfileValue(item.selectedRuleProfile || item.ruleProfileKey)) {
        message.warning('请先选择规则模式，再确认当前报告。');
        return item;
      }

      if (!String(item.projectName || '').trim()) {
        message.warning('请输入项目名，再确认当前报告。');
        return item;
      }

      if (!String(item.projectPhase || '').trim()) {
        message.warning('请选择项目阶段，再确认当前报告。');
        return item;
      }

      return {
        ...item,
        parameterConfirmationStatus: 'confirmed'
      };
    }));
  };

  const updateFileRuleProfileSelection = (reportPath, value) => {
    setFiles((prev) => prev.map((item) => {
      if (item.path !== reportPath) {
        return item;
      }

      return {
        ...item,
        selectedRuleProfile: normalizeRuleProfileValue(value),
        parameterConfirmationStatus: 'pending'
      };
    }));
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

      const detectedContext = detectReportContext(file.name);
      const newItem = {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        name: file.name,
        path: file.path,
        bundleKey: getBundleKey(file.name),
        reportKind: getReportKind(file.name),
        reportContext: detectedContext,
        reportPanelSelections: createEmptyReportPanelSelections(),
        parameterConfirmationStatus: 'pending',
        contextInspectionStatus: 'pending',
        contextInspectionError: '',
        suggestedRuleProfile: '',
        suggestedRuleProfileReason: '',
        selectedRuleProfile: '',
        availableRuleProfiles: [],
        needsRuleConfirmation: true,
        ruleSelectionSource: '',
        ruleSelectionReason: '',
        status: 'pending',
        items: 0,
        ruleProfileKey: '',
        outputPath: '',
        outputName: '',
        error: '',
        skippedItems: [],
        unmatchedItems: [],
        audit: null,
        projectName: detectedContext.projectName || '',
        projectPhase: detectedContext.projectPhase || ''
      };

      setProcessedConclusion(null);
      setFiles((prev) => {
        const exists = prev.some((item) => item.path === file.path);
        return exists ? prev : [newItem, ...prev];
      });

      inspectUploadedReport(file.path)
        .then((inspection) => {
          if (!inspection) {
            return;
          }

          setFiles((prev) => prev.map((item) => {
            if (item.path !== file.path) {
              return item;
            }

            const backendContext = inspection.reportContext || {};
            return {
              ...item,
              reportContext: backendContext,
              contextInspectionStatus: 'success',
              contextInspectionError: '',
              parameterConfirmationStatus: 'pending',
              suggestedRuleProfile: normalizeRuleProfileValue(inspection.suggestedRuleProfile),
              suggestedRuleProfileReason: inspection.suggestedRuleProfileReason || '',
              selectedRuleProfile: normalizeRuleProfileValue(inspection.suggestedRuleProfile),
              availableRuleProfiles: Array.isArray(inspection.availableRuleProfiles) ? inspection.availableRuleProfiles : [],
              needsRuleConfirmation: Boolean(inspection.needsRuleConfirmation),
              reportPanelSelections: normalizePanelSelections(
                inspection.suggestedReportPanelSelections
                || inspection.reportContext?.reportPanelSelections
                || item.reportPanelSelections
              ),
              projectName: backendContext.projectName || item.projectName || '',
              projectPhase: backendContext.projectPhase || item.projectPhase || ''
            };
          }));
        })
        .catch((error) => {
          setFiles((prev) => prev.map((item) => {
            if (item.path !== file.path) {
              return item;
            }

            return {
              ...item,
              contextInspectionStatus: 'error',
              contextInspectionError: error?.message || '读取参数上下文失败'
            };
          }));
          message.warning(error?.message || `读取 ${file.name} 的参数上下文失败，将保留手动选择。`);
        });

      message.success(`已添加报告: ${file.name}`);
    }

    if (target === 'rules') {
      setProcessedConclusion(null);
      setRuleFile({ name: file.name, path: file.path });
      refreshRuleSuggestionsForReports(file.path);
      if (presetChecklistTemplate) {
        setChecklistConfirmationStatus('pending');
        resolveAndLoadPresetChecklistTemplate(presetChecklistTemplate, file.path);
      }
      message.success(`已上传规则: ${file.name}`);
    }

    if (target === 'checklist') {
      setProcessedConclusion(null);
      setPresetChecklistTemplate('');
      setPresetChecklistPath('');
      setChecklistSelectionSource('upload');
      setChecklistConfirmationStatus('pending');
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
    setReportPanelSelections(createEmptyReportPanelSelections());
    setReportPanelDynamicOptions({});
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
      await window.electron.testDataCollection.showOutputInFolder(record.outputPath);
    } catch (error) {
      message.error(error?.message || '打开输出目录失败');
    }
  };

  const openInfoModal = (config) => {
    modal.info({
      closable: true,
      keyboard: true,
      maskClosable: true,
      ...config
    });
  };

  const showDetails = (record) => {
    openInfoModal({
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
            {(record.reportPanelSelections || record.reportContext?.reportPanelSelections)
              ? `B13=${(record.reportPanelSelections || record.reportContext?.reportPanelSelections).B13 || '-'} / B15=${(record.reportPanelSelections || record.reportContext?.reportPanelSelections).B15 || '-'} / C15=${(record.reportPanelSelections || record.reportContext?.reportPanelSelections).C15 || '-'} / D15=${(record.reportPanelSelections || record.reportContext?.reportPanelSelections).D15 || '-'}`
              : '未指定'}
          </Paragraph>
          <Paragraph>
            <Text strong>命中规则数：</Text> {record.items || 0}
          </Paragraph>
          <Paragraph>
            <Text strong>生效规则 Profile：</Text> {record.ruleProfileKey || '未标记'}
          </Paragraph>
          <Paragraph>
            <Text strong>规则选择来源：</Text> {record.ruleSelectionSource || '未标记'}
            {record.ruleSelectionReason ? ` / ${record.ruleSelectionReason}` : ''}
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
    openInfoModal({
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
    openInfoModal({
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
    openInfoModal({
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
    openInfoModal({
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

    // 如果没有上传 checklist，但选择了预设模板，则允许继续（后端会使用内置模板）
    if (!checklistFile?.path && !presetChecklistPath) {
      message.warning('请先上传 checklist Excel 文件，或选择预设模板。');
      return;
    }

    if (checklistConfirmationStatus !== 'confirmed') {
      message.warning('请先确认当前 checklist 模板或 checklist 文件，再开始处理。');
      return;
    }

    if (uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0) {
      const confirmed = await new Promise((resolve) => {
        modal.confirm({
          title: '当前仅导入 Word 报告',
          content: 'Word 提取的准确性和稳定性低于 Excel。若继续处理，输出结果必须由人工逐项确认后才能使用。',
          okText: '继续并人工确认',
          cancelText: '返回检查',
          okButtonProps: { danger: true },
          onOk: () => resolve(true),
          onCancel: () => resolve(false)
        });
      });

      if (!confirmed) {
        return;
      }
    }

    const unconfirmedReports = files.filter(
      (item) => item.parameterConfirmationStatus !== 'confirmed'
    );
    if (unconfirmedReports.length > 0) {
      message.warning(`还有 ${unconfirmedReports.length} 份报告未确认配置，请先确认后再开始处理。`);
      return;
    }

    const reportsWithoutRuleProfile = files.filter((item) => !normalizeRuleProfileValue(item.selectedRuleProfile || item.ruleProfileKey));
    if (reportsWithoutRuleProfile.length > 0) {
      message.warning(`还有 ${reportsWithoutRuleProfile.length} 份报告没有选择规则模式，请先补齐。`);
      return;
    }

    const duplicateOutputConfigs = findDuplicateOutputConfigs(files);
    if (duplicateOutputConfigs.length > 0) {
      const confirmed = await new Promise((resolve) => {
        modal.confirm({
          title: '检测到重复的输出命名配置',
          content: `有 ${duplicateOutputConfigs.length} 组报告会生成相同的逻辑文件名。系统会自动补报告名后缀避免覆盖，但建议你先确认这些报告的 Network / Vocoder / 项目名 / 项目阶段 是否确实一致。`,
          okText: '确认继续',
          cancelText: '返回检查',
          onOk: () => resolve(true),
          onCancel: () => resolve(false)
        });
      });

      if (!confirmed) {
        return;
      }
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
      const response = await window.electron.testDataCollection.processReports({
        runId,
        reportPaths: files.map((item) => item.path),
        checklistPath: checklistFile?.path || presetChecklistPath || null,
        rulePath: ruleFile?.path || null,
        customer: selectedCustomer,
        reportPanelSelections: null,
        reportPanelSelectionsByPath: uploadSummary.excelCount > 0
          ? Object.fromEntries(
            files
              .filter((item) => item.reportKind === 'excel')
              .map((item) => [item.path, normalizePanelSelections(item.reportPanelSelections || {})])
          )
          : null,
        ruleProfileOverridesByPath: Object.fromEntries(
          files.map((item) => [item.path, normalizeRuleProfileValue(item.selectedRuleProfile || item.ruleProfileKey)])
        ),
        reportProjectMetaByPath: Object.fromEntries(
          files.map((item) => [item.path, {
            projectName: String(item.projectName || '').trim(),
            projectPhase: String(item.projectPhase || '').trim()
          }])
        )
      });

      const resultMap = new Map(response.results.map((item) => [item.reportPath, item]));
      recordDataCollectionResults(response.results);
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
          selectedRuleProfile: normalizeRuleProfileValue(result.ruleProfileKey || item.selectedRuleProfile || item.ruleProfileKey),
          suggestedRuleProfile: normalizeRuleProfileValue(result.reportContext?.suggestedRuleProfile || item.suggestedRuleProfile),
          suggestedRuleProfileReason: result.reportContext?.suggestedRuleProfileReason || item.suggestedRuleProfileReason,
          ruleSelectionSource: result.ruleSelectionSource || item.ruleSelectionSource || '',
          ruleSelectionReason: result.ruleSelectionReason || item.ruleSelectionReason || '',
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
      const errorMessage = error?.message || '执行测试数据收集失败';
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
  const uploadSummary = useMemo(() => buildUploadSummary(files, checklistFile, presetChecklistPath), [files, checklistFile, presetChecklistPath]);
  const excelReportRecords = useMemo(() => files.filter((item) => item.reportKind === 'excel'), [files]);
  const isMultiExcelMode = excelReportRecords.length > 1;
  const hasExcelSource = excelReportRecords.length > 0;
  const conclusionData = useMemo(() => buildConclusionData(uploadSummary, processedConclusion), [uploadSummary, processedConclusion]);

  const exportRules = async () => {
    setExportingRules(true);

    try {
      const result = await window.electron.testDataCollection.exportRules(ruleFile?.path || null);
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
      title: '自动识别状态',
      key: 'autoDetectionStatus',
      width: 320,
      render: (_, record) => {
        const autoDetectionMeta = getAutoDetectionMeta(record, isMultiExcelMode);

        return (
          <div>
            <div style={{ marginBottom: 6 }}>
              <Tag color={autoDetectionMeta.color}>{autoDetectionMeta.label}</Tag>
            </div>
            <Text type="secondary" className="report-checker-table-text">
              {autoDetectionMeta.detail}
            </Text>
            {autoDetectionMeta.tags.length > 0 ? (
              <div style={{ marginTop: 8 }}>
                <Space size={[4, 4]} wrap>
                  {autoDetectionMeta.tags.map((tag) => <Tag key={`${record.id}-${tag}`}>{tag}</Tag>)}
                </Space>
              </div>
            ) : null}
          </div>
        );
      }
    },
    {
      title: '操作',
      key: 'action',
      width: 120,
      render: (_, record) => (
        <Space wrap={false}>
          <Button danger size="small" icon={<DeleteOutlined />} onClick={() => removeReport(record.id)}>删除</Button>
        </Space>
      )
    }
  ];

  const resultColumns = [
    {
      title: '文件名',
      dataIndex: 'name',
      key: 'name',
      width: 260,
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
      width: 360,
      ellipsis: true,
      render: (_, record) => {
        if (!record.outputName) {
          return <Text type="secondary">处理完成后显示</Text>;
        }

        return (
          <Text className="report-checker-table-text" ellipsis={{ tooltip: record.outputName }}>
            {record.outputName}
          </Text>
        );
      }
    },
    {
      title: '结果操作',
      key: 'resultAction',
      width: 220,
      render: (_, record) => (
        <Space wrap={false}>
          <Button type="primary" size="small" onClick={() => showDetails(record)}>详情</Button>
          <Button size="small" disabled={!record.outputPath} onClick={() => openOutputFolder(record)}>打开目录</Button>
        </Space>
      )
    }
  ];

  const resultRecords = useMemo(
    () => files.filter((item) => item.status === 'success' || item.status === 'error' || item.status === 'processing'),
    [files]
  );

  return (
    <div className="page-container">
      <Card className="report-checker-card report-checker-section-card report-checker-guide-card" title="使用说明">
        <div className="report-checker-note-list">
          <Paragraph style={{ marginBottom: 10 }}>先上传报告、checklist 和可选规则，再执行测试数据收集，最后在结论窗口查看覆盖性、文档审查和一致性结果。</Paragraph>
          <Paragraph style={{ marginBottom: 10 }}>Excel 是默认主数据源，优先用于 checklist 填表与覆盖性评估；Word 只建议作为兼容补充通道。</Paragraph>
          <Paragraph style={{ marginBottom: 10 }}>跨网络或跨 codec 的一致性审查只在存在可比样本时触发。</Paragraph>
          <Paragraph style={{ marginBottom: 0 }}>如果当前批次仅导入 Word 报告，UI 会重点提示并要求人工确认后再使用输出结果。</Paragraph>
        </div>
      </Card>

      {uploadSummary.wordCount > 0 && uploadSummary.excelCount === 0 ? (
        <Alert
          style={{ marginBottom: 16 }}
          type="warning"
          showIcon
          message="当前仅导入 Word 报告"
          description="Word 提取准确性和稳定性低于 Excel。若继续处理，本批次生成的 checklist 结果必须由人工逐项确认。"
        />
      ) : null}

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
                  ? `已选择 ${uploadSummary.totalReports} 份报告；确认报告参数和 checklist 模板后即可执行。`
                  : '还没有选择测试报告'}
              </span>
            </div>
          </div>

          <p style={{ marginBottom: '24px', color: 'var(--text-light)' }}>
            Excel 报告优先用于 checklist 填表、漏测/重测/多测候选评估与跨报告一致性检查；Word 报告可参与提取，但默认只作为兼容补充与证据审查来源。
          </p>
          <p style={{ marginTop: '-12px', marginBottom: '24px', color: 'var(--text-light)' }}>
            .doc 报告会先在后台尝试转成 .docx；若当前批次只有 Word，系统仍会继续写入 checklist，但结果必须人工逐项确认。
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
              <UploadOutlined className="report-checker-upload-placeholder__icon report-checker-upload-placeholder__icon--large" />
              <p className="report-checker-upload-placeholder__title report-checker-upload-placeholder__title--large">
                拖拽报告文件到此处，或点击上方按钮上传
              </p>
              <p className="report-checker-upload-placeholder__subtitle">
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
            <div style={{ marginBottom: 12 }}>
              <div className="report-checker-template-row">
                <Text strong className="report-checker-template-label">预设模板</Text>
                <div className="report-checker-template-select-wrap">
                  <Select
                    showSearch
                    allowClear
                    value={presetChecklistTemplate || undefined}
                    className="report-checker-template-select"
                    placeholder="根据报告模式自动选择"
                    onChange={(value) => {
                      setPresetChecklistTemplate(value || '');
                      if (value) {
                        setChecklistFile(null);
                        setChecklistSelectionSource('manual');
                        setChecklistConfirmationStatus('pending');
                        resolveAndLoadPresetChecklistTemplate(value);
                      } else {
                        setChecklistSelectionSource('');
                        setChecklistConfirmationStatus('pending');
                        setPresetChecklistPath('');
                        resetChecklistReportPanelOptions();
                      }
                    }}
                    options={CHECKLIST_TEMPLATE_OPTIONS}
                  />
                </div>
                <Space size={[8, 8]} wrap className="report-checker-template-meta">
                  {suggestedChecklistTemplate ? (
                    <Tag color="blue">系统预选: {formatChecklistTemplateLabel(suggestedChecklistTemplate)}</Tag>
                  ) : null}
                  <Tag color={checklistConfirmationStatus === 'confirmed' ? 'green' : 'gold'}>
                    {checklistConfirmationStatus === 'confirmed' ? '模板已确认' : '模板待确认'}
                  </Tag>
                  <Button
                    size="small"
                    type={checklistConfirmationStatus === 'confirmed' ? 'default' : 'primary'}
                    disabled={checklistConfirmationStatus === 'confirmed' || (!checklistFile?.path && !presetChecklistPath)}
                    onClick={confirmChecklistSelection}
                  >
                    {checklistConfirmationStatus === 'confirmed' ? '已确认' : '确认'}
                  </Button>
                </Space>
              </div>
              <Text type="secondary" style={{ display: 'block', marginTop: 4, fontSize: 12 }}>
                上传测试报告后会先自动预选模板；你可以确认当前模板，也可以手动切换后再确认。上传自定义 checklist 可覆盖预设选择。
              </Text>
            </div>

            <Upload.Dragger
              customRequest={({ file, onSuccess }) => handleUpload(file, 'checklist', onSuccess)}
              accept=".xlsx,.xls"
              showUploadList={false}
              className="report-checker-upload"
              style={compactUploadDraggerStyle}
            >
              {checklistFile ? (
                <div style={{ padding: '8px 0' }}>
                  <Text strong className="report-checker-upload-placeholder__title" style={{ display: 'block', fontSize: '16px', marginBottom: '6px' }}>
                    已选择 checklist
                  </Text>
                  <Text className="report-checker-upload-placeholder__file" style={{ fontSize: '14px', wordBreak: 'break-all' }}>
                    {checklistFile.name}
                  </Text>
                </div>
              ) : presetChecklistPath ? (
                <div style={{ padding: '8px 0' }}>
                  <Text strong className="report-checker-upload-placeholder__title" style={{ display: 'block', fontSize: '16px', marginBottom: '6px' }}>
                    已启用预设模板
                  </Text>
                  <Text className="report-checker-upload-placeholder__file" style={{ fontSize: '14px', wordBreak: 'break-all' }}>
                    {presetChecklistTemplate}
                  </Text>
                </div>
              ) : (
                <>
                  <UploadOutlined className="report-checker-upload-placeholder__icon" />
                  <p className="report-checker-upload-placeholder__title">
                    拖拽 checklist 文件到此处，或点击上传
                  </p>
                  <p className="report-checker-upload-placeholder__subtitle">
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
              <UploadOutlined className="report-checker-upload-placeholder__icon" />
              <p className="report-checker-upload-placeholder__title">
                拖拽规则文件到此处，或点击上传
              </p>
              <p className="report-checker-upload-placeholder__subtitle">
                支持格式: JSON / JSON5；不上传时默认使用内置规则
              </p>
            </Upload.Dragger>
            {ruleFile && (
              <p className="report-checker-upload-placeholder__file" style={{ marginTop: '12px' }}>
                已选择: {ruleFile.name}
              </p>
            )}
          </Card>
        </div>
      </div>

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
                : presetChecklistPath
                  ? `预设模板 ${presetChecklistTemplate} / ${reportPanelMeta.reportSheetName || 'Report'} 页`
                : '先上传 checklist 后读取参数'}
            </Text>
            <div style={{ marginTop: 8 }}>
              <Text type="secondary">
                {files.length > 0
                  ? '每份报告都要确认规则模式；Excel 额外确认 Report 参数；checklist 模板也需要确认。未确认前不能开始解析。'
                  : '上传报告后，这里会显示规则预选和参数确认入口。'}
              </Text>
            </div>
          </div>
        </div>
        <Paragraph type="secondary" style={{ marginTop: 16, marginBottom: 0 }}>
          客户、已确认的规则模式以及 Excel 参数会随本次任务提交后端，用于规则分发、动态模板切换与 Report 页回写。
        </Paragraph>

        {files.length > 0 ? (
          <div style={{ marginTop: 24, paddingTop: 20, borderTop: '1px solid #f0f0f0' }}>
            <Text strong style={{ display: 'block', marginBottom: 8 }}>报告配置确认</Text>
            <Paragraph type="secondary" style={{ marginBottom: 16 }}>
              系统会先按文件名预选规则模式。你可以直接接受，也可以手动切换；如果是 Excel 报告，还需要同时确认 Report 参数。
            </Paragraph>

            <div style={{ display: 'grid', gap: 16 }}>
              {files.map((record) => {
                const rowSelections = normalizePanelSelections(record.reportPanelSelections || record.reportContext?.reportPanelSelections || {});
                const detectedTags = buildDetectedTags(record.reportContext || {});
                const confirmationMeta = getParameterConfirmationMeta(record);
                const ruleSelectionMeta = getRuleSelectionMeta(record);
                const ruleProfileOptions = getRuleProfileOptions(record);
                const isConfirmed = record.parameterConfirmationStatus === 'confirmed';
                const configFields = [
                  {
                    key: 'project-name',
                    label: '项目名',
                    type: 'input',
                    span: 2,
                    extracted: false,
                    value: record.projectName || '',
                    placeholder: '自动识别'
                  },
                  {
                    key: 'project-phase',
                    label: '阶段',
                    type: 'select',
                    extracted: false,
                    value: record.projectPhase || undefined,
                    placeholder: '自动识别',
                    options: PROJECT_PHASE_OPTIONS.map((opt) => ({ label: opt, value: opt }))
                  },
                  {
                    key: 'rule-profile',
                    label: '规则模式',
                    type: 'select',
                    extracted: false,
                    value: normalizeRuleProfileValue(record.selectedRuleProfile || record.ruleProfileKey) || undefined,
                    placeholder: '请选择规则模式',
                    options: ruleProfileOptions,
                    notFoundContent: ruleProfileOptions.length === 0 ? '暂无可用规则模式' : null
                  },
                  ...(record.reportKind === 'excel'
                    ? REPORT_PANEL_FIELDS.map((field) => {
                      const fieldOptions = getOptionsForCell(reportPanelMeta, field.cell, rowSelections);

                      return {
                        key: field.cell,
                        label: `${REPORT_PANEL_FIELD_SHORT_LABELS[field.cell] || field.label} (${field.cell})`,
                        type: 'select',
                        extracted: true,
                        value: rowSelections[field.cell] || undefined,
                        placeholder: '请选择参数值',
                        options: fieldOptions.map((item) => ({ label: item, value: item })),
                        notFoundContent: fieldOptions.length === 0 ? '暂无候选值' : null,
                        cell: field.cell
                      };
                    })
                    : [])
                ];

                return (
                  <div
                    key={record.id}
                    className={`report-parameter-card report-parameter-card--${record.parameterConfirmationStatus === 'confirmed' ? 'confirmed' : 'pending'}`}
                  >
                    <div className="report-parameter-card__header">
                      <div className="report-parameter-card__summary">
                        <Text strong className="report-parameter-card__name">{record.name}</Text>
                        <Space size={[8, 8]} wrap className="report-parameter-card__tags">
                          {detectedTags.length > 0
                            ? detectedTags.map((tag) => <Tag key={`${record.id}-${tag}`} color="blue">{tag}</Tag>)
                            : <Tag>未识别到上下文</Tag>}
                          {record.reportContext?.measurementObject ? <Tag color="default">{record.reportContext.measurementObject}</Tag> : null}
                        </Space>
                        <div style={{ marginTop: 10 }}>
                          <Tag color={ruleSelectionMeta.color}>{ruleSelectionMeta.label}</Tag>
                          <Text type="secondary" style={{ marginLeft: 8 }}>{ruleSelectionMeta.detail}</Text>
                        </div>
                      </div>
                      <Space size={8} className="report-parameter-card__actions">
                        <Tag color={confirmationMeta.color}>{confirmationMeta.label}</Tag>
                        <Button
                          type={record.parameterConfirmationStatus === 'confirmed' ? 'default' : 'primary'}
                          size="small"
                          disabled={record.parameterConfirmationStatus === 'confirmed'}
                          onClick={() => confirmFileReportPanelSelection(record.path)}
                        >
                          {record.parameterConfirmationStatus === 'confirmed' ? '已确认' : '确认'}
                        </Button>
                      </Space>
                    </div>

                    <div className="report-parameter-card__config-grid">
                      {configFields.map((field) => {
                        const pillClassName = [
                          'report-config-pill',
                          field.span === 2 ? 'report-config-pill--span-2' : '',
                          field.extracted ? 'report-config-pill--extracted' : '',
                          isConfirmed ? 'report-config-pill--disabled' : ''
                        ].filter(Boolean).join(' ');

                        return (
                          <div key={`${record.id}-${field.key}`} className={pillClassName}>
                            <div className="report-config-pill__label">{field.label}</div>
                            <div className="report-config-pill__control">
                              {field.type === 'input' ? (
                                <input
                                  type="text"
                                  className="report-config-pill__input ant-input"
                                  value={field.value}
                                  placeholder={field.placeholder}
                                  disabled={isConfirmed}
                                  onChange={(e) => {
                                    setFiles((prev) => prev.map((item) => {
                                      if (item.path !== record.path) return item;
                                      return { ...item, projectName: e.target.value.toUpperCase(), parameterConfirmationStatus: 'pending' };
                                    }));
                                  }}
                                />
                              ) : (
                                <Select
                                  showSearch
                                  allowClear={!isConfirmed}
                                  disabled={isConfirmed}
                                  value={field.value}
                                  className="report-config-pill__select"
                                  popupClassName="report-config-pill__dropdown"
                                  placeholder={field.placeholder}
                                  onChange={(value) => {
                                    if (field.key === 'project-phase') {
                                      setFiles((prev) => prev.map((item) => {
                                        if (item.path !== record.path) return item;
                                        return { ...item, projectPhase: value || '', parameterConfirmationStatus: 'pending' };
                                      }));
                                      return;
                                    }

                                    if (field.key === 'rule-profile') {
                                      updateFileRuleProfileSelection(record.path, value);
                                      return;
                                    }

                                    updateFileReportPanelSelection(record.path, field.cell, value);
                                  }}
                                  options={field.options}
                                  notFoundContent={field.notFoundContent ?? null}
                                />
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        ) : null}
      </Card>

      <Card className="report-checker-card report-checker-action-card">
        <div className="report-checker-action-panel">
          <div>
            <div className="report-checker-action-title">开始收集并输出结论</div>
            <div className="report-checker-action-text">上传完成后，从这里启动处理。结果会先更新进度，再在下方结论窗口集中展示。</div>
          </div>
          <Button
            type="primary"
            icon={<CheckCircleOutlined />}
            loading={processing}
            onClick={processReports}
            className="report-checker-primary-action"
          >
            开始收集并输出结论
          </Button>
        </div>
      </Card>

      {(processing || progressState.completed > 0) && (
        <Card className="report-checker-card report-checker-progress-card">
          <div className="report-checker-progress-header">
            <div>
              <div className="report-checker-progress-title">任务进度</div>
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

      {resultRecords.length > 0 ? (
        <Card className="report-checker-card report-checker-section-card" title="数据收集结果">
          <Paragraph type="secondary" style={{ marginBottom: 16 }}>
            这里集中展示报告处理后的收集结果、输出文件和结果操作。处理完成后可直接查看详情或打开输出目录，不需要再回到上传列表。
          </Paragraph>
          <Table
            className="report-checker-table report-checker-report-table"
            columns={resultColumns}
            dataSource={resultRecords}
            rowKey="id"
            scroll={{ x: 1120 }}
            pagination={{ pageSize: 8 }}
          />
        </Card>
      ) : null}

      <Card className="report-checker-card report-checker-conclusion-card" title="结论输出">
        {conclusionData.sourcePolicy ? (
          <Alert
            style={{ marginBottom: 16 }}
            type={conclusionData.sourcePolicy.manualConfirmationLevel === 'required' ? 'warning' : (conclusionData.sourcePolicy.manualConfirmationLevel === 'recommended' ? 'info' : 'success')}
            showIcon
            message={`数据源策略：${conclusionData.sourcePolicy.summary}`}
            description={`默认优先级：Excel 主、Word 辅；当前置信度：${conclusionData.sourcePolicy.confidenceLabel}；${conclusionData.sourcePolicy.detail}`}
          />
        ) : null}

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
              <Space size={6} wrap>
                <Tag color={getStatusMeta(conclusionData.wordAudit.status).color}>{getStatusMeta(conclusionData.wordAudit.status).label}</Tag>
                <Tag color={conclusionData.sourcePolicy?.confidence === 'low' ? 'red' : conclusionData.sourcePolicy?.confidence === 'medium' ? 'gold' : 'green'}>
                  置信度 {conclusionData.sourcePolicy?.confidenceLabel || '未触发'}
                </Tag>
              </Space>
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

          <button type="button" className="report-checker-insight-card" onClick={() => openInfoModal({ title: '当前建议', width: 760, content: (<div style={{ marginTop: 16 }}>{conclusionData.suggestedActions.length > 0 ? conclusionData.suggestedActions.map((item) => <Paragraph key={item} style={{ marginBottom: 8 }}>{item}</Paragraph>) : <Paragraph style={{ marginBottom: 0 }}>当前没有新增建议。</Paragraph>}</div>) })}>
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
                    <Tag color={bundle.sourceMode === 'word' ? 'red' : bundle.sourceMode === 'excel+word' ? 'gold' : 'green'}>
                      {bundle.sourceMode === 'word' ? '必须人工确认' : bundle.sourceMode === 'excel+word' ? '建议人工复核' : 'Excel 主通道'}
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
            执行数据收集后，这里会按 Excel 填表、Word 审查和跨报告一致性输出完整结论。
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

export default TestDataCollectionPage;
