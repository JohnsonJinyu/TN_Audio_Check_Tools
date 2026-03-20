const path = require('path');
const fs = require('fs/promises');

function createReportExtractor({
  parseReport,
  applyResultsToChecklist,
  resolveRowBasedValue,
  resolveAnchorValue,
  resolveTableValue,
  resolveRegexValue,
  analyzeExcelReport,
  analyzeWordReport
}) {
  function resolveTerminalModeFromInterface(value) {
    const normalized = String(value || '').trim().toLowerCase();
    if (!normalized) {
      return '';
    }

    if (normalized.includes('handset') || normalized === 'ha') {
      return 'HA';
    }

    if (normalized.includes('headset') || normalized === 'he' || normalized === 'hs') {
      return 'HE';
    }

    if (normalized.includes('handsfree') || normalized === 'hh' || normalized === 'hf') {
      return 'HH';
    }

    return '';
  }

  function normalizeReportPanelSelections(reportPanelSelections = {}) {
    const normalized = {
      B13: String(reportPanelSelections.B13 || '').trim(),
      B15: String(reportPanelSelections.B15 || '').trim(),
      C15: String(reportPanelSelections.C15 || '').trim(),
      D15: String(reportPanelSelections.D15 || '').trim()
    };

    if (!normalized.B13 && !normalized.B15 && !normalized.C15 && !normalized.D15) {
      return null;
    }

    return normalized;
  }

  function mergePanelSelections(primarySelections, fallbackSelections) {
    const primary = normalizeReportPanelSelections(primarySelections) || {};
    const fallback = normalizeReportPanelSelections(fallbackSelections) || {};

    const merged = {
      B13: primary.B13 || fallback.B13 || '',
      B15: primary.B15 || fallback.B15 || '',
      C15: primary.C15 || fallback.C15 || '',
      D15: primary.D15 || fallback.D15 || ''
    };

    return normalizeReportPanelSelections(merged);
  }

  function formatNetworkLabel(value) {
    const normalized = String(value || '').trim().toUpperCase();
    const networkMap = {
      VOLTE: 'VoLTE',
      VOWIFI: 'VoWiFi',
      VONR: 'VoNR',
      VOIP: 'VoIP',
      WCDMA: 'WCDMA',
      GSM: 'GSM'
    };

    return networkMap[normalized] || '';
  }

  function buildAutoReportPanelSelections(reportContext = {}) {
    const vocoder = buildNormalizedVocoder(reportContext);
    const selections = {
      B13: String(reportContext.headsetInterface || '').trim(),
      B15: formatNetworkLabel(reportContext.network),
      C15: vocoder.full || '',
      D15: String(reportContext.bitrate || '').trim()
    };

    return normalizeReportPanelSelections(selections);
  }

  function mergeReportContext(baseReportContext = {}, { customer, reportPanelSelections } = {}) {
    const merged = {
      ...baseReportContext
    };

    const normalizedCustomer = String(customer || '').trim();
    if (normalizedCustomer) {
      merged.customer = normalizedCustomer;
    }

    const normalizedSelections = mergePanelSelections(reportPanelSelections, buildAutoReportPanelSelections(baseReportContext));
    if (!normalizedSelections) {
      return merged;
    }

    merged.reportPanelSelections = normalizedSelections;

    if (normalizedSelections.B15) {
      merged.network = normalizedSelections.B15;
    }

    if (normalizedSelections.C15) {
      merged.vocoder = normalizedSelections.C15;
      const vocoderMatch = normalizedSelections.C15.toUpperCase().match(/(EVS|AMR)(?:[_\-\s]*(NB|WB|SWB|SB))?/);
      if (vocoderMatch) {
        if (vocoderMatch[1]) {
          merged.codec = vocoderMatch[1];
        }

        if (vocoderMatch[2]) {
          merged.bandwidth = vocoderMatch[2] === 'SB' ? 'SWB' : vocoderMatch[2];
        }
      }
    }

    if (normalizedSelections.D15) {
      merged.bitrate = normalizedSelections.D15;
    }

    const terminalMode = resolveTerminalModeFromInterface(normalizedSelections.B13);
    if (terminalMode) {
      merged.terminalMode = terminalMode;
    }

    return merged;
  }

  function inspectReportContext(baseReportContext = {}, { customer, reportPanelSelections } = {}) {
    return mergeReportContext(baseReportContext, { customer, reportPanelSelections });
  }

  function resolveTerminalModeKey(value) {
    const terminalMode = String(value || '').trim().toUpperCase();
    if (terminalMode === 'HA') {
      return 'handset';
    }

    if (terminalMode === 'HH' || terminalMode === 'HF') {
      return 'handsfree';
    }

    if (terminalMode === 'HE' || terminalMode === 'HS') {
      return 'headset';
    }

    return '';
  }

  function resolveChecklistModeKeyFromName(checklistPath) {
    const checklistName = path.basename(checklistPath || '').toLowerCase();
    if (checklistName.includes('handsfree')) {
      return 'handsfree';
    }

    if (checklistName.includes('headset')) {
      return 'headset';
    }

    if (checklistName.includes('handset')) {
      return 'handset';
    }

    return '';
  }

  function normalizeChecklistTemplateName(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '');
  }

  function buildChecklistModeTokens(modeKey) {
    if (modeKey === 'handset') {
      return ['handset'];
    }

    if (modeKey === 'handsfree') {
      return ['handsfree'];
    }

    if (modeKey === 'headset') {
      return ['headset'];
    }

    return [];
  }

  async function findModeChecklistTemplate({ checklistPath, reportPath, targetModeKey }) {
    if (!targetModeKey) {
      return '';
    }

    const candidateDirs = [];
    const pushDir = (dirPath) => {
      if (!dirPath) {
        return;
      }

      const normalizedDir = path.resolve(dirPath);
      if (!candidateDirs.includes(normalizedDir)) {
        candidateDirs.push(normalizedDir);
      }
    };

    pushDir(path.dirname(path.dirname(reportPath || '')));
    pushDir(path.dirname(reportPath || ''));
    pushDir(path.dirname(checklistPath || ''));

    const modeTokens = buildChecklistModeTokens(targetModeKey);
    for (const dirPath of candidateDirs) {
      try {
        const entries = await fs.readdir(dirPath, { withFileTypes: true });
        const matchedFile = entries.find((entry) => {
          if (!entry.isFile()) {
            return false;
          }

          if (!entry.name.toLowerCase().endsWith('.xlsx')) {
            return false;
          }

          const normalizedName = normalizeChecklistTemplateName(entry.name);
          const isVoiceTuningTemplate = normalizedName.includes('voicetuningchecklist');
          if (!isVoiceTuningTemplate) {
            return false;
          }

          return modeTokens.some((token) => normalizedName.includes(token));
        });

        if (matchedFile) {
          return path.join(dirPath, matchedFile.name);
        }
      } catch {
        // ignore and continue scanning next candidate directory
      }
    }

    return '';
  }

  async function resolveChecklistPathForReport(checklistPath, reportContext = {}, reportPath = '') {
    if (!checklistPath) {
      return checklistPath;
    }

    const targetModeKey = resolveTerminalModeKey(reportContext?.terminalMode);
    if (!targetModeKey) {
      return checklistPath;
    }

    const modeSpecificTemplate = await findModeChecklistTemplate({
      checklistPath,
      reportPath,
      targetModeKey
    });
    if (modeSpecificTemplate) {
      return modeSpecificTemplate;
    }

    const currentModeKey = resolveChecklistModeKeyFromName(checklistPath);
    if (!currentModeKey || currentModeKey === targetModeKey) {
      return checklistPath;
    }

    const modeNameMap = {
      handset: 'Handset',
      handsfree: 'Handsfree',
      headset: 'Headset'
    };
    const nextModeName = modeNameMap[targetModeKey];
    const checklistDir = path.dirname(checklistPath);
    const checklistName = path.basename(checklistPath);
    const candidateName = checklistName.replace(/handset|handsfree|headset/i, nextModeName);
    const candidatePath = path.join(checklistDir, candidateName);

    try {
      await fs.access(candidatePath);
      return candidatePath;
    } catch {
      return checklistPath;
    }
  }

  function resolveRuleProfileKey(reportData, checklistPath) {
    const modeKey = resolveTerminalModeKey(reportData?.reportContext?.terminalMode);
    if (modeKey === 'handset') {
      return 'handset';
    }

    if (modeKey === 'handsfree') {
      return 'handsfree';
    }

    if (modeKey === 'headset') {
      return 'headset';
    }

    const checklistModeKey = resolveChecklistModeKeyFromName(checklistPath);
    if (checklistModeKey) {
      return checklistModeKey;
    }

    return 'handset';
  }

  function resolveRulesForReport(rules, reportData, checklistPath) {
    if (!rules?.ruleProfiles) {
      return {
        activeRules: rules,
        profileKey: 'default'
      };
    }

    const profileKey = resolveRuleProfileKey(reportData, checklistPath);
    const activeRules = rules.ruleProfiles[profileKey]
      || rules.ruleProfiles[rules.defaultProfileKey]
      || Object.values(rules.ruleProfiles)[0];

    if (!activeRules?.extractItemList) {
      throw new Error(`未找到可用的规则 profile: ${profileKey}`);
    }

    return {
      activeRules,
      profileKey
    };
  }

  function normalizeBandwidth(value) {
    const normalized = String(value || '').trim().toUpperCase();
    if (!normalized) {
      return '';
    }

    if (['SB', 'SWB'].includes(normalized)) {
      return 'SWB';
    }

    if (normalized.endsWith('SWB')) {
      return 'SWB';
    }

    if (normalized.endsWith('WB')) {
      return 'WB';
    }

    if (normalized.endsWith('NB')) {
      return 'NB';
    }

    return normalized;
  }

  function normalizeNetwork(value) {
    const normalized = String(value || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!normalized) {
      return '';
    }

    const networkAliases = {
      VOLTE: 'VOLTE',
      VOWIFI: 'VOWIFI',
      VONR: 'VONR',
      VOIP: 'VOIP',
      WCDMA: 'WCDMA',
      GSM: 'GSM'
    };

    return networkAliases[normalized] || normalized;
  }

  function normalizeVocoderFamily(value) {
    const normalized = String(value || '').trim().toUpperCase();
    if (!normalized) {
      return '';
    }

    if (normalized.includes('EVS')) {
      return 'EVS';
    }

    if (normalized.includes('AMR')) {
      return 'AMR';
    }

    return normalized.replace(/[^A-Z0-9]/g, '');
  }

  function buildNormalizedVocoder(reportContext = {}) {
    const explicitVocoder = String(reportContext.vocoder || '').trim().toUpperCase();
    if (explicitVocoder) {
      const vocoderMatch = explicitVocoder.match(/(EVS|AMR)(?:[_\-\s]*(NB|WB|SWB|SB))?/);
      if (vocoderMatch) {
        const family = vocoderMatch[1];
        const bandwidth = (vocoderMatch[2] || '').replace('SB', 'SWB');
        return {
          family,
          full: bandwidth ? `${family}_${bandwidth}` : family
        };
      }
    }

    const family = normalizeVocoderFamily(reportContext.codec);
    const bandwidth = normalizeBandwidth(reportContext.bandwidth);
    if (!family) {
      return { family: '', full: '' };
    }

    return {
      family,
      full: bandwidth ? `${family}_${bandwidth}` : family
    };
  }

  function normalizeCustomer(value) {
    return String(value || '').trim().toUpperCase();
  }

  function isTokenApplicable({ reportValue, allowedTokens, excludedTokens, reasonPrefix, dimension }) {
    if (allowedTokens.length > 0 && !allowedTokens.includes(reportValue)) {
      return {
        applicable: false,
        reason: `${reasonPrefix} ${reportValue || 'unknown'} 不要求测试该项`,
        skipContext: {
          dimension,
          actual: reportValue || '',
          include: allowedTokens,
          exclude: excludedTokens
        }
      };
    }

    if (excludedTokens.includes(reportValue)) {
      return {
        applicable: false,
        reason: `${reasonPrefix} ${reportValue || 'unknown'} 不要求测试该项`,
        skipContext: {
          dimension,
          actual: reportValue || '',
          include: allowedTokens,
          exclude: excludedTokens
        }
      };
    }

    return { applicable: true };
  }

  function isItemApplicable(reportData, item) {
    const reportTerminalMode = String(reportData?.reportContext?.terminalMode || '').trim().toUpperCase();
    const rawApplicableTerminalModes = [
      ...(Array.isArray(item?.applicableTerminalModes) ? item.applicableTerminalModes : []),
      ...(Array.isArray(item?.applicableModes) ? item.applicableModes : [])
    ];
    const applicableTerminalModes = [...new Set(rawApplicableTerminalModes
      .map((mode) => String(mode || '').trim().toUpperCase())
      .filter(Boolean))];

    const rawExcludedTerminalModes = [
      ...(Array.isArray(item?.excludedTerminalModes) ? item.excludedTerminalModes : []),
      ...(Array.isArray(item?.excludedModes) ? item.excludedModes : [])
    ];
    const excludedTerminalModes = [...new Set(rawExcludedTerminalModes
      .map((mode) => String(mode || '').trim().toUpperCase())
      .filter(Boolean))];

    if (applicableTerminalModes.length > 0 && !applicableTerminalModes.includes(reportTerminalMode)) {
      return {
        applicable: false,
        reason: `当前模式 ${reportTerminalMode || 'unknown'} 不要求测试该项`,
        skipContext: {
          dimension: 'terminalMode',
          actual: reportTerminalMode || '',
          include: applicableTerminalModes,
          exclude: excludedTerminalModes
        }
      };
    }

    if (excludedTerminalModes.includes(reportTerminalMode)) {
      return {
        applicable: false,
        reason: `当前模式 ${reportTerminalMode} 不要求测试该项`,
        skipContext: {
          dimension: 'terminalMode',
          actual: reportTerminalMode,
          include: applicableTerminalModes,
          exclude: excludedTerminalModes
        }
      };
    }

    const reportNetwork = normalizeNetwork(reportData?.reportContext?.network);
    const applicableNetworks = Array.isArray(item?.applicableNetworks)
      ? item.applicableNetworks.map(normalizeNetwork).filter(Boolean)
      : [];
    const excludedNetworks = Array.isArray(item?.excludedNetworks)
      ? item.excludedNetworks.map(normalizeNetwork).filter(Boolean)
      : [];
    if (reportNetwork) {
      const networkApplicability = isTokenApplicable({
        reportValue: reportNetwork,
        allowedTokens: applicableNetworks,
        excludedTokens: excludedNetworks,
        reasonPrefix: '当前网络',
        dimension: 'network'
      });
      if (!networkApplicability.applicable) {
        return networkApplicability;
      }
    }

    const reportVocoder = buildNormalizedVocoder(reportData?.reportContext || {});
    const applicableVocoders = Array.isArray(item?.applicableVocoders)
      ? item.applicableVocoders.map((token) => String(token || '').trim().toUpperCase()).filter(Boolean)
      : [];
    const excludedVocoders = Array.isArray(item?.excludedVocoders)
      ? item.excludedVocoders.map((token) => String(token || '').trim().toUpperCase()).filter(Boolean)
      : [];
    if (reportVocoder.family) {
      const matchesToken = (token) => (token.includes('_') ? token === reportVocoder.full : token === reportVocoder.family);
      if (applicableVocoders.length > 0 && !applicableVocoders.some(matchesToken)) {
        return {
          applicable: false,
          reason: `当前 Vocoder ${reportVocoder.full || reportVocoder.family} 不要求测试该项`,
          skipContext: {
            dimension: 'vocoder',
            actual: reportVocoder.full || reportVocoder.family,
            include: applicableVocoders,
            exclude: excludedVocoders
          }
        };
      }

      if (excludedVocoders.some(matchesToken)) {
        return {
          applicable: false,
          reason: `当前 Vocoder ${reportVocoder.full || reportVocoder.family} 不要求测试该项`,
          skipContext: {
            dimension: 'vocoder',
            actual: reportVocoder.full || reportVocoder.family,
            include: applicableVocoders,
            exclude: excludedVocoders
          }
        };
      }
    }

    const reportCustomer = normalizeCustomer(reportData?.reportContext?.customer);
    const applicableCustomers = Array.isArray(item?.applicableCustomers)
      ? item.applicableCustomers.map(normalizeCustomer).filter(Boolean)
      : [];
    const excludedCustomers = Array.isArray(item?.excludedCustomers)
      ? item.excludedCustomers.map(normalizeCustomer).filter(Boolean)
      : [];
    if (reportCustomer) {
      const customerApplicability = isTokenApplicable({
        reportValue: reportCustomer,
        allowedTokens: applicableCustomers,
        excludedTokens: excludedCustomers,
        reasonPrefix: '当前客户',
        dimension: 'customer'
      });
      if (!customerApplicability.applicable) {
        return customerApplicability;
      }
    }

    const reportBandwidth = normalizeBandwidth(reportData?.reportContext?.bandwidth);
    if (!reportBandwidth) {
      return { applicable: true };
    }

    const applicableBandwidths = Array.isArray(item?.applicableBandwidths)
      ? item.applicableBandwidths.map(normalizeBandwidth).filter(Boolean)
      : [];
    const excludedBandwidths = Array.isArray(item?.excludedBandwidths)
      ? item.excludedBandwidths.map(normalizeBandwidth).filter(Boolean)
      : [];

    if (applicableBandwidths.length > 0 && !applicableBandwidths.includes(reportBandwidth)) {
      return {
        applicable: false,
        reason: `当前带宽 ${reportBandwidth} 不要求测试该项`,
        skipContext: {
          dimension: 'bandwidth',
          actual: reportBandwidth,
          include: applicableBandwidths,
          exclude: excludedBandwidths
        }
      };
    }

    if (excludedBandwidths.includes(reportBandwidth)) {
      return {
        applicable: false,
        reason: `当前带宽 ${reportBandwidth} 不要求测试该项`,
        skipContext: {
          dimension: 'bandwidth',
          actual: reportBandwidth,
          include: applicableBandwidths,
          exclude: excludedBandwidths
        }
      };
    }

    return { applicable: true };
  }

  async function processSingleReport({ reportPath, checklistPath, rules, customer, reportPanelSelections, reportPanelSelectionsOverride }) {
    const reportData = await parseReport(reportPath);
    const mergedReportContext = mergeReportContext(reportData?.reportContext || {}, {
      customer,
      reportPanelSelections: reportPanelSelectionsOverride || reportPanelSelections
    });
    const normalizedReportData = {
      ...reportData,
      reportContext: mergedReportContext
    };
    const reportKind = reportData?.reportFormat === 'xlsx' ? 'excel' : 'word';
    const checklistPathForReport = await resolveChecklistPathForReport(checklistPath, mergedReportContext, reportPath);
    const { activeRules, profileKey } = resolveRulesForReport(rules, normalizedReportData, checklistPathForReport);

    if (reportKind === 'word') {
      return {
        reportPath,
        reportKind,
        reportFormat: reportData?.reportFormat || 'docx',
        bundleKey: path.parse(reportPath).name,
        reportContext: mergedReportContext,
        ruleProfileKey: profileKey,
        outputPath: '',
        totalItems: 0,
        matchedItems: 0,
        skippedItems: [],
        unmatchedItems: [],
        extractedItems: [],
        audit: analyzeWordReport({ reportPath, reportData: normalizedReportData })
      };
    }

    if (!checklistPathForReport) {
      throw new Error('Excel 报告处理需要 checklist 文件。');
    }

    const textNormalizeConfig = activeRules.globalMatchConfig?.textNormalize || {};
    const globalMatchConfig = activeRules.globalMatchConfig || {};

    const extractedResultsByItemId = new Map();
    const extractedItems = activeRules.extractItemList.map((item) => {
      const applicability = isItemApplicable(normalizedReportData, item);
      if (!applicability.applicable) {
        const skippedResult = {
          itemId: item.itemId,
          checklistDesc: item.checklistDesc,
          outputCell: item.outputCell,
          extractType: item.extractType,
          matched: false,
          skipped: true,
          reason: applicability.reason,
          skipContext: applicability.skipContext || null
        };

        extractedResultsByItemId.set(item.itemId, skippedResult);
        return skippedResult;
      }

      let extractionResult;

      if (['summary_table_match', 'formula_calc', 'status_judge'].includes(item.extractType)) {
        extractionResult = resolveRowBasedValue(normalizedReportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);
      } else if (item.extractType === 'detail_anchor_extract') {
        extractionResult = resolveAnchorValue(normalizedReportData, item, textNormalizeConfig, globalMatchConfig, extractedResultsByItemId);
      } else if (item.extractType === 'table_struct_extract') {
        extractionResult = resolveTableValue(normalizedReportData, item, textNormalizeConfig);
      } else if (item.extractType === 'regex_extract') {
        extractionResult = resolveRegexValue(normalizedReportData, item, textNormalizeConfig, globalMatchConfig);
      } else {
        extractionResult = { matched: false, reason: `暂不支持的提取类型: ${item.extractType}` };
      }

      const itemResult = {
        itemId: item.itemId,
        checklistDesc: item.checklistDesc,
        outputCell: item.outputCell,
        extractType: item.extractType,
        ...extractionResult
      };

      extractedResultsByItemId.set(item.itemId, itemResult);
      return itemResult;
    });

    const outputPath = await applyResultsToChecklist(checklistPathForReport, reportPath, extractedItems, mergedReportContext);
    const matchedItems = extractedItems.filter((item) => item.matched).length;
    const skippedItems = extractedItems.filter((item) => item.skipped).map((item) => ({
      itemId: item.itemId,
      checklistDesc: item.checklistDesc,
      outputCell: item.outputCell,
      reason: item.reason || '当前场景不要求测试',
      skipContext: item.skipContext || null
    }));
    const unmatchedItems = extractedItems.filter((item) => !item.matched && !item.skipped).map((item) => ({
      itemId: item.itemId,
      checklistDesc: item.checklistDesc,
      outputCell: item.outputCell,
      reason: item.reason || '未命中'
    }));

    return {
      reportPath,
      reportKind,
      reportFormat: reportData?.reportFormat || 'xlsx',
      bundleKey: path.parse(reportPath).name,
      reportContext: mergedReportContext,
      checklistPathUsed: checklistPathForReport,
      ruleProfileKey: profileKey,
      outputPath,
      totalItems: extractedItems.length,
      matchedItems,
      skippedItems,
      unmatchedItems,
      extractedItems,
      audit: analyzeExcelReport({ reportPath, reportData: normalizedReportData, extractedItems })
    };
  }

  return {
    processSingleReport,
    inspectReportContext
  };
}

module.exports = {
  createReportExtractor
};
