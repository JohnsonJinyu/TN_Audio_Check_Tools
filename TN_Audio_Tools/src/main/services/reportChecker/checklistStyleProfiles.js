const SKIPPED_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFD9D9D9' }
};

const STYLE_PROFILES = {
  handset: {
    key: 'handset',
    skippedFill: SKIPPED_FILL
  },
  handsfree: {
    key: 'handsfree',
    skippedFill: SKIPPED_FILL
  },
  headset: {
    key: 'headset',
    skippedFill: SKIPPED_FILL
  }
};

function normalizeModeKey(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!normalized) {
    return '';
  }

  if (['ha', 'handset'].includes(normalized)) {
    return 'handset';
  }

  if (['hh', 'hf', 'handsfree'].includes(normalized)) {
    return 'handsfree';
  }

  if (['he', 'hs', 'headset'].includes(normalized)) {
    return 'headset';
  }

  return normalized;
}

function resolveChecklistStyleProfile({ reportContext, sheetName } = {}) {
  const reportModeKey = normalizeModeKey(reportContext?.terminalMode);
  if (STYLE_PROFILES[reportModeKey]) {
    return STYLE_PROFILES[reportModeKey];
  }

  const sheetModeKey = normalizeModeKey(sheetName);
  if (STYLE_PROFILES[sheetModeKey]) {
    return STYLE_PROFILES[sheetModeKey];
  }

  return STYLE_PROFILES.handset;
}

module.exports = {
  resolveChecklistStyleProfile
};