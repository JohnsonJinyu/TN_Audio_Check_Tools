const Store = require('electron-store');

const THEME_OPTIONS = new Set(['auto', 'light', 'dark']);
const LANGUAGE_OPTIONS = new Set(['zh-cn', 'zh-tw', 'en-us']);
const AUDIO_FORMAT_OPTIONS = new Set(['mp3', 'wav', 'flac', 'aac']);
const BITRATE_OPTIONS = new Set(['128', '192', '256', '320']);
const SAMPLE_RATE_OPTIONS = new Set(['original', '44100', '48000', '96000']);
const CONCURRENCY_OPTIONS = new Set([1, 2, 4, 8]);

const DEFAULT_APP_SETTINGS = Object.freeze({
  appearance: {
    theme: 'auto',
    language: 'zh-cn'
  },
  system: {
    enableTray: false,
    launchMinimizedToTray: false
  },
  files: {
    defaultOutputDirectory: '',
    maxConcurrentTasks: 4
  },
  audio: {
    defaultOutputFormat: 'mp3',
    defaultBitrate: '192',
    defaultSampleRate: '44100'
  }
});

const store = new Store({
  name: 'app-settings',
  defaults: cloneDefaults()
});

function cloneDefaults() {
  return JSON.parse(JSON.stringify(DEFAULT_APP_SETTINGS));
}

function selectAllowedValue(value, allowedValues, fallbackValue) {
  return allowedValues.has(value) ? value : fallbackValue;
}

function normalizeBoolean(value) {
  return Boolean(value);
}

function normalizeMaxConcurrentTasks(value) {
  const numericValue = Number(value);
  if (!Number.isFinite(numericValue)) {
    return DEFAULT_APP_SETTINGS.files.maxConcurrentTasks;
  }

  return CONCURRENCY_OPTIONS.has(numericValue)
    ? numericValue
    : DEFAULT_APP_SETTINGS.files.maxConcurrentTasks;
}

function normalizeOutputDirectory(value) {
  return typeof value === 'string' ? value.trim() : '';
}

function normalizeSettings(input = {}) {
  const normalized = {
    appearance: {
      theme: selectAllowedValue(
        input?.appearance?.theme,
        THEME_OPTIONS,
        DEFAULT_APP_SETTINGS.appearance.theme
      ),
      language: selectAllowedValue(
        input?.appearance?.language,
        LANGUAGE_OPTIONS,
        DEFAULT_APP_SETTINGS.appearance.language
      )
    },
    system: {
      enableTray: normalizeBoolean(input?.system?.enableTray),
      launchMinimizedToTray: normalizeBoolean(input?.system?.launchMinimizedToTray)
    },
    files: {
      defaultOutputDirectory: normalizeOutputDirectory(input?.files?.defaultOutputDirectory),
      maxConcurrentTasks: normalizeMaxConcurrentTasks(input?.files?.maxConcurrentTasks)
    },
    audio: {
      defaultOutputFormat: selectAllowedValue(
        input?.audio?.defaultOutputFormat,
        AUDIO_FORMAT_OPTIONS,
        DEFAULT_APP_SETTINGS.audio.defaultOutputFormat
      ),
      defaultBitrate: selectAllowedValue(
        input?.audio?.defaultBitrate,
        BITRATE_OPTIONS,
        DEFAULT_APP_SETTINGS.audio.defaultBitrate
      ),
      defaultSampleRate: selectAllowedValue(
        input?.audio?.defaultSampleRate,
        SAMPLE_RATE_OPTIONS,
        DEFAULT_APP_SETTINGS.audio.defaultSampleRate
      )
    }
  };

  if (!normalized.system.enableTray) {
    normalized.system.launchMinimizedToTray = false;
  }

  return normalized;
}

function getSettings() {
  return normalizeSettings(store.store || {});
}

function saveSettings(nextSettings) {
  const normalizedSettings = normalizeSettings(nextSettings);
  store.store = normalizedSettings;
  return normalizedSettings;
}

function resetSettings() {
  const normalizedDefaults = cloneDefaults();
  store.store = normalizedDefaults;
  return normalizedDefaults;
}

module.exports = {
  DEFAULT_APP_SETTINGS,
  getSettings,
  normalizeSettings,
  resetSettings,
  saveSettings
};