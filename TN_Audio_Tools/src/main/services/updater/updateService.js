const { app, dialog, shell } = require('electron');
const axios = require('axios');
const log = require('electron-log');

const UPDATE_PROVIDER = Object.freeze({
  owner: 'JohnsonJinyu',
  repo: 'TN_Audio_Check_Tools'
});
const DEFAULT_DOWNLOAD_MIRROR = process.env.TN_AUDIO_UPDATE_MIRROR || 'https://ghfast.top/';
const DEFAULT_UPDATE_MANIFEST_FILE = 'update-manifest.json';
const DEFAULT_UPDATE_MANIFEST_PATH = `TN_Audio_Tools/${DEFAULT_UPDATE_MANIFEST_FILE}`;
const DEFAULT_UPDATE_REQUEST_TIMEOUT = 8000;
const DEFAULT_UPDATE_MANIFEST_URLS = [
  `https://gitee.com/lingyu_mayun/${UPDATE_PROVIDER.repo}/raw/master/${DEFAULT_UPDATE_MANIFEST_PATH}`,
  `https://cdn.jsdelivr.net/gh/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}@master/${DEFAULT_UPDATE_MANIFEST_PATH}`,
  `${normalizeMirrorPrefix(DEFAULT_DOWNLOAD_MIRROR)}https://raw.githubusercontent.com/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}/master/${DEFAULT_UPDATE_MANIFEST_PATH}`,
  `https://raw.githubusercontent.com/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}/master/${DEFAULT_UPDATE_MANIFEST_PATH}`
].filter(Boolean);

let initialized = false;
let getMainWindow = () => null;
let currentCheckSource = 'idle';
let promptedAvailableVersion = null;

function normalizeMirrorPrefix(prefix) {
  if (typeof prefix !== 'string') {
    return '';
  }

  const trimmed = prefix.trim();
  if (!trimmed) {
    return '';
  }

  return trimmed.endsWith('/') ? trimmed : `${trimmed}/`;
}

function getConfiguredManifestUrls() {
  const rawValue = process.env.TN_AUDIO_UPDATE_CONFIG_URLS || process.env.TN_AUDIO_UPDATE_CONFIG_URL || '';
  const configuredUrls = rawValue
    .split(/[\n,;]/)
    .map((entry) => entry.trim())
    .filter(Boolean);

  return configuredUrls.length ? configuredUrls : DEFAULT_UPDATE_MANIFEST_URLS;
}

function buildMirrorUrl(rawUrl) {
  if (!rawUrl) {
    return null;
  }

  const prefix = normalizeMirrorPrefix(DEFAULT_DOWNLOAD_MIRROR);
  return prefix ? `${prefix}${rawUrl}` : rawUrl;
}

function getReleaseTag(version) {
  if (!version) {
    return null;
  }

  return String(version).startsWith('v') ? String(version) : `v${version}`;
}

function encodePathSegment(value) {
  return encodeURIComponent(value).replace(/%2F/g, '/');
}

function resolveAssetName(info = {}) {
  if (info?.files?.[0]?.url || info?.path) {
    return info?.files?.[0]?.url || info?.path || null;
  }

  const downloadUrl = info.downloadUrl || info.mirrorUrl || null;
  if (!downloadUrl) {
    return null;
  }

  try {
    const pathname = new URL(downloadUrl).pathname || '';
    const token = pathname.split('/').filter(Boolean).pop();
    return token ? decodeURIComponent(token) : null;
  } catch (error) {
    return null;
  }
}

function buildUpdateLinks(info = {}) {
  const version = info.version || null;
  const tag = getReleaseTag(version);
  const assetName = resolveAssetName(info);

  if (!tag) {
    return {
      assetName: null,
      githubDownloadUrl: null,
      externalDownloadUrl: null,
      releasePageUrl: null,
      mirrorName: normalizeMirrorPrefix(DEFAULT_DOWNLOAD_MIRROR) ? 'ghfast' : null
    };
  }

  const releasePageUrl = `https://github.com/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}/releases/tag/${encodePathSegment(tag)}`;
  const fallbackGithubDownloadUrl = assetName
    ? `https://github.com/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}/releases/download/${encodePathSegment(tag)}/${encodePathSegment(assetName)}`
    : null;
  const githubDownloadUrl = info.downloadUrl || fallbackGithubDownloadUrl;
  const externalDownloadUrl = info.mirrorUrl || buildMirrorUrl(githubDownloadUrl);

  return {
    assetName,
    githubDownloadUrl,
    externalDownloadUrl,
    releasePageUrl,
    mirrorName: normalizeMirrorPrefix(DEFAULT_DOWNLOAD_MIRROR) ? 'ghfast' : null
  };
}

function stripVersionPrefix(version) {
  return String(version || '').trim().replace(/^v/i, '');
}

function compareVersions(left, right) {
  const leftParts = stripVersionPrefix(left).split('.').map((part) => Number.parseInt(part, 10) || 0);
  const rightParts = stripVersionPrefix(right).split('.').map((part) => Number.parseInt(part, 10) || 0);
  const total = Math.max(leftParts.length, rightParts.length);

  for (let index = 0; index < total; index += 1) {
    const leftValue = leftParts[index] || 0;
    const rightValue = rightParts[index] || 0;

    if (leftValue > rightValue) {
      return 1;
    }

    if (leftValue < rightValue) {
      return -1;
    }
  }

  return 0;
}

function ensureReleaseNotes(value) {
  if (!value) {
    return null;
  }

  if (Array.isArray(value)) {
    return value.filter(Boolean).join('\n');
  }

  if (typeof value === 'string') {
    return value.trim() || null;
  }

  return null;
}

function normalizeRemoteManifest(payload = {}, sourceUrl) {
  const version = stripVersionPrefix(payload.version);
  if (!version) {
    throw new Error('远程版本清单缺少 version 字段。');
  }

  const releaseName = payload.releaseName || payload.name || null;
  const releaseDate = payload.releaseDate || payload.publishedAt || null;
  const releaseNotes = ensureReleaseNotes(payload.releaseNotes || payload.notes || payload.changelog);
  const downloadUrl = payload.downloadUrl || payload.releasePageUrl || null;
  const mirrorUrl = payload.mirrorUrl || payload.externalDownloadUrl || null;
  const info = {
    version,
    releaseName,
    releaseDate,
    releaseNotes,
    downloadUrl,
    mirrorUrl
  };
  const links = buildUpdateLinks(info);

  return {
    version,
    releaseName,
    releaseDate,
    releaseNotes,
    mandatory: Boolean(payload.mandatory),
    sourceUrl,
    ...links
  };
}

async function fetchRemoteManifest() {
  const manifestUrls = getConfiguredManifestUrls();
  const errors = [];

  for (const manifestUrl of manifestUrls) {
    try {
      const response = await axios.get(manifestUrl, {
        timeout: DEFAULT_UPDATE_REQUEST_TIMEOUT,
        responseType: 'json',
        headers: {
          Accept: 'application/json, text/plain, */*',
          'Cache-Control': 'no-cache'
        }
      });

      return normalizeRemoteManifest(response.data || {}, manifestUrl);
    } catch (error) {
      const message = getErrorMessage(error);
      errors.push(`${manifestUrl}: ${message}`);
      log.warn('[updater] failed to fetch manifest from', manifestUrl, message);
    }
  }

  const detail = errors.length ? `已尝试地址：${errors.join(' | ')}` : '未配置远程版本清单地址。';
  throw new Error(`无法获取远程版本信息。${detail}`);
}

function isPortableApp() {
  return Boolean(process.env.PORTABLE_EXECUTABLE_FILE || process.env.PORTABLE_EXECUTABLE_DIR);
}

function isUpdaterSupported() {
  return process.platform === 'win32' && app.isPackaged && !isPortableApp();
}

function getUnsupportedMessage() {
  if (!app.isPackaged) {
    return '开发模式不支持在线更新，请使用安装版进行验证。';
  }

  if (isPortableApp()) {
    return '便携版不支持在线更新，请改用安装版。';
  }

  return '当前运行环境不支持在线更新。';
}

function createInitialState() {
  const supported = isUpdaterSupported();
  return {
    status: supported ? 'idle' : 'unsupported',
    currentVersion: app.getVersion(),
    latestVersion: null,
    releaseName: null,
    releaseDate: null,
    releaseNotes: null,
    checking: false,
    available: false,
    downloading: false,
    downloaded: false,
    installReady: false,
    unsupported: !supported,
    progressPercent: 0,
    transferred: 0,
    total: 0,
    bytesPerSecond: 0,
    assetName: null,
    githubDownloadUrl: null,
    externalDownloadUrl: null,
    releasePageUrl: null,
    mirrorName: null,
    lastCheckedAt: null,
    lastCheckSource: null,
    lastDownloadedAt: null,
    error: supported ? null : getUnsupportedMessage()
  };
}

let updateState = createInitialState();

function getWindow() {
  const candidate = typeof getMainWindow === 'function' ? getMainWindow() : null;
  if (!candidate || candidate.isDestroyed()) {
    return null;
  }

  return candidate;
}

function emitState() {
  const windowRef = getWindow();
  if (!windowRef) {
    return;
  }

  windowRef.webContents.send('app-update:state-changed', updateState);
}

function setState(patch) {
  updateState = {
    ...updateState,
    ...patch
  };
  emitState();
}

function normalizeReleaseNotes(releaseNotes) {
  if (!releaseNotes) {
    return null;
  }

  if (typeof releaseNotes === 'string') {
    return releaseNotes;
  }

  if (Array.isArray(releaseNotes)) {
    return releaseNotes
      .map((entry) => entry?.note || '')
      .filter(Boolean)
      .join('\n\n');
  }

  return null;
}

function getErrorMessage(error) {
  if (!error) {
    return '检查更新时发生未知错误。';
  }

  if (typeof error === 'string') {
    return error;
  }

  if (error.stack) {
    log.error('[updater] error stack:', error.stack);
  }

  if (error.response?.data?.error) {
    return String(error.response.data.error);
  }

  return error.message || '检查更新时发生未知错误。';
}

async function promptForBrowserDownload() {
  const targetUrl = updateState.externalDownloadUrl || updateState.githubDownloadUrl || updateState.releasePageUrl;
  const targetVersion = updateState.latestVersion || null;

  if (!targetUrl) {
    return false;
  }

  if (targetVersion && promptedAvailableVersion === targetVersion) {
    return false;
  }

  promptedAvailableVersion = targetVersion;

  const result = await dialog.showMessageBox(getWindow(), {
    type: 'info',
    buttons: ['打开下载页', '稍后'],
    defaultId: 0,
    cancelId: 1,
    noLink: true,
    title: '发现新版本',
    message: `检测到新版本 ${targetVersion}`,
    detail: updateState.externalDownloadUrl
      ? '是否现在在浏览器中打开下载地址？'
      : '是否现在在浏览器中打开发布页？'
  });

  if (result.response === 0) {
    await shell.openExternal(targetUrl);
    return true;
  }

  return false;
}

async function promptForDownload(info) {
  return promptForBrowserDownload(info);
}

function handleUpdaterError(error) {
  const message = getErrorMessage(error);
  log.error('[updater] update failed:', message);
  setState({
    status: updateState.available ? 'available' : 'idle',
    checking: false,
    downloading: false,
    downloaded: false,
    installReady: false,
    error: message
  });
  currentCheckSource = 'idle';

  return { ok: false, error: message };
}

function initializeUpdateService(options = {}) {
  if (initialized) {
    return;
  }

  getMainWindow = options.getMainWindow || (() => null);
  log.transports.file.level = 'info';
  initialized = true;
  emitState();
}

async function checkForUpdates(options = {}) {
  const source = options.manual ? 'manual' : 'auto';

  if (!isUpdaterSupported()) {
    const message = getUnsupportedMessage();
    setState({
      status: 'unsupported',
      checking: false,
      unsupported: true,
      lastCheckSource: source,
      lastCheckedAt: new Date().toISOString(),
      error: options.manual ? message : null
    });
    return { ok: false, unsupported: true, message };
  }

  if (updateState.checking) {
    return { ok: false, busy: true, state: updateState };
  }

  currentCheckSource = source;
  setState({
    status: 'checking',
    checking: true,
    available: false,
    downloading: false,
    downloaded: false,
    installReady: false,
    error: null,
    lastCheckedAt: new Date().toISOString(),
    lastCheckSource: source,
    unsupported: false
  });

  try {
    const remoteInfo = await fetchRemoteManifest();
    const isNewVersion = compareVersions(remoteInfo.version, app.getVersion()) > 0;

    if (isNewVersion) {
      setState({
        status: 'available',
        checking: false,
        available: true,
        downloading: false,
        downloaded: false,
        installReady: false,
        latestVersion: remoteInfo.version,
        releaseName: remoteInfo.releaseName,
        releaseDate: remoteInfo.releaseDate,
        releaseNotes: normalizeReleaseNotes(remoteInfo.releaseNotes),
        assetName: remoteInfo.assetName,
        githubDownloadUrl: remoteInfo.githubDownloadUrl,
        externalDownloadUrl: remoteInfo.externalDownloadUrl,
        releasePageUrl: remoteInfo.releasePageUrl,
        mirrorName: remoteInfo.externalDownloadUrl ? (remoteInfo.mirrorName || '下载源') : null,
        progressPercent: 0,
        transferred: 0,
        total: 0,
        bytesPerSecond: 0,
        error: null
      });

      const triggerSource = currentCheckSource;
      currentCheckSource = 'idle';

      if (triggerSource === 'auto') {
        promptForDownload(remoteInfo).catch((promptError) => {
          log.error('[updater] failed to show browser download prompt:', promptError);
        });
      }

      return { ok: true, available: true, version: remoteInfo.version };
    }

    const currentVersion = app.getVersion();
    setState({
      status: 'up-to-date',
      checking: false,
      available: false,
      downloading: false,
      downloaded: false,
      installReady: false,
      latestVersion: currentVersion,
      releaseName: remoteInfo.releaseName,
      releaseDate: remoteInfo.releaseDate,
      releaseNotes: normalizeReleaseNotes(remoteInfo.releaseNotes),
      assetName: remoteInfo.assetName,
      githubDownloadUrl: remoteInfo.githubDownloadUrl,
      externalDownloadUrl: remoteInfo.externalDownloadUrl,
      releasePageUrl: remoteInfo.releasePageUrl,
      mirrorName: remoteInfo.externalDownloadUrl ? (remoteInfo.mirrorName || '下载源') : null,
      progressPercent: 0,
      transferred: 0,
      total: 0,
      bytesPerSecond: 0,
      error: null
    });
    currentCheckSource = 'idle';
    return { ok: true };
  } catch (error) {
    return handleUpdaterError(error);
  }
}

async function downloadUpdate(options = {}) {
  const source = options.source || 'manual';

  if (!isUpdaterSupported()) {
    const message = getUnsupportedMessage();
    setState({
      status: 'unsupported',
      checking: false,
      downloading: false,
      unsupported: true,
      error: message
    });
    return { ok: false, unsupported: true, message };
  }

  try {
    const targetUrl = getExternalDownloadUrl(true);
    if (!targetUrl) {
      return { ok: false, message: '当前没有可用的下载地址。' };
    }

    await shell.openExternal(targetUrl);
    return { ok: true, source, openedInBrowser: true, url: targetUrl };
  } catch (error) {
    return handleUpdaterError(error);
  }
}

function quitAndInstallUpdate() {
  return {
    ok: false,
    message: '当前版本采用浏览器下载模式，请在下载完成后手动运行安装包。'
  };
}

function getUpdateState() {
  return updateState;
}

function getExternalDownloadUrl(preferMirror = true) {
  if (preferMirror && updateState.externalDownloadUrl) {
    return updateState.externalDownloadUrl;
  }

  return updateState.githubDownloadUrl || updateState.releasePageUrl || null;
}

module.exports = {
  initializeUpdateService,
  checkForUpdates,
  downloadUpdate,
  quitAndInstallUpdate,
  getUpdateState,
  getExternalDownloadUrl
};