const { app, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const log = require('electron-log');

let initialized = false;
let getMainWindow = () => null;
let currentCheckSource = 'idle';
let promptedAvailableVersion = null;
let promptedDownloadedVersion = null;

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

  return error.message || '检查更新时发生未知错误。';
}

async function promptForDownload(info) {
  if (promptedAvailableVersion === info.version) {
    return;
  }

  promptedAvailableVersion = info.version;

  const result = await dialog.showMessageBox(getWindow(), {
    type: 'info',
    buttons: ['立即下载', '稍后'],
    defaultId: 0,
    cancelId: 1,
    noLink: true,
    title: '发现新版本',
    message: `检测到新版本 ${info.version}`,
    detail: '是否现在开始下载更新安装包？'
  });

  if (result.response === 0) {
    await downloadUpdate({ source: 'auto' });
  }
}

async function promptForInstall(info) {
  if (promptedDownloadedVersion === info.version) {
    return;
  }

  promptedDownloadedVersion = info.version;

  const result = await dialog.showMessageBox(getWindow(), {
    type: 'info',
    buttons: ['立即安装', '稍后'],
    defaultId: 0,
    cancelId: 1,
    noLink: true,
    title: '更新已下载完成',
    message: `新版本 ${info.version} 已下载完成`,
    detail: '立即重启应用并完成安装？'
  });

  if (result.response === 0) {
    quitAndInstallUpdate();
  }
}

function handleUpdaterError(error) {
  const message = getErrorMessage(error);
  log.error('[updater] update failed:', message);
  setState({
    status: updateState.available ? 'available' : 'idle',
    checking: false,
    downloading: false,
    error: message
  });
  currentCheckSource = 'idle';
  return { ok: false, error: message };
}

function registerAutoUpdaterEvents() {
  autoUpdater.on('checking-for-update', () => {
    log.info('[updater] checking for update');
    setState({
      status: 'checking',
      checking: true,
      error: null,
      lastCheckedAt: new Date().toISOString(),
      lastCheckSource: currentCheckSource,
      unsupported: false
    });
  });

  autoUpdater.on('update-available', (info) => {
    log.info('[updater] update available:', info.version);
    setState({
      status: 'available',
      checking: false,
      available: true,
      downloading: false,
      downloaded: false,
      installReady: false,
      latestVersion: info.version,
      releaseName: info.releaseName || null,
      releaseDate: info.releaseDate || null,
      releaseNotes: normalizeReleaseNotes(info.releaseNotes),
      progressPercent: 0,
      transferred: 0,
      total: 0,
      bytesPerSecond: 0,
      error: null
    });

    const triggerSource = currentCheckSource;
    currentCheckSource = 'idle';

    if (triggerSource === 'auto') {
      promptForDownload(info).catch((error) => {
        log.error('[updater] failed to show download prompt:', error);
      });
    }
  });

  autoUpdater.on('update-not-available', (info) => {
    log.info('[updater] update not available');
    setState({
      status: 'up-to-date',
      checking: false,
      available: false,
      downloading: false,
      downloaded: false,
      installReady: false,
      latestVersion: info?.version || app.getVersion(),
      releaseName: info?.releaseName || null,
      releaseDate: info?.releaseDate || null,
      releaseNotes: normalizeReleaseNotes(info?.releaseNotes),
      progressPercent: 0,
      transferred: 0,
      total: 0,
      bytesPerSecond: 0,
      error: null
    });
    currentCheckSource = 'idle';
  });

  autoUpdater.on('download-progress', (progress) => {
    setState({
      status: 'downloading',
      checking: false,
      downloading: true,
      downloaded: false,
      installReady: false,
      progressPercent: Number(progress.percent || 0).toFixed(1),
      transferred: progress.transferred || 0,
      total: progress.total || 0,
      bytesPerSecond: progress.bytesPerSecond || 0,
      error: null
    });
  });

  autoUpdater.on('update-downloaded', (info) => {
    log.info('[updater] update downloaded:', info.version);
    setState({
      status: 'downloaded',
      checking: false,
      available: true,
      downloading: false,
      downloaded: true,
      installReady: true,
      latestVersion: info.version,
      releaseName: info.releaseName || null,
      releaseDate: info.releaseDate || null,
      releaseNotes: normalizeReleaseNotes(info.releaseNotes),
      progressPercent: 100,
      lastDownloadedAt: new Date().toISOString(),
      error: null
    });

    promptForInstall(info).catch((error) => {
      log.error('[updater] failed to show install prompt:', error);
    });
  });

  autoUpdater.on('error', (error) => {
    handleUpdaterError(error);
  });
}

function initializeUpdateService(options = {}) {
  if (initialized) {
    return;
  }

  getMainWindow = options.getMainWindow || (() => null);
  log.transports.file.level = 'info';
  autoUpdater.logger = log;
  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = false;
  autoUpdater.allowPrerelease = false;
  autoUpdater.allowDowngrade = false;

  registerAutoUpdaterEvents();
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

  try {
    await autoUpdater.checkForUpdates();
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

  if (updateState.downloaded) {
    return { ok: true, alreadyDownloaded: true };
  }

  if (updateState.downloading) {
    return { ok: false, busy: true, state: updateState };
  }

  setState({
    status: 'downloading',
    checking: false,
    downloading: true,
    error: null
  });

  try {
    await autoUpdater.downloadUpdate();
    return { ok: true, source };
  } catch (error) {
    return handleUpdaterError(error);
  }
}

function quitAndInstallUpdate() {
  if (!updateState.downloaded) {
    return { ok: false, message: '当前没有已下载完成的更新。' };
  }

  setState({
    status: 'installing',
    checking: false,
    downloading: false,
    error: null
  });

  setImmediate(() => {
    autoUpdater.quitAndInstall(false, true);
  });

  return { ok: true };
}

function getUpdateState() {
  return updateState;
}

module.exports = {
  initializeUpdateService,
  checkForUpdates,
  downloadUpdate,
  quitAndInstallUpdate,
  getUpdateState
};