const fs = require('fs');
const fsPromises = require('fs/promises');
const path = require('path');
const { spawn } = require('child_process');
const { app, dialog, shell } = require('electron');
const axios = require('axios');
const log = require('electron-log');

const UPDATE_PROVIDER = Object.freeze({
  owner: 'JohnsonJinyu',
  repo: 'TN_Audio_Check_Tools'
});

const DEFAULT_UPDATE_MANIFEST_FILE = 'update-manifest.json';
const DEFAULT_UPDATE_MANIFEST_PATH = `TN_Audio_Tools/${DEFAULT_UPDATE_MANIFEST_FILE}`;
const DEFAULT_MANIFEST_URLS = Object.freeze([
  `https://gitee.com/lingyu_mayun/${UPDATE_PROVIDER.repo}/raw/master/${DEFAULT_UPDATE_MANIFEST_PATH}`,
  `https://raw.githubusercontent.com/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}/master/${DEFAULT_UPDATE_MANIFEST_PATH}`,
  `https://cdn.jsdelivr.net/gh/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}@master/${DEFAULT_UPDATE_MANIFEST_PATH}`
]);
const DEFAULT_DOWNLOAD_MIRROR = process.env.TN_AUDIO_UPDATE_MIRROR || 'https://ghfast.top/';
const DEFAULT_REQUEST_TIMEOUT = 8000;
const DEFAULT_PROBE_TIMEOUT = 2500;
const UPDATE_DOWNLOAD_DIRECTORY = 'updates';

let initialized = false;
let getMainWindow = () => null;
let currentCheckSource = 'idle';
let promptedAvailableVersion = null;
let promptedDownloadedVersion = null;
let promptedExternalDownloadVersion = null;
let remoteManifest = null;
let downloadedInstallerPath = null;

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

function buildMirrorUrl(rawUrl) {
  if (!rawUrl) {
    return null;
  }

  const prefix = normalizeMirrorPrefix(DEFAULT_DOWNLOAD_MIRROR);
  return prefix ? `${prefix}${rawUrl}` : rawUrl;
}

function stripVersionPrefix(version) {
  return String(version || '').trim().replace(/^v/i, '');
}

function compareVersions(left, right) {
  const leftParts = stripVersionPrefix(left).split('.').map((part) => Number.parseInt(part, 10) || 0);
  const rightParts = stripVersionPrefix(right).split('.').map((part) => Number.parseInt(part, 10) || 0);
  const maxLength = Math.max(leftParts.length, rightParts.length);

  for (let index = 0; index < maxLength; index += 1) {
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

function getReleaseTag(version) {
  if (!version) {
    return null;
  }

  return String(version).startsWith('v') ? String(version) : `v${version}`;
}

function encodePathSegment(value) {
  return encodeURIComponent(value).replace(/%2F/g, '/');
}

function getRepositoryReleasePage(version) {
  const tag = getReleaseTag(version);
  if (!tag) {
    return `https://github.com/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}/releases`;
  }

  return `https://github.com/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}/releases/tag/${encodePathSegment(tag)}`;
}

function getDefaultInstallerName(version) {
  if (!version) {
    return null;
  }

  return `TN Audio Toolkit Setup ${stripVersionPrefix(version)}.exe`;
}

function getDefaultDownloadCandidates(version, assetName) {
  const tag = getReleaseTag(version);
  const fileName = assetName || getDefaultInstallerName(version);

  if (!tag || !fileName) {
    return [];
  }

  const githubUrl = `https://github.com/${UPDATE_PROVIDER.owner}/${UPDATE_PROVIDER.repo}/releases/download/${encodePathSegment(tag)}/${encodePathSegment(fileName)}`;
  const mirroredUrl = buildMirrorUrl(githubUrl);
  const candidates = [];

  if (mirroredUrl) {
    candidates.push({
      url: mirroredUrl,
      name: 'GitHub 镜像加速',
      channel: 'mirror'
    });
  }

  candidates.push({
    url: githubUrl,
    name: 'GitHub Release',
    channel: 'direct'
  });

  return candidates;
}

function parseListEnv(value) {
  return String(value || '')
    .split(/[\n,;]/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function getConfiguredManifestUrls() {
  const configuredUrls = [
    ...parseListEnv(process.env.TN_AUDIO_UPDATE_CONFIG_URLS),
    ...parseListEnv(process.env.TN_AUDIO_UPDATE_CONFIG_URL)
  ];

  const merged = configuredUrls.length > 0 ? configuredUrls : [...DEFAULT_MANIFEST_URLS];
  return [...new Set(merged)];
}

function normalizeReleaseNotes(releaseNotes) {
  if (!releaseNotes) {
    return null;
  }

  if (typeof releaseNotes === 'string') {
    return releaseNotes.trim() || null;
  }

  if (Array.isArray(releaseNotes)) {
    return releaseNotes
      .map((entry) => {
        if (typeof entry === 'string') {
          return entry.trim();
        }

        if (entry && typeof entry === 'object') {
          return String(entry.note || entry.text || '').trim();
        }

        return '';
      })
      .filter(Boolean)
      .join('\n\n') || null;
  }

  return null;
}

function normalizeDownloadEntry(entry, index = 0) {
  if (!entry) {
    return null;
  }

  if (typeof entry === 'string') {
    const trimmed = entry.trim();
    return trimmed
      ? { url: trimmed, name: index === 0 ? '主下载源' : `备用下载源 ${index + 1}`, channel: index === 0 ? 'primary' : 'fallback' }
      : null;
  }

  if (typeof entry !== 'object') {
    return null;
  }

  const url = String(entry.url || entry.href || '').trim();
  if (!url) {
    return null;
  }

  return {
    url,
    name: String(entry.name || entry.label || (index === 0 ? '主下载源' : `备用下载源 ${index + 1}`)).trim(),
    channel: String(entry.channel || entry.type || (index === 0 ? 'primary' : 'fallback')).trim()
  };
}

function dedupeDownloadEntries(entries) {
  const seen = new Set();
  const normalizedEntries = [];

  entries.forEach((entry) => {
    if (!entry?.url || seen.has(entry.url)) {
      return;
    }

    seen.add(entry.url);
    normalizedEntries.push(entry);
  });

  return normalizedEntries;
}

function ensureManifestShape(payload, manifestUrl) {
  if (!payload || typeof payload !== 'object') {
    throw new Error(`远程清单格式无效: ${manifestUrl}`);
  }

  const version = stripVersionPrefix(payload.version);
  if (!version) {
    throw new Error(`远程清单缺少 version 字段: ${manifestUrl}`);
  }

  const releaseName = payload.releaseName || `TN Audio Toolkit ${version}`;
  const releasePageUrl = String(payload.releasePageUrl || payload.releaseUrl || getRepositoryReleasePage(version)).trim();
  const assetName = String(payload.assetName || payload.fileName || getDefaultInstallerName(version) || '').trim() || null;
  const explicitDownloads = Array.isArray(payload.downloads)
    ? payload.downloads
    : [payload.downloadUrl, payload.mirrorUrl].filter(Boolean);
  const downloads = dedupeDownloadEntries([
    ...explicitDownloads.map((entry, index) => normalizeDownloadEntry(entry, index)).filter(Boolean),
    ...getDefaultDownloadCandidates(version, assetName)
  ]);

  if (downloads.length === 0) {
    throw new Error(`远程清单缺少可用下载地址: ${manifestUrl}`);
  }

  const directCandidate = downloads.find((entry) => entry.channel !== 'mirror') || downloads[0];
  const mirrorCandidate = downloads.find((entry) => entry.channel === 'mirror') || null;

  return {
    version,
    releaseName,
    releaseDate: payload.releaseDate || null,
    releaseNotes: normalizeReleaseNotes(payload.releaseNotes),
    mandatory: Boolean(payload.mandatory),
    manifestUrl,
    assetName,
    releasePageUrl,
    downloads,
    githubDownloadUrl: directCandidate?.url || null,
    externalDownloadUrl: mirrorCandidate?.url || directCandidate?.url || releasePageUrl,
    mirrorName: mirrorCandidate?.name || null
  };
}

function isUpdateCheckSupported() {
  return process.platform === 'win32';
}

function getUnsupportedMessage() {
  if (process.platform !== 'win32') {
    return '当前仅支持 Windows 平台在线更新。';
  }

  return null;
}

function createInitialState() {
  const supported = isUpdateCheckSupported();

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
    downloadSourceName: null,
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

function toReadableUrlError(url, error) {
  const message = getErrorMessage(error);
  return `${url}: ${message}`;
}

async function promptForExternalDownload() {
  const targetUrl = getExternalDownloadUrl(true);
  const targetVersion = updateState.latestVersion || null;

  if (!targetUrl) {
    return false;
  }

  if (targetVersion && promptedExternalDownloadVersion === targetVersion) {
    return false;
  }

  promptedExternalDownloadVersion = targetVersion;

  const result = await dialog.showMessageBox(getWindow(), {
    type: 'warning',
    buttons: ['打开浏览器下载', '稍后'],
    defaultId: 0,
    cancelId: 1,
    noLink: true,
    title: '应用内下载失败',
    message: '当前应用内下载失败，已切换到浏览器备用下载方案。',
    detail: `是否在浏览器中打开 ${updateState.assetName || '安装包'} 的下载地址？`
  });

  if (result.response === 0) {
    await shell.openExternal(targetUrl);
    return true;
  }

  return false;
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
    detail: '应用会自动选择更快的下载源，并在失败时切换备用地址。'
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
    detail: '将关闭当前应用并启动安装程序。'
  });

  if (result.response === 0) {
    quitAndInstallUpdate();
  }
}

async function fetchRemoteManifest() {
  const manifestUrls = getConfiguredManifestUrls();
  const failures = [];

  for (const manifestUrl of manifestUrls) {
    try {
      const response = await axios.get(manifestUrl, {
        timeout: DEFAULT_REQUEST_TIMEOUT,
        responseType: 'json',
        validateStatus: (status) => status >= 200 && status < 300
      });

      return ensureManifestShape(response.data, manifestUrl);
    } catch (error) {
      failures.push(toReadableUrlError(manifestUrl, error));
    }
  }

  throw new Error(`无法获取远程版本信息。已尝试地址: ${failures.join(' | ')}`);
}

async function safeUnlink(targetPath) {
  if (!targetPath) {
    return;
  }

  try {
    await fsPromises.unlink(targetPath);
  } catch (error) {
    if (error?.code !== 'ENOENT') {
      log.warn('[updater] failed to remove temp file:', targetPath, error.message);
    }
  }
}

function getDownloadDirectory(version) {
  return path.join(app.getPath('userData'), UPDATE_DOWNLOAD_DIRECTORY, stripVersionPrefix(version));
}

async function ensureDownloadDirectory(version) {
  const directory = getDownloadDirectory(version);
  await fsPromises.mkdir(directory, { recursive: true });
  return directory;
}

async function probeDownloadCandidate(candidate) {
  const startedAt = Date.now();

  try {
    await axios.head(candidate.url, {
      timeout: DEFAULT_PROBE_TIMEOUT,
      maxRedirects: 5,
      validateStatus: (status) => status >= 200 && status < 400
    });
  } catch (_) {
    const response = await axios.get(candidate.url, {
      timeout: DEFAULT_PROBE_TIMEOUT,
      maxRedirects: 5,
      responseType: 'stream',
      headers: {
        Range: 'bytes=0-0'
      },
      validateStatus: (status) => status === 200 || status === 206
    });

    response.data.destroy();
  }

  return {
    candidate,
    latency: Date.now() - startedAt
  };
}

async function rankDownloadCandidates(candidates) {
  const probeResults = await Promise.allSettled(candidates.map((candidate) => probeDownloadCandidate(candidate)));
  const successful = probeResults
    .filter((result) => result.status === 'fulfilled')
    .map((result) => result.value)
    .sort((left, right) => left.latency - right.latency)
    .map((result) => result.candidate);
  const successfulUrls = new Set(successful.map((candidate) => candidate.url));
  const remaining = candidates.filter((candidate) => !successfulUrls.has(candidate.url));

  return successful.length > 0 ? [...successful, ...remaining] : candidates;
}

async function streamDownload(candidate, manifest) {
  const downloadDirectory = await ensureDownloadDirectory(manifest.version);
  const assetName = manifest.assetName || getDefaultInstallerName(manifest.version) || 'TN-Audio-Toolkit-Setup.exe';
  const tempFilePath = path.join(downloadDirectory, `${assetName}.download`);
  const finalFilePath = path.join(downloadDirectory, assetName);

  await safeUnlink(tempFilePath);

  const response = await axios.get(candidate.url, {
    timeout: DEFAULT_REQUEST_TIMEOUT,
    maxRedirects: 5,
    responseType: 'stream',
    validateStatus: (status) => status >= 200 && status < 300
  });

  const total = Number.parseInt(response.headers['content-length'] || '0', 10) || 0;
  const startedAt = Date.now();
  let transferred = 0;

  await new Promise((resolve, reject) => {
    const writer = fs.createWriteStream(tempFilePath);

    response.data.on('data', (chunk) => {
      transferred += chunk.length;
      const elapsedSeconds = Math.max((Date.now() - startedAt) / 1000, 0.1);
      const bytesPerSecond = Math.round(transferred / elapsedSeconds);
      const progressPercent = total > 0 ? Number(((transferred / total) * 100).toFixed(1)) : 0;

      setState({
        status: 'downloading',
        checking: false,
        downloading: true,
        downloaded: false,
        installReady: false,
        progressPercent,
        transferred,
        total,
        bytesPerSecond,
        downloadSourceName: candidate.name,
        error: null
      });
    });

    response.data.on('error', reject);
    writer.on('error', reject);
    writer.on('finish', resolve);
    response.data.pipe(writer);
  }).catch(async (error) => {
    await safeUnlink(tempFilePath);
    throw error;
  });

  await safeUnlink(finalFilePath);
  await fsPromises.rename(tempFilePath, finalFilePath);
  return finalFilePath;
}

function setAvailableState(manifest) {
  setState({
    status: 'available',
    checking: false,
    available: true,
    downloading: false,
    downloaded: false,
    installReady: false,
    latestVersion: manifest.version,
    releaseName: manifest.releaseName,
    releaseDate: manifest.releaseDate,
    releaseNotes: manifest.releaseNotes,
    assetName: manifest.assetName,
    githubDownloadUrl: manifest.githubDownloadUrl,
    externalDownloadUrl: manifest.externalDownloadUrl,
    releasePageUrl: manifest.releasePageUrl,
    mirrorName: manifest.mirrorName,
    progressPercent: 0,
    transferred: 0,
    total: 0,
    bytesPerSecond: 0,
    downloadSourceName: null,
    error: null
  });
}

function setUpToDateState(manifest, source) {
  setState({
    status: 'up-to-date',
    checking: false,
    available: false,
    downloading: false,
    downloaded: false,
    installReady: false,
    latestVersion: manifest?.version || app.getVersion(),
    releaseName: manifest?.releaseName || null,
    releaseDate: manifest?.releaseDate || null,
    releaseNotes: manifest?.releaseNotes || null,
    assetName: manifest?.assetName || null,
    githubDownloadUrl: manifest?.githubDownloadUrl || null,
    externalDownloadUrl: manifest?.externalDownloadUrl || null,
    releasePageUrl: manifest?.releasePageUrl || null,
    mirrorName: manifest?.mirrorName || null,
    progressPercent: 0,
    transferred: 0,
    total: 0,
    bytesPerSecond: 0,
    downloadSourceName: null,
    lastCheckedAt: new Date().toISOString(),
    lastCheckSource: source,
    error: null
  });
}

function setCheckStartState(source) {
  setState({
    status: 'checking',
    checking: true,
    downloading: false,
    error: null,
    lastCheckedAt: new Date().toISOString(),
    lastCheckSource: source,
    unsupported: false
  });
}

function handleUpdaterError(error) {
  const message = getErrorMessage(error);
  const shouldOfferExternalDownload = Boolean(updateState.downloading && getExternalDownloadUrl(true));
  log.error('[updater] update failed:', message);
  setState({
    status: updateState.available ? 'available' : 'idle',
    checking: false,
    downloading: false,
    installReady: false,
    error: message
  });
  currentCheckSource = 'idle';

  if (shouldOfferExternalDownload) {
    promptForExternalDownload().catch((promptError) => {
      log.error('[updater] failed to show external download prompt:', promptError);
    });
  }

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

  if (!isUpdateCheckSupported()) {
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
  setCheckStartState(source);

  try {
    const manifest = await fetchRemoteManifest();
    remoteManifest = manifest;

    if (compareVersions(manifest.version, app.getVersion()) <= 0) {
      setUpToDateState(manifest, source);
      currentCheckSource = 'idle';
      return { ok: true, available: false, version: manifest.version };
    }

    setAvailableState(manifest);
    setState({
      lastCheckedAt: new Date().toISOString(),
      lastCheckSource: source
    });
    const triggerSource = currentCheckSource;
    currentCheckSource = 'idle';

    if (triggerSource === 'auto') {
      promptForDownload(manifest).catch((error) => {
        log.error('[updater] failed to show download prompt:', error);
      });
    }

    return { ok: true, available: true, version: manifest.version };
  } catch (error) {
    return handleUpdaterError(error);
  }
}

async function downloadUpdate(options = {}) {
  const source = options.source || 'manual';

  if (!isUpdateCheckSupported()) {
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

  if (updateState.downloading) {
    return { ok: false, busy: true, state: updateState };
  }

  const manifest = remoteManifest;
  if (!manifest || compareVersions(manifest.version, app.getVersion()) <= 0) {
    const checkResult = await checkForUpdates({ manual: source === 'manual' });
    if (!checkResult.ok || !remoteManifest || compareVersions(remoteManifest.version, app.getVersion()) <= 0) {
      return checkResult.ok
        ? { ok: false, message: '当前没有可下载的新版本。' }
        : checkResult;
    }
  }

  const activeManifest = remoteManifest;
  if (!activeManifest?.downloads?.length) {
    return handleUpdaterError(new Error('当前版本缺少可用下载地址。'));
  }

  if (updateState.downloaded && downloadedInstallerPath) {
    try {
      await fsPromises.access(downloadedInstallerPath);
      return { ok: true, alreadyDownloaded: true, filePath: downloadedInstallerPath };
    } catch (_) {
      downloadedInstallerPath = null;
    }
  }

  setState({
    status: 'downloading',
    checking: false,
    downloading: true,
    downloaded: false,
    installReady: false,
    progressPercent: 0,
    transferred: 0,
    total: 0,
    bytesPerSecond: 0,
    error: null,
    downloadSourceName: null
  });

  try {
    const orderedCandidates = await rankDownloadCandidates(activeManifest.downloads);
    const failures = [];

    for (const candidate of orderedCandidates) {
      try {
        downloadedInstallerPath = await streamDownload(candidate, activeManifest);
        setState({
          status: 'downloaded',
          checking: false,
          available: true,
          downloading: false,
          downloaded: true,
          installReady: true,
          latestVersion: activeManifest.version,
          releaseName: activeManifest.releaseName,
          releaseDate: activeManifest.releaseDate,
          releaseNotes: activeManifest.releaseNotes,
          assetName: activeManifest.assetName,
          githubDownloadUrl: activeManifest.githubDownloadUrl,
          externalDownloadUrl: activeManifest.externalDownloadUrl,
          releasePageUrl: activeManifest.releasePageUrl,
          mirrorName: activeManifest.mirrorName,
          progressPercent: 100,
          lastDownloadedAt: new Date().toISOString(),
          downloadSourceName: candidate.name,
          error: null
        });

        promptForInstall(activeManifest).catch((error) => {
          log.error('[updater] failed to show install prompt:', error);
        });

        return { ok: true, source, filePath: downloadedInstallerPath, downloadUrl: candidate.url };
      } catch (error) {
        failures.push(toReadableUrlError(candidate.url, error));
      }
    }

    return handleUpdaterError(new Error(`所有下载地址均失败。${failures.join(' | ')}`));
  } catch (error) {
    return handleUpdaterError(error);
  }
}

function quitAndInstallUpdate() {
  if (!updateState.downloaded || !downloadedInstallerPath) {
    return { ok: false, message: '当前没有已下载完成的更新。' };
  }

  setState({
    status: 'installing',
    checking: false,
    downloading: false,
    error: null
  });

  try {
    const child = spawn(downloadedInstallerPath, [], {
      detached: true,
      stdio: 'ignore'
    });

    child.unref();
    setImmediate(() => {
      app.quit();
    });

    return { ok: true, filePath: downloadedInstallerPath };
  } catch (error) {
    return handleUpdaterError(error);
  }
}

function getUpdateState() {
  return updateState;
}

function getExternalDownloadUrl(preferMirror = true) {
  const manifest = remoteManifest;
  if (!manifest) {
    return null;
  }

  if (preferMirror) {
    const mirrorCandidate = manifest.downloads.find((entry) => entry.channel === 'mirror');
    if (mirrorCandidate?.url) {
      return mirrorCandidate.url;
    }
  }

  return manifest.downloads[0]?.url || manifest.releasePageUrl || null;
}

module.exports = {
  initializeUpdateService,
  checkForUpdates,
  downloadUpdate,
  quitAndInstallUpdate,
  getUpdateState,
  getExternalDownloadUrl
};