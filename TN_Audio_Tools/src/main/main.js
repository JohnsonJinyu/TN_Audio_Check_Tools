if (process.env.NODE_ENV === 'development') {
  process.env.ELECTRON_DISABLE_SECURITY_WARNINGS = 'true';
}

require('./services/testDataExtraction/runtimePolyfills');

const progressBus = require('./services/reportReview/progressBus');

const fs = require('fs/promises');
const os = require('os');
const { app, BrowserWindow, Menu, Tray, ipcMain, shell, dialog, nativeImage } = require('electron');
const isDev = require('electron-is-dev');
const path = require('path');
const { existsSync } = require('fs');
const {
  initializeUpdateService,
  checkForUpdates,
  downloadUpdate,
  quitAndInstallUpdate,
  getUpdateState,
  getExternalDownloadUrl
} = require('./services/updater/updateService');
const {
  processReports,
  resolveBundledRulesPath,
  buildExportableRulesContent,
  parseChecklistReportOptions,
  inspectReport
} = require('./services/testDataExtraction');
const { reviewWordReport, reviewPairedReport, runCrossReportChecks } = require('./services/reportReview');
const {
  DEFAULT_APP_SETTINGS,
  getSettings,
  normalizeSettings,
  resetSettings,
  saveSettings
} = require('./services/settingsService');

// Avoid GPU process crashes on some Windows drivers/VM environments.
app.disableHardwareAcceleration();
app.commandLine.appendSwitch('disable-http-cache');
app.commandLine.appendSwitch('disable-gpu-shader-disk-cache');
if (process.platform === 'win32') {
  // Dev mode should use executable path, packaged app should use stable app id.
  app.setAppUserModelId(app.isPackaged ? 'com.tnaudio.toolkit' : process.execPath);
}

let mainWindow;
let tray = null;
let isQuitting = false;

function getIconPath() {
  const candidatePaths = isDev
    ? [path.join(__dirname, '../../assets/icon.ico')]
    : [
        path.join(process.resourcesPath, 'icon.ico'),
        path.join(__dirname, '../../assets/icon.ico')
      ];

  const resolved = candidatePaths.find((candidatePath) => existsSync(candidatePath));
  return resolved || candidatePaths[0];
}

function emitSettingsChanged(settings) {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return;
  }

  mainWindow.webContents.send('app-settings:changed', settings);
}

function showMainWindow() {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return;
  }

  if (mainWindow.isMinimized()) {
    mainWindow.restore();
  }

  mainWindow.show();
  mainWindow.focus();
}

function hideMainWindowToTray() {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return;
  }

  mainWindow.hide();
}

function destroyTray() {
  if (!tray) {
    return;
  }

  tray.destroy();
  tray = null;
}

function buildTrayMenu() {
  return Menu.buildFromTemplate([
    { label: '显示主窗口', click: () => showMainWindow() },
    { type: 'separator' },
    {
      label: '退出',
      click: () => {
        isQuitting = true;
        app.quit();
      }
    }
  ]);
}

function ensureTray() {
  const appSettings = getSettings();
  if (!appSettings.system.enableTray) {
    destroyTray();
    return;
  }

  if (!tray) {
    const trayIcon = nativeImage.createFromPath(getIconPath());
    tray = new Tray(trayIcon);
    tray.setToolTip('TN Audio Toolkit');
    tray.on('click', () => {
      if (mainWindow && !mainWindow.isDestroyed() && mainWindow.isVisible()) {
        hideMainWindowToTray();
      } else {
        showMainWindow();
      }
    });
    tray.on('double-click', () => showMainWindow());
  }

  tray.setContextMenu(buildTrayMenu());
}

function applyDesktopBehavior() {
  const appSettings = getSettings();
  ensureTray();

  if (!appSettings.system.enableTray && mainWindow && !mainWindow.isDestroyed() && !mainWindow.isVisible()) {
    showMainWindow();
  }
}

async function clearRuntimeCache() {
  const windowRef = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null;
  if (windowRef) {
    await windowRef.webContents.session.clearCache();
  }

  const tempRoot = os.tmpdir();
  const entries = await fs.readdir(tempRoot, { withFileTypes: true }).catch(() => []);
  const removableDirectories = entries
    .filter((entry) => entry.isDirectory() && entry.name.startsWith('tn-audio-report-'))
    .map((entry) => fs.rm(path.join(tempRoot, entry.name), { recursive: true, force: true }));

  await Promise.all(removableDirectories);

  return {
    clearedBrowserCache: Boolean(windowRef),
    removedTempDirectories: removableDirectories.length
  };
}

function createWindow() {
  const appSettings = getSettings();
  const shouldStartHidden = appSettings.system.enableTray && appSettings.system.launchMinimizedToTray;

  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1000,
    minHeight: 600,
    show: !shouldStartHidden,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    icon: getIconPath()
  });

  if (process.platform === 'win32') {
    const winIcon = nativeImage.createFromPath(getIconPath());
    if (!winIcon.isEmpty()) {
      mainWindow.setIcon(winIcon);
    }
  }

  const startUrl = isDev
    ? (process.env.ELECTRON_RENDERER_URL || 'http://localhost:3123')
    : `file://${path.join(__dirname, '../../build/index.html')}`;

  mainWindow.webContents.on('did-fail-load', (_, errorCode, errorDescription, validatedURL) => {
    if (!isDev) {
      return;
    }

    const html = `
      <!DOCTYPE html>
      <html lang="zh-CN">
        <head>
          <meta charset="UTF-8" />
          <title>开发服务未启动</title>
          <style>
            body {
              margin: 0;
              font-family: "Segoe UI", sans-serif;
              background: linear-gradient(160deg, #f7f8fc 0%, #eef1f8 100%);
              color: #1f2937;
              display: flex;
              align-items: center;
              justify-content: center;
              min-height: 100vh;
            }
            .panel {
              width: min(720px, calc(100vw - 48px));
              background: #ffffff;
              border: 1px solid #dbe3f0;
              border-radius: 16px;
              box-shadow: 0 18px 48px rgba(31, 41, 55, 0.12);
              padding: 28px 32px;
            }
            h1 {
              margin: 0 0 12px;
              font-size: 24px;
            }
            p {
              margin: 0 0 12px;
              line-height: 1.6;
            }
            code {
              background: #f3f4f6;
              border-radius: 6px;
              padding: 2px 6px;
              font-family: Consolas, monospace;
            }
            ul {
              margin: 16px 0;
              padding-left: 20px;
              line-height: 1.7;
            }
            .hint {
              margin-top: 18px;
              padding: 12px 14px;
              background: #f9fafb;
              border-left: 4px solid #2563eb;
              border-radius: 8px;
            }
          </style>
        </head>
        <body>
          <div class="panel">
            <h1>React 开发服务没有启动</h1>
            <p>Electron 当前尝试加载 <code>${validatedURL || startUrl}</code>，但没有连上，所以窗口显示为空白。</p>
            <ul>
              <li>完整启动开发环境：<code>npm start</code></li>
              <li>如果 React 已经在跑，只需要启动桌面端：<code>npm run electron-dev</code></li>
              <li>如果 3000 端口被旧进程占用，先清掉旧进程再重启</li>
            </ul>
            <p>错误信息：<code>${errorCode} / ${errorDescription}</code></p>
            <div class="hint">现在这个白屏不是页面组件渲染报错，而是开发模式下没有成功连接到本地前端服务。</div>
          </div>
        </body>
      </html>
    `;

    mainWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
  });

  mainWindow.loadURL(startUrl);

  // 全局图表分析进度转发：任何模块 emit 的进度都推送到渲染进程
  progressBus.on(progressBus.events.CHART_PROGRESS_EVENT, function(data) {
    try {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('chart-analysis-progress', data);
      }
    } catch (_) {}
  });

  progressBus.on(progressBus.events.REVIEW_PROGRESS_EVENT, function(data) {
    try {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('report-review-progress', data);
      }
    } catch (_) {}
  });

  if (isDev) {
    // mainWindow.webContents.openDevTools(); // 启动时不自动打开，可通过菜单手动打开
  }

  mainWindow.on('minimize', (event) => {
    if (!getSettings().system.enableTray) {
      return;
    }

    event.preventDefault();
    hideMainWindowToTray();
  });

  mainWindow.on('close', (event) => {
    if (isQuitting || !getSettings().system.enableTray) {
      return;
    }

    event.preventDefault();
    hideMainWindowToTray();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

function initializeApp() {
  createWindow();
  createMenu();
  applyDesktopBehavior();
  initializeUpdateService({
    getMainWindow: () => mainWindow
  });

  setTimeout(() => {
    checkForUpdates({ manual: false }).catch(() => {});
  }, 5000);
}

app.on('ready', initializeApp);

app.on('before-quit', () => {
  isQuitting = true;
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
    applyDesktopBehavior();
    return;
  }

  showMainWindow();
});

// IPC 处理程序示例
ipcMain.handle('get-version', () => {
  return app.getVersion();
});

ipcMain.handle('app-settings:get', () => {
  return getSettings();
});

ipcMain.handle('app-settings:defaults', () => {
  return DEFAULT_APP_SETTINGS;
});

ipcMain.handle('app-settings:save', async (_, payload) => {
  const savedSettings = saveSettings(normalizeSettings(payload || {}));
  applyDesktopBehavior();
  emitSettingsChanged(savedSettings);
  return savedSettings;
});

ipcMain.handle('app-settings:reset', async () => {
  const resetValue = resetSettings();
  applyDesktopBehavior();
  emitSettingsChanged(resetValue);
  return resetValue;
});

ipcMain.handle('app-settings:choose-output-directory', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: '选择默认输出目录',
    properties: ['openDirectory', 'createDirectory']
  });

  return {
    canceled,
    filePath: canceled ? '' : (filePaths[0] || '')
  };
});

ipcMain.handle('app-settings:clear-cache', async () => {
  return clearRuntimeCache();
});

ipcMain.handle('app-update:get-state', () => {
  return getUpdateState();
});

ipcMain.handle('app-update:check-for-updates', async () => {
  return checkForUpdates({ manual: true });
});

ipcMain.handle('app-update:download-update', async () => {
  return downloadUpdate({ source: 'manual' });
});

ipcMain.handle('app-update:open-external-download', async (_, payload) => {
  const preferMirror = payload?.preferMirror !== false;
  const targetUrl = getExternalDownloadUrl(preferMirror);

  if (!targetUrl) {
    return { ok: false, message: '当前没有可用的外部下载地址。' };
  }

  await shell.openExternal(targetUrl);
  return { ok: true, url: targetUrl };
});

ipcMain.handle('app-update:quit-and-install', () => {
  return quitAndInstallUpdate();
});

ipcMain.handle('report-checker:process-reports', async (_, payload) => {
  const runId = payload?.runId || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const appSettings = getSettings();

  return processReports({
    ...payload,
    maxConcurrentTasks: payload?.maxConcurrentTasks || appSettings.files.maxConcurrentTasks,
    outputDirectory: payload?.outputDirectory || appSettings.files.defaultOutputDirectory || '',
    appPath: app.getAppPath(),
    onProgress: (progressPayload) => {
      if (_.sender.isDestroyed()) {
        return;
      }

      _.sender.send('report-checker:progress', {
        runId,
        ...progressPayload
      });
    }
  });
});

ipcMain.handle('report-checker:show-output-in-folder', async (_, filePath) => {
  if (!filePath) {
    throw new Error('缺少输出文件路径');
  }

  await shell.showItemInFolder(filePath);
  return true;
});

ipcMain.handle('report-checker:get-checklist-report-options', async (_, checklistPath) => {
  return parseChecklistReportOptions(checklistPath);
});

ipcMain.handle('report-checker:inspect-report-context', async (_, payload) => {
  return inspectReport(payload?.reportPath, {
    customer: payload?.customer,
    reportPanelSelections: payload?.reportPanelSelections
  });
});

ipcMain.handle('report-checker:export-rules', async (_, customRulePath) => {
  const sourcePath = customRulePath || await resolveBundledRulesPath(app.getAppPath());
  await fs.access(sourcePath);

  const defaultName = path.basename(sourcePath);
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title: '导出规则文件',
    defaultPath: defaultName,
    filters: [
      { name: '规则文件', extensions: ['json5', 'json'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });

  if (canceled || !filePath) {
    return { canceled: true };
  }

  const outputPath = path.extname(filePath)
    ? filePath
    : `${filePath}${path.extname(defaultName) || '.json5'}`;

  const exportableContent = await buildExportableRulesContent(sourcePath);
  await fs.writeFile(outputPath, exportableContent, 'utf8');
  return {
    canceled: false,
    filePath: outputPath
  };
});

ipcMain.handle('report-review:review-word-report', async (_, payload) => {
  if (!payload?.reportPath) {
    throw new Error('缺少报告路径');
  }

  const result = await reviewWordReport(payload.reportPath);
  return result;
});

ipcMain.handle('report-review:review-paired-report', async (_, payload) => {
  if (!payload?.docxPath || !payload?.xlsxPath) {
    throw new Error('配对审查需要同时提供 .docx 和 .xlsx 文件路径');
  }

  const result = await reviewPairedReport(payload.docxPath, payload.xlsxPath);
  return result;
});

ipcMain.handle('report-review:run-cross-report-checks', async (_, payload) => {
  if (!Array.isArray(payload?.results) || payload.results.length < 2) {
    throw new Error('跨报告对比需要至少2份审查结果');
  }

  return runCrossReportChecks(payload.results);
});

ipcMain.handle('report-review:analyze-chart-images', async (event, payload) => {
  const { reportPath, testDataFacts } = payload || {};
  if (!reportPath) {
    throw new Error('缺少报告路径');
  }

  const settings = getSettings();
  if (!settings.llm?.enabled || !settings.llm?.apiKey || !settings.llm?.apiUrl) {
    return {
      status: 'review',
      issues: [{ severity: 'review', message: '请在设置中启用AI图表分析并配置API地址和Key' }],
      evidence: ['LLM图表分析未配置或未启用']
    };
  }

  try {
    const { extractReportImages } = require('./services/reportReview/imageExtractor');
    const { analyzeGroupedCharts } = require('./services/reportReview/llmService');

    const { images, warnings } = await extractReportImages(reportPath);
    var chartImages = (images || []);
    var _ev = (warnings || []).slice();

    if (chartImages.length === 0) {
      return {
        status: 'review',
        issues: [{ severity: 'review', message: '未在报告中检测到频响/响度曲线图，无法进行AI验证' }],
        evidence: _ev.length > 0 ? _ev : ['报告中无可提取的图片，请确认报告包含图表']
      };
    }

    return await analyzeGroupedCharts({
      images: chartImages,
      testDataFacts: testDataFacts || {},
      settings: settings.llm
    }, function(progress) {
      try { progressBus.emitChartProgress({ imageCurrent: progress.current, imageTotal: progress.total, fileName: progress.fileName || '', imageCount: progress.imageCount || 0, status: progress.status || 'analyzing', detail: progress.detail || '' }); } catch (_) {}
    });
  } catch (error) {
    return {
      status: 'review',
      issues: [{ severity: 'review', message: 'AI图表分析异常: ' + (error.message || '未知错误') }],
      evidence: ['分析过程出错: ' + (error.message || '')]
    };
  }
});

ipcMain.handle('llm:test-connection', async (_, payload) => {
  var p = payload || {};
  var settings = getSettings();
  var apiUrl = String(p.apiUrl || settings.llm?.apiUrl || '').replace(/\/+$/, '');
  var apiKey = String(p.apiKey || settings.llm?.apiKey || '');
  var model = String(p.model || settings.llm?.model || 'claude-sonnet-4-20250514');

  if (!apiUrl || !apiKey) {
    return { ok: false, message: 'API地址或Key未填写' };
  }

  try {
    var axios = require('axios');
    await axios.post(apiUrl + '/v1/chat/completions', {
      model: model,
      max_tokens: 10,
      messages: [{ role: 'user', content: '回复OK' }]
    }, {
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
      timeout: 15000
    });
    return { ok: true, message: '连接成功，API可用 (' + model + ')' };
  } catch (e) {
    var msg = '连接失败: ';
    if (e.response && e.response.status === 401) msg += 'API Key无效(401)';
    else if (e.response && e.response.status === 403) msg += '无权限访问(403)';
    else if (e.code === 'ECONNREFUSED' || e.code === 'ENOTFOUND') msg += '无法连接到 ' + apiUrl;
    else if (e.code === 'ETIMEDOUT' || e.code === 'ECONNABORTED') msg += '连接超时';
    else msg += (e.response?.status || e.message);
    return { ok: false, message: msg };
  }
});

ipcMain.handle('dialog:open-file', async (_, options = {}) => {
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: options.title || '选择文件',
    filters: options.filters || [{ name: '所有文件', extensions: ['*'] }],
    properties: options.properties || ['openFile']
  });

  return {
    canceled,
    filePath: filePaths
  };
});

ipcMain.handle('report-review:upload-word-report', async (_, payload) => {
  if (!payload?.filePath) {
    throw new Error('缺少上传的报告文件路径');
  }

  // 验证文件是否存在
  try {
    await fs.access(payload.filePath);
  } catch (err) {
    throw new Error(`上传的报告文件不存在: ${payload.filePath}`);
  }

  // 验证文件扩展名
  const ext = path.extname(payload.filePath).toLowerCase();
  if (!['.doc', '.docx'].includes(ext)) {
    throw new Error(`不支持的文件格式: ${ext}，仅支持 .doc 和 .docx`);
  }

  // 直接调用审查函数
  const result = await reviewWordReport(payload.filePath);
  return result;
});

// 创建菜单
const createMenu = () => {
  const template = [
    {
      label: '文件',
      submenu: [
        { label: '退出', accelerator: 'CmdOrCtrl+Q', click: () => app.quit() }
      ]
    },
    {
      label: '编辑',
      submenu: [
        { label: '撤销', accelerator: 'CmdOrCtrl+Z', role: 'undo' },
        { label: '重做', accelerator: 'CmdOrCtrl+Y', role: 'redo' },
        { type: 'separator' },
        { label: '剪切', accelerator: 'CmdOrCtrl+X', role: 'cut' },
        { label: '复制', accelerator: 'CmdOrCtrl+C', role: 'copy' },
        { label: '粘贴', accelerator: 'CmdOrCtrl+V', role: 'paste' }
      ]
    },
    {
      label: '帮助',
      submenu: [
        { label: '检查更新', click: () => checkForUpdates({ manual: true }) },
        { label: '关于', click: () => console.log('About') }
      ]
    }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
};
