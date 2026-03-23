require('./services/testDataExtraction/runtimePolyfills');

const fs = require('fs/promises');
const { app, BrowserWindow, Menu, ipcMain, shell, dialog } = require('electron');
const isDev = require('electron-is-dev');
const path = require('path');
const {
  processReports,
  DEFAULT_RULES_RELATIVE_PATH,
  buildExportableRulesContent,
  parseChecklistReportOptions,
  inspectReport
} = require('./services/testDataExtraction');
const { reviewWordReport } = require('./services/reportReview');

// Avoid GPU process crashes on some Windows drivers/VM environments.
app.disableHardwareAcceleration();

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1000,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    icon: path.join(__dirname, '../../assets/icon.ico')
  });

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

  if (isDev) {
    mainWindow.webContents.openDevTools();
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
  }
});

// IPC 处理程序示例
ipcMain.handle('get-version', () => {
  return app.getVersion();
});

ipcMain.handle('report-checker:process-reports', async (_, payload) => {
  const runId = payload?.runId || `${Date.now()}-${Math.random().toString(16).slice(2)}`;

  return processReports({
    ...payload,
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
  const sourcePath = customRulePath || path.join(app.getAppPath(), DEFAULT_RULES_RELATIVE_PATH);
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
        { label: '关于', click: () => console.log('About') }
      ]
    }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
};

app.on('ready', createMenu);
