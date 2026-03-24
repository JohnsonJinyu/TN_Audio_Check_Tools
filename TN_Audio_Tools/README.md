# TN Audio Toolkit

面向音频测试场景的数据收集、报告审查与频谱分析工具，基于 Electron + React + Node.js 构建。

## 当前进展

- 当前版本：1.1.3
- 已落地：测试数据收集、报告审查、设置持久化、托盘行为、安装版在线更新
- 本次新增：应用内更新镜像下载、下载失败自动回退到浏览器镜像下载
- 已验证：`npm run build` 可正常生成 Windows 安装包与便携版
- 持续建设中：频谱分析、批处理、音频转换等模块仍需继续完善业务能力

## 功能模块

### 1. 🎵 仪表盘 (Dashboard)
- 应用首页
- 快速统计信息
- 最近使用记录

### 2. 📋 测试数据收集 (Report Checker)
- 导入测试报告、checklist 和规则文件
- 执行测试数据收集并生成结论
- 展示处理结果与输出文件

### 3. 🔎 报告审查 (Report Review)
- 汇总文档完整性与曲线章节审查范围
- 查看最近处理记录
- 快速打开输出目录

### 4. 📊 频谱分析 (Spectrum Analyzer)
- 实时频谱分析
- 多种分析类型支持
- FFT、小波变换等算法
- 详细的分析参数配置
- 批量转换支持

### 6. ⚙️ 批量处理 (Batch Processor)
- 批量格式转换
- 批量报告检查
- 批量频谱分析
- 多线程处理

### 7. 🔧 设置 (Settings)
- 应用外观配置
- 文件和音频默认设置
- 系统偏好配置

## 技术栈

- **前端框架**: React 18.2.0
- **UI 组件库**: Ant Design 5.11.2
- **桌面应用**: Electron 27.0.0
- **构建工具**: Electron Builder
- **样式**: CSS 自定义主题
- **包管理**: npm

## 项目结构

```
.
├── public/                 # 静态资源
│   └── index.html         # HTML 入口
├── src/
│   ├── main/              # Electron 主进程
│   │   ├── main.js        # 主应用文件
│   │   └── preload.js     # 预加载脚本
│   ├── renderer/          # React 前端应用
│   │   ├── App.jsx        # 主应用组件
│   │   ├── App.css        # 应用样式
│   │   ├── index.jsx      # 应用入口
│   │   ├── index.css      # 全局样式
│   │   ├── pages/         # 页面组件
│   │   │   ├── Dashboard.jsx
│   │   │   ├── TestDataCollectionPage.jsx
│   │   │   ├── ReportReview.jsx
│   │   │   ├── SpectrumAnalyzer.jsx
│   │   │   └── Settings.jsx
│   │   └── styles/        # 样式文件
│   │       └── pages.css
│   └── utils/             # 工具函数
├── package.json           # 项目配置
└── .gitignore            # Git 忽略文件
```

## 安装和运行

### 环境要求
- Node.js 14.0.0 或更高版本
- npm 6.0.0 或更高版本

### 安装依赖

```bash
npm install
```

### 开发模式

```bash
npm run dev
```

此命令会同时启动 React 开发服务器和 Electron 应用。

### 构建生产版本

```bash
npm run build
```

构建完成后，可执行文件位于 `dist/` 目录。

## 使用说明

### 键盘快捷键
- `Ctrl+I`: 快速导入文件
- `Ctrl+S`: 保存设置
- `Ctrl+Q` / `Ctrl+W`: 退出应用

### 文件支持

**音频格式**
- MP3, WAV, FLAC, AAC, OGG, M4A 等

**报告格式**
- Excel (.xlsx, .xls)
- CSV
- PDF

## 开发指南

### 添加新的页面

1. 在 `src/renderer/pages/` 中创建新组件
2. 在 `App.jsx` 中导入并注册路由
3. 在菜单中添加对应的菜单项

### 样式定制

所有样式均使用 CSS 变量定义，可在 `src/renderer/App.css` 中修改：

```css
:root {
  --primary-color: #1890ff;
  --secondary-color: #722ed1;
  /* ... */
}
```

### IPC 通信

在 Electron 主进程中定义 IPC 处理程序：

```javascript
ipcMain.handle('channel-name', async (event, data) => {
  // 处理业务逻辑
  return result;
});
```

在渲染进程中调用：

```javascript
window.electron.ipcRenderer.invoke('channel-name', data);
```

## 常见问题

### 应用无法启动
- 检查 Node.js 版本是否符合要求
- 尝试删除 `node_modules` 目录重新安装依赖

### 模块未找到错误
- 确保 `npm install` 已成功执行
- 清除 npm 缓存: `npm cache clean --force`

## 许可证

MIT

## 贡献

欢迎提交 Issue 和 Pull Request！

## 联系方式

开发者: JohnsonJinyu
