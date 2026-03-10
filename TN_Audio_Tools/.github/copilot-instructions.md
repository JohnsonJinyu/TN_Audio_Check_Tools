# TN Audio Toolkit - 项目开发指南

本文件提供项目特定的开发指南和说明。

## 项目概述

TN Audio Toolkit 是一个基于 Electron + React + Node.js 的现代化音频处理和分析应用。应用采用现代化的 UI 设计，提供多个音频处理模块。

## 项目结构说明

### 主要目录结构

```
src/
├── main/               # Electron 主进程代码
│   ├── main.js        # 应用主文件，窗口创建和管理
│   └── preload.js     # 预加载脚本，IPC 通信桥接
│
└── renderer/          # React 前端应用
    ├── App.jsx        # 主应用组件（导航、布局）
    ├── App.css        # 全局应用样式
    ├── index.jsx      # 应用入口文件
    ├── index.css      # 全局基础样式
    │
    ├── pages/         # 页面组件
    │   ├── Dashboard.jsx          # 仪表盘
    │   ├── ReportChecker.jsx      # 报告检查（核心功能）
    │   ├── AudioPlayer.jsx        # 音频播放器
    │   ├── SpectrumAnalyzer.jsx   # 频谱分析
    │   ├── AudioConverter.jsx     # 格式转换
    │   ├── BatchProcessor.jsx     # 批量处理
    │   └── Settings.jsx           # 设置页面
    │
    └── styles/        # 样式文件
        └── pages.css  # 页面通用样式
```

## 核心功能模块

### 1. Dashboard（仪表盘）
- 应用首页和欢迎页面
- 快速统计信息卡片
- 工具导航
- **实现状态**: UI 框架完成，逻辑待实现

### 2. ReportChecker（报告检查）- 核心功能
- 导入音频测试报告
- 自动验证报告完整性和有效性
- 显示检查结果
- **实现状态**: UI 框架完成，核心逻辑待实现

### 3. 其他模块
- AudioPlayer：音频播放
- SpectrumAnalyzer：频谱分析
- AudioConverter：格式转换
- BatchProcessor：批量处理
- Settings：应用设置
- **实现状态**: UI 框架完成，业务逻辑待实现

## 技术栈详解

### 前端 (React)
- React 18.2.0：用户界面框架
- Ant Design 5.11.2：企业级 UI 组件库
- CSS 自定义属性：主题管理

### 桌面应用 (Electron)
- Electron 27.0.0：跨平台桌面应用框架
- Electron Builder：打包和发行工具
- electron-store：数据持久化
- electron-log：日志管理

## 样式设计系统

### 颜色系统
```css
--primary-color: #1890ff      /* 主色 - 蓝色 */
--secondary-color: #722ed1    /* 次色 - 紫色 */
--success-color: #52c41a      /* 成功 - 绿色 */
--warning-color: #faad14      /* 警告 - 橙色 */
--error-color: #f5222d        /* 错误 - 红色 */
```

### 设计特点
- 渐变色背景的侧边栏和标题
- 圆角卡片设计（border-radius: 6-8px）
- 平滑的过渡动画（transition: all 0.3s ease）
- 响应式布局（支持移动设备）
- 深色侧边栏 + 浅色主内容区

## IPC 通信指南

### 在主进程中添加处理程序

```javascript
// src/main/main.js
ipcMain.handle('channel-name', async (event, data) => {
  // 处理业务逻辑
  return result;
});
```

### 在渲染进程中调用

```javascript
// src/renderer/pages/*.jsx
const result = await window.electron.ipcRenderer.invoke('channel-name', data);
```

## 开发流程

### 1. 启动开发环境
```bash
npm install        # 安装依赖
npm run dev        # 启动开发模式
```

### 2. 实现功能步骤
1. 在对应的页面组件中完成 UI（已完成）
2. 在 IPC 处理程序中实现后端逻辑
3. 在前端组件中调用 IPC 方法
4. 测试和调试

### 3. 打包发行
```bash
npm run build      # 构建生产版本
```

## 数据流

```
User UI (React Component)
    ↓
IPC Renderer (window.electron.ipcRenderer)
    ↓
IPC Main (ipcMain.handle)
    ↓
Node.js Logic / External Process
    ↓
Return Result to Renderer
    ↓
Update UI State
```

## 常见开发任务

### 添加新的菜单项
1. 在 `App.jsx` 的 `menuItems` 数组中添加新项
2. 创建对应的页面组件
3. 在 `renderContent()` 中添加渲染逻辑

### 修改主题颜色
编辑 `src/renderer/App.css` 中的 CSS 变量

### 添加新的工具函数
在 `src/utils/` 目录下创建相应的工具文件

## 调试

### 开发者工具
在开发模式下，Electron 会自动打开 DevTools，可以：
- 检查 React 组件
- 查看网络请求
- 调试 JavaScript 代码
- 查看控制台日志

### 日志记录
```javascript
import log from 'electron-log';

log.info('Message to log');
log.error('Error message');
```

## 构建和打包

### Windows 构建
```bash
npm run build      # 生成 .exe 和 .exe 便携版
```

### 输出目录
- `build/`：React 编译输出
- `dist/`：最终可执行文件

## 性能优化建议

1. **代码分割**: 使用 React.lazy() 进行页面级代码分割
2. **图片优化**: 使用适当的图片格式和大小
3. **缓存策略**: 使用 electron-store 缓存用户数据
4. **异步处理**: IPC 通信使用 async/await

## 安全注意事项

1. **上下文隔离**: 已在 preload.js 中启用
2. **nodeIntegration**: 已禁用，使用 IPC 通信
3. **远程内容**: 避免加载不可信的远程内容
4. **用户数据**: 使用 electron-store 安全存储

## 扩展指南

### 添加新的音频处理算法
1. 在 Node.js 后端实现算法
2. 通过 IPC 暴露接口
3. 在对应页面调用

### 集成第三方库
- 音频处理: tone.js, web-audio-api
- 频谱分析: essentia.js
- 格式转换: ffmpeg

## 后续开发优先级

1. ✅ UI 框架完成
2. ⏳ 实现报告检查核心逻辑
3. ⏳ 实现音频播放功能
4. ⏳ 实现频谱分析
5. ⏳ 实现格式转换
6. ⏳ 实现批量处理

## 常见问题

**Q: 如何在开发时修改窗口大小?**
A: 编辑 `src/main/main.js` 中的 BrowserWindow 配置

**Q: 如何添加应用图标?**
A: 将图标文件放在 `assets/` 目录，更新 package.json build 配置

**Q: 如何实现文件拖放?**
A: 使用 Electron 的 drag-drop 事件，在预加载脚本中设置

## 相关资源

- [React 官方文档](https://react.dev)
- [Electron 官方文档](https://www.electronjs.org/docs)
- [Ant Design 文档](https://ant.design)
- [electron-builder 文档](https://www.electron.build)

---

**最后更新**: 2026-03-10
**开发者**: JohnsonJinyu
