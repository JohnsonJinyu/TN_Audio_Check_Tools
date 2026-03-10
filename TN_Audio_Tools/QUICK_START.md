# TN Audio Toolkit - 快速开始指南

## 📦 项目已完成部分

### ✅ UI 框架 100% 完成
这个项目包含了完整的现代化 UI 设计，所有 7 个页面都已经设计和实现：

1. **Dashboard（仪表盘）** - 应用首页
2. **ReportChecker（报告检查）** - 核心功能模块
3. **AudioPlayer（音频播放器）** - 音频播放功能
4. **SpectrumAnalyzer（频谱分析）** - 频谱分析工具
5. **AudioConverter（音频转换）** - 格式转换工具
6. **BatchProcessor（批量处理）** - 批量处理功能
7. **Settings（设置）** - 应用设置页面

### ✨ 设计特色
- 紫色渐变侧边栏 (#667eea → #764ba2)
- 蓝色主要颜色方案 (#1890ff)
- Ant Design UI 组件库
- 响应式布局（支持移动设备）
- 平滑的动画和过渡效果
- 现代化的卡片设计

## 🚀 3分钟快速启动

### 第 1 步：安装依赖
```bash
npm install
```

### 第 2 步：启动开发模式
```bash
npm run dev
```

此命令会自动：
- 启动 React 开发服务器 (http://localhost:3000)
- 启动 Electron 应用窗口

### 第 3 步：等待应用加载
- React 开发服务器启动完成后（约 10-20 秒）
- Electron 窗口会自动打开
- 开发者工具会自动打开（左侧 Console 面板）

## 📝 推荐开发流程

### 开发新功能时：
1. 修改页面组件文件（`src/renderer/pages/*.jsx`）
2. 在浏览器的 React DevTools 中查看实时效果
3. 如果需要后端逻辑，在 `src/main/main.js` 中添加 IPC 处理程序

### 修改样式时：
1. 全局样式：编辑 `src/renderer/App.css`
2. 页面样式：编辑 `src/renderer/styles/pages.css`
3. 组件样式：直接在组件的 `style` 属性中修改

### 测试应用时：
- 使用 `npm run dev` 在开发模式下测试
- 使用开发者工具调试（F12 打开）
- 修改后会自动热重载

## 🔧 常用命令

```bash
# 启动开发模式（推荐）
npm run dev

# 仅启动 React 开发服务器
npm run react-start

# 仅启动 Electron 应用
npm run electron-dev

# 构建生产版本
npm run build

# 编译 React（用于生产构建）
npm run react-build
```

## 📂 重要文件位置

| 文件 | 用途 | 编辑建议 |
|------|------|---------|
| `src/renderer/App.jsx` | 主应用组件和导航 | ⭐ 菜单配置 |
| `src/renderer/App.css` | 全局样式和主题 | ⭐ 颜色修改 |
| `src/renderer/pages/` | 7个页面组件 | ⭐ 添加功能 |
| `src/renderer/styles/pages.css` | 页面通用样式 | UI 样式 |
| `src/main/main.js` | Electron 主进程 | ⭐ IPC 处理程序 |
| `src/main/preload.js` | IPC 桥接脚本 | 通信配置 |
| `package.json` | 项目配置和依赖 | 依赖管理 |

## 💡 实现功能的步骤

以"报告检查"功能为例：

### 1. 设计 UI（已完成）
文件：`src/renderer/pages/ReportChecker.jsx`
- 已有上传区域、表格、操作按钮

### 2. 添加状态管理
```javascript
const [files, setFiles] = useState([]);
const [checking, setChecking] = useState(false);
```

### 3. 在主进程添加 IPC 处理程序
```javascript
// src/main/main.js
ipcMain.handle('check-report', async (event, filePath) => {
  // 实现检查逻辑
  return results;
});
```

### 4. 在组件中调用 IPC
```javascript
const handleCheck = async (file) => {
  const result = await window.electron.ipcRenderer.invoke('check-report', file.path);
  // 更新 UI
};
```

## 🎨 自定义主题

编辑 `src/renderer/App.css` 中的 CSS 变量：

```css
:root {
  --primary-color: #1890ff;      /* 修改主颜色 */
  --secondary-color: #722ed1;    /* 修改副颜色 */
  --success-color: #52c41a;      /* 修改成功颜色 */
  /* ... */
}
```

## 📊 项目统计

- **文件总数**: 23 个
- **代码行数**: ~2000+ 行
- **页面组件**: 7 个
- **UI 组件**: Ant Design (60+ 个)
- **包依赖**: 15 个

## 🐛 常见问题

### Q: 应用启动时黑屏？
A: 这是正常的，需要等待 React 开发服务器启动（通常 10-20 秒）

### Q: 如何修改窗口大小？
A: 编辑 `src/main/main.js` 中的 BrowserWindow 配置
```javascript
mainWindow = new BrowserWindow({
  width: 1400,  // 修改宽度
  height: 900,  // 修改高度
  // ...
});
```

### Q: 如何添加新的菜单项？
A: 编辑 `src/renderer/App.jsx` 中的 `menuItems` 数组

### Q: 如何禁用开发者工具？
A: 在 `src/main/main.js` 中注释掉：
```javascript
// if (isDev) {
//   mainWindow.webContents.openDevTools();
// }
```

## 📚 学习资源

- [React 官方文档](https://react.dev)
- [Electron 官方文档](https://www.electronjs.org)
- [Ant Design 文档](https://ant.design)
- [项目开发指南](.github/copilot-instructions.md)

## ✅ 检查清单

启动前检查：
- [ ] Node.js 已安装（v14+）
- [ ] npm 已安装（v6+）
- [ ] 依赖已安装（`npm install` 完成）

启动后检查：
- [ ] React 开发服务器已启动
- [ ] Electron 窗口已打开
- [ ] 可以看到仪表盘页面
- [ ] 左侧菜单可以切换页面

## 🎯 下一步

1. **执行 `npm install`** - 安装所有依赖
2. **执行 `npm run dev`** - 启动开发模式
3. **浏览应用** - 点击左侧菜单切换不同页面
4. **开始开发** - 修改页面组件添加功能

---

**享受开发！** 🎉

有任何问题，请参考 [开发指南](.github/copilot-instructions.md) 或 [项目说明](README.md)。
