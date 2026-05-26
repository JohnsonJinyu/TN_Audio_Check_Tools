# Scripts 导航

这个目录按职责拆成四层，避免开发脚本、发布脚本和诊断脚本继续混放在一起。

## 目录说明

- dev/
  - dev-runner.js: 本地开发入口，负责启动 React 开发服务并拉起 Electron
- release/
  - release-preflight.js: 发布前自检，检查 gh 登录、版本号和构建目标
- diagnostics/
  - analyze-3gpp-corpus.js: 3GPP 语料批量差异分析
  - run-3gpp-regression.js: 3GPP 回归摘要输出
  - debug-report-checker.js: 提取结果定向比对/诊断
  - diag-borders.js: checklist 边框问题诊断
  - inspect-template.js: 模板样式与 merge 基线检查
  - test_review.js: Word 报告审查快速验证
  - test_full_review.js: Word/xlsx 报告审查综合验证
- tools/
  - cleanup-office-processes.ps1: 清理 Office 相关残留进程
  - generate-app-icon.ps1: 生成应用图标相关资源

## 常用命令

```powershell
npm run dev
npm run release:preflight
npm run diag:review
npm run diag:review:full
npm run diag:3gpp
npm run diag:3gpp:summary
```

## 维护约定

- 直接参与 npm 主流程的入口脚本放到 dev/ 或 release/。
- 一次性排查、回归验证、模板诊断脚本放到 diagnostics/。
- PowerShell 辅助工具放到 tools/。
- 如果某个诊断脚本已经长期稳定并成为常用工具，再考虑提升到 package.json 的 npm scripts。
