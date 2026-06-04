# TN Audio Toolkit

专业音频测试报告 AI 检查填写工具 —— 面向 ACQUA 6.0 测试系统的桌面端自动化解决方案。

## 简介

TN Audio Toolkit 是一款 Electron 桌面应用，专为音频测试工程师设计，用于自动化处理 ACQUA 6.0 生成的音频测试报告。核心能力包括：

- **数据提取**：从结构化 Excel 测试报告中自动提取测量数据
- **自动填表**：将提取数据写入 Voice Tuning Checklist 模板
- **报告审查**：对 Word 测试报告进行 17 项文档完整性检查
- **AI 图表分析**：通过多模态 LLM 对频响/响度曲线进行智能分析
- **跨报告一致性校验**：自动检测不同网络/编解码器间的数据一致性

当前支持终端模式：手机 (HA)、免提 (HH/HF)、耳机 (HE/HS)、电气接口 (EI)。

## 功能概览

### 仪表盘
- 审查统计（总数、通过率、分类分布）
- 快捷入口卡片
- 最近审查动态

### 测试数据收集
- 拖拽上传 ACQUA 测试报告（.xlsx / .xls / .doc / .docx）
- 上传 Voice Tuning Checklist 模板
- 自动识别终端模式并匹配规则配置
- 级联参数选择（接口 → 网络 → 编解码器 → 码率）
- 批量处理与实时进度追踪
- 结论输出：Excel 覆盖率、Word 审查结果、跨报告一致性

### 报告审查
- 成对报告审查（.docx + .xlsx）：17 项检查
- 单文档审查（.doc / .docx）：8 项文档结构检查
- 检查维度：文档完整性、曲线章节定位、POLQA 配置、时序、元数据、内容一致性
- AI 图表分析：频响曲线与响度曲线的趋势一致性、单调性、数值交叉验证
- 审查历史记录与详情查看

### 频谱分析（开发中）
- FFT / 小波 / 傅里叶分析框架（界面已预留）

### 设置
- 外观：自动/亮色/暗色主题，5 种设计风格（新拟物、玻璃拟态、纸张质感、柔和简约、经典专业）
- 系统托盘行为、默认输出目录、并发任务数
- LLM API 配置（OpenAI 兼容接口）与连通性测试
- 版本更新检测与下载

## 技术栈

| 技术 | 用途 |
|------|------|
| Electron 27 | 桌面应用壳 |
| React 18 + Ant Design 5 | 前端 UI |
| electron-builder | 打包与发布（NSIS 安装包 + 便携版） |
| electron-store | 本地设置持久化 |
| exceljs / xlsx | Excel 文件读写 |
| mammoth / word-extractor | Word 文档解析 |
| adm-zip / jszip | 直接操作 xlsx zip 结构 |
| recharts | 数据可视化 |
| axios | HTTP 客户端（LLM API、更新检查） |
| electron-updater | 应用内自动更新 |

## 快速开始

### 环境要求

- Node.js 18+
- npm 9+
- Windows 10/11（当前仅支持 Windows 平台）

### 开发运行

```bash
# 克隆仓库（国内推荐 Gitee）
git clone git@gitee.com:lingyu_mayun/TN_Audio_Check_Tools.git
cd TN_Audio_Check_Tools/TN_Audio_Tools

# 安装依赖
npm install

# 启动开发环境（React 开发服务器 + Electron）
npm start
```

### 构建

```bash
# 构建可分发的安装包
npm run build

# 正式发布（含前置检查 + GitHub Release 发布 + Gitee 同步）
npm run release
```

## 项目结构

```
TN_Audio_Tools/
├── src/
│   ├── main/                          # Electron 主进程
│   │   ├── main.js                    # 窗口管理、IPC、菜单、托盘
│   │   ├── preload.js                 # contextBridge 安全暴露 API
│   │   └── services/
│   │       ├── testDataExtraction/    # 数据提取管道（规则→解析→提取→填表）
│   │       ├── reportReview/          # 报告审查管道（17项检查 + AI图表）
│   │       ├── updater/               # 自动更新服务
│   │       └── settingsService.js     # 设置持久化
│   └── renderer/                      # React 渲染进程
│       ├── App.jsx                    # 根组件（路由、主题、侧边栏）
│       ├── pages/
│       │   ├── Dashboard.jsx          # 仪表盘
│       │   ├── TestDataCollectionPage.jsx  # 测试数据收集
│       │   ├── reportReviewModule/    # 报告审查
│       │   ├── SpectrumAnalyzer.jsx   # 频谱分析（开发中）
│       │   └── Settings.jsx           # 设置
│       ├── modules/testDataExtraction/config/  # 规则文件与 Checklist 模板
│       └── styles/                    # 主题 CSS + 设计风格定义
├── scripts/
│   ├── release/                       # 发布脚本
│   ├── diagnostics/                   # 诊断工具
│   └── dev/                           # 开发启动器
├── docs/                              # 项目文档
├── build/                             # React 构建输出
├── dist/                              # Electron 打包输出
├── package.json
└── update-manifest.json               # 远程版本清单（旧客户端兼容）
```

## 文档

完整文档见 [docs/](docs/) 目录：

| 文档 | 说明 |
|------|------|
| [音频测试报告全局解决方案](docs/音频测试报告全局解决方案.md) | 产品愿景、架构思路、迭代历史 |
| [数据提取系统架构说明](docs/数据提取系统架构说明.md) | 6层数据管道技术详设 |
| [报告审查说明与开发记录](docs/模块开发记录/报告审查说明与开发记录.md) | 审查功能交互流程与开发迭代 |
| [Git 开发与发布 SOP](docs/Git开发与发布SOP.md) | 分支策略、开发流程、发布步骤 |
| [AI 设计规范开发指南](docs/AI设计规范开发指南.md) | 多风格设计系统配置指南 |

## 协作方式

### 仓库架构

本项目采用 **Gitee 主仓 + GitHub 镜像** 双仓库架构：

- **Gitee**（主仓）：团队克隆、日常拉取、Release 下载
- **GitHub**（镜像）：个人开发、备份

`git push origin <branch>` 会同时推送到两个远端。

### 分支策略

- `master`：稳定发布基线
- `dev`：日常开发主线
- `feature/*`：功能分支（从 `dev` 创建，完成后合回 `dev`）

### 发布流程

1. 功能完成合回 `dev`
2. `dev` 合并到 `master`
3. 更新 `package.json`、`update-manifest.json` 版本号
4. 执行 `npm run release`（自动构建 + GitHub Release + Gitee Release 同步）

详见 [Git 开发与发布 SOP](docs/Git开发与发布SOP.md)。

## 许可证

内部工具，未开放外部使用许可。
