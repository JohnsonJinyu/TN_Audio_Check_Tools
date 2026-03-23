# testDataExtraction 模块说明

## 目录定位

`testDataExtraction` 目录负责测试数据提取与 checklist 回填的主流程拆分。现在的设计目标是把“入口编排、报告解析、规则提取、Excel 写入、样式处理、文档转换”分开，避免所有逻辑继续堆在 `index.js` 之外的单个大文件里。

## 模块职责

### `index.js`

- 组合入口。
- 负责装配各个子模块并导出 `processReports`。
- 不再直接承载解析细节、提取规则和 Excel 样式实现。

### `reportRunner.js`

- 负责批量处理流程编排。
- 校验报告、checklist、规则文件路径。
- 统一收集每个报告的成功/失败结果。

### `reportSource.js`

- 负责规则文件读取。
- 负责按扩展名分发报告解析流程。
- `.xlsx` / `.xls` 走结构化 Excel 报告解析。
- `.doc` / `.docx` 继续走原有 Word 报告解析。
- 输出统一的搜索数据结构，供后续提取模块使用。

### `xlsxReportSource.js`

- 负责 ACQUA 导出的 `.xlsx` / `.xls` 结构化报告解析。
- 主要读取 `Detailed` / `Values` 工作表并构造成统一搜索数据。
- 与 Word 解析链路分离，便于后续独立扩展 Excel 主方案。

### `reportConverter.js`

- 负责 `.doc` 到临时 `.docx` 的转换链路。
- 使用 `word-extractor` 提取文本并通过 `jszip` 生成标准 `.docx` 包。
- 只返回转换结果，不参与内容解析。

### `reportAnalysis.js`

- 负责搜索数据构建和提取规则匹配。
- 包含行匹配、锚点提取、表格提取、正则提取等逻辑。
- 这是当前规则判断最集中的模块。

### `reportExtractor.js`

- 负责单份报告的提取执行。
- 根据规则的 `extractType` 调度 `reportAnalysis` 提供的各类提取函数。
- 汇总命中项、未命中项，并调用 checklist 写入模块生成输出。

### `checklistWriter.js`

- 负责把提取结果写回 checklist。
- 第一阶段使用 `xlsx` 落盘，避免直接处理模板时被 shared formula 卡住。
- 第二阶段只补充 I 列数值格式和 skip 单元格灰填充，其他边框和填充保持模板原样。

### `checklistStyleProfiles.js`

- 负责管理 Handset / Handsfree / Headset 三类 checklist 的样式 profile。
- 当前只保留模式识别和 skip 灰填充配置。

### `checklistLibraryStyler.js`

- 负责基于 xlsx zip 结构直接补充样式。
- 当前处理数字格式和 skip 项灰填充，不依赖 Office / WPS COM。
- 不重写模板边框和底色，只在现有样式基础上追加必要样式变体。

## 当前调用关系

主入口调用顺序如下：

1. `index.js` 组装依赖。
2. `reportRunner.js` 接收批量报告并校验参数。
3. `reportSource.js` 读取规则并解析单份报告。
4. `reportExtractor.js` 按规则驱动 `reportAnalysis.js` 提取值。
5. `checklistWriter.js` 把结果写入 checklist。
6. `checklistLibraryStyler.js` 直接修改输出 xlsx 的样式 XML。
7. `checklistStyleProfiles.js` 为 `checklistWriter.js` 和库样式层提供统一的样式 profile。

## 样式处理策略

当前 checklist 样式链路是纯库实现：

1. `xlsx` 负责稳定写出数据。
2. `adm-zip` 直接修改 workbook / worksheet / styles XML。
3. 模板边框、填充和 sheet 顺序尽量保持原模板，只对需要变化的值、数字格式和 skip 灰填充做最小补丁。

这么做的原因是：
1. 避免依赖用户本机安装 Excel / WPS；
2. 避免 COM 弹窗和用户误关闭；
3. 保持输出链路在无 Office 环境下也可稳定运行。

## 后续维护建议

- 如果后面要继续瘦身，优先考虑再拆 `reportAnalysis.js`。
- 如果要新增提取类型，优先改 `reportExtractor.js` 的分发和 `reportAnalysis.js` 的实现。
- 如果要调整 checklist 样式，优先改 `checklistLibraryStyler.js`，保持纯库实现，不要重新引入 COM 依赖。