# reportChecker 模块说明

## 目录定位

`reportChecker` 目录负责测试报告检查填写的主流程拆分。现在的设计目标是把“入口编排、报告解析、规则提取、Excel 写入、样式处理、文档转换”分开，避免所有逻辑继续堆在 `reportCheckerService.js` 里。

## 模块职责

### `reportCheckerService.js`

- 组合入口。
- 负责装配各个子模块并导出 `processReports`。
- 不再直接承载解析细节、提取规则和 Excel 样式实现。

### `reportRunner.js`

- 负责批量处理流程编排。
- 校验报告、checklist、规则文件路径。
- 统一收集每个报告的成功/失败结果。

### `reportSource.js`

- 负责规则文件读取。
- 负责 `.doc` / `.docx` 报告内容解析。
- 输出统一的搜索数据结构，供后续提取模块使用。

### `reportConverter.js`

- 负责 `.doc` 到临时 `.docx` 的转换链路。
- 优先尝试 Word COM，其次 WPS COM，最后 LibreOffice。
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
- 第二阶段统一处理样式。

### `excelComStyler.js`

- 负责从 Node 调起 PowerShell COM 样式脚本。
- 只做脚本调用和结果回传，不关心具体边框规则。

### `applyChecklistStyles.ps1`

- 负责 Excel/WPS COM 样式落地。
- 当前规则是：`A3:K75` 外围粗线，内部细线。
- 同时处理对齐和 I 列数值格式。

## 当前调用关系

主入口调用顺序如下：

1. `reportCheckerService.js` 组装依赖。
2. `reportRunner.js` 接收批量报告并校验参数。
3. `reportSource.js` 读取规则并解析单份报告。
4. `reportExtractor.js` 按规则驱动 `reportAnalysis.js` 提取值。
5. `checklistWriter.js` 把结果写入 checklist。
6. `excelComStyler.js` 调起 `applyChecklistStyles.ps1` 做样式处理。
7. 若 COM 不可用，`checklistWriter.js` 回退到 `ExcelJS` 样式方案。

## 样式处理策略

当前样式链路是 COM 优先、库方案兜底：

1. `xlsx` 负责稳定写出数据。
2. `applyChecklistStyles.ps1` 优先尝试以下 COM 引擎：
   - `Excel.Application`
   - `ket.Application`
   - `et.Application`
3. 若 COM 不可用，再回退到 `ExcelJS`。

这么做的原因是：合并单元格、边框和模板样式恢复这类行为，用 COM 更接近人工在 Excel 或 WPS 里直接操作的结果。

## 后续维护建议

- 如果后面要继续瘦身，优先考虑再拆 `reportAnalysis.js`。
- 如果要新增提取类型，优先改 `reportExtractor.js` 的分发和 `reportAnalysis.js` 的实现。
- 如果要调整 checklist 样式，优先改 `applyChecklistStyles.ps1`，再确认 `ExcelJS` fallback 是否需要同步。