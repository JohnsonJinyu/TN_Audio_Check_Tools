# AI 设计规范开发指南

> 配置一次，所有项目、所有 AI 工具（Claude Code / Copilot / Codex）自动遵循统一设计规范。

---

## 一、已完成的全局配置

你本机已配置好以下基础设施，**不需要再做任何全局配置**：

| 配置项 | 位置 | 作用 |
|--------|------|------|
| Claude Code 全局指令 | `~/.claude/CLAUDE.md` | Claude Code 启动时自动读取，指引它去找项目的 AGENTS.md + DESIGN.md |
| Codex 全局指令 | `~/.codex/AGENTS.md` | Codex 启动时自动读取，同上 |
| Copilot 全局指令 | VS Code `settings.json` | Copilot agent mode 自动发现 AGENTS.md + 遵循 DESIGN.md |
| DESIGN.md 模板库 | `~/awesome-design-md/design-md/` | 70+ 品牌设计规范，直接复制使用 |
| VS Code Snippet | 输入 `designmd` → Tab | 一键生成空白 DESIGN.md 模板 |
| VS Code Snippet | 输入 `agents` → Tab | 一键生成 AGENTS.md 模板（含 DESIGN.md 引用） |

---

## 二、新项目快速上手（3 步）

### 步骤 1：生成 AGENTS.md
在项目根目录新建 `AGENTS.md`，输入 `agents`，按 Tab，填入项目信息。

### 步骤 2：选择或创建 DESIGN.md

**选项 A：从 70+ 品牌库直接复制**（推荐，见下方风格目录）
```powershell
# 示例：用 Linear 极简风
Copy-Item ~/awesome-design-md/design-md/linear.app/DESIGN.md ./DESIGN.md
```

**选项 B：用 snippet 自己写**
在项目根目录新建 `DESIGN.md`，输入 `designmd`，按 Tab，填入你的设计规范。

### 步骤 3：创建符号链接（让 Claude Code 和 Copilot 都能读到）
```powershell
New-Item -ItemType SymbolicLink -Path CLAUDE.md -Target AGENTS.md
New-Item -Force -ItemType SymbolicLink -Path .github/copilot-instructions.md -Target ..\AGENTS.md
git config core.symlinks true
```

之后对任何 AI 工具说：**"请遵循 AGENTS.md 和 DESIGN.md 来构建界面"** 即可。

---

## 三、风格选型指南

### 按项目类型推荐

| 项目类型 | 推荐风格 | 一句话特点 |
|----------|----------|------------|
| **SaaS 管理后台** | [Linear](#linearapp), [Notion](#notion), [Vercel](#vercel) | 极简专业、数据密集、信息层级清晰 |
| **开发者工具/CLI** | [Warp](#warp), [Raycast](#raycast), [Terminal](#voltagent) | 暗色终端风、代码优先、技术感 |
| **企业官网/品牌站** | [Apple](#apple), [Stripe](#stripe), [IBM](#ibm) | 大气留白、摄影驱动、品牌识别强 |
| **AI/LLM 产品** | [Claude](#claude), [OpenCode AI](#opencodeai), [Mistral](#mistralai) | 温暖人文或炫酷科技、对话式 UI |
| **电商/消费品牌** | [Nike](#nike), [Airbnb](#airbnb), [Tesla](#tesla) | 大面积摄影、极简 UI chrome、情感驱动 |
| **仪表盘/数据平台** | [ClickHouse](#clickhouse), [PostHog](#posthog), [Sentry](#sentry) | 数据密集、仪表盘布局、技术色彩 |
| **金融/支付** | [Stripe](#stripe), [Coinbase](#coinbase), [Revolut](#revolut) | 专业可信、渐变紫色系、精密排版 |
| **文档/知识库** | [Mintlify](#mintlify), [MongoDB](#mongodb), [Notion](#notion) | 阅读优先、清晰层级、代码高亮友好 |
| **创意/作品集** | [Framer](#framer), [Figma](#figma), [PlayStation](#playstation) | 大胆配色、动效驱动、年轻化 |

### 按视觉基调推荐

#### 亮色/白底为主（Clean Light）
适合：SaaS、企业站、文档站
- **[Apple](#apple)** — 白底 + 仅一种 Action Blue，极致减法，摄影驱动
- **[Vercel](#vercel)** — 白底黑字 + 渐变网格装饰，Geist 字体
- **[Notion](#notion)** — 暖白底 + 紫色 CTA + 柔和表面
- **[Stripe](#stripe)** — 深蓝 + 电光紫渐变，细体排版
- **[Supabase](#supabase)** — 白底 + 翡翠绿单色 CTA
- **[Airbnb](#airbnb)** — 暖珊瑚色 + 大圆角 + 摄影驱动
- **[Resend](#resend)** — 极简暗白 + 等宽字体强调

#### 暗色为主（Dark Mode）
适合：开发者工具、CLI、仪表盘、AI 工具
- **[Linear](#linearapp)** — 近纯黑画布 (#010102) + 唯一薰衣草蓝 CTA
- **[Cursor](#cursor)** — 暗色 IDE 风 + 渐变点缀
- **[Warp](#warp)** — 暗色终端 IDE + 块式命令 UI
- **[Raycast](#raycast)** — 暗色发射台 + 活力渐变
- **[Sentry](#sentry)** — 暗色仪表盘 + 粉紫点缀
- **[VoltAgent](#voltagent)** — 虚空黑 + 翡翠绿 + 终端原生
- **[ElevenLabs](#elevenlabs)** — 暗色电影感 + 音频波形美学

#### 暖调/人文感（Warm/Editorial）
适合：内容平台、写作工具、教育产品
- **[Claude](#claude)** — 奶油暖底 + 珊瑚 CTA + 衬线标题
- **[WIRED](#wired)** — 新闻纸白 + 定制衬线 + 墨蓝链接
- **[Mastercard](#mastercard)** — 暖奶油底色 + 轨道圆形 + 社论感
- **[Notion](#notion)** — 暖白极简 + 柔粉/薄荷/薰衣草卡片色

#### 极简/克制（Ultra-Minimal）
适合：奢侈品牌、高端产品、CEO 个人站
- **[Tesla](#tesla)** — 激进减法：全屏摄影 + 单一蓝色 CTA + 零装饰
- **[Ferrari](#ferrari)** — 明暗强烈对比 + 法拉利红极端克制使用
- **[Bugatti](#bugatti)** — 影院黑底 + 单色调 + 纪念碑式排版
- **[Nike](#nike)** — 黑白 UI + 全大写 Futura + 满版摄影

#### 活力/多彩（Playful/Vibrant）
适合：创意工具、社交、游戏
- **[Figma](#figma)** — 多色系统 + 趣味专业并存
- **[PostHog](#posthog)** — 趣味刺猬品牌 + 开发者友好暗色 UI
- **[Spotify](#spotify)** — 活力绿 + 暗底 + 专辑封面驱动
- **[PlayStation](#playstation)** — 三色通道布局 + 悬浮缩放交互

---

## 四、完整品牌目录

### AI & LLM 平台
- <a id="claude"></a>**Claude** — Anthropic 的 AI 助手。暖陶土色点缀，干净社论布局
- <a id="cohere"></a>**Cohere** — 企业 AI 平台。活力渐变，数据密集型仪表盘
- <a id="elevenlabs"></a>**ElevenLabs** — AI 语音平台。暗色电影感 UI，音频波形美学
- <a id="minimax"></a>**Minimax** — AI 模型提供商。大胆暗色界面 + 霓虹点缀
- <a id="mistralai"></a>**Mistral AI** — 开源 LLM。法式工程极简，紫色调
- <a id="ollama"></a>**Ollama** — 本地运行 LLM。终端优先，单色简洁
- <a id="opencodeai"></a>**OpenCode AI** — AI 编码平台。开发者中心暗色主题
- <a id="replicate"></a>**Replicate** — 云端运行 ML 模型。干净白底，代码优先
- <a id="runwayml"></a>**Runway** — AI 创意工具。电影社论美学，暗色英雄区
- <a id="togetherai"></a>**Together AI** — 开源 AI 基础设施。技术蓝图风
- <a id="voltagent"></a>**VoltAgent** — AI Agent 框架。虚空黑画布，翡翠绿点缀
- <a id="xai"></a>**xAI** — Elon Musk 的 AI。鲜明单色，未来极简

### 开发者工具
- <a id="cursor"></a>**Cursor** — AI 代码编辑器。暗色界面 + 渐变点缀
- <a id="expo"></a>**Expo** — React Native 平台。暗色主题 + 紧凑字距
- <a id="lovable"></a>**Lovable** — AI 全栈构建器。趣味渐变，友好开发者感
- <a id="raycast"></a>**Raycast** — 效率启动器。暗色金属风 + 活力渐变
- <a id="superhuman"></a>**Superhuman** — 快速邮件客户端。暗色高级 UI，紫色光晕
- <a id="vercel"></a>**Vercel** — 前端部署平台。黑白精准，Geist 字体
- <a id="warp"></a>**Warp** — 现代终端。暗色 IDE 风，块式命令 UI

### 后端/数据库/DevOps
- <a id="clickhouse"></a>**ClickHouse** — 高性能分析数据库。黄色点缀，技术文档风
- <a id="composio"></a>**Composio** — 工具集成平台。暗色 + 多彩集成图标
- <a id="hashicorp"></a>**HashiCorp** — IaC 工具。企业整洁黑白风
- <a id="mongodb"></a>**MongoDB** — 文档数据库。绿叶品牌 + 开发者文档
- <a id="posthog"></a>**PostHog** — 产品分析。趣味刺猬品牌，开发者暗色 UI
- <a id="sanity"></a>**Sanity** — 无头 CMS。暗色社论营销 + 珊瑚红 CTA
- <a id="sentry"></a>**Sentry** — 错误监控。暗色仪表盘，粉紫点缀
- <a id="supabase"></a>**Supabase** — 开源 Firebase。暗色翡翠主题，代码优先

### 生产力 & SaaS
- <a id="cal"></a>**Cal.com** — 开源日程。干净中性 UI
- <a id="intercom"></a>**Intercom** — 客户通讯。友好蓝色，对话式 UI
- <a id="linearapp"></a>**Linear** — 技术项目管理。极致极简，薰衣草蓝点缀
- <a id="mintlify"></a>**Mintlify** — 文档平台。绿色点缀，阅读优化
- <a id="notion"></a>**Notion** — 一体化工作区。暖色极简，衬线标题
- <a id="resend"></a>**Resend** — 邮件 API。暗色极简，等宽字体
- <a id="zapier"></a>**Zapier** — 自动化平台。暖橙色，插画驱动

### 设计 & 创意工具
- <a id="airtable"></a>**Airtable** — 表格数据库混合体。多彩友好
- <a id="clay"></a>**Clay** — 创意机构。有机形状，柔渐变
- <a id="figma"></a>**Figma** — 协作设计工具。多彩活力
- <a id="framer"></a>**Framer** — 建站工具。大胆黑白蓝，动效优先
- <a id="miro"></a>**Miro** — 白板协作。亮黄点缀，无限画布感
- <a id="webflow"></a>**Webflow** — 可视化建站。蓝色点缀，营销站美学

### 金融 & 加密
- <a id="binance"></a>**Binance** — 加密货币交易所。币安黄 + 单色暗底
- <a id="coinbase"></a>**Coinbase** — 加密货币交易。干净蓝色，信任优先
- <a id="kraken"></a>**Kraken** — 加密交易。紫色点缀暗色 UI
- <a id="mastercard"></a>**Mastercard** — 支付网络。暖奶油 + 轨道圆形
- <a id="revolut"></a>**Revolut** — 数字银行。暗色界面 + 渐变卡片
- <a id="stripe"></a>**Stripe** — 支付基础设施。签名紫色渐变 + 细体优雅
- <a id="wise"></a>**Wise** — 跨境汇款。亮绿点缀，友好清晰

### 电商 & 零售
- <a id="airbnb"></a>**Airbnb** — 旅游平台。暖珊瑚色 + 摄影驱动 + 圆角 UI
- <a id="meta"></a>**Meta** — 科技零售。摄影优先 + Meta 蓝 CTA
- <a id="nike"></a>**Nike** — 运动零售。黑白 UI + 全大写 Futura + 满版摄影
- <a id="shopify"></a>**Shopify** — 电商平台。暗色电影感 + 霓虹绿 + 超细标题字体
- <a id="starbucks"></a>**Starbucks** — 咖啡零售。四层大地绿系 + 暖奶油 + 专属 SoDoSans 字体

### 媒体 & 消费科技
- <a id="apple"></a>**Apple** — 消费电子。高级留白，SF Pro，电影感摄影
- <a id="ibm"></a>**IBM** — 企业科技。Carbon 设计系统，结构化蓝
- <a id="nvidia"></a>**NVIDIA** — GPU 计算。绿黑能量，技术力量感
- <a id="pinterest"></a>**Pinterest** — 视觉发现。红色点缀 + 瀑布流
- <a id="playstation"></a>**PlayStation** — 游戏主机。三色通道 + 悬浮放大交互
- <a id="spacex"></a>**SpaceX** — 太空科技。鲜明黑白 + 满版摄影
- <a id="spotify"></a>**Spotify** — 音乐流媒体。活力绿 + 暗底 + 粗体排版
- <a id="theverge"></a>**The Verge** — 科技编辑媒体。酸薄荷 + 紫外光点缀
- <a id="uber"></a>**Uber** — 出行平台。大胆黑白 + 紧凑字体
- <a id="vodafone"></a>**Vodafone** — 全球电信。纪念碑式大写 + Vodafone 红
- <a id="wired"></a>**WIRED** — 科技杂志。新闻纸白 + 定制衬线 + 墨蓝链接

### 汽车
- <a id="bmw"></a>**BMW** — 豪华汽车。暗色高级表面，德国工程精度
- <a id="bmw-m"></a>**BMW M** — 性能汽车。赛道灵感对比 + M 三色点缀
- <a id="bugatti"></a>**Bugatti** — 顶级超跑。影院黑，单色克制，纪念碑式排版
- <a id="ferrari"></a>**Ferrari** — 豪华跑车。明暗强烈对比 + 法拉利红极度克制
- <a id="lamborghini"></a>**Lamborghini** — 豪华超跑。真黑教堂 + 金色点缀
- <a id="renault"></a>**Renault** — 法国汽车。极光渐变 + 零圆角按钮
- <a id="tesla"></a>**Tesla** — 电动车。激进减法：100vh 全屏摄影 + 单一蓝色 CTA

---

## 五、如何指定 / 切换风格

### 指定风格
在项目根目录有 DESIGN.md 后，直接用自然语言告诉 AI：

| 场景 | 话术示例 |
|------|----------|
| 第一次指定 | "请遵循 DESIGN.md 中的设计规范来构建这个界面" |
| 强调具体要求 | "按钮用 DESIGN.md 里定义的主色（primary），圆角按组件规范来" |
| 引用特定章节 | "参考 DESIGN.md 第 2 节 Color Palette 和第 4 节 Component Stylings 来调整卡片样式" |
| Claude 特定 | "Please follow the DESIGN.md in the project root for all UI work" |
| Copilot 特定 | "根据 DESIGN.md 的配色方案，这个表单的提交按钮应该是什么颜色？" |

### 切换风格
只需替换 DESIGN.md 文件：

```powershell
# 当前用的是 Linear，想换成 Stripe 风格
Remove-Item ./DESIGN.md
Copy-Item ~/awesome-design-md/design-md/stripe/DESIGN.md ./DESIGN.md
```

然后告诉 AI：**"DESIGN.md 已更新，请用新的设计规范重新构建这个界面"**。

### 混合风格（进阶）
你也可以让 AI 混合多个风格：
> "参考 DESIGN.md 中 Linear 的配色方案和 Notion 的组件样式，结合两者的优点来设计这个页面。"

或在 DESIGN.md 中混合引用：
```markdown
## 2. Color Palette
Based on Stripe DESIGN.md (purple gradient + dark navy)
## 4. Component Stylings
Based on Linear (tight spacing, hairline borders)
```

### 预览风格效果
在浏览器中打开 getdesign.md 网站预览（需联网）：
- 网址格式：`https://getdesign.md/<brand>/design-md`
- 例如：https://getdesign.md/stripe/design-md

或者直接阅读 YAML 头部的 `description` 字段也能清楚了解风格特点。

---

## 六、VS Code Snippets 速查

| Snippet 触发词 | 用途 | 生成内容 |
|---------------|------|----------|
| `agents` + Tab | 快速生成 AGENTS.md | 完整技术规范模板（含 DESIGN.md 引用） |
| `designmd` + Tab | 快速生成 DESIGN.md | 完整 9 模块空白设计规范模板 |

---

## 七、常见问题

**Q: 没有 UI 需求的项目还需要 DESIGN.md 吗？**
不需要。纯后端/CLI 项目只需要 AGENTS.md 就够了。DESIGN.md 专治 AI 写前端 UI 时「每次风格都不一样」的毛病。

**Q: 我同时用 Claude Code 和 Copilot，会冲突吗？**
不会。它们都会通过自己的全局配置去读取同一份 AGENTS.md 和 DESIGN.md，保持一致性。

**Q: DESIGN.md 可以提交到 Git 吗？**
可以且应该提交。它是项目的设计真相来源，和代码一起版本管理。

**Q: 为什么用 AGENTS.md 而不是 CLAUDE.md？**
AGENTS.md 是跨工具标准（Copilot + Codex + Cursor + Windsurf 等 20+ 工具都支持），CLAUDE.md 只是 Claude Code 专用。用符号链接 `CLAUDE.md → AGENTS.md` 实现一份内容两个文件名都生效。

---

> 本地 DESIGN.md 库位置：`C:\Users\Lenovo\awesome-design-md\design-md\`
> 下次忘了怎么用，搜索「AI设计规范」打开本文档即可。
