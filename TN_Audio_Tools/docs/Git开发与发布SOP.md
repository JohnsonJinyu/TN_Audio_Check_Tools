# Git 开发与发布 SOP

**更新日期**：2026-06-04

这份文档只回答一个问题：这个仓库后续应该怎么做日常开发、功能分支合并和正式发布。

当前约定：

- `dev` 是日常开发主线
- `master` 是稳定发布主线
- 新需求默认从 `dev` 拉功能分支
- 发版时再把 `dev` 合到 `master`

---

## 一、分支职责

### `dev`

用途：

- 日常开发基线
- 新功能、修复、重构先合到这里
- 联调、自测、阶段性集成都以它为准

要求：

- 开发前先同步最新 `dev`
- 不建议长期堆积未验证的大改动
- 小功能可以快速合回，但要保持提交清晰

### `master`

用途：

- 正式发布基线
- 打 tag、发 Release、承载线上稳定版本

要求：

- 不在 `master` 上做日常开发
- 只接收已经在 `dev` 上验证过的内容
- 每次发布后，`master` 和 `dev` 尽量重新对齐

### `feature/*`

用途：

- 单个需求
- 单个修复
- 单个重构主题

要求：

- 一个分支只做一类事情
- 默认从 `dev` 创建
- 完成后优先合回 `dev`

建议命名：

- `feature/report-review-xxx`
- `feature/update-flow-xxx`
- `fix/settings-xxx`
- `refactor/test-data-xxx`

---

## 二、日常开发标准流程

### 1. 开始新需求前

先回到开发基线：

```bash
git switch dev
git pull origin dev
```

然后从 `dev` 拉新分支：

```bash
git switch -c feature/xxx
```

### 2. 在功能分支开发

开发过程中只在当前功能分支提交：

```bash
git add .
git commit -m "feat: xxx"
```

提交建议：

- 一次提交只表达一个明确意图
- 不把无关改动混进同一个提交
- 先保证能回退，再追求提交数量少

### 3. 开发过程中同步主线

如果开发周期较长，期间 `dev` 可能前进，需要同步一次：

```bash
git switch dev
git pull origin dev
git switch feature/xxx
git merge dev
```

如果你更偏好线性历史，也可以改用 rebase，但当前仓库不强制。

### 4. 功能完成后合回 `dev`

```bash
git switch dev
git pull origin dev
git merge feature/xxx
git push origin dev
```

合并后建议删除已经完成的功能分支：

```bash
git branch -d feature/xxx
```

如果远端也存在：

```bash
git push origin --delete feature/xxx
```

---

## 三、发布标准流程

发布时，不直接从功能分支发版，而是从 `dev` 收口到 `master`。

### 1. 发布前收口

确保 `dev` 是本次准备发布的最终内容：

```bash
git switch dev
git pull origin dev
```

建议检查：

- 功能是否都已合回 `dev`
- 工作区是否干净
- 版本号是否已更新
- `update-manifest.json` 是否与版本一致

### 2. 合并到 `master`

```bash
git switch master
git pull origin master
git merge dev
git push origin master
```

如果发布后希望继续保持两个主干整齐，确认 `master` 没有额外热修时，可以把 `dev` 再对齐到 `master`。

### 3. 打包与发布

当前仓库正式发布命令：

```bash
npm run release
```

发布前脚本会校验：

- GitHub CLI 已登录
- `package.json` 版本合法
- 远端不存在同名 Release
- Windows 构建目标包含 `nsis`
- Gitee API Token 已配置（警告，非阻断）

### 3.5 同步 Release 资产到 Gitee

`npm run release` 完成后会自动执行 `release:gitee-sync`，将构建产物同步到 Gitee Releases。

依赖前提：

- 环境变量 `GITEE_API_TOKEN` 已配置（Gitee 私人令牌，需 `projects` 权限）
- 令牌获取方式：Gitee → 设置 → 私人令牌 → 生成新令牌

同步内容：

- NSIS 安装包 exe
- blockmap 文件
- latest.yml

如果 Gitee 同步失败（网络、Token 等问题），不会阻断发布流程。可单独重新执行：

```bash
npm run release:gitee-sync
```

也可以手动通过 Gitee Web 界面补充上传资产。

### 4. 发布后验证

至少检查下面四项：

- GitHub Release 已创建
- Release 资产包含 `latest.yml`、安装版 exe、blockmap
- `update-manifest.json` 已同步到对应版本
- 远端 manifest 能返回正确版本

---

## 四、这个仓库的特殊注意点

### 1. `update-manifest.json` 不能漏

这个项目的旧客户端主要依赖远程 `update-manifest.json` 做版本检查，不是只看 GitHub Release。

所以每次正式发版时，至少要同步这三处：

- `package.json`
- `package-lock.json`
- `update-manifest.json`

如果 Release 已发布，但 manifest 没更新，旧客户端仍然可能看不到新版本。

### 2. `origin` 以 Gitee 为主，GitHub 为镜像

当前仓库的 `origin` 配置：

- **fetch/pull 来源**：Gitee（国内访问更快）
- **push 目标**：同时推送到 Gitee 和 GitHub（双 push URL）

这意味着：

- `git pull` / `git fetch` 默认从 Gitee 拉取
- `git push origin master` 会同时推到 Gitee 和 GitHub
- 新同事 clone 仓库时默认从 Gitee clone

如果 Gitee 不可用，可以临时从 GitHub 拉取：

```bash
git pull git@github.com:JohnsonJinyu/TN_Audio_Check_Tools.git <branch>
```

如果已经用 GitHub clone 了仓库，需要更新 remote：

```bash
git remote set-url origin git@gitee.com:lingyu_mayun/TN_Audio_Check_Tools.git
```

双 push 时如果一侧失败，单独重试：

```bash
git push git@gitee.com:lingyu_mayun/TN_Audio_Check_Tools.git <branch>
git push git@github.com:JohnsonJinyu/TN_Audio_Check_Tools.git <branch>
```

### 3. 验证 manifest 时不要只信 raw 页面缓存

如果远端 raw 页面看起来还是旧版本，优先用下面方式确认源站真实状态：

- GitHub API
- 分支 SHA
- 带时间戳的 no-cache 请求

不要只凭浏览器里直接打开的 raw 页面判断是否推送成功。

### 4. 发布后验证清单

每次正式发版后，至少检查以下六项：

1. GitHub Release 已创建，资产包含 `latest.yml`、安装版 exe、blockmap
2. Gitee Release 已创建，资产与 GitHub 一致
3. `update-manifest.json` 已同步到对应版本，且 `downloads` 数组包含 Gitee 下载地址
4. 远端 manifest（Gitee raw）能返回正确版本号和下载 URL
5. 应用内「检查更新」能正常检测到新版本
6. 应用内下载能通过 Gitee 链路完成

---

## 五、不推荐的做法

- 不要直接在 `master` 上做日常开发
- 不要从 `master` 开常规功能分支
- 不要把多个需求混在一个 `feature` 分支里
- 不要长时间让 `dev` 和 `master` 完全脱节不整理
- 不要发布后只传 Release 资产、不同步 manifest

---

## 六、最简工作模板

### 日常开发

```bash
git switch dev
git pull origin dev
git switch -c feature/xxx

# 开发...

git add .
git commit -m "feat: xxx"
git push origin feature/xxx
```

### 功能完成

```bash
git switch dev
git pull origin dev
git merge feature/xxx
git push origin dev
git branch -d feature/xxx
```

### 正式发布

```bash
# 1. 确保环境变量已配置（可选，不配置则跳过 Gitee 同步）
set GITEE_API_TOKEN=your_gitee_token

# 2. 合并到 master 并推送（同时推 GitHub + Gitee）
git switch master
git pull origin master
git merge dev
git push origin master

# 3. 构建并发布（自动上传 GitHub Release + Gitee Release）
npm run release

# 4. 如果 Gitee 同步失败，单独重试
npm run release:gitee-sync
```

---

## 七、一句话原则

日常开发站在 `dev`，单个需求放进 `feature/*`，功能完成先合回 `dev`，正式发版再从 `dev` 合到 `master`。