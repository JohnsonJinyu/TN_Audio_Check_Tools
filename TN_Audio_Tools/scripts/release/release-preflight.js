#!/usr/bin/env node

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const projectRoot = path.resolve(__dirname, '..', '..');
const packageJsonPath = path.join(projectRoot, 'package.json');

function run(command, options = {}) {
  return execSync(command, {
    cwd: projectRoot,
    stdio: ['ignore', 'pipe', 'pipe'],
    encoding: 'utf8',
    ...options
  });
}

function fail(message, details) {
  console.error(`\n[preflight] FAIL: ${message}`);
  if (details) {
    console.error(details);
  }
  process.exit(1);
}

function info(message) {
  console.log(`[preflight] ${message}`);
}

function warn(message) {
  console.warn(`[preflight] WARN: ${message}`);
}

function loadPackageJson() {
  if (!fs.existsSync(packageJsonPath)) {
    fail('package.json 不存在，无法执行发布前检查。');
  }

  try {
    return JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
  } catch (error) {
    fail('package.json 解析失败。', error.message);
  }
}

function checkGitHubCli() {
  try {
    const version = run('gh --version').split('\n')[0].trim();
    info(`GitHub CLI: ${version}`);
  } catch (error) {
    fail('未检测到 gh 命令，请先安装 GitHub CLI。', error.stderr || error.message);
  }

  try {
    run('gh auth status');
    info('GitHub 认证状态正常。');
  } catch (error) {
    fail('gh 未登录或认证失效，请先执行 gh auth login。', error.stderr || error.message);
  }
}

function checkGitState() {
  try {
    const status = run('git status --porcelain').trim();
    if (status) {
      warn('当前工作区有未提交变更，建议先提交后再发布。');
    } else {
      info('Git 工作区干净。');
    }
  } catch (error) {
    warn(`无法检查 git 状态: ${error.message}`);
  }
}

function resolvePublishTarget(pkg) {
  const publish = pkg?.build?.publish;
  if (!Array.isArray(publish) || publish.length === 0) {
    fail('build.publish 未配置，无法确认发布目标。');
  }

  const githubTarget = publish.find((item) => item && item.provider === 'github');
  if (!githubTarget) {
    fail('build.publish 未配置 github provider。');
  }

  const owner = githubTarget.owner;
  const repo = githubTarget.repo;

  if (!owner || !repo) {
    fail('github provider 缺少 owner/repo 配置。');
  }

  return { owner, repo };
}

function checkVersion(pkg, owner, repo) {
  const version = pkg?.version;
  if (!version || !/^\d+\.\d+\.\d+([-.][0-9A-Za-z.]+)?$/.test(version)) {
    fail(`package.json version 非法: ${version || '(empty)'}`);
  }

  const tag = `v${version}`;
  info(`准备发布版本: ${version} (${tag})`);

  try {
    run(`gh release view ${tag} --repo ${owner}/${repo}`);
    fail(`远端已存在 Release ${tag}，请先删除或升级版本号。`);
  } catch (error) {
    const stderr = (error.stderr || '').toLowerCase();
    if (stderr.includes('not found') || stderr.includes('could not resolve to a release')) {
      info(`远端未发现同名 Release: ${tag}`);
      return;
    }

    fail('检查远端 Release 状态失败。', error.stderr || error.message);
  }
}

function checkBuildTargets(pkg) {
  const targets = pkg?.build?.win?.target;
  if (!Array.isArray(targets)) {
    warn('未检测到 build.win.target 数组，请确认包含 nsis 目标。');
    return;
  }

  const names = targets
    .map((target) => (typeof target === 'string' ? target : target?.target))
    .filter(Boolean);

  if (!names.includes('nsis')) {
    fail('build.win.target 未包含 nsis，electron-updater 无法生成 latest.yml。');
  }

  info(`Windows 构建目标: ${names.join(', ')}`);
}

function checkGiteeToken() {
  const token = process.env.GITEE_API_TOKEN;
  if (!token) {
    warn('GITEE_API_TOKEN 未设置，构建后不会同步到 Gitee Release。');
    warn('如需同步，请设置环境变量: set GITEE_API_TOKEN=your_token');
    return;
  }

  try {
    const https = require('https');
    const url = `https://gitee.com/api/v5/user?access_token=${encodeURIComponent(token)}`;
    https.get(url, { timeout: 10000 }, (res) => {
      let body = '';
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => {
        if (res.statusCode === 200) {
          try {
            const user = JSON.parse(body);
            info(`Gitee Token 有效 (用户: ${user.login || user.name || 'unknown'})。`);
          } catch {
            warn('Gitee Token 验证响应解析失败，继续。');
          }
        } else {
          warn(`Gitee Token 验证失败 (HTTP ${res.statusCode})，同步可能不成功。`);
        }
      });
    }).on('error', () => {
      warn('无法连接 Gitee API，请检查网络。');
    });
  } catch (err) {
    warn(`Gitee Token 验证异常: ${err.message}`);
  }
}

function main() {
  info('开始执行发布前自检...');

  const pkg = loadPackageJson();
  const { owner, repo } = resolvePublishTarget(pkg);

  checkGitHubCli();
  checkGitState();
  checkVersion(pkg, owner, repo);
  checkBuildTargets(pkg);
  checkGiteeToken();

  info('发布前自检通过，可以执行 npm run release。');
}

main();
