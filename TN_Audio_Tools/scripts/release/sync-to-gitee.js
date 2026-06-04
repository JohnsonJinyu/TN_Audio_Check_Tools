#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const https = require('https');
const { execSync } = require('child_process');

const GITEE_OWNER = 'lingyu_mayun';
const GITEE_REPO = 'TN_Audio_Check_Tools';
const GITEE_API_BASE = 'gitee.com/api/v5';

const projectRoot = path.resolve(__dirname, '..', '..');
const packageJsonPath = path.join(projectRoot, 'package.json');
const distDir = path.join(projectRoot, 'dist');

function getToken() {
  return process.env.GITEE_API_TOKEN || '';
}

function info(msg) {
  console.log(`[gitee-sync] ${msg}`);
}

function warn(msg) {
  console.warn(`[gitee-sync] WARN: ${msg}`);
}

function fail(msg) {
  console.error(`[gitee-sync] ERROR: ${msg}`);
}

function apiRequest(method, apiPath, body, isUpload) {
  return new Promise((resolve, reject) => {
    const token = getToken();
    const url = new URL(`https://${GITEE_API_BASE}${apiPath}`);
    url.searchParams.set('access_token', token);

    const options = {
      method,
      hostname: GITEE_API_BASE.split('/')[0],
      path: `${url.pathname}${url.search}`,
      headers: {
        'User-Agent': 'TN-Audio-Toolkit/1.0',
        Accept: 'application/json'
      },
      timeout: 120000
    };

    if (body && !isUpload) {
      options.headers['Content-Type'] = 'application/json';
    }

    const req = https.request(options, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        let data;
        try {
          data = JSON.parse(raw);
        } catch {
          data = raw;
        }
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve({ statusCode: res.statusCode, data });
        } else {
          const err = new Error(`Gitee API ${res.statusCode}: ${typeof data === 'string' ? data : JSON.stringify(data)}`);
          err.statusCode = res.statusCode;
          err.body = data;
          reject(err);
        }
      });
    });

    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Request timed out'));
    });

    if (body) {
      if (isUpload) {
        req.write(body);
      } else {
        req.write(JSON.stringify(body));
      }
    }
    req.end();
  });
}

function loadPackageJson() {
  return JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
}

function findArtifacts(version) {
  const files = fs.readdirSync(distDir);
  const versionStr = version.replace(/^v/, '');

  const setupExe = files.find((f) =>
    f.endsWith('.exe') && f.includes('Setup') && f.includes(versionStr)
  );
  const blockmap = files.find((f) =>
    f.endsWith('.exe.blockmap') && f.includes(versionStr)
  );
  const latestYml = files.find((f) => f === 'latest.yml');

  const results = [];
  if (setupExe) results.push({ name: setupExe, path: path.join(distDir, setupExe), type: 'exe' });
  if (blockmap) results.push({ name: blockmap, path: path.join(distDir, blockmap), type: 'blockmap' });
  if (latestYml) results.push({ name: latestYml, path: path.join(distDir, latestYml), type: 'latest.yml' });

  return results;
}

function checkFileSize(filePath) {
  const stats = fs.statSync(filePath);
  const sizeMB = stats.size / (1024 * 1024);
  if (sizeMB > 100) {
    warn(`${path.basename(filePath)} 大小为 ${sizeMB.toFixed(1)}MB，超过 Gitee 附件限制 100MB，可能上传失败。`);
  }
}

async function getReleaseByTag(tag) {
  try {
    const { data } = await apiRequest(
      'GET',
      `/repos/${GITEE_OWNER}/${GITEE_REPO}/releases/tags/${encodeURIComponent(tag)}`
    );
    return data;
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

async function createRelease(tag, version) {
  const body = {
    tag_name: tag,
    name: `TN Audio Toolkit ${version}`,
    body: `Release ${tag} — 资产同步自 GitHub Release。`,
    target_commitish: 'master',
    prerelease: false
  };

  const { data } = await apiRequest(
    'POST',
    `/repos/${GITEE_OWNER}/${GITEE_REPO}/releases`,
    body
  );
  return data;
}

function uploadFile(releaseId, filePath, fileName) {
  return new Promise((resolve, reject) => {
    const token = getToken();
    const boundary = `----FormBoundary${Date.now()}`;
    const fileContent = fs.readFileSync(filePath);
    const ext = path.extname(fileName).toLowerCase();

    let mimeType = 'application/octet-stream';
    if (ext === '.yml') mimeType = 'text/yaml';
    if (ext === '.exe') mimeType = 'application/vnd.microsoft.portable-executable';

    const header = [
      `--${boundary}`,
      `Content-Disposition: form-data; name="file"; filename="${fileName}"`,
      `Content-Type: ${mimeType}`,
      '',
      ''
    ].join('\r\n');

    const footer = `\r\n--${boundary}--\r\n`;
    const headerBuf = Buffer.from(header, 'utf8');
    const footerBuf = Buffer.from(footer, 'utf8');
    const body = Buffer.concat([headerBuf, fileContent, footerBuf]);

    const apiPath = `/repos/${GITEE_OWNER}/${GITEE_REPO}/releases/${releaseId}/attach_files`;
    const url = new URL(`https://${GITEE_API_BASE}${apiPath}`);
    url.searchParams.set('access_token', token);

    const options = {
      method: 'POST',
      hostname: GITEE_API_BASE.split('/')[0],
      path: `${url.pathname}${url.search}`,
      headers: {
        'User-Agent': 'TN-Audio-Toolkit/1.0',
        'Content-Type': `multipart/form-data; boundary=${boundary}`,
        'Content-Length': String(body.length)
      },
      timeout: 300000
    };

    const req = https.request(options, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        let data;
        try { data = JSON.parse(raw); } catch { data = raw; }
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(data);
        } else {
          reject(new Error(`Upload failed ${res.statusCode}: ${typeof data === 'string' ? data : JSON.stringify(data)}`));
        }
      });
    });

    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Upload timed out'));
    });
    req.write(body);
    req.end();
  });
}

async function main() {
  info('开始同步 Release 资产到 Gitee...');

  const token = getToken();
  if (!token) {
    fail('GITEE_API_TOKEN 未设置。请设置环境变量后重试。');
    info('获取令牌: Gitee -> 设置 -> 私人令牌 -> 生成新令牌（权限: projects）');
    process.exit(0);
  }

  if (!fs.existsSync(distDir)) {
    fail('dist/ 目录不存在，请先执行 electron-builder 构建。');
    process.exit(0);
  }

  const pkg = loadPackageJson();
  const version = pkg.version;
  const tag = `v${version}`;
  info(`版本: ${version}, Tag: ${tag}`);

  const artifacts = findArtifacts(version);
  if (artifacts.length === 0) {
    fail(`dist/ 中未找到版本 ${version} 的构建产物。`);
    process.exit(0);
  }
  info(`找到 ${artifacts.length} 个构建产物: ${artifacts.map((a) => a.name).join(', ')}`);

  artifacts.forEach((a) => checkFileSize(a.path));

  let release = await getReleaseByTag(tag);
  if (release) {
    info(`Release ${tag} 已存在 (id=${release.id})。`);
    if (release.assets && release.assets.length > 0) {
      info(`已有 ${release.assets.length} 个附件，跳过上传（幂等）。`);
      info('Gitee Release: ' + (release.html_url || `https://gitee.com/${GITEE_OWNER}/${GITEE_REPO}/releases/tag/${tag}`));
      return;
    }
  } else {
    info(`创建新 Release: ${tag}`);
    release = await createRelease(tag, version);
    info(`Release 已创建 (id=${release.id})。`);
  }

  for (const artifact of artifacts) {
    info(`上传: ${artifact.name} (${(fs.statSync(artifact.path).size / (1024 * 1024)).toFixed(1)}MB)...`);
    try {
      await uploadFile(release.id, artifact.path, artifact.name);
      info(`${artifact.name} 上传成功。`);
    } catch (err) {
      fail(`${artifact.name} 上传失败: ${err.message}`);
      warn('可手动通过 Gitee Web 界面补充上传。');
    }
  }

  info('Gitee Release: ' + (release.html_url || `https://gitee.com/${GITEE_OWNER}/${GITEE_REPO}/releases/tag/${tag}`));
  info('同步完成。');
}

main().catch((err) => {
  fail(err.message);
  process.exit(0);
});
