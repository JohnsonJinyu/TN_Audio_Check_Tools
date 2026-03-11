const http = require('http');
const { spawn } = require('child_process');

const HOST = '127.0.0.1';
const PORT = 3000;
const STARTUP_TIMEOUT_MS = 90000;
const POLL_INTERVAL_MS = 1500;

function getNpmCommand() {
  return process.platform === 'win32' ? 'npm.cmd' : 'npm';
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function checkDevServer() {
  return new Promise((resolve) => {
    const request = http.get(
      {
        host: HOST,
        port: PORT,
        path: '/',
        timeout: 2500
      },
      (response) => {
        response.resume();
        resolve(response.statusCode >= 200 && response.statusCode < 500);
      }
    );

    request.on('timeout', () => {
      request.destroy();
      resolve(false);
    });

    request.on('error', () => resolve(false));
  });
}

function spawnCommand(command, args, options = {}) {
  const baseOptions = {
    cwd: process.cwd(),
    stdio: 'inherit',
    windowsHide: false,
    ...options
  };

  if (process.platform === 'win32') {
    const commandLine = [command, ...args].join(' ');

    return spawn(commandLine, {
      ...baseOptions,
      shell: true
    });
  }

  return spawn(command, args, {
    ...baseOptions,
    shell: false
  });
}

function terminateChild(child) {
  if (!child || child.killed || child.exitCode !== null) {
    return Promise.resolve();
  }

  if (process.platform === 'win32') {
    return new Promise((resolve) => {
      const killer = spawn('taskkill', ['/pid', String(child.pid), '/t', '/f'], {
        stdio: 'ignore',
        windowsHide: true
      });

      killer.on('error', () => resolve());
      killer.on('close', () => resolve());
    });
  }

  child.kill('SIGTERM');
  return Promise.resolve();
}

async function waitForDevServer(reactProcess) {
  const deadline = Date.now() + STARTUP_TIMEOUT_MS;

  while (Date.now() < deadline) {
    if (reactProcess && reactProcess.exitCode !== null) {
      throw new Error('React 开发服务提前退出，未能成功启动。');
    }

    if (await checkDevServer()) {
      return true;
    }

    await sleep(POLL_INTERVAL_MS);
  }

  throw new Error('等待 React 开发服务超时，请检查 3000 端口是否被其它程序占用。');
}

async function startDev() {
  const npmCommand = getNpmCommand();
  let reactProcess = null;
  let cleaningUp = false;

  const cleanup = async () => {
    if (cleaningUp) {
      return;
    }

    cleaningUp = true;
    await terminateChild(reactProcess);
  };

  process.on('SIGINT', async () => {
    await cleanup();
    process.exit(130);
  });

  process.on('SIGTERM', async () => {
    await cleanup();
    process.exit(143);
  });

  try {
    const devServerUp = await checkDevServer();

    if (devServerUp) {
      console.log('检测到 React 开发服务已经在 3000 端口运行，直接启动 Electron。');
    } else {
      console.log('未检测到 React 开发服务，正在启动 react-scripts。');
      reactProcess = spawnCommand(npmCommand, ['run', 'react-start'], {
        env: {
          ...process.env,
          BROWSER: 'none'
        }
      });

      reactProcess.on('exit', (code) => {
        if (code !== 0 && !cleaningUp) {
          console.error(`React 开发服务退出，退出码: ${code}`);
        }
      });

      await waitForDevServer(reactProcess);
      console.log('React 开发服务已就绪，正在启动 Electron。');
    }

    const electronProcess = spawnCommand(npmCommand, ['run', 'electron-dev']);

    electronProcess.on('exit', async (code) => {
      await cleanup();
      process.exit(code ?? 0);
    });
  } catch (error) {
    await cleanup();
    console.error(error.message || error);
    process.exit(1);
  }
}

if (require.main === module) {
  startDev();
}

module.exports = {
  startDev,
  checkDevServer
};