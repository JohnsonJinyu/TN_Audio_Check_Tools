const http = require('http');
const net = require('net');
const { spawn } = require('child_process');

const HOST = '127.0.0.1';
const DEFAULT_PORT = 3123;
const MAX_PORT_SEARCH_COUNT = 20;
const STARTUP_TIMEOUT_MS = 90000;
const POLL_INTERVAL_MS = 1500;

function getNpmCommand() {
  return process.platform === 'win32' ? 'npm.cmd' : 'npm';
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function checkDevServer(port = DEFAULT_PORT) {
  return new Promise((resolve) => {
    const request = http.get(
      {
        host: HOST,
        port,
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

function isPortAvailable(port) {
  return new Promise((resolve) => {
    const server = net.createServer();

    server.once('error', () => {
      resolve(false);
    });

    server.once('listening', () => {
      server.close(() => resolve(true));
    });

    server.listen(port, HOST);
  });
}

async function findAvailablePort(startPort = DEFAULT_PORT) {
  for (let offset = 0; offset < MAX_PORT_SEARCH_COUNT; offset += 1) {
    const candidatePort = startPort + offset;
    if (await isPortAvailable(candidatePort)) {
      return candidatePort;
    }
  }

  throw new Error(`未能在 ${startPort}-${startPort + MAX_PORT_SEARCH_COUNT - 1} 之间找到可用端口。`);
}

async function resolveDevServerContext() {
  if (await checkDevServer(DEFAULT_PORT)) {
    return {
      port: DEFAULT_PORT,
      url: `http://${HOST}:${DEFAULT_PORT}`,
      hasRunningServer: true
    };
  }

  const port = await findAvailablePort(DEFAULT_PORT);
  return {
    port,
    url: `http://${HOST}:${port}`,
    hasRunningServer: false
  };
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

async function waitForDevServer(reactProcess, port) {
  const deadline = Date.now() + STARTUP_TIMEOUT_MS;

  while (Date.now() < deadline) {
    if (reactProcess && reactProcess.exitCode !== null) {
      throw new Error('React 开发服务提前退出，未能成功启动。');
    }

    if (await checkDevServer(port)) {
      return true;
    }

    await sleep(POLL_INTERVAL_MS);
  }

  throw new Error(`等待 React 开发服务超时，请检查 ${port} 端口是否被其它程序占用。`);
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
    const devServerContext = await resolveDevServerContext();

    if (devServerContext.hasRunningServer) {
      console.log(`检测到 React 开发服务已经在 ${devServerContext.port} 端口运行，直接启动 Electron。`);
    } else {
      if (devServerContext.port !== DEFAULT_PORT) {
        console.log(`3000 端口不可用，将改用 ${devServerContext.port} 端口启动 React 开发服务。`);
      } else {
        console.log('未检测到 React 开发服务，正在启动 react-scripts。');
      }

      reactProcess = spawnCommand(npmCommand, ['run', 'react-start'], {
        env: {
          ...process.env,
          BROWSER: 'none',
          HOST,
          PORT: String(devServerContext.port),
          WDS_SOCKET_HOST: HOST,
          WDS_SOCKET_PORT: String(devServerContext.port),
          WDS_SOCKET_PATH: '/ws'
        }
      });

      reactProcess.on('exit', (code) => {
        if (code !== 0 && !cleaningUp) {
          console.error(`React 开发服务退出，退出码: ${code}`);
        }
      });

      await waitForDevServer(reactProcess, devServerContext.port);
      console.log('React 开发服务已就绪，正在启动 Electron。');
    }

    const electronProcess = spawnCommand(npmCommand, ['run', 'electron-dev'], {
      env: {
        ...process.env,
        ELECTRON_DISABLE_SECURITY_WARNINGS: 'true',
        ELECTRON_RENDERER_URL: devServerContext.url
      }
    });

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