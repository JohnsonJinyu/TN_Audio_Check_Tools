const path = require('path');
const { spawn } = require('child_process');

const COM_STYLE_TIMEOUT_MS = 120000;

function runProcess(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const { timeoutMs = 0, ...spawnOptions } = options;
    const child = spawn(command, args, {
      windowsHide: true,
      ...spawnOptions
    });

    let stdout = '';
    let stderr = '';
    let finished = false;

    const timeoutHandle = timeoutMs > 0
      ? setTimeout(() => {
          if (finished) {
            return;
          }

          finished = true;
          child.kill();
          reject(new Error(`Command timed out after ${timeoutMs}ms: ${command}`));
        }, timeoutMs)
      : null;

    const finalize = (handler) => {
      if (finished) {
        return;
      }

      finished = true;
      if (timeoutHandle) {
        clearTimeout(timeoutHandle);
      }
      handler();
    };

    child.stdout?.on('data', (chunk) => {
      stdout += String(chunk);
    });

    child.stderr?.on('data', (chunk) => {
      stderr += String(chunk);
    });

    child.on('error', (error) => {
      finalize(() => reject(error));
    });

    child.on('close', (code) => {
      finalize(() => {
        if (code === 0) {
          resolve({ stdout, stderr });
          return;
        }

        reject(new Error(stderr || stdout || `Command failed: ${command}`));
      });
    });
  });
}

async function styleChecklistWithCom({ outputPath, sheetName, decimalCells, percentCells, skippedCells }) {
  const scriptPath = path.join(__dirname, 'applyChecklistStyles.ps1');
  const { stdout } = await runProcess(
    'powershell',
    [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      scriptPath,
      '-WorkbookPath',
      outputPath,
      '-SheetName',
      sheetName,
      '-DecimalCells',
      (decimalCells || []).join(','),
      '-PercentCells',
      (percentCells || []).join(','),
      '-SkippedCells',
      (skippedCells || []).join(',')
    ],
    { timeoutMs: COM_STYLE_TIMEOUT_MS }
  );

  const engineMatch = stdout.match(/STYLE_ENGINE=([^\r\n]+)/);
  return engineMatch ? engineMatch[1].trim() : 'COM';
}

module.exports = {
  styleChecklistWithCom
};
