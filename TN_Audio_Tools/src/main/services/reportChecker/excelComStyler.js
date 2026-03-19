const fs = require('fs/promises');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const COM_STYLE_TIMEOUT_MS = 300000;
const COM_UPDATE_BATCH_SIZE = 12;

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

async function styleChecklistWithCom({ outputPath, sheetName, decimalCells, percentCells, skippedCells, valueUpdates, reportUpdates }) {
  let styleEngine = 'COM';

  if (Array.isArray(reportUpdates) && reportUpdates.length > 0) {
    styleEngine = await runChecklistComPass({
      outputPath,
      sheetName,
      valueUpdates: [],
      reportUpdates,
      decimalCells: [],
      percentCells: [],
      skippedCells: []
    });
  }

  for (const batch of chunkArray(valueUpdates || [], COM_UPDATE_BATCH_SIZE)) {
    styleEngine = await runChecklistComPass({
      outputPath,
      sheetName,
      valueUpdates: batch,
      reportUpdates: [],
      decimalCells: [],
      percentCells: [],
      skippedCells: []
    });
  }

  styleEngine = await runChecklistComPass({
    outputPath,
    sheetName,
    valueUpdates: [],
    reportUpdates: [],
    decimalCells,
    percentCells,
    skippedCells
  });

  return styleEngine;
}

async function runChecklistComPass({ outputPath, sheetName, valueUpdates, reportUpdates, decimalCells, percentCells, skippedCells }) {
  const scriptPath = path.join(__dirname, 'applyChecklistStyles.ps1');
  const payloadPath = path.join(
    os.tmpdir(),
    `report-checker-com-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}.json`
  );

  await fs.writeFile(payloadPath, JSON.stringify({
    valueUpdates: valueUpdates || [],
    reportUpdates: reportUpdates || []
  }), 'utf8');

  let stdout = '';
  try {
    const result = await runProcess(
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
        (skippedCells || []).join(','),
        '-UpdatePayloadPath',
        payloadPath
      ],
      { timeoutMs: COM_STYLE_TIMEOUT_MS }
    );
    stdout = result.stdout;
  } finally {
    await fs.unlink(payloadPath).catch(() => {});
  }

  const engineMatch = stdout.match(/STYLE_ENGINE=([^\r\n]+)/);
  return engineMatch ? engineMatch[1].trim() : 'COM';
}

function chunkArray(items, chunkSize) {
  const source = Array.isArray(items) ? items : [];
  if (source.length === 0) {
    return [];
  }

  const chunks = [];
  for (let index = 0; index < source.length; index += chunkSize) {
    chunks.push(source.slice(index, index + chunkSize));
  }
  return chunks;
}

module.exports = {
  styleChecklistWithCom
};
