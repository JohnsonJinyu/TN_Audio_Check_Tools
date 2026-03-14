const fs = require('fs/promises');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

function createReportConverter({ childProcessTimeoutMs, libreOfficeCandidatePaths }) {
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

  async function pathExists(targetPath) {
    try {
      await fs.access(targetPath);
      return true;
    } catch {
      return false;
    }
  }

  async function queryWindowsAppPath(executableName) {
    const registryKeys = [
      `HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${executableName}`,
      `HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${executableName}`
    ];

    for (const registryKey of registryKeys) {
      try {
        const { stdout } = await runProcess('reg', ['QUERY', registryKey, '/ve']);
        const match = stdout.match(/REG_SZ\s+(.+)$/m);
        if (match && match[1]) {
          const executablePath = match[1].trim();
          if (await pathExists(executablePath)) {
            return executablePath;
          }
        }
      } catch {
        // Continue probing other registry keys.
      }
    }

    return null;
  }

  async function findLibreOfficeExecutable() {
    const appPath = await queryWindowsAppPath('soffice.exe');
    if (appPath) {
      return appPath;
    }

    for (const candidatePath of libreOfficeCandidatePaths) {
      if (await pathExists(candidatePath)) {
        return candidatePath;
      }
    }

    return 'soffice';
  }

  async function convertDocWithLibreOffice(reportPath, outputDir) {
    const candidates = [await findLibreOfficeExecutable(), 'soffice.exe'];

    for (const candidate of candidates) {
      try {
        await runProcess(candidate, ['--headless', '--convert-to', 'docx', '--outdir', outputDir, reportPath], {
          timeoutMs: childProcessTimeoutMs
        });
        const convertedPath = path.join(outputDir, `${path.parse(reportPath).name}.docx`);
        if (await pathExists(convertedPath)) {
          return convertedPath;
        }
      } catch {
        // Try the next converter candidate.
      }
    }

    return null;
  }

  async function convertDocWithCom(reportPath, outputDir, progId, formatCode) {
    const convertedPath = path.join(outputDir, `${path.parse(reportPath).name}.docx`);
    const escapedInput = reportPath.replace(/'/g, "''");
    const escapedOutput = convertedPath.replace(/'/g, "''");
    const script = [
      "$ErrorActionPreference = 'Stop'",
      `$application = New-Object -ComObject '${progId}'`,
      '$application.Visible = $false',
      '$application.DisplayAlerts = 0',
      `$document = $application.Documents.Open('${escapedInput}')`,
      'if ($document.PSObject.Methods.Name -contains "SaveAs2") {',
      `  $document.SaveAs2('${escapedOutput}', ${formatCode})`,
      '} else {',
      `  $document.SaveAs('${escapedOutput}', ${formatCode})`,
      '}',
      '$document.Close()',
      '$application.Quit()'
    ].join('; ');

    try {
      await runProcess('powershell', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script], {
        timeoutMs: childProcessTimeoutMs
      });
      return (await pathExists(convertedPath)) ? convertedPath : null;
    } catch {
      return null;
    }
  }

  async function convertDocWithWord(reportPath, outputDir) {
    return convertDocWithCom(reportPath, outputDir, 'Word.Application', 16);
  }

  async function convertDocWithWps(reportPath, outputDir) {
    const wpsProgIds = ['kwps.Application', 'wps.Application'];

    for (const progId of wpsProgIds) {
      const convertedPath = await convertDocWithCom(reportPath, outputDir, progId, 12);
      if (convertedPath) {
        return convertedPath;
      }
    }

    return null;
  }

  async function convertDocToTemporaryDocx(reportPath) {
    const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'tn-audio-report-'));

    try {
      const wordPath = await convertDocWithWord(reportPath, tempDir);
      if (wordPath) {
        return { tempDir, convertedPath: wordPath };
      }

      const wpsPath = await convertDocWithWps(reportPath, tempDir);
      if (wpsPath) {
        return { tempDir, convertedPath: wpsPath };
      }

      const libreOfficePath = await convertDocWithLibreOffice(reportPath, tempDir);
      if (libreOfficePath) {
        return { tempDir, convertedPath: libreOfficePath };
      }

      await fs.rm(tempDir, { recursive: true, force: true });
      return null;
    } catch (error) {
      await fs.rm(tempDir, { recursive: true, force: true });
      throw error;
    }
  }

  return {
    convertDocToTemporaryDocx
  };
}

module.exports = {
  createReportConverter
};
