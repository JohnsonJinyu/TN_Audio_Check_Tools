$root = Split-Path -Parent $PSScriptRoot
$assetsDir = Join-Path $root 'assets'
$pngPath = Join-Path $assetsDir 'icon-preview.png'
$icoPath = Join-Path $assetsDir 'icon.ico'

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw '未找到 node，无法生成标准 ico 文件。'
}

if (-not (Test-Path $pngPath)) {
  throw "缺少图标源文件: $pngPath"
}

# Use png-to-ico to produce a standards-compliant multi-size Windows icon.
$iconGenerator = @'
const fs = require('fs');
const pngToIco = require('png-to-ico').default;

async function main() {
  const [pngPath, icoPath] = process.argv.slice(2);
  const icoBuffer = await pngToIco(pngPath);
  fs.writeFileSync(icoPath, icoBuffer);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
'@

$scriptPath = Join-Path $assetsDir 'generate-icon.cjs'
Set-Content -Path $scriptPath -Value $iconGenerator -Encoding UTF8
$exitCode = 0

try {
  node $scriptPath $pngPath $icoPath
  $exitCode = $LASTEXITCODE
} finally {
  Remove-Item $scriptPath -Force -ErrorAction SilentlyContinue
}

if ($exitCode -ne 0) {
  throw "生成 icon.ico 失败，退出码: $exitCode"
}

Write-Output "Generated icon.ico ($((Get-Item $icoPath).Length) bytes) and icon-preview.png"
