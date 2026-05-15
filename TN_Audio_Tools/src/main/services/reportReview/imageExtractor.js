const fs = require('fs/promises');
const fsSync = require('fs');
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');
const JSZip = require('jszip');

var IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.wmf']);

var MAX_IMAGE_SIZE_BYTES = 5 * 1024 * 1024;

// 响度/频响曲线图关键词
var CHART_INCLUDE = /Receiving\s*Dir|Sending\s*Dir|MAX-[1-9]|RLR|SLR|STMR|Frequency\s*Response|频响|响度|SND|RCV|NOM\b|Sensitivity|DRP|ERP|LIN\s*EQ|DF\s*AVG/i;

// 排除非频响/响度的图
var CHART_EXCLUDE = /Distortion|Echo\s*[Cc]ontrol|Background\s*Noise|Ambient\s*Noise|Acoustic\s*Shock|D-Value|ANR|Single\s*Talk|Double\s*Talk|Preparation|Level\s*vs\.?\s*Time|Speech\s*quality/i;

function guessContentType(fileName) {
  var ext = path.extname(fileName).toLowerCase();
  if (ext === '.png') return 'image/png';
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg';
  if (ext === '.gif') return 'image/gif';
  if (ext === '.bmp') return 'image/bmp';
  if (ext === '.webp') return 'image/webp';
  if (ext === '.wmf') return 'image/wmf';
  return 'image/png';
}

// WMF→PNG 批量转换 (Windows PowerShell + .NET System.Drawing)
function convertWmfBatchToPng(wmfEntries) {
  return new Promise(function(resolve) {
    if (wmfEntries.length === 0) { resolve([]); return; }

    var tmpDir = os.tmpdir();
    var batchId = Date.now();
    var fileList = [];

    wmfEntries.forEach(function(entry, idx) {
      var wmfPath = path.join(tmpDir, 'tn_wmf_' + batchId + '_' + idx + '.wmf');
      var pngPath = path.join(tmpDir, 'tn_wmf_' + batchId + '_' + idx + '.png');
      fileList.push({ wmf: wmfPath, png: pngPath });
      fsSync.writeFileSync(wmfPath, entry.buffer);
    });

    // 构建批量转换PowerShell脚本
    var psLines = [
      '[Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null'
    ];
    fileList.forEach(function(f) {
      psLines.push(
        'try {',
        '  $img = [System.Drawing.Image]::FromFile("' + f.wmf.replace(/\\/g, '\\\\') + '")',
        '  $img.Save("' + f.png.replace(/\\/g, '\\\\') + '", [System.Drawing.Imaging.ImageFormat]::Png)',
        '  $img.Dispose()',
        '  Write-Output "OK"',
        '} catch {',
        '  Write-Output "ERR:$($_.Exception.Message)"',
        '}'
      );
    });

    var psScript = psLines.join('\n');

    execFile('powershell', ['-NoProfile', '-Command', psScript], { timeout: 60000, windowsHide: true }, async function(err) {
      // 收集结果
      var results = [];
      for (var i = 0; i < fileList.length; i++) {
        var f = fileList[i];
        try {
          if (fsSync.existsSync(f.png)) {
            var pngBuf = await fs.readFile(f.png);
            results.push({ ok: true, pngBuffer: pngBuf, sizeBytes: pngBuf.length });
          } else {
            results.push({ ok: false });
          }
        } catch (_) {
          results.push({ ok: false });
        }
        // 清理临时文件
        try { fsSync.unlinkSync(f.wmf); } catch (_) {}
        try { fsSync.unlinkSync(f.png); } catch (_) {}
      }
      resolve(results);
    });
  });
}

async function extractReportImages(reportPath) {
  var buffer;
  try { buffer = await fs.readFile(reportPath); } catch (_) { return []; }

  var zip;
  try { zip = await JSZip.loadAsync(buffer); } catch (_) { return []; }

  // 读取关系文件 rId → media/xxx.wmf
  var imageRelMap = {};
  try {
    var relsFile = zip.file('word/_rels/document.xml.rels');
    if (relsFile) {
      var relsXml = await relsFile.async('string');
      var relMatches = relsXml.matchAll(/<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*\/?>/gi);
      for (var _m of relMatches) {
        var rid = _m[1];
        var target = _m[2];
        if (/image/i.test(target) || IMAGE_EXTENSIONS.has(path.extname(target).toLowerCase())) {
          imageRelMap[rid] = target;
        }
      }
    }
  } catch (_) {}

  // 解析 document.xml 获取图片引用上下文
  var imageContextMap = {};
  try {
    var docFile = zip.file('word/document.xml');
    if (docFile) {
      var docXml = await docFile.async('string');
      var paragraphs = docXml.match(/<w:p[ >][\s\S]*?<\/w:p>/gi) || [];
      var paraTexts = paragraphs.map(function(p) {
        var texts = p.match(/<w:t[^>]*>([^<]*)<\/w:t>/gi) || [];
        return texts.map(function(t) { return t.replace(/<\/?w:t[^>]*>/gi, ''); }).join('');
      }).filter(Boolean);

      // DrawingML: <a:blip r:embed="rIdN"/>
      var dmlMatches = docXml.matchAll(/<(?:\w+:)?blip[^>]*(?:\w+:)?embed="([^"]*)"[^>]*\/?>/gi);
      // VML: <v:imagedata r:id="rIdN"/>
      var vmlMatches = docXml.matchAll(/<v:imagedata[^>]*r:id="([^"]*)"[^>]*\/?>/gi);

      var allRefs = [];
      for (var _dm of Array.from(dmlMatches)) allRefs.push(_dm[1]);
      for (var _vm of Array.from(vmlMatches)) allRefs.push(_vm[1]);

      allRefs.forEach(function(rid, idx) {
        var target = imageRelMap[rid];
        if (!target) return;
        var contextParts = [];
        var startIdx = Math.max(0, idx - 1);
        var endIdx = Math.min(paraTexts.length - 1, idx + 1);
        for (var i = startIdx; i <= endIdx; i++) {
          if (paraTexts[i]) contextParts.push(paraTexts[i]);
        }
        var contextText = contextParts.join(' | ');
        if (!imageContextMap[target] || contextText.length > imageContextMap[target].length) {
          imageContextMap[target] = contextText;
        }
      });
    }
  } catch (_) {}

  // 提取 word/media/ 下的图片
  var mediaFiles = zip.file(/^word\/media\/.+/i) || [];
  var images = [];
  var wmfToConvert = [];

  for (var i = 0; i < mediaFiles.length; i++) {
    var entry = mediaFiles[i];
    var fileName = path.basename(entry.name);
    var ext = path.extname(fileName).toLowerCase();
    if (!IMAGE_EXTENSIONS.has(ext)) continue;

    var imgBuffer = await entry.async('nodebuffer');
    if (imgBuffer.length > MAX_IMAGE_SIZE_BYTES) continue;

    var target = 'media/' + fileName;
    var contextText = imageContextMap[target] || ('图片: ' + fileName);

    // 智能筛选：只保留响度/频响曲线图
    if (CHART_INCLUDE.test(contextText) && !CHART_EXCLUDE.test(contextText)) {
      if (ext === '.wmf') {
        wmfToConvert.push({ index: i, fileName: fileName, buffer: imgBuffer, contextText: contextText });
      } else {
        images.push({
          index: i, fileName: fileName,
          base64: imgBuffer.toString('base64'), contentType: guessContentType(fileName),
          contextText: contextText.slice(0, 500), sizeBytes: imgBuffer.length
        });
      }
    }
  }

  // 批量转换 WMF → PNG
  if (wmfToConvert.length > 0) {
    var pngResults = await convertWmfBatchToPng(wmfToConvert);
    pngResults.forEach(function(result, idx) {
      if (result.ok && result.pngBuffer) {
        var src = wmfToConvert[idx];
        images.push({
          index: src.index, fileName: src.fileName.replace(/\.wmf$/i, '.png'),
          base64: result.pngBuffer.toString('base64'), contentType: 'image/png',
          contextText: src.contextText.slice(0, 500), sizeBytes: result.sizeBytes
        });
      }
    });
  }

  return images;
}

module.exports = { extractReportImages };
