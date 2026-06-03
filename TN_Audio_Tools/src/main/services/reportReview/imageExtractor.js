const fs = require('fs/promises');
const fsSync = require('fs');
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');
const JSZip = require('jszip');

var { extractVolumeLevelFromTitle } = require('./volumeLevelUtils');

var IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.wmf', '.emf']);

var MAX_IMAGE_SIZE_BYTES = 5 * 1024 * 1024;

// 排除项（优先级最高 — 明显不相关的测试）
var EXCLUDE = /Delay|Distortion|Echo|MOS|POLQA|TOSQA|Acoustic\s*Shock|Ambient\s*Noise|D-Value|ANR|Background|Single\s*Talk|Double\s*Talk|Preparation|Level\s*vs|Speech\s*quality|Stability|Sidetone|Channel|idle\s*noise/i;

// FR 参考图 — 兼容实际报告中的标题变体:
// - Sensitivity, frequency RCV/SND
// - Sensitivity, frequency character. ... MAX/NOM/MIN
// - Sensitivity frequency charact. Sending ...
var FR_REFERENCE = /Sensitivity,?\s*frequency(?:\s*(?:RCV|SND|Receiving|Sending|character(?:istic)?\.?|charact\.?))?/i;
// 旧格式兼容: "Frequency Response"、"频响"、"频率响应"
var FR_REFERENCE_LEGACY = /Frequency\s*Response|频响|频率响应/i;

// 响度 RLR（接收方向）— 新规则精确格式: "Loudness Rating RCV (RLR)"
var LOUDNESS_RLR = /Loudness\s*Rating\s*(?:RCV|Receiving)\s*(?:\(?RLR\)?)?/i;
// 响度 SLR（发送方向）— 新规则精确格式: "Loudness Rating SND (SLR)"
var LOUDNESS_SLR = /Loudness\s*Rating\s*(?:SND|Sending)\s*(?:\(?SLR\)?)?/i;

function classifyImage(contextText, fileName) {
  var combined = (contextText + ' ' + (fileName || '')).toUpperCase();
  if (EXCLUDE.test(combined)) return { category: 'excluded', volumeLevel: null };

  var volumeLevel = extractVolumeLevelFromTitle(contextText) || extractVolumeLevelFromTitle(fileName) || null;

  if (FR_REFERENCE.test(combined)) return { category: 'fr_reference', volumeLevel: volumeLevel };
  if (FR_REFERENCE_LEGACY.test(combined)) return { category: 'fr_reference', volumeLevel: volumeLevel };
  if (LOUDNESS_RLR.test(combined)) return { category: 'loudness_rlr', volumeLevel: volumeLevel };
  if (LOUDNESS_SLR.test(combined)) return { category: 'loudness_slr', volumeLevel: volumeLevel };

  // 兜底仅限“响度语义不完整但仍明显是响度图”的标题，避免把 P.863 AC 等方向图误判成 RLR/SLR。
  if (/(?:MAX(?:-\d+)?|NOM|MIN)\b/i.test(combined)
    && /Sending|Receiving|SND|RCV|TX|RX|Dir/i.test(combined)
    && /Loudness|Rating|RLR|SLR/i.test(combined)) {
    return { category: 'loudness_rlr', volumeLevel: volumeLevel };
  }
  return { category: 'excluded', volumeLevel: null };
}

function extractNearbyText(xml, index, radius) {
  var start = Math.max(0, index - radius);
  var end = Math.min(xml.length, index + radius);
  return xml.slice(start, end)
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

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
  var warnings = [];

  var buffer;
  try { buffer = await fs.readFile(reportPath); } catch (e) { console.warn('[imageExtractor] 读取文件失败:', e.message); return { images: [], warnings: ['读取报告文件失败: ' + e.message] }; }

  var zip;
  try { zip = await JSZip.loadAsync(buffer); } catch (e) { console.warn('[imageExtractor] ZIP 解析失败:', e.message); return { images: [], warnings: ['无法解析报告文件（可能为 .doc 二进制格式，请使用 .docx 格式）: ' + e.message] }; }

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
  } catch (e) { console.warn('[imageExtractor] 关系文件解析失败:', e.message); }

  // 解析 document.xml 获取图片引用上下文
  var imageContextMap = {};
  try {
    var docFile = zip.file('word/document.xml');
    if (docFile) {
      var docXml = await docFile.async('string');
      // DrawingML: <a:blip r:embed="rIdN"/>
      var dmlMatches = Array.from(docXml.matchAll(/<(?:\w+:)?blip[^>]*(?:\w+:)?embed="([^"]*)"[^>]*\/?>/gi));
      // VML: <v:imagedata r:id="rIdN"/>
      var vmlMatches = Array.from(docXml.matchAll(/<v:imagedata[^>]*r:id="([^"]*)"[^>]*\/?>/gi));

      var allRefs = dmlMatches.map(function(m) {
        return { rid: m[1], index: m.index || 0 };
      }).concat(vmlMatches.map(function(m) {
        return { rid: m[1], index: m.index || 0 };
      }));

      allRefs.forEach(function(ref) {
        var rid = ref.rid;
        var target = imageRelMap[rid];
        if (!target) return;

        var contextText = extractNearbyText(docXml, ref.index || 0, 1200);
        if (!imageContextMap[target] || contextText.length > imageContextMap[target].length) {
          imageContextMap[target] = contextText;
        }
      });
    }
  } catch (e) { console.warn('[imageExtractor] document.xml 解析失败:', e.message); }

  // 提取 word/media/ 下的图片
  var mediaFiles = zip.file(/^word\/media\/.+/i) || [];
  var images = [];
  var wmfToConvert = [];
  var totalImageCount = 0;
  var filteredCount = 0;
  var skippedExtensions = {};
  var allMediaNames = [];

  for (var i = 0; i < mediaFiles.length; i++) {
    var entry = mediaFiles[i];
    var fileName = path.basename(entry.name);
    var ext = path.extname(fileName).toLowerCase();
    allMediaNames.push(fileName);
    if (!IMAGE_EXTENSIONS.has(ext)) {
      skippedExtensions[ext] = (skippedExtensions[ext] || 0) + 1;
      continue;
    }

    var imgBuffer = await entry.async('nodebuffer');
    if (imgBuffer.length > MAX_IMAGE_SIZE_BYTES) { totalImageCount++; continue; }

    totalImageCount++;
    var target = 'media/' + fileName;
    var contextText = imageContextMap[target] || ('图片: ' + fileName);

    // 智能分类：FR 参考 / 响度 RLR / 响度 SLR / 排除
    var classification = classifyImage(contextText, fileName);
    if (classification.category !== 'excluded') {
      filteredCount++;
      if (ext === '.wmf' || ext === '.emf') {
        wmfToConvert.push({ index: i, fileName: fileName, buffer: imgBuffer, contextText: contextText, category: classification.category, volumeLevel: classification.volumeLevel });
      } else {
        images.push({
          index: i, fileName: fileName, category: classification.category,
          volumeLevel: classification.volumeLevel,
          base64: imgBuffer.toString('base64'), contentType: guessContentType(fileName),
          contextText: contextText.slice(0, 500), sizeBytes: imgBuffer.length
        });
      }
    }
  }

  if (totalImageCount > 0) {
    var catCounts = {};
    images.forEach(function(img) { var c = img.category || '?'; catCounts[c] = (catCounts[c] || 0) + 1; });
    var excludedCount = totalImageCount - filteredCount;
    warnings.push('图片分类: ' + JSON.stringify(catCounts) + (excludedCount > 0 ? ', 排除' + excludedCount + '张(Delay/Distortion等)' : ''));
  }

  // 诊断日志
  if (allMediaNames.length === 0) {
    warnings.push('Word文档中未找到 word/media/ 目录或该目录为空。请确认报告为 .docx 格式（非 .doc 二进制格式）。');
  } else {
    var skippedExtList = Object.keys(skippedExtensions);
    if (skippedExtList.length > 0) {
      warnings.push('media目录文件(' + allMediaNames.length + '个): ' + allMediaNames.slice(0, 10).join(', ') + (allMediaNames.length > 10 ? '...' : ''));
      warnings.push('跳过的文件类型: ' + skippedExtList.map(function(e) { return e + '(' + skippedExtensions[e] + '个)'; }).join(', '));
    }
  }

  // 批量转换 WMF/EMF → PNG
  if (wmfToConvert.length > 0) {
    var pngResults = await convertWmfBatchToPng(wmfToConvert);
    var wmfFailed = 0;
    pngResults.forEach(function(result, idx) {
      if (result.ok && result.pngBuffer) {
        var src = wmfToConvert[idx];
        images.push({
          index: src.index, fileName: src.fileName.replace(/\.(wmf|emf)$/i, '.png'),
          category: src.category || 'unknown',
          volumeLevel: src.volumeLevel || null,
          base64: result.pngBuffer.toString('base64'), contentType: 'image/png',
          contextText: src.contextText.slice(0, 500), sizeBytes: result.sizeBytes
        });
      } else {
        wmfFailed++;
      }
    });
    if (wmfFailed > 0) {
      warnings.push(wmfFailed + ' 张 WMF 图片转换失败，已跳过。');
    }
  }

  return { images: images, warnings: warnings };
}

module.exports = { extractReportImages };
