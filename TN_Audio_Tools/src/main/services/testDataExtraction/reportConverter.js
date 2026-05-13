const fs = require('fs/promises');
const os = require('os');
const path = require('path');
const { execSync } = require('child_process');
const JSZip = require('jszip');

function createReportConverter({ wordExtractor }) {
  if (!wordExtractor || typeof wordExtractor.extract !== 'function') {
    throw new Error('reportConverter 需要注入可用的 wordExtractor 实例。');
  }

  function escapeXml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function normalizeDocText(rawText) {
    return String(rawText || '')
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/\u0000/g, '')
      .trim();
  }

  function buildDocxDocumentXml(rawText) {
    const lines = normalizeDocText(rawText)
      .split('\n')
      .map((line) => line.trimEnd());

    const paragraphXml = lines
      .map((line) => {
        if (!line) {
          return '<w:p/>';
        }

        return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`;
      })
      .join('');

    return [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">',
      `<w:body>${paragraphXml || '<w:p/>'}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body>`,
      '</w:document>'
    ].join('');
  }

  async function buildDocxFromText(rawText, outputPath) {
    const zip = new JSZip();

    zip.file('[Content_Types].xml', [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      '</Types>'
    ].join(''));

    zip.file('_rels/.rels', [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
      '</Relationships>'
    ].join(''));

    zip.file('docProps/app.xml', [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
      '<Application>TN Audio Toolkit</Application>',
      '</Properties>'
    ].join(''));

    const now = new Date().toISOString();
    zip.file('docProps/core.xml', [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
      '<dc:title>Converted from DOC</dc:title>',
      '<dc:creator>TN Audio Toolkit</dc:creator>',
      '<cp:lastModifiedBy>TN Audio Toolkit</cp:lastModifiedBy>',
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`,
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`,
      '</cp:coreProperties>'
    ].join(''));

    zip.file('word/document.xml', buildDocxDocumentXml(rawText));

    const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    await fs.writeFile(outputPath, buffer);
  }

  function isComProgIdAvailable(progId) {
    try {
      execSync(
        `powershell -NoProfile -Command "try { $x = New-Object -ComObject '${progId}'; $x.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x); Write-Host 'OK' } catch { exit 1 }"`,
        { stdio: 'pipe', timeout: 8000 }
      );
      return true;
    } catch (_) {
      return false;
    }
  }

  async function convertDocViaCom(progId, inputPath, outputPath) {
    const psScript = [
      `$app = New-Object -ComObject '${progId}'`,
      `$app.Visible = $false`,
      `$app.DisplayAlerts = 0`,
      `$doc = $app.Documents.Open('${inputPath.replace(/\\/g, '\\\\')}')`,
      `$doc.SaveAs2('${outputPath.replace(/\\/g, '\\\\')}', 16)`,
      `$doc.Close()`,
      `$app.Quit()`,
      `[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)`,
      `[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)`,
    ].join('; ');

    execSync(`powershell -NoProfile -Command "${psScript}"`, {
      stdio: 'pipe',
      timeout: 30000,
      windowsHide: true
    });

    try {
      await fs.access(outputPath);
      return true;
    } catch (_) {
      return false;
    }
  }

  function isLibreOfficeAvailable() {
    try {
      const candidates = ['libreoffice', 'soffice'];
      for (const cmd of candidates) {
        try {
          execSync(`"${cmd}" --version`, { stdio: 'pipe', timeout: 5000 });
          return cmd;
        } catch (_) { /* try next */ }
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  async function convertDocViaLibreOffice(reportPath, tempDir) {
    const libreOfficeCmd = isLibreOfficeAvailable();
    if (!libreOfficeCmd) return null;

    const inputName = path.parse(reportPath).base;
    execSync(
      `"${libreOfficeCmd}" --headless --convert-to docx --outdir "${tempDir}" "${reportPath}"`,
      { stdio: 'pipe', timeout: 30000 }
    );

    const expectedOutput = path.join(tempDir, inputName.replace(/\.doc$/i, '.docx'));
    try {
      await fs.access(expectedOutput);
      return { tempDir, convertedPath: expectedOutput };
    } catch (_) {
      return null;
    }
  }

  async function extractDocText(reportPath) {
    const extracted = await wordExtractor.extract(reportPath);
    return [
      extracted.getHeaders?.() || '',
      extracted.getBody?.() || '',
      extracted.getFootnotes?.() || '',
      extracted.getEndnotes?.() || '',
      extracted.getTextboxes?.() || ''
    ].filter(Boolean).join('\n');
  }

  async function convertDocToTemporaryDocx(reportPath) {
    const outputDir = path.dirname(reportPath);
    const convertedPath = path.join(outputDir, `${path.parse(reportPath).name}.docx`);

    // ① Word.Application COM（最高保真度）
    if (isComProgIdAvailable('Word.Application')) {
      try {
        const ok = await convertDocViaCom('Word.Application', reportPath, convertedPath);
        if (ok) return { tempDir: outputDir, convertedPath };
      } catch (_) { /* 静默回退 */ }
    }

    // ② KWPS.Application COM（WPS，已预装）
    if (isComProgIdAvailable('KWPS.Application')) {
      try {
        const ok = await convertDocViaCom('KWPS.Application', reportPath, convertedPath);
        if (ok) return { tempDir: outputDir, convertedPath };
      } catch (_) { /* 静默回退 */ }
    }

    // ③ LibreOffice CLI
    try {
      const loResult = await convertDocViaLibreOffice(reportPath, outputDir);
      if (loResult) return loResult;
    } catch (_) { /* 静默回退 */ }

    // ④ word-extractor 文本重建（永久保底）
    const rawText = normalizeDocText(await extractDocText(reportPath));
    if (!rawText) {
      return null;
    }

    try {
      await buildDocxFromText(rawText, convertedPath);
      return { tempDir: outputDir, convertedPath };
    } catch (error) {
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
