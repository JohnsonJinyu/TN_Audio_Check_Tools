const fs = require('fs/promises');
const JSZip = require('jszip');
const cheerio = require('cheerio');

function getLocalName(node) {
  return String(node?.name || '').split(':').pop();
}

function normalizeCellText(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractNodeText($, node) {
  const parts = [];

  $(node).contents().each((_, child) => {
    const localName = getLocalName(child);

    if (child.type === 'text') {
      const text = normalizeCellText(child.data);
      if (text) {
        parts.push(text);
      }
      return;
    }

    if (localName === 't') {
      const text = normalizeCellText($(child).text());
      if (text) {
        parts.push(text);
      }
      return;
    }

    if (localName === 'tab') {
      parts.push('\t');
      return;
    }

    if (localName === 'br' || localName === 'cr') {
      parts.push('\n');
      return;
    }

    if (child.type === 'tag') {
      const nestedText = extractNodeText($, child);
      if (nestedText) {
        parts.push(nestedText);
      }
    }
  });

  return parts
    .join(' ')
    .replace(/\s*\t\s*/g, '\t')
    .replace(/\s*\n\s*/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .trim();
}

function extractParagraphText($, paragraphNode) {
  const text = extractNodeText($, paragraphNode);
  return normalizeCellText(text.replace(/\n+/g, ' '));
}

function extractCellText($, cellNode) {
  const paragraphTexts = $(cellNode)
    .children()
    .toArray()
    .filter((child) => getLocalName(child) === 'p')
    .map((paragraphNode) => extractParagraphText($, paragraphNode))
    .filter(Boolean);

  if (paragraphTexts.length > 0) {
    return paragraphTexts.join(' ');
  }

  return normalizeCellText(extractNodeText($, cellNode));
}

function parseDocumentXml(documentXml) {
  const $ = cheerio.load(documentXml, {
    xmlMode: true,
    decodeEntities: true
  });

  const body = $('w\\:body, body').first();
  const lines = [];
  const tables = [];
  let tableIndex = 0;

  body.children().each((_, child) => {
    const localName = getLocalName(child);

    if (localName === 'p') {
      const text = extractParagraphText($, child);
      if (text) {
        lines.push(text);
      }
      return;
    }

    if (localName !== 'tbl') {
      return;
    }

    const rows = $(child)
      .children()
      .toArray()
      .filter((rowNode) => getLocalName(rowNode) === 'tr')
      .map((rowNode) => $(rowNode)
        .children()
        .toArray()
        .filter((cellNode) => getLocalName(cellNode) === 'tc')
        .map((cellNode) => extractCellText($, cellNode))
        .filter(Boolean))
      .filter((row) => row.length > 0);

    if (rows.length === 0) {
      return;
    }

    tables.push({ tableIndex, rows });
    tableIndex += 1;

    rows.forEach((row) => {
      const text = row.join(' | ').trim();
      if (text) {
        lines.push(text);
      }
    });
  });

  return {
    lines,
    tables
  };
}

async function parseDocxStructuredData(reportPath) {
  const buffer = await fs.readFile(reportPath);
  const zip = await JSZip.loadAsync(buffer);
  const documentEntry = zip.file('word/document.xml');

  if (!documentEntry) {
    return { lines: [], tables: [] };
  }

  const documentXml = await documentEntry.async('string');
  return parseDocumentXml(documentXml);
}

module.exports = {
  parseDocxStructuredData
};