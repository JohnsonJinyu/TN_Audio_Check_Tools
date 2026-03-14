/**
 * Inspect original template's borders as baseline.
 */
const fs = require('fs');
const JSZip = require('jszip');

async function main() {
  const filePath = process.argv[2];
  const buf = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buf);
  const sheetXml = await zip.file('xl/worksheets/sheet2.xml').async('string');
  const stylesXml = await zip.file('xl/styles.xml').async('string');

  const xfs = [...stylesXml.matchAll(/<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/gi)].map(m => m[0]);
  const borders = [...stylesXml.matchAll(/<border\b[\s\S]*?<\/border>|<border\b[^>]*\/>/gi)].map(m => m[0]);

  console.log('Total borders:', borders.length);
  borders.forEach((b, i) => console.log('B' + i + ':', b.substring(0, 120)));

  // Check specific cells
  ['A3', 'B3', 'C3', 'D3', 'K3', 'K6', 'K7', 'A5', 'A6', 'B7'].forEach(cell => {
    const re = new RegExp('<c\\b[^>]*r="' + cell + '"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)', 'i');
    const m = sheetXml.match(re);
    const styleId = m?.[0]?.match(/s="(\d+)"/)?.[1];
    const xf = styleId !== undefined ? xfs[parseInt(styleId, 10)] : null;
    const borderId = xf?.match(/borderId="(\d+)"/)?.[1];
    const bo = borderId !== undefined ? borders[parseInt(borderId, 10)] : null;
    console.log(cell + ': node=' + (m?.[0]?.substring(0, 80) || 'MISSING'));
    if (xf) console.log('  xf=' + xf.substring(0, 100));
    if (bo) console.log('  border=' + bo.substring(0, 120));
  });

  const merges = [...sheetXml.matchAll(/<mergeCell\s+ref="([^"]+)"\s*\/>/gi)].map(m => m[1]);
  console.log('\nMerges (first 15):', merges.slice(0, 15).join(', '));
  console.log('Total merges:', merges.length);
}

main().catch(console.error);
