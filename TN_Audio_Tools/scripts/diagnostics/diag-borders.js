/**
 * Diagnostic: check all cells A3:K75 in sheet2.xml for missing/broken border styles.
 */
const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');
const XLSX = require('xlsx');

async function main() {
  const xlsxPath = process.argv[2];
  if (!xlsxPath) {
    console.error('Usage: node diag-borders.js <path-to-xlsx>');
    process.exit(1);
  }

  const buf = fs.readFileSync(xlsxPath);
  const zip = await JSZip.loadAsync(buf);

  // Resolve sheet2 path
  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const sheetRel = relsXml.match(/Id="rId(\d+)"[^>]+Target="worksheets\/sheet(\d+)\.xml"/gi);
  // Just use sheet2.xml directly
  const sheetXml = await zip.file('xl/worksheets/sheet2.xml').async('string');
  const stylesXml = await zip.file('xl/styles.xml').async('string');

  // Parse xf list
  const cellXfsContent = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] || '';
  const xfList = [...cellXfsContent.matchAll(/<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/gi)].map(m => m[0]);

  // Parse borders list
  const bordersContent = stylesXml.match(/<borders[^>]*>([\s\S]*?)<\/borders>/i)?.[1] || '';
  const borderList = [...bordersContent.matchAll(/<border\b[\s\S]*?<\/border>|<border\b[^>]*\/>/gi)].map(m => m[0]);

  // Check all cells A3:K75
  const cols = ['A','B','C','D','E','F','G','H','I','J','K'];
  const noStyle = [];
  const noBorder = [];
  const borderOK = [];

  for (let r = 3; r <= 75; r++) {
    for (const col of cols) {
      const cell = col + r;
      const cellMatch = sheetXml.match(new RegExp(`<c\\b[^>]*r="${cell}"[^>]*s="(\\d+)"[^>]*(?:\/>|>[\\s\\S]*?<\\/c>)`, 'i'));
      if (!cellMatch) {
        noStyle.push(cell);
        continue;
      }
      const sid = parseInt(cellMatch[1], 10);
      const xf = xfList[sid];
      if (!xf) {
        noBorder.push(`${cell}(xf${sid}=MISSING)`);
        continue;
      }
      const applyBorder = xf.match(/applyBorder="(\d+)"/i)?.[1];
      const borderId = xf.match(/borderId="(\d+)"/i)?.[1];
      if (applyBorder !== '1' || borderId === undefined) {
        noBorder.push(`${cell}(xf${sid},applyBorder=${applyBorder},borderId=${borderId})`);
      } else {
        const bId = parseInt(borderId, 10);
        const border = borderList[bId];
        if (!border) {
          noBorder.push(`${cell}(xf${sid},borderId=${bId}=MISSING_BORDER)`);
        } else {
          // Check if all 4 sides are defined (not just empty <left/>)
          const left = border.match(/<left\s+style="([^"]+)"/i)?.[1];
          const right = border.match(/<right\s+style="([^"]+)"/i)?.[1];
          const top = border.match(/<top\s+style="([^"]+)"/i)?.[1];
          const bottom = border.match(/<bottom\s+style="([^"]+)"/i)?.[1];
          if (!left || !right || !top || !bottom) {
            noBorder.push(`${cell}(xf${sid},borderId=${bId},sides:L=${left||'?'} R=${right||'?'} T=${top||'?'} B=${bottom||'?'})`);
          } else {
            borderOK.push(cell);
          }
        }
      }
    }
  }

  console.log(`Total cells: ${(cols.length * 73)}`);
  console.log(`✅ Border OK: ${borderOK.length}`);
  console.log(`❌ No style (no s attr): ${noStyle.length}`);
  if (noStyle.length) console.log('  ' + noStyle.join(', '));
  console.log(`❌ Missing/bad border: ${noBorder.length}`);
  if (noBorder.length) console.log('  ' + noBorder.slice(0, 80).join('\n  '));

  // Also show merge cells
  const merges = [...sheetXml.matchAll(/<mergeCell\s+ref="([^"]+)"\s*\/>/gi)].map(m => m[1]);
  console.log(`\nMerge cells: ${merges.join(', ')}`);
}

main().catch(e => { console.error(e); process.exit(1); });
