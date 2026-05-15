export function normalizeReportPaths(filePaths) {
  if (!Array.isArray(filePaths)) {
    return [];
  }

  const supportedExtensions = new Set(['.doc', '.docx', '.xlsx']);
  const normalized = filePaths
    .map((filePath) => String(filePath || '').trim())
    .filter(Boolean)
    .filter((filePath) => {
      const lowerCasePath = filePath.toLowerCase();
      return Array.from(supportedExtensions).some((extension) => lowerCasePath.endsWith(extension));
    });

  return Array.from(new Set(normalized));
}

export function getReportName(filePath) {
  return String(filePath || '').split('\\').pop() || filePath;
}

export function getFileBaseName(filePath) {
  const name = getReportName(filePath);
  return name.replace(/\.(docx?|xlsx)$/i, '');
}

export function isDocFile(filePath) {
  const lower = (filePath || '').toLowerCase();
  return lower.endsWith('.doc') && !lower.endsWith('.docx');
}

export function getProgressLabel(filePath, prefix, baseName) {
  const label = baseName || getReportName(filePath);
  return isDocFile(filePath) ? `正在转换Word格式: ${label}` : `${prefix}: ${label}`;
}

export function detectFilePairs(filePaths) {
  const byBase = new Map();

  filePaths.forEach((filePath) => {
    const lowerPath = filePath.toLowerCase();
    let ext = '';
    if (lowerPath.endsWith('.xlsx')) ext = '.xlsx';
    else if (lowerPath.endsWith('.docx')) ext = '.docx';
    else if (lowerPath.endsWith('.doc')) ext = '.doc';

    const baseName = filePath.slice(0, filePath.length - ext.length);
    const normalizedKey = baseName.toLowerCase();

    if (!byBase.has(normalizedKey)) {
      byBase.set(normalizedKey, { baseName: getFileBaseName(filePath), docx: null, xlsx: null });
    }

    const group = byBase.get(normalizedKey);
    if (ext === '.docx' || ext === '.doc') {
      group.docx = group.docx || filePath;
    } else if (ext === '.xlsx') {
      group.xlsx = filePath;
    }
  });

  const pairs = [];
  const solo = [];

  byBase.forEach((group) => {
    if (group.docx && group.xlsx) {
      pairs.push({ baseName: group.baseName, docx: group.docx, xlsx: group.xlsx });
    } else {
      if (group.docx) solo.push(group.docx);
      if (group.xlsx) solo.push(group.xlsx);
    }
  });

  return { pairs, solo };
}
