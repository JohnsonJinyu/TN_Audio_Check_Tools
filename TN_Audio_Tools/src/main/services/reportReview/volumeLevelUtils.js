/**
 * 测试等级体系工具模块
 *
 * 定义8级规范等级：
 *   MAX(0) > MAX-1(1) > MAX-2(2) > MAX-3(NOM)(3)
 *   > MAX-4(4) > MAX-5(5) > MAX-6(6) > MAX-7(MIN)(7)
 */

var CANONICAL_LEVELS = [
  'MAX',
  'MAX-1',
  'MAX-2',
  'MAX-3(NOM)',
  'MAX-4',
  'MAX-5',
  'MAX-6',
  'MAX-7(MIN)',
];

var ORDINAL_MAP = {};
CANONICAL_LEVELS.forEach(function (level, idx) {
  ORDINAL_MAP[level.toUpperCase()] = idx;
});

var ALIAS_MAP = {
  'NOM': 'MAX-3(NOM)',
  'MAX-3': 'MAX-3(NOM)',
  'MAX3': 'MAX-3(NOM)',
  'MAX-3NOM': 'MAX-3(NOM)',
  'MAX3(NOM)': 'MAX-3(NOM)',
  'MIN': 'MAX-7(MIN)',
  'MAX-7': 'MAX-7(MIN)',
  'MAX7': 'MAX-7(MIN)',
  'MAX-7MIN': 'MAX-7(MIN)',
  'MAX7(MIN)': 'MAX-7(MIN)',
  // 数字格式（ACQUA xlsx VOLUME_CTRL列可能使用数字步进值）
  '0': 'MAX',
  '-0': 'MAX',
  '-1': 'MAX-1',
  '1': 'MAX-1',
  '-2': 'MAX-2',
  '2': 'MAX-2',
  '-3': 'MAX-3(NOM)',
  '3': 'MAX-3(NOM)',
  '-4': 'MAX-4',
  '4': 'MAX-4',
  '-5': 'MAX-5',
  '5': 'MAX-5',
  '-6': 'MAX-6',
  '6': 'MAX-6',
  '-7': 'MAX-7(MIN)',
  '7': 'MAX-7(MIN)',
};

/**
 * 将原始等级字符串规范化为标准格式
 * "NOM" → "MAX-3(NOM)", "MAX-3" → "MAX-3(NOM)", "min" → "MAX-7(MIN)"
 * 无法识别时返回原始值（去空白）
 */
function normalizeVolumeLevel(raw) {
  if (!raw && raw !== 0) return null;
  var s = String(raw).trim();
  if (!s) return null;
  var upper = s.toUpperCase();
  if (ALIAS_MAP[upper]) return ALIAS_MAP[upper];
  // 直接匹配标准名称（大小写不敏感）
  for (var i = 0; i < CANONICAL_LEVELS.length; i++) {
    if (upper === CANONICAL_LEVELS[i].toUpperCase()) return CANONICAL_LEVELS[i];
  }
  return s;
}

function getVolumeLevelOrdinal(level) {
  if (!level) return 999;
  var upper = String(level).toUpperCase();
  if (ORDINAL_MAP[upper] !== undefined) return ORDINAL_MAP[upper];
  // 尝试规范化后再查
  var normalized = normalizeVolumeLevel(level);
  if (normalized && ORDINAL_MAP[normalized.toUpperCase()] !== undefined) {
    return ORDINAL_MAP[normalized.toUpperCase()];
  }
  return 999;
}

function compareVolumeLevel(a, b) {
  var ordA = getVolumeLevelOrdinal(a);
  var ordB = getVolumeLevelOrdinal(b);
  if (ordA < ordB) return -1;
  if (ordA > ordB) return 1;
  return 0;
}

function isNomLevel(level) {
  var normalized = normalizeVolumeLevel(level);
  return normalized === 'MAX-3(NOM)';
}

function isMinLevel(level) {
  var normalized = normalizeVolumeLevel(level);
  return normalized === 'MAX-7(MIN)';
}

/**
 * 从标题文本中提取测试等级
 * 匹配模式: "BIN MAX", "BIN MAX-3(NOM)", "MAX HHNB" 等
 */
var VOLUME_LEVEL_TITLE_RE = /\b(?:BIN\s+)?(MAX(?:[-\s]*\d+)?(?:\(NOM\))?|NOM|MAX[-\s]*\d+\s*\(MIN\)|MIN)(?:\s+HHNB|\s+WB|\s+NB|\s+SWB|\s+SSW|\s+FB)?(?:\s|$)/i;

function extractVolumeLevelFromTitle(titleText) {
  if (!titleText) return null;
  var m = titleText.match(VOLUME_LEVEL_TITLE_RE);
  if (!m) return null;
  return normalizeVolumeLevel(m[1]);
}

module.exports = {
  CANONICAL_LEVELS,
  normalizeVolumeLevel,
  getVolumeLevelOrdinal,
  compareVolumeLevel,
  isNomLevel,
  isMinLevel,
  extractVolumeLevelFromTitle,
};
