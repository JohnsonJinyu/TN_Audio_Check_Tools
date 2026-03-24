const { buildReviewFacts } = require('../reportFacts');

function checkPolqaConfiguration(wordData, reviewFacts = buildReviewFacts('', wordData)) {
  const issues = [];
  const evidence = [];

  const polqaFacts = reviewFacts.polqa || { algorithm: [], version: [], reference: [] };
  const foundPolqa = polqaFacts.algorithm.length > 0;

  polqaFacts.algorithm.slice(0, 3).forEach((item) => evidence.push(`✓ 找到 POLQA 算法线索：${item}`));
  polqaFacts.version.slice(0, 3).forEach((item) => evidence.push(`✓ 找到算法版本线索：${item}`));
  polqaFacts.reference.slice(0, 3).forEach((item) => evidence.push(`✓ 找到参考音源线索：${item}`));

  if (!foundPolqa) {
    issues.push({
      severity: 'review',
      message: '未在报告中明确识别到 POLQA 或 MOS-LQO 相关信息，需要人工确认'
    });
    evidence.push('POLQA 配置：未确认');
  } else {
    if (polqaFacts.version.length === 0) {
      issues.push({
        severity: 'warning',
        message: 'POLQA 相关信息不完整，未找到算法版本说明'
      });
      evidence.push('缺少版本信息');
    }

    if (polqaFacts.reference.length === 0) {
      issues.push({
        severity: 'review',
        message: 'POLQA 相关信息中未找到明确的参考音源说明，需要人工复核'
      });
      evidence.push('缺少参考音源信息');
    }
  }

  if (issues.length === 0) {
    evidence.push('✓ POLQA 配置信息基本完整');
    return { issues, evidence, status: 'pass' };
  }

  return { issues, evidence, status: issues.some((item) => item.severity === 'review') ? 'review' : 'warning' };
}

module.exports = {
  checkPolqaConfiguration
};