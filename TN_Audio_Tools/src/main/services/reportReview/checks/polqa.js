function checkPolqaConfiguration(wordData) {
  const issues = [];
  const evidence = [];

  const reportText = (wordData.paragraphs || []).join(' ');

  const polqaPatterns = [
    { regex: /POLQA|MOS-LQO|P\.863/i, label: 'POLQA 算法' },
    { regex: /algorithm\s+version|algo.*version|算法.*版本|版本[\s:：]*[\d.]+/i, label: '算法版本' },
    { regex: /reference.*signal|ref.*source|参考.*音源|音源[\s:：]*[^,.;。，；]*[a-zA-Z]+/i, label: '参考音源' }
  ];

  let foundPolqa = false;

  polqaPatterns.forEach((pattern) => {
    const matches = reportText.match(pattern.regex);
    if (matches) {
      evidence.push(`✓ 找到 ${pattern.label}：${matches[0]}`);
      if (pattern.label === 'POLQA 算法') {
        foundPolqa = true;
      }
    }
  });

  if (!foundPolqa) {
    issues.push({
      severity: 'review',
      message: '未在报告中明确识别到 POLQA 或 MOS-LQO 相关信息，需要人工确认'
    });
    evidence.push('POLQA 配置：未确认');
  } else {
    if (!reportText.match(/version|版本/i)) {
      issues.push({
        severity: 'warning',
        message: 'POLQA 相关信息不完整，未找到算法版本说明'
      });
      evidence.push('缺少版本信息');
    }
  }

  if (issues.length === 0) {
    evidence.push('✓ POLQA 配置信息基本完整');
    return { issues, evidence, status: 'pass' };
  }

  return { issues, evidence, status: 'review' };
}

module.exports = {
  checkPolqaConfiguration
};