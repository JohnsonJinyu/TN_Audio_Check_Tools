import React, { useEffect, useRef, useState } from 'react';
import { Alert, Button, Card, Col, Collapse, Divider, Progress, Row, Space, Spin, Tag } from 'antd';
import { CheckOutlined, CloseOutlined, ExclamationOutlined, ExperimentOutlined } from '@ant-design/icons';
import MonotonicityChart from './MonotonicityChart';
import LoudnessFrComparisonCard from './components/LoudnessFrComparisonCard';

function getStatusColor(status) {
  const colorMap = {
    pass: 'green',
    warning: 'orange',
    review: 'blue',
    error: 'red'
  };

  return colorMap[status] || 'default';
}

function getStatusIcon(status) {
  if (status === 'pass') return <CheckOutlined />;
  if (status === 'error') return <CloseOutlined />;
  return <ExclamationOutlined />;
}

function getStatusText(status) {
  const textMap = {
    pass: '通过',
    warning: '有警告',
    review: '需人工复核',
    error: '有错误'
  };

  return textMap[status] || status;
}

function groupComparisonCards(cards) {
  return (cards || []).reduce(function(acc, card) {
    var key = card.direction || 'unknown';
    if (!acc[key]) acc[key] = [];
    acc[key].push(card);
    return acc;
  }, {});
}

function getComparisonGroupTitle(direction) {
  if (direction === 'RCV') return '接收方向同等级曲线对比（RLR vs 接收频响）';
  if (direction === 'SND') return '发送方向曲线对比（SLR vs 发送频响）';
  return '方向待确认的同等级曲线对比';
}

function formatTrendSummary(summaryText) {
  var text = String(summaryText || '').trim();
  if (!text) return [];

  return text
    .split(/\s*[;；]\s*(?=\[[A-Z]{3}\]|$)/)
    .map(function(item) { return String(item || '').trim(); })
    .filter(Boolean)
    .map(function(item, index) {
      var match = item.match(/^(\[[A-Z]{3}\]\s*\[[^\]]+\])\s*(.*)$/);
      return {
        key: 'trend-summary-' + index,
        label: match ? match[1] : '',
        content: match ? match[2] : item,
      };
    });
}

export default function ReviewResultContent({ resultData }) {
  const [aiAnalyzing, setAiAnalyzing] = useState({});
  const [aiResults, setAiResults] = useState({});
  const [aiProgress, setAiProgress] = useState({});
  const progressUnsubscribeRef = useRef({});

  useEffect(() => {
    return () => {
      Object.values(progressUnsubscribeRef.current).forEach(fn => { if (typeof fn === 'function') fn(); });
    };
  }, []);

  if (!resultData) {
    return null;
  }

  const { report, reviewResult } = resultData;
  const reportPath = report?.reportPath || report?.path || '';

  const handleAiReanalyze = async (sectionKey) => {
    setAiAnalyzing(prev => ({ ...prev, [sectionKey]: true }));
    setAiResults(prev => ({ ...prev, [sectionKey]: null }));
    setAiProgress(prev => ({ ...prev, [sectionKey]: { current: 0, total: 0, fileName: '' } }));

    try {
      if (typeof progressUnsubscribeRef.current[sectionKey] === 'function') {
        progressUnsubscribeRef.current[sectionKey]();
      }
    } catch (_) {}

    var unsub = null;
    try {
      if (window.electron && window.electron.reportReview && window.electron.reportReview.onChartAnalysisProgress) {
        unsub = window.electron.reportReview.onChartAnalysisProgress(function(data) {
          setAiProgress(prev => ({ ...prev, [sectionKey]: data }));
        });
        progressUnsubscribeRef.current[sectionKey] = unsub;
      }
    } catch (_) {}

    try {
      const res = await window.electron.reportReview.analyzeChartImages({ reportPath });
      setAiResults(prev => ({ ...prev, [sectionKey]: res }));
    } catch (e) {
      setAiResults(prev => ({ ...prev, [sectionKey]: { status: 'error', issues: [{ severity: 'error', message: 'AI分析失败: ' + (e.message || '未知错误') }] } }));
    } finally {
      setAiAnalyzing(prev => ({ ...prev, [sectionKey]: false }));
      try { if (typeof unsub === 'function') unsub(); } catch (_) {}
      try { delete progressUnsubscribeRef.current[sectionKey]; } catch (_) {}
    }
  };
  if (!reviewResult) {
    return <Alert type="error" message="审查结果数据格式异常，无法显示" />;
  }

  const summary = reviewResult.summary || {
    passedChecks: 0,
    warningChecks: 0,
    reviewChecks: 0,
    errorChecks: 0
  };
  const CROSS_REPORT_SECTION_KEYS = new Set(['contentSameCodecDiffNetwork', 'contentSameNetworkDiffCodec']);
  let sections = Array.isArray(report?.sections) ? report.sections : [];
  var trendSummaryItems = formatTrendSummary(
    sections.find(function(section) { return section.key === 'contentLoudnessFRTrend'; })?.conclusion
  );

  if (resultData.hideCrossReportSections) {
    sections = sections.filter(function(s) { return !CROSS_REPORT_SECTION_KEYS.has(s.key); });
  }

  if (sections.length === 0) {
    return (
      <Alert
        type="warning"
        showIcon
        message="当前历史记录缺少可展示的详细检查项"
        description="这条记录可能来自旧版本数据，或审查结果未完整保存。"
      />
    );
  }

  return (
    <div>
      <Row gutter={[16, 16]} style={{ marginBottom: 24 }}>
        <Col xs={24}>
          <Alert
            message={`整体评估：${getStatusText(reviewResult.overallStatus)}`}
            type={reviewResult.overallStatus === 'pass' ? 'success' : (
              reviewResult.overallStatus === 'error' ? 'error' : (
                reviewResult.overallStatus === 'warning' ? 'warning' : 'info'
              )
            )}
            icon={getStatusIcon(reviewResult.overallStatus)}
            showIcon
            style={{ marginBottom: 16 }}
          />
        </Col>

        <Col xs={24} sm={6}>
          <Card size="small" title="审查统计">
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: '8px 0', color: '#52c41a', fontSize: 16, fontWeight: 'bold' }}>{summary.passedChecks}</p>
              <p style={{ margin: 0, fontSize: 12, color: 'var(--text-light)' }}>通过项</p>
            </div>
          </Card>
        </Col>

        <Col xs={24} sm={6}>
          <Card size="small">
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: '8px 0', color: '#faad14', fontSize: 16, fontWeight: 'bold' }}>{summary.warningChecks}</p>
              <p style={{ margin: 0, fontSize: 12, color: 'var(--text-light)' }}>警告项</p>
            </div>
          </Card>
        </Col>

        <Col xs={24} sm={6}>
          <Card size="small">
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: '8px 0', color: '#1677ff', fontSize: 16, fontWeight: 'bold' }}>{summary.reviewChecks}</p>
              <p style={{ margin: 0, fontSize: 12, color: 'var(--text-light)' }}>需复核</p>
            </div>
          </Card>
        </Col>

        <Col xs={24} sm={6}>
          <Card size="small">
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: '8px 0', color: '#ff4d4f', fontSize: 16, fontWeight: 'bold' }}>{summary.errorChecks}</p>
              <p style={{ margin: 0, fontSize: 12, color: 'var(--text-light)' }}>错误项</p>
            </div>
          </Card>
        </Col>
      </Row>

      <Divider />

      <h3 style={{ marginBottom: 16 }}>详细检查结果</h3>
      <Collapse
        items={sections.map((section, sectionIndex) => ({
          key: section.key || `section-${sectionIndex}`,
          label: (
            <Space>
              <Tag color={getStatusColor(section.status)}>{getStatusText(section.status || 'review')}</Tag>
              <span style={{ fontWeight: 500 }}>{section.title || '未命名检查项'}</span>
            </Space>
          ),
          children: (
            <div>
              <p style={{ marginBottom: 12, color: 'var(--text-light)' }}>{section.description || '无详细说明'}</p>

              {section.key === 'contentLoudnessFRTrend' && Array.isArray(section.comparisonCards) && section.comparisonCards.length > 0 && (
                <div style={{ marginBottom: 20 }}>
                  <Card
                    size="small"
                    style={{
                      marginBottom: 16,
                      borderRadius: 24,
                      border: '1px solid #e7e5e4',
                      background: '#fafaf9',
                    }}
                    bodyStyle={{ padding: 20 }}
                  >
                    <div style={{ fontSize: 12, color: '#78716c', textTransform: 'uppercase', letterSpacing: 0.8, marginBottom: 8 }}>
                      趋势摘要
                    </div>
                    {trendSummaryItems.length > 0 ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                        {trendSummaryItems.map(function(item) {
                          return (
                            <div
                              key={item.key}
                              style={{
                                padding: '12px 14px',
                                borderRadius: 16,
                                background: '#ffffff',
                                border: '1px solid #e7e5e4'
                              }}
                            >
                              {item.label ? (
                                <div style={{ marginBottom: 6, fontSize: 12, fontWeight: 600, color: '#57534e', letterSpacing: 0.2 }}>
                                  {item.label}
                                </div>
                              ) : null}
                              <div style={{ fontSize: 15, color: '#1c1917', lineHeight: 1.85, whiteSpace: 'pre-wrap' }}>
                                {item.content}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div style={{ fontSize: 15, color: '#1c1917', lineHeight: 1.85, whiteSpace: 'pre-wrap' }}>
                        {section.conclusion || '已按方向与等级生成曲线对比卡片，可直接查看 FR 与响度图的对应关系。'}
                      </div>
                    )}
                  </Card>

                  {Object.entries(groupComparisonCards(section.comparisonCards)).map(function(entry) {
                    var direction = entry[0];
                    var cards = entry[1];
                    var shouldHideLevel = direction === 'SND' && cards.length === 1;
                    return (
                      <div key={direction} style={{ marginBottom: 20 }}>
                        <div style={{ marginBottom: 12, display: 'flex', alignItems: 'center', gap: 8 }}>
                          <Tag color="blue" style={{ borderRadius: 999, margin: 0 }}>{direction}</Tag>
                          <span style={{ fontSize: 15, fontWeight: 500, color: '#292524' }}>
                            {getComparisonGroupTitle(direction)}
                          </span>
                        </div>
                        <Row gutter={[16, 16]}>
                          {cards.map(function(card, idx) {
                            return (
                              <Col xs={24} key={direction + '-' + card.level + '-' + idx}>
                                <LoudnessFrComparisonCard card={card} hideLevel={shouldHideLevel} />
                              </Col>
                            );
                          })}
                        </Row>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* 判定结论 */}
              {section.conclusion && section.key !== 'contentLoudnessFRTrend' && (
                <div style={{ marginBottom: 12 }}>
                  <h4 style={{ marginBottom: 8 }}>判定结论：</h4>
                  <p style={{ margin: 0, color: 'var(--text-light)', fontSize: 12 }}>{section.conclusion}</p>
                </div>
              )}

              {/* 逐项检查清单 — 显示每一项的对错状态 */}
              {Array.isArray(section.checklist) && section.checklist.length > 0 && (
                <div style={{ marginBottom: 16 }}>
                  {section.key === 'contentLoudnessFRTrend' ? (
                    <Collapse
                      ghost
                      size="small"
                      items={[{
                        key: 'trend-checklist',
                        label: (
                          <span style={{ fontSize: 12, color: 'var(--text-light)' }}>
                            查看结构化明细（{section.checklist.length} 项）
                          </span>
                        ),
                        children: (
                          <div style={{ overflowX: 'auto' }}>
                            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                              <thead>
                                <tr style={{ borderBottom: '2px solid var(--border-color)', backgroundColor: 'var(--surface-muted)' }}>
                                  {section.checklist[0]?.direction != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>方向</th>}
                                  {section.checklist[0]?.volumeLevel != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>等级</th>}
                                  {section.checklist[0]?.role != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>类型</th>}
                                  {section.checklist[0]?.imageIndex != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>图#</th>}
                                  <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>结果</th>
                                  <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>详情</th>
                                </tr>
                              </thead>
                              <tbody>
                                {section.checklist.map((item, idx) => (
                                  <tr key={idx} style={{ borderBottom: '1px solid var(--border-color)', backgroundColor: item.status === 'error' ? '#fff2f0' : item.status === 'warning' ? '#fffbe6' : 'transparent' }}>
                                    {item.direction != null && <td style={{ padding: '6px 8px' }}><Tag color="blue" style={{ margin: 0 }}>{item.direction}</Tag></td>}
                                    {item.volumeLevel != null && <td style={{ padding: '6px 8px' }}>{item.volumeLevel || '-'}</td>}
                                    {item.role != null && <td style={{ padding: '6px 8px' }}>{item.role === 'reference' ? 'FR基准' : item.role === 'loudness_rlr' ? 'RLR响度' : item.role === 'loudness_slr' ? 'SLR响度' : item.role || '-'}</td>}
                                    {item.imageIndex != null && <td style={{ padding: '6px 8px' }}>#{item.imageIndex}</td>}
                                    <td style={{ padding: '6px 8px' }}>
                                      <Tag color={item.status === 'pass' ? 'success' : item.status === 'warning' ? 'orange' : item.status === 'error' ? 'red' : 'default'} style={{ margin: 0 }}>
                                        {item.status === 'pass' ? '通过' : item.status === 'warning' ? '警告' : item.status === 'error' ? '错误' : item.status}
                                      </Tag>
                                    </td>
                                    <td style={{ padding: '6px 8px', maxWidth: 350, fontSize: 11 }}>
                                      {item.detail || '-'}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )
                      }]}
                    />
                  ) : (
                    <>
                      <h4 style={{ marginBottom: 8 }}>
                        逐项检查清单（{section.checklist.length} 项）：
                        <span style={{ fontSize: 12, fontWeight: 'normal', color: 'var(--text-light)', marginLeft: 12 }}>
                          通过 {section.checklist.filter(function(c) { return c.status === 'pass'; }).length}
                          {section.checklist.filter(function(c) { return c.status === 'warning'; }).length > 0 && ' | 警告 ' + section.checklist.filter(function(c) { return c.status === 'warning'; }).length}
                          {section.checklist.filter(function(c) { return c.status === 'error'; }).length > 0 && ' | 错误 ' + section.checklist.filter(function(c) { return c.status === 'error'; }).length}
                        </span>
                      </h4>
                      <div style={{ overflowX: 'auto' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                          <thead>
                            <tr style={{ borderBottom: '2px solid var(--border-color)', backgroundColor: 'var(--surface-muted)' }}>
                              {section.checklist[0]?.direction != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>方向</th>}
                              {section.checklist[0]?.volumeLevel != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>等级</th>}
                              {section.checklist[0]?.role != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>类型</th>}
                              {section.checklist[0]?.fromLevel != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>从</th>}
                              {section.checklist[0]?.toLevel != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>到</th>}
                              {section.checklist[0]?.imageIndex != null && <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>图#</th>}
                              <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>结果</th>
                              <th style={{ padding: '6px 8px', textAlign: 'left', whiteSpace: 'nowrap' }}>详情</th>
                            </tr>
                          </thead>
                          <tbody>
                            {section.checklist.map((item, idx) => (
                              <tr key={idx} style={{ borderBottom: '1px solid var(--border-color)', backgroundColor: item.status === 'error' ? '#fff2f0' : item.status === 'warning' ? '#fffbe6' : 'transparent' }}>
                                {item.direction != null && <td style={{ padding: '6px 8px' }}><Tag color="blue" style={{ margin: 0 }}>{item.direction}</Tag></td>}
                                {item.volumeLevel != null && <td style={{ padding: '6px 8px' }}>{item.volumeLevel || '-'}</td>}
                                {item.role != null && <td style={{ padding: '6px 8px' }}>{item.role === 'reference' ? 'FR基准' : item.role === 'loudness_rlr' ? 'RLR响度' : item.role === 'loudness_slr' ? 'SLR响度' : item.role || '-'}</td>}
                                {item.fromLevel != null && <td style={{ padding: '6px 8px' }}>{item.fromLevel}</td>}
                                {item.toLevel != null && <td style={{ padding: '6px 8px' }}>{item.toLevel}</td>}
                                {item.imageIndex != null && <td style={{ padding: '6px 8px' }}>#{item.imageIndex}</td>}
                                <td style={{ padding: '6px 8px' }}>
                                  <Tag color={item.status === 'pass' ? 'success' : item.status === 'warning' ? 'orange' : item.status === 'error' ? 'red' : 'default'} style={{ margin: 0 }}>
                                    {item.status === 'pass' ? '通过' : item.status === 'warning' ? '警告' : item.status === 'error' ? '错误' : item.status}
                                  </Tag>
                                </td>
                                <td style={{ padding: '6px 8px', maxWidth: 350, fontSize: 11 }}>
                                  {item.detail || '-'}
                                  {(item.expectedBehavior || item.actualBehavior) && (
                                    <div style={{ marginTop: 2, color: 'var(--text-light)' }}>
                                      {item.expectedBehavior && <div>期望: {item.expectedBehavior}</div>}
                                      {item.actualBehavior && <div>实际: {item.actualBehavior}</div>}
                                    </div>
                                  )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </>
                  )}
                </div>
              )}

              {/* 响度-等级趋势折线图 */}
              {section.chartData && Object.keys(section.chartData).length > 0 && (
                <MonotonicityChart data={section.chartData} />
              )}

              {/* 异常问题项 — 仅当有error/warning */}
              {Array.isArray(section.issues) && section.issues.length > 0 && (
                <div style={{ marginBottom: 16 }}>
                  <h4 style={{ marginBottom: 8 }}>异常详情：</h4>
                  {section.issues.filter(function(i) { return i.severity === 'error'; }).map((issue, idx) => (
                    <Alert
                      key={'err-' + idx}
                      message={
                        <Space wrap size={[4, 4]}>
                          {issue.meta?.direction && <Tag color="blue">{issue.meta.direction}</Tag>}
                          {issue.meta?.volumeLevel && <Tag color="geekblue">{issue.meta.volumeLevel}</Tag>}
                          {issue.meta?.frequencyRange && <Tag color="purple">{issue.meta.frequencyRange}</Tag>}
                          {issue.meta?.imageIndex != null && <Tag>图#{issue.meta.imageIndex}</Tag>}
                          <span>{issue.message}</span>
                        </Space>
                      }
                      description={issue.meta?.expectedBehavior && (
                        <div style={{ fontSize: 12 }}>
                          <div>期望: {issue.meta.expectedBehavior}</div>
                          <div>实际: {issue.meta.actualBehavior}</div>
                        </div>
                      )}
                      type="error"
                      showIcon
                      style={{ marginBottom: 8 }}
                    />
                  ))}
                  {section.issues.filter(function(i) { return i.severity === 'warning'; }).map((issue, idx) => (
                    <Alert
                      key={'warn-' + idx}
                      message={
                        <Space wrap size={[4, 4]}>
                          {issue.meta?.direction && <Tag color="blue">{issue.meta.direction}</Tag>}
                          {issue.meta?.volumeLevel && <Tag color="geekblue">{issue.meta.volumeLevel}</Tag>}
                          {issue.meta?.frequencyRange && <Tag color="purple">{issue.meta.frequencyRange}</Tag>}
                          <span>{issue.message}</span>
                        </Space>
                      }
                      type="warning"
                      showIcon
                      style={{ marginBottom: 8 }}
                    />
                  ))}
                </div>
              )}

              {/* 诊断证据 */}
              {Array.isArray(section.evidence) && section.evidence.length > 0 && (
                <div style={{ marginBottom: 12 }}>
                  <h4 style={{ marginBottom: 8 }}>诊断依据：</h4>
                  <ul style={{ marginBottom: 0, paddingLeft: 20 }}>
                    {section.evidence.map((item, idx) => (
                      <li key={idx} style={{ marginBottom: 4, color: 'var(--text-light)', fontSize: 12 }}>{item}</li>
                    ))}
                  </ul>
                </div>
              )}

              {/* 分析日志 — 可折叠，默认关闭 */}
              {Array.isArray(section.logs) && section.logs.length > 0 && (
                <Collapse
                  ghost
                  size="small"
                  items={[{
                    key: 'logs',
                    label: <span style={{ fontSize: 11, color: 'var(--text-light)' }}>分析日志（{section.logs.length} 条）</span>,
                    children: (
                      <ul style={{ margin: 0, paddingLeft: 20 }}>
                        {section.logs.map((item, idx) => (
                          <li key={idx} style={{ marginBottom: 2, color: 'var(--text-light)', fontSize: 11 }}>{item}</li>
                        ))}
                      </ul>
                    )
                  }]}
                />
              )}

              {section.status === 'review' && reportPath && (
                <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px dashed var(--border-color)' }}>
                  <Space direction="vertical" style={{ width: '100%' }}>
                    <Space>
                      <Button
                        type="primary"
                        size="small"
                        icon={<ExperimentOutlined />}
                        loading={aiAnalyzing[section.key]}
                        onClick={() => handleAiReanalyze(section.key)}
                      >
                        AI 分析图表
                      </Button>
                      <span style={{ fontSize: 12, color: 'var(--text-light)' }}>逐张分析报告中的频响曲线图</span>
                    </Space>
                    {aiProgress[section.key] && (aiProgress[section.key].imageTotal > 0 || aiProgress[section.key].imageCount > 0) && (
                      <div style={{ marginTop: 8, padding: 10, background: 'var(--surface-muted)', borderRadius: 6, border: '1px solid var(--border-color)' }}>
                        <div style={{ marginBottom: 6, display: 'flex', alignItems: 'center', gap: 8 }}>
                          <Tag color={aiProgress[section.key].status === 'done' ? 'success' : aiProgress[section.key].status === 'analyzing' ? 'processing' : aiProgress[section.key].status === 'preparing' ? 'default' : 'default'}>
                            {aiProgress[section.key].status === 'done' ? '已完成' : aiProgress[section.key].status === 'analyzing' ? 'AI分析中' : aiProgress[section.key].status === 'preparing' ? '准备中' : aiProgress[section.key].status || '处理中'}
                          </Tag>
                          <span style={{ fontSize: 12, fontWeight: 500 }}>
                            {aiProgress[section.key].status === 'preparing' ? '图片 ' + aiProgress[section.key].imageCurrent + '/' + aiProgress[section.key].imageTotal : ''}
                            {aiProgress[section.key].status === 'analyzing' && aiProgress[section.key].imageTotal > 1 ? '批次 ' + aiProgress[section.key].imageCurrent + '/' + aiProgress[section.key].imageTotal : ''}
                            {aiProgress[section.key].status === 'done' ? '分析完成' : ''}
                          </span>
                        </div>
                        {aiProgress[section.key].status === 'preparing' && aiProgress[section.key].imageTotal > 0 && (
                          <Progress percent={Math.round((aiProgress[section.key].imageCurrent / aiProgress[section.key].imageTotal) * 100)} size="small" style={{ marginBottom: 4 }} strokeColor="#1677ff" />
                        )}
                        {aiProgress[section.key].status === 'analyzing' && aiProgress[section.key].imageTotal > 1 && (
                          <Progress percent={Math.round((aiProgress[section.key].imageCurrent / aiProgress[section.key].imageTotal) * 100)} size="small" style={{ marginBottom: 4 }} />
                        )}
                        <div style={{ marginTop: 4 }}>
                          <span style={{ fontSize: 11, color: 'var(--text-light)' }}>
                            {aiProgress[section.key].status === 'preparing' && aiProgress[section.key].fileName ? '准备: ' + aiProgress[section.key].fileName : ''}
                            {aiProgress[section.key].status === 'analyzing' ? '批次: ' + (aiProgress[section.key].fileName || '') + (aiProgress[section.key].imageCount ? ' (' + aiProgress[section.key].imageCount + ' 张曲线图)' : '') : ''}
                            {aiProgress[section.key].status === 'done' ? (aiProgress[section.key].fileName || '') + (aiProgress[section.key].imageCount ? ' (' + aiProgress[section.key].imageCount + ' 张)' : '') : ''}
                          </span>
                        </div>
                        {aiProgress[section.key].detail && (
                          <div style={{ marginTop: 2 }}>
                            <span style={{ fontSize: 10, color: 'var(--text-light)', fontStyle: 'italic' }}>{aiProgress[section.key].detail}</span>
                          </div>
                        )}
                      </div>
                    )}
                  </Space>
                  {aiResults[section.key] && (
                    <Alert
                      style={{ marginTop: 8 }}
                      type={aiResults[section.key].status === 'pass' ? 'success' : aiResults[section.key].status === 'error' ? 'error' : 'warning'}
                      showIcon
                      message={aiResults[section.key].status === 'pass' ? 'AI 分析通过' : 'AI 分析结果'}
                      description={aiResults[section.key].issues?.map(i => i.message).join('; ') || '无具体问题'}
                    />
                  )}
                </div>
              )}

              {section.data && section.key === 'tableOfContents' && Array.isArray(section.data.chapters) && section.data.chapters.length > 0 && (
                <div style={{ marginTop: 16 }}>
                  <h4 style={{ marginBottom: 8 }}>识别到的章节：</h4>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid var(--border-color)' }}>
                        <th style={{ padding: 8, textAlign: 'left' }}>章节号</th>
                        <th style={{ padding: 8, textAlign: 'left' }}>标题</th>
                      </tr>
                    </thead>
                    <tbody>
                      {section.data.chapters.map((chapter, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--border-color)' }}>
                          <td style={{ padding: 8 }}>{chapter.number}</td>
                          <td style={{ padding: 8 }}>{chapter.title}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {section.data && section.key === 'engineers' && Array.isArray(section.data.engineers) && section.data.engineers.length > 0 && (
                <div style={{ marginTop: 16 }}>
                  <h4 style={{ marginBottom: 8 }}>识别到的人员：</h4>
                  <ul style={{ paddingLeft: 20 }}>
                    {section.data.engineers.map((engineer, idx) => (
                      <li key={idx} style={{ marginBottom: 4 }}>{engineer.name}</li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          )
        }))}
      />
    </div>
  );
}
