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

const HIDDEN_SECTION_KEYS = new Set(['engineers']);

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
  sections = sections.filter(function(section) {
    return section && !HIDDEN_SECTION_KEYS.has(section.key);
  });
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
    <div className="report-review-result">
      {/* 1. 整体评估横幅：白色高级卡片背景+精细立体阴影增强层次 */}
      <div className="report-review-result__hero" style={{
        padding: '24px 28px',
        background: 'var(--surface-color)',
        border: '1px solid var(--border-color)',
        borderLeft: `6px solid ${reviewResult.overallStatus === 'pass' ? 'var(--status-pass)' : reviewResult.overallStatus === 'error' ? 'var(--status-error)' : reviewResult.overallStatus === 'warning' ? 'var(--status-warn)' : 'var(--status-info)'}`,
        borderRadius: 12,
        display: 'flex',
        alignItems: 'center',
        gap: 20,
        marginBottom: 32,
        boxShadow: 'var(--shadow-color)'
      }}>
        <div style={{ fontSize: 32, color: reviewResult.overallStatus === 'pass' ? 'var(--status-pass)' : reviewResult.overallStatus === 'error' ? 'var(--status-error)' : reviewResult.overallStatus === 'warning' ? 'var(--status-warn)' : 'var(--status-info)', display: 'flex', alignItems: 'center' }}>
          {getStatusIcon(reviewResult.overallStatus)}
        </div>
        <div>
          <div className="report-review-result__eyebrow" style={{ fontSize: 12, color: 'var(--text-light)', marginBottom: 4, letterSpacing: 0.5, textTransform: 'uppercase', fontWeight: 600 }}>Overall Assessment</div>
          <div style={{ fontSize: 20, fontWeight: 600, color: 'var(--text-color)' }}>
            整体评估：<span style={{ color: reviewResult.overallStatus === 'pass' ? 'var(--status-pass)' : reviewResult.overallStatus === 'error' ? 'var(--status-error)' : reviewResult.overallStatus === 'warning' ? 'var(--status-warn)' : 'var(--status-info)' }}>{getStatusText(reviewResult.overallStatus)}</span>
          </div>
        </div>
      </div>

      {/* 2. 审查统计卡：白色背景+顶部色彩提示线+悬浮阴影，增加空间感和层次 */}
      <div style={{ marginBottom: 40 }}>
        <div className="report-review-result__section-heading" style={{ marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
          <div className="report-review-result__section-bar" style={{ width: 4, height: 16, background: 'var(--status-info)', borderRadius: 2 }}></div>
          <h3 style={{ margin: 0, fontSize: 16, color: 'var(--text-color)', fontWeight: 600 }}>审查统计</h3>
        </div>
        <Row gutter={[16, 16]}>
          <Col xs={12} sm={6}>
            <div className="report-review-result__stat-card report-review-result__stat-card--pass" style={{ padding: '24px 20px', background: 'var(--surface-color)', borderRadius: 12, border: '1px solid var(--border-color)', borderTop: '4px solid var(--status-pass)', boxShadow: 'var(--shadow-color)' }}>
              <div style={{ fontSize: 14, color: 'var(--text-light)', marginBottom: 8, fontWeight: 500 }}>通过项</div>
              <div style={{ fontSize: 36, fontWeight: 600, color: 'var(--text-color)', lineHeight: 1 }}>{summary.passedChecks}</div>
            </div>
          </Col>
          <Col xs={12} sm={6}>
            <div className="report-review-result__stat-card report-review-result__stat-card--warning" style={{ padding: '24px 20px', background: 'var(--surface-color)', borderRadius: 12, border: '1px solid var(--border-color)', borderTop: '4px solid var(--status-warn)', boxShadow: 'var(--shadow-color)' }}>
              <div style={{ fontSize: 14, color: 'var(--text-light)', marginBottom: 8, fontWeight: 500 }}>警告项</div>
              <div style={{ fontSize: 36, fontWeight: 600, color: 'var(--text-color)', lineHeight: 1 }}>{summary.warningChecks}</div>
            </div>
          </Col>
          <Col xs={12} sm={6}>
            <div className="report-review-result__stat-card report-review-result__stat-card--review" style={{ padding: '24px 20px', background: 'var(--surface-color)', borderRadius: 12, border: '1px solid var(--border-color)', borderTop: '4px solid var(--status-info)', boxShadow: 'var(--shadow-color)' }}>
              <div style={{ fontSize: 14, color: 'var(--text-light)', marginBottom: 8, fontWeight: 500 }}>需复核</div>
              <div style={{ fontSize: 36, fontWeight: 600, color: 'var(--text-color)', lineHeight: 1 }}>{summary.reviewChecks}</div>
            </div>
          </Col>
          <Col xs={12} sm={6}>
            <div className="report-review-result__stat-card report-review-result__stat-card--error" style={{ padding: '24px 20px', background: 'var(--surface-color)', borderRadius: 12, border: '1px solid var(--border-color)', borderTop: '4px solid var(--status-error)', boxShadow: 'var(--shadow-color)' }}>
              <div style={{ fontSize: 14, color: 'var(--text-light)', marginBottom: 8, fontWeight: 500 }}>错误项</div>
              <div style={{ fontSize: 36, fontWeight: 600, color: 'var(--text-color)', lineHeight: 1 }}>{summary.errorChecks}</div>
            </div>
          </Col>
        </Row>
      </div>

      <div className="report-review-result__section-heading" style={{ marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
        <div className="report-review-result__section-bar" style={{ width: 4, height: 16, background: 'var(--status-info)', borderRadius: 2 }}></div>
        <h3 style={{ margin: 0, fontSize: 16, color: 'var(--text-color)', fontWeight: 600 }}>详细检查结果</h3>
      </div>
      <Collapse
        className="report-review-result__collapse"
        expandIconPosition="end"
        style={{ 
          background: 'transparent', 
          border: 'none', 
          borderRadius: 12,
          boxShadow: 'none',
          overflow: 'visible'
        }}
        bordered={false}
        items={sections.map((section, sectionIndex) => ({
          key: section.key || `section-${sectionIndex}`,
          style: { 
            background: 'var(--surface-color)',
            border: '1px solid var(--border-color)',
            borderRadius: 16,
            marginBottom: sectionIndex === sections.length - 1 ? 0 : 14
          },
          label: (
            <div className="report-review-result__panel-label" style={{ padding: '8px 4px' }}>
              <Space size="middle">
                <Tag color={getStatusColor(section.status)} style={{ margin: 0, padding: '0 8px', borderRadius: 4 }}>
                  {getStatusText(section.status || 'review')}
                </Tag>
                <span className="report-review-result__panel-title" style={{ fontWeight: 600, fontSize: 15, color: 'var(--text-color)' }}>{section.title || '未命名检查项'}</span>
              </Space>
            </div>
          ),
          children: (
            <div className="report-review-result__panel-body" style={{ padding: '0 16px 24px 48px' }}>
              <div style={{ marginBottom: 16, color: 'var(--text-light)', fontSize: 14, lineHeight: 1.6 }}>
                {section.description || '无详细说明'}
              </div>

              {section.key === 'contentLoudnessFRTrend' && Array.isArray(section.comparisonCards) && section.comparisonCards.length > 0 && (
                <div style={{ marginBottom: 20 }}>
                  <Card
                    className="report-review-result__nested-card"
                    size="small"
                    style={{
                      marginBottom: 20,
                      borderRadius: 12,
                      border: '1px solid var(--border-color)',
                      background: 'var(--surface-color)',
                      boxShadow: 'var(--shadow-color)',
                    }}
                    bodyStyle={{ padding: 0 }}
                  >
                    <div className="report-review-result__nested-card-header" style={{ padding: '14px 20px', borderBottom: '1px solid var(--border-color)', background: 'var(--surface-elevated)', borderRadius: '12px 12px 0 0', display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span style={{ width: 3, height: 14, background: 'var(--primary-color)', borderRadius: 2 }}></span>
                      <span style={{ fontSize: 14, color: 'var(--text-color)', fontWeight: 600 }}>趋势摘要</span>
                    </div>
                    <div style={{ padding: 20 }}>
                      {trendSummaryItems.length > 0 ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
                          {trendSummaryItems.map(function(item, index) {
                            return (
                              <div key={item.key}>
                                {item.label && (
                                  <div style={{ marginBottom: 6, display: 'flex', alignItems: 'center', gap: 6 }}>
                                    <span className="report-review-result__summary-badge" style={{ fontSize: 12, fontWeight: 600, color: 'var(--primary-color)', background: 'color-mix(in srgb, var(--primary-color) 10%, transparent)', padding: '2px 8px', borderRadius: 4, border: '1px solid color-mix(in srgb, var(--primary-color) 30%, transparent)' }}>
                                      {item.label}
                                    </span>
                                  </div>
                                )}
                                <div style={{ fontSize: 14, color: 'var(--text-color)', lineHeight: 1.7, whiteSpace: 'pre-wrap' }}>
                                  {item.content}
                                </div>
                                {index !== trendSummaryItems.length - 1 && (
                                  <div style={{ height: 1, background: 'var(--border-color)', marginTop: 16 }}></div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      ) : (
                        <div style={{ fontSize: 14, color: 'var(--text-color)', lineHeight: 1.7, whiteSpace: 'pre-wrap' }}>
                          {section.conclusion || '已按方向与等级生成曲线对比卡片，可直接查看 FR 与响度图的对应关系。'}
                        </div>
                      )}
                    </div>
                  </Card>

                  {Object.entries(groupComparisonCards(section.comparisonCards)).map(function(entry) {
                    var direction = entry[0];
                    var cards = entry[1];
                    var shouldHideLevel = direction === 'SND' && cards.length === 1;
                    return (
                      <div key={direction} style={{ marginBottom: 20 }}>
                        <div style={{ marginBottom: 12, display: 'flex', alignItems: 'center', gap: 8 }}>
                          <Tag color="blue" style={{ borderRadius: 999, margin: 0 }}>{direction}</Tag>
                          <span style={{ fontSize: 15, fontWeight: 500, color: 'var(--text-color)' }}>
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
                <div style={{ marginBottom: 16, display: 'flex', flexDirection: 'column', gap: 6 }}>
                  <span style={{ fontSize: 14, fontWeight: 600, color: 'var(--text-color)' }}>判定结论：</span>
                  <div style={{ color: 'var(--text-light)', fontSize: 14, lineHeight: 1.6 }}>{section.conclusion}</div>
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
                          <span style={{ fontSize: 13, color: 'var(--text-light)' }}>
                            查看结构化明细（{section.checklist.length} 项）
                          </span>
                        ),
                        children: (
                          <div style={{ overflowX: 'auto' }}>
                            <table className="report-review-result__table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                              <thead>
                                <tr style={{ borderBottom: '1px solid var(--border-color)', backgroundColor: 'var(--surface-elevated)' }}>
                                  {section.checklist[0]?.direction != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>方向</th>}
                                  {section.checklist[0]?.volumeLevel != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>等级</th>}
                                  {section.checklist[0]?.role != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>类型</th>}
                                  {section.checklist[0]?.imageIndex != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>图#</th>}
                                  <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>结果</th>
                                  <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>详情</th>
                                </tr>
                              </thead>
                              <tbody>
                                {section.checklist.map((item, idx) => (
                                  <tr key={idx} style={{ borderBottom: '1px solid var(--border-color)', backgroundColor: item.status === 'error' ? 'color-mix(in srgb, var(--status-error) 8%, transparent)' : item.status === 'warning' ? 'color-mix(in srgb, var(--status-warn) 8%, transparent)' : 'transparent' }}>
                                    {item.direction != null && <td style={{ padding: '10px 10px' }}><Tag color="blue" style={{ margin: 0 }}>{item.direction}</Tag></td>}
                                    {item.volumeLevel != null && <td style={{ padding: '10px 10px', color: 'var(--text-light)' }}>{item.volumeLevel || '-'}</td>}
                                    {item.role != null && <td style={{ padding: '10px 10px', color: 'var(--text-light)' }}>{item.role === 'reference' ? 'FR基准' : item.role === 'loudness_rlr' ? 'RLR响度' : item.role === 'loudness_slr' ? 'SLR响度' : item.role || '-'}</td>}
                                    {item.imageIndex != null && <td style={{ padding: '10px 10px', color: 'var(--text-light)' }}>#{item.imageIndex}</td>}
                                    <td style={{ padding: '10px 10px' }}>
                                      <Tag color={item.status === 'pass' ? 'success' : item.status === 'warning' ? 'orange' : item.status === 'error' ? 'red' : 'default'} style={{ margin: 0 }}>
                                        {item.status === 'pass' ? '通过' : item.status === 'warning' ? '警告' : item.status === 'error' ? '错误' : item.status}
                                      </Tag>
                                    </td>
                                    <td style={{ padding: '10px 10px', maxWidth: 350, fontSize: 13, color: 'var(--text-light)', lineHeight: 1.5 }}>
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
                      <h4 style={{ marginBottom: 8, color: 'var(--text-color)', fontWeight: 600 }}>
                        逐项检查清单（{section.checklist.length} 项）：
                        <span style={{ fontSize: 13, fontWeight: 'normal', color: 'var(--text-light)', marginLeft: 12 }}>
                          通过 {section.checklist.filter(function(c) { return c.status === 'pass'; }).length}
                          {section.checklist.filter(function(c) { return c.status === 'warning'; }).length > 0 && ' | 警告 ' + section.checklist.filter(function(c) { return c.status === 'warning'; }).length}
                          {section.checklist.filter(function(c) { return c.status === 'error'; }).length > 0 && ' | 错误 ' + section.checklist.filter(function(c) { return c.status === 'error'; }).length}
                        </span>
                      </h4>
                      <div style={{ overflowX: 'auto' }}>
                        <table className="report-review-result__table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                          <thead>
                            <tr style={{ borderBottom: '1px solid var(--border-color)', backgroundColor: 'var(--surface-elevated)' }}>
                              {section.checklist[0]?.direction != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>方向</th>}
                              {section.checklist[0]?.volumeLevel != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>等级</th>}
                              {section.checklist[0]?.role != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>类型</th>}
                              {section.checklist[0]?.fromLevel != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>从</th>}
                              {section.checklist[0]?.toLevel != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>到</th>}
                              {section.checklist[0]?.imageIndex != null && <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>图#</th>}
                              <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>结果</th>
                              <th style={{ padding: '8px 10px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--text-color)', fontWeight: 600 }}>详情</th>
                            </tr>
                          </thead>
                          <tbody>
                            {section.checklist.map((item, idx) => (
                              <tr key={idx} style={{ borderBottom: '1px solid var(--border-color)', backgroundColor: item.status === 'error' ? 'color-mix(in srgb, var(--status-error) 8%, transparent)' : item.status === 'warning' ? 'color-mix(in srgb, var(--status-warn) 8%, transparent)' : 'transparent' }}>
                                {item.direction != null && <td style={{ padding: '10px 10px' }}><Tag color="blue" style={{ margin: 0 }}>{item.direction}</Tag></td>}
                                {item.volumeLevel != null && <td style={{ padding: '10px 10px', color: 'var(--text-light)' }}>{item.volumeLevel || '-'}</td>}
                                {item.role != null && <td style={{ padding: '10px 10px', color: 'var(--text-light)' }}>{item.role === 'reference' ? 'FR基准' : item.role === 'loudness_rlr' ? 'RLR响度' : item.role === 'loudness_slr' ? 'SLR响度' : item.role || '-'}</td>}
                                {item.fromLevel != null && <td style={{ padding: '10px 10px', color: 'var(--text-light)' }}>{item.fromLevel}</td>}
                                {item.toLevel != null && <td style={{ padding: '10px 10px', color: 'var(--text-light)' }}>{item.toLevel}</td>}
                                {item.imageIndex != null && <td style={{ padding: '10px 10px', color: 'var(--text-light)' }}>#{item.imageIndex}</td>}
                                <td style={{ padding: '10px 10px' }}>
                                  <Tag color={item.status === 'pass' ? 'success' : item.status === 'warning' ? 'orange' : item.status === 'error' ? 'red' : 'default'} style={{ margin: 0 }}>
                                    {item.status === 'pass' ? '通过' : item.status === 'warning' ? '警告' : item.status === 'error' ? '错误' : item.status}
                                  </Tag>
                                </td>
                                <td style={{ padding: '10px 10px', maxWidth: 350, fontSize: 13, color: 'var(--text-light)', lineHeight: 1.5 }}>
                                  {item.detail || '-'}
                                  {(item.expectedBehavior || item.actualBehavior) && (
                                    <div style={{ marginTop: 4, color: 'var(--text-light)', fontSize: 12 }}>
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
                <div style={{ marginBottom: 20 }}>
                  <div style={{ fontSize: 14, fontWeight: 600, color: 'var(--text-color)', marginBottom: 8 }}>异常详情：</div>
                  {section.issues.filter(function(i) { return i.severity === 'error'; }).map((issue, idx) => (
                    <Alert
                      key={'err-' + idx}
                      message={
                        <Space wrap size={[4, 4]}>
                          {issue.meta?.direction && <Tag color="blue" bordered={false}>{issue.meta.direction}</Tag>}
                          {issue.meta?.volumeLevel && <Tag color="geekblue" bordered={false}>{issue.meta.volumeLevel}</Tag>}
                          {issue.meta?.frequencyRange && <Tag color="purple" bordered={false}>{issue.meta.frequencyRange}</Tag>}
                          {issue.meta?.imageIndex != null && <Tag bordered={false}>图#{issue.meta.imageIndex}</Tag>}
                          <span style={{ color: 'var(--text-color)', fontWeight: 500 }}>{issue.message}</span>
                        </Space>
                      }
                      description={issue.meta?.expectedBehavior && (
                        <div style={{ fontSize: 13, color: 'var(--text-light)', marginTop: 4 }}>
                          <div>期望: {issue.meta.expectedBehavior}</div>
                          <div>实际: {issue.meta.actualBehavior}</div>
                        </div>
                      )}
                      type="error"
                      showIcon
                      style={{ marginBottom: 8, border: 'none', backgroundColor: 'color-mix(in srgb, var(--status-error) 8%, transparent)' }}
                    />
                  ))}
                  {section.issues.filter(function(i) { return i.severity === 'warning'; }).map((issue, idx) => (
                    <Alert
                      key={'warn-' + idx}
                      message={
                        <Space wrap size={[4, 4]}>
                          {issue.meta?.direction && <Tag color="blue" bordered={false}>{issue.meta.direction}</Tag>}
                          {issue.meta?.volumeLevel && <Tag color="geekblue" bordered={false}>{issue.meta.volumeLevel}</Tag>}
                          {issue.meta?.frequencyRange && <Tag color="purple" bordered={false}>{issue.meta.frequencyRange}</Tag>}
                          <span style={{ color: 'var(--text-color)', fontWeight: 500 }}>{issue.message}</span>
                        </Space>
                      }
                      type="warning"
                      showIcon
                      style={{ marginBottom: 8, border: 'none', backgroundColor: 'color-mix(in srgb, var(--status-warn) 8%, transparent)' }}
                    />
                  ))}
                </div>
              )}

              {/* 3. 诊断依据块：统一采用柔和警示/信息块形式 */}
              {Array.isArray(section.evidence) && section.evidence.length > 0 && (
                <div className="report-review-result__evidence" style={{ 
                  marginBottom: 16, 
                  marginTop: 12,
                  padding: '14px 18px', 
                  background: 'var(--surface-elevated)', 
                  border: '1px solid var(--border-color)', 
                  borderRadius: 8 
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
                    <span style={{ fontSize: 15 }}>💡</span>
                    <span style={{ fontSize: 14, fontWeight: 600, color: 'var(--text-color)' }}>诊断依据</span>
                  </div>
                  <ul style={{ margin: 0, paddingLeft: 24, display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {section.evidence.map((item, idx) => (
                      <li key={idx} style={{ color: 'var(--text-light)', fontSize: 14, lineHeight: 1.6 }}>
                        {item}
                      </li>
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
                    style: { border: 'none', padding: 0 },
                    label: <span style={{ fontSize: 12, color: 'var(--text-light)' }}>分析日志（{section.logs.length} 条）</span>,
                    children: (
                      <ul style={{ margin: 0, paddingLeft: 24 }}>
                        {section.logs.map((item, idx) => (
                          <li key={idx} style={{ marginBottom: 4, color: 'var(--text-light)', fontSize: 12, lineHeight: 1.5 }}>{item}</li>
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
                      <div className="report-review-result__progress-card" style={{ marginTop: 8, padding: 10, background: 'var(--surface-muted)', borderRadius: 6, border: '1px solid var(--border-color)' }}>
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
                          <Progress percent={Math.round((aiProgress[section.key].imageCurrent / aiProgress[section.key].imageTotal) * 100)} size="small" style={{ marginBottom: 4 }} strokeColor="var(--status-info)" />
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
                  <table className="report-review-result__table" style={{ width: '100%', borderCollapse: 'collapse' }}>
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
