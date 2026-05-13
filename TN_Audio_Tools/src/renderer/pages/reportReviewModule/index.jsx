import React, { useMemo, useState } from 'react';
import { Alert, App as AntdApp, Button, Card, Col, Collapse, Empty, Modal, Progress, Row, Space, Table, Tag } from 'antd';
import { InfoCircleOutlined } from '@ant-design/icons';
import { clearWordReviewHistory, readWordReviewHistory, recordWordReviewResult } from '../../modules/reportReview/storage';
import { reviewAreas } from './constants';
import { createReviewHistoryColumns } from './reviewHistoryColumns';
import { buildReviewDigest, getReviewSectionsByStatus, reviewStatusColor, reviewStatusText } from './reviewSummary';
import ReviewResultContent from './ReviewResultContent';
import '../../styles/pages.css';

function normalizeReportPaths(filePaths) {
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

function getReportName(filePath) {
  return String(filePath || '').split('\\').pop() || filePath;
}

function getFileBaseName(filePath) {
  const name = getReportName(filePath);
  return name.replace(/\.(docx?|xlsx)$/i, '');
}

function isDocFile(filePath) {
  const lower = (filePath || '').toLowerCase();
  return lower.endsWith('.doc') && !lower.endsWith('.docx');
}

function getProgressLabel(filePath, prefix, baseName) {
  const label = baseName || getReportName(filePath);
  return isDocFile(filePath) ? `正在转换Word格式: ${label}` : `${prefix}: ${label}`;
}

function detectFilePairs(filePaths) {
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

export default function ReportReviewPage() {
  const { message, modal } = AntdApp.useApp();
  const isDarkTheme = typeof document !== 'undefined' && document.documentElement.dataset.theme === 'dark';
  const reviewDropzoneBaseColor = isDarkTheme ? '#24314a' : '#f5f7fa';
  const reviewDropzoneHoverColor = isDarkTheme ? '#2a3955' : '#e6f7ff';
  const reviewDropzoneTitleColor = isDarkTheme ? '#f4f7ff' : '#262626';
  const reviewDropzoneTextColor = isDarkTheme ? '#c8d3e8' : '#8c8c8c';
  const reviewSelectionPanelColor = isDarkTheme
    ? { backgroundColor: '#24314a', border: '1px solid #425272', textColor: '#dbe5f7', accentColor: '#9ab1ff', metaColor: '#b8c7e6' }
    : { backgroundColor: '#e6f7ff', border: '1px solid #91d5ff', textColor: '#0050b3', accentColor: '#0050b3', metaColor: '#4b6381' };
  const [wordReviewHistory, setWordReviewHistory] = useState(() => readWordReviewHistory() || []);
  const [reviewLoading, setReviewLoading] = useState(false);
  const [selectedReportPaths, setSelectedReportPaths] = useState([]);
  const [batchProgress, setBatchProgress] = useState(null);
  const [detailModalVisible, setDetailModalVisible] = useState(false);
  const [detailModalData, setDetailModalData] = useState(null);
  const [statusDetailModalVisible, setStatusDetailModalVisible] = useState(false);
  const [selectedStatusDetail, setSelectedStatusDetail] = useState(null);
  const [reviewAreaModalVisible, setReviewAreaModalVisible] = useState(false);
  const [selectedReviewArea, setSelectedReviewArea] = useState(null);
  const [crossReportResults, setCrossReportResults] = useState(null);
  const [crossReportModalVisible, setCrossReportModalVisible] = useState(false);

  const safeWordReviewHistory = Array.isArray(wordReviewHistory) ? wordReviewHistory : [];

  const latestReviewDigests = useMemo(
    () => safeWordReviewHistory.slice(0, 3).map((record) => ({
      ...record,
      digest: buildReviewDigest(record?.result)
    })),
    [safeWordReviewHistory]
  );

  const openStatusDetail = (record, status) => {
    setSelectedStatusDetail({ record, status });
    setStatusDetailModalVisible(true);
  };

  const historyColumns = useMemo(() => createReviewHistoryColumns(
    (record) => {
      setDetailModalData(record);
      setDetailModalVisible(true);
    },
    (record, status) => {
      openStatusDetail(record, status);
    }
  ), []);

  const openReviewAreaDetail = (area) => {
    setSelectedReviewArea(area);
    setReviewAreaModalVisible(true);
  };

  const filteredStatusSections = useMemo(() => {
    if (!selectedStatusDetail?.record?.result || !selectedStatusDetail?.status) {
      return [];
    }

    return getReviewSectionsByStatus(selectedStatusDetail.record.result, selectedStatusDetail.status);
  }, [selectedStatusDetail]);

  const handleReportSelection = (filePaths) => {
    const nextFilePaths = normalizeReportPaths(filePaths);

    if (nextFilePaths.length === 0) {
      message.warning('未检测到可用的 .doc、.docx 或 .xlsx 报告');
      return;
    }

    setSelectedReportPaths(nextFilePaths);
    setBatchProgress(null);

    const { pairs, solo } = detectFilePairs(nextFilePaths);
    const parts = [];
    if (pairs.length > 0) parts.push(`${pairs.length} 组配对`);
    if (solo.length > 0) parts.push(`${solo.length} 份单独报告`);
    message.success(`已选择 ${parts.join(' + ')}`);
  };

  const removeSelectedReport = (reportPath) => {
    setSelectedReportPaths((currentPaths) => currentPaths.filter((item) => item !== reportPath));
  };

  const handleSelectFile = async () => {
    try {
      const result = await window.electron.ipcRenderer.invoke('dialog:open-file', {
        filters: [
          { name: '报告文件', extensions: ['doc', 'docx', 'xlsx'] },
          { name: '所有文件', extensions: ['*'] }
        ],
        properties: ['openFile', 'multiSelections']
      });

      if (!result.canceled && result.filePath && result.filePath.length > 0) {
        handleReportSelection(result.filePath);
      }
    } catch (error) {
      message.error(error?.message || '选择文件失败');
      setSelectedReportPaths([]);
    }
  };

  const performReview = async () => {
    if (selectedReportPaths.length === 0) {
      return;
    }

    const { pairs, solo } = detectFilePairs(selectedReportPaths);
    const totalTasks = pairs.length + solo.length;

    setReviewLoading(true);
    setBatchProgress({
      total: totalTasks,
      completed: 0,
      successCount: 0,
      failedCount: 0,
      currentFileName: pairs.length > 0
        ? getProgressLabel(pairs[0].docx, '配对审查', pairs[0].baseName)
        : getProgressLabel(solo[0], '正在审查', getReportName(solo[0]))
    });

    try {
      let successCount = 0;
      let failedCount = 0;
      const failedReports = [];
      const allResults = [];
      let taskIndex = 0;

      // 处理配对文件
      for (const pair of pairs) {
        setBatchProgress({
          total: totalTasks,
          completed: taskIndex,
          successCount,
          failedCount,
          currentFileName: getProgressLabel(pair.docx, '配对审查', pair.baseName)
        });

        try {
          const result = await window.electron.reportReview.reviewPairedReport({
            docxPath: pair.docx,
            xlsxPath: pair.xlsx
          });

          recordWordReviewResult(pair.docx, result, pair.xlsx);
          allResults.push(result);
          successCount += 1;
        } catch (error) {
          failedCount += 1;
          failedReports.push({
            reportPath: `${pair.baseName} (配对)`,
            message: error?.message || '配对审查失败'
          });
        }

        taskIndex += 1;
        setBatchProgress({
          total: totalTasks,
          completed: taskIndex,
          successCount,
          failedCount,
          currentFileName: getProgressLabel(pair.docx, '配对审查', pair.baseName)
        });
      }

      // 处理单独文件
      for (const reportPath of solo) {
        setBatchProgress({
          total: totalTasks,
          completed: taskIndex,
          successCount,
          failedCount,
          currentFileName: getProgressLabel(reportPath, '正在审查', getReportName(reportPath))
        });

        try {
          const result = await window.electron.reportReview.reviewWordReport({
            reportPath
          });

          recordWordReviewResult(reportPath, result);
          allResults.push(result);
          successCount += 1;
        } catch (error) {
          failedCount += 1;
          failedReports.push({
            reportPath,
            message: error?.message || '审查失败'
          });
        }

        taskIndex += 1;
        setBatchProgress({
          total: totalTasks,
          completed: taskIndex,
          successCount,
          failedCount,
          currentFileName: getProgressLabel(reportPath, '正在审查', getReportName(reportPath))
        });
      }

      setWordReviewHistory(readWordReviewHistory());

      // 跨报告对比：≥2份结果时自动触发
      if (allResults.length >= 2) {
        try {
          const crossResults = await window.electron.reportReview.runCrossReportChecks({
            results: allResults
          });
          setCrossReportResults({
            results: crossResults,
            checkedAt: new Date().toISOString(),
            reportCount: crossResults.length
          });
        } catch (_) {
          setCrossReportResults(null);
        }
      } else {
        setCrossReportResults(null);
      }

      if (failedReports.length === 0) {
        const pairMsg = pairs.length > 0 ? `${pairs.length} 组配对、` : '';
        message.success(`批量审查完成，${pairMsg}${solo.length} 份单独报告，共 ${successCount} 项审查`);
      } else {
        modal.warning({
          title: '批量审查完成，部分报告失败',
          width: 680,
          content: (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <div>{`成功 ${successCount} 项，失败 ${failedCount} 项。失败项如下：`}</div>
              <div style={{ maxHeight: 240, overflowY: 'auto', paddingRight: 4 }}>
                {failedReports.map((item) => (
                  <div key={item.reportPath} style={{ marginBottom: 10, padding: '10px 12px', borderRadius: 10, background: 'color-mix(in srgb, #cf1322 10%, var(--surface-color))', border: '1px solid color-mix(in srgb, #cf1322 28%, var(--border-color))' }}>
                    <div style={{ fontWeight: 600, color: '#cf1322', marginBottom: 4 }}>{getReportName(item.reportPath)}</div>
                    <div style={{ color: 'var(--text-light)', fontSize: 12, marginBottom: 4 }}>{typeof item.reportPath === 'string' ? item.reportPath : ''}</div>
                    <div style={{ color: 'var(--text-color)' }}>{item.message}</div>
                  </div>
                ))}
              </div>
            </div>
          )
        });
      }
    } catch (error) {
      message.error(error?.message || '审查失败');
    } finally {
      setBatchProgress((currentProgress) => currentProgress ? {
        ...currentProgress,
        currentFileName: null
      } : null);
      setReviewLoading(false);
    }
  };

  return (
    <div className="page-container">
      <Row gutter={[24, 24]}>
        <Col xs={24}>
          <Card
            className="report-checker-card"
            style={{ borderColor: '#d6e4ff' }}
            styles={{ body: { padding: '10px 24px' } }}
          >
            <Row gutter={[24, 8]} align="middle" justify="center">
              <Col xs={0} lg={4} />
              <Col xs={24} lg={16}>
                <div style={{ width: '100%', maxWidth: 760, margin: '0 auto', display: 'flex', flexDirection: 'column', gap: 12 }}>
                  <div
                    onClick={handleSelectFile}
                    onDragOver={(event) => {
                      event.preventDefault();
                      event.currentTarget.style.backgroundColor = reviewDropzoneHoverColor;
                    }}
                    onDragLeave={(event) => {
                      event.preventDefault();
                      event.currentTarget.style.backgroundColor = reviewDropzoneBaseColor;
                    }}
                    onDrop={(event) => {
                      event.preventDefault();
                      event.currentTarget.style.backgroundColor = reviewDropzoneBaseColor;
                      const files = Array.from(event.dataTransfer.files || []);
                      const filePaths = files.map((file) => file.path).filter(Boolean);
                      if (filePaths.length > 0) {
                        handleReportSelection(filePaths);
                      }
                    }}
                    style={{
                      border: '2px dashed #1677ff',
                      borderRadius: 16,
                      padding: '28px 28px',
                      textAlign: 'center',
                      backgroundColor: reviewDropzoneBaseColor,
                      width: '100%',
                      minHeight: 150,
                      display: 'flex',
                      flexDirection: 'column',
                      justifyContent: 'center',
                      cursor: 'pointer',
                      transition: 'all 0.2s',
                      userSelect: 'none'
                    }}
                    onMouseEnter={(event) => {
                      event.currentTarget.style.borderColor = '#40a9ff';
                    }}
                    onMouseLeave={(event) => {
                      event.currentTarget.style.borderColor = '#1677ff';
                    }}
                  >
                    <div style={{ fontSize: 32, marginBottom: 10 }}>📁</div>
                    <div style={{ fontSize: 20, fontWeight: 600, color: reviewDropzoneTitleColor, marginBottom: 6 }}>点击或拖拽选择多个文件</div>
                    <div style={{ fontSize: 13, color: reviewDropzoneTextColor }}>支持一次选择或拖入多个 .doc / .docx / .xlsx 报告</div>
                    <div style={{ fontSize: 12, color: reviewDropzoneTextColor, marginTop: 4, opacity: 0.8 }}>建议同时选择同名的 .docx 和 .xlsx，配对后可获得最完整的审查结果</div>
                  </div>

                  {selectedReportPaths.length > 0 && (
                    <div
                      style={{
                        width: '100%',
                        padding: '12px 16px',
                        backgroundColor: reviewSelectionPanelColor.backgroundColor,
                        border: reviewSelectionPanelColor.border,
                        borderRadius: 12,
                        fontSize: 13
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', marginBottom: 10 }}>
                        <div style={{ color: reviewSelectionPanelColor.accentColor, fontWeight: 600 }}>
                          {'✓ 已选择 '}
                          {(() => {
                            const { pairs, solo } = detectFilePairs(selectedReportPaths);
                            const parts = [];
                            if (pairs.length > 0) parts.push(`${pairs.length} 组配对`);
                            if (solo.length > 0) parts.push(`${solo.length} 份单独报告`);
                            return parts.join(' + ');
                          })()}
                        </div>
                        <Button size="small" onClick={() => setSelectedReportPaths([])} disabled={reviewLoading}>清空已选</Button>
                      </div>

                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                        {(() => {
                          const { pairs, solo } = detectFilePairs(selectedReportPaths);
                          return (
                            <>
                              {pairs.map((pair) => (
                                <Tag
                                  key={pair.docx}
                                  closable={!reviewLoading}
                                  onClose={(event) => {
                                    event.preventDefault();
                                    removeSelectedReport(pair.docx);
                                    removeSelectedReport(pair.xlsx);
                                  }}
                                  color="purple"
                                  style={{ marginInlineEnd: 0, padding: '4px 10px', borderRadius: 999, maxWidth: '100%', overflow: 'hidden', textOverflow: 'ellipsis' }}
                                >
                                  {pair.baseName} ({pair.docx.toLowerCase().endsWith('.doc') && !pair.docx.toLowerCase().endsWith('.docx') ? '.doc' : '.docx'} + .xlsx)
                                </Tag>
                              ))}
                              {solo.map((reportPath) => (
                                <Tag
                                  key={reportPath}
                                  closable={!reviewLoading}
                                  onClose={(event) => {
                                    event.preventDefault();
                                    removeSelectedReport(reportPath);
                                  }}
                                  style={{ marginInlineEnd: 0, padding: '4px 10px', borderRadius: 999, maxWidth: '100%', overflow: 'hidden', textOverflow: 'ellipsis' }}
                                >
                                  {getReportName(reportPath)}
                                </Tag>
                              ))}
                            </>
                          );
                        })()}
                      </div>

                      {batchProgress && (
                        <div style={{ marginTop: 14 }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, marginBottom: 6, color: reviewSelectionPanelColor.accentColor, fontWeight: 500 }}>
                            <span>
                              {reviewLoading ? `正在处理：${batchProgress.currentFileName || '-'}` : '本轮批量审查已结束'}
                            </span>
                            <span>{batchProgress.completed}/{batchProgress.total}</span>
                          </div>
                          <Progress
                            percent={batchProgress.total > 0 ? Math.round((batchProgress.completed / batchProgress.total) * 100) : 0}
                            status={reviewLoading ? 'active' : 'normal'}
                            strokeColor="#1677ff"
                          />
                          <div style={{ display: 'flex', gap: 16, color: reviewSelectionPanelColor.metaColor, fontSize: 12 }}>
                            <span>成功 {batchProgress.successCount}</span>
                            <span>失败 {batchProgress.failedCount}</span>
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </Col>
              <Col xs={24} lg={4}>
                <div style={{ width: '100%', maxWidth: 220, marginLeft: 'auto', display: 'flex', alignItems: 'center' }}>
                  <Button
                    type="primary"
                    size="large"
                    onClick={performReview}
                    loading={reviewLoading}
                    disabled={selectedReportPaths.length === 0}
                    style={{ width: '100%', height: 52, borderRadius: 14, fontSize: 18, fontWeight: 600 }}
                    block
                  >
                    {reviewLoading ? '批量审查中...' : `开始审查${selectedReportPaths.length > 0 ? `（${selectedReportPaths.length}）` : ''}`}
                  </Button>
                </div>
              </Col>
            </Row>
          </Card>
        </Col>

        {crossReportResults && crossReportResults.results.length >= 2 ? (
          <Col xs={24}>
            <Card
              className="report-checker-card"
              title="本轮批量对比概览"
              extra={<span style={{ color: 'var(--text-light)', fontSize: 12 }}>{crossReportResults.checkedAt ? new Date(crossReportResults.checkedAt).toLocaleString() : ''}</span>}
              style={{ borderColor: '#d6e4ff' }}
            >
              {(() => {
                var reports = crossReportResults.results;
                var codecs = [...new Set(reports.map(function(r) {
                  return r.reviewResult?.reportFacts?.metadata?.codec || '?';
                }))];
                var networks = [...new Set(reports.map(function(r) {
                  return r.reviewResult?.reportFacts?.metadata?.network || '?';
                }))];

                return (
                  <div>
                    <div style={{ display: 'flex', gap: 24, flexWrap: 'wrap', marginBottom: 14 }}>
                      <div><span style={{ color: 'var(--text-light)' }}>报告数：</span><strong>{reports.length}</strong></div>
                      <div><span style={{ color: 'var(--text-light)' }}>Codec：</span><strong>{codecs.join(' · ')}</strong></div>
                      <div><span style={{ color: 'var(--text-light)' }}>网络：</span><strong>{networks.join(' · ')}</strong></div>
                    </div>
                    <Table
                      dataSource={reports.map(function(r) {
                        var s = r.reviewResult?.summary || {};
                        var md = r.reviewResult?.reportFacts?.metadata || {};
                        var name = getReportName(r.docxPath || r.reportPath || '');
                        return {
                          key: name,
                          name: name,
                          network: md.network || '-',
                          codec: md.codec || '-',
                          passed: s.passedChecks || 0,
                          warning: s.warningChecks || 0,
                          review: s.reviewChecks || 0,
                          error: s.errorChecks || 0,
                          total: s.totalChecks || 17,
                          status: r.reviewResult?.overallStatus || 'unknown',
                          record: r
                        };
                      })}
                      rowKey="key"
                      size="small"
                      pagination={false}
                      columns={[
                        { title: '报告', dataIndex: 'name', key: 'name', ellipsis: true, width: 180 },
                        { title: '网络', dataIndex: 'network', key: 'network', width: 70 },
                        { title: 'Codec', dataIndex: 'codec', key: 'codec', width: 70 },
                        { title: '通过', dataIndex: 'passed', key: 'passed', width: 58, align: 'center', render: function(v) { return <span style={{ color: '#52c41a', fontWeight: 600 }}>{v}</span>; } },
                        { title: '警告', dataIndex: 'warning', key: 'warning', width: 58, align: 'center', render: function(v) { return <span style={{ color: v > 0 ? '#faad14' : undefined, fontWeight: v > 0 ? 600 : undefined }}>{v}</span>; } },
                        { title: '复核', dataIndex: 'review', key: 'review', width: 58, align: 'center', render: function(v) { return <span style={{ color: v > 0 ? '#1677ff' : undefined, fontWeight: v > 0 ? 600 : undefined }}>{v}</span>; } },
                        { title: '错误', dataIndex: 'error', key: 'error', width: 58, align: 'center', render: function(v) { return <span style={{ color: v > 0 ? '#f5222d' : undefined, fontWeight: v > 0 ? 600 : undefined }}>{v}</span>; } },
                        { title: '总计', dataIndex: 'total', key: 'total', width: 54, align: 'center' },
                        { title: '状态', dataIndex: 'status', key: 'status', width: 70, render: function(v) { return <Tag color={v === 'pass' ? 'success' : (v === 'error' ? 'error' : (v === 'warning' ? 'warning' : 'processing'))}>{v === 'pass' ? '通过' : (v === 'error' ? '错误' : (v === 'warning' ? '警告' : '复核'))}</Tag>; } }
                      ]}
                    />
                  </div>
                );
              })()}
            </Card>
          </Col>
        ) : (latestReviewDigests.length > 0 ? (
          <Col xs={24}>
            <Card
              className="report-checker-card"
              title="最近审查结论"
              extra={<span style={{ color: 'var(--text-light)', fontSize: 12 }}>批量对比时自动切换为对比概览</span>}
            >
              <Row gutter={[16, 16]}>
                {latestReviewDigests.slice(0, 1).map(function(record) {
                  return (
                    <Col key={record.id} xs={24}>
                      <div
                        onClick={function() { setDetailModalData(record); setDetailModalVisible(true); }}
                        style={{ cursor: 'pointer', padding: '16px 20px', borderRadius: 14, background: 'var(--surface-color)', border: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}
                      >
                        <div>
                          <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 4 }}>{record.reportName}</div>
                          <div style={{ fontSize: 12, color: 'var(--text-light)' }}>
                            {record.checkedAt ? new Date(record.checkedAt).toLocaleString() : '-'}
                            {' · '}通过 {record.digest.summary.passedChecks} / 警告 {record.digest.summary.warningChecks} / 复核 {record.digest.summary.reviewChecks} / 错误 {record.digest.summary.errorChecks}
                          </div>
                        </div>
                        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                          <Button size="small" onClick={function(e) {
                            e.stopPropagation();
                            setDetailModalData(record);
                            setDetailModalVisible(true);
                          }}>查看详情</Button>
                          <Tag color={record.digest.statusColor}>{record.digest.statusText}</Tag>
                        </div>
                      </div>
                    </Col>
                  );
                })}
              </Row>
            </Card>
          </Col>
        ) : null)}

        <Col xs={24}>
          <Card
            className="report-checker-card"
            title="最近审查记录"
            extra={safeWordReviewHistory.length > 0 ? (
              <Button
                danger
                size="small"
                onClick={() => {
                  modal.confirm({
                    title: '清空审查历史',
                    content: '确定要清空所有审查记录吗？此操作无法撤销。',
                    okText: '确定',
                    cancelText: '取消',
                    okButtonProps: { danger: true },
                    onOk() {
                      clearWordReviewHistory();
                      setWordReviewHistory([]);
                      message.success('审查历史已清空');
                    }
                  });
                }}
              >
                清空列表
              </Button>
            ) : null}
          >
            {safeWordReviewHistory.length > 0 ? (
              <Table columns={historyColumns} dataSource={safeWordReviewHistory} rowKey="id" pagination={{ pageSize: 6 }} scroll={{ x: 960 }} />
            ) : (
              <Empty description="暂无审查记录" style={{ margin: '24px 0' }} />
            )}
          </Card>
        </Col>

        {crossReportResults && crossReportResults.results.length >= 2 && (
          <Col xs={24}>
            <Card
              className="report-checker-card"
              title={`批量对比结果（跨报告） — ${crossReportResults.reportCount} 份报告`}
              extra={<span style={{ color: 'var(--text-light)', fontSize: 12 }}>{crossReportResults.checkedAt ? new Date(crossReportResults.checkedAt).toLocaleString() : ''}</span>}
              style={{ borderColor: '#d6e4ff' }}
            >
              {(() => {
                const reports = crossReportResults.results;
                const codecCheck = reports[0]?.reviewResult?.checks?.contentSameCodecDiffNetwork;
                const networkCheck = reports[0]?.reviewResult?.checks?.contentSameNetworkDiffCodec;

                // 按 codec 分组渲染响度数据
                const codecGroups = {};
                reports.forEach(function(r) {
                  var md = r.reviewResult?.reportFacts?.metadata || {};
                  var tf = r.reviewResult?.testDataFacts;
                  var codec = md.codec || '未知';
                  if (!codecGroups[codec]) codecGroups[codec] = [];
                  var metrics = tf?.loudnessMetrics || [];
                  var minRow = metrics.length > 0 ? metrics.reduce(function(a, b) {
                    return (a.rlr || 99) < (b.rlr || 99) ? a : b;
                  }) : null;
                  codecGroups[codec].push({
                    name: getReportName(r.reportPath || r.docxPath || ''),
                    network: md.network || '-',
                    slr: minRow?.slr != null ? Number(minRow.slr).toFixed(1) : '-',
                    rlr: minRow?.rlr != null ? Number(minRow.rlr).toFixed(1) : '-',
                    stmr: minRow?.stmr != null ? Number(minRow.stmr).toFixed(1) : '-',
                    path: r.reportPath || r.docxPath
                  });
                });

                return (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
                    {/* 同Codec不同网络 */}
                    {Object.keys(codecGroups).map(function(codec) {
                      var items = codecGroups[codec];
                      if (items.length < 2) return null;
                      var diffCheck = codecCheck;
                      return (
                        <div key={'codec-' + codec} style={{ border: '1px solid var(--border-color)', borderRadius: 12, padding: '16px 20px' }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
                            <div style={{ fontWeight: 600, fontSize: 16 }}>同 Codec 不同网络响度差异 — {codec}</div>
                            <Tag color={diffCheck?.status === 'pass' ? 'success' : (diffCheck?.status === 'warning' ? 'warning' : (diffCheck?.status === 'error' ? 'error' : 'default'))}>
                              {diffCheck?.status === 'pass' ? '通过' : (diffCheck?.status === 'warning' ? '警告' : (diffCheck?.status === 'error' ? '错误' : '待对比'))}
                            </Tag>
                          </div>
                          <Table
                            dataSource={items}
                            rowKey="name"
                            size="small"
                            pagination={false}
                            columns={[
                              { title: '报告', dataIndex: 'name', key: 'name', ellipsis: true },
                              { title: '网络', dataIndex: 'network', key: 'network', width: 80 },
                              { title: 'RLR', dataIndex: 'rlr', key: 'rlr', width: 70 },
                              { title: 'SLR', dataIndex: 'slr', key: 'slr', width: 70 },
                              { title: 'STMR', dataIndex: 'stmr', key: 'stmr', width: 70 }
                            ]}
                          />
                          {diffCheck?.issues && diffCheck.issues.length > 0 && (
                            <div style={{ marginTop: 10 }}>
                              {diffCheck.issues.map(function(iss, i) {
                                return (
                                  <Alert
                                    key={i}
                                    type={iss.severity === 'error' ? 'error' : (iss.severity === 'warning' ? 'warning' : 'info')}
                                    showIcon
                                    message={iss.message}
                                    style={{ marginBottom: 6 }}
                                  />
                                );
                              })}
                            </div>
                          )}
                          {diffCheck?.evidence && diffCheck.evidence.length > 0 && (() => {
                            var evList = diffCheck.evidence;
                            var pairLines = evList.filter(function(l) { return /diff=\d+\.\d+d?B/.test(l); });
                            var summaryLine = evList.filter(function(l) { return l.indexOf('✓') === 0 && l.indexOf('diff') === -1; })[0] || '';
                            var headerLines = evList.filter(function(l) { return l.indexOf(':') > -1 && l.indexOf('个') > -1 && l.indexOf('diff') === -1; });
                            return (
                              <div style={{ marginTop: 14, padding: '14px 16px', background: diffCheck?.status === 'pass' ? '#f6ffed' : (diffCheck?.status === 'warning' ? '#fffbe6' : '#fff2f0'), borderRadius: 10, border: '1px solid ' + (diffCheck?.status === 'pass' ? '#b7eb8f' : (diffCheck?.status === 'warning' ? '#ffe58f' : '#ffa39e')) }}>
                                <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 10, color: diffCheck?.status === 'pass' ? '#389e0d' : (diffCheck?.status === 'warning' ? '#d48806' : '#cf1322') }}>
                                  {summaryLine || (diffCheck?.status === 'pass' ? '✓ 所有差异均在1dB以内' : '存在超出阈值的差异')}
                                </div>
                                {headerLines.length > 0 && (
                                  <div style={{ fontSize: 13, color: '#555', marginBottom: 8 }}>{headerLines.join(' | ')}</div>
                                )}
                                {pairLines.length > 0 && (
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                                    {pairLines.map(function(line, i) {
                                      var isOK = line.indexOf('✓') > -1;
                                      return (
                                        <div key={i} style={{ fontSize: 13, color: isOK ? '#389e0d' : '#cf1322', fontFamily: "'JetBrains Mono', 'Fira Code', monospace", padding: '3px 0' }}>
                                          {isOK ? '✓ ' : '✗ '}{line.replace(/^\s+/, '').replace(/ ✓$/, '')}
                                        </div>
                                      );
                                    })}
                                  </div>
                                )}
                              </div>
                            );
                          })()}
                        </div>
                      );
                    })}

                    {/* 同网络不同Codec */}
                    <div style={{ border: '1px solid var(--border-color)', borderRadius: 12, padding: '16px 20px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
                        <div style={{ fontWeight: 600, fontSize: 16 }}>同网络不同 Codec 响度差异</div>
                        <Tag color={networkCheck?.status === 'pass' ? 'success' : (networkCheck?.status === 'warning' ? 'warning' : (networkCheck?.status === 'error' ? 'error' : 'default'))}>
                          {networkCheck?.status === 'pass' ? '通过' : (networkCheck?.status === 'warning' ? '警告' : (networkCheck?.status === 'error' ? '错误' : '待对比'))}
                        </Tag>
                      </div>
                      {networkCheck?.issues && networkCheck.issues.length > 0 && (
                        <div>
                          {networkCheck.issues.map(function(iss, i) {
                            return (
                              <Alert
                                key={i}
                                type={iss.severity === 'error' ? 'error' : (iss.severity === 'warning' ? 'warning' : 'info')}
                                showIcon
                                message={iss.message}
                                style={{ marginBottom: 6 }}
                              />
                            );
                          })}
                        </div>
                      )}
                      {networkCheck?.evidence && networkCheck.evidence.length > 0 && (() => {
                        var evList = networkCheck.evidence;
                        var pairLines = evList.filter(function(l) { return /diff=\d+\.\d+d?B/.test(l); });
                        var summaryLine = evList.filter(function(l) { return l.indexOf('✓') === 0 && l.indexOf('diff') === -1; })[0] || '';
                        var hasMultiple = evList.some(function(l) { return l.indexOf('个codec') > -1; });
                        if (!hasMultiple && pairLines.length === 0) {
                          return <div style={{ marginTop: 10, color: 'var(--text-light)', fontSize: 13 }}>各报告 codec 相同，无跨 codec 对比数据</div>;
                        }
                        return (
                          <div style={{ marginTop: 14, padding: '14px 16px', background: networkCheck?.status === 'pass' ? '#f6ffed' : (networkCheck?.status === 'warning' ? '#fffbe6' : '#fff2f0'), borderRadius: 10, border: '1px solid ' + (networkCheck?.status === 'pass' ? '#b7eb8f' : (networkCheck?.status === 'warning' ? '#ffe58f' : '#ffa39e')) }}>
                            <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 10, color: networkCheck?.status === 'pass' ? '#389e0d' : (networkCheck?.status === 'warning' ? '#d48806' : '#cf1322') }}>
                              {summaryLine || (networkCheck?.status === 'pass' ? '✓ 所有差异均在1dB以内' : '存在超出阈值的差异')}
                            </div>
                            {pairLines.length > 0 && (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                                {pairLines.map(function(line, i) {
                                  var isOK = line.indexOf('✓') > -1;
                                  return (
                                    <div key={i} style={{ fontSize: 13, color: isOK ? '#389e0d' : '#cf1322', fontFamily: "'JetBrains Mono', 'Fira Code', monospace", padding: '3px 0' }}>
                                      {isOK ? '✓ ' : '✗ '}{line.replace(/^\s+/, '').replace(/ ✓$/, '')}
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                );
              })()}
            </Card>
          </Col>
        )}

        <Col xs={24}>
          <Card
            className="report-checker-card"
            title="检查范围说明"
            extra={<span style={{ color: 'var(--text-light)', fontSize: 12 }}>默认收起，按需展开查看</span>}
          >
            <Collapse
              ghost
              items={[
                {
                  key: 'review-scope',
                  label: '展开查看系统检查范围与说明',
                  children: (
                    <div>
                      <div className="review-area-section-title">系统检查范围</div>
                      <Row gutter={[16, 16]}>
                        {reviewAreas.map((area) => (
                          <Col key={area.title} xs={24} sm={12} md={8}>
                            <Card
                              className="tool-card review-area-card"
                              hoverable
                              style={{ height: '100%' }}
                              onClick={() => openReviewAreaDetail(area)}
                            >
                              <Space direction="vertical" size={12} style={{ width: '100%' }}>
                                <div className="review-area-card__top">
                                  <Tag color={area.color} style={{ width: 'fit-content', marginInlineEnd: 0 }}>{area.tag}</Tag>
                                  <span className="review-area-card__link">点击查看说明</span>
                                </div>
                                <h3 style={{ margin: 0, fontSize: 18 }}>{area.title}</h3>
                                <p style={{ margin: 0, color: 'var(--text-light)', minHeight: 44 }}>{area.description}</p>
                                <div className="review-area-card__footer">
                                  <InfoCircleOutlined /> 查看检查内容与适用场景
                                </div>
                              </Space>
                            </Card>
                          </Col>
                        ))}
                      </Row>
                    </div>
                  )
                }
              ]}
            />
          </Card>
        </Col>
      </Row>

      <Modal
        title={`报告审查详情：${detailModalData?.reportName || ''}`}
        open={detailModalVisible}
        onCancel={() => {
          setDetailModalVisible(false);
          setDetailModalData(null);
        }}
        width={900}
        footer={null}
      >
        {detailModalData?.result ? (
          <ReviewResultContent resultData={detailModalData.result} />
        ) : (
          <Alert type="warning" showIcon message="当前记录没有可展示的详情数据" />
        )}
      </Modal>

      <Modal
        title={selectedStatusDetail ? `${selectedStatusDetail.record?.reportName || ''} - ${reviewStatusText[selectedStatusDetail.status] || selectedStatusDetail.status} 明细` : '状态明细'}
        open={statusDetailModalVisible}
        onCancel={() => {
          setStatusDetailModalVisible(false);
          setSelectedStatusDetail(null);
        }}
        footer={null}
        width={860}
      >
        {selectedStatusDetail ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <Alert
              type={selectedStatusDetail.status === 'pass' ? 'success' : (selectedStatusDetail.status === 'warning' ? 'warning' : (selectedStatusDetail.status === 'error' ? 'error' : 'info'))}
              showIcon
              message={`当前共 ${filteredStatusSections.length} 项“${reviewStatusText[selectedStatusDetail.status] || selectedStatusDetail.status}”检查`}
              description="点击表格中的统计标签即可按类别筛选，不必再从完整详情里逐项查找。"
            />

            {filteredStatusSections.length > 0 ? filteredStatusSections.map((section) => (
              <Card
                key={`${section.key}-${section.status}`}
                size="small"
                title={(
                  <Space>
                    <Tag color={reviewStatusColor[section.status] || 'default'} style={{ marginInlineEnd: 0 }}>
                      {reviewStatusText[section.status] || section.status}
                    </Tag>
                    <span>{section.title || '未命名检查项'}</span>
                  </Space>
                )}
                className="review-status-detail-card"
              >
                <div style={{ color: 'var(--text-light)', marginBottom: 12, lineHeight: 1.7 }}>
                  {section.description || '无详细说明'}
                </div>

                {Array.isArray(section.issues) && section.issues.length > 0 ? (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: section.evidence?.length ? 14 : 0 }}>
                    {section.issues.map((issue, index) => (
                      <Alert
                        key={`${section.key}-issue-${index}`}
                        type={issue?.severity === 'error' ? 'error' : (issue?.severity === 'warning' ? 'warning' : 'info')}
                        showIcon
                        message={issue?.message || '未提供问题说明'}
                      />
                    ))}
                  </div>
                ) : (
                  <Alert
                    type={section.status === 'pass' ? 'success' : 'info'}
                    showIcon
                    style={{ marginBottom: section.evidence?.length ? 14 : 0 }}
                    message={section.status === 'pass' ? '当前检查项已通过，未记录问题项。' : '当前检查项没有单独记录问题描述。'}
                  />
                )}

                {Array.isArray(section.evidence) && section.evidence.length > 0 && (
                  <div>
                    <div style={{ fontWeight: 600, color: 'var(--text-color)', marginBottom: 8 }}>证据记录</div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                      {section.evidence.map((evidence, index) => (
                        <div
                          key={`${section.key}-evidence-${index}`}
                          style={{ padding: '10px 12px', borderRadius: 10, border: '1px solid var(--border-color)', background: 'var(--surface-elevated)', color: 'var(--text-light)', lineHeight: 1.65 }}
                        >
                          {evidence}
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </Card>
            )) : (
              <Empty description={`当前记录没有“${reviewStatusText[selectedStatusDetail.status] || selectedStatusDetail.status}”项`} />
            )}
          </div>
        ) : null}
      </Modal>

      <Modal
        title={selectedReviewArea ? `${selectedReviewArea.tag} - ${selectedReviewArea.title}` : '检查范围说明'}
        open={reviewAreaModalVisible}
        onCancel={() => {
          setReviewAreaModalVisible(false);
          setSelectedReviewArea(null);
        }}
        footer={null}
        width={640}
      >
        {selectedReviewArea ? (
          <div>
            <Alert
              type="info"
              showIcon
              style={{ marginBottom: 16 }}
              message={selectedReviewArea.description}
            />
            <Space direction="vertical" size={12} style={{ width: '100%' }}>
              {(selectedReviewArea.details || []).map((detail, index) => (
                <div key={`${selectedReviewArea.title}-${index}`} className="review-area-modal__item">
                  <span className="review-area-modal__index">{index + 1}</span>
                  <span>{detail}</span>
                </div>
              ))}
            </Space>
          </div>
        ) : null}
      </Modal>
    </div>
  );
}
