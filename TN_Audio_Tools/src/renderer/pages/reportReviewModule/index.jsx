import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Alert, App as AntdApp, Button, Card, Col, Collapse, Empty, Modal, Progress, Row, Space, Table, Tag } from 'antd';
import { InfoCircleOutlined } from '@ant-design/icons';
import { clearWordReviewHistory, readWordReviewHistory, recordWordReviewResult } from '../../modules/reportReview/storage';
import { reviewAreas } from './constants';
import { createReviewHistoryColumns } from './reviewHistoryColumns';
import { buildReviewDigest, getReviewSectionsByStatus, reviewStatusColor, reviewStatusText } from './reviewSummary';
import { normalizeReportPaths, getReportName, getFileBaseName, isDocFile, getProgressLabel, detectFilePairs } from './utils';
import DetailModal from './components/DetailModal';
import StatusDetailModal from './components/StatusDetailModal';
import ReviewAreaModal from './components/ReviewAreaModal';
import CrossReportPanel from './components/CrossReportPanel';
import { useTheme } from '../../ThemeContext';
import '../../styles/pages.css';

export default function ReportReviewPage() {
  const { message, modal } = AntdApp.useApp();
  const isDarkTheme = useTheme() === 'dark';
  const reviewDropzoneBaseColor = isDarkTheme ? '#221d38' : '#faf8ff';
  const reviewDropzoneHoverColor = isDarkTheme ? '#2a2245' : '#f3eeff';
  const reviewDropzoneTitleColor = isDarkTheme ? '#f4f0ff' : '#22075e';
  const reviewDropzoneTextColor = isDarkTheme ? '#c3b8e4' : '#8c8c8c';
  const reviewSelectionPanelColor = isDarkTheme
    ? { backgroundColor: '#221d38', border: '1px solid #4a3a72', textColor: '#dbd4f7', accentColor: '#b37feb', metaColor: '#c3b8e4' }
    : { backgroundColor: '#f9f0ff', border: '1px solid #d3adf7', textColor: '#531dab', accentColor: '#722ed1', metaColor: '#4b3d6e' };
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
  const [chartProgress, setChartProgress] = useState(null);
  const chartProgressUnsubRef = useRef(null);

  useEffect(() => {
    return () => {
      try { if (typeof chartProgressUnsubRef.current === 'function') chartProgressUnsubRef.current(); } catch (_) {}
    };
  }, []);

  const listenForChartProgress = () => {
    try {
      if (typeof chartProgressUnsubRef.current === 'function') chartProgressUnsubRef.current();
      if (window.electron && window.electron.reportReview && window.electron.reportReview.onChartAnalysisProgress) {
        var unsub = window.electron.reportReview.onChartAnalysisProgress(function(data) {
          setChartProgress(data);
        });
        chartProgressUnsubRef.current = unsub;
      }
    } catch (_) {}
  };

  const stopChartProgress = () => {
    try { if (typeof chartProgressUnsubRef.current === 'function') chartProgressUnsubRef.current(); } catch (_) {}
    chartProgressUnsubRef.current = null;
    setChartProgress(null);
  };

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

    listenForChartProgress();
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
      stopChartProgress();
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
            className="report-checker-card report-review-main-card"
            style={{ borderColor: '#d3adf7' }}
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
                      border: '2px dashed #722ed1',
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
                      event.currentTarget.style.borderColor = '#9254de';
                    }}
                    onMouseLeave={(event) => {
                      event.currentTarget.style.borderColor = '#722ed1';
                    }}
                  >
                    <div style={{ fontSize: 32, marginBottom: 10 }}>🔍</div>
                    <div style={{ fontSize: 20, fontWeight: 600, color: reviewDropzoneTitleColor, marginBottom: 6 }}>点击或拖拽选择多个文件</div>
                    <div style={{ fontSize: 13, color: reviewDropzoneTextColor }}>支持一次选择或拖入多个 .doc / .docx / .xlsx 报告</div>
                    <div style={{ fontSize: 12, color: reviewDropzoneTextColor, marginTop: 4, opacity: 0.8 }}>建议同时选择同名的 .docx 和 .xlsx，配对后可获得最完整的审查结果</div>
                  </div>

                  {selectedReportPaths.length > 0 && (() => {
                    const { pairs: selPairs, solo: selSolo } = detectFilePairs(selectedReportPaths);
                    const hasDocLegacy = selectedReportPaths.some((p) => isDocFile(p));
                    const hasDocxOrDoc = selectedReportPaths.some((p) => {
                      const lp = p.toLowerCase();
                      return lp.endsWith('.docx') || (lp.endsWith('.doc') && !lp.endsWith('.docx'));
                    });
                    const hasXlsx = selectedReportPaths.some((p) => p.toLowerCase().endsWith('.xlsx'));
                    const xlsxOnlyNoPair = !hasDocxOrDoc && hasXlsx && selPairs.length === 0;

                    return (
                      <>
                        {hasDocLegacy && (
                          <Alert
                            type="warning"
                            showIcon
                            style={{ marginBottom: 10 }}
                            message="检测到 .doc 格式报告"
                            description=".doc 为旧版二进制格式，需本机安装 Microsoft Word 或 WPS 才能获得完整检查能力。建议将报告另存为 .docx 格式后重新导入，或在装有 Word 的电脑上使用本工具。"
                          />
                        )}
                        {xlsxOnlyNoPair && (
                          <Alert
                            type="info"
                            showIcon
                            style={{ marginBottom: 10 }}
                            message="仅检测到 xlsx 报告，缺少 Word 文档"
                            description="文档结构检查（目录/章节/元数据/POLQA等8项）仅对Word报告有效。建议同时导入同名的 .docx 文件，与 .xlsx 配对后可获得完整的17项检查。"
                          />
                        )}
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
                                const parts = [];
                                if (selPairs.length > 0) parts.push(`${selPairs.length} 组配对`);
                                if (selSolo.length > 0) parts.push(`${selSolo.length} 份单独报告`);
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
                          {chartProgress && chartProgress.imageTotal > 0 && (
                            <div style={{ marginTop: 10, padding: 8, background: 'rgba(0,0,0,0.03)', borderRadius: 6 }}>
                              <div style={{ fontSize: 11, color: reviewSelectionPanelColor.metaColor, marginBottom: 4 }}>
                                图表分析 {chartProgress.imageCurrent}/{chartProgress.imageTotal}: {chartProgress.fileName}
                              </div>
                              <Progress
                                percent={Math.round((chartProgress.imageCurrent / chartProgress.imageTotal) * 100)}
                                size="small"
                                format={() => chartProgress.imageCurrent + '/' + chartProgress.imageTotal}
                              />
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  </>
                );
              })()}
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

        <CrossReportPanel
          crossReportResults={crossReportResults}
          latestReviewDigests={latestReviewDigests}
          onOpenDetail={(record) => { setDetailModalData(record); setDetailModalVisible(true); }}
          onOpenStatusDetail={(crossRecord, status) => {
            setSelectedStatusDetail({
              record: { result: crossRecord.reviewResult, reportName: getReportName(crossRecord.docxPath || crossRecord.reportPath || '') },
              status
            });
            setStatusDetailModalVisible(true);
          }}
        />

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

      <DetailModal
        open={detailModalVisible}
        reportName={detailModalData?.reportName}
        resultData={detailModalData?.result}
        hideCrossReportSections={!crossReportResults || crossReportResults.results?.length < 2}
        onClose={() => { setDetailModalVisible(false); setDetailModalData(null); }}
      />

      <StatusDetailModal
        open={statusDetailModalVisible}
        selectedDetail={selectedStatusDetail}
        filteredSections={filteredStatusSections}
        onClose={() => { setStatusDetailModalVisible(false); setSelectedStatusDetail(null); }}
      />

      <ReviewAreaModal
        open={reviewAreaModalVisible}
        area={selectedReviewArea}
        onClose={() => { setReviewAreaModalVisible(false); setSelectedReviewArea(null); }}
      />
    </div>
  );
}
