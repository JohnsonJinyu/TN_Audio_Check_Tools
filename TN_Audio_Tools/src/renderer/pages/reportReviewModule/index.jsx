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

  const supportedExtensions = new Set(['.doc', '.docx']);
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
      message.warning('未检测到可用的 .doc 或 .docx 报告');
      return;
    }

    setSelectedReportPaths(nextFilePaths);
    setBatchProgress(null);
    message.success(`已选择 ${nextFilePaths.length} 份报告`);
  };

  const removeSelectedReport = (reportPath) => {
    setSelectedReportPaths((currentPaths) => currentPaths.filter((item) => item !== reportPath));
  };

  const handleSelectFile = async () => {
    try {
      const result = await window.electron.ipcRenderer.invoke('dialog:open-file', {
        filters: [
          { name: 'Word 文档', extensions: ['doc', 'docx'] },
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

    setReviewLoading(true);
    setBatchProgress({
      total: selectedReportPaths.length,
      completed: 0,
      successCount: 0,
      failedCount: 0,
      currentFileName: getReportName(selectedReportPaths[0])
    });

    try {
      let successCount = 0;
      let failedCount = 0;
      const failedReports = [];

      for (let index = 0; index < selectedReportPaths.length; index += 1) {
        const reportPath = selectedReportPaths[index];

        setBatchProgress({
          total: selectedReportPaths.length,
          completed: index,
          successCount,
          failedCount,
          currentFileName: getReportName(reportPath)
        });

        try {
          const result = await window.electron.reportReview.reviewWordReport({
            reportPath
          });

          recordWordReviewResult(reportPath, result);
          successCount += 1;
        } catch (error) {
          failedCount += 1;
          failedReports.push({
            reportPath,
            message: error?.message || '审查失败'
          });
        }

        setBatchProgress({
          total: selectedReportPaths.length,
          completed: index + 1,
          successCount,
          failedCount,
          currentFileName: getReportName(reportPath)
        });
      }

      setWordReviewHistory(readWordReviewHistory());

      if (failedReports.length === 0) {
        message.success(`批量审查完成，共处理 ${successCount} 份报告`);
      } else {
        modal.warning({
          title: '批量审查完成，部分报告失败',
          width: 680,
          content: (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <div>{`成功 ${successCount} 份，失败 ${failedCount} 份。失败项如下：`}</div>
              <div style={{ maxHeight: 240, overflowY: 'auto', paddingRight: 4 }}>
                {failedReports.map((item) => (
                  <div key={item.reportPath} style={{ marginBottom: 10, padding: '10px 12px', borderRadius: 10, background: '#fff2f0', border: '1px solid #ffccc7' }}>
                    <div style={{ fontWeight: 600, color: '#cf1322', marginBottom: 4 }}>{getReportName(item.reportPath)}</div>
                    <div style={{ color: '#8c8c8c', fontSize: 12, marginBottom: 4 }}>{item.reportPath}</div>
                    <div style={{ color: '#5c0011' }}>{item.message}</div>
                  </div>
                ))}
              </div>
            </div>
          )
        });
      }
    } catch (error) {
      message.error(error?.message || 'Word 审查失败');
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
                    <div style={{ fontSize: 13, color: reviewDropzoneTextColor }}>支持一次选择或拖入多个 .doc / .docx 报告</div>
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
                        <div style={{ color: reviewSelectionPanelColor.accentColor, fontWeight: 600 }}>✓ 已选择 {selectedReportPaths.length} 份报告</div>
                        <Button size="small" onClick={() => setSelectedReportPaths([])} disabled={reviewLoading}>清空已选</Button>
                      </div>

                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                        {selectedReportPaths.map((reportPath) => (
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

        {latestReviewDigests.length > 0 && (
          <Col xs={24}>
            <Card
              className="report-checker-card"
              title="最近审查结论"
              extra={<span style={{ color: '#8c8c8c', fontSize: 12 }}>无需点详情即可先看初步判断</span>}
            >
              <Row gutter={[16, 16]}>
                {latestReviewDigests.map((record) => (
                  <Col key={record.id} xs={24} lg={8}>
                    <Card
                      className="review-digest-card"
                      size="small"
                      hoverable
                      onClick={() => {
                        setDetailModalData(record);
                        setDetailModalVisible(true);
                      }}
                      data-status={record.digest.overallStatus}
                      style={{
                        height: '100%',
                        borderRadius: 14,
                        borderColor: record.digest.theme.border,
                        background: `linear-gradient(180deg, #ffffff 0%, ${record.digest.theme.soft} 100%)`
                      }}
                      styles={{ body: { display: 'flex', flexDirection: 'column', gap: 10, height: '100%' } }}
                    >
                      <div className="review-digest-card__bar" style={{ backgroundColor: record.digest.theme.accent }} />

                      <div className="review-digest-card__header">
                        <div className="review-digest-card__meta">
                          <div className="review-digest-card__name">
                            {record.reportName}
                          </div>
                          <div className="review-digest-card__time">
                            {record.checkedAt ? new Date(record.checkedAt).toLocaleString() : '-'}
                          </div>
                        </div>
                        <Tag color={record.digest.statusColor} style={{ marginInlineEnd: 0 }}>
                          {record.digest.statusText}
                        </Tag>
                      </div>

                      <div className="review-digest-card__headline" style={{ color: record.digest.theme.title }}>
                        {record.digest.headline}
                      </div>

                      <div className="review-digest-card__detail">
                        {record.digest.detail}
                      </div>

                      <div className="review-digest-card__stats">
                        <span
                          className="review-digest-card__pill review-digest-card__pill--interactive"
                          onClick={(event) => {
                            event.stopPropagation();
                            openStatusDetail(record, 'pass');
                          }}
                          style={{ backgroundColor: record.digest.theme.soft, color: record.digest.theme.muted, borderColor: record.digest.theme.border }}
                        >
                          通过 {record.digest.summary.passedChecks}
                        </span>
                        <span
                          className="review-digest-card__pill review-digest-card__pill--warning review-digest-card__pill--interactive"
                          onClick={(event) => {
                            event.stopPropagation();
                            openStatusDetail(record, 'warning');
                          }}
                        >
                          警告 {record.digest.summary.warningChecks}
                        </span>
                        <span
                          className="review-digest-card__pill review-digest-card__pill--review review-digest-card__pill--interactive"
                          onClick={(event) => {
                            event.stopPropagation();
                            openStatusDetail(record, 'review');
                          }}
                        >
                          复核 {record.digest.summary.reviewChecks}
                        </span>
                        <span
                          className="review-digest-card__pill review-digest-card__pill--error review-digest-card__pill--interactive"
                          onClick={(event) => {
                            event.stopPropagation();
                            openStatusDetail(record, 'error');
                          }}
                        >
                          错误 {record.digest.summary.errorChecks}
                        </span>
                      </div>

                      <div className="review-digest-card__hint">
                        点击卡片查看完整详情，点击统计标签可按状态筛选明细
                      </div>
                    </Card>
                  </Col>
                ))}
              </Row>
            </Card>
          </Col>
        )}

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
            extra={<span style={{ color: '#8c8c8c', fontSize: 12 }}>默认收起，按需展开查看</span>}
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
                                <p style={{ margin: 0, color: '#667085', minHeight: 44 }}>{area.description}</p>
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
                <div style={{ color: '#595959', marginBottom: 12, lineHeight: 1.7 }}>
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
                    <div style={{ fontWeight: 600, color: '#44506b', marginBottom: 8 }}>证据记录</div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                      {section.evidence.map((evidence, index) => (
                        <div
                          key={`${section.key}-evidence-${index}`}
                          style={{ padding: '10px 12px', borderRadius: 10, border: '1px solid #e8eefc', background: '#f8fbff', color: '#4b6381', lineHeight: 1.65 }}
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
