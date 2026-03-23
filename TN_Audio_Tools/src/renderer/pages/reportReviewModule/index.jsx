import React, { useMemo, useState } from 'react';
import { Alert, Button, Card, Col, Empty, message, Modal, Row, Space, Statistic, Table, Tag } from 'antd';
import { CheckCircleOutlined, FileTextOutlined, SearchOutlined } from '@ant-design/icons';
import { clearWordReviewHistory, readWordReviewHistory, recordWordReviewResult } from '../../modules/reportReview/storage';
import { reviewAreas } from './constants';
import { createReviewHistoryColumns } from './reviewHistoryColumns';
import ReviewResultContent from './ReviewResultContent';
import '../../styles/pages.css';

export default function ReportReviewPage() {
  const [wordReviewHistory, setWordReviewHistory] = useState(() => readWordReviewHistory() || []);
  const [reviewLoading, setReviewLoading] = useState(false);
  const [selectedReportPath, setSelectedReportPath] = useState(null);
  const [detailModalVisible, setDetailModalVisible] = useState(false);
  const [detailModalData, setDetailModalData] = useState(null);

  const safeWordReviewHistory = Array.isArray(wordReviewHistory) ? wordReviewHistory : [];

  const reviewStats = useMemo(() => {
    const total = safeWordReviewHistory.length;
    const passed = safeWordReviewHistory.filter((item) => item?.result?.reviewResult?.overallStatus === 'pass').length;
    const warning = safeWordReviewHistory.filter((item) => item?.result?.reviewResult?.overallStatus === 'warning').length;
    const error = safeWordReviewHistory.filter((item) => ['review', 'error'].includes(item?.result?.reviewResult?.overallStatus)).length;

    return { total, passed, warning, error };
  }, [safeWordReviewHistory]);

  const historyColumns = useMemo(() => createReviewHistoryColumns((record) => {
    setDetailModalData(record);
    setDetailModalVisible(true);
  }), []);

  const handleSelectFile = async () => {
    try {
      const result = await window.electron.ipcRenderer.invoke('dialog:open-file', {
        filters: [
          { name: 'Word 文档', extensions: ['doc', 'docx'] },
          { name: '所有文件', extensions: ['*'] }
        ],
        properties: ['openFile']
      });

      if (!result.canceled && result.filePath && result.filePath.length > 0) {
        const filePath = result.filePath[0];
        setSelectedReportPath(filePath);
        message.success(`已选择文件：${filePath.split('\\').pop()}`);
      }
    } catch (error) {
      message.error(error?.message || '选择文件失败');
      setSelectedReportPath(null);
    }
  };

  const performReview = async () => {
    if (!selectedReportPath) {
      return;
    }

    setReviewLoading(true);
    try {
      const result = await window.electron.reportReview.reviewWordReport({
        reportPath: selectedReportPath
      });

      recordWordReviewResult(selectedReportPath, result);
      setWordReviewHistory(readWordReviewHistory());
      message.success('报告审查成功，结果已保存到历史记录');
    } catch (error) {
      message.error(error?.message || 'Word 审查失败');
    } finally {
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
            bodyStyle={{ padding: '10px 24px' }}
          >
            <Row gutter={[24, 8]} align="middle" justify="center">
              <Col xs={0} lg={4} />
              <Col xs={24} lg={16}>
                <div style={{ width: '100%', maxWidth: 760, margin: '0 auto', display: 'flex', flexDirection: 'column', gap: 12 }}>
                  <div
                    onClick={handleSelectFile}
                    onDragOver={(event) => {
                      event.preventDefault();
                      event.currentTarget.style.backgroundColor = '#e6f7ff';
                    }}
                    onDragLeave={(event) => {
                      event.preventDefault();
                      event.currentTarget.style.backgroundColor = '#f5f7fa';
                    }}
                    onDrop={(event) => {
                      event.preventDefault();
                      event.currentTarget.style.backgroundColor = '#f5f7fa';
                      const files = event.dataTransfer.files;
                      if (files && files.length > 0 && files[0].path) {
                        const filePath = files[0].path;
                        setSelectedReportPath(filePath);
                        message.success(`已选择文件：${filePath.split('\\').pop()}`);
                      }
                    }}
                    style={{
                      border: '2px dashed #1677ff',
                      borderRadius: 16,
                      padding: '28px 28px',
                      textAlign: 'center',
                      backgroundColor: '#f5f7fa',
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
                    <div style={{ fontSize: 20, fontWeight: 600, color: '#262626', marginBottom: 6 }}>点击或拖拽选择文件</div>
                    <div style={{ fontSize: 13, color: '#8c8c8c' }}>支持 .doc 和 .docx 格式</div>
                  </div>

                  {selectedReportPath && (
                    <div
                      style={{
                        width: '100%',
                        padding: '12px 16px',
                        backgroundColor: '#e6f7ff',
                        border: '1px solid #91d5ff',
                        borderRadius: 12,
                        fontSize: 13
                      }}
                    >
                      <div style={{ color: '#0050b3', marginBottom: 6, fontWeight: 600 }}>✓ 已选择文件</div>
                      <div style={{ color: '#1677ff', wordBreak: 'break-all', fontSize: 18, lineHeight: 1.35 }}>
                        {selectedReportPath.split('\\').pop()}
                      </div>
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
                    disabled={!selectedReportPath}
                    style={{ width: '100%', height: 52, borderRadius: 14, fontSize: 18, fontWeight: 600 }}
                    block
                  >
                    开始审查
                  </Button>
                </div>
              </Col>
            </Row>
          </Card>
        </Col>

        <Col xs={24} sm={8}>
          <Card className="dashboard-stat-card" hoverable>
            <Statistic title="审查报告数" value={reviewStats.total} prefix={<FileTextOutlined />} valueStyle={{ color: '#ff7a45' }} />
          </Card>
        </Col>
        <Col xs={24} sm={8}>
          <Card className="dashboard-stat-card" hoverable>
            <Statistic title="全部通过" value={reviewStats.passed} prefix={<CheckCircleOutlined />} valueStyle={{ color: '#52c41a' }} />
          </Card>
        </Col>
        <Col xs={24} sm={8}>
          <Card className="dashboard-stat-card" hoverable>
            <Statistic title="待处理（警告/错误）" value={reviewStats.warning + reviewStats.error} prefix={<SearchOutlined />} valueStyle={{ color: reviewStats.warning + reviewStats.error > 0 ? '#faad14' : '#1677ff' }} />
          </Card>
        </Col>

        {reviewAreas.map((area) => (
          <Col key={area.title} xs={24} md={8}>
            <Card className="tool-card" hoverable style={{ height: '100%' }}>
              <Space direction="vertical" size={12} style={{ width: '100%' }}>
                <Tag color={area.color} style={{ width: 'fit-content' }}>{area.tag}</Tag>
                <h3 style={{ margin: 0, fontSize: 18 }}>{area.title}</h3>
                <p style={{ margin: 0, color: '#667085', minHeight: 44 }}>{area.description}</p>
              </Space>
            </Card>
          </Col>
        ))}

        <Col xs={24}>
          <Card
            className="report-checker-card"
            title="最近审查记录"
            extra={safeWordReviewHistory.length > 0 ? (
              <Button
                danger
                size="small"
                onClick={() => {
                  Modal.confirm({
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
              <Table columns={historyColumns} dataSource={safeWordReviewHistory} rowKey="id" pagination={{ pageSize: 6 }} scroll={{ x: 860 }} />
            ) : (
              <Empty description="暂无审查记录" style={{ margin: '24px 0' }} />
            )}
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
    </div>
  );
}
