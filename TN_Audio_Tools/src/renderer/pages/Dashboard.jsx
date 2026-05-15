import React, { useEffect, useMemo, useState } from 'react';
import { App as AntdApp, Button, Card, Col, Divider, Empty, Modal, Row, Space, Statistic, Table, Tag } from 'antd';
import {
  FileTextOutlined,
  LineChartOutlined,
  SearchOutlined,
  ArrowRightOutlined,
  FolderOpenOutlined
} from '@ant-design/icons';
import {
  clearReportHistory,
  DASHBOARD_STATS_UPDATED_EVENT,
  readDashboardData
} from '../modules/dashboard/storage';
import {
  readWordReviewHistory,
  WORD_REVIEW_HISTORY_UPDATED_EVENT
} from '../modules/reportReview/storage';
import '../styles/pages.css';

function buildDashboardSnapshot() {
  const collectionData = readDashboardData();
  const reviewHistory = readWordReviewHistory();
  const safeReviewHistory = Array.isArray(reviewHistory) ? reviewHistory : [];
  const passedReviewCount = safeReviewHistory.filter(function(item) {
    return item?.result?.reviewResult?.overallStatus === 'pass';
  }).length;

  var now = Date.now();
  var sevenDaysAgo = now - 7 * 24 * 60 * 60 * 1000;
  var recent7DaysReviews = safeReviewHistory.filter(function(item) {
    return item?.checkedAt && new Date(item.checkedAt).getTime() > sevenDaysAgo;
  });

  var passRate = safeReviewHistory.length > 0
    ? Math.round((passedReviewCount / safeReviewHistory.length) * 100)
    : 0;

  var categoryStats = { structure: { pass: 0, total: 0 }, metadata: { pass: 0, total: 0 }, timing: { pass: 0, total: 0 }, content: { pass: 0, total: 0 } };
  var structureKeys = ['tableOfContents', 'tableOfContentsPages', 'chaptersAlignment'];
  var metadataKeys = ['basicInfo', 'testItemConsistency', 'namePollution', 'engineers', 'polqa'];
  var timingKeys = ['timingAdjacentInterval', 'timingTotalSpan', 'timingDelayOrder', 'timingSidetoneDelayOrder', 'timingBgnConnectionOrder'];
  var contentKeys = ['contentLoudnessFRTrend', 'contentCurveValueCorroboration', 'contentSameCodecDiffNetwork', 'contentSameNetworkDiffCodec'];

  safeReviewHistory.forEach(function(item) {
    var checks = item?.result?.reviewResult?.checks || {};
    structureKeys.forEach(function(k) { if (checks[k]) { categoryStats.structure.total++; if (checks[k].status === 'pass') categoryStats.structure.pass++; } });
    metadataKeys.forEach(function(k) { if (checks[k]) { categoryStats.metadata.total++; if (checks[k].status === 'pass') categoryStats.metadata.pass++; } });
    timingKeys.forEach(function(k) { if (checks[k]) { categoryStats.timing.total++; if (checks[k].status === 'pass') categoryStats.timing.pass++; } });
    contentKeys.forEach(function(k) { if (checks[k]) { categoryStats.content.total++; if (checks[k].status === 'pass') categoryStats.content.pass++; } });
  });

  var recentReviews = safeReviewHistory.slice(0, 5).map(function(item) {
    return {
      id: item.id,
      reportName: item.reportName,
      checkedAt: item.checkedAt,
      status: item?.result?.reviewResult?.overallStatus || 'unknown',
      summary: item?.result?.reviewResult?.summary || {}
    };
  });

  return {
    ...collectionData,
    reviewHistoryCount: safeReviewHistory.length,
    passedReviewCount: passedReviewCount,
    passRate: passRate,
    recent7DaysCount: recent7DaysReviews.length,
    recentReviews: recentReviews,
    categoryStats: categoryStats
  };
}

function Dashboard({ onNavigate }) {
  const { message, modal } = AntdApp.useApp();
  const [dashboardData, setDashboardData] = useState(buildDashboardSnapshot);
  const [historyVisible, setHistoryVisible] = useState(false);

  useEffect(() => {
    const refreshDashboardData = () => {
      setDashboardData(buildDashboardSnapshot());
    };

    window.addEventListener(DASHBOARD_STATS_UPDATED_EVENT, refreshDashboardData);
    window.addEventListener(WORD_REVIEW_HISTORY_UPDATED_EVENT, refreshDashboardData);

    return () => {
      window.removeEventListener(DASHBOARD_STATS_UPDATED_EVENT, refreshDashboardData);
      window.removeEventListener(WORD_REVIEW_HISTORY_UPDATED_EVENT, refreshDashboardData);
    };
  }, []);

  const tools = [
    {
      title: '测试数据收集',
      description: '上传报告、checklist 和规则文件，统一收集测试数据并生成结论',
      icon: <FileTextOutlined />,
      color: '#ff7a45',
      stats: `${dashboardData.checkedReports} 份报告`,
      pageKey: 'report-checker'
    },
    {
      title: '报告审查',
      description: '查看审查范围、最近处理结果和输出文件历史',
      icon: <SearchOutlined />,
      color: '#1677ff',
      stats: dashboardData.reviewHistoryCount > 0
        ? (`${dashboardData.passedReviewCount}/${dashboardData.reviewHistoryCount} 通过 (${dashboardData.passRate}%)`)
        : '暂无审查记录',
      pageKey: 'report-review'
    },
    {
      title: '频谱分析',
      description: '实时分析音频的频谱特性，提供可视化展示',
      icon: <LineChartOutlined />,
      color: '#722ed1',
      stats: '实时分析',
      pageKey: 'spectrum'
    }
  ];

  const handleNavigate = (pageKey) => {
    if (typeof onNavigate === 'function' && pageKey) {
      onNavigate(pageKey);
    }
  };

  const openReportHistory = () => {
    setDashboardData(buildDashboardSnapshot());
    setHistoryVisible(true);
  };

  const handleQuickStatClick = (stat) => {
    if (stat.key === 'checkedReports') {
      openReportHistory();
      return;
    }

    handleNavigate(stat.pageKey);
  };

  const openHistoryOutputFolder = async (record) => {
    if (!record.outputPath) {
      message.warning('该历史记录没有输出文件可打开。');
      return;
    }

    try {
      await window.electron.testDataCollection.showOutputInFolder(record.outputPath);
    } catch (error) {
      message.error(error?.message || '打开输出目录失败');
    }
  };

  const handleClearReportHistory = () => {
    if (dashboardData.reportHistory.length === 0) {
      return;
    }

    modal.confirm({
      title: '清空数据收集历史',
      content: '清空后将删除当前本地保存的数据收集历史，且不可恢复。',
      okText: '确认清空',
      cancelText: '取消',
      okButtonProps: { danger: true },
      onOk: () => {
        clearReportHistory();
        setDashboardData(buildDashboardSnapshot());
        message.success('已清空数据收集历史');
      }
    });
  };

  const quickStats = [
    {
      key: 'checkedReports',
      title: '已收集报告',
      value: dashboardData.checkedReports,
      suffix: '份',
      prefix: <FileTextOutlined />,
      color: '#ff7a45',
      pageKey: 'report-checker'
    },
    {
      key: 'reviewHistory',
      title: '报告审查记录',
      value: dashboardData.reviewHistoryCount,
      suffix: dashboardData.passRate > 0 ? ('通过率 ' + dashboardData.passRate + '%') : '',
      prefix: <SearchOutlined />,
      color: '#1677ff',
      pageKey: 'report-review'
    },
    {
      key: 'recentActivity',
      title: '近7天审查',
      value: dashboardData.recent7DaysCount,
      suffix: '次',
      prefix: <LineChartOutlined />,
      color: '#722ed1',
      pageKey: 'report-review'
    }
  ];

  const historyColumns = useMemo(() => ([
    {
      title: '报告文件',
      dataIndex: 'reportName',
      key: 'reportName',
      ellipsis: true
    },
    {
      title: '检查时间',
      dataIndex: 'checkedAt',
      key: 'checkedAt',
      width: 180,
      render: (value) => value ? new Date(value).toLocaleString() : '-'
    },
    {
      title: '状态',
      dataIndex: 'status',
      key: 'status',
      width: 100,
      render: (status) => {
        const colorMap = {
          success: 'green',
          error: 'red'
        };

        return <Tag color={colorMap[status] || 'default'}>{status}</Tag>;
      }
    },
    {
      title: '命中项',
      dataIndex: 'matchedItems',
      key: 'matchedItems',
      width: 90
    },
    {
      title: '输出文件',
      dataIndex: 'outputName',
      key: 'outputName',
      ellipsis: true,
      render: (value) => value || '未生成'
    },
    {
      title: '操作',
      key: 'action',
      width: 120,
      render: (_, record) => (
        <Button
          size="small"
          icon={<FolderOpenOutlined />}
          disabled={!record.outputPath}
          onClick={() => openHistoryOutputFolder(record)}
        >
          打开目录
        </Button>
      )
    }
  ]), []);

  return (
    <div className="page-container dashboard">
      <Row gutter={[24, 24]}>
        {/* 欢迎信息 */}
        <Col xs={24}>
          <Card className="welcome-card" style={{
            background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            color: '#fff',
            border: 'none'
          }}>
            <h1 style={{ fontSize: '28px', marginBottom: '12px' }}>
              欢迎使用 TN Audio Toolkit
            </h1>
            <p style={{ fontSize: '16px', opacity: 0.9, marginBottom: 0 }}>
              面向音频测试场景的数据收集、报告审查与频谱分析工作台。
            </p>
          </Card>
        </Col>

        {/* 快速统计 */}
        {quickStats.map((stat) => (
          <Col key={stat.title} xs={24} sm={12} md={8}>
            <Card
              className="dashboard-stat-card"
              hoverable
              style={{ cursor: 'pointer' }}
              onClick={() => handleQuickStatClick(stat)}
            >
              <Statistic
                title={stat.title}
                value={stat.value}
                prefix={stat.prefix}
                suffix={stat.suffix ? <span style={{ fontSize: 14, color: stat.color }}>{stat.suffix}</span> : undefined}
                valueStyle={{ color: stat.color }}
              />
            </Card>
          </Col>
        ))}

        {/* 工具卡片 */}
        <Col xs={24}>
          <h2 style={{ marginBottom: '16px', fontSize: '20px', fontWeight: 'bold' }}>
            主要功能
          </h2>
        </Col>

        {tools.map((tool, index) => (
          <Col key={index} xs={24} sm={12} md={12} lg={8}>
            <Card 
              className="tool-card"
              hoverable
              style={{ height: '100%', cursor: 'pointer' }}
              onClick={() => handleNavigate(tool.pageKey)}
            >
              <div 
                className="tool-icon"
                style={{ 
                  color: tool.color,
                  fontSize: '32px',
                  marginBottom: '12px'
                }}
              >
                {tool.icon}
              </div>
              <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '8px' }}>
                {tool.title}
              </h3>
              <p style={{ 
                fontSize: '12px', 
                color: 'var(--text-light)',
                marginBottom: '12px',
                minHeight: '36px'
              }}>
                {tool.description}
              </p>
              <Tag color={tool.color} style={{ marginBottom: '12px' }}>
                {tool.stats}
              </Tag>
              <div style={{ marginTop: '12px' }}>
                <Button 
                  type="primary" 
                  size="small"
                  style={{ width: '100%' }}
                  onClick={(event) => {
                    event.stopPropagation();
                    handleNavigate(tool.pageKey);
                  }}
                >
                  进入 <ArrowRightOutlined />
                </Button>
              </div>
            </Card>
          </Col>
        ))}

        {/* 最近审查活动 */}
        <Col xs={24}>
          <Divider />
          <h2 style={{ marginBottom: '16px', fontSize: '20px', fontWeight: 'bold' }}>
            最近审查活动
          </h2>
          <Card>
            {dashboardData.recentReviews && dashboardData.recentReviews.length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {dashboardData.recentReviews.map(function(record) {
                  var statusColorMap = { pass: 'success', warning: 'warning', review: 'processing', error: 'error' };
                  var statusTextMap = { pass: '通过', warning: '警告', review: '复核', error: '错误' };
                  return (
                    <div
                      key={record.id}
                      style={{
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        padding: '10px 16px',
                        borderRadius: 10,
                        border: '1px solid var(--border-color)',
                        background: 'var(--surface-color)'
                      }}
                    >
                      <div>
                        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 2 }}>{record.reportName}</div>
                        <div style={{ fontSize: 12, color: 'var(--text-light)' }}>
                          {record.checkedAt ? new Date(record.checkedAt).toLocaleString() : '-'}
                          {' · '}通过 {record.summary.passedChecks || 0} / 警告 {record.summary.warningChecks || 0} / 复核 {record.summary.reviewChecks || 0} / 错误 {record.summary.errorChecks || 0}
                        </div>
                      </div>
                      <Tag color={statusColorMap[record.status] || 'default'}>
                        {statusTextMap[record.status] || record.status}
                      </Tag>
                    </div>
                  );
                })}
              </div>
            ) : (
              <p style={{ color: 'var(--text-light)', textAlign: 'center', margin: '40px 0' }}>
                暂无审查记录 — 前往"报告审查"开始使用
              </p>
            )}
          </Card>
        </Col>

        {/* 快速提示 */}
        <Col xs={24}>
          <Card 
            className="dashboard-tip-card"
            type="inner"
            title="💡 快速提示"
          >
            <Space direction="vertical" style={{ width: '100%' }}>
              <p>• 先进入“测试数据收集”上传报告、checklist 和规则文件</p>
              <p>• 在“报告审查”面板查看最近批次结果和输出目录</p>
              <p>• 点击 Settings（设置）配置应用偏好</p>
            </Space>
          </Card>
        </Col>
      </Row>

      <Modal
        title="测试数据收集历史记录"
        open={historyVisible}
        onCancel={() => setHistoryVisible(false)}
        width={900}
        footer={[
          <Button
            key="clear-history"
            danger
            disabled={dashboardData.reportHistory.length === 0}
            onClick={handleClearReportHistory}
          >
            清空记录
          </Button>,
          <Button key="go-report-checker" onClick={() => {
            setHistoryVisible(false);
            handleNavigate('report-checker');
          }}>
            前往测试数据收集
          </Button>,
          <Button key="close-history" type="primary" onClick={() => setHistoryVisible(false)}>
            关闭
          </Button>
        ]}
      >
        {dashboardData.reportHistory.length > 0 ? (
          <Table
            className="dashboard-history-table"
            columns={historyColumns}
            dataSource={dashboardData.reportHistory}
            rowKey="id"
            pagination={{ pageSize: 6 }}
            scroll={{ x: 820 }}
          />
        ) : (
          <Empty description="暂无数据收集历史记录" style={{ margin: '24px 0' }} />
        )}
      </Modal>
    </div>
  );
}

export default Dashboard;
