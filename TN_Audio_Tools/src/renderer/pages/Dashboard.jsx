import React, { useMemo, useState } from 'react';
import { Card, Row, Col, Statistic, Button, Space, Tag, Divider, Modal, Table, Empty, message } from 'antd';
import {
  FileTextOutlined,
  AudioOutlined,
  LineChartOutlined,
  SwapOutlined,
  ArrowRightOutlined,
  FolderOpenOutlined
} from '@ant-design/icons';
import { clearReportHistory, readDashboardData } from '../modules/dashboard/storage';
import '../styles/pages.css';

function Dashboard({ onNavigate }) {
  const [dashboardData, setDashboardData] = useState(readDashboardData);
  const [historyVisible, setHistoryVisible] = useState(false);
  const tools = [
    {
      title: '音频报告检查',
      description: '检查和验证音频测试报告的完整性和有效性',
      icon: <FileTextOutlined />,
      color: '#ff7a45',
      stats: `${dashboardData.checkedReports} 个报告`,
      pageKey: 'report-checker'
    },
    {
      title: '音频播放器',
      description: '播放各种格式的音频文件，并进行基础分析',
      icon: <AudioOutlined />,
      color: '#13c2c2',
      stats: '支持多种格式',
      pageKey: 'audio-player'
    },
    {
      title: '频谱分析',
      description: '实时分析音频的频谱特性，提供可视化展示',
      icon: <LineChartOutlined />,
      color: '#722ed1',
      stats: '实时分析',
      pageKey: 'spectrum'
    },
    {
      title: '音频转换',
      description: '转换音频格式，调整采样率和码率',
      icon: <SwapOutlined />,
      color: '#faad14',
      stats: '多格式支持',
      pageKey: 'converter'
    }
  ];

  const handleNavigate = (pageKey) => {
    if (typeof onNavigate === 'function' && pageKey) {
      onNavigate(pageKey);
    }
  };

  const openReportHistory = () => {
    setDashboardData(readDashboardData());
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
      await window.electron.reportChecker.showOutputInFolder(record.outputPath);
    } catch (error) {
      message.error(error?.message || '打开输出目录失败');
    }
  };

  const handleClearReportHistory = () => {
    if (dashboardData.reportHistory.length === 0) {
      return;
    }

    Modal.confirm({
      title: '清空检查历史记录',
      content: '清空后将删除当前本地保存的报告检查历史，且不可恢复。',
      okText: '确认清空',
      cancelText: '取消',
      okButtonProps: { danger: true },
      onOk: () => {
        clearReportHistory();
        setDashboardData(readDashboardData());
        message.success('已清空检查历史记录');
      }
    });
  };

  const quickStats = [
    {
      key: 'processedAudio',
      title: '已处理音频',
      value: dashboardData.processedAudio,
      prefix: <AudioOutlined />,
      color: '#1890ff',
      pageKey: 'audio-player'
    },
    {
      key: 'checkedReports',
      title: '已检查报告',
      value: dashboardData.checkedReports,
      prefix: <FileTextOutlined />,
      color: '#ff7a45',
      pageKey: 'report-checker'
    },
    {
      key: 'analysisSuccess',
      title: '分析成功',
      value: dashboardData.analysisSuccess,
      prefix: <LineChartOutlined />,
      color: '#722ed1',
      pageKey: 'spectrum'
    },
    {
      key: 'conversionsCompleted',
      title: '转换完成',
      value: dashboardData.conversionsCompleted,
      prefix: <SwapOutlined />,
      color: '#faad14',
      pageKey: 'converter'
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
              一站式音频处理和分析解决方案。选择下方的工具开始使用。
            </p>
          </Card>
        </Col>

        {/* 快速统计 */}
        {quickStats.map((stat) => (
          <Col key={stat.title} xs={24} sm={12} md={6}>
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
          <Col key={index} xs={24} sm={12} md={12} lg={6}>
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
                color: '#8c8c8c',
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

        {/* 最近使用 */}
        <Col xs={24}>
          <Divider />
          <h2 style={{ marginBottom: '16px', fontSize: '20px', fontWeight: 'bold' }}>
            最近使用
          </h2>
          <Card>
            <p style={{ color: '#8c8c8c', textAlign: 'center', margin: '40px 0' }}>
              暂无最近使用的记录
            </p>
          </Card>
        </Col>

        {/* 快速提示 */}
        <Col xs={24}>
          <Card 
            type="inner"
            title="💡 快速提示"
            style={{ backgroundColor: '#f6f8fb', border: '1px solid #e6f7ff' }}
          >
            <Space direction="vertical" style={{ width: '100%' }}>
              <p>• 使用 Ctrl+I 快速打开音频导入对话框</p>
              <p>• 在左侧菜单中选择相应工具开始使用</p>
              <p>• 点击 Settings（设置）配置应用偏好</p>
            </Space>
          </Card>
        </Col>
      </Row>

      <Modal
        title="已检查报告历史记录"
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
            前往报告检查
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
          <Empty description="暂无检查历史记录" style={{ margin: '24px 0' }} />
        )}
      </Modal>
    </div>
  );
}

export default Dashboard;
