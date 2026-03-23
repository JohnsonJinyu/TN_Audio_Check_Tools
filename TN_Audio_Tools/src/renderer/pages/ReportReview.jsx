import React, { useMemo, useState } from 'react';
import { Card, Row, Col, Statistic, Button, Space, Tag, Table, Empty, message } from 'antd';
import { SearchOutlined, FileTextOutlined, FolderOpenOutlined, CheckCircleOutlined } from '@ant-design/icons';
import { readDashboardData } from '../modules/dashboard/storage';
import '../styles/pages.css';

function ReportReview({ onNavigate }) {
  const [dashboardData] = useState(readDashboardData);

  const successRate = dashboardData.checkedReports > 0
    ? Math.round((dashboardData.analysisSuccess / dashboardData.checkedReports) * 100)
    : 0;

  const reviewAreas = [
    {
      title: '文档完整性审查',
      description: '核对 Word 报告章节、附件与必要说明是否齐全。',
      tag: 'Word 审查',
      color: 'blue'
    },
    {
      title: '曲线章节定位',
      description: '检查响度、频响等关键章节是否被正确识别并纳入结论。',
      tag: '曲线证据',
      color: 'purple'
    },
    {
      title: '跨报告一致性复核',
      description: '对可比样本进行交叉复核，识别异常差异和补充动作。',
      tag: '一致性',
      color: 'gold'
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
      title: '处理时间',
      dataIndex: 'checkedAt',
      key: 'checkedAt',
      width: 180,
      render: (value) => value ? new Date(value).toLocaleString() : '-'
    },
    {
      title: '状态',
      dataIndex: 'status',
      key: 'status',
      width: 110,
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
          onClick={() => openOutputFolder(record)}
        >
          打开目录
        </Button>
      )
    }
  ]), []);

  const openOutputFolder = async (record) => {
    if (!record.outputPath) {
      message.warning('该记录没有输出文件可打开。');
      return;
    }

    try {
      await window.electron.testDataCollection.showOutputInFolder(record.outputPath);
    } catch (error) {
      message.error(error?.message || '打开输出目录失败');
    }
  };

  return (
    <div className="page-container">
      <Row gutter={[24, 24]}>
        <Col xs={24}>
          <Card className="report-checker-card" style={{ borderColor: '#d6e4ff' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 16, alignItems: 'center', flexWrap: 'wrap' }}>
              <div>
                <h2 style={{ marginBottom: 8, fontSize: 24 }}>报告审查面板</h2>
                <p style={{ marginBottom: 0, color: '#667085' }}>
                  这里聚合当前工具集中的审查关注点，并提供最近一批处理结果的快速回看入口。
                </p>
              </div>
              <Button type="primary" icon={<FileTextOutlined />} onClick={() => onNavigate?.('report-checker')}>
                前往测试数据收集
              </Button>
            </div>
          </Card>
        </Col>

        <Col xs={24} sm={8}>
          <Card className="dashboard-stat-card" hoverable>
            <Statistic title="累计处理报告" value={dashboardData.checkedReports} prefix={<FileTextOutlined />} valueStyle={{ color: '#ff7a45' }} />
          </Card>
        </Col>
        <Col xs={24} sm={8}>
          <Card className="dashboard-stat-card" hoverable>
            <Statistic title="成功处理" value={dashboardData.analysisSuccess} prefix={<CheckCircleOutlined />} valueStyle={{ color: '#52c41a' }} />
          </Card>
        </Col>
        <Col xs={24} sm={8}>
          <Card className="dashboard-stat-card" hoverable>
            <Statistic title="成功率" value={successRate} suffix="%" prefix={<SearchOutlined />} valueStyle={{ color: '#1677ff' }} />
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
          <Card className="report-checker-card" title="最近审查记录">
            {dashboardData.reportHistory.length > 0 ? (
              <Table
                columns={historyColumns}
                dataSource={dashboardData.reportHistory}
                rowKey="id"
                pagination={{ pageSize: 6 }}
                scroll={{ x: 860 }}
              />
            ) : (
              <Empty description="暂无可审查的历史记录" style={{ margin: '24px 0' }} />
            )}
          </Card>
        </Col>
      </Row>
    </div>
  );
}

export default ReportReview;