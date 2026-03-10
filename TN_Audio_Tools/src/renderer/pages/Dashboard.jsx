import React from 'react';
import { Card, Row, Col, Statistic, Button, Space, Tag, Divider } from 'antd';
import {
  FileTextOutlined,
  AudioOutlined,
  LineChartOutlined,
  SwapOutlined,
  ArrowRightOutlined
} from '@ant-design/icons';
import '../styles/pages.css';

function Dashboard() {
  const tools = [
    {
      title: '音频报告检查',
      description: '检查和验证音频测试报告的完整性和有效性',
      icon: <FileTextOutlined />,
      color: '#ff7a45',
      stats: '0 个报告'
    },
    {
      title: '音频播放器',
      description: '播放各种格式的音频文件，并进行基础分析',
      icon: <AudioOutlined />,
      color: '#13c2c2',
      stats: '支持多种格式'
    },
    {
      title: '频谱分析',
      description: '实时分析音频的频谱特性，提供可视化展示',
      icon: <LineChartOutlined />,
      color: '#722ed1',
      stats: '实时分析'
    },
    {
      title: '音频转换',
      description: '转换音频格式，调整采样率和码率',
      icon: <SwapOutlined />,
      color: '#faad14',
      stats: '多格式支持'
    }
  ];

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
        <Col xs={24} sm={12} md={6}>
          <Card>
            <Statistic
              title="已处理音频"
              value={0}
              prefix={<AudioOutlined />}
              valueStyle={{ color: '#1890ff' }}
            />
          </Card>
        </Col>
        <Col xs={24} sm={12} md={6}>
          <Card>
            <Statistic
              title="已检查报告"
              value={0}
              prefix={<FileTextOutlined />}
              valueStyle={{ color: '#ff7a45' }}
            />
          </Card>
        </Col>
        <Col xs={24} sm={12} md={6}>
          <Card>
            <Statistic
              title="分析成功"
              value={0}
              prefix={<LineChartOutlined />}
              valueStyle={{ color: '#722ed1' }}
            />
          </Card>
        </Col>
        <Col xs={24} sm={12} md={6}>
          <Card>
            <Statistic
              title="转换完成"
              value={0}
              prefix={<SwapOutlined />}
              valueStyle={{ color: '#faad14' }}
            />
          </Card>
        </Col>

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
    </div>
  );
}

export default Dashboard;
