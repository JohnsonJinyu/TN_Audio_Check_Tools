import React, { useState } from 'react';
import { Card, Button, Upload, Form, Select, Table, Progress, Space, Checkbox, Row, Col } from 'antd';
import { UploadOutlined, DeleteOutlined, PlayCircleOutlined } from '@ant-design/icons';
import '../styles/pages.css';

function BatchProcessor() {
  const [files, setFiles] = useState([]);

  const handleFileUpload = (file) => {
    // 实现逻辑待定
    console.log('Upload files:', file);
  };

  const columns = [
    {
      title: '文件名',
      dataIndex: 'name',
      key: 'name'
    },
    {
      title: '操作',
      dataIndex: 'operation',
      key: 'operation'
    },
    {
      title: '状态',
      dataIndex: 'status',
      key: 'status'
    },
    {
      title: '进度',
      dataIndex: 'progress',
      key: 'progress',
      render: (progress) => <Progress percent={progress} />
    }
  ];

  return (
    <div className="page-container">
      <Card title="批量处理">
        <Form layout="vertical">
          <Form.Item label="操作类型" required>
            <Select
              placeholder="选择批量操作"
              disabled
              options={[
                { label: '批量格式转换', value: 'convert' },
                { label: '批量检查报告', value: 'check' },
                { label: '批量频谱分析', value: 'analyze' },
                { label: '批量编码转换', value: 'encode' }
              ]}
            />
          </Form.Item>

          <Row gutter={16}>
            <Col xs={24} md={12}>
              <Form.Item label="输出格式">
                <Select
                  placeholder="选择输出格式"
                  disabled
                  options={[
                    { label: 'MP3', value: 'mp3' },
                    { label: 'WAV', value: 'wav' },
                    { label: 'FLAC', value: 'flac' }
                  ]}
                />
              </Form.Item>
            </Col>
            <Col xs={24} md={12}>
              <Form.Item label="比特率">
                <Select
                  placeholder="选择比特率"
                  disabled
                  defaultValue="192"
                  options={[
                    { label: '128 kbps', value: '128' },
                    { label: '192 kbps', value: '192' },
                    { label: '256 kbps', value: '256' },
                    { label: '320 kbps', value: '320' }
                  ]}
                />
              </Form.Item>
            </Col>
          </Row>

          <Space style={{ marginBottom: '16px' }}>
            <Checkbox disabled>覆盖原文件</Checkbox>
            <Checkbox disabled>转换完成后删除原文件</Checkbox>
            <Checkbox disabled>自动启用多线程</Checkbox>
          </Space>

          <Space>
            <Upload
              customRequest={({ file }) => handleFileUpload(file)}
              multiple
              accept="audio/*"
            >
              <Button type="primary" icon={<UploadOutlined />}>
                添加文件
              </Button>
            </Upload>
            <Button icon={<PlayCircleOutlined />} disabled>
              开始处理
            </Button>
            <Button danger disabled>
              清空列表
            </Button>
          </Space>
        </Form>
      </Card>

      <Card title="处理队列" style={{ marginTop: '24px' }}>
        {files.length === 0 ? (
          <p style={{ textAlign: 'center', color: '#8c8c8c', padding: '40px 0' }}>
            暂无处理任务
          </p>
        ) : (
          <Table columns={columns} dataSource={files} rowKey="id" pagination={false} />
        )}
      </Card>

      <Card title="处理统计" style={{ marginTop: '24px' }}>
        <Row gutter={16}>
          <Col xs={24} sm={12} md={6}>
            <Card type="inner">
              <div style={{ textAlign: 'center' }}>
                <p style={{ color: '#8c8c8c', fontSize: '12px', marginBottom: '8px' }}>
                  总文件数
                </p>
                <p style={{ fontSize: '24px', fontWeight: 'bold' }}>0</p>
              </div>
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card type="inner">
              <div style={{ textAlign: 'center' }}>
                <p style={{ color: '#8c8c8c', fontSize: '12px', marginBottom: '8px' }}>
                  已完成
                </p>
                <p style={{ fontSize: '24px', fontWeight: 'bold', color: '#52c41a' }}>0</p>
              </div>
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card type="inner">
              <div style={{ textAlign: 'center' }}>
                <p style={{ color: '#8c8c8c', fontSize: '12px', marginBottom: '8px' }}>
                  处理中
                </p>
                <p style={{ fontSize: '24px', fontWeight: 'bold', color: '#1890ff' }}>0</p>
              </div>
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card type="inner">
              <div style={{ textAlign: 'center' }}>
                <p style={{ color: '#8c8c8c', fontSize: '12px', marginBottom: '8px' }}>
                  失败
                </p>
                <p style={{ fontSize: '24px', fontWeight: 'bold', color: '#f5222d' }}>0</p>
              </div>
            </Card>
          </Col>
        </Row>
      </Card>
    </div>
  );
}

export default BatchProcessor;
