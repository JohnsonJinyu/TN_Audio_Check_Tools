import React, { useState } from 'react';
import { Card, Button, Upload, Table, Space, Tag, Modal, message } from 'antd';
import { UploadOutlined, DeleteOutlined, CheckCircleOutlined } from '@ant-design/icons';
import '../styles/pages.css';

function ReportChecker() {
  const [files, setFiles] = useState([]);

  const handleUpload = (file) => {
    // 实现逻辑待定
    console.log('Upload file:', file);
  };

  const columns = [
    {
      title: '文件名',
      dataIndex: 'name',
      key: 'name',
      render: (text) => <span>{text}</span>
    },
    {
      title: '状态',
      dataIndex: 'status',
      key: 'status',
      render: (status) => {
        const colors = {
          success: 'green',
          error: 'red',
          pending: 'blue'
        };
        return <Tag color={colors[status] || 'default'}>{status}</Tag>;
      }
    },
    {
      title: '检查项',
      dataIndex: 'items',
      key: 'items',
      render: (items) => <span>{items || 0}</span>
    },
    {
      title: '操作',
      key: 'action',
      render: (_, record) => (
        <Space>
          <Button type="primary" size="small">详情</Button>
          <Button danger size="small" icon={<DeleteOutlined />}>删除</Button>
        </Space>
      )
    }
  ];

  return (
    <div className="page-container">
      <Card 
        title="音频测试报告检查"
        extra={
          <Upload
            customRequest={({ file }) => handleUpload(file)}
            multiple
            accept=".xlsx,.xls,.csv,.pdf"
          >
            <Button type="primary" icon={<UploadOutlined />}>
              上传报告
            </Button>
          </Upload>
        }
      >
        <p style={{ marginBottom: '24px', color: '#8c8c8c' }}>
          导入音频测试报告，系统将自动检查报告的完整性、有效性和数据准确性。
        </p>

        {files.length === 0 ? (
          <div style={{
            textAlign: 'center',
            padding: '60px 20px',
            backgroundColor: '#fafafa',
            borderRadius: '6px',
            border: '2px dashed #d9d9d9'
          }}>
            <UploadOutlined style={{ fontSize: '48px', color: '#bfbfbf', marginBottom: '16px' }} />
            <p style={{ fontSize: '16px', color: '#262626', marginBottom: '8px' }}>
              拖拽报告文件到此处，或点击上方按钮上传
            </p>
            <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
              支持格式: Excel (.xlsx, .xls), CSV, PDF
            </p>
          </div>
        ) : (
          <Table
            columns={columns}
            dataSource={files}
            rowKey="id"
            pagination={{ pageSize: 10 }}
          />
        )}
      </Card>

      <Card title="检查规则" style={{ marginTop: '24px' }}>
        <ul style={{ paddingLeft: '20px' }}>
          <li>待实现</li>
          <li>待实现</li>
          <li>待实现</li>
        </ul>
      </Card>
    </div>
  );
}

export default ReportChecker;
