import React, { useState } from 'react';
import { Card, Button, Upload, Table, Space, Tag, Modal, message } from 'antd';
import { UploadOutlined, DeleteOutlined, CheckCircleOutlined } from '@ant-design/icons';
import '../styles/pages.css';

function ReportChecker() {
  const [files, setFiles] = useState([]);
  const [ruleFile, setRuleFile] = useState(null);
  const [checklistFile, setChecklistFile] = useState(null);

  const handleUpload = (file, target, onSuccess) => {
    if (target === 'report') {
      const newItem = {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        name: file.name,
        status: 'pending',
        items: 0
      };
      setFiles((prev) => [newItem, ...prev]);
      message.success(`已添加报告: ${file.name}`);
    }

    if (target === 'rules') {
      setRuleFile(file);
      message.success(`已上传规则: ${file.name}`);
    }

    if (target === 'checklist') {
      setChecklistFile(file);
      message.success(`已上传 checklist: ${file.name}`);
    }

    if (onSuccess) {
      setTimeout(() => onSuccess('ok'), 0);
    }
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
            customRequest={({ file, onSuccess }) => handleUpload(file, 'report', onSuccess)}
            multiple
            accept=".doc,.docx"
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
          <Upload.Dragger
            customRequest={({ file, onSuccess }) => handleUpload(file, 'report', onSuccess)}
            multiple
            accept=".doc,.docx"
            showUploadList={false}
            style={{ padding: '24px' }}
          >
            <UploadOutlined style={{ fontSize: '48px', color: '#bfbfbf', marginBottom: '16px' }} />
            <p style={{ fontSize: '16px', color: '#262626', marginBottom: '8px' }}>
              拖拽报告文件到此处，或点击上方按钮上传
            </p>
            <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
              支持格式: Word (.doc, .docx)
            </p>
          </Upload.Dragger>
        ) : (
          <Table
            columns={columns}
            dataSource={files}
            rowKey="id"
            pagination={{ pageSize: 10 }}
          />
        )}
      </Card>

      <Card
        title="上传规则"
        style={{ marginTop: '24px' }}
        extra={
          <Upload
            customRequest={({ file, onSuccess }) => handleUpload(file, 'rules', onSuccess)}
            accept=".json,.xlsx,.xls,.csv"
          >
            <Button type="primary" icon={<UploadOutlined />}>
              上传规则
            </Button>
          </Upload>
        }
      >
        <Upload.Dragger
          customRequest={({ file, onSuccess }) => handleUpload(file, 'rules', onSuccess)}
          accept=".json,.xlsx,.xls,.csv"
          showUploadList={false}
          style={{ padding: '12px 16px', minHeight: '120px' }}
        >
          <UploadOutlined style={{ fontSize: '28px', color: '#bfbfbf', marginBottom: '8px' }} />
          <p style={{ fontSize: '14px', color: '#262626', marginBottom: '4px' }}>
            拖拽规则文件到此处，或点击上传
          </p>
          <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
            支持格式: JSON, Excel, CSV
          </p>
        </Upload.Dragger>
        {ruleFile && (
          <p style={{ marginTop: '12px', color: '#595959' }}>
            已选择: {ruleFile.name}
          </p>
        )}
      </Card>

      <Card
        title="上传 checklist"
        style={{ marginTop: '24px' }}
        extra={
          <Upload
            customRequest={({ file, onSuccess }) => handleUpload(file, 'checklist', onSuccess)}
            accept=".json,.xlsx,.xls,.csv"
          >
            <Button type="primary" icon={<UploadOutlined />}>
              上传 checklist
            </Button>
          </Upload>
        }
      >
        <Upload.Dragger
          customRequest={({ file, onSuccess }) => handleUpload(file, 'checklist', onSuccess)}
          accept=".json,.xlsx,.xls,.csv"
          showUploadList={false}
          style={{ padding: '12px 16px', minHeight: '120px' }}
        >
          <UploadOutlined style={{ fontSize: '28px', color: '#bfbfbf', marginBottom: '8px' }} />
          <p style={{ fontSize: '14px', color: '#262626', marginBottom: '4px' }}>
            拖拽 checklist 文件到此处，或点击上传
          </p>
          <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
            支持格式: JSON, Excel, CSV
          </p>
        </Upload.Dragger>
        {checklistFile && (
          <p style={{ marginTop: '12px', color: '#595959' }}>
            已选择: {checklistFile.name}
          </p>
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
