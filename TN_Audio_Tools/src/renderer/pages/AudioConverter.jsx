import React, { useState } from 'react';
import { Card, Button, Upload, Form, Select, InputNumber, Space, Table, Progress } from 'antd';
import { UploadOutlined, DeleteOutlined, CloudDownloadOutlined } from '@ant-design/icons';
import '../styles/pages.css';

function AudioConverter() {
  const [files, setFiles] = useState([]);
  const [form] = Form.useForm();

  const handleFileUpload = (file) => {
    // 实现逻辑待定
    console.log('Upload audio file:', file);
  };

  const columns = [
    {
      title: '文件名',
      dataIndex: 'name',
      key: 'name'
    },
    {
      title: '原始格式',
      dataIndex: 'format',
      key: 'format'
    },
    {
      title: '目标格式',
      dataIndex: 'targetFormat',
      key: 'targetFormat'
    },
    {
      title: '进度',
      dataIndex: 'progress',
      key: 'progress',
      render: (progress) => <Progress percent={progress} />
    },
    {
      title: '操作',
      key: 'action',
      render: () => (
        <Space>
          <Button type="link" size="small" icon={<CloudDownloadOutlined />}>下载</Button>
          <Button type="link" danger size="small" icon={<DeleteOutlined />}>删除</Button>
        </Space>
      )
    }
  ];

  return (
    <div className="page-container">
      <Card title="音频格式转换">
        <Form form={form} layout="vertical">
          <Form.Item label="输出格式" required>
            <Select
              placeholder="选择输出格式"
              disabled
              options={[
                { label: 'MP3', value: 'mp3' },
                { label: 'WAV', value: 'wav' },
                { label: 'FLAC', value: 'flac' },
                { label: 'AAC', value: 'aac' },
                { label: 'OGG', value: 'ogg' }
              ]}
            />
          </Form.Item>

          <Form.Item label="比特率 (kbps)">
            <InputNumber 
              defaultValue={192}
              disabled
              min={32}
              max={320}
              style={{ width: '100%' }}
            />
          </Form.Item>

          <Form.Item label="采样率 (Hz)">
            <Select
              placeholder="选择采样率"
              disabled
              options={[
                { label: '保持原采样率', value: 'original' },
                { label: '44.1 kHz (CD)', value: '44100' },
                { label: '48 kHz', value: '48000' },
                { label: '96 kHz', value: '96000' }
              ]}
            />
          </Form.Item>

          <Form.Item label="声道">
            <Select
              placeholder="选择声道"
              disabled
              options={[
                { label: '保持原声道', value: 'original' },
                { label: '单声道', value: 'mono' },
                { label: '立体声', value: 'stereo' }
              ]}
            />
          </Form.Item>

          <Space>
            <Upload
              customRequest={({ file }) => handleFileUpload(file)}
              multiple
              accept="audio/*"
            >
              <Button type="primary" icon={<UploadOutlined />}>
                选择文件
              </Button>
            </Upload>
            <Button disabled>开始转换</Button>
          </Space>
        </Form>
      </Card>

      <Card title="转换列表" style={{ marginTop: '24px' }}>
        {files.length === 0 ? (
          <p style={{ textAlign: 'center', color: '#8c8c8c', padding: '40px 0' }}>
            暂无转换任务
          </p>
        ) : (
          <Table columns={columns} dataSource={files} rowKey="id" pagination={false} />
        )}
      </Card>

      <Card title="转换设置" style={{ marginTop: '24px' }}>
        <Space direction="vertical" style={{ width: '100%' }}>
          <div>
            <label style={{ display: 'block', marginBottom: '8px', fontSize: '12px' }}>
              默认输出位置
            </label>
            <input 
              type="text" 
              disabled
              placeholder="选择输出文件夹"
              style={{
                width: '100%',
                padding: '8px 12px',
                border: '1px solid #d9d9d9',
                borderRadius: '4px',
                backgroundColor: '#fafafa'
              }}
            />
          </div>
          <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <input type="checkbox" disabled />
            <span>转换完成后自动打开文件夹</span>
          </label>
        </Space>
      </Card>
    </div>
  );
}

export default AudioConverter;
