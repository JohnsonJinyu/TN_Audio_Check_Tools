import React, { useState } from 'react';
import { Card, Button, Slider, Space, Upload, List, Tag } from 'antd';
import { 
  UploadOutlined, 
  PlayCircleOutlined, 
  PauseCircleOutlined,
  DeleteOutlined
} from '@ant-design/icons';
import '../styles/pages.css';

function AudioPlayer() {
  const [playlist, setPlaylist] = useState([]);
  const [isPlaying, setIsPlaying] = useState(false);

  const handleFileUpload = (file) => {
    // 实现逻辑待定
    console.log('Upload audio file:', file);
  };

  return (
    <div className="page-container">
      <Card 
        title="音频播放器"
        extra={
          <Upload
            customRequest={({ file }) => handleFileUpload(file)}
            multiple
            accept="audio/*"
          >
            <Button type="primary" icon={<UploadOutlined />}>
              添加音频
            </Button>
          </Upload>
        }
      >
        <div style={{ marginBottom: '24px' }}>
          <div style={{
            backgroundColor: '#f5f5f5',
            borderRadius: '6px',
            padding: '20px',
            textAlign: 'center'
          }}>
            <div style={{ 
              fontSize: '48px', 
              marginBottom: '16px',
              color: '#13c2c2'
            }}>
              🎵
            </div>
            <p style={{ marginBottom: '16px', color: '#8c8c8c' }}>
              当前未选择音频文件
            </p>
            <Space>
              <Button 
                type="primary" 
                size="large"
                icon={isPlaying ? <PauseCircleOutlined /> : <PlayCircleOutlined />}
                disabled
              >
                {isPlaying ? '暂停' : '播放'}
              </Button>
            </Space>
          </div>

          <div style={{ marginTop: '20px' }}>
            <p style={{ marginBottom: '8px', color: '#8c8c8c', fontSize: '12px' }}>
              进度: 00:00 / 00:00
            </p>
            <Slider 
              defaultValue={0}
              disabled
              style={{ marginBottom: '20px' }}
            />
          </div>

          <div style={{
            display: 'flex',
            gap: '12px',
            marginTop: '16px'
          }}>
            <span style={{ fontSize: '12px', color: '#8c8c8c' }}>音量</span>
            <Slider 
              defaultValue={70}
              style={{ width: '200px' }}
              disabled
            />
          </div>
        </div>
      </Card>

      <Card title="播放列表" style={{ marginTop: '24px' }}>
        {playlist.length === 0 ? (
          <p style={{ textAlign: 'center', color: '#8c8c8c', padding: '40px 0' }}>
            暂无音频文件
          </p>
        ) : (
          <List
            dataSource={playlist}
            renderItem={(item) => (
              <List.Item
                actions={[
                  <Button type="link" size="small">播放</Button>,
                  <Button type="link" danger size="small" icon={<DeleteOutlined />}>删除</Button>
                ]}
              >
                <List.Item.Meta
                  title={item.name}
                  description={item.duration}
                />
              </List.Item>
            )}
          />
        )}
      </Card>

      <Card title="音频信息" style={{ marginTop: '24px' }}>
        <p style={{ color: '#8c8c8c', textAlign: 'center', padding: '40px 0' }}>
          选择一个音频文件查看详细信息
        </p>
      </Card>
    </div>
  );
}

export default AudioPlayer;
