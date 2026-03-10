import React, { useState } from 'react';
import { Card, Button, Upload, Row, Col, Select, Space, Switch } from 'antd';
import { UploadOutlined, DownloadOutlined } from '@ant-design/icons';
import '../styles/pages.css';

function SpectrumAnalyzer() {
  const [audioFile, setAudioFile] = useState(null);

  const handleFileUpload = (file) => {
    // 实现逻辑待定
    console.log('Upload audio file:', file);
  };

  return (
    <div className="page-container">
      <Card 
        title="音频频谱分析"
        extra={
          <Upload
            customRequest={({ file }) => handleFileUpload(file)}
            accept="audio/*"
          >
            <Button type="primary" icon={<UploadOutlined />}>
              选择音频
            </Button>
          </Upload>
        }
      >
        <Row gutter={[16, 16]}>
          {/* 控制面板 */}
          <Col xs={24} md={6}>
            <Card type="inner" title="分析设置">
              <Space direction="vertical" style={{ width: '100%' }}>
                <div>
                  <label style={{ display: 'block', marginBottom: '8px', fontSize: '12px' }}>
                    分析类型
                  </label>
                  <Select 
                    defaultValue="fft"
                    style={{ width: '100%' }}
                    disabled
                    options={[
                      { label: 'FFT 分析', value: 'fft' },
                      { label: '小波变换', value: 'wavelet' },
                      { label: '傅里叶变换', value: 'fourier' }
                    ]}
                  />
                </div>

                <div>
                  <label style={{ display: 'block', marginBottom: '8px', fontSize: '12px' }}>
                    窗口函数
                  </label>
                  <Select 
                    defaultValue="hann"
                    style={{ width: '100%' }}
                    disabled
                    options={[
                      { label: 'Hann', value: 'hann' },
                      { label: 'Hamming', value: 'hamming' },
                      { label: 'Blackman', value: 'blackman' }
                    ]}
                  />
                </div>

                <div>
                  <label style={{ display: 'block', marginBottom: '8px', fontSize: '12px' }}>
                    频率范围 (Hz)
                  </label>
                  <Select 
                    defaultValue="0-20000"
                    style={{ width: '100%' }}
                    disabled
                    options={[
                      { label: '0 - 20 kHz (全音程)', value: '0-20000' },
                      { label: '0 - 10 kHz', value: '0-10000' },
                      { label: '20 Hz - 20 kHz (人类听觉)', value: '20-20000' }
                    ]}
                  />
                </div>

                <div>
                  <label style={{ display: 'block', marginBottom: '8px', fontSize: '12px' }}>
                    时域分析
                  </label>
                  <Switch disabled />
                </div>

                <Button type="primary" block disabled>
                  开始分析
                </Button>
              </Space>
            </Card>
          </Col>

          {/* 频谱显示 */}
          <Col xs={24} md={18}>
            <Card type="inner" title="频谱图">
              <div style={{
                backgroundColor: '#f5f5f5',
                borderRadius: '6px',
                padding: '40px',
                textAlign: 'center',
                height: '400px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#8c8c8c'
              }}>
                <div>
                  <div style={{ fontSize: '48px', marginBottom: '16px' }}>📊</div>
                  <p>选择音频文件开始分析</p>
                  <p style={{ fontSize: '12px', marginTop: '8px' }}>
                    频谱图将在此处显示
                  </p>
                </div>
              </div>
            </Card>
          </Col>
        </Row>
      </Card>

      {/* 分析结果 */}
      <Card title="分析结果" style={{ marginTop: '24px' }}>
        <Row gutter={[16, 16]}>
          <Col xs={24} sm={12} md={6}>
            <Card type="inner">
              <div style={{ textAlign: 'center' }}>
                <p style={{ color: '#8c8c8c', fontSize: '12px', marginBottom: '8px' }}>
                  峰值频率
                </p>
                <p style={{ fontSize: '20px', fontWeight: 'bold' }}>
                  -- Hz
                </p>
              </div>
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card type="inner">
              <div style={{ textAlign: 'center' }}>
                <p style={{ color: '#8c8c8c', fontSize: '12px', marginBottom: '8px' }}>
                  能量
                </p>
                <p style={{ fontSize: '20px', fontWeight: 'bold' }}>
                  -- dB
                </p>
              </div>
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card type="inner">
              <div style={{ textAlign: 'center' }}>
                <p style={{ color: '#8c8c8c', fontSize: '12px', marginBottom: '8px' }}>
                  THD
                </p>
                <p style={{ fontSize: '20px', fontWeight: 'bold' }}>
                  -- %
                </p>
              </div>
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card type="inner">
              <div style={{ textAlign: 'center' }}>
                <p style={{ color: '#8c8c8c', fontSize: '12px', marginBottom: '8px' }}>
                  SNR
                </p>
                <p style={{ fontSize: '20px', fontWeight: 'bold' }}>
                  -- dB
                </p>
              </div>
            </Card>
          </Col>
        </Row>
      </Card>
    </div>
  );
}

export default SpectrumAnalyzer;
