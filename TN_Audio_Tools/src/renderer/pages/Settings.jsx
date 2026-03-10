import React from 'react';
import { Card, Form, Select, Switch, Button, Space, Row, Col, Divider, Input } from 'antd';
import '../styles/pages.css';

function Settings() {
  const [form] = Form.useForm();

  return (
    <div className="page-container">
      <Card title="应用设置">
        <Form form={form} layout="vertical">
          <h3 style={{ marginTop: '24px', marginBottom: '16px', fontSize: '16px', fontWeight: '600' }}>
            外观设置
          </h3>
          <Row gutter={16}>
            <Col xs={24} md={12}>
              <Form.Item label="主题">
                <Select
                  defaultValue="auto"
                  disabled
                  options={[
                    { label: '自动', value: 'auto' },
                    { label: '亮色', value: 'light' },
                    { label: '暗色', value: 'dark' }
                  ]}
                />
              </Form.Item>
            </Col>
            <Col xs={24} md={12}>
              <Form.Item label="语言">
                <Select
                  defaultValue="zh-cn"
                  disabled
                  options={[
                    { label: '简体中文', value: 'zh-cn' },
                    { label: '繁體中文', value: 'zh-tw' },
                    { label: 'English', value: 'en-us' }
                  ]}
                />
              </Form.Item>
            </Col>
          </Row>

          <Form.Item label="启用系统托盘">
            <Switch disabled />
          </Form.Item>

          <Form.Item label="开启时最小化到托盘">
            <Switch disabled />
          </Form.Item>

          <Divider />

          <h3 style={{ marginTop: '24px', marginBottom: '16px', fontSize: '16px', fontWeight: '600' }}>
            文件设置
          </h3>

          <Form.Item label="默认输出目录">
            <Input 
              placeholder="选择默认输出目录"
              disabled
              addonAfter={
                <Button type="link" size="small" disabled>
                  浏览
                </Button>
              }
            />
          </Form.Item>

          <Form.Item label="最大并发任务数">
            <Select
              defaultValue="4"
              disabled
              options={[
                { label: '1', value: '1' },
                { label: '2', value: '2' },
                { label: '4', value: '4' },
                { label: '8', value: '8' }
              ]}
            />
          </Form.Item>

          <Divider />

          <h3 style={{ marginTop: '24px', marginBottom: '16px', fontSize: '16px', fontWeight: '600' }}>
            音频设置
          </h3>

          <Form.Item label="默认输出格式">
            <Select
              defaultValue="mp3"
              disabled
              options={[
                { label: 'MP3', value: 'mp3' },
                { label: 'WAV', value: 'wav' },
                { label: 'FLAC', value: 'flac' },
                { label: 'AAC', value: 'aac' }
              ]}
            />
          </Form.Item>

          <Form.Item label="默认比特率 (kbps)">
            <Select
              defaultValue="192"
              disabled
              options={[
                { label: '128', value: '128' },
                { label: '192', value: '192' },
                { label: '256', value: '256' },
                { label: '320', value: '320' }
              ]}
            />
          </Form.Item>

          <Form.Item label="默认采样率">
            <Select
              defaultValue="44100"
              disabled
              options={[
                { label: '保持原采样率', value: 'original' },
                { label: '44.1 kHz', value: '44100' },
                { label: '48 kHz', value: '48000' },
                { label: '96 kHz', value: '96000' }
              ]}
            />
          </Form.Item>

          <Divider />

          <h3 style={{ marginTop: '24px', marginBottom: '16px', fontSize: '16px', fontWeight: '600' }}>
            关于
          </h3>

          <Row gutter={16}>
            <Col xs={24} md={12}>
              <Card type="inner">
                <p><strong>应用名称：</strong>TN Audio Toolkit</p>
                <p><strong>版本：</strong>1.0.0</p>
                <p><strong>构建日期：</strong>2026-03-10</p>
              </Card>
            </Col>
            <Col xs={24} md={12}>
              <Card type="inner">
                <p><strong>开发者：</strong>JohnsonJinyu</p>
                <p><strong>仓库：</strong>TN_Audio_Check_Tools</p>
              </Card>
            </Col>
          </Row>

          <Space style={{ marginTop: '24px' }}>
            <Button type="primary" disabled>保存设置</Button>
            <Button disabled>恢复默认</Button>
            <Button danger disabled>清除缓存</Button>
          </Space>
        </Form>
      </Card>
    </div>
  );
}

export default Settings;
