import React, { useEffect, useMemo, useState } from 'react';
import { Alert, Button, Card, Col, Divider, Form, Input, Progress, Row, Select, Space, Switch, Tag, Typography } from 'antd';
import '../styles/pages.css';

const statusMeta = {
  unsupported: { label: '不可用', color: 'default' },
  idle: { label: '待检查', color: 'default' },
  checking: { label: '检查中', color: 'processing' },
  'up-to-date': { label: '已是最新版本', color: 'success' },
  available: { label: '发现更新', color: 'warning' },
  downloading: { label: '下载中', color: 'processing' },
  downloaded: { label: '下载完成', color: 'success' },
  installing: { label: '安装中', color: 'processing' }
};

function formatDateTime(value) {
  if (!value) {
    return '暂无';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '暂无';
  }

  return date.toLocaleString('zh-CN', { hour12: false });
}

function Settings() {
  const [form] = Form.useForm();
  const [appVersion, setAppVersion] = useState('');
  const [updateState, setUpdateState] = useState(null);
  const [checkingManually, setCheckingManually] = useState(false);
  const [downloadingManually, setDownloadingManually] = useState(false);

  useEffect(() => {
    let mounted = true;

    const loadInitialData = async () => {
      try {
        const [version, state] = await Promise.all([
          window.electron.appInfo.getVersion(),
          window.electron.updates.getState()
        ]);

        if (!mounted) {
          return;
        }

        setAppVersion(version || '');
        setUpdateState(state);
      } catch (error) {
        if (!mounted) {
          return;
        }

        setUpdateState({
          status: 'unsupported',
          error: error?.message || '读取更新状态失败。'
        });
      }
    };

    loadInitialData();
    const unsubscribe = window.electron.updates.onStateChanged((nextState) => {
      if (mounted) {
        setUpdateState(nextState);
      }
    });

    return () => {
      mounted = false;
      unsubscribe();
    };
  }, []);

  const currentStatus = updateState?.status || 'idle';
  const currentStatusMeta = statusMeta[currentStatus] || statusMeta.idle;
  const canCheckForUpdates = currentStatus !== 'checking' && currentStatus !== 'downloading' && currentStatus !== 'installing';
  const canDownload = updateState?.available && !updateState?.downloaded && currentStatus !== 'downloading' && !updateState?.unsupported;
  const canInstall = Boolean(updateState?.downloaded);

  const releaseNotes = useMemo(() => {
    if (!updateState?.releaseNotes) {
      return null;
    }

    return String(updateState.releaseNotes).trim();
  }, [updateState]);

  const handleManualCheck = async () => {
    setCheckingManually(true);
    try {
      await window.electron.updates.checkForUpdates();
    } finally {
      setCheckingManually(false);
    }
  };

  const handleDownloadUpdate = async () => {
    setDownloadingManually(true);
    try {
      await window.electron.updates.downloadUpdate();
    } finally {
      setDownloadingManually(false);
    }
  };

  const handleInstallUpdate = async () => {
    await window.electron.updates.quitAndInstall();
  };

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
            版本更新
          </h3>

          <Card type="inner" className="settings-update-card">
            <Space direction="vertical" size={16} style={{ width: '100%' }}>
              <div className="settings-update-header">
                <div>
                  <div className="settings-update-label">当前状态</div>
                  <Space size={12} wrap>
                    <Typography.Title level={4} style={{ margin: 0 }}>
                      {updateState?.latestVersion && updateState.latestVersion !== appVersion
                        ? `v${appVersion || '...'} -> v${updateState.latestVersion}`
                        : `v${appVersion || '...'}`}
                    </Typography.Title>
                    <Tag color={currentStatusMeta.color}>{currentStatusMeta.label}</Tag>
                  </Space>
                </div>
                <Space wrap>
                  <Button
                    type="primary"
                    onClick={handleManualCheck}
                    loading={checkingManually || currentStatus === 'checking'}
                    disabled={!canCheckForUpdates}
                  >
                    检查更新
                  </Button>
                  <Button
                    onClick={handleDownloadUpdate}
                    loading={downloadingManually || currentStatus === 'downloading'}
                    disabled={!canDownload}
                  >
                    下载更新
                  </Button>
                  <Button type="primary" ghost onClick={handleInstallUpdate} disabled={!canInstall}>
                    重启安装
                  </Button>
                </Space>
              </div>

              {updateState?.error ? (
                <Alert
                  type={updateState.status === 'unsupported' ? 'warning' : 'error'}
                  showIcon
                  message={updateState.error}
                />
              ) : null}

              <Row gutter={16}>
                <Col xs={24} md={8}>
                  <Card size="small" className="settings-update-metric">
                    <div className="settings-update-label">最近检查</div>
                    <strong>{formatDateTime(updateState?.lastCheckedAt)}</strong>
                    <div className="settings-update-hint">
                      {updateState?.lastCheckSource === 'manual'
                        ? '手动检查'
                        : updateState?.lastCheckSource === 'auto'
                          ? '启动自动检查'
                          : '暂无记录'}
                    </div>
                  </Card>
                </Col>
                <Col xs={24} md={8}>
                  <Card size="small" className="settings-update-metric">
                    <div className="settings-update-label">目标版本</div>
                    <strong>{updateState?.latestVersion ? `v${updateState.latestVersion}` : '暂无'}</strong>
                    <div className="settings-update-hint">{updateState?.releaseName || '等待新版本信息'}</div>
                  </Card>
                </Col>
                <Col xs={24} md={8}>
                  <Card size="small" className="settings-update-metric">
                    <div className="settings-update-label">最近下载</div>
                    <strong>{formatDateTime(updateState?.lastDownloadedAt)}</strong>
                    <div className="settings-update-hint">仅安装版支持在线更新</div>
                  </Card>
                </Col>
              </Row>

              {currentStatus === 'downloading' ? (
                <div>
                  <Progress percent={Number(updateState?.progressPercent || 0)} />
                  <div className="settings-update-hint">
                    已下载 {updateState?.transferred || 0} / {updateState?.total || 0} 字节，速度 {updateState?.bytesPerSecond || 0} B/s
                  </div>
                </div>
              ) : null}

              {releaseNotes ? (
                <Alert
                  type="info"
                  showIcon
                  message="版本说明"
                  description={<div className="settings-update-release-notes">{releaseNotes}</div>}
                />
              ) : null}
            </Space>
          </Card>

          <Divider />

          <h3 style={{ marginTop: '24px', marginBottom: '16px', fontSize: '16px', fontWeight: '600' }}>
            关于
          </h3>

          <Row gutter={16}>
            <Col xs={24} md={12}>
              <Card type="inner">
                <p><strong>应用名称：</strong>TN Audio Toolkit</p>
                <p><strong>版本：</strong>{appVersion || '读取中'}</p>
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
