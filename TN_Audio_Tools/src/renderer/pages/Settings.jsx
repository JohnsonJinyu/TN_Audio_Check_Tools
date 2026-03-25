import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Alert, App as AntdApp, Button, Card, Col, Divider, Form, Input, Row, Select, Space, Spin, Switch, Tag, Typography } from 'antd';
import '../styles/pages.css';

function SettingSection({ title, description, children, extra = null }) {
  return (
    <section className="settings-section">
      <div className="settings-section__header">
        <div>
          <h3 className="settings-section__title">{title}</h3>
          {description ? <p className="settings-section__description">{description}</p> : null}
        </div>
        {extra ? <div className="settings-section__extra">{extra}</div> : null}
      </div>
      <div className="settings-section__body">{children}</div>
    </section>
  );
}

const fallbackSettings = {
  appearance: {
    theme: 'auto',
    language: 'zh-cn'
  },
  system: {
    enableTray: false,
    launchMinimizedToTray: false
  },
  files: {
    defaultOutputDirectory: '',
    maxConcurrentTasks: 4
  },
  audio: {
    defaultOutputFormat: 'mp3',
    defaultBitrate: '192',
    defaultSampleRate: '44100'
  }
};
const localeMap = {
  'zh-cn': 'zh-CN',
  'zh-tw': 'zh-TW',
  'en-us': 'en-US'
};
const APPEARANCE_PREVIEW_EVENT = 'app-settings:appearance-preview';

const statusMeta = {
  unsupported: { label: '不可用', color: 'default' },
  idle: { label: '待检查', color: 'default' },
  checking: { label: '检查中', color: 'processing' },
  'up-to-date': { label: '已是最新版本', color: 'success' },
  available: { label: '发现更新', color: 'warning' }
};

function formatDateTime(value, language = 'zh-cn') {
  if (!value) {
    return '暂无';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '暂无';
  }

  return date.toLocaleString(localeMap[language] || 'zh-CN', { hour12: false });
}

function normalizeSettings(settings = fallbackSettings) {
  return JSON.parse(JSON.stringify(settings || fallbackSettings));
}

function settingsEqual(left, right) {
  return JSON.stringify(left || {}) === JSON.stringify(right || {});
}

function emitAppearancePreview(appearance) {
  if (typeof window === 'undefined' || typeof window.dispatchEvent !== 'function' || !appearance) {
    return;
  }

  window.dispatchEvent(new CustomEvent(APPEARANCE_PREVIEW_EVENT, {
    detail: {
      theme: appearance.theme || fallbackSettings.appearance.theme,
      language: appearance.language || fallbackSettings.appearance.language
    }
  }));
}

function Settings() {
  const { message } = AntdApp.useApp();
  const electronApi = typeof window !== 'undefined' ? window.electron : null;
  const hasSettingsBridge = Boolean(electronApi?.settings?.get && electronApi?.settings?.save);
  const hasUpdatesBridge = Boolean(electronApi?.updates?.getState);
  const hasAppInfoBridge = Boolean(electronApi?.appInfo?.getVersion);
  const [form] = Form.useForm();
  const [appVersion, setAppVersion] = useState('');
  const [appSettings, setAppSettings] = useState(normalizeSettings(fallbackSettings));
  const [updateState, setUpdateState] = useState(null);
  const [checkingManually, setCheckingManually] = useState(false);
  const [draftSettings, setDraftSettings] = useState(normalizeSettings(appSettings || fallbackSettings));
  const [saving, setSaving] = useState(false);
  const [clearingCache, setClearingCache] = useState(false);
  const [autoSaveMessage, setAutoSaveMessage] = useState('设置将自动保存');
  const autoSaveTimerRef = useRef(null);
  const isHydratingRef = useRef(true);
  const skipNextAutoSaveRef = useRef(false);

  useEffect(() => {
    return () => {
      if (autoSaveTimerRef.current) {
        clearTimeout(autoSaveTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    let mounted = true;

    const loadInitialData = async () => {
      try {
        const [version, state, settings] = await Promise.all([
          hasAppInfoBridge ? electronApi.appInfo.getVersion() : Promise.resolve(''),
          hasUpdatesBridge ? electronApi.updates.getState() : Promise.resolve(null),
          hasSettingsBridge ? electronApi.settings.get() : Promise.resolve(fallbackSettings)
        ]);

        if (!mounted) {
          return;
        }

        setAppVersion(version || '');
        setUpdateState(state);
        setAppSettings(normalizeSettings(settings || fallbackSettings));
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
    const unsubscribe = hasUpdatesBridge && typeof electronApi?.updates?.onStateChanged === 'function'
      ? electronApi.updates.onStateChanged((nextState) => {
        if (mounted) {
          setUpdateState(nextState);
        }
      })
      : () => {};
    const unsubscribeSettings = hasSettingsBridge && typeof electronApi?.settings?.onChanged === 'function'
      ? electronApi.settings.onChanged((nextSettings) => {
        if (!mounted || !nextSettings) {
          return;
        }

        setAppSettings(normalizeSettings(nextSettings));
      })
      : () => {};

    return () => {
      mounted = false;
      unsubscribe();
      unsubscribeSettings();
    };
  }, []);

  useEffect(() => {
    const nextSettings = normalizeSettings(appSettings || fallbackSettings);
    form.setFieldsValue(nextSettings);
    setDraftSettings(nextSettings);
    isHydratingRef.current = true;
    setAutoSaveMessage('设置将自动保存');
  }, [appSettings, form]);

  const currentStatus = updateState?.status || 'idle';
  const currentStatusMeta = statusMeta[currentStatus] || statusMeta.idle;
  const canCheckForUpdates = currentStatus !== 'checking';
  const canOpenExternalDownload = Boolean(
    updateState?.available && (updateState?.externalDownloadUrl || updateState?.githubDownloadUrl || updateState?.releasePageUrl)
    && !updateState?.unsupported
  );
  const effectiveSettings = normalizeSettings(appSettings || fallbackSettings);
  const selectedLanguage = draftSettings?.appearance?.language || effectiveSettings.appearance.language;
  const trayEnabled = Boolean(draftSettings?.system?.enableTray);

  const releaseNotes = useMemo(() => {
    if (!updateState?.releaseNotes) {
      return null;
    }

    return String(updateState.releaseNotes).trim();
  }, [updateState]);

  const persistSettings = async (nextSettings, successText = '设置已自动保存') => {
    if (!hasSettingsBridge) {
      setAutoSaveMessage('当前环境不支持设置持久化');
      return;
    }

    setSaving(true);
    try {
      const savedSettings = await electronApi.settings.save(nextSettings);
      const normalized = normalizeSettings(savedSettings);
      form.setFieldsValue(normalized);
      setAppSettings(normalized);
      setDraftSettings(normalized);
      setAutoSaveMessage(successText);
    } catch (error) {
      setAutoSaveMessage(error?.message || '自动保存失败');
      message.error(error?.message || '保存设置失败');
    } finally {
      setSaving(false);
    }
  };

  const scheduleAutoSave = (nextSettings) => {
    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
    }

    setAutoSaveMessage('正在自动保存...');
    autoSaveTimerRef.current = setTimeout(() => {
      autoSaveTimerRef.current = null;
      persistSettings(nextSettings);
    }, 350);
  };

  const handleManualCheck = async () => {
    if (!hasUpdatesBridge) {
      message.warning('当前环境不支持在线更新。');
      return;
    }

    setCheckingManually(true);
    try {
      await electronApi.updates.checkForUpdates();
    } finally {
      setCheckingManually(false);
    }
  };

  const handleOpenExternalDownload = async () => {
    if (!hasUpdatesBridge || typeof electronApi?.updates?.openExternalDownload !== 'function') {
      message.warning('当前环境不支持外部下载。');
      return;
    }

    const result = await electronApi.updates.openExternalDownload({ preferMirror: true });
    if (!result?.ok) {
      message.warning(result?.message || '当前没有可用的外部下载地址。');
      return;
    }

    message.success('已在浏览器中打开下载地址。');
  };

  const handleBrowseOutputDirectory = async () => {
    if (!hasSettingsBridge || typeof electronApi?.settings?.chooseOutputDirectory !== 'function') {
      message.warning('当前环境不支持目录选择。');
      return;
    }

    const result = await electronApi.settings.chooseOutputDirectory();
    if (result?.canceled || !result?.filePath) {
      return;
    }

    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = null;
    }

    skipNextAutoSaveRef.current = true;
    form.setFieldValue(['files', 'defaultOutputDirectory'], result.filePath);
    const nextSettings = normalizeSettings(form.getFieldsValue(true));
    setDraftSettings(nextSettings);
    if (!settingsEqual(effectiveSettings, nextSettings)) {
      await persistSettings(nextSettings, '输出目录已自动保存');
    }
  };

  const handleClearOutputDirectory = async () => {
    if (!draftSettings?.files?.defaultOutputDirectory) {
      return;
    }

    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = null;
    }

    skipNextAutoSaveRef.current = true;
    form.setFieldValue(['files', 'defaultOutputDirectory'], '');
    const nextSettings = normalizeSettings(form.getFieldsValue(true));
    setDraftSettings(nextSettings);
    await persistSettings(nextSettings, '输出目录已清空');
  };

  const handleValuesChange = (_, allValues) => {
    const normalized = normalizeSettings(allValues);
    if (!normalized.system.enableTray) {
      normalized.system.launchMinimizedToTray = false;
      form.setFieldValue(['system', 'launchMinimizedToTray'], false);
    }

    setDraftSettings(normalized);
    emitAppearancePreview(normalized.appearance);

    if (skipNextAutoSaveRef.current) {
      skipNextAutoSaveRef.current = false;
      return;
    }

    if (isHydratingRef.current) {
      isHydratingRef.current = false;
      return;
    }

    if (settingsEqual(effectiveSettings, normalized)) {
      setAutoSaveMessage('设置已自动保存');
      return;
    }

    scheduleAutoSave(normalized);
  };

  const handleResetSettings = async () => {
    if (!hasSettingsBridge) {
      message.warning('当前环境不支持恢复默认设置。');
      return;
    }

    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = null;
    }

    setSaving(true);
    try {
      const resetValue = await electronApi.settings.reset();
      const normalized = normalizeSettings(resetValue);
      form.setFieldsValue(normalized);
      setAppSettings(normalized);
      setDraftSettings(normalized);
      emitAppearancePreview(normalized.appearance);
      setAutoSaveMessage('已恢复默认设置');
      message.success('已恢复默认设置。');
    } catch (error) {
      message.error(error?.message || '恢复默认设置失败');
    } finally {
      setSaving(false);
    }
  };

  const handleClearCache = async () => {
    if (!hasSettingsBridge || typeof electronApi?.settings?.clearCache !== 'function') {
      message.warning('当前环境不支持清理缓存。');
      return;
    }

    setClearingCache(true);
    try {
      const result = await electronApi.settings.clearCache();
      const removedCount = Number(result?.removedTempDirectories || 0);
      message.success(`缓存已清理，移除了 ${removedCount} 个临时目录。`);
    } catch (error) {
      message.error(error?.message || '清理缓存失败');
    } finally {
      setClearingCache(false);
    }
  };

  return (
    <div className="page-container">
      <Card title="应用设置">
        {!hasSettingsBridge ? (
          <Alert
            style={{ marginBottom: 16 }}
            type="warning"
            showIcon
            message="当前未检测到桌面端设置桥接。"
            description="设置页已进入安全降级模式，因此不会因为 preload 注入缺失而直接崩溃。可继续查看页面，但设置保存、目录选择、缓存清理和在线更新会受限。"
          />
        ) : null}
        <Form form={form} layout="vertical" onValuesChange={handleValuesChange}>
          <div className="settings-action-bar">
            <div className="settings-action-bar__meta">
              <div className="settings-action-bar__title">个性化设置</div>
              <Typography.Text className="settings-action-bar__status" type={autoSaveMessage.includes('失败') ? 'danger' : 'secondary'}>
                {saving ? '正在保存...' : autoSaveMessage}
              </Typography.Text>
            </div>
            <Space wrap>
              <Button onClick={handleResetSettings} loading={saving}>
                恢复默认
              </Button>
              <Button danger onClick={handleClearCache} loading={clearingCache}>
                清除缓存
              </Button>
            </Space>
          </div>

          <SettingSection
            title="外观与系统"
            description="管理主题、语言和桌面驻留行为。主题会即时切换，语言当前主要影响组件本地化和时间格式。"
          >
            <div className="settings-grid settings-grid--two">
              <div className="settings-grid__cell">
                <Form.Item label="主题" name={['appearance', 'theme']}>
                  <Select
                    options={[
                      { label: '自动', value: 'auto' },
                      { label: '亮色', value: 'light' },
                      { label: '暗色', value: 'dark' }
                    ]}
                  />
                </Form.Item>
              </div>
              <div className="settings-grid__cell">
                <Form.Item label="语言" name={['appearance', 'language']}>
                  <Select
                    options={[
                      { label: '简体中文', value: 'zh-cn' },
                      { label: '繁體中文', value: 'zh-tw' },
                      { label: 'English', value: 'en-us' }
                    ]}
                  />
                </Form.Item>
              </div>
            </div>

            <Alert
              type="info"
              showIcon
              message="主题会立即切换。页面文案目前仍以中文为主，语言设置主要作用于组件文案和日期时间格式。"
            />

            <div className="settings-grid settings-grid--two">
              <div className="settings-switch-card">
                <div className="settings-switch-card__header">
                  <div>
                    <div className="settings-switch-card__title">启用系统托盘</div>
                    <div className="settings-switch-card__description">最小化或关闭时允许应用继续驻留后台。</div>
                  </div>
                  <Form.Item name={['system', 'enableTray']} valuePropName="checked" className="settings-switch-card__control">
                    <Switch />
                  </Form.Item>
                </div>
              </div>

              <div className="settings-switch-card">
                <div className="settings-switch-card__header">
                  <div>
                    <div className="settings-switch-card__title">开启时最小化到托盘</div>
                    <div className="settings-switch-card__description">启用系统托盘后，启动应用时直接进入后台托盘。</div>
                  </div>
                  <Form.Item
                    name={['system', 'launchMinimizedToTray']}
                    valuePropName="checked"
                    className="settings-switch-card__control"
                  >
                    <Switch disabled={!trayEnabled} />
                  </Form.Item>
                </div>
              </div>
            </div>
          </SettingSection>

          <SettingSection
            title="文件处理"
            description="控制测试数据收集的输出位置和处理并发，直接影响生成 checklist 的落盘方式与批处理效率。"
            extra={<Typography.Text type="secondary">当前默认并发：{effectiveSettings.files.maxConcurrentTasks}</Typography.Text>}
          >
            <div className="settings-grid settings-grid--two">
              <div className="settings-grid__cell settings-grid__cell--full">
                <Form.Item
                  label="默认输出目录"
                  name={['files', 'defaultOutputDirectory']}
                  extra="测试数据收集生成的 checklist 输出会优先写入这里。留空时仍写回报告同目录。"
                >
                  <div className="settings-path-picker">
                    <Input
                      className="settings-path-picker__input"
                      readOnly
                      placeholder="选择默认输出目录"
                    />
                    <Space.Compact className="settings-path-picker__actions">
                      <Button type="link" size="small" onClick={handleBrowseOutputDirectory}>
                        浏览
                      </Button>
                      <Button type="link" size="small" onClick={handleClearOutputDirectory} disabled={!draftSettings?.files?.defaultOutputDirectory}>
                        清空
                      </Button>
                    </Space.Compact>
                  </div>
                </Form.Item>
              </div>

              <div className="settings-grid__cell">
                <Form.Item
                  label="最大并发任务数"
                  name={['files', 'maxConcurrentTasks']}
                  extra="该值会直接影响测试数据收集的报告并行处理数量。"
                >
                  <Select
                    options={[
                      { label: '1', value: 1 },
                      { label: '2', value: 2 },
                      { label: '4', value: 4 },
                      { label: '8', value: 8 }
                    ]}
                  />
                </Form.Item>
              </div>

              <div className="settings-inline-note">
                <div className="settings-inline-note__title">处理建议</div>
                <div className="settings-inline-note__text">普通批量场景建议使用 2 到 4 并发，避免磁盘与 Office 进程争抢资源。</div>
              </div>
            </div>
          </SettingSection>

          <SettingSection
            title="音频偏好"
            description="这些参数会被持久保存，供后续音频导出与转码模块直接读取；当前版本的数据收集与报告审查不会改写源音频。"
          >
            <Alert
              type="info"
              showIcon
              message="当前阶段先完成偏好沉淀，后续新增音频导出链路时会直接复用这里的设置。"
            />

            <Form.Item label="默认输出格式" name={['audio', 'defaultOutputFormat']}>
              <Select
                options={[
                  { label: 'MP3', value: 'mp3' },
                  { label: 'WAV', value: 'wav' },
                  { label: 'FLAC', value: 'flac' },
                  { label: 'AAC', value: 'aac' }
                ]}
              />
            </Form.Item>

            <Form.Item label="默认比特率 (kbps)" name={['audio', 'defaultBitrate']}>
              <Select
                options={[
                  { label: '128', value: '128' },
                  { label: '192', value: '192' },
                  { label: '256', value: '256' },
                  { label: '320', value: '320' }
                ]}
              />
            </Form.Item>

            <Form.Item label="默认采样率" name={['audio', 'defaultSampleRate']}>
              <Select
                options={[
                  { label: '保持原采样率', value: 'original' },
                  { label: '44.1 kHz', value: '44100' },
                  { label: '48 kHz', value: '48000' },
                  { label: '96 kHz', value: '96000' }
                ]}
              />
            </Form.Item>
          </SettingSection>

          <SettingSection
            title="版本更新"
            description="查看当前版本状态并手动检查新版本。发现新版后，下载将在浏览器中完成，以优先保证国内网络下的可更新性。"
          >
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
                  <Button onClick={handleOpenExternalDownload} disabled={!canOpenExternalDownload}>
                    浏览器下载
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
                    <strong>{formatDateTime(updateState?.lastCheckedAt, selectedLanguage)}</strong>
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
                    <div className="settings-update-label">更新方式</div>
                    <strong>{canOpenExternalDownload ? '浏览器下载' : '检查后提供入口'}</strong>
                    <div className="settings-update-hint">内部工具优先保证可检查新版本，安装包下载在浏览器中完成。</div>
                  </Card>
                </Col>
              </Row>

              {canOpenExternalDownload ? (
                <Alert
                  type="info"
                  showIcon
                  message="下载说明"
                  description={
                    updateState?.externalDownloadUrl
                      ? `点击“浏览器下载”后会优先打开可直达的下载地址。当前下载源：${updateState?.mirrorName || '自定义下载源'}。`
                      : '点击“浏览器下载”后会在浏览器中打开发布页或下载页。'
                  }
                />
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
          </SettingSection>

          <SettingSection
            title="关于应用"
            description="查看当前桌面端版本、构建时间和项目来源信息。"
          >
            <Row gutter={16}>
              <Col xs={24} md={12}>
                <Card type="inner">
                  <p><strong>应用名称：</strong>TN Audio Toolkit</p>
                  <p><strong>版本：</strong>{appVersion || '读取中'}</p>
                  <p><strong>构建日期：</strong>2026-03-24</p>
                </Card>
              </Col>
              <Col xs={24} md={12}>
                <Card type="inner">
                  <p><strong>开发者：</strong>JohnsonJinyu</p>
                  <p><strong>仓库：</strong>TN_Audio_Check_Tools</p>
                </Card>
              </Col>
            </Row>
          </SettingSection>

          {!appSettings ? (
            <div style={{ marginTop: 16 }}>
              <Spin size="small" /> <span style={{ marginLeft: 8 }}>设置加载中</span>
            </div>
          ) : null}

          <Space direction="vertical" size={4} style={{ marginTop: 16 }}>
            <Typography.Text type="secondary">
              当前生效主题：{effectiveSettings.appearance.theme === 'auto' ? '跟随系统' : effectiveSettings.appearance.theme}
            </Typography.Text>
            <Typography.Text type="secondary">
              当前默认并发：{effectiveSettings.files.maxConcurrentTasks}
            </Typography.Text>
          </Space>
        </Form>
      </Card>
    </div>
  );
}

export default Settings;
