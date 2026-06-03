import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Alert, App as AntdApp, Button, Collapse, Descriptions, Divider, Dropdown, Form, Input, Progress, Radio, Select, Space, Spin, Switch, Tag, Typography } from 'antd';
import '../styles/pages.css';

const designStyleLabelMap = {
  neumorphism: '新拟态',
  glassmorphism: '玻璃拟态',
  'paper-craft': '纸张质感',
  'soft-clean': '柔和简约',
  'classic-pro': '经典专业'
};

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
    theme: 'light',
    designStyle: 'neumorphism'
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
  },
  llm: {
    enabled: false,
    apiUrl: '',
    apiKey: '',
    model: '',
    maxImagesPerAnalysis: 12
  }
};
const APPEARANCE_PREVIEW_EVENT = 'app-settings:appearance-preview';

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
      designStyle: appearance.designStyle || fallbackSettings.appearance.designStyle
    }
  }));
}

function Settings() {
  const { message } = AntdApp.useApp();
  const electronApi = typeof window !== 'undefined' ? window.electron : null;
  const hasSettingsBridge = Boolean(electronApi?.settings?.get && electronApi?.settings?.save);
  const hasImmediateSettingsBridge = Boolean(electronApi?.settings?.saveImmediate);
  const hasUpdatesBridge = Boolean(electronApi?.updates?.getState);
  const hasAppInfoBridge = Boolean(electronApi?.appInfo?.getVersion);
  const [form] = Form.useForm();
  const [appVersion, setAppVersion] = useState('');
  const [appSettings, setAppSettings] = useState(normalizeSettings(fallbackSettings));
  const [updateState, setUpdateState] = useState(null);
  const [checkingManually, setCheckingManually] = useState(false);
  const [downloadingManually, setDownloadingManually] = useState(false);
  const [draftSettings, setDraftSettings] = useState(normalizeSettings(appSettings || fallbackSettings));
  const [saving, setSaving] = useState(false);
  const [clearingCache, setClearingCache] = useState(false);
  const [autoSaveMessage, setAutoSaveMessage] = useState('设置将自动保存');
  const autoSaveTimerRef = useRef(null);
  const isHydratingRef = useRef(true);
  const skipNextAutoSaveRef = useRef(false);
  const latestDraftSettingsRef = useRef(normalizeSettings(fallbackSettings));
  const latestAppSettingsRef = useRef(normalizeSettings(fallbackSettings));
  const persistSeqRef = useRef(0);

  const flushPendingSettings = () => {
    if (!hasImmediateSettingsBridge || !autoSaveTimerRef.current) {
      return;
    }

    clearTimeout(autoSaveTimerRef.current);
    autoSaveTimerRef.current = null;

    const nextSettings = normalizeSettings(latestDraftSettingsRef.current || fallbackSettings);
    const currentSettings = normalizeSettings(latestAppSettingsRef.current || fallbackSettings);
    if (settingsEqual(currentSettings, nextSettings)) {
      return;
    }

    electronApi.settings.saveImmediate(nextSettings);
  };

  useEffect(() => {
    latestAppSettingsRef.current = normalizeSettings(appSettings || fallbackSettings);
  }, [appSettings]);

  useEffect(() => {
    latestDraftSettingsRef.current = normalizeSettings(draftSettings || fallbackSettings);
  }, [draftSettings]);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return () => {};
    }

    const handleBeforeUnload = () => {
      flushPendingSettings();
    };

    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
      flushPendingSettings();
      if (autoSaveTimerRef.current) {
        clearTimeout(autoSaveTimerRef.current);
      }
    };
  }, [hasImmediateSettingsBridge]);

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
    /* Skip if form already holds identical values — prevents Radio.Button flicker */
    const currentFormSettings = normalizeSettings(form.getFieldsValue(true));
    if (JSON.stringify(currentFormSettings) === JSON.stringify(nextSettings)) {
      return;
    }
    isHydratingRef.current = true;
    form.setFieldsValue(nextSettings);
    setDraftSettings(nextSettings);
    setAutoSaveMessage('设置将自动保存');

    // Ant Design 的 setFieldsValue 不会把这次初始化和后续用户修改区分开，
    // 如果把 hydration 标记留到第一次 onValuesChange 再清掉，会吞掉用户的第一次真实修改。
    Promise.resolve().then(() => {
      isHydratingRef.current = false;
    });
  }, [appSettings, form]);

  const currentStatus = updateState?.status || 'idle';
  const currentStatusMeta = statusMeta[currentStatus] || statusMeta.idle;
  const canCheckForUpdates = currentStatus !== 'checking' && currentStatus !== 'downloading' && currentStatus !== 'installing';
  const canDownload = updateState?.available && !updateState?.downloaded && currentStatus !== 'downloading' && !updateState?.unsupported;
  const canInstall = Boolean(updateState?.downloaded);
  const canOpenExternalDownload = Boolean(
    (updateState?.externalDownloadUrl || updateState?.githubDownloadUrl || updateState?.releasePageUrl)
    && !updateState?.unsupported
  );
  const effectiveSettings = normalizeSettings(appSettings || fallbackSettings);
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

    const seq = ++persistSeqRef.current;
    setSaving(true);
    try {
      const savedSettings = await electronApi.settings.save(nextSettings);
      /* Ignore stale results from superseded save calls */
      if (seq !== persistSeqRef.current) {
        return;
      }
      const normalized = normalizeSettings(savedSettings);
      setAppSettings(normalized);
      setDraftSettings(normalized);
      setAutoSaveMessage(successText);
    } catch (error) {
      if (seq !== persistSeqRef.current) {
        return;
      }
      setAutoSaveMessage(error?.message || '自动保存失败');
      message.error(error?.message || '保存设置失败');
    } finally {
      if (seq === persistSeqRef.current) {
        setSaving(false);
      }
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

  const handleDownloadUpdate = async () => {
    if (!hasUpdatesBridge) {
      message.warning('当前环境不支持在线更新。');
      return;
    }

    setDownloadingManually(true);
    try {
      await electronApi.updates.downloadUpdate();
    } finally {
      setDownloadingManually(false);
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

    message.success('已在浏览器中打开镜像下载地址。');
  };

  const handleInstallUpdate = async () => {
    if (!hasUpdatesBridge) {
      message.warning('当前环境不支持在线更新。');
      return;
    }

    await electronApi.updates.quitAndInstall();
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

  const handleValuesChange = (changedValues, allValues) => {
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
      return;
    }

    if (settingsEqual(effectiveSettings, normalized)) {
      setAutoSaveMessage('设置已自动保存');
      return;
    }

    if (changedValues?.appearance) {
      if (autoSaveTimerRef.current) {
        clearTimeout(autoSaveTimerRef.current);
        autoSaveTimerRef.current = null;
      }
      persistSettings(normalized, '外观设置已保存');
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

  const handleTestConnection = async () => {
    const apiUrl = form.getFieldValue(['llm', 'apiUrl']) || '';
    const apiKey = form.getFieldValue(['llm', 'apiKey']) || '';
    const model = form.getFieldValue(['llm', 'model']) || '';
    if (!apiUrl || !apiKey) {
      message.warning('请先填写API地址和Key');
      return;
    }
    message.loading({ content: '正在测试连接...', key: 'llm-test', duration: 0 });
    try {
      const res = await window.electron.reportReview.testLlmConnection({ apiUrl, apiKey, model });
      if (res.ok) {
        message.success({ content: res.message, key: 'llm-test' });
      } else {
        message.error({ content: res.message, key: 'llm-test', duration: 8 });
      }
    } catch (e) {
      message.error({ content: '测试失败: ' + (e.message || '未知错误'), key: 'llm-test', duration: 8 });
    }
  };

  return (
    <div className="page-container settings-page">
      <div className="settings-page-header">
        <h2 className="settings-page-title">应用设置</h2>
        <p className="settings-page-subtitle">管理外观、文件输出、AI集成和更新偏好。更改将自动保存。</p>
      </div>

      {!hasSettingsBridge ? (
        <Alert
          style={{ marginBottom: 20 }}
          type="warning"
          showIcon
          message="当前未检测到桌面端设置桥接。"
          description="设置页已进入安全降级模式，因此不会因为 preload 注入缺失而直接崩溃。可继续查看页面，但设置保存、目录选择、缓存清理和在线更新会受限。"
        />
      ) : null}

      <Form form={form} layout="vertical" onValuesChange={handleValuesChange}>
          <SettingSection
            title="外观与系统"
            description="管理主题、设计风格、语言和桌面驻留行为。更改即时生效。"
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
            </div>

            <Form.Item label="设计风格" name={['appearance', 'designStyle']}>
              <Radio.Group optionType="button" buttonStyle="solid">
                <Radio.Button value="neumorphism">新拟态</Radio.Button>
                <Radio.Button value="glassmorphism">玻璃拟态</Radio.Button>
                <Radio.Button value="paper-craft">纸张质感</Radio.Button>
                <Radio.Button value="soft-clean">柔和简约</Radio.Button>
                <Radio.Button value="classic-pro">经典专业</Radio.Button>
              </Radio.Group>
            </Form.Item>

            <div className="settings-switch-list">
              <div className="settings-switch-row">
                <div className="settings-switch-row__text">
                  <div className="settings-switch-row__title">启用系统托盘</div>
                  <div className="settings-switch-row__description">最小化或关闭时允许应用继续驻留后台。</div>
                </div>
                <Form.Item name={['system', 'enableTray']} valuePropName="checked" className="settings-switch-row__control">
                  <Switch />
                </Form.Item>
              </div>
              <div className="settings-switch-row">
                <div className="settings-switch-row__text">
                  <div className="settings-switch-row__title">开启时最小化到托盘</div>
                  <div className="settings-switch-row__description">启用系统托盘后，启动应用时直接进入后台托盘。</div>
                </div>
                <Form.Item
                  name={['system', 'launchMinimizedToTray']}
                  valuePropName="checked"
                  className="settings-switch-row__control"
                >
                  <Switch disabled={!trayEnabled} />
                </Form.Item>
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

            <div className="settings-grid settings-grid--two">
              <div className="settings-grid__cell">
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
              </div>
              <div className="settings-grid__cell">
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
              </div>
              <div className="settings-grid__cell">
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
              </div>
            </div>
          </SettingSection>

          <SettingSection
            title="AI 图表分析"
            description="利用大模型视觉能力分析报告中的频率响应曲线图表，验证响度与频响趋势是否一致。需配置公司内部API凭据。API Key仅存储在本地，不会上传至任何第三方。"
          >
            <div className="settings-switch-list" style={{ marginBottom: 20 }}>
              <div className="settings-switch-row">
                <div className="settings-switch-row__text">
                  <div className="settings-switch-row__title">启用 AI 图表分析</div>
                  <div className="settings-switch-row__description">开启后，在审查详情中可手动触发 AI 图表验证。</div>
                </div>
                <Form.Item name={['llm', 'enabled']} valuePropName="checked" className="settings-switch-row__control">
                  <Switch />
                </Form.Item>
              </div>
            </div>

            <div className="settings-grid settings-grid--two">
              <div className="settings-grid__cell settings-grid__cell--full">
                <Form.Item label="API 地址" name={['llm', 'apiUrl']}>
                  <Input placeholder="https://llm.your-company.com" />
                </Form.Item>
              </div>
              <div className="settings-grid__cell settings-grid__cell--full">
                <Form.Item label="API Key" name={['llm', 'apiKey']}>
                  <Input.Password placeholder="sk-..." />
                </Form.Item>
              </div>
              <div className="settings-grid__cell settings-grid__cell--full">
                <Button onClick={handleTestConnection}>测试连接</Button>
              </div>
              <div className="settings-grid__cell">
                <Form.Item label="模型名称" name={['llm', 'model']}>
                  <Input placeholder="claude-sonnet-4-20250514" />
                </Form.Item>
              </div>
            </div>

          </SettingSection>

          <SettingSection
            title="版本更新"
            description="查看当前版本状态，手动触发检查、下载和安装。国内网络较慢时，可优先使用镜像下载在浏览器中获取安装包。"
          >
            <Space direction="vertical" size={20} style={{ width: '100%' }}>
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
                  <Dropdown menu={{ items: [
                    { key: 'download', label: '应用内下载', disabled: !canDownload, onClick: handleDownloadUpdate },
                    { key: 'mirror', label: '镜像下载', disabled: !canOpenExternalDownload, onClick: handleOpenExternalDownload },
                    { key: 'install', label: '重启安装', disabled: !canInstall, onClick: handleInstallUpdate }
                  ]}}>
                    <Button>更多操作</Button>
                  </Dropdown>
                </Space>
              </div>

              {updateState?.error ? (
                <Alert
                  type={updateState.status === 'unsupported' ? 'warning' : 'error'}
                  showIcon
                  message={updateState.error}
                />
              ) : null}

              <Descriptions size="small" column={3} bordered style={{ marginTop: 16 }}>
                <Descriptions.Item label="最近检查">{formatDateTime(updateState?.lastCheckedAt)}</Descriptions.Item>
                <Descriptions.Item label="目标版本">{updateState?.latestVersion ? `v${updateState.latestVersion}` : '暂无'}</Descriptions.Item>
                <Descriptions.Item label="最近下载">{formatDateTime(updateState?.lastDownloadedAt)}</Descriptions.Item>
              </Descriptions>

              {currentStatus === 'downloading' ? (
                <div>
                  <Progress percent={Number(updateState?.progressPercent || 0)} />
                  <div className="settings-update-hint">
                    已下载 {updateState?.transferred || 0} / {updateState?.total || 0} 字节，速度 {updateState?.bytesPerSecond || 0} B/s
                  </div>
                </div>
              ) : null}

              {canOpenExternalDownload ? (
                <Alert
                  type="info"
                  showIcon
                  message="下载加速建议"
                  description={
                    updateState?.externalDownloadUrl
                      ? '如果应用内下载速度偏慢，可从更多操作菜单选择镜像下载，使用浏览器直接下载 ' + (updateState?.assetName || '安装包') + '。'
                      : '如果应用内下载速度偏慢，可通过更多操作菜单选择镜像下载，在浏览器中打开发布页。'
                  }
                />
              ) : null}

              {releaseNotes ? (
                <Collapse items={[{ key: 'release-notes', label: '版本说明', children: <div className="settings-update-release-notes">{releaseNotes}</div> }]} />
              ) : null}
            </Space>
          </SettingSection>

          <SettingSection
            title="关于应用"
            description="查看当前桌面端版本、构建时间和项目来源信息。"
          >
            <Descriptions column={2} size="small">
              <Descriptions.Item label="应用名称">TN Audio Toolkit</Descriptions.Item>
              <Descriptions.Item label="版本">{appVersion || '读取中'}</Descriptions.Item>
              <Descriptions.Item label="构建日期">2026-03-24</Descriptions.Item>
              <Descriptions.Item label="开发者">JohnsonJinyu</Descriptions.Item>
              <Descriptions.Item label="仓库">TN_Audio_Check_Tools</Descriptions.Item>
            </Descriptions>
          </SettingSection>

          {!appSettings ? (
            <div style={{ marginTop: 16 }}>
              <Spin size="small" /> <span style={{ marginLeft: 8 }}>设置加载中</span>
            </div>
          ) : null}

          <div className="settings-footer">
            <Space wrap>
              <Button onClick={handleResetSettings} loading={saving}>
                恢复默认
              </Button>
              <Button danger onClick={handleClearCache} loading={clearingCache}>
                清除缓存
              </Button>
            </Space>
            <Typography.Text type="secondary">
              {autoSaveMessage.includes('失败') ? autoSaveMessage : ''}
              &ensp;{effectiveSettings.appearance.theme === 'auto' ? '主题：跟随系统' : `主题：${effectiveSettings.appearance.theme}`}
              &ensp;·&ensp;设计风格：{designStyleLabelMap[effectiveSettings.appearance.designStyle] || '新拟态'}
              &ensp;·&ensp;并发：{effectiveSettings.files.maxConcurrentTasks}
            </Typography.Text>
          </div>
        </Form>
    </div>
  );
}

export default Settings;
