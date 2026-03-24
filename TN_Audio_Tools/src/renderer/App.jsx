import React, { useEffect, useMemo, useState } from 'react';
import { Alert, App as AntdApp, ConfigProvider, Layout, Menu, theme as antdTheme } from 'antd';
import zhCN from 'antd/locale/zh_CN';
import zhTW from 'antd/locale/zh_TW';
import enUS from 'antd/locale/en_US';
import {
  FileTextOutlined,
  LineChartOutlined,
  SearchOutlined,
  SettingOutlined,
  HomeOutlined
} from '@ant-design/icons';
import './App.css';
import Dashboard from './pages/Dashboard';
import TestDataCollectionPage from './pages/TestDataCollectionPage';
import ReportReview from './pages/ReportReview';
import SpectrumAnalyzer from './pages/SpectrumAnalyzer';
import Settings from './pages/Settings';

const { Header, Sider, Content } = Layout;
const languageLocaleMap = {
  'zh-cn': zhCN,
  'zh-tw': zhTW,
  'en-us': enUS
};
const APPEARANCE_PREVIEW_EVENT = 'app-settings:appearance-preview';
const fallbackAppearanceSettings = {
  theme: 'auto',
  language: 'zh-cn'
};

function appearanceEqual(left = fallbackAppearanceSettings, right = fallbackAppearanceSettings) {
  return left?.theme === right?.theme && left?.language === right?.language;
}

function App() {
  const electronApi = typeof window !== 'undefined' ? window.electron : null;
  const hasSettingsBridge = Boolean(electronApi?.settings?.get && electronApi?.settings?.onChanged);
  const hasAppInfoBridge = Boolean(electronApi?.appInfo?.getVersion);
  const [collapsed, setCollapsed] = useState(false);
  const [currentPage, setCurrentPage] = useState('dashboard');
  const [mountedPages, setMountedPages] = useState(() => new Set(['dashboard']));
  const [appVersion, setAppVersion] = useState('');
  const [appearanceSettings, setAppearanceSettings] = useState(fallbackAppearanceSettings);
  const [prefersDarkMode, setPrefersDarkMode] = useState(() => {
    if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') {
      return false;
    }

    return window.matchMedia('(prefers-color-scheme: dark)').matches;
  });

  useEffect(() => {
    let mounted = true;

    Promise.all([
      hasAppInfoBridge ? electronApi.appInfo.getVersion() : Promise.resolve(''),
      hasSettingsBridge ? electronApi.settings.get() : Promise.resolve(null)
    ])
      .then(([version, settings]) => {
        if (!mounted) {
          return;
        }

        setAppVersion(version || '');
        setAppearanceSettings(settings?.appearance || fallbackAppearanceSettings);
      })
      .catch(() => {
        if (!mounted) {
          return;
        }

        setAppVersion('');
        setAppearanceSettings(fallbackAppearanceSettings);
      });

    const unsubscribe = hasSettingsBridge
      ? electronApi.settings.onChanged((nextSettings) => {
        if (!mounted || !nextSettings?.appearance) {
          return;
        }

        setAppearanceSettings((currentValue) => (
          appearanceEqual(currentValue, nextSettings.appearance)
            ? currentValue
            : nextSettings.appearance
        ));
      })
      : () => {};

    return () => {
      mounted = false;
      unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') {
      return () => {};
    }

    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    const handleThemeChange = (event) => {
      setPrefersDarkMode(event.matches);
    };

    setPrefersDarkMode(mediaQuery.matches);
    if (typeof mediaQuery.addEventListener === 'function') {
      mediaQuery.addEventListener('change', handleThemeChange);
      return () => mediaQuery.removeEventListener('change', handleThemeChange);
    }

    mediaQuery.addListener(handleThemeChange);
    return () => mediaQuery.removeListener(handleThemeChange);
  }, []);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return () => {};
    }

    const handleAppearancePreview = (event) => {
      const nextAppearance = event?.detail;
      if (!nextAppearance) {
        return;
      }

      setAppearanceSettings((currentValue) => (
        appearanceEqual(currentValue, nextAppearance)
          ? currentValue
          : {
            theme: nextAppearance.theme || fallbackAppearanceSettings.theme,
            language: nextAppearance.language || fallbackAppearanceSettings.language
          }
      ));
    };

    window.addEventListener(APPEARANCE_PREVIEW_EVENT, handleAppearancePreview);
    return () => window.removeEventListener(APPEARANCE_PREVIEW_EVENT, handleAppearancePreview);
  }, []);

  const language = appearanceSettings?.language || fallbackAppearanceSettings.language;
  const selectedTheme = appearanceSettings?.theme || fallbackAppearanceSettings.theme;
  const resolvedTheme = selectedTheme === 'auto'
    ? (prefersDarkMode ? 'dark' : 'light')
    : selectedTheme;

  useEffect(() => {
    document.documentElement.dataset.theme = resolvedTheme;
    document.documentElement.lang = language;
  }, [language, resolvedTheme]);

  const configProviderLocale = useMemo(
    () => languageLocaleMap[language] || zhCN,
    [language]
  );

  const themeConfig = useMemo(
    () => ({
      algorithm: resolvedTheme === 'dark' ? antdTheme.darkAlgorithm : antdTheme.defaultAlgorithm,
      token: {
        borderRadius: 10,
        colorPrimary: resolvedTheme === 'dark' ? '#7ab8ff' : '#4c6ef5'
      }
    }),
    [resolvedTheme]
  );

  const pageMeta = {
    dashboard: {
      title: '仪表盘',
      description: '快速进入测试数据收集、报告审查和分析模块。'
    },
    'report-checker': {
      title: '测试数据收集',
      description: '上传报告、checklist 与规则文件后，统一执行数据收集并生成结论。'
    },
    'report-review': {
      title: '报告审查',
      description: '集中查看审查范围、处理结果和历史输出记录。'
    },
    spectrum: {
      title: '频谱分析',
      description: '实时查看音频频谱与波形特性。'
    },
    settings: {
      title: '设置',
      description: '配置应用偏好与处理选项。'
    }
  };

  const currentPageMeta = pageMeta[currentPage] || pageMeta.dashboard;

  const menuItems = [
    {
      key: 'dashboard',
      icon: <HomeOutlined />,
      label: '仪表盘',
      title: '应用主页'
    },
    {
      type: 'divider'
    },
    {
      key: 'report-checker',
      icon: <FileTextOutlined />,
      label: '测试数据收集',
      title: '测试报告数据收集'
    },
    {
      key: 'report-review',
      icon: <SearchOutlined />,
      label: '报告审查',
      title: '查看报告审查面板'
    },
    {
      key: 'spectrum',
      icon: <LineChartOutlined />,
      label: '频谱分析',
      title: '音频频谱分析工具'
    },
    {
      type: 'divider'
    },
    {
      key: 'settings',
      icon: <SettingOutlined />,
      label: '设置',
      title: '应用设置'
    }
  ];

  const navigateToPage = (pageKey) => {
    setCurrentPage(pageKey);
    setMountedPages((prev) => {
      if (prev.has(pageKey)) {
        return prev;
      }

      const next = new Set(prev);
      next.add(pageKey);
      return next;
    });
  };

  const pageComponents = {
    dashboard: <Dashboard onNavigate={navigateToPage} />,
    'report-checker': <TestDataCollectionPage />,
    'report-review': <ReportReview onNavigate={navigateToPage} />,
    spectrum: <SpectrumAnalyzer />,
    settings: <Settings />
  };

  const renderContent = () => Array.from(mountedPages).map((pageKey) => (
    <div
      key={pageKey}
      style={{ display: currentPage === pageKey ? 'block' : 'none', height: '100%' }}
    >
      {pageComponents[pageKey]}
    </div>
  ));

  return (
    <ConfigProvider locale={configProviderLocale} theme={themeConfig}>
      <AntdApp>
        <Layout style={{ height: '100vh' }}>
          <Sider 
            trigger={null} 
            collapsible 
            collapsed={collapsed}
            width={220}
            className="sider"
          >
            <div className="logo">
              <span className="logo-icon">🎵</span>
              {!collapsed && <span className="logo-text">音频工具集</span>}
            </div>
            <Menu
              theme="dark"
              mode="inline"
              selectedKeys={[currentPage]}
              items={menuItems.map(item => {
                if (item.type === 'divider') {
                  return item;
                }
                return {
                  ...item,
                  onClick: () => navigateToPage(item.key),
                  title: undefined
                };
              })}
              style={{ marginTop: '10px' }}
            />
          </Sider>

          <Layout>
            <Header className="header">
              <div className="header-left">
                <button 
                  className="trigger-btn"
                  onClick={() => setCollapsed(!collapsed)}
                  title={collapsed ? '展开菜单' : '收起菜单'}
                >
                  {collapsed ? '▶' : '◀'}
                </button>
                <div className="header-meta">
                  <h1 className="header-title">{currentPageMeta.title}</h1>
                  <p className="header-description">{currentPageMeta.description}</p>
                </div>
              </div>
              <div className="header-right">
                <span className="version">{appVersion ? `v${appVersion}` : '版本读取中'}</span>
              </div>
            </Header>

            <Content className="content">
              {!hasSettingsBridge ? (
                <Alert
                  type="warning"
                  showIcon
                  style={{ marginBottom: 16 }}
                  message="桌面端桥接未加载，设置与本地文件能力暂不可用。"
                  description="页面已切换为安全降级模式，因此不会因为 window.electron.settings 缺失而白屏。请从 Electron 桌面端入口启动应用，或检查 preload 是否成功加载。"
                />
              ) : null}
              {renderContent()}
            </Content>
          </Layout>
        </Layout>
      </AntdApp>
    </ConfigProvider>
  );
}

export default App;
