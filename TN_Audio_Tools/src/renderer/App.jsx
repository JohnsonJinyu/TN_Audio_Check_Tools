import React, { useEffect, useMemo, useState } from 'react';
import { Alert, App as AntdApp, Badge, ConfigProvider, Layout, Menu, theme as antdTheme } from 'antd';
import { ThemeContext } from './ThemeContext';
import zhCN from 'antd/locale/zh_CN';
import {
  FileTextOutlined,
  LineChartOutlined,
  SearchOutlined,
  SettingOutlined,
  HomeOutlined
} from '@ant-design/icons';
import './App.css';
import './styles/theme-glassmorphism.css';
import './styles/theme-neumorphism.css';
import './styles/theme-paper-craft.css';
import './styles/theme-soft-clean.css';
import './styles/theme-classic-pro.css';
import appLogo from './assets/app-logo.svg';
import Dashboard from './pages/Dashboard';
import TestDataCollectionPage from './pages/TestDataCollectionPage';
import ReportReview from './pages/ReportReview';
import SpectrumAnalyzer from './pages/SpectrumAnalyzer';
import Settings from './pages/Settings';

const { Header, Sider, Content } = Layout;
const APPEARANCE_PREVIEW_EVENT = 'app-settings:appearance-preview';
const fallbackAppearanceSettings = {
  theme: 'light',
  designStyle: 'neumorphism'
};

/* Map legacy design styles to new ones for backward compatibility */
const LEGACY_STYLE_MAP = {
  'apple-light': 'soft-clean',
  'elevenlabs': 'classic-pro',
  'linear': 'glassmorphism',
  'claude': 'paper-craft',
  'vercel': 'classic-pro'
};

/* Each design style declares whether its sidebar is light or dark */
const SIDEBAR_THEME_MAP = {
  'glassmorphism': 'light',
  'neumorphism': 'light',
  'paper-craft': 'light',
  'soft-clean': 'light',
  'classic-pro': 'dark'
};

/* Theme tokens per design style — used for ConfigProvider */
const DESIGN_STYLE_TOKENS = {
  'glassmorphism':  { colorPrimary: '#6366f1', borderRadius: 14, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" },
  'neumorphism':    { colorPrimary: '#6b7fa8', borderRadius: 16, fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" },
  'paper-craft':    { colorPrimary: '#3d3929', borderRadius: 4,  fontFamily: "'Merriweather', 'Georgia', serif" },
  'soft-clean':     { colorPrimary: '#6366f1', borderRadius: 12, fontFamily: "'Inter', 'Plus Jakarta Sans', -apple-system, sans-serif" },
  'classic-pro':    { colorPrimary: '#1a56db', borderRadius: 6,  fontFamily: "'Inter', 'Roboto', -apple-system, sans-serif" }
};

function appearanceEqual(left = fallbackAppearanceSettings, right = fallbackAppearanceSettings) {
  return left?.theme === right?.theme
    && left?.designStyle === right?.designStyle;
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
            designStyle: nextAppearance.designStyle || fallbackAppearanceSettings.designStyle
          }
      ));
    };

    window.addEventListener(APPEARANCE_PREVIEW_EVENT, handleAppearancePreview);
    return () => window.removeEventListener(APPEARANCE_PREVIEW_EVENT, handleAppearancePreview);
  }, []);

  const selectedTheme = appearanceSettings?.theme || fallbackAppearanceSettings.theme;
  const rawDesignStyle = appearanceSettings?.designStyle || fallbackAppearanceSettings.designStyle;
  const designStyle = LEGACY_STYLE_MAP[rawDesignStyle] || rawDesignStyle;

  const resolvedTheme = selectedTheme === 'auto'
    ? (prefersDarkMode ? 'dark' : 'light')
    : selectedTheme;

  const sidebarTheme = SIDEBAR_THEME_MAP[designStyle] || 'dark';

  /* Keep data-design-style and data-theme in sync with React state */
  useEffect(() => {
    if (typeof document !== 'undefined' && document.documentElement) {
      document.documentElement.dataset.designStyle = designStyle;
      document.documentElement.dataset.theme = resolvedTheme;
    }
  }, [designStyle, resolvedTheme]);

  const themeConfig = useMemo(
    () => {
      const tokens = DESIGN_STYLE_TOKENS[designStyle] || DESIGN_STYLE_TOKENS['neumorphism'];
      const baseToken = {
        borderRadius: tokens.borderRadius,
        fontFamily: tokens.fontFamily
      };
      if (resolvedTheme === 'dark') {
        return {
          algorithm: antdTheme.darkAlgorithm,
          token: { ...baseToken, colorPrimary: tokens.colorPrimary }
        };
      }
      return {
        algorithm: antdTheme.defaultAlgorithm,
        token: { ...baseToken, colorPrimary: tokens.colorPrimary }
      };
    },
    [resolvedTheme, designStyle]
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
      label: <span>频谱分析 <Badge count="Beta" size="small" style={{ backgroundColor: 'var(--status-warn)', fontSize: 10, verticalAlign: 'middle' }} /></span>,
      title: '音频频谱分析工具 (开发中)'
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
    <ConfigProvider locale={zhCN} theme={themeConfig}>
      <ThemeContext.Provider value={resolvedTheme}>
      <AntdApp>
        <Layout style={{ height: '100vh' }}>
          <Sider
            trigger={null}
            collapsible
            collapsed={collapsed}
            width={220}
            className={`sider ${sidebarTheme === 'light' ? 'sider--light' : ''}`}
          >
            <div className={`logo ${sidebarTheme === 'light' ? 'logo--light' : ''}`}>
              <img className="logo-icon" src={appLogo} alt="TN Audio Toolkit" />
              {!collapsed && <span className="logo-text">音频工具集</span>}
            </div>
            <Menu
              theme={sidebarTheme}
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
      </ThemeContext.Provider>
    </ConfigProvider>
  );
}

export default App;
