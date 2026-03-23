import React, { useState } from 'react';
import { Layout, Menu } from 'antd';
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

function App() {
  const [collapsed, setCollapsed] = useState(false);
  const [currentPage, setCurrentPage] = useState('dashboard');
  const [mountedPages, setMountedPages] = useState(() => new Set(['dashboard']));

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
              title: undefined // Ant Design Menu 中去掉 title
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
            <span className="version">v1.0.0</span>
          </div>
        </Header>

        <Content className="content">
          {renderContent()}
        </Content>
      </Layout>
    </Layout>
  );
}

export default App;
