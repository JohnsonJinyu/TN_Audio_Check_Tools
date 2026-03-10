import React, { useState } from 'react';
import { Layout, Menu, Tooltip } from 'antd';
import {
  FileTextOutlined,
  AudioOutlined,
  LineChartOutlined,
  SwapOutlined,
  BgColorsOutlined,
  SettingOutlined,
  HomeOutlined
} from '@ant-design/icons';
import './App.css';
import Dashboard from './pages/Dashboard';
import ReportChecker from './pages/ReportChecker';
import AudioPlayer from './pages/AudioPlayer';
import SpectrumAnalyzer from './pages/SpectrumAnalyzer';
import AudioConverter from './pages/AudioConverter';
import BatchProcessor from './pages/BatchProcessor';
import Settings from './pages/Settings';

const { Header, Sider, Content } = Layout;

function App() {
  const [collapsed, setCollapsed] = useState(false);
  const [currentPage, setCurrentPage] = useState('dashboard');

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
      label: '报告检查',
      title: '音频测试报告检查'
    },
    {
      key: 'audio-player',
      icon: <AudioOutlined />,
      label: '音频播放',
      title: '播放和分析音频文件'
    },
    {
      key: 'spectrum',
      icon: <LineChartOutlined />,
      label: '频谱分析',
      title: '音频频谱分析工具'
    },
    {
      key: 'converter',
      icon: <SwapOutlined />,
      label: '音频转换',
      title: '音频格式转换和处理'
    },
    {
      key: 'batch',
      icon: <BgColorsOutlined />,
      label: '批量处理',
      title: '批量转换和处理音频'
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

  const renderContent = () => {
    switch (currentPage) {
      case 'dashboard':
        return <Dashboard />;
      case 'report-checker':
        return <ReportChecker />;
      case 'audio-player':
        return <AudioPlayer />;
      case 'spectrum':
        return <SpectrumAnalyzer />;
      case 'converter':
        return <AudioConverter />;
      case 'batch':
        return <BatchProcessor />;
      case 'settings':
        return <Settings />;
      default:
        return <Dashboard />;
    }
  };

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
              onClick: () => setCurrentPage(item.key),
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
            <h1 className="header-title">TN Audio Toolkit</h1>
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
