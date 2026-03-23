import React from 'react';
import { Alert, Card, Col, Collapse, Divider, Row, Space, Tag } from 'antd';
import { CheckOutlined, CloseOutlined, ExclamationOutlined } from '@ant-design/icons';

function getStatusColor(status) {
  const colorMap = {
    pass: 'green',
    warning: 'orange',
    review: 'blue',
    error: 'red'
  };

  return colorMap[status] || 'default';
}

function getStatusIcon(status) {
  if (status === 'pass') return <CheckOutlined />;
  if (status === 'error') return <CloseOutlined />;
  return <ExclamationOutlined />;
}

function getStatusText(status) {
  const textMap = {
    pass: '通过',
    warning: '有警告',
    review: '需人工复核',
    error: '有错误'
  };

  return textMap[status] || status;
}

export default function ReviewResultContent({ resultData }) {
  if (!resultData) {
    return null;
  }

  const { report, reviewResult } = resultData;
  if (!reviewResult) {
    return <Alert type="error" message="审查结果数据格式异常，无法显示" />;
  }

  const summary = reviewResult.summary || {
    passedChecks: 0,
    warningChecks: 0,
    reviewChecks: 0,
    errorChecks: 0
  };
  const sections = Array.isArray(report?.sections) ? report.sections : [];

  if (sections.length === 0) {
    return (
      <Alert
        type="warning"
        showIcon
        message="当前历史记录缺少可展示的详细检查项"
        description="这条记录可能来自旧版本数据，或审查结果未完整保存。"
      />
    );
  }

  return (
    <div>
      <Row gutter={[16, 16]} style={{ marginBottom: 24 }}>
        <Col xs={24}>
          <Alert
            message={`整体评估：${getStatusText(reviewResult.overallStatus)}`}
            type={reviewResult.overallStatus === 'pass' ? 'success' : (
              reviewResult.overallStatus === 'error' ? 'error' : (
                reviewResult.overallStatus === 'warning' ? 'warning' : 'info'
              )
            )}
            icon={getStatusIcon(reviewResult.overallStatus)}
            showIcon
            style={{ marginBottom: 16 }}
          />
        </Col>

        <Col xs={24} sm={6}>
          <Card size="small" title="审查统计">
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: '8px 0', color: '#52c41a', fontSize: 16, fontWeight: 'bold' }}>{summary.passedChecks}</p>
              <p style={{ margin: 0, fontSize: 12, color: '#999' }}>通过项</p>
            </div>
          </Card>
        </Col>

        <Col xs={24} sm={6}>
          <Card size="small">
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: '8px 0', color: '#faad14', fontSize: 16, fontWeight: 'bold' }}>{summary.warningChecks}</p>
              <p style={{ margin: 0, fontSize: 12, color: '#999' }}>警告项</p>
            </div>
          </Card>
        </Col>

        <Col xs={24} sm={6}>
          <Card size="small">
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: '8px 0', color: '#1677ff', fontSize: 16, fontWeight: 'bold' }}>{summary.reviewChecks}</p>
              <p style={{ margin: 0, fontSize: 12, color: '#999' }}>需复核</p>
            </div>
          </Card>
        </Col>

        <Col xs={24} sm={6}>
          <Card size="small">
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: '8px 0', color: '#ff4d4f', fontSize: 16, fontWeight: 'bold' }}>{summary.errorChecks}</p>
              <p style={{ margin: 0, fontSize: 12, color: '#999' }}>错误项</p>
            </div>
          </Card>
        </Col>
      </Row>

      <Divider />

      <h3 style={{ marginBottom: 16 }}>详细检查结果</h3>
      <Collapse
        items={sections.map((section, sectionIndex) => ({
          key: section.key || `section-${sectionIndex}`,
          label: (
            <Space>
              <Tag color={getStatusColor(section.status)}>{getStatusText(section.status || 'review')}</Tag>
              <span style={{ fontWeight: 500 }}>{section.title || '未命名检查项'}</span>
            </Space>
          ),
          children: (
            <div>
              <p style={{ marginBottom: 12, color: '#666' }}>{section.description || '无详细说明'}</p>

              {Array.isArray(section.issues) && section.issues.length > 0 && (
                <div style={{ marginBottom: 16 }}>
                  <h4 style={{ marginBottom: 8 }}>问题项：</h4>
                  {section.issues.map((issue, idx) => (
                    <Alert
                      key={idx}
                      message={issue?.message || '未提供问题说明'}
                      type={issue?.severity === 'error' ? 'error' : (issue?.severity === 'warning' ? 'warning' : 'info')}
                      showIcon
                      style={{ marginBottom: 8 }}
                    />
                  ))}
                </div>
              )}

              {Array.isArray(section.evidence) && section.evidence.length > 0 && (
                <div>
                  <h4 style={{ marginBottom: 8 }}>证据记录：</h4>
                  <ul style={{ marginBottom: 0, paddingLeft: 20 }}>
                    {section.evidence.map((item, idx) => (
                      <li key={idx} style={{ marginBottom: 4, color: '#666', fontSize: 12 }}>{item}</li>
                    ))}
                  </ul>
                </div>
              )}

              {section.data && section.key === 'tableOfContents' && Array.isArray(section.data.chapters) && section.data.chapters.length > 0 && (
                <div style={{ marginTop: 16 }}>
                  <h4 style={{ marginBottom: 8 }}>识别到的章节：</h4>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid #ddd' }}>
                        <th style={{ padding: 8, textAlign: 'left' }}>章节号</th>
                        <th style={{ padding: 8, textAlign: 'left' }}>标题</th>
                      </tr>
                    </thead>
                    <tbody>
                      {section.data.chapters.map((chapter, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid #eee' }}>
                          <td style={{ padding: 8 }}>{chapter.number}</td>
                          <td style={{ padding: 8 }}>{chapter.title}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {section.data && section.key === 'engineers' && Array.isArray(section.data.engineers) && section.data.engineers.length > 0 && (
                <div style={{ marginTop: 16 }}>
                  <h4 style={{ marginBottom: 8 }}>识别到的人员：</h4>
                  <ul style={{ paddingLeft: 20 }}>
                    {section.data.engineers.map((engineer, idx) => (
                      <li key={idx} style={{ marginBottom: 4 }}>{engineer.name}</li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          )
        }))}
      />
    </div>
  );
}
