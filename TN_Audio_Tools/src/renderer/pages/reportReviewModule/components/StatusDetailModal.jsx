import React, { useEffect, useMemo, useState } from 'react';
import { Alert, Collapse, Empty, Modal, Space, Tag } from 'antd';
import { reviewStatusColor, reviewStatusText } from '../reviewSummary';
import { reportReviewModalMotionProps } from './modalMotion';

export default function StatusDetailModal(props) {
  const { open, selectedDetail, filteredSections, onClose } = props;
  const modalRootClassName = 'report-review-detail-modal report-review-status-modal';
  const [activeKeys, setActiveKeys] = useState([]);

  useEffect(() => {
    if (!open) {
      setActiveKeys([]);
    }
  }, [open, selectedDetail]);

  const normalizedActiveKeys = useMemo(() => {
    if (Array.isArray(activeKeys)) {
      return activeKeys.map((key) => String(key));
    }

    if (activeKeys === undefined || activeKeys === null) {
      return [];
    }

    return [String(activeKeys)];
  }, [activeKeys]);

  if (!selectedDetail) {
    return (
      <Modal
        title="状态明细"
        rootClassName={modalRootClassName}
        {...reportReviewModalMotionProps}
        open={open}
        onCancel={onClose}
        footer={null}
        width={860}
      >
        <Empty description="未选择状态" />
      </Modal>
    );
  }

  const status = selectedDetail.status;
  const record = selectedDetail.record;
  const alertType = status === 'pass' ? 'success' : (status === 'warning' ? 'warning' : (status === 'error' ? 'error' : 'info'));
  const statusColor = status === 'pass'
    ? 'var(--status-pass)'
    : (status === 'warning' ? 'var(--status-warn)' : (status === 'error' ? 'var(--status-error)' : 'var(--status-info)'));
  const collapseItems = filteredSections.map((section, sectionIndex) => {
    const panelKey = String(section.key || `section-${sectionIndex}`);
    const isActive = normalizedActiveKeys.includes(panelKey);

    return {
      key: panelKey,
      style: {
        background: 'var(--surface-color)',
        border: '1px solid var(--border-color)',
        borderRadius: 16,
        marginBottom: sectionIndex === filteredSections.length - 1 ? 0 : 14
      },
      label: (
        <div className="report-review-result__panel-label" style={{ padding: '8px 4px' }}>
          <Space size="middle">
            <Tag color={reviewStatusColor[section.status] || 'default'} style={{ margin: 0, padding: '0 8px', borderRadius: 4 }}>
              {reviewStatusText[section.status] || section.status}
            </Tag>
            <span className="report-review-result__panel-title" style={{ fontWeight: 600, fontSize: 15, color: 'var(--text-color)' }}>
              {section.title || '未命名检查项'}
            </span>
          </Space>
        </div>
      ),
      children: isActive ? (
        <div className="report-review-result__panel-body report-review-status-detail__panel-body" style={{ padding: '0 16px 24px 48px' }}>
          <div style={{ color: 'var(--text-light)', marginBottom: 16, lineHeight: 1.7, fontSize: 14 }}>
            {section.description || '无详细说明'}
          </div>

          {Array.isArray(section.issues) && section.issues.length > 0 ? (
            <div className="report-review-status-detail__issue-list" style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: section.evidence?.length ? 16 : 0 }}>
              {section.issues.map((issue, index) => (
                <Alert
                  key={`${section.key}-issue-${index}`}
                  type={issue?.severity === 'error' ? 'error' : (issue?.severity === 'warning' ? 'warning' : 'info')}
                  showIcon
                  message={issue?.message || '未提供问题说明'}
                />
              ))}
            </div>
          ) : (
            <Alert
              type={section.status === 'pass' ? 'success' : alertType}
              showIcon
              style={{ marginBottom: section.evidence?.length ? 16 : 0 }}
              message={section.status === 'pass' ? '当前检查项已通过，未记录问题项。' : '当前检查项没有单独记录问题描述。'}
            />
          )}

          {Array.isArray(section.evidence) && section.evidence.length > 0 && (
            <div className="report-review-result__evidence report-review-status-detail__evidence" style={{ marginBottom: 0, marginTop: 12, padding: '14px 18px', background: 'var(--surface-elevated)', border: '1px solid var(--border-color)', borderRadius: 8 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
                <span style={{ fontSize: 15 }}>💡</span>
                <span style={{ fontSize: 14, fontWeight: 600, color: 'var(--text-color)' }}>证据记录</span>
              </div>
              <ul style={{ margin: 0, paddingLeft: 24, display: 'flex', flexDirection: 'column', gap: 8 }}>
                {section.evidence.map((evidence, index) => (
                  <li key={`${section.key}-evidence-${index}`} style={{ color: 'var(--text-light)', fontSize: 14, lineHeight: 1.6 }}>
                    {evidence}
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      ) : null
    };
  });

  return (
    <Modal
      title={`${record?.reportName || ''} - ${reviewStatusText[status] || status} 明细`}
      rootClassName={modalRootClassName}
      {...reportReviewModalMotionProps}
      open={open}
      onCancel={onClose}
      footer={null}
      width={860}
    >
      <div className="report-review-result report-review-status-detail">
        <div
          className="report-review-result__hero report-review-status-detail__hero"
          style={{
            padding: '24px 28px',
            background: 'var(--surface-color)',
            border: '1px solid var(--border-color)',
            borderLeft: `6px solid ${statusColor}`,
            borderRadius: 12,
            display: 'flex',
            alignItems: 'center',
            gap: 20,
            marginBottom: 32,
            boxShadow: 'var(--shadow-color)'
          }}
        >
          <div style={{ minWidth: 0 }}>
            <div className="report-review-result__eyebrow" style={{ fontSize: 12, color: 'var(--text-light)', marginBottom: 4, letterSpacing: 0.5, textTransform: 'uppercase', fontWeight: 600 }}>
              Status Focus
            </div>
            <div style={{ fontSize: 20, fontWeight: 600, color: 'var(--text-color)', marginBottom: 8 }}>
              当前共 {filteredSections.length} 项“{reviewStatusText[status] || status}”检查
            </div>
            <div style={{ color: 'var(--text-light)', fontSize: 14, lineHeight: 1.7 }}>
              点击表格中的统计标签即可按类别筛选，不必再从完整详情里逐项查找。
            </div>
          </div>
        </div>

        <div className="report-review-result__section-heading" style={{ marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
          <div className="report-review-result__section-bar" style={{ width: 4, height: 16, background: statusColor, borderRadius: 2 }}></div>
          <h3 style={{ margin: 0, fontSize: 16, color: 'var(--text-color)', fontWeight: 600 }}>筛选检查结果</h3>
        </div>

        {filteredSections.length > 0 ? (
          <Collapse
            className="report-review-result__collapse report-review-status-detail__collapse"
            expandIconPosition="end"
            bordered={false}
            destroyInactivePanel
            activeKey={normalizedActiveKeys}
            onChange={setActiveKeys}
            style={{ background: 'transparent', border: 'none', borderRadius: 12, boxShadow: 'none', overflow: 'visible' }}
            items={collapseItems}
          />
        ) : (
          <Empty description={`当前记录没有“${reviewStatusText[status] || status}”项`} />
        )}
      </div>
    </Modal>
  );
}
