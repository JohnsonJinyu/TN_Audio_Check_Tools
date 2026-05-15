import React from 'react';
import { Alert, Card, Empty, Modal, Space, Tag } from 'antd';
import { reviewStatusColor, reviewStatusText } from '../reviewSummary';

export default function StatusDetailModal(props) {
  const { open, selectedDetail, filteredSections, onClose } = props;

  if (!selectedDetail) {
    return (
      <Modal title="状态明细" open={open} onCancel={onClose} footer={null} width={860}>
        <Empty description="未选择状态" />
      </Modal>
    );
  }

  const status = selectedDetail.status;
  const record = selectedDetail.record;
  const alertType = status === 'pass' ? 'success' : (status === 'warning' ? 'warning' : (status === 'error' ? 'error' : 'info'));

  return (
    <Modal
      title={`${record?.reportName || ''} - ${reviewStatusText[status] || status} 明细`}
      open={open}
      onCancel={onClose}
      footer={null}
      width={860}
    >
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <Alert
          type={alertType}
          showIcon
          message={`当前共 ${filteredSections.length} 项"${reviewStatusText[status] || status}"检查`}
          description="点击表格中的统计标签即可按类别筛选，不必再从完整详情里逐项查找。"
        />

        {filteredSections.length > 0 ? filteredSections.map((section) => (
          <Card
            key={`${section.key}-${section.status}`}
            size="small"
            title={(
              <Space>
                <Tag color={reviewStatusColor[section.status] || 'default'} style={{ marginInlineEnd: 0 }}>
                  {reviewStatusText[section.status] || section.status}
                </Tag>
                <span>{section.title || '未命名检查项'}</span>
              </Space>
            )}
            className="review-status-detail-card"
          >
            <div style={{ color: 'var(--text-light)', marginBottom: 12, lineHeight: 1.7 }}>
              {section.description || '无详细说明'}
            </div>

            {Array.isArray(section.issues) && section.issues.length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: section.evidence?.length ? 14 : 0 }}>
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
                type={section.status === 'pass' ? 'success' : 'info'}
                showIcon
                style={{ marginBottom: section.evidence?.length ? 14 : 0 }}
                message={section.status === 'pass' ? '当前检查项已通过，未记录问题项。' : '当前检查项没有单独记录问题描述。'}
              />
            )}

            {Array.isArray(section.evidence) && section.evidence.length > 0 && (
              <div>
                <div style={{ fontWeight: 600, color: 'var(--text-color)', marginBottom: 8 }}>证据记录</div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {section.evidence.map((evidence, index) => (
                    <div
                      key={`${section.key}-evidence-${index}`}
                      style={{ padding: '10px 12px', borderRadius: 10, border: '1px solid var(--border-color)', background: 'var(--surface-elevated)', color: 'var(--text-light)', lineHeight: 1.65 }}
                    >
                      {evidence}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </Card>
        )) : (
          <Empty description={`当前记录没有"${reviewStatusText[status] || status}"项`} />
        )}
      </div>
    </Modal>
  );
}
