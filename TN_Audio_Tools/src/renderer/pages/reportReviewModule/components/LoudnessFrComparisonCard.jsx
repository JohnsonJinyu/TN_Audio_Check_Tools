import React, { useState } from 'react';
import { Card, Col, Modal, Row, Space, Tag } from 'antd';

var STATUS_META = {
  pass: { color: 'success', text: '通过' },
  warning: { color: 'orange', text: '警告' },
  review: { color: 'blue', text: '复核' },
  error: { color: 'red', text: '错误' },
};

function getCardCopy(direction) {
  if (direction === 'RCV') {
    return {
      heading: 'Receiving',
      pairLabel: 'RLR 响度曲线 vs 接收频响曲线',
      frTitle: '接收频响图 (Sensitivity, frequency RCV)',
      loudnessTitle: '接收响度图 (Loudness Rating RCV / RLR)',
      frEmptyText: '当前等级未找到可对照的接收频响基准图',
      loudnessEmptyText: '当前等级未找到可对照的 RLR 响度曲线图',
    };
  }
  if (direction === 'SND') {
    return {
      heading: 'Sending',
      pairLabel: 'SLR 响度曲线 vs 发送频响曲线',
      frTitle: '发送频响图 (Sensitivity, frequency SND)',
      loudnessTitle: '发送响度图 (Loudness Rating SND / SLR)',
      frEmptyText: '当前等级未找到可对照的发送频响基准图',
      loudnessEmptyText: '当前等级未找到可对照的 SLR 响度曲线图',
    };
  }
  return {
    heading: 'Unknown',
    pairLabel: '方向待确认的响度曲线 vs 频响曲线',
    frTitle: '频响图',
    loudnessTitle: '响度曲线图',
    frEmptyText: '当前等级未找到可对照的频响基准图',
    loudnessEmptyText: '当前等级未找到可对照的响度曲线图',
  };
}

function PreviewPane({ title, image, emptyText, onPreview }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--text-light)', letterSpacing: 0.2 }}>
        {title}
      </div>
      <div
        style={{
          minHeight: 180,
          borderRadius: 8,
          border: '1px solid var(--border-color)',
          background: 'var(--surface-muted)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          overflow: 'hidden',
          padding: 8,
        }}
      >
        {image ? (
          <img
            src={image.src}
            alt={image.fileName || title}
            onClick={function() { if (onPreview) onPreview(image, title); }}
            style={{ width: '100%', maxHeight: 180, objectFit: 'contain', display: 'block', cursor: 'zoom-in' }}
          />
        ) : (
          <div style={{ fontSize: 12, color: 'var(--text-light)', textAlign: 'center', lineHeight: 1.6 }}>{emptyText}</div>
        )}
      </div>
      {image && (
        <div 
          style={{ 
            fontSize: 11, 
            color: 'var(--text-light)', 
            lineHeight: 1.5,
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            maxWidth: '100%' 
          }}
          title={image.contextText || image.fileName || ''}
        >
          {image.contextText || image.fileName || ''}
        </div>
      )}
    </div>
  );
}

export default function LoudnessFrComparisonCard({ card, hideLevel }) {
  if (!card) return null;

  var [previewState, setPreviewState] = useState({ open: false, image: null, title: '' });
  var statusMeta = STATUS_META[card.status] || STATUS_META.review;
  var copy = getCardCopy(card.direction);
  var titleText = hideLevel
    ? (card.direction === 'SND' ? '发送侧曲线' : '同组曲线')
    : (card.level || '等级未知');
  var loudnessImage = Array.isArray(card.loudnessImages) ? card.loudnessImages[0] : null;
  var highlightFinding = Array.isArray(card.findings)
    ? card.findings.find(function(item) { return item.status === 'error' || item.status === 'warning'; }) || card.findings[0]
    : null;

  function openPreview(image, title) {
    setPreviewState({ open: true, image: image, title: title || image.fileName || '曲线大图预览' });
  }

  return (
    <>
      <Card
        size="small"
        style={{
          borderRadius: 12,
          border: '1px solid var(--border-color)',
          background: 'var(--surface-color)',
          boxShadow: '0 1px 2px 0 rgba(0, 0, 0, 0.03)',
        }}
        bodyStyle={{ padding: 0 }}
      >
        <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--border-color)', background: 'var(--surface-muted)', borderRadius: '12px 12px 0 0', display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'flex-start' }}>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 4 }}>
              <div style={{ fontSize: 18, lineHeight: 1.2, fontWeight: 600, color: 'var(--text-color)' }}>{titleText}</div>
              <Space size={6}>
                <Tag color="geekblue" style={{ borderRadius: 4, margin: 0 }}>{card.direction || 'unknown'}</Tag>
                <Tag color={statusMeta.color} style={{ borderRadius: 4, margin: 0 }}>{statusMeta.text}</Tag>
              </Space>
            </div>
            <div style={{ fontSize: 12, color: 'var(--text-light)', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ textTransform: 'uppercase', letterSpacing: 0.5, fontWeight: 500 }}>{copy.heading}</span>
              <span style={{ color: 'var(--border-color)' }}>|</span>
              <span>{copy.pairLabel}</span>
            </div>
          </div>
        </div>

        <div style={{ padding: 20 }}>
          <Row gutter={[24, 20]}>
            <Col xs={24} lg={12}>
              <PreviewPane
                title={copy.frTitle}
                image={card.frImage}
                emptyText={copy.frEmptyText}
                onPreview={openPreview}
              />
            </Col>
            <Col xs={24} lg={12}>
              <PreviewPane
                title={copy.loudnessTitle}
                image={loudnessImage}
                emptyText={copy.loudnessEmptyText}
                onPreview={openPreview}
              />
            </Col>
          </Row>

          <div
            style={{
              marginTop: 24,
              borderLeft: `3px solid ${statusMeta.color === 'success' ? 'var(--status-pass)' : statusMeta.color === 'red' ? 'var(--status-error)' : statusMeta.color === 'orange' ? 'var(--status-warn)' : 'var(--status-info)'}`,
              background: statusMeta.color === 'success' ? 'color-mix(in srgb, var(--status-pass) 8%, transparent)' : statusMeta.color === 'red' ? 'color-mix(in srgb, var(--status-error) 8%, transparent)' : statusMeta.color === 'orange' ? 'color-mix(in srgb, var(--status-warn) 8%, transparent)' : 'color-mix(in srgb, var(--status-info) 8%, transparent)',
              padding: '12px 16px',
              borderRadius: '0 8px 8px 0',
            }}
          >
            <div style={{ fontSize: 14, color: 'var(--text-color)', fontWeight: 500 }}>{card.summary || '暂无摘要。'}</div>
            {highlightFinding && (highlightFinding.expectedBehavior || highlightFinding.actualBehavior) && (
              <div style={{ marginTop: 8, fontSize: 13, color: 'var(--text-light)', lineHeight: 1.6 }}>
                {highlightFinding.expectedBehavior && <div style={{ marginBottom: 4 }}><strong>期望：</strong>{highlightFinding.expectedBehavior}</div>}
                {highlightFinding.actualBehavior && <div><strong>实际：</strong>{highlightFinding.actualBehavior}</div>}
              </div>
            )}
          </div>
        </div>
      </Card>

      <Modal
        open={previewState.open}
        onCancel={function() { setPreviewState({ open: false, image: null, title: '' }); }}
        footer={null}
        width={980}
        title={previewState.title || '曲线大图预览'}
        centered
      >
        {previewState.image && (
          <div style={{ background: 'var(--surface-color)', borderRadius: 16, border: '1px solid var(--border-color)', padding: 16 }}>
            <img
              src={previewState.image.src}
              alt={previewState.image.fileName || previewState.title}
              style={{ width: '100%', maxHeight: '70vh', objectFit: 'contain', display: 'block' }}
            />
            <div style={{ marginTop: 12, fontSize: 12, color: 'var(--text-light)', lineHeight: 1.6 }}>
              {(previewState.image.contextText || previewState.image.fileName || '').slice(0, 240)}
            </div>
          </div>
        )}
      </Modal>
    </>
  );
}
