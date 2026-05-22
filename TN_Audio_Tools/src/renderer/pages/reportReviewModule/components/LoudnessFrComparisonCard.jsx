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
      <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--text-secondary, #292524)', letterSpacing: 0.2 }}>
        {title}
      </div>
      <div
        style={{
          minHeight: 220,
          borderRadius: 16,
          border: '1px solid #e7e5e4',
          background: '#fafaf9',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          overflow: 'hidden',
          padding: 12,
        }}
      >
        {image ? (
          <img
            src={image.src}
            alt={image.fileName || title}
            onClick={function() { if (onPreview) onPreview(image, title); }}
            style={{ width: '100%', maxHeight: 196, objectFit: 'contain', display: 'block', cursor: 'zoom-in' }}
          />
        ) : (
          <div style={{ fontSize: 12, color: '#78716c', textAlign: 'center', lineHeight: 1.6 }}>{emptyText}</div>
        )}
      </div>
      {image && (
        <div style={{ fontSize: 11, color: '#78716c', lineHeight: 1.5 }}>
          {(image.contextText || image.fileName || '').slice(0, 120)}
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
          borderRadius: 24,
          border: '1px solid #e7e5e4',
          background: '#ffffff',
          boxShadow: '0 10px 30px rgba(12,10,9,0.04)',
        }}
        bodyStyle={{ padding: 20 }}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12, marginBottom: 16, alignItems: 'flex-start' }}>
          <div>
            <div style={{ fontSize: 12, color: '#78716c', textTransform: 'uppercase', letterSpacing: 0.8, marginBottom: 4 }}>
              {copy.heading}
            </div>
            <div style={{ fontSize: 22, lineHeight: 1.2, fontWeight: 500, color: '#1c1917' }}>{titleText}</div>
            <div style={{ fontSize: 12, color: '#78716c', marginTop: 6, lineHeight: 1.5 }}>{copy.pairLabel}</div>
          </div>
          <Space size={6} wrap>
            <Tag color="geekblue" style={{ borderRadius: 999 }}>{card.direction || 'unknown'}</Tag>
            <Tag color={statusMeta.color} style={{ borderRadius: 999 }}>{statusMeta.text}</Tag>
          </Space>
        </div>

        <Row gutter={[20, 16]}>
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
            marginTop: 16,
            borderRadius: 18,
            padding: '14px 16px',
            background: '#f5f5f4',
            border: '1px solid #ece7e2',
          }}
        >
          <div style={{ fontSize: 12, color: '#78716c', marginBottom: 6, letterSpacing: 0.3 }}>摘要</div>
          <div style={{ fontSize: 14, color: '#292524', lineHeight: 1.7 }}>{card.summary || '暂无摘要。'}</div>
          {highlightFinding && (highlightFinding.expectedBehavior || highlightFinding.actualBehavior) && (
            <div style={{ marginTop: 10, fontSize: 12, color: '#57534e', lineHeight: 1.6 }}>
              {highlightFinding.expectedBehavior && <div>期望：{highlightFinding.expectedBehavior}</div>}
              {highlightFinding.actualBehavior && <div>实际：{highlightFinding.actualBehavior}</div>}
            </div>
          )}
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
          <div style={{ background: '#fafaf9', borderRadius: 16, border: '1px solid #e7e5e4', padding: 16 }}>
            <img
              src={previewState.image.src}
              alt={previewState.image.fileName || previewState.title}
              style={{ width: '100%', maxHeight: '70vh', objectFit: 'contain', display: 'block' }}
            />
            <div style={{ marginTop: 12, fontSize: 12, color: '#78716c', lineHeight: 1.6 }}>
              {(previewState.image.contextText || previewState.image.fileName || '').slice(0, 240)}
            </div>
          </div>
        )}
      </Modal>
    </>
  );
}
