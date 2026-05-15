import React from 'react';
import { Alert, Modal, Space } from 'antd';

export default function ReviewAreaModal(props) {
  const { open, area, onClose } = props;

  return (
    <Modal
      title={area ? `${area.tag} - ${area.title}` : '检查范围说明'}
      open={open}
      onCancel={onClose}
      footer={null}
      width={640}
    >
      {area ? (
        <div>
          <Alert
            type="info"
            showIcon
            style={{ marginBottom: 16 }}
            message={area.description}
          />
          <Space direction="vertical" size={12} style={{ width: '100%' }}>
            {(area.details || []).map((detail, index) => (
              <div key={`${area.title}-${index}`} className="review-area-modal__item">
                <span className="review-area-modal__index">{index + 1}</span>
                <span>{detail}</span>
              </div>
            ))}
          </Space>
        </div>
      ) : null}
    </Modal>
  );
}
