import React from 'react';
import { Alert, Modal } from 'antd';
import ReviewResultContent from '../ReviewResultContent';

export default function DetailModal(props) {
  const { open, reportName, resultData, hideCrossReportSections, onClose } = props;

  return (
    <Modal
      title={`报告审查详情：${reportName || ''}`}
      open={open}
      onCancel={onClose}
      width={900}
      footer={null}
    >
      {resultData ? (
        <ReviewResultContent resultData={{
          ...resultData,
          hideCrossReportSections: !!hideCrossReportSections
        }} />
      ) : (
        <Alert type="warning" showIcon message="当前记录没有可展示的详情数据" />
      )}
    </Modal>
  );
}
