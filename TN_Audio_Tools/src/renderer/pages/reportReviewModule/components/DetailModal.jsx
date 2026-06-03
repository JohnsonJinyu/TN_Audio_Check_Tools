import React from 'react';
import { Alert, Modal } from 'antd';
import ReviewResultContent from '../ReviewResultContent';
import { reportReviewModalMotionProps } from './modalMotion';

export default function DetailModal(props) {
  const { open, reportName, resultData, hideCrossReportSections, onClose } = props;

  return (
    <Modal
      title={`报告审查详情：${reportName || ''}`}
      rootClassName="report-review-detail-modal"
      {...reportReviewModalMotionProps}
      open={open}
      onCancel={onClose}
      width={1100}
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
