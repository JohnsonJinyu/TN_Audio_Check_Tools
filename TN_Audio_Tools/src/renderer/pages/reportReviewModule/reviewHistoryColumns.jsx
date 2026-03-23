import React from 'react';
import { Button, Tag } from 'antd';
import { reviewStatusColor, reviewStatusText } from './reviewSummary';

export function createReviewHistoryColumns(onOpenDetail, onOpenStatusDetail) {
  return [
    {
      title: '报告文件',
      dataIndex: 'reportName',
      key: 'reportName',
      ellipsis: true
    },
    {
      title: '审查时间',
      dataIndex: 'checkedAt',
      key: 'checkedAt',
      width: 180,
      render: (value) => (value ? new Date(value).toLocaleString() : '-')
    },
    {
      title: '审查状态',
      dataIndex: ['result', 'reviewResult', 'overallStatus'],
      key: 'status',
      width: 100,
      render: (status) => <Tag color={reviewStatusColor[status] || 'default'}>{reviewStatusText[status] || status || '-'}</Tag>
    },
    {
      title: '检查统计',
      key: 'summary',
      width: 300,
      render: (_, record) => {
        const summary = record?.result?.reviewResult?.summary || {};
        const statusItems = [
          { key: 'pass', label: '通过', color: 'green', count: Number(summary.passedChecks) || 0 },
          { key: 'warning', label: '警告', color: 'orange', count: Number(summary.warningChecks) || 0 },
          { key: 'review', label: '复核', color: 'blue', count: Number(summary.reviewChecks) || 0 },
          { key: 'error', label: '错误', color: 'red', count: Number(summary.errorChecks) || 0 }
        ];

        return (
          <div style={{ display: 'flex', flexWrap: 'nowrap', gap: 6, whiteSpace: 'nowrap' }}>
            {statusItems.map((item) => (
              <Tag
                key={item.key}
                color={item.color}
                onClick={() => onOpenStatusDetail(record, item.key)}
                style={{ marginInlineEnd: 0, cursor: 'pointer', userSelect: 'none' }}
              >
                {item.label} {item.count}
              </Tag>
            ))}
          </div>
        );
      }
    },
    {
      title: '操作',
      key: 'action',
      width: 120,
      render: (_, record) => (
        <Button size="small" type="link" onClick={() => onOpenDetail(record)}>
          查看详情
        </Button>
      )
    }
  ];
}
