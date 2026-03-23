import React from 'react';
import { Button, Tag } from 'antd';

const statusText = {
  pass: '通过',
  warning: '有警告',
  review: '需复核',
  error: '有错误'
};

const statusColor = {
  pass: 'green',
  warning: 'orange',
  review: 'blue',
  error: 'red'
};

export function createReviewHistoryColumns(onOpenDetail) {
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
      render: (status) => <Tag color={statusColor[status] || 'default'}>{statusText[status] || status || '-'}</Tag>
    },
    {
      title: '检查统计',
      key: 'summary',
      width: 150,
      render: (_, record) => {
        const summary = record?.result?.reviewResult?.summary || {};
        return `通过${summary.passedChecks || 0}/${summary.warningChecks || 0}警/${summary.reviewChecks || 0}复`;
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
