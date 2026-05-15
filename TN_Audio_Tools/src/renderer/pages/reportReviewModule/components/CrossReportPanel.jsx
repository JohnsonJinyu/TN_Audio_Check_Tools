import React from 'react';
import { Alert, Button, Card, Col, Row, Table, Tag } from 'antd';
import { getReportName } from '../utils';

export default function CrossReportPanel(props) {
  const { crossReportResults, latestReviewDigests, onOpenDetail, onOpenStatusDetail } = props;

  // Cross-report comparison overview
  if (crossReportResults && crossReportResults.results && crossReportResults.results.length >= 2) {
    var reports = crossReportResults.results;
    var codecs = [...new Set(reports.map(function(r) {
      return r.reviewResult?.reportFacts?.metadata?.codec || '?';
    }))];
    var networks = [...new Set(reports.map(function(r) {
      return r.reviewResult?.reportFacts?.metadata?.network || '?';
    }))];

    return (
      <>
        <Col xs={24}>
          <Card
            className="report-checker-card"
            title="本轮批量对比概览"
            extra={<span style={{ color: 'var(--text-light)', fontSize: 12 }}>{crossReportResults.checkedAt ? new Date(crossReportResults.checkedAt).toLocaleString() : ''}</span>}
            style={{ borderColor: '#d6e4ff' }}
          >
            <div>
              <div style={{ display: 'flex', gap: 24, flexWrap: 'wrap', marginBottom: 14 }}>
                <div><span style={{ color: 'var(--text-light)' }}>报告数：</span><strong>{reports.length}</strong></div>
                <div><span style={{ color: 'var(--text-light)' }}>Codec：</span><strong>{codecs.join(' · ')}</strong></div>
                <div><span style={{ color: 'var(--text-light)' }}>网络：</span><strong>{networks.join(' · ')}</strong></div>
              </div>
              <Table
                dataSource={reports.map(function(r) {
                  var s = r.reviewResult?.summary || {};
                  var md = r.reviewResult?.reportFacts?.metadata || {};
                  var name = getReportName(r.docxPath || r.reportPath || '');
                  return {
                    key: name,
                    name: name,
                    network: md.network || '-',
                    codec: md.codec || '-',
                    passed: s.passedChecks || 0,
                    warning: s.warningChecks || 0,
                    review: s.reviewChecks || 0,
                    error: s.errorChecks || 0,
                    total: s.totalChecks || 17,
                    status: r.reviewResult?.overallStatus || 'unknown',
                    record: r
                  };
                })}
                rowKey="key"
                size="small"
                pagination={false}
                columns={[
                  { title: '报告', dataIndex: 'name', key: 'name', ellipsis: true, width: 180 },
                  { title: '网络', dataIndex: 'network', key: 'network', width: 70 },
                  { title: 'Codec', dataIndex: 'codec', key: 'codec', width: 70 },
                  { title: '通过', dataIndex: 'passed', key: 'passed', width: 58, align: 'center', render: function(v, row) { return v > 0 ? <a onClick={function(e) { e.stopPropagation(); if (onOpenStatusDetail) onOpenStatusDetail(row.record, 'pass'); }} style={{ color: '#52c41a', fontWeight: 600, cursor: 'pointer' }}>{v}</a> : <span style={{ color: '#52c41a' }}>0</span>; } },
                  { title: '警告', dataIndex: 'warning', key: 'warning', width: 58, align: 'center', render: function(v, row) { return v > 0 ? <a onClick={function(e) { e.stopPropagation(); if (onOpenStatusDetail) onOpenStatusDetail(row.record, 'warning'); }} style={{ color: '#faad14', fontWeight: 600, cursor: 'pointer' }}>{v}</a> : <span>0</span>; } },
                  { title: '复核', dataIndex: 'review', key: 'review', width: 58, align: 'center', render: function(v, row) { return v > 0 ? <a onClick={function(e) { e.stopPropagation(); if (onOpenStatusDetail) onOpenStatusDetail(row.record, 'review'); }} style={{ color: '#1677ff', fontWeight: 600, cursor: 'pointer' }}>{v}</a> : <span>0</span>; } },
                  { title: '错误', dataIndex: 'error', key: 'error', width: 58, align: 'center', render: function(v, row) { return v > 0 ? <a onClick={function(e) { e.stopPropagation(); if (onOpenStatusDetail) onOpenStatusDetail(row.record, 'error'); }} style={{ color: '#f5222d', fontWeight: 600, cursor: 'pointer' }}>{v}</a> : <span>0</span>; } },
                  { title: '总计', dataIndex: 'total', key: 'total', width: 54, align: 'center' },
                  { title: '状态', dataIndex: 'status', key: 'status', width: 70, render: function(v) { return <Tag color={v === 'pass' ? 'success' : (v === 'error' ? 'error' : (v === 'warning' ? 'warning' : 'processing'))}>{v === 'pass' ? '通过' : (v === 'error' ? '错误' : (v === 'warning' ? '警告' : '复核'))}</Tag>; } }
                ]}
              />
            </div>
          </Card>
        </Col>

        <Col xs={24}>
          <Card
            className="report-checker-card"
            title={`批量对比结果（跨报告） — ${crossReportResults.reportCount} 份报告`}
            extra={<span style={{ color: 'var(--text-light)', fontSize: 12 }}>{crossReportResults.checkedAt ? new Date(crossReportResults.checkedAt).toLocaleString() : ''}</span>}
            style={{ borderColor: '#d6e4ff' }}
          >
            <CrossReportDetail reports={reports} />
          </Card>
        </Col>
      </>
    );
  }

  // Latest review digest (single report)
  if (latestReviewDigests && latestReviewDigests.length > 0) {
    return (
      <Col xs={24}>
        <Card
          className="report-checker-card"
          title="最近审查结论"
          extra={<span style={{ color: 'var(--text-light)', fontSize: 12 }}>批量对比时自动切换为对比概览</span>}
        >
          <Row gutter={[16, 16]}>
            {latestReviewDigests.slice(0, 1).map(function(record) {
              return (
                <Col key={record.id} xs={24}>
                  <div
                    onClick={function() { onOpenDetail(record); }}
                    style={{ cursor: 'pointer', padding: '16px 20px', borderRadius: 14, background: 'var(--surface-color)', border: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}
                  >
                    <div>
                      <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 4 }}>{record.reportName}</div>
                      <div style={{ fontSize: 12, color: 'var(--text-light)' }}>
                        {record.checkedAt ? new Date(record.checkedAt).toLocaleString() : '-'}
                        {' · '}通过 {record.digest.summary.passedChecks} / 警告 {record.digest.summary.warningChecks} / 复核 {record.digest.summary.reviewChecks} / 错误 {record.digest.summary.errorChecks}
                      </div>
                    </div>
                    <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                      <Button size="small" onClick={function(e) {
                        e.stopPropagation();
                        onOpenDetail(record);
                      }}>查看详情</Button>
                      <Tag color={record.digest.statusColor}>{record.digest.statusText}</Tag>
                    </div>
                  </div>
                </Col>
              );
            })}
          </Row>
        </Card>
      </Col>
    );
  }

  return null;
}

function CrossReportDetail(props) {
  var reports = props.reports;
  var codecCheck = reports[0]?.reviewResult?.checks?.contentSameCodecDiffNetwork;
  var networkCheck = reports[0]?.reviewResult?.checks?.contentSameNetworkDiffCodec;

  var codecGroups = {};
  reports.forEach(function(r) {
    var md = r.reviewResult?.reportFacts?.metadata || {};
    var tf = r.reviewResult?.testDataFacts;
    var codec = md.codec || '未知';
    if (!codecGroups[codec]) codecGroups[codec] = [];
    var metrics = tf?.loudnessMetrics || [];
    var minRow = metrics.length > 0 ? metrics.reduce(function(a, b) {
      return (a.rlr || 99) < (b.rlr || 99) ? a : b;
    }) : null;
    codecGroups[codec].push({
      name: getReportName(r.reportPath || r.docxPath || ''),
      network: md.network || '-',
      slr: minRow?.slr != null ? Number(minRow.slr).toFixed(1) : '-',
      rlr: minRow?.rlr != null ? Number(minRow.rlr).toFixed(1) : '-',
      stmr: minRow?.stmr != null ? Number(minRow.stmr).toFixed(1) : '-',
      path: r.reportPath || r.docxPath
    });
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      {Object.keys(codecGroups).map(function(codec) {
        var items = codecGroups[codec];
        if (items.length < 2) return null;
        return (
          <div key={'codec-' + codec} style={{ border: '1px solid var(--border-color)', borderRadius: 12, padding: '16px 20px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
              <div style={{ fontWeight: 600, fontSize: 16 }}>同 Codec 不同网络响度差异 — {codec}</div>
              <Tag color={codecCheck?.status === 'pass' ? 'success' : (codecCheck?.status === 'warning' ? 'warning' : (codecCheck?.status === 'error' ? 'error' : 'default'))}>
                {codecCheck?.status === 'pass' ? '通过' : (codecCheck?.status === 'warning' ? '警告' : (codecCheck?.status === 'error' ? '错误' : '待对比'))}
              </Tag>
            </div>
            <Table
              dataSource={items}
              rowKey="name"
              size="small"
              pagination={false}
              columns={[
                { title: '报告', dataIndex: 'name', key: 'name', ellipsis: true },
                { title: '网络', dataIndex: 'network', key: 'network', width: 80 },
                { title: 'RLR', dataIndex: 'rlr', key: 'rlr', width: 70 },
                { title: 'SLR', dataIndex: 'slr', key: 'slr', width: 70 },
                { title: 'STMR', dataIndex: 'stmr', key: 'stmr', width: 70 }
              ]}
            />
            {codecCheck?.issues && codecCheck.issues.length > 0 && (
              <div style={{ marginTop: 10 }}>
                {codecCheck.issues.map(function(iss, i) {
                  return (
                    <Alert
                      key={i}
                      type={iss.severity === 'error' ? 'error' : (iss.severity === 'warning' ? 'warning' : 'info')}
                      showIcon
                      message={iss.message}
                      style={{ marginBottom: 6 }}
                    />
                  );
                })}
              </div>
            )}
            {codecCheck?.evidence && codecCheck.evidence.length > 0 && <EvidencePanel checkResult={codecCheck} />}
          </div>
        );
      })}

      <div style={{ border: '1px solid var(--border-color)', borderRadius: 12, padding: '16px 20px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div style={{ fontWeight: 600, fontSize: 16 }}>同网络不同 Codec 响度差异</div>
          <Tag color={networkCheck?.status === 'pass' ? 'success' : (networkCheck?.status === 'warning' ? 'warning' : (networkCheck?.status === 'error' ? 'error' : 'default'))}>
            {networkCheck?.status === 'pass' ? '通过' : (networkCheck?.status === 'warning' ? '警告' : (networkCheck?.status === 'error' ? '错误' : '待对比'))}
          </Tag>
        </div>
        {networkCheck?.issues && networkCheck.issues.length > 0 && (
          <div>
            {networkCheck.issues.map(function(iss, i) {
              return (
                <Alert
                  key={i}
                  type={iss.severity === 'error' ? 'error' : (iss.severity === 'warning' ? 'warning' : 'info')}
                  showIcon
                  message={iss.message}
                  style={{ marginBottom: 6 }}
                />
              );
            })}
          </div>
        )}
        {networkCheck?.evidence && networkCheck.evidence.length > 0 ? (
          <EvidencePanel checkResult={networkCheck} />
        ) : (
          <div style={{ marginTop: 10, color: 'var(--text-light)', fontSize: 13 }}>各报告 codec 相同，无跨 codec 对比数据</div>
        )}
      </div>
    </div>
  );
}

function EvidencePanel(props) {
  var checkResult = props.checkResult;
  var evList = checkResult.evidence;
  var pairLines = evList.filter(function(l) { return /diff=\d+\.\d+d?B/.test(l); });
  var summaryLine = evList.filter(function(l) { return l.indexOf('✓') === 0 && l.indexOf('diff') === -1; })[0] || '';
  var headerLines = evList.filter(function(l) { return l.indexOf(':') > -1 && l.indexOf('个') > -1 && l.indexOf('diff') === -1; });

  return (
    <div style={{ marginTop: 14, padding: '14px 16px', background: checkResult?.status === 'pass' ? '#f6ffed' : (checkResult?.status === 'warning' ? '#fffbe6' : '#fff2f0'), borderRadius: 10, border: '1px solid ' + (checkResult?.status === 'pass' ? '#b7eb8f' : (checkResult?.status === 'warning' ? '#ffe58f' : '#ffa39e')) }}>
      <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 10, color: checkResult?.status === 'pass' ? '#389e0d' : (checkResult?.status === 'warning' ? '#d48806' : '#cf1322') }}>
        {summaryLine || (checkResult?.status === 'pass' ? '✓ 所有差异均在1dB以内' : '存在超出阈值的差异')}
      </div>
      {headerLines.length > 0 && (
        <div style={{ fontSize: 13, color: '#555', marginBottom: 8 }}>{headerLines.join(' | ')}</div>
      )}
      {pairLines.length > 0 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          {pairLines.map(function(line, i) {
            var isOK = line.indexOf('✓') > -1;
            return (
              <div key={i} style={{ fontSize: 13, color: isOK ? '#389e0d' : '#cf1322', fontFamily: "'JetBrains Mono', 'Fira Code', monospace", padding: '3px 0' }}>
                {isOK ? '✓ ' : '✗ '}{line.replace(/^\s+/, '').replace(/ ✓$/, '')}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
