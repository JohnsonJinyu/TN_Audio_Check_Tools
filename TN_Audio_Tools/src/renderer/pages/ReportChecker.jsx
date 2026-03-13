import React, { useState } from 'react';
import { Card, Button, Upload, Table, Space, Tag, Modal, message, Typography } from 'antd';
import { UploadOutlined, DeleteOutlined, CheckCircleOutlined, ExportOutlined } from '@ant-design/icons';
import { recordReportCheckResults } from '../modules/dashboard/storage';
import '../styles/pages.css';

const { Text, Paragraph } = Typography;
const compactUploadDraggerStyle = { padding: '10px 14px', minHeight: '96px' };

function getOutputFileName(outputPath) {
  if (!outputPath) {
    return '';
  }

  const normalizedPath = String(outputPath).replace(/\\/g, '/');
  const segments = normalizedPath.split('/');
  return segments[segments.length - 1] || outputPath;
}

function ReportChecker() {
  const [files, setFiles] = useState([]);
  const [ruleFile, setRuleFile] = useState(null);
  const [checklistFile, setChecklistFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [exportingRules, setExportingRules] = useState(false);

  const handleUpload = (file, target, onSuccess) => {
    if (!file.path) {
      message.error('当前环境未提供本地文件路径，无法执行桌面端文件处理。');
      return;
    }

    if (target === 'report') {
      const extension = file.name.toLowerCase();
      if (!extension.endsWith('.doc') && !extension.endsWith('.docx')) {
        message.error('当前后台仅支持 .doc / .docx 测试报告。');
        return;
      }

      const newItem = {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        name: file.name,
        path: file.path,
        status: 'pending',
        items: 0,
        outputPath: '',
        outputName: '',
        error: '',
        unmatchedItems: []
      };

      setFiles((prev) => {
        const exists = prev.some((item) => item.path === file.path);
        return exists ? prev : [newItem, ...prev];
      });
      message.success(`已添加报告: ${file.name}`);
    }

    if (target === 'rules') {
      setRuleFile({ name: file.name, path: file.path });
      message.success(`已上传规则: ${file.name}`);
    }

    if (target === 'checklist') {
      setChecklistFile({ name: file.name, path: file.path });
      message.success(`已上传 checklist: ${file.name}`);
    }

    if (onSuccess) {
      setTimeout(() => onSuccess('ok'), 0);
    }
  };

  const removeReport = (reportId) => {
    setFiles((prev) => prev.filter((item) => item.id !== reportId));
  };

  const openOutputFolder = async (record) => {
    if (!record.outputPath) {
      message.warning('该报告还没有生成输出文件。');
      return;
    }

    try {
      await window.electron.reportChecker.showOutputInFolder(record.outputPath);
    } catch (error) {
      message.error(error?.message || '打开输出目录失败');
    }
  };

  const showDetails = (record) => {
    Modal.info({
      title: record.name,
      width: 760,
      content: (
        <div style={{ marginTop: 16 }}>
          <Paragraph>
            <Text strong>状态：</Text> {record.status}
          </Paragraph>
          <Paragraph>
            <Text strong>命中规则数：</Text> {record.items || 0}
          </Paragraph>
          <Paragraph>
            <Text strong>输出文件：</Text> {record.outputPath || '尚未生成'}
          </Paragraph>
          {record.error ? (
            <Paragraph type="danger">
              <Text strong>错误：</Text> {record.error}
            </Paragraph>
          ) : null}
          {record.unmatchedItems?.length ? (
            <div>
              <Text strong>未命中规则：</Text>
              <div style={{ maxHeight: 240, overflow: 'auto', marginTop: 8, paddingRight: 8 }}>
                {record.unmatchedItems.slice(0, 20).map((item) => (
                  <Paragraph key={`${record.id}-${item.itemId}`} style={{ marginBottom: 8 }}>
                    {item.outputCell} - {item.checklistDesc} ({item.reason})
                  </Paragraph>
                ))}
              </div>
            </div>
          ) : null}
        </div>
      )
    });
  };

  const processReports = async () => {
    if (files.length === 0) {
      message.warning('请先上传至少一个 Word 测试报告。');
      return;
    }

    if (!checklistFile?.path) {
      message.warning('请先上传 checklist Excel 文件。');
      return;
    }

    setProcessing(true);
    setFiles((prev) => prev.map((item) => ({ ...item, status: 'processing', error: '' })));

    try {
      const response = await window.electron.reportChecker.processReports({
        reportPaths: files.map((item) => item.path),
        checklistPath: checklistFile.path,
        rulePath: ruleFile?.path || null
      });

      const resultMap = new Map(response.results.map((item) => [item.reportPath, item]));
      recordReportCheckResults(response.results);

      setFiles((prev) => prev.map((item) => {
        const result = resultMap.get(item.path);
        if (!result) {
          return item;
        }

        if (result.status === 'error') {
          return {
            ...item,
            status: 'error',
            error: result.error,
            items: 0,
            outputPath: '',
            outputName: '',
            unmatchedItems: []
          };
        }

        return {
          ...item,
          status: 'success',
          items: result.matchedItems,
          outputPath: result.outputPath,
          outputName: getOutputFileName(result.outputPath),
          unmatchedItems: result.unmatchedItems || [],
          error: ''
        };
      }));

      const successCount = response.results.filter((item) => item.status === 'success').length;
      const errorCount = response.results.length - successCount;
      message.success(`处理完成：成功 ${successCount} 份，失败 ${errorCount} 份。`);
    } catch (error) {
      const errorMessage = error?.message || '执行报告检查失败';
      setFiles((prev) => prev.map((item) => ({ ...item, status: 'error', error: errorMessage })));
      message.error(errorMessage);
    } finally {
      setProcessing(false);
    }
  };

  const exportRules = async () => {
    setExportingRules(true);

    try {
      const result = await window.electron.reportChecker.exportRules(ruleFile?.path || null);
      if (result?.canceled) {
        return;
      }

      message.success(`规则已导出到: ${result.filePath}`);
    } catch (error) {
      message.error(error?.message || '导出规则失败');
    } finally {
      setExportingRules(false);
    }
  };

  const columns = [
    {
      title: '文件名',
      dataIndex: 'name',
      key: 'name',
      width: 320,
      ellipsis: true,
      render: (text) => (
        <Text className="report-checker-table-text" ellipsis={{ tooltip: text }}>
          {text}
        </Text>
      )
    },
    {
      title: '状态',
      dataIndex: 'status',
      key: 'status',
      width: 100,
      render: (status) => {
        const colors = {
          success: 'green',
          error: 'red',
          pending: 'blue',
          processing: 'gold'
        };
        return <Tag color={colors[status] || 'default'}>{status}</Tag>;
      }
    },
    {
      title: '检查项',
      dataIndex: 'items',
      key: 'items',
      width: 96,
      render: (items) => <span>{items || 0}</span>
    },
    {
      title: '输出文件名',
      dataIndex: 'outputName',
      key: 'outputName',
      width: 420,
      ellipsis: true,
      render: (_, record) => {
        if (!record.outputName) {
          return <Text type="secondary">处理完成后显示</Text>;
        }

        return (
          <Text className="report-checker-table-text" ellipsis={{ tooltip: record.outputName }}>
            {record.outputName}
          </Text>
        );
      }
    },
    {
      title: '操作',
      key: 'action',
      width: 220,
      render: (_, record) => (
        <Space wrap={false}>
          <Button type="primary" size="small" onClick={() => showDetails(record)}>详情</Button>
          <Button size="small" disabled={!record.outputPath} onClick={() => openOutputFolder(record)}>打开目录</Button>
          <Button danger size="small" icon={<DeleteOutlined />} onClick={() => removeReport(record.id)}>删除</Button>
        </Space>
      )
    }
  ];

  return (
    <div className="page-container">
      <div className="report-checker-toolbar">
        <Button
          type="primary"
          icon={<CheckCircleOutlined />}
          loading={processing}
          onClick={processReports}
          className="report-checker-primary-action"
        >
          开始检查并生成 Excel
        </Button>
      </div>

      <Card 
        className="report-checker-card report-checker-main-card"
        title="上传测试报告"
        extra={
          <div className="report-checker-actions">
            <Upload
              customRequest={({ file, onSuccess }) => handleUpload(file, 'report', onSuccess)}
              multiple
              accept=".doc,.docx"
              showUploadList={false}
            >
              <Button icon={<UploadOutlined />} className="report-checker-upload-action report-checker-section-action">
                上传报告
              </Button>
            </Upload>
          </div>
        }
      >
        <p style={{ marginBottom: '24px', color: '#8c8c8c' }}>
          上传真实 .doc/.docx 测试报告和 I 列为空的 checklist，系统会按 moto_rules_for_analysis.json5 提取报告数据并回填到 checklist，最终在报告目录下生成新的 Excel 文件。
        </p>
        <p style={{ marginTop: '-12px', marginBottom: '24px', color: '#8c8c8c' }}>
          注意：.doc 报告会先做转换，通常比 .docx 慢；处理完成后可在表格里查看输出文件名，或点击“打开目录”直接定位结果。
        </p>

        {files.length === 0 ? (
          <Upload.Dragger
            customRequest={({ file, onSuccess }) => handleUpload(file, 'report', onSuccess)}
            multiple
            accept=".doc,.docx"
            showUploadList={false}
            className="report-checker-upload report-checker-upload-report"
            style={{ padding: '16px 18px', minHeight: '108px' }}
          >
            <UploadOutlined style={{ fontSize: '36px', color: '#bfbfbf', marginBottom: '10px' }} />
            <p style={{ fontSize: '15px', color: '#262626', marginBottom: '6px' }}>
              拖拽报告文件到此处，或点击上方按钮上传
            </p>
            <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
              当前支持格式: Word (.doc, .docx)
            </p>
          </Upload.Dragger>
        ) : (
          <Table
            className="report-checker-table"
            columns={columns}
            dataSource={files}
            rowKey="id"
            scroll={{ x: 1180 }}
            pagination={{ pageSize: 10 }}
          />
        )}
      </Card>

      <Card
        className="report-checker-card report-checker-section-card"
        title="上传 checklist"
        style={{ marginTop: '24px' }}
        extra={
          <Upload
            customRequest={({ file, onSuccess }) => handleUpload(file, 'checklist', onSuccess)}
            accept=".xlsx,.xls"
            showUploadList={false}
          >
            <Button icon={<UploadOutlined />} className="report-checker-upload-action report-checker-section-action">
              上传 checklist
            </Button>
          </Upload>
        }
      >
        <Upload.Dragger
          customRequest={({ file, onSuccess }) => handleUpload(file, 'checklist', onSuccess)}
          accept=".xlsx,.xls"
          showUploadList={false}
          className="report-checker-upload"
          style={compactUploadDraggerStyle}
        >
          {checklistFile ? (
            <div style={{ padding: '8px 0' }}>
              <Text strong style={{ display: 'block', fontSize: '16px', color: '#262626', marginBottom: '6px' }}>
                已选择 checklist
              </Text>
              <Text style={{ fontSize: '14px', color: '#595959', wordBreak: 'break-all' }}>
                {checklistFile.name}
              </Text>
            </div>
          ) : (
            <>
              <UploadOutlined style={{ fontSize: '24px', color: '#bfbfbf', marginBottom: '6px' }} />
              <p style={{ fontSize: '14px', color: '#262626', marginBottom: '3px' }}>
                拖拽 checklist 文件到此处，或点击上传
              </p>
              <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
                支持格式: Excel (.xlsx, .xls)
              </p>
            </>
          )}
        </Upload.Dragger>
        {checklistFile && (
          <p style={{ marginTop: '12px', color: '#595959' }}>
            已选择: {checklistFile.name}
          </p>
        )}
      </Card>

      <Card
        className="report-checker-card report-checker-section-card"
        title="上传规则"
        style={{ marginTop: '24px' }}
        extra={
          <Space>
            <Button
              icon={<ExportOutlined />}
              className="report-checker-upload-action report-checker-section-action"
              loading={exportingRules}
              onClick={exportRules}
            >
              导出规则
            </Button>
            <Upload
              customRequest={({ file, onSuccess }) => handleUpload(file, 'rules', onSuccess)}
              accept=".json,.json5"
              showUploadList={false}
            >
              <Button icon={<UploadOutlined />} className="report-checker-upload-action report-checker-section-action">
                上传规则
              </Button>
            </Upload>
          </Space>
        }
      >
        <Upload.Dragger
          customRequest={({ file, onSuccess }) => handleUpload(file, 'rules', onSuccess)}
          accept=".json,.json5"
          showUploadList={false}
          className="report-checker-upload"
          style={compactUploadDraggerStyle}
        >
          <UploadOutlined style={{ fontSize: '24px', color: '#bfbfbf', marginBottom: '6px' }} />
          <p style={{ fontSize: '14px', color: '#262626', marginBottom: '3px' }}>
            拖拽规则文件到此处，或点击上传
          </p>
          <p style={{ color: '#8c8c8c', fontSize: '12px' }}>
            支持格式: JSON / JSON5；不上传时默认使用内置规则
          </p>
        </Upload.Dragger>
        {ruleFile && (
          <p style={{ marginTop: '12px', color: '#595959' }}>
            已选择: {ruleFile.name}
          </p>
        )}
      </Card>

      <Card title="检查规则" style={{ marginTop: '24px' }}>
        <ul style={{ paddingLeft: '20px' }}>
          <li>报告解析当前使用主进程后台执行，避免浏览器沙箱限制本地文件读写。</li>
          <li>默认读取内置 JSON5 规则，也可以手动上传其它规则文件覆盖。</li>
          <li>“导出规则”会导出当前使用的规则文件；若未上传自定义规则，则导出内置默认规则。</li>
          <li>.doc 报告会优先在后台临时转为 .docx 后解析，解析完成后自动清理临时文件。</li>
          <li>生成的 Excel 是在空白 checklist 基础上回填数据后的结果文件，默认保存到测试报告原目录。</li>
        </ul>
      </Card>
    </div>
  );
}

export default ReportChecker;
