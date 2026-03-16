const XLSX = require('xlsx');

function createXlsxReportSource() {
  function normalizeCellValue(value) {
    if (value === undefined || value === null) {
      return '';
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    return value;
  }

  function stringifyCellValue(value) {
    const normalized = normalizeCellValue(value);
    if (normalized === '') {
      return '';
    }

    return typeof normalized === 'string' ? normalized.trim() : String(normalized);
  }

  function joinNonEmpty(parts) {
    return parts
      .map((part) => stringifyCellValue(part))
      .filter(Boolean)
      .join(' | ');
  }

  function buildSemanticAliases(text) {
    const sourceText = stringifyCellValue(text);
    if (!sourceText) {
      return [];
    }

    const aliases = [];

    if (/\bg-mos\b/i.test(sourceText) && !/\bglobal\b/i.test(sourceText)) {
      aliases.push('Global');
    }

    if (/\bn-mos\b/i.test(sourceText) && !/\bnoise\b/i.test(sourceText)) {
      aliases.push('Noise');
    }

    if (/\bs-mos\b/i.test(sourceText) && !/\bspeech\b/i.test(sourceText)) {
      aliases.push('Speech');
    }

    if (/P\.863/i.test(sourceText)) {
      if (!/MOS-LQO/i.test(sourceText)) {
        aliases.push('MOS-LQO');
      }
      if (!/informative/i.test(sourceText)) {
        aliases.push('informative');
      }
      if (!/POLQA/i.test(sourceText)) {
        aliases.push('POLQA');
      }
    }

    return aliases;
  }

  function buildRowText(parts, aliasSourceText) {
    const aliases = buildSemanticAliases(aliasSourceText);
    return joinNonEmpty([...parts, ...aliases]);
  }

  function readDetailedRows(workbook) {
    const worksheet = workbook.Sheets.Detailed;
    if (!worksheet) {
      return [];
    }

    return XLSX.utils.sheet_to_json(worksheet, { defval: '' });
  }

  function readValuesMatrix(workbook) {
    const worksheet = workbook.Sheets.Values;
    if (!worksheet) {
      return [];
    }

    return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  }

  function buildDetailedRowContexts(rows) {
    return rows.map((row, index) => {
      const descriptorText = joinNonEmpty([
        row.SMD,
        row.Name,
        row.ResultName,
        row.BGNScenario,
        row.Description,
        row.Direction,
        row.VolumeCTRL,
        row.RequirementType,
        row.Comment,
        row.UseCase,
        row.Bandwidth
      ]);
      const inlineMetricText = joinNonEmpty([row.Name, normalizeCellValue(row.Value), row.Unit]);

      const headers = [
        'Descriptor',
        'Value',
        'Status',
        'Unit',
        'Description',
        'Comment',
        'Direction',
        'VolumeCTRL',
        'Bandwidth',
        'BGNScenario',
        'RequirementType',
        'ResultName',
        'UseCase',
        'Measurement Object'
      ];

      const cells = [
        descriptorText,
        normalizeCellValue(row.Value),
        row.Status,
        row.Unit,
        row.Description,
        row.Comment,
        row.Direction,
        row.VolumeCTRL,
        row.Bandwidth,
        row.BGNScenario,
        row.RequirementType,
        row.ResultName,
        row.UseCase,
        row['Measurement Object']
      ];

      return {
        tableIndex: 0,
        rowIndex: index + 1,
        headers,
        cells,
        text: buildRowText([
          descriptorText,
          inlineMetricText,
          row.Status,
          row.Description,
          row.Direction,
          row.VolumeCTRL,
          row.BGNScenario,
          row.RequirementType,
          row.ResultName,
          row.UseCase,
          row.Bandwidth,
          row['Measurement Object']
        ], `${row.SMD} ${row.Name} ${row.Description} ${row.ResultName}`),
        sourceKind: 'xlsx-detailed'
      };
    }).filter((rowContext) => rowContext.text);
  }

  function buildValuesDerivedRows(valuesMatrix) {
    if (!Array.isArray(valuesMatrix) || valuesMatrix.length < 3) {
      return [];
    }

    const headerRow = valuesMatrix[0] || [];
    const unitRow = valuesMatrix[1] || [];
    const rows = [];

    for (let rowIndex = 2; rowIndex < valuesMatrix.length; rowIndex += 1) {
      const currentRow = valuesMatrix[rowIndex] || [];
      const measurementObject = currentRow[0];

      for (let columnIndex = 1; columnIndex < headerRow.length; columnIndex += 1) {
        const descriptor = headerRow[columnIndex];
        const value = normalizeCellValue(currentRow[columnIndex]);
        if (!descriptor || value === '') {
          continue;
        }

        const headers = ['Descriptor', 'Value', 'Unit', 'Measurement Object'];
        const cells = [descriptor, value, unitRow[columnIndex], measurementObject];

        rows.push({
          tableIndex: 1,
          rowIndex: rows.length + 1,
          headers,
          cells,
          text: buildRowText([
            descriptor,
            joinNonEmpty([descriptor, value, unitRow[columnIndex]]),
            measurementObject
          ], descriptor),
          sourceKind: 'xlsx-values'
        });
      }
    }

    return rows;
  }

  async function parseXlsxReport(reportPath) {
    const workbook = XLSX.readFile(reportPath, { cellDates: true });
    const detailedRows = readDetailedRows(workbook);
    if (detailedRows.length === 0) {
      throw new Error('未在 Excel 测试报告中找到 Detailed 工作表或有效数据。');
    }

    const detailedRowContexts = buildDetailedRowContexts(detailedRows);
    const valuesDerivedRows = buildValuesDerivedRows(readValuesMatrix(workbook));
    const lines = [...detailedRowContexts, ...valuesDerivedRows]
      .map((rowContext) => rowContext.text)
      .filter(Boolean);

    const detailedHeaders = [
      'Descriptor',
      'Value',
      'Status',
      'Unit',
      'Description',
      'Comment',
      'Direction',
      'VolumeCTRL',
      'Bandwidth',
      'BGNScenario',
      'RequirementType',
      'ResultName',
      'UseCase',
      'Measurement Object'
    ];

    return {
      reportFormat: 'xlsx',
      rawText: lines.join('\n'),
      html: '',
      lines,
      ambientNoiseBlocks: [],
      tables: [
        {
          tableIndex: 0,
          rows: [detailedHeaders, ...detailedRowContexts.map((rowContext) => rowContext.cells)]
        }
      ],
      tableRows: detailedRowContexts,
      fallbackRows: valuesDerivedRows
    };
  }

  return {
    parseXlsxReport
  };
}

module.exports = {
  createXlsxReportSource
};