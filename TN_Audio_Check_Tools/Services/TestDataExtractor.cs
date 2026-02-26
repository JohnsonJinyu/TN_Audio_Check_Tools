using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TN_Audio_Check_Tools.Services
{
    /// <summary>
    /// 从 Word .docx 文档中提取表格和数据的工具类
    /// </summary>
    public static class TestDataExtractor
    {
        /// <summary>
        /// 从 .docx 提取"Status Overview"标题下的第一个表格为 DataTable
        /// </summary>
        public static DataTable ExtractStatusOverviewTableAsDataTable(string docxPath)
        {
            if (string.IsNullOrWhiteSpace(docxPath)) throw new ArgumentNullException(nameof(docxPath));
            if (!File.Exists(docxPath)) throw new FileNotFoundException("Docx file not found.", docxPath);

            var defaultKeywords = new[] { "Status Overview", "Status overview", "Status Overview:", "状态概览" };
            var tableData = ExtractTablesAfterHeader(docxPath, defaultKeywords, maxTables: 1);

            if (tableData == null || tableData.Count == 0)
            {
                // 返回空DataTable如果找不到表格
                return new DataTable();
            }

            var table = tableData[0];
            var result = new DataTable();

            if (table.Count == 0)
                return result;

            // 第一行为列标题
            var headers = table[0];
            foreach (var header in headers)
            {
                var columnName = string.IsNullOrWhiteSpace(header) ? $"Column{result.Columns.Count + 1}" : header;
                result.Columns.Add(columnName);
            }

            // 后续行为数据
            for (int rowIndex = 1; rowIndex < table.Count; rowIndex++)
            {
                var row = table[rowIndex];
                var dataRow = result.NewRow();
                for (int colIndex = 0; colIndex < headers.Count; colIndex++)
                {
                    dataRow[colIndex] = colIndex < row.Count ? row[colIndex] : "";
                }
                result.Rows.Add(dataRow);
            }

            return result;
        }

        /// <summary>
        /// 提取 Word 文档中在指定标题后的所有表格
        /// </summary>
        public static List<List<List<string>>> ExtractTablesAfterHeader(string docxPath, IEnumerable<string> headerKeywords, int maxTables = 1)
        {
            if (string.IsNullOrWhiteSpace(docxPath)) throw new ArgumentNullException(nameof(docxPath));
            if (!File.Exists(docxPath)) throw new FileNotFoundException("Docx file not found.", docxPath);

            var keywords = (headerKeywords ?? Array.Empty<string>()).Where(k => !string.IsNullOrWhiteSpace(k)).Select(k => k.Trim()).ToList();
            var foundTables = new List<List<List<string>>>();

            using var doc = WordprocessingDocument.Open(docxPath, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return foundTables;

            bool foundHeader = keywords.Count == 0;
            foreach (var element in body.Elements())
            {
                if (!foundHeader)
                {
                    if (element is Paragraph p)
                    {
                        var text = string.Concat(p.Descendants<Text>().Select(t => t.Text));
                        foreach (var kw in keywords)
                        {
                            if (!string.IsNullOrEmpty(kw) && text.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                foundHeader = true;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    if (element is Table tbl)
                    {
                        var tableData = new List<List<string>>();
                        foreach (var tr in tbl.Elements<TableRow>())
                        {
                            var row = new List<string>();
                            foreach (var tc in tr.Elements<TableCell>())
                            {
                                var cellText = string.Concat(tc.Descendants<Text>().Select(t => t.Text));
                                row.Add(cellText?.Trim() ?? string.Empty);
                            }
                            tableData.Add(row);
                        }
                        foundTables.Add(tableData);
                        if (foundTables.Count >= maxTables) break;
                    }
                }
            }

            // 如果没有找到表格，返回第一个表格作为备选
            if (foundTables.Count == 0)
            {
                foreach (var tbl in body.Elements<Table>().Take(maxTables))
                {
                    var tableData = new List<List<string>>();
                    foreach (var tr in tbl.Elements<TableRow>())
                    {
                        var row = new List<string>();
                        foreach (var tc in tr.Elements<TableCell>())
                        {
                            var cellText = string.Concat(tc.Descendants<Text>().Select(t => t.Text));
                            row.Add(cellText?.Trim() ?? string.Empty);
                        }
                        tableData.Add(row);
                    }
                    foundTables.Add(tableData);
                }
            }

            return foundTables;
        }

        // 保留其他方法的存根以兼容现有代码
        public static string ExtractPlainTextFromDocx(string docxPath) => throw new NotImplementedException();

        public static Dictionary<string, string> ParseKeyValuePairs(string plainText) => new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static Dictionary<string, string> ExtractKeyValuePairsFromDocx(string docxPath) => new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static List<Dictionary<string, string>> ExtractStatusOverviewRecords(string docxPath, IEnumerable<string>? headerKeywords = null) => new List<Dictionary<string, string>>();

        public static List<Dictionary<string, string>> ExtractFromMultipleFiles(IEnumerable<string> docxPaths) => new List<Dictionary<string, string>>();

        public static void ExportToCsv(string csvPath, IEnumerable<Dictionary<string, string>> records, bool includeBom = true) { }
    }
}
