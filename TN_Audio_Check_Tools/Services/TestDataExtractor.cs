using System;
using System.Collections.Generic;

namespace TN_Audio_Check_Tools.Services
{
    // 业务实现已清空。保留类和方法签名以避免引用中断。
    public static class TestDataExtractor
    {
        public static string ExtractPlainTextFromDocx(string docxPath) => throw new NotImplementedException();

        public static Dictionary<string, string> ParseKeyValuePairs(string plainText) => new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static Dictionary<string, string> ExtractKeyValuePairsFromDocx(string docxPath) => new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static List<List<List<string>>> ExtractTablesAfterHeader(string docxPath, IEnumerable<string> headerKeywords, int maxTables = 1) => new List<List<List<string>>>();

        public static List<Dictionary<string, string>> ExtractStatusOverviewRecords(string docxPath, IEnumerable<string>? headerKeywords = null) => new List<Dictionary<string, string>>();

        public static List<Dictionary<string, string>> ExtractFromMultipleFiles(IEnumerable<string> docxPaths) => new List<Dictionary<string, string>>();

        public static void ExportToCsv(string csvPath, IEnumerable<Dictionary<string, string>> records, bool includeBom = true) { }
    }
}
