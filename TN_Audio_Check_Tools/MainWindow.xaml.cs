using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Data;
using TN_Audio_Check_Tools.Services;
using OfficeOpenXml;

namespace TN_Audio_Check_Tools
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private record FileEntry(string FullPath)
        {
            public string FullPath { get; } = FullPath;
            public string FileName => System.IO.Path.GetFileName(FullPath);
            public override string ToString() => FileName;
        }

        private void ResultsList_CopyClick(object sender, RoutedEventArgs e)
        {
            CopySelectedResultsToClipboard();
        }

        private void ResultsList_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.C)
            {
                CopySelectedResultsToClipboard();
                e.Handled = true;
            }
        }

        private void CopySelectedResultsToClipboard()
        {
            try
            {
                if (ResultsList.SelectedItems == null || ResultsList.SelectedItems.Count == 0) return;
                var sb = new StringBuilder();
                foreach (var it in ResultsList.SelectedItems)
                {
                    if (it != null) sb.AppendLine(it.ToString());
                }
                var text = sb.ToString();
                if (!string.IsNullOrEmpty(text)) Clipboard.SetText(text);
            }
            catch
            {
                // ignore clipboard errors  ff
            }
        }

        private ObservableCollection<FileEntry> FileGroupA { get; } = new ObservableCollection<FileEntry>();
        private ObservableCollection<FileEntry> FileGroupB { get; } = new ObservableCollection<FileEntry>();
        private ObservableCollection<string> Results { get; } = new ObservableCollection<string>();
        private ProcessingManager ProcessingMgr { get; } = new ProcessingManager();

        public MainWindow()
        {
            InitializeComponent();

            ListA.ItemsSource = FileGroupA;
            ListB.ItemsSource = FileGroupB;
            ResultsList.ItemsSource = Results;

            // 绑定 ProcessingManager 的日志到 Results
            ProcessingMgr.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(ProcessingManager.StatusMessage))
                {
                    StatusText.Text = ProcessingMgr.StatusMessage;
        }
            };
            ProcessingMgr.Logs.CollectionChanged += (s, e) =>
            {
                if (e.NewItems != null)
                {
                    foreach (var item in e.NewItems)
                    {
                        Results.Add(item?.ToString() ?? "");
                    }
                }
            };
        }

        private void AddFilesA_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog() { Multiselect = true, Title = "选择 测试报告 文件" };
            if (dlg.ShowDialog() == true)
            {
                AddFilesToCollection(dlg.FileNames, FileGroupA);
            }
        }

        private void AddFilesB_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog() { Multiselect = false, Title = "选择 汇总表模板 文件" };
            if (dlg.ShowDialog() == true)
            {
                // only single template file allowed
                FileGroupB.Clear();
                FileGroupB.Add(new FileEntry(dlg.FileName));
            }
        }

        private void RemoveSelectedA_Click(object sender, RoutedEventArgs e)
        {
            var items = ListA.SelectedItems.Cast<FileEntry>().ToList();
            foreach (var it in items) FileGroupA.Remove(it);
        }

        private void RemoveSelectedB_Click(object sender, RoutedEventArgs e)
        {
            var items = ListB.SelectedItems.Cast<FileEntry>().ToList();
            foreach (var it in items) FileGroupB.Remove(it);
        }

        private void ClearA_Click(object sender, RoutedEventArgs e) => FileGroupA.Clear();
        private void ClearB_Click(object sender, RoutedEventArgs e) => FileGroupB.Clear();

        private void List_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void ListA_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddFilesToCollection(files, FileGroupA);
            }
        }

        private void ListB_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    FileGroupB.Clear();
                    FileGroupB.Add(new FileEntry(files[0]));
            }
        }
        }

        private void AddFilesToCollection(IEnumerable<string> files, ObservableCollection<FileEntry> collection)
        {
            foreach (var f in files)
            {
                if (!collection.Any(x => string.Equals(x.FullPath, f, StringComparison.OrdinalIgnoreCase)))
                {
                    collection.Add(new FileEntry(f));
                }
            }
        }

        private void Process_Click(object sender, RoutedEventArgs e)
        {
            Results.Clear();

            var total = FileGroupA.Count + FileGroupB.Count;
            if (total == 0)
            {
                StatusText.Text = "无文件可处理";
                ProgressBar.Value = 0;
                return;
            }

            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = total;
            int processed = 0;

            Results.Add($"Group A: {FileGroupA.Count} file(s)");
            foreach (var f in FileGroupA)
            {
                Results.Add($"A: {f.FileName} ({f.FullPath})");
                processed++;
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Value = processed;
                    StatusText.Text = $"处理 {processed}/{total}...";
                }, System.Windows.Threading.DispatcherPriority.Background);
            }

            Results.Add($"Group B: {FileGroupB.Count} file(s)");
            foreach (var f in FileGroupB)
            {
                Results.Add($"B: {f.FileName} ({f.FullPath})");
                processed++;
                Dispatcher.Invoke(() =>
                {
                    ProgressBar.Value = processed;
                    StatusText.Text = $"处理 {processed}/{total}...";
                }, System.Windows.Threading.DispatcherPriority.Background);
            }

            StatusText.Text = "处理完成";
        }

        private void ClearResults_Click(object sender, RoutedEventArgs e) => Results.Clear();

        // .doc -> .docx 的转换逻辑已移至 Services.DocConverter.ConvertDocToDocx

        private async void ProcessResults_Click(object sender, RoutedEventArgs e)
        {
            var files = FileGroupA.Select(f => f.FullPath).ToList();
            if (files.Count == 0)
            {
                StatusText.Text = "请先添加测试报告文件到列表";
                return;
            }

            Results.Clear();
            ProcessingMgr.Initialize(files.Count);
            ProgressBar.Maximum = files.Count;
            ProgressBar.Value = 0;

            var convertedFiles = new List<string>();

            // ===== Phase 1: Convert .doc files to .docx =====
            ProcessingMgr.StartPhase(ProcessingPhase.Converting, "第一阶段：格式转换");

            for (int i = 0; i < files.Count; i++)
            {
                var path = files[i];
                string convertedPath = path;

                try
                {
                    if (System.IO.Path.GetExtension(path).Equals(".doc", StringComparison.OrdinalIgnoreCase))
                    {
                        ProcessingMgr.UpdateItemProgress(i, System.IO.Path.GetFileName(path));
                        ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] {System.IO.Path.GetFileName(path)}: 开始转换 .doc -> .docx");

                        var conv = DocConverter.ConvertDocToDocxWithDetails(path);
                        var tempConverted = conv.Path;
                        var method = conv.Method;
                        var error = conv.Error;

                        if (!string.IsNullOrEmpty(tempConverted))
                        {
                            bool fileExists = false;
                            long fileSize = 0;
                            try
                            {
                                if (File.Exists(tempConverted))
                                {
                                    var fi = new FileInfo(tempConverted);
                                    fileSize = fi.Length;
                                    fileExists = (fileSize > 0);
                                }
                            }
                            catch (Exception fileCheckEx)
                            {
                                ProcessingMgr.AddLog($"[诊断] 文件检查异常: {fileCheckEx.Message}");
                            }

                            if (fileExists)
                            {
                                convertedPath = tempConverted;
                                ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] 转换成功 ({method}) -> {System.IO.Path.GetFileName(tempConverted)} ({fileSize} 字节)");
                            }
                            else
                            {
                                ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] 转换声称成功但文件未找到或为空");
                                ProcessingMgr.AddLog($"[诊断] 期望路径: {tempConverted}");
                                ProcessingMgr.AddLog($"[诊断] 方法: {method}, 错误: {error}");
                                convertedPath = path;
                            }
                        }
                        else
                        {
                            ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] 转换失败：方法={method} 错误={error}");
                            convertedPath = path;
                        }
                    }
                    else
                    {
                        ProcessingMgr.UpdateItemProgress(i, System.IO.Path.GetFileName(path));
                        ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] {System.IO.Path.GetFileName(path)} 已是 .docx 格式，无需转换");
                    }
                }
                catch (Exception ex)
                {
                    ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] 转换过程异常: {ex.Message}");
                    convertedPath = path;
                }

                convertedFiles.Add(convertedPath);
                ProgressBar.Value = i + 1;
                await Task.Delay(20);
            }

            ProcessingMgr.AddLog(string.Empty);
            await Task.Delay(300);

            // ===== Phase 2: Extract data from all files =====
            ProcessingMgr.StartPhase(ProcessingPhase.Extracting, "第二阶段：数据提取");
            ProgressBar.Value = 0;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var allDataTables = new List<(string fileName, DataTable table)>();

            for (int i = 0; i < convertedFiles.Count; i++)
            {
                var usedPath = convertedFiles[i];
                var originalFileName = System.IO.Path.GetFileName(files[i]);

                try
                {
                    ProcessingMgr.UpdateItemProgress(i, originalFileName);

                    // 提取"Status Overview"表格为DataTable
                    var statusTable = await Task.Run(() => TestDataExtractor.ExtractStatusOverviewTableAsDataTable(usedPath));

                    if (statusTable != null && statusTable.Rows.Count > 0)
                    {
                        allDataTables.Add((originalFileName, statusTable));
                        ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] {originalFileName}: 成功提取表格，共 {statusTable.Rows.Count} 行");
                    }
                    else
                    {
                        ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] {originalFileName}: 未找到或表格为空");
                    }
                }
                catch (Exception ex)
                {
                    ProcessingMgr.AddLog($"[{i + 1}/{files.Count}] {originalFileName}: 提取失败 {ex.Message}");
                }

                ProgressBar.Value = i + 1;
                await Task.Delay(20);
            }

            ProcessingMgr.AddLog(string.Empty);
            ProcessingMgr.AddLog($"总计：已提取 {allDataTables.Count} 个表格");

            // ===== Phase 3: Save to Excel =====
            if (allDataTables.Count > 0)
            {
                ProcessingMgr.StartPhase(ProcessingPhase.Extracting, "第三阶段：保存结果");

                var sfd = new SaveFileDialog() 
                { 
                    Filter = "Excel 工作簿|*.xlsx", 
                    FileName = $"StatusOverview_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        await Task.Run(() => SaveDataTablesToExcel(sfd.FileName, allDataTables));
                        ProcessingMgr.AddLog($"已保存到: {sfd.FileName}");
                        StatusText.Text = "已保存为 Excel";
                        MessageBox.Show(this, $"已成功保存 {allDataTables.Count} 个表格到:\n{sfd.FileName}", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        ProcessingMgr.AddLog($"保存失败: {ex.Message}");
                        MessageBox.Show(this, $"保存失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    ProcessingMgr.AddLog("已取消保存");
                }
            }

            ProcessingMgr.Complete();
        }

        /// <summary>
        /// 将多个DataTable保存到Excel文件
        /// </summary>
        private void SaveDataTablesToExcel(string filePath, List<(string fileName, DataTable table)> dataTables)
        {
            using var package = new ExcelPackage(new FileInfo(filePath));

            // 删除默认工作表
            while (package.Workbook.Worksheets.Count > 0)
            {
                package.Workbook.Worksheets.Delete(package.Workbook.Worksheets.First());
            }

            foreach (var (fileName, table) in dataTables)
            {
                // 创建工作表（去除文件扩展名，限制长度）
                var sheetName = System.IO.Path.GetFileNameWithoutExtension(fileName);
                sheetName = sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;

                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                // 写入列标题
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = table.Columns[col].ColumnName;
                    // 设置标题行为粗体
                    worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                }

                // 写入数据行
                for (int row = 0; row < table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = table.Rows[row][col];
                    }
                }

                // 自动调整列宽
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    worksheet.Column(col).AutoFit();
                }
            }

            package.Save();
        }
    }
}
