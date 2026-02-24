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

        private ObservableCollection<FileEntry> FileGroupA { get; } = new ObservableCollection<FileEntry>();
        private ObservableCollection<FileEntry> FileGroupB { get; } = new ObservableCollection<FileEntry>();
        private ObservableCollection<string> Results { get; } = new ObservableCollection<string>();

        public MainWindow()
        {
            InitializeComponent();

            ListA.ItemsSource = FileGroupA;
            ListB.ItemsSource = FileGroupB;
            ResultsList.ItemsSource = Results;
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
            var dlg = new OpenFileDialog() { Multiselect = true, Title = "选择 文件 (Group B)" };
            if (dlg.ShowDialog() == true)
            {
                AddFilesToCollection(dlg.FileNames, FileGroupB);
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
                AddFilesToCollection(files, FileGroupB);
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
            Results.Add($"Group A: {FileGroupA.Count} file(s)");
            foreach (var f in FileGroupA)
            {
                Results.Add($"A: {f.FileName} ({f.FullPath})");
            }

            Results.Add($"Group B: {FileGroupB.Count} file(s)");
            foreach (var f in FileGroupB)
            {
                Results.Add($"B: {f.FileName} ({f.FullPath})");
            }
        }

        private void ClearResults_Click(object sender, RoutedEventArgs e) => Results.Clear();
    }
}
