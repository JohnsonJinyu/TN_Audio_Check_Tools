using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace TN_Audio_Check_Tools.Services
{
    /// <summary>
    /// 管理整个处理流程的阶段和状态5
    /// </summary>
    public class ProcessingManager : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private ProcessingPhase _currentPhase = ProcessingPhase.Idle;
        private int _currentProgress = 0;
        private int _totalItems = 0;
        private string _statusMessage = "就绪";
        private string _currentFileName = "";
        private bool _isProcessing = false;

        public ObservableCollection<string> Logs { get; } = new ObservableCollection<string>();

        /// <summary>
        /// 当前处理阶段
        /// </summary>
        public ProcessingPhase CurrentPhase
        {
            get => _currentPhase;
            set => SetProperty(ref _currentPhase, value);
        }

        /// <summary>
        /// 当前进度（0 到 TotalItems）
        /// </summary>
        public int CurrentProgress
        {
            get => _currentProgress;
            set => SetProperty(ref _currentProgress, value);
        }

        /// <summary>
        /// 总项目数
        /// </summary>
        public int TotalItems
        {
            get => _totalItems;
            set => SetProperty(ref _totalItems, value);
        }

        /// <summary>
        /// 状态栏显示的消息
        /// </summary>
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        /// <summary>
        /// 当前处理的文件名
        /// </summary>
        public string CurrentFileName
        {
            get => _currentFileName;
            set => SetProperty(ref _currentFileName, value);
        }

        /// <summary>
        /// 是否正在处理
        /// </summary>
        public bool IsProcessing
        {
            get => _isProcessing;
            set => SetProperty(ref _isProcessing, value);
        }

        /// <summary>
        /// 初始化处理流程
        /// </summary>
        public void Initialize(int totalItems)
        {
            Logs.Clear();
            TotalItems = totalItems;
            CurrentProgress = 0;
            CurrentFileName = "";
            IsProcessing = true;
        }

        /// <summary>
        /// 开始新阶段
        /// </summary>
        public void StartPhase(ProcessingPhase phase, string phaseDescription)
        {
            CurrentPhase = phase;
            CurrentProgress = 0;
            CurrentFileName = "";
            AddLog($"========== {phaseDescription} ==========");
            UpdateStatus();
        }

        /// <summary>
        /// 更新当前项目的进度
        /// </summary>
        public void UpdateItemProgress(int itemIndex, string fileName)
        {
            CurrentProgress = itemIndex + 1;
            CurrentFileName = fileName;
            UpdateStatus();
        }

        /// <summary>
        /// 完成处理
        /// </summary>
        public void Complete()
        {
            IsProcessing = false;
            CurrentPhase = ProcessingPhase.Completed;
            StatusMessage = "已完成（已收集结果，但暂不保存为 Excel）";
        }

        /// <summary>
        /// 添加日志
        /// </summary>
        public void AddLog(string message)
        {
            Logs.Add(message);
        }

        /// <summary>
        /// 获取当前阶段的中文描述
        /// </summary>
        private string GetPhaseDescription()
        {
            return CurrentPhase switch
            {
                ProcessingPhase.Converting => "转换",
                ProcessingPhase.Extracting => "提取",
                ProcessingPhase.Completed => "已完成",
                _ => "处理中"
            };
        }

        /// <summary>
        /// 更新状态消息
        /// </summary>
        private void UpdateStatus()
        {
            string phaseDesc = GetPhaseDescription();
            if (TotalItems > 0)
            {
                StatusMessage = $"{phaseDesc} {CurrentProgress}/{TotalItems}：{CurrentFileName}";
            }
            else
            {
                StatusMessage = phaseDesc;
            }
        }

        protected void SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = "")
        {
            if (!Equals(field, value))
            {
                field = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }

    /// <summary>
    /// 处理阶段枚举
    /// </summary>
    public enum ProcessingPhase
    {
        Idle,           // 空闲
        Converting,     // 转换阶段
        Extracting,     // 提取阶段
        Completed       // 完成
    }
}
