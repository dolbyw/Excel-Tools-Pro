using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace ExcelToolsPro.Models;

/// <summary>
/// 处理任务模型
/// </summary>
public class ProcessingTask : INotifyPropertyChanged, IDisposable
{
    private TaskStatus _status = TaskStatus.Pending;
    private float _progress = 0f;
    private string _currentFile = string.Empty;
    private string _errorMessage = string.Empty;
    private bool _disposed = false;

    public ProcessingTask()
    {
        Id = Guid.NewGuid().ToString();
        CreatedAt = DateTime.UtcNow;
        UpdatedAt = DateTime.UtcNow;
        InputFiles = new ObservableCollection<FileInfo>();
        Logs = new ObservableCollection<ProcessingLog>();
        CancellationTokenSource = new CancellationTokenSource();
    }

    /// <summary>
    /// 任务ID
    /// </summary>
    public string Id { get; }

    /// <summary>
    /// 任务类型
    /// </summary>
    public TaskType TaskType { get; set; }

    /// <summary>
    /// 任务状态
    /// </summary>
    public TaskStatus Status
    {
        get => _status;
        set
        {
            if (SetProperty(ref _status, value))
            {
                UpdatedAt = DateTime.UtcNow;
                OnPropertyChanged(nameof(StatusText));
            }
        }
    }

    /// <summary>
    /// 进度百分比 (0-100)
    /// </summary>
    public float Progress
    {
        get => _progress;
        set
        {
            if (SetProperty(ref _progress, Math.Clamp(value, 0f, 100f)))
            {
                UpdatedAt = DateTime.UtcNow;
                OnPropertyChanged(nameof(ProgressText));
            }
        }
    }

    /// <summary>
    /// 当前处理的文件
    /// </summary>
    public string CurrentFile
    {
        get => _currentFile;
        set => SetProperty(ref _currentFile, value);
    }

    /// <summary>
    /// 输入文件列表
    /// </summary>
    public ObservableCollection<FileInfo> InputFiles { get; }

    /// <summary>
    /// 输出路径
    /// </summary>
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>
    /// 配置信息
    /// </summary>
    public AppConfig? Config { get; set; }

    /// <summary>
    /// 创建时间
    /// </summary>
    public DateTime CreatedAt { get; }

    /// <summary>
    /// 更新时间
    /// </summary>
    public DateTime UpdatedAt { get; private set; }

    /// <summary>
    /// 日志列表
    /// </summary>
    public ObservableCollection<ProcessingLog> Logs { get; }

    /// <summary>
    /// 取消令牌源
    /// </summary>
    public CancellationTokenSource CancellationTokenSource { get; }

    /// <summary>
    /// 错误消息
    /// </summary>
    public string ErrorMessage
    {
        get => _errorMessage;
        set => SetProperty(ref _errorMessage, value);
    }

    /// <summary>
    /// 状态文本
    /// </summary>
    public string StatusText => Status switch
    {
        TaskStatus.Pending => "等待中",
        TaskStatus.Processing => "处理中",
        TaskStatus.Completed => "已完成",
        TaskStatus.Failed => "失败",
        TaskStatus.Cancelled => "已取消",
        _ => "未知"
    };

    /// <summary>
    /// 进度文本
    /// </summary>
    public string ProgressText => $"{Progress:F1}%";

    /// <summary>
    /// 是否正在处理
    /// </summary>
    public bool IsProcessing => Status == TaskStatus.Processing;

    /// <summary>
    /// 是否已完成
    /// </summary>
    public bool IsCompleted => Status is TaskStatus.Completed or TaskStatus.Failed or TaskStatus.Cancelled;

    /// <summary>
    /// 添加日志
    /// </summary>
    public void AddLog(LogLevel level, string message, string? details = null)
    {
        var log = new ProcessingLog
        {
            Timestamp = DateTime.UtcNow,
            Level = level,
            Message = message,
            Details = details
        };
        
        var app = System.Windows.Application.Current;
        if (app?.Dispatcher != null)
        {
            app.Dispatcher.BeginInvoke(new Action(() => Logs.Add(log)));
        }
        else
        {
            Logs.Add(log);
        }
    }

    /// <summary>
    /// 取消任务
    /// </summary>
    public void Cancel()
    {
        if (!CancellationTokenSource.Token.IsCancellationRequested)
        {
            CancellationTokenSource.Cancel();
            Status = TaskStatus.Cancelled;
            AddLog(LogLevel.Information, "任务已取消");
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            CancellationTokenSource?.Dispose();
            _disposed = true;
        }
    }
}

/// <summary>
/// 任务类型枚举
/// </summary>
public enum TaskType
{
    /// <summary>
    /// 合并文件
    /// </summary>
    Merge,
    
    /// <summary>
    /// 拆分文件
    /// </summary>
    Split
}

/// <summary>
/// 任务状态枚举
/// </summary>
public enum TaskStatus
{
    /// <summary>
    /// 等待中
    /// </summary>
    Pending,
    
    /// <summary>
    /// 处理中
    /// </summary>
    Processing,
    
    /// <summary>
    /// 已完成
    /// </summary>
    Completed,
    
    /// <summary>
    /// 失败
    /// </summary>
    Failed,
    
    /// <summary>
    /// 已取消
    /// </summary>
    Cancelled
}

/// <summary>
/// 处理日志模型
/// </summary>
public class ProcessingLog
{
    /// <summary>
    /// 时间戳
    /// </summary>
    public DateTime Timestamp { get; set; }

    /// <summary>
    /// 日志级别
    /// </summary>
    public LogLevel Level { get; set; }

    /// <summary>
    /// 消息
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// 详细信息
    /// </summary>
    public string? Details { get; set; }

    /// <summary>
    /// 格式化的时间文本
    /// </summary>
    public string TimeText => Timestamp.ToString("HH:mm:ss");

    /// <summary>
    /// 级别文本
    /// </summary>
    public string LevelText => Level switch
    {
        LogLevel.Trace => "跟踪",
        LogLevel.Debug => "调试",
        LogLevel.Information => "信息",
        LogLevel.Warning => "警告",
        LogLevel.Error => "错误",
        LogLevel.Critical => "严重",
        _ => "未知"
    };
}

/// <summary>
/// 日志级别枚举
/// </summary>
public enum LogLevel
{
    Trace,
    Debug,
    Information,
    Warning,
    Error,
    Critical
}