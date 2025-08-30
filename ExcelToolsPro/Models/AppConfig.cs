using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelToolsPro.Models;

/// <summary>
/// 应用程序配置模型
/// </summary>
public class AppConfig : INotifyPropertyChanged
{
    private bool _addHeaders = true;
    private bool _dedupeHeaders = true;
    private string _outputFormat = "xlsx";
    private int _maxMemoryMB = 512;
    private int _concurrentTasks = 2;
    private string? _lastOutputDirectory;
    private bool _enableDarkMode = false;
    private string _language = "zh-CN";
    private bool _minimizeToTray = true;
    private bool _enableAutoThrottle = true;
    private int _chunkSizeMB = 1;
    private TimeSpan _processingTimeout = TimeSpan.FromMinutes(30);
    private int _customSplitRowCount = 1000;
    private string _mergeNamingTemplate = "{timestamp}_merged";
    private string _splitNamingTemplate = "{filename}_{index}";
    private bool _useCustomNaming = false;

    /// <summary>
    /// 是否自动添加表头
    /// </summary>
    public bool AddHeaders
    {
        get => _addHeaders;
        set => SetProperty(ref _addHeaders, value);
    }

    /// <summary>
    /// 是否去重表头
    /// </summary>
    public bool DedupeHeaders
    {
        get => _dedupeHeaders;
        set => SetProperty(ref _dedupeHeaders, value);
    }

    /// <summary>
    /// 输出格式
    /// </summary>
    public string OutputFormat
    {
        get => _outputFormat;
        set => SetProperty(ref _outputFormat, value);
    }

    /// <summary>
    /// 最大内存使用量(MB)
    /// </summary>
    public int MaxMemoryMB
    {
        get => _maxMemoryMB;
        set => SetProperty(ref _maxMemoryMB, value);
    }

    /// <summary>
    /// 并发任务数
    /// </summary>
    public int ConcurrentTasks
    {
        get => _concurrentTasks;
        set => SetProperty(ref _concurrentTasks, value);
    }

    /// <summary>
    /// 上次输出目录
    /// </summary>
    public string? LastOutputDirectory
    {
        get => _lastOutputDirectory;
        set => SetProperty(ref _lastOutputDirectory, value);
    }

    /// <summary>
    /// 是否启用深色模式
    /// </summary>
    public bool EnableDarkMode
    {
        get => _enableDarkMode;
        set => SetProperty(ref _enableDarkMode, value);
    }

    /// <summary>
    /// 语言设置
    /// </summary>
    public string Language
    {
        get => _language;
        set => SetProperty(ref _language, value);
    }

    /// <summary>
    /// 是否最小化到托盘
    /// </summary>
    public bool MinimizeToTray
    {
        get => _minimizeToTray;
        set => SetProperty(ref _minimizeToTray, value);
    }

    /// <summary>
    /// 是否启用自动限流
    /// </summary>
    public bool EnableAutoThrottle
    {
        get => _enableAutoThrottle;
        set => SetProperty(ref _enableAutoThrottle, value);
    }

    /// <summary>
    /// 分块大小(MB)
    /// </summary>
    public int ChunkSizeMB
    {
        get => _chunkSizeMB;
        set => SetProperty(ref _chunkSizeMB, value);
    }

    /// <summary>
    /// 处理超时时间
    /// </summary>
    public TimeSpan ProcessingTimeout
    {
        get => _processingTimeout;
        set => SetProperty(ref _processingTimeout, value);
    }

    /// <summary>
    /// 自定义分割行数
    /// </summary>
    public int CustomSplitRowCount
    {
        get => _customSplitRowCount;
        set => SetProperty(ref _customSplitRowCount, value);
    }

    /// <summary>
    /// 合并文件命名模板
    /// </summary>
    public string MergeNamingTemplate
    {
        get => _mergeNamingTemplate;
        set => SetProperty(ref _mergeNamingTemplate, value);
    }

    /// <summary>
    /// 拆分文件命名模板
    /// </summary>
    public string SplitNamingTemplate
    {
        get => _splitNamingTemplate;
        set => SetProperty(ref _splitNamingTemplate, value);
    }

    /// <summary>
    /// 是否使用自定义命名
    /// </summary>
    public bool UseCustomNaming
    {
        get => _useCustomNaming;
        set => SetProperty(ref _useCustomNaming, value);
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
}