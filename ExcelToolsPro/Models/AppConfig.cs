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
    private int _maxDegreeOfParallelism = Environment.ProcessorCount;
    private int _progressThrottleMs = 100;
    private bool _largeFileMode = false;
    private int _largeFileSizeThresholdMB = 50;
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
    private int _ioBufferSizeKB = 64;
    private string _csvEncoding = "UTF-8";
    private bool _csvIncludeBom = true;
    private bool _useAsyncIO = true;
    private int _previewPageSize = 500;
    private int _maxPreviewItems = 2000;
    private int _htmlParseTimeoutMs = 500;
    private bool _enableHtmlParseOptimization = true;
    private int _htmlContentMaxSizeKB = 1024;

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
    /// 最大并发度（1-16）
    /// </summary>
    public int MaxDegreeOfParallelism
    {
        get => _maxDegreeOfParallelism;
        set => SetProperty(ref _maxDegreeOfParallelism, Math.Max(1, Math.Min(16, value)));
    }

    /// <summary>
    /// 进度节流间隔（毫秒）
    /// </summary>
    public int ProgressThrottleMs
    {
        get => _progressThrottleMs;
        set => SetProperty(ref _progressThrottleMs, Math.Max(50, Math.Min(1000, value)));
    }

    /// <summary>
    /// 是否启用低内存模式
    /// </summary>
    public bool LargeFileMode
    {
        get => _largeFileMode;
        set => SetProperty(ref _largeFileMode, value);
    }

    /// <summary>
    /// 大文件阈值（MB），超过此值自动启用低内存模式
    /// </summary>
    public int LargeFileSizeThresholdMB
    {
        get => _largeFileSizeThresholdMB;
        set => SetProperty(ref _largeFileSizeThresholdMB, Math.Max(10, Math.Min(1000, value)));
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

    /// <summary>
    /// I/O缓冲区大小（KB）
    /// </summary>
    public int IOBufferSizeKB
    {
        get => _ioBufferSizeKB;
        set => SetProperty(ref _ioBufferSizeKB, Math.Max(4, Math.Min(1024, value)));
    }

    /// <summary>
    /// CSV文件编码
    /// </summary>
    public string CsvEncoding
    {
        get => _csvEncoding;
        set => SetProperty(ref _csvEncoding, value);
    }

    /// <summary>
    /// CSV文件是否包含BOM
    /// </summary>
    public bool CsvIncludeBom
    {
        get => _csvIncludeBom;
        set => SetProperty(ref _csvIncludeBom, value);
    }

    /// <summary>
    /// 是否使用异步I/O
    /// </summary>
    public bool UseAsyncIO
    {
        get => _useAsyncIO;
        set => SetProperty(ref _useAsyncIO, value);
    }

    /// <summary>
    /// 预览分页大小
    /// </summary>
    public int PreviewPageSize
    {
        get => _previewPageSize;
        set => SetProperty(ref _previewPageSize, Math.Max(50, Math.Min(1000, value)));
    }

    /// <summary>
    /// 最大预览项目数
    /// </summary>
    public int MaxPreviewItems
    {
        get => _maxPreviewItems;
        set => SetProperty(ref _maxPreviewItems, Math.Max(100, Math.Min(10000, value)));
    }

    /// <summary>
    /// HTML解析超时时间（毫秒）
    /// </summary>
    public int HtmlParseTimeoutMs
    {
        get => _htmlParseTimeoutMs;
        set => SetProperty(ref _htmlParseTimeoutMs, Math.Max(100, Math.Min(5000, value)));
    }

    /// <summary>
    /// 是否启用HTML解析优化
    /// </summary>
    public bool EnableHtmlParseOptimization
    {
        get => _enableHtmlParseOptimization;
        set => SetProperty(ref _enableHtmlParseOptimization, value);
    }

    /// <summary>
    /// HTML内容最大处理大小（KB）
    /// </summary>
    public int HtmlContentMaxSizeKB
    {
        get => _htmlContentMaxSizeKB;
        set => SetProperty(ref _htmlContentMaxSizeKB, Math.Max(64, Math.Min(10240, value)));
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