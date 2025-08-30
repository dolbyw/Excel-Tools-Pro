namespace ExcelToolsPro.Services;

/// <summary>
/// 性能监控服务接口
/// </summary>
public interface IPerformanceMonitorService
{
    /// <summary>
    /// 获取系统性能指标
    /// </summary>
    Task<SystemMetrics> GetSystemMetricsAsync();

    /// <summary>
    /// 更新性能设置
    /// </summary>
    Task<bool> UpdatePerformanceSettingsAsync(PerformanceSettings settings);

    /// <summary>
    /// 性能指标更新事件
    /// </summary>
    event EventHandler<SystemMetrics>? MetricsUpdated;

    /// <summary>
    /// 开始监控
    /// </summary>
    void StartMonitoring();

    /// <summary>
    /// 停止监控
    /// </summary>
    void StopMonitoring();
}

/// <summary>
/// 系统性能指标
/// </summary>
public class SystemMetrics
{
    /// <summary>
    /// 内存使用百分比
    /// </summary>
    public float MemoryUsagePercent { get; set; }

    /// <summary>
    /// CPU使用百分比
    /// </summary>
    public float CpuUsagePercent { get; set; }

    /// <summary>
    /// 可用磁盘空间(GB)
    /// </summary>
    public double AvailableDiskSpaceGB { get; set; }

    /// <summary>
    /// 处理速度(文件/分钟)
    /// </summary>
    public float ProcessingSpeedFilesPerMin { get; set; }

    /// <summary>
    /// 时间戳
    /// </summary>
    public DateTime Timestamp { get; set; }
}

/// <summary>
/// 性能设置
/// </summary>
public class PerformanceSettings
{
    /// <summary>
    /// 最大内存使用量(MB)
    /// </summary>
    public int MaxMemoryMB { get; set; } = 512;

    /// <summary>
    /// 最大并发文件数
    /// </summary>
    public int MaxConcurrentFiles { get; set; } = 2;

    /// <summary>
    /// 是否启用自动限流
    /// </summary>
    public bool EnableAutoThrottle { get; set; } = true;
}