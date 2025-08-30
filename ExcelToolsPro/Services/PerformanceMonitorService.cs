using Microsoft.Extensions.Logging;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Timers;

namespace ExcelToolsPro.Services;

/// <summary>
/// 性能监控服务实现
/// </summary>
public class PerformanceMonitorService : IPerformanceMonitorService, IDisposable
{
    private readonly ILogger<PerformanceMonitorService> _logger;
    private readonly System.Timers.Timer _monitoringTimer;
    private readonly System.Timers.Timer _healthCheckTimer;
    private PerformanceCounter? _cpuCounter;
    private PerformanceCounter? _memoryCounter;
    private PerformanceSettings _settings;
    private bool _disposed = false;
    private volatile bool _isUpdating = false; // 防止定时器回调重入
    private float _totalMemoryMB = 0f; // 缓存总物理内存，避免频繁WMI查询
    private bool _monitoringActive = false; // 跟踪监控是否处于激活状态
    private int _consecutiveErrors = 0; // 连续错误计数
    private DateTime _lastSuccessfulUpdate = DateTime.Now;
    private readonly object _lockObject = new object();
    private bool _performanceCountersInitialized = false;
    private int _initializationRetryCount = 0;
    private const int MAX_INITIALIZATION_RETRIES = 3;
    private const int MAX_CONSECUTIVE_ERRORS = 5;
    private const int WMI_TIMEOUT_MS = 5000;

    public event EventHandler<SystemMetrics>? MetricsUpdated;

    public PerformanceMonitorService(ILogger<PerformanceMonitorService> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _settings = new PerformanceSettings();
        
        try
        {
            // 异步初始化性能计数器，避免阻塞构造函数
            _ = Task.Run(async () => await InitializePerformanceCountersWithRetryAsync().ConfigureAwait(false));
            
            _monitoringTimer = new System.Timers.Timer(10000); // 每10秒更新一次，减少频率
            _monitoringTimer.AutoReset = false; // 避免重入，手动重启
            _monitoringTimer.Elapsed += OnTimerElapsed;
            
            // 健康检查定时器，每分钟检查一次
            _healthCheckTimer = new System.Timers.Timer(60000);
            _healthCheckTimer.AutoReset = true;
            _healthCheckTimer.Elapsed += OnHealthCheckElapsed;
            _healthCheckTimer.Start();
            
            _logger.LogInformation("性能监控服务初始化完成");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "性能监控服务初始化失败");
            throw;
        }
    }
    
    private async Task InitializePerformanceCountersWithRetryAsync()
    {
        while (_initializationRetryCount < MAX_INITIALIZATION_RETRIES && !_performanceCountersInitialized && !_disposed)
        {
            try
            {
                _initializationRetryCount++;
                _logger.LogDebug("开始性能计数器初始化，尝试次数: {RetryCount}/{MaxRetries}", _initializationRetryCount, MAX_INITIALIZATION_RETRIES);
                
                await Task.Delay(1000 * _initializationRetryCount).ConfigureAwait(false); // 递增延迟
                
                await InitializePerformanceCountersAsync().ConfigureAwait(false);
                
                _performanceCountersInitialized = true;
                _logger.LogInformation("性能计数器初始化成功，尝试次数: {RetryCount}", _initializationRetryCount);
                break;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "性能计数器初始化失败，尝试次数: {RetryCount}/{MaxRetries}", _initializationRetryCount, MAX_INITIALIZATION_RETRIES);
                
                if (_initializationRetryCount >= MAX_INITIALIZATION_RETRIES)
                {
                    _logger.LogError("性能计数器初始化最终失败，将使用降级模式");
                }
            }
        }
    }
    
    private async Task InitializePerformanceCountersAsync()
    {
        try
        {
            // 清理旧的计数器
            _cpuCounter?.Dispose();
            _memoryCounter?.Dispose();
            
            _cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
            _memoryCounter = new PerformanceCounter("Memory", "Available MBytes");
            
            // 初始化CPU计数器（第一次调用通常返回0）
            _ = _cpuCounter.NextValue();
            
            // 预取并缓存总物理内存，使用超时机制
            await GetTotalPhysicalMemoryWithTimeoutAsync().ConfigureAwait(false);
            
            _logger.LogDebug("性能计数器初始化完成");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "初始化性能计数器时发生错误");
            throw;
        }
    }
    
    private async Task GetTotalPhysicalMemoryWithTimeoutAsync()
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(WMI_TIMEOUT_MS));
            
            await Task.Run(() =>
            {
                try
                {
                    using var searcher = new ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem");
                    searcher.Options.Timeout = TimeSpan.FromMilliseconds(WMI_TIMEOUT_MS);
                    
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        cts.Token.ThrowIfCancellationRequested();
                        var totalMemoryBytes = Convert.ToUInt64(obj["TotalPhysicalMemory"]);
                        _totalMemoryMB = (float)(totalMemoryBytes / (1024.0 * 1024.0));
                        _logger.LogDebug("获取到总物理内存: {TotalMemoryMB}MB", _totalMemoryMB);
                        break;
                    }
                }
                catch (OperationCanceledException)
                {
                    _logger.LogWarning("WMI查询超时，使用默认内存值");
                    throw;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "WMI查询失败，使用默认内存值");
                    throw;
                }
            }, cts.Token).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "获取总物理内存失败，使用默认值8GB");
            _totalMemoryMB = 8192f; // 默认8GB
        }
    }

    public async Task<SystemMetrics> GetSystemMetricsAsync()
    {
        return await Task.Run(() =>
        {
            var metrics = new SystemMetrics
            {
                Timestamp = DateTime.Now
            };

            try
            {
                // CPU使用率
                if (_cpuCounter != null)
                {
                    metrics.CpuUsagePercent = _cpuCounter.NextValue();
                }

                // 内存使用率
                if (_memoryCounter != null)
                {
                    var availableMemoryMB = _memoryCounter.NextValue();
                    var totalMemoryMB = GetTotalPhysicalMemoryMB();
                    if (totalMemoryMB > 0)
                    {
                        metrics.MemoryUsagePercent = ((totalMemoryMB - availableMemoryMB) / totalMemoryMB) * 100;
                    }
                }
                else
                {
                    // 备用方法：使用GC获取当前进程内存使用情况
                    var processMemoryMB = GC.GetTotalMemory(false) / (1024.0 * 1024.0);
                    metrics.MemoryUsagePercent = (float)(processMemoryMB / _settings.MaxMemoryMB * 100);
                }

                // 可用磁盘空间
                metrics.AvailableDiskSpaceGB = GetAvailableDiskSpace();

                // 处理速度（这里是占位符，实际应用中需要根据具体处理情况计算）
                metrics.ProcessingSpeedFilesPerMin = 0.0f;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "获取系统性能指标时发生错误");
            }

            return metrics;
        }).ConfigureAwait(false);
    }

    public async Task<bool> UpdatePerformanceSettingsAsync(PerformanceSettings settings)
    {
        try
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _logger.LogInformation("性能设置已更新");
            return await Task.FromResult(true).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "更新性能设置时发生错误");
            return false;
        }
    }

    public void StartMonitoring()
    {
        try
        {
            if (!_monitoringTimer.Enabled)
            {
                _monitoringActive = true;
                _monitoringTimer.Start();
                _logger.LogInformation("性能监控已启动");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "启动性能监控时发生错误");
        }
    }

    public void StopMonitoring()
    {
        try
        {
            if (_monitoringTimer.Enabled)
            {
                _monitoringActive = false;
                _monitoringTimer.Stop();
                _logger.LogInformation("性能监控已停止");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "停止性能监控时发生错误");
        }
    }

    private async void OnTimerElapsed(object? sender, ElapsedEventArgs e)
    {
        if (_disposed) return;
        if (_isUpdating) return;
        
        lock (_lockObject)
        {
            if (_isUpdating) return;
            _isUpdating = true;
        }
        
        try
        {
            // 在后台线程获取性能指标，避免阻塞UI线程
            var metrics = await GetSystemMetricsAsync().ConfigureAwait(false);
            
            // 直接触发事件，由订阅方自行决定是否切换到UI线程
            try
            {
                MetricsUpdated?.Invoke(this, metrics);
                _consecutiveErrors = 0; // 重置错误计数
                _lastSuccessfulUpdate = DateTime.Now;
            }
            catch (Exception eventEx)
            {
                _logger.LogError(eventEx, "触发MetricsUpdated事件时发生错误");
                _consecutiveErrors++;
            }
        }
        catch (Exception ex)
        {
            _consecutiveErrors++;
            _logger.LogError(ex, "更新性能指标时发生错误，连续错误次数: {ConsecutiveErrors}", _consecutiveErrors);
            
            // 如果连续错误过多，尝试重新初始化性能计数器
            if (_consecutiveErrors >= MAX_CONSECUTIVE_ERRORS)
            {
                _logger.LogWarning("连续错误次数过多，尝试重新初始化性能计数器");
                _ = Task.Run(async () => await ReinitializePerformanceCountersAsync().ConfigureAwait(false));
            }
        }
        finally
        {
            lock (_lockObject)
            {
                _isUpdating = false;
            }
            
            // 使用指数退避策略重启定时器
            if (_monitoringActive && !_disposed)
            {
                var delay = Math.Min(10000 * Math.Pow(2, Math.Min(_consecutiveErrors, 3)), 60000); // 最大1分钟
                _monitoringTimer.Interval = delay;
                _monitoringTimer.Start();
            }
        }
    }
    
    private async void OnHealthCheckElapsed(object? sender, ElapsedEventArgs e)
    {
        if (_disposed) return;
        
        try
        {
            await PerformHealthCheckAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "健康检查时发生错误");
        }
    }
    
    private async Task PerformHealthCheckAsync()
    {
        try
        {
            var timeSinceLastUpdate = DateTime.Now - _lastSuccessfulUpdate;
            
            // 检查是否长时间没有成功更新
            if (timeSinceLastUpdate.TotalMinutes > 5)
            {
                _logger.LogWarning("性能监控长时间未更新，上次成功更新时间: {LastUpdate}", _lastSuccessfulUpdate);
                
                // 尝试重新初始化
                await ReinitializePerformanceCountersAsync().ConfigureAwait(false);
            }
            
            // 检查性能计数器状态
            if (!_performanceCountersInitialized && _initializationRetryCount < MAX_INITIALIZATION_RETRIES)
            {
                _logger.LogInformation("检测到性能计数器未初始化，尝试重新初始化");
                await InitializePerformanceCountersWithRetryAsync().ConfigureAwait(false);
            }
            
            _logger.LogDebug("健康检查完成，连续错误次数: {ConsecutiveErrors}, 上次更新: {LastUpdate}", 
                _consecutiveErrors, _lastSuccessfulUpdate);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "执行健康检查时发生错误");
        }
    }
    
    private async Task ReinitializePerformanceCountersAsync()
    {
        try
        {
            _logger.LogInformation("开始重新初始化性能计数器");
            
            _performanceCountersInitialized = false;
            _initializationRetryCount = 0;
            _consecutiveErrors = 0;
            
            await InitializePerformanceCountersWithRetryAsync().ConfigureAwait(false);
            
            _logger.LogInformation("性能计数器重新初始化完成");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "重新初始化性能计数器失败");
        }
    }

    private float GetTotalPhysicalMemoryMB()
    {
        try
        {
            if (_totalMemoryMB > 0f)
            {
                return _totalMemoryMB;
            }

            // 未命中缓存时，后台异步刷新，当前返回一个安全默认值，避免阻塞
            _ = Task.Run(() =>
            {
                try
                {
                    using var searcher = new ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem");
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var totalMemoryBytes = Convert.ToUInt64(obj["TotalPhysicalMemory"]);
                        _totalMemoryMB = (float)(totalMemoryBytes / (1024.0 * 1024.0));
                        break;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "后台刷新总物理内存失败");
                }
            });

            return 8192f; // 默认8GB
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "获取总物理内存时发生错误，使用默认值");
            return 8192f; // 默认8GB
        }
    }

    private double GetAvailableDiskSpace()
    {
        try
        {
            var drive = new DriveInfo(Path.GetPathRoot(Environment.CurrentDirectory) ?? "C:\\");
            return drive.AvailableFreeSpace / (1024.0 * 1024.0 * 1024.0); // 转换为GB
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "获取可用磁盘空间时发生错误");
            return 0;
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true; // 先设置标志，防止其他操作
            
            try
            {
                _logger.LogDebug("开始释放性能监控服务资源");
                
                // 停止监控
                StopMonitoring();
                
                // 释放定时器
                try
                {
                    _monitoringTimer?.Stop();
                    _monitoringTimer?.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "释放监控定时器时发生错误");
                }
                
                try
                {
                    _healthCheckTimer?.Stop();
                    _healthCheckTimer?.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "释放健康检查定时器时发生错误");
                }
                
                // 释放性能计数器
                try
                {
                    _cpuCounter?.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "释放CPU性能计数器时发生错误");
                }
                
                try
                {
                    _memoryCounter?.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "释放内存性能计数器时发生错误");
                }
                
                _logger.LogInformation("性能监控服务资源释放完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "释放性能监控服务资源时发生严重错误");
            }
        }
    }
}