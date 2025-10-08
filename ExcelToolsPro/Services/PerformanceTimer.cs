using Microsoft.Extensions.Logging;
using System.Diagnostics;

namespace ExcelToolsPro.Services;

/// <summary>
/// 性能计时器，用于统一的性能监控和结构化日志记录
/// </summary>
public class PerformanceTimer : IDisposable
{
    private readonly ILogger _logger;
    private readonly string _operationName;
    private readonly string _operationId;
    private readonly Stopwatch _stopwatch;
    private readonly Dictionary<string, object> _context;
    private bool _disposed = false;

    public PerformanceTimer(ILogger logger, string operationName, string? operationId = null)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _operationName = operationName ?? throw new ArgumentNullException(nameof(operationName));
        _operationId = operationId ?? Guid.NewGuid().ToString("N")[..8];
        _stopwatch = Stopwatch.StartNew();
        _context = new Dictionary<string, object>();

        // 记录操作开始
        _logger.LogInformation("[{OperationId}] {OperationName} 开始", _operationId, _operationName);
    }

    /// <summary>
    /// 添加上下文信息
    /// </summary>
    public PerformanceTimer WithContext(string key, object value)
    {
        _context[key] = value;
        return this;
    }

    /// <summary>
    /// 添加多个上下文信息
    /// </summary>
    public PerformanceTimer WithContext(Dictionary<string, object> context)
    {
        foreach (var kvp in context)
        {
            _context[kvp.Key] = kvp.Value;
        }
        return this;
    }

    /// <summary>
    /// 记录中间检查点
    /// </summary>
    public void Checkpoint(string checkpointName, object? additionalData = null)
    {
        var elapsedMs = _stopwatch.ElapsedMilliseconds;
        
        var logData = new Dictionary<string, object>
        {
            ["OperationId"] = _operationId,
            ["OperationName"] = _operationName,
            ["CheckpointName"] = checkpointName,
            ["ElapsedMs"] = elapsedMs
        };

        // 添加上下文信息
        foreach (var kvp in _context)
        {
            logData[kvp.Key] = kvp.Value;
        }

        // 添加额外数据
        if (additionalData != null)
        {
            if (additionalData is Dictionary<string, object> dict)
            {
                foreach (var kvp in dict)
                {
                    logData[kvp.Key] = kvp.Value;
                }
            }
            else
            {
                logData["AdditionalData"] = additionalData;
            }
        }

        _logger.LogDebug("[{OperationId}] {OperationName} 检查点: {CheckpointName}, 耗时: {ElapsedMs}ms, 数据: {@LogData}",
            _operationId, _operationName, checkpointName, elapsedMs, logData);
    }

    /// <summary>
    /// 记录错误
    /// </summary>
    public void LogError(Exception exception, string? message = null)
    {
        var elapsedMs = _stopwatch.ElapsedMilliseconds;
        
        var logData = new Dictionary<string, object>
        {
            ["OperationId"] = _operationId,
            ["OperationName"] = _operationName,
            ["ElapsedMs"] = elapsedMs,
            ["ExceptionType"] = exception.GetType().Name,
            ["ExceptionMessage"] = exception.Message
        };

        // 添加上下文信息
        foreach (var kvp in _context)
        {
            logData[kvp.Key] = kvp.Value;
        }

        var logMessage = message ?? $"{_operationName} 执行失败";
        _logger.LogError(exception, "[{OperationId}] {LogMessage}, 耗时: {ElapsedMs}ms, 数据: {@LogData}",
            _operationId, logMessage, elapsedMs, logData);
    }

    /// <summary>
    /// 记录错误（字符串消息）
    /// </summary>
    public void LogError(string message, object? additionalData = null)
    {
        var elapsedMs = _stopwatch.ElapsedMilliseconds;
        
        var logData = new Dictionary<string, object>
        {
            ["OperationId"] = _operationId,
            ["OperationName"] = _operationName,
            ["ElapsedMs"] = elapsedMs
        };

        // 添加上下文信息
        foreach (var kvp in _context)
        {
            logData[kvp.Key] = kvp.Value;
        }

        // 添加额外数据
        if (additionalData != null)
        {
            logData["AdditionalData"] = additionalData;
        }

        _logger.LogError("[{OperationId}] {OperationName} 错误: {Message}, 耗时: {ElapsedMs}ms, 数据: {@LogData}",
            _operationId, _operationName, message, elapsedMs, logData);
    }

    /// <summary>
    /// 记录警告
    /// </summary>
    public void LogWarning(string message, object? additionalData = null)
    {
        var elapsedMs = _stopwatch.ElapsedMilliseconds;
        
        var logData = new Dictionary<string, object>
        {
            ["OperationId"] = _operationId,
            ["OperationName"] = _operationName,
            ["ElapsedMs"] = elapsedMs
        };

        // 添加上下文信息
        foreach (var kvp in _context)
        {
            logData[kvp.Key] = kvp.Value;
        }

        // 添加额外数据
        if (additionalData != null)
        {
            logData["AdditionalData"] = additionalData;
        }

        _logger.LogWarning("[{OperationId}] {OperationName} 警告: {Message}, 耗时: {ElapsedMs}ms, 数据: {@LogData}",
            _operationId, _operationName, message, elapsedMs, logData);
    }

    /// <summary>
    /// 获取当前耗时
    /// </summary>
    public long ElapsedMilliseconds => _stopwatch.ElapsedMilliseconds;

    /// <summary>
    /// 获取操作ID
    /// </summary>
    public string OperationId => _operationId;

    /// <summary>
    /// 添加上下文信息
    /// </summary>
    public void AddContext(string key, object value)
    {
        _context[key] = value;
    }

    public void Dispose()
    {
        if (_disposed) return;

        _stopwatch.Stop();
        var elapsedMs = _stopwatch.ElapsedMilliseconds;
        
        var logData = new Dictionary<string, object>
        {
            ["OperationId"] = _operationId,
            ["OperationName"] = _operationName,
            ["ElapsedMs"] = elapsedMs
        };

        // 添加上下文信息
        foreach (var kvp in _context)
        {
            logData[kvp.Key] = kvp.Value;
        }

        // 根据耗时选择日志级别
        if (elapsedMs > 10000) // 超过10秒
        {
            _logger.LogWarning("[{OperationId}] {OperationName} 完成（耗时较长）, 耗时: {ElapsedMs}ms, 数据: {@LogData}",
                _operationId, _operationName, elapsedMs, logData);
        }
        else if (elapsedMs > 1000) // 超过1秒
        {
            _logger.LogInformation("[{OperationId}] {OperationName} 完成, 耗时: {ElapsedMs}ms, 数据: {@LogData}",
                _operationId, _operationName, elapsedMs, logData);
        }
        else
        {
            _logger.LogDebug("[{OperationId}] {OperationName} 完成, 耗时: {ElapsedMs}ms, 数据: {@LogData}",
                _operationId, _operationName, elapsedMs, logData);
        }

        _disposed = true;
    }
}

/// <summary>
/// 性能计时器扩展方法
/// </summary>
public static class PerformanceTimerExtensions
{
    /// <summary>
    /// 创建性能计时器
    /// </summary>
    public static PerformanceTimer StartTimer(this ILogger logger, string operationName, string? operationId = null)
    {
        return new PerformanceTimer(logger, operationName, operationId);
    }

    /// <summary>
    /// 创建性能计时器（带上下文）
    /// </summary>
    public static PerformanceTimer CreateTimer(ILogger logger, string operationName, object? context = null, string? operationId = null)
    {
        var timer = new PerformanceTimer(logger, operationName, operationId);
        if (context != null)
        {
            timer.AddContext("Context", context);
        }
        return timer;
    }

    /// <summary>
    /// 使用性能计时器执行操作
    /// </summary>
    public static T WithTimer<T>(this ILogger logger, string operationName, Func<PerformanceTimer, T> operation, string? operationId = null)
    {
        using var timer = new PerformanceTimer(logger, operationName, operationId);
        try
        {
            return operation(timer);
        }
        catch (Exception ex)
        {
            timer.LogError(ex);
            throw;
        }
    }

    /// <summary>
    /// 使用性能计时器执行异步操作
    /// </summary>
    public static async Task<T> WithTimerAsync<T>(this ILogger logger, string operationName, Func<PerformanceTimer, Task<T>> operation, string? operationId = null)
    {
        using var timer = new PerformanceTimer(logger, operationName, operationId);
        try
        {
            return await operation(timer).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            timer.LogError(ex);
            throw;
        }
    }

    /// <summary>
    /// 使用性能计时器执行异步操作（无返回值）
    /// </summary>
    public static async Task WithTimerAsync(this ILogger logger, string operationName, Func<PerformanceTimer, Task> operation, string? operationId = null)
    {
        using var timer = new PerformanceTimer(logger, operationName, operationId);
        try
        {
            await operation(timer).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            timer.LogError(ex);
            throw;
        }
    }
}