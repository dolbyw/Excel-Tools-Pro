using ExcelToolsPro.Models;
using Microsoft.Extensions.Logging;

namespace ExcelToolsPro.Services.Reporting;

/// <summary>
/// 处理过程报告器接口
/// </summary>
public interface IProcessingReporter
{
    /// <summary>
    /// 开始操作报告
    /// </summary>
    /// <param name="operationType">操作类型</param>
    /// <param name="operationId">操作ID</param>
    /// <param name="context">操作上下文</param>
    /// <returns>操作会话</returns>
    IOperationSession StartOperation(
        string operationType,
        string operationId,
        OperationContext context);
    
    /// <summary>
    /// 报告进度
    /// </summary>
    /// <param name="operationId">操作ID</param>
    /// <param name="progress">进度百分比（0-100）</param>
    /// <param name="message">进度消息</param>
    /// <param name="details">详细信息</param>
    void ReportProgress(
        string operationId,
        float progress,
        string? message = null,
        object? details = null);
    
    /// <summary>
    /// 报告检查点
    /// </summary>
    /// <param name="operationId">操作ID</param>
    /// <param name="checkpointName">检查点名称</param>
    /// <param name="details">详细信息</param>
    void ReportCheckpoint(
        string operationId,
        string checkpointName,
        object? details = null);
    
    /// <summary>
    /// 报告警告
    /// </summary>
    /// <param name="operationId">操作ID</param>
    /// <param name="message">警告消息</param>
    /// <param name="details">详细信息</param>
    void ReportWarning(
        string operationId,
        string message,
        object? details = null);
    
    /// <summary>
    /// 报告错误
    /// </summary>
    /// <param name="operationId">操作ID</param>
    /// <param name="error">错误信息</param>
    /// <param name="exception">异常对象</param>
    void ReportError(
        string operationId,
        string error,
        Exception? exception = null);
    
    /// <summary>
    /// 完成操作报告
    /// </summary>
    /// <param name="operationId">操作ID</param>
    /// <param name="result">操作结果</param>
    void CompleteOperation(
        string operationId,
        OperationResult result);
    
    /// <summary>
    /// 获取操作统计信息
    /// </summary>
    /// <param name="operationId">操作ID</param>
    /// <returns>统计信息</returns>
    OperationStatistics? GetOperationStatistics(string operationId);
    
    /// <summary>
    /// 创建进度节流器
    /// </summary>
    /// <param name="progress">进度报告器</param>
    /// <param name="throttleMs">节流间隔（毫秒）</param>
    /// <param name="minProgressDelta">最小进度变化</param>
    /// <returns>进度节流器</returns>
    IProgressThrottler CreateProgressThrottler(
        IProgress<float>? progress,
        int throttleMs = 100,
        float minProgressDelta = 1f);
}

/// <summary>
/// 操作会话接口
/// </summary>
public interface IOperationSession : IDisposable
{
    /// <summary>
    /// 操作ID
    /// </summary>
    string OperationId { get; }
    
    /// <summary>
    /// 操作类型
    /// </summary>
    string OperationType { get; }
    
    /// <summary>
    /// 开始时间
    /// </summary>
    DateTime StartTime { get; }
    
    /// <summary>
    /// 报告检查点
    /// </summary>
    /// <param name="checkpointName">检查点名称</param>
    /// <param name="details">详细信息</param>
    void Checkpoint(string checkpointName, object? details = null);
    
    /// <summary>
    /// 报告进度
    /// </summary>
    /// <param name="progress">进度百分比</param>
    /// <param name="message">进度消息</param>
    /// <param name="details">详细信息</param>
    void Progress(float progress, string? message = null, object? details = null);
    
    /// <summary>
    /// 报告警告
    /// </summary>
    /// <param name="message">警告消息</param>
    /// <param name="details">详细信息</param>
    void Warning(string message, object? details = null);
    
    /// <summary>
    /// 报告错误
    /// </summary>
    /// <param name="error">错误信息</param>
    /// <param name="exception">异常对象</param>
    void Error(string error, Exception? exception = null);
    
    /// <summary>
    /// 完成操作
    /// </summary>
    /// <param name="result">操作结果</param>
    void Complete(OperationResult result);
}

/// <summary>
/// 进度节流器接口
/// </summary>
public interface IProgressThrottler
{
    /// <summary>
    /// 报告进度
    /// </summary>
    /// <param name="progress">进度百分比</param>
    /// <param name="forceReport">是否强制报告</param>
    void Report(float progress, bool forceReport = false);
    
    /// <summary>
    /// 最后报告的进度
    /// </summary>
    float LastReportedProgress { get; }
    
    /// <summary>
    /// 最后报告时间
    /// </summary>
    DateTime LastReportTime { get; }
}

/// <summary>
/// 操作上下文
/// </summary>
public class OperationContext
{
    /// <summary>
    /// 文件路径列表
    /// </summary>
    public List<string> FilePaths { get; set; } = new();
    
    /// <summary>
    /// 输出目录
    /// </summary>
    public string? OutputDirectory { get; set; }
    
    /// <summary>
    /// 预估文件大小（字节）
    /// </summary>
    public long EstimatedSize { get; set; }
    
    /// <summary>
    /// 预估处理时间（毫秒）
    /// </summary>
    public long EstimatedDurationMs { get; set; }
    
    /// <summary>
    /// 是否启用低内存模式
    /// </summary>
    public bool LowMemoryMode { get; set; }
    
    /// <summary>
    /// 并发度
    /// </summary>
    public int Concurrency { get; set; }
    
    /// <summary>
    /// 自定义属性
    /// </summary>
    public Dictionary<string, object> Properties { get; set; } = new();
}

/// <summary>
/// 操作结果
/// </summary>
public class OperationResult
{
    /// <summary>
    /// 是否成功
    /// </summary>
    public bool Success { get; set; }
    
    /// <summary>
    /// 结果消息
    /// </summary>
    public string? Message { get; set; }
    
    /// <summary>
    /// 输出文件路径列表
    /// </summary>
    public List<string> OutputFiles { get; set; } = new();
    
    /// <summary>
    /// 处理的文件数
    /// </summary>
    public int ProcessedFiles { get; set; }
    
    /// <summary>
    /// 失败的文件数
    /// </summary>
    public int FailedFiles { get; set; }
    
    /// <summary>
    /// 处理的总行数
    /// </summary>
    public long ProcessedRows { get; set; }
    
    /// <summary>
    /// 警告列表
    /// </summary>
    public List<string> Warnings { get; set; } = new();
    
    /// <summary>
    /// 错误列表
    /// </summary>
    public List<string> Errors { get; set; } = new();
    
    /// <summary>
    /// 性能指标
    /// </summary>
    public Dictionary<string, object> Metrics { get; set; } = new();
}

/// <summary>
/// 操作统计信息
/// </summary>
public class OperationStatistics
{
    /// <summary>
    /// 操作ID
    /// </summary>
    public string OperationId { get; set; } = string.Empty;
    
    /// <summary>
    /// 操作类型
    /// </summary>
    public string OperationType { get; set; } = string.Empty;
    
    /// <summary>
    /// 开始时间
    /// </summary>
    public DateTime StartTime { get; set; }
    
    /// <summary>
    /// 结束时间
    /// </summary>
    public DateTime? EndTime { get; set; }
    
    /// <summary>
    /// 总耗时（毫秒）
    /// </summary>
    public long TotalElapsedMs => EndTime.HasValue ? 
        (long)(EndTime.Value - StartTime).TotalMilliseconds : 
        (long)(DateTime.Now - StartTime).TotalMilliseconds;
    
    /// <summary>
    /// 当前进度
    /// </summary>
    public float CurrentProgress { get; set; }
    
    /// <summary>
    /// 检查点列表
    /// </summary>
    public List<Checkpoint> Checkpoints { get; set; } = new();
    
    /// <summary>
    /// 警告数量
    /// </summary>
    public int WarningCount { get; set; }
    
    /// <summary>
    /// 错误数量
    /// </summary>
    public int ErrorCount { get; set; }
    
    /// <summary>
    /// 峰值内存使用（字节）
    /// </summary>
    public long PeakMemoryUsage { get; set; }
    
    /// <summary>
    /// 平均CPU使用率
    /// </summary>
    public double AverageCpuUsage { get; set; }
}

/// <summary>
/// 检查点
/// </summary>
public class Checkpoint
{
    /// <summary>
    /// 检查点名称
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// 时间戳
    /// </summary>
    public DateTime Timestamp { get; set; }
    
    /// <summary>
    /// 从开始的耗时（毫秒）
    /// </summary>
    public long ElapsedMs { get; set; }
    
    /// <summary>
    /// 详细信息
    /// </summary>
    public object? Details { get; set; }
}