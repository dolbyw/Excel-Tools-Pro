using System.IO;
using System.Text;

namespace ExcelToolsPro.Services;

/// <summary>
/// 文件系统服务接口
/// </summary>
public interface IFileSystemService
{
    /// <summary>
    /// 检查文件是否存在
    /// </summary>
    bool FileExists(string filePath);

    /// <summary>
    /// 检查目录是否存在
    /// </summary>
    bool DirectoryExists(string directoryPath);

    /// <summary>
    /// 创建目录
    /// </summary>
    void CreateDirectory(string directoryPath);

    /// <summary>
    /// 删除文件
    /// </summary>
    void DeleteFile(string filePath);

    /// <summary>
    /// 删除目录
    /// </summary>
    void DeleteDirectory(string directoryPath, bool recursive = false);

    /// <summary>
    /// 获取文件大小
    /// </summary>
    long GetFileSize(string filePath);

    /// <summary>
    /// 获取文件最后修改时间
    /// </summary>
    DateTime GetLastWriteTime(string filePath);

    /// <summary>
    /// 复制文件
    /// </summary>
    void CopyFile(string sourceFilePath, string destinationFilePath, bool overwrite = false);

    /// <summary>
    /// 移动文件
    /// </summary>
    void MoveFile(string sourceFilePath, string destinationFilePath);

    /// <summary>
    /// 获取临时文件路径
    /// </summary>
    string GetTempFilePath(string extension = ".tmp");

    /// <summary>
    /// 清理临时文件
    /// </summary>
    Task CleanupTempFilesAsync(TimeSpan olderThan);

    /// <summary>
    /// 异步复制文件
    /// </summary>
    Task CopyFileAsync(string sourceFilePath, string destinationFilePath, bool overwrite = false, CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步移动文件
    /// </summary>
    Task MoveFileAsync(string sourceFilePath, string destinationFilePath, CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步读取文件内容
    /// </summary>
    Task<string> ReadAllTextAsync(string filePath, Encoding? encoding = null, CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步写入文件内容
    /// </summary>
    Task WriteAllTextAsync(string filePath, string content, Encoding? encoding = null, CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步读取文件字节
    /// </summary>
    Task<byte[]> ReadAllBytesAsync(string filePath, CancellationToken cancellationToken = default);

    /// <summary>
    /// 异步写入文件字节
    /// </summary>
    Task WriteAllBytesAsync(string filePath, byte[] bytes, CancellationToken cancellationToken = default);

    /// <summary>
    /// 创建异步文件流（用于大文件处理）
    /// </summary>
    Task<FileStream> CreateFileStreamAsync(string filePath, FileMode mode, FileAccess access, FileShare share = FileShare.Read, int bufferSize = 4096, CancellationToken cancellationToken = default);
}

/// <summary>
/// 错误恢复服务接口
/// </summary>
public interface IErrorRecoveryService
{
    /// <summary>
    /// 重试失败的操作
    /// </summary>
    Task<RetryResult> RetryFailedOperationAsync(string taskId, RetryOptions? options = null);

    /// <summary>
    /// 获取处理检查点
    /// </summary>
    Task<CheckpointInfo> GetProcessingCheckpointAsync(string taskId);

    /// <summary>
    /// 清理临时文件
    /// </summary>
    Task<CleanupResult> CleanupTempFilesAsync(string? taskId = null);
}

/// <summary>
/// 重试结果
/// </summary>
public class RetryResult
{
    /// <summary>
    /// 是否成功
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// 新任务ID
    /// </summary>
    public string? NewTaskId { get; set; }

    /// <summary>
    /// 消息
    /// </summary>
    public string Message { get; set; } = string.Empty;
}

/// <summary>
/// 重试选项
/// </summary>
public class RetryOptions
{
    /// <summary>
    /// 最大重试次数
    /// </summary>
    public int MaxRetries { get; set; } = 3;

    /// <summary>
    /// 重试延迟
    /// </summary>
    public TimeSpan RetryDelay { get; set; } = TimeSpan.FromSeconds(1);

    /// <summary>
    /// 是否使用指数退避
    /// </summary>
    public bool UseExponentialBackoff { get; set; } = true;
}

/// <summary>
/// 检查点信息
/// </summary>
public class CheckpointInfo
{
    /// <summary>
    /// 检查点
    /// </summary>
    public ProcessingCheckpoint? Checkpoint { get; set; }

    /// <summary>
    /// 是否可以恢复
    /// </summary>
    public bool CanResume { get; set; }
}

/// <summary>
/// 处理检查点
/// </summary>
public class ProcessingCheckpoint
{
    /// <summary>
    /// 任务ID
    /// </summary>
    public string TaskId { get; set; } = string.Empty;

    /// <summary>
    /// 已完成的文件
    /// </summary>
    public List<string> CompletedFiles { get; set; } = new();

    /// <summary>
    /// 当前文件
    /// </summary>
    public string? CurrentFile { get; set; }

    /// <summary>
    /// 进度百分比
    /// </summary>
    public float ProgressPercent { get; set; }

    /// <summary>
    /// 时间戳
    /// </summary>
    public DateTime Timestamp { get; set; }

    /// <summary>
    /// 临时文件
    /// </summary>
    public List<string> TempFiles { get; set; } = new();
}

/// <summary>
/// 清理结果
/// </summary>
public class CleanupResult
{
    /// <summary>
    /// 已清理的文件
    /// </summary>
    public string[] CleanedFiles { get; set; } = Array.Empty<string>();

    /// <summary>
    /// 释放的空间(MB)
    /// </summary>
    public long FreedSpaceMB { get; set; }
}