using ClosedXML.Excel;
using ExcelToolsPro.Models;
using System.Text;

namespace ExcelToolsPro.Services.Storage;

/// <summary>
/// Excel存储服务接口
/// </summary>
public interface IExcelStorage
{
    /// <summary>
    /// 保存工作簿到文件
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <param name="filePath">文件路径</param>
    /// <param name="options">保存选项</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>保存结果</returns>
    Task<StorageResult> SaveWorkbookAsync(
        XLWorkbook workbook,
        string filePath,
        SaveOptions? options = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 保存CSV数据到文件
    /// </summary>
    /// <param name="csvData">CSV数据行</param>
    /// <param name="filePath">文件路径</param>
    /// <param name="options">保存选项</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>保存结果</returns>
    Task<StorageResult> SaveCsvDataAsync(
        List<string> csvData,
        string filePath,
        SaveOptions? options = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 批量保存工作簿列表
    /// </summary>
    /// <param name="workbooks">工作簿列表</param>
    /// <param name="outputDirectory">输出目录</param>
    /// <param name="fileNameTemplate">文件名模板</param>
    /// <param name="options">保存选项</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>保存结果列表</returns>
    Task<List<StorageResult>> SaveWorkbooksBatchAsync(
        List<XLWorkbook> workbooks,
        string outputDirectory,
        string fileNameTemplate,
        SaveOptions? options = null,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 创建临时文件
    /// </summary>
    /// <param name="extension">文件扩展名</param>
    /// <param name="prefix">文件名前缀</param>
    /// <returns>临时文件路径</returns>
    string CreateTempFile(string extension = ".xlsx", string prefix = "ExcelTools_");
    
    /// <summary>
    /// 清理临时文件
    /// </summary>
    /// <param name="filePaths">文件路径列表</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>清理结果</returns>
    Task<CleanupResult> CleanupTempFilesAsync(
        List<string> filePaths,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 验证输出目录
    /// </summary>
    /// <param name="outputDirectory">输出目录</param>
    /// <param name="createIfNotExists">如果不存在是否创建</param>
    /// <returns>验证结果</returns>
    Task<ValidationResult> ValidateOutputDirectoryAsync(
        string outputDirectory,
        bool createIfNotExists = true);
    
    /// <summary>
    /// 获取文件大小
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>文件大小（字节）</returns>
    Task<long> GetFileSizeAsync(string filePath);
    
    /// <summary>
    /// 检查磁盘空间
    /// </summary>
    /// <param name="directory">目录路径</param>
    /// <param name="requiredBytes">需要的字节数</param>
    /// <returns>是否有足够空间</returns>
    Task<bool> HasSufficientDiskSpaceAsync(string directory, long requiredBytes);
    
    /// <summary>
    /// 原子性保存文件（先保存到临时文件，再替换目标文件）
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <param name="filePath">目标文件路径</param>
    /// <param name="options">保存选项</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>保存结果</returns>
    Task<StorageResult> SaveWorkbookAtomicAsync(
        XLWorkbook workbook,
        string filePath,
        SaveOptions? options = null,
        CancellationToken cancellationToken = default);
}

/// <summary>
/// 存储结果
/// </summary>
public class StorageResult
{
    /// <summary>
    /// 是否成功
    /// </summary>
    public bool Success { get; set; }
    
    /// <summary>
    /// 文件路径
    /// </summary>
    public string FilePath { get; set; } = string.Empty;
    
    /// <summary>
    /// 文件大小（字节）
    /// </summary>
    public long FileSize { get; set; }
    
    /// <summary>
    /// 错误消息
    /// </summary>
    public string? ErrorMessage { get; set; }
    
    /// <summary>
    /// 保存耗时（毫秒）
    /// </summary>
    public long ElapsedMilliseconds { get; set; }
    
    /// <summary>
    /// 是否使用了原子保存
    /// </summary>
    public bool UsedAtomicSave { get; set; }
}

/// <summary>
/// 清理结果
/// </summary>
public class CleanupResult
{
    /// <summary>
    /// 成功清理的文件数
    /// </summary>
    public int SuccessCount { get; set; }
    
    /// <summary>
    /// 失败的文件数
    /// </summary>
    public int FailureCount { get; set; }
    
    /// <summary>
    /// 失败的文件列表
    /// </summary>
    public List<string> FailedFiles { get; set; } = new();
    
    /// <summary>
    /// 释放的磁盘空间（字节）
    /// </summary>
    public long FreedBytes { get; set; }
}

/// <summary>
/// 验证结果
/// </summary>
public class ValidationResult
{
    /// <summary>
    /// 是否有效
    /// </summary>
    public bool IsValid { get; set; }
    
    /// <summary>
    /// 错误消息
    /// </summary>
    public string? ErrorMessage { get; set; }
    
    /// <summary>
    /// 目录是否存在
    /// </summary>
    public bool DirectoryExists { get; set; }
    
    /// <summary>
    /// 是否有写入权限
    /// </summary>
    public bool HasWritePermission { get; set; }
    
    /// <summary>
    /// 可用磁盘空间（字节）
    /// </summary>
    public long AvailableSpace { get; set; }
}

/// <summary>
/// 保存选项
/// </summary>
public class SaveOptions
{
    /// <summary>
    /// 是否覆盖现有文件
    /// </summary>
    public bool OverwriteExisting { get; set; } = true;
    
    /// <summary>
    /// 是否创建备份
    /// </summary>
    public bool CreateBackup { get; set; } = false;
    
    /// <summary>
    /// 是否使用原子保存
    /// </summary>
    public bool UseAtomicSave { get; set; } = true;
    
    /// <summary>
    /// I/O缓冲区大小（字节）
    /// </summary>
    public int BufferSize { get; set; } = 65536; // 64KB
    
    /// <summary>
    /// 文件编码（用于CSV等文本文件）
    /// </summary>
    public Encoding? Encoding { get; set; }
    
    /// <summary>
    /// 是否压缩（用于Excel文件）
    /// </summary>
    public bool Compress { get; set; } = true;
    
    /// <summary>
    /// 超时时间（毫秒）
    /// </summary>
    public int TimeoutMs { get; set; } = 30000; // 30秒
}