using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services;

/// <summary>
/// Excel处理服务接口
/// </summary>
public interface IExcelProcessingService
{
    /// <summary>
    /// 合并Excel文件
    /// </summary>
    Task<ProcessingResult> MergeExcelFilesAsync(
        MergeRequest request, 
        IProgress<float>? progress = null, 
        CancellationToken cancellationToken = default);

    /// <summary>
    /// 拆分Excel文件
    /// </summary>
    Task<ProcessingResult> SplitExcelFileAsync(
        SplitRequest request, 
        IProgress<float>? progress = null, 
        CancellationToken cancellationToken = default);

    /// <summary>
    /// 验证Excel文件
    /// </summary>
    Task<ValidationResult> ValidateExcelFilesAsync(
        string[] filePaths, 
        CancellationToken cancellationToken = default);
}

/// <summary>
/// 合并请求
/// </summary>
public class MergeRequest
{
    /// <summary>
    /// 文件路径列表
    /// </summary>
    public string[] FilePaths { get; set; } = Array.Empty<string>();

    /// <summary>
    /// 输出目录
    /// </summary>
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>
    /// 是否添加表头
    /// </summary>
    public bool AddHeaders { get; set; } = true;

    /// <summary>
    /// 是否去重表头
    /// </summary>
    public bool DedupeHeaders { get; set; } = true;
}

/// <summary>
/// 拆分请求
/// </summary>
public class SplitRequest
{
    /// <summary>
    /// 文件路径
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// 输出目录
    /// </summary>
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>
    /// 拆分方式
    /// </summary>
    public SplitMode SplitBy { get; set; } = SplitMode.BySheet;

    /// <summary>
    /// 按行拆分时每个文件的行数
    /// </summary>
    public int? RowsPerFile { get; set; }

    /// <summary>
    /// 是否添加表头
    /// </summary>
    public bool AddHeaders { get; set; } = true;
}

/// <summary>
/// 拆分模式
/// </summary>
public enum SplitMode
{
    /// <summary>
    /// 按工作表拆分
    /// </summary>
    BySheet,
    
    /// <summary>
    /// 按行数拆分
    /// </summary>
    ByRows
}

/// <summary>
/// 处理结果
/// </summary>
public class ProcessingResult
{
    /// <summary>
    /// 是否成功
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// 消息
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// 输出文件
    /// </summary>
    public string? OutputFile { get; set; }

    /// <summary>
    /// 输出文件列表（用于拆分操作）
    /// </summary>
    public List<string>? OutputFiles { get; set; }

    /// <summary>
    /// 错误列表
    /// </summary>
    public List<string> Errors { get; set; } = new();
}

/// <summary>
/// 验证结果
/// </summary>
public class ValidationResult
{
    /// <summary>
    /// 有效文件
    /// </summary>
    public string[] ValidFiles { get; set; } = Array.Empty<string>();

    /// <summary>
    /// 无效文件
    /// </summary>
    public string[] InvalidFiles { get; set; } = Array.Empty<string>();

    /// <summary>
    /// HTML伪装文件
    /// </summary>
    public string[] HtmlFiles { get; set; } = Array.Empty<string>();

    /// <summary>
    /// 文件错误信息
    /// </summary>
    public Dictionary<string, string> FileErrors { get; set; } = new();
}