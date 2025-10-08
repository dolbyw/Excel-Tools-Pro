using ClosedXML.Excel;
using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services.Transformation;

/// <summary>
/// Excel转换器接口
/// </summary>
public interface IExcelTransformer
{
    /// <summary>
    /// 合并多个工作表到一个工作表
    /// </summary>
    /// <param name="sourceWorkbooks">源工作簿列表</param>
    /// <param name="targetWorksheet">目标工作表</param>
    /// <param name="options">合并选项</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>合并结果</returns>
    Task<TransformationResult> MergeWorksheetsAsync(
        List<XLWorkbook> sourceWorkbooks,
        IXLWorksheet targetWorksheet,
        MergeOptions options,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 将数据行合并到工作表
    /// </summary>
    /// <param name="dataRows">数据行</param>
    /// <param name="targetWorksheet">目标工作表</param>
    /// <param name="startRow">起始行</param>
    /// <param name="options">合并选项</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>转换结果</returns>
    Task<TransformationResult> MergeDataRowsAsync(
        List<string[]> dataRows,
        IXLWorksheet targetWorksheet,
        int startRow,
        MergeOptions options,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 按工作表拆分工作簿
    /// </summary>
    /// <param name="sourceWorkbook">源工作簿</param>
    /// <param name="options">拆分选项</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>拆分后的工作簿列表</returns>
    Task<List<XLWorkbook>> SplitByWorksheetsAsync(
        XLWorkbook sourceWorkbook,
        SplitOptions options,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 按行数拆分工作表
    /// </summary>
    /// <param name="sourceWorksheet">源工作表</param>
    /// <param name="rowsPerFile">每个文件的行数</param>
    /// <param name="options">拆分选项</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>拆分后的工作簿列表</returns>
    Task<List<XLWorkbook>> SplitByRowsAsync(
        IXLWorksheet sourceWorksheet,
        int rowsPerFile,
        SplitOptions options,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 将工作表转换为CSV数据
    /// </summary>
    /// <param name="worksheet">工作表</param>
    /// <param name="options">转换选项</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>CSV行数据</returns>
    Task<List<string>> ConvertToCsvDataAsync(
        IXLWorksheet worksheet,
        ConversionOptions options,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 将CSV数据转换为工作表
    /// </summary>
    /// <param name="csvData">CSV数据行</param>
    /// <param name="targetWorksheet">目标工作表</param>
    /// <param name="options">转换选项</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>转换结果</returns>
    Task<TransformationResult> ConvertFromCsvDataAsync(
        List<string[]> csvData,
        IXLWorksheet targetWorksheet,
        ConversionOptions options,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
}

/// <summary>
/// 转换结果
/// </summary>
public class TransformationResult
{
    /// <summary>
    /// 是否成功
    /// </summary>
    public bool Success { get; set; }
    
    /// <summary>
    /// 处理的行数
    /// </summary>
    public int ProcessedRows { get; set; }
    
    /// <summary>
    /// 处理的列数
    /// </summary>
    public int ProcessedColumns { get; set; }
    
    /// <summary>
    /// 错误消息
    /// </summary>
    public string? ErrorMessage { get; set; }
    
    /// <summary>
    /// 警告消息列表
    /// </summary>
    public List<string> Warnings { get; set; } = new();
    
    /// <summary>
    /// 处理耗时（毫秒）
    /// </summary>
    public long ElapsedMilliseconds { get; set; }
}

/// <summary>
/// 合并选项
/// </summary>
public class MergeOptions
{
    /// <summary>
    /// 是否包含表头
    /// </summary>
    public bool IncludeHeaders { get; set; } = true;
    
    /// <summary>
    /// 是否去重表头
    /// </summary>
    public bool DedupeHeaders { get; set; } = true;
    
    /// <summary>
    /// 批处理大小
    /// </summary>
    public int BatchSize { get; set; } = 1000;
    
    /// <summary>
    /// 是否启用低内存模式
    /// </summary>
    public bool LowMemoryMode { get; set; } = false;
}

/// <summary>
/// 拆分选项
/// </summary>
public class SplitOptions
{
    /// <summary>
    /// 是否保留表头
    /// </summary>
    public bool PreserveHeaders { get; set; } = true;
    
    /// <summary>
    /// 输出文件名模板
    /// </summary>
    public string FileNameTemplate { get; set; } = "{filename}_Split_{index:D3}";
    
    /// <summary>
    /// 批处理大小
    /// </summary>
    public int BatchSize { get; set; } = 1000;
    
    /// <summary>
    /// 是否启用低内存模式
    /// </summary>
    public bool LowMemoryMode { get; set; } = false;
}

/// <summary>
/// 转换选项
/// </summary>
public class ConversionOptions
{
    /// <summary>
    /// 字符编码
    /// </summary>
    public string Encoding { get; set; } = "UTF-8";
    
    /// <summary>
    /// 是否包含BOM
    /// </summary>
    public bool IncludeBom { get; set; } = true;
    
    /// <summary>
    /// CSV分隔符
    /// </summary>
    public char Delimiter { get; set; } = ',';
    
    /// <summary>
    /// 是否包含表头
    /// </summary>
    public bool IncludeHeaders { get; set; } = true;
    
    /// <summary>
    /// 批处理大小
    /// </summary>
    public int BatchSize { get; set; } = 1000;
    
    /// <summary>
    /// 是否启用低内存模式
    /// </summary>
    public bool LowMemoryMode { get; set; } = false;
}