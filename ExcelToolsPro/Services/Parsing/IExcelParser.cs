using ClosedXML.Excel;
using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services.Parsing;

/// <summary>
/// Excel解析器接口
/// </summary>
public interface IExcelParser
{
    /// <summary>
    /// 解析Excel文件并返回工作簿
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>Excel工作簿</returns>
    Task<XLWorkbook> ParseExcelFileAsync(string filePath, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 解析CSV文件并返回数据行
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="config">应用配置</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>数据行集合</returns>
    Task<List<string[]>> ParseCsvFileAsync(
        string filePath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 解析HTML表格文件并返回数据行
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="config">应用配置</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>数据行集合</returns>
    Task<List<List<string>>> ParseHtmlTableFileAsync(
        string filePath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 检查文件是否为HTML格式的XLS文件
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>是否为HTML格式</returns>
    Task<bool> IsHtmlDisguisedFileAsync(string filePath, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 获取Excel文件的基本信息
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>文件信息</returns>
    Task<ExcelFileInfo> GetExcelFileInfoAsync(string filePath, CancellationToken cancellationToken = default);
}

/// <summary>
/// Excel文件信息
/// </summary>
public class ExcelFileInfo
{
    /// <summary>
    /// 工作表数量
    /// </summary>
    public int WorksheetCount { get; set; }
    
    /// <summary>
    /// 总行数
    /// </summary>
    public int TotalRows { get; set; }
    
    /// <summary>
    /// 总列数
    /// </summary>
    public int TotalColumns { get; set; }
    
    /// <summary>
    /// 文件大小（字节）
    /// </summary>
    public long FileSize { get; set; }
    
    /// <summary>
    /// 是否为HTML格式
    /// </summary>
    public bool IsHtmlFormat { get; set; }
    
    /// <summary>
    /// 工作表信息列表
    /// </summary>
    public List<WorksheetInfo> Worksheets { get; set; } = new();
}

/// <summary>
/// 工作表信息
/// </summary>
public class WorksheetInfo
{
    /// <summary>
    /// 工作表名称
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// 行数
    /// </summary>
    public int RowCount { get; set; }
    
    /// <summary>
    /// 列数
    /// </summary>
    public int ColumnCount { get; set; }
}