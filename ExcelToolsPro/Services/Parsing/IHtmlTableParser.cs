using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services.Parsing;

/// <summary>
/// HTML表格解析器接口
/// </summary>
public interface IHtmlTableParser
{
    /// <summary>
    /// 解析HTML内容中的表格
    /// </summary>
    /// <param name="htmlContent">HTML内容</param>
    /// <param name="config">应用配置</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>表格行数据</returns>
    Task<List<List<string>>> ParseHtmlTableAsync(
        string htmlContent, 
        AppConfig config, 
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 从文件解析HTML表格
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="config">应用配置</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>表格行数据</returns>
    Task<List<List<string>>> ParseHtmlTableFromFileAsync(
        string filePath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 验证HTML内容是否包含有效表格
    /// </summary>
    /// <param name="htmlContent">HTML内容</param>
    /// <returns>是否包含有效表格</returns>
    bool ContainsValidTable(string htmlContent);
    
    /// <summary>
    /// 获取HTML表格的基本信息
    /// </summary>
    /// <param name="htmlContent">HTML内容</param>
    /// <returns>表格信息</returns>
    HtmlTableInfo GetTableInfo(string htmlContent);
}

/// <summary>
/// HTML表格信息
/// </summary>
public class HtmlTableInfo
{
    /// <summary>
    /// 表格数量
    /// </summary>
    public int TableCount { get; set; }
    
    /// <summary>
    /// 预估行数
    /// </summary>
    public int EstimatedRows { get; set; }
    
    /// <summary>
    /// 预估列数
    /// </summary>
    public int EstimatedColumns { get; set; }
    
    /// <summary>
    /// 是否包含表头
    /// </summary>
    public bool HasHeaders { get; set; }
    
    /// <summary>
    /// 内容大小（字符数）
    /// </summary>
    public int ContentSize { get; set; }
}