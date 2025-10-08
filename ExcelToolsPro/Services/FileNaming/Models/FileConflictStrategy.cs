namespace ExcelToolsPro.Services.FileNaming.Models;

/// <summary>
/// 文件冲突处理策略
/// </summary>
public enum FileConflictStrategy
{
    /// <summary>
    /// 覆盖现有文件
    /// </summary>
    Overwrite,
    
    /// <summary>
    /// 跳过（抛出异常）
    /// </summary>
    Skip,
    
    /// <summary>
    /// 追加数字后缀（默认）
    /// </summary>
    AppendNumber,
    
    /// <summary>
    /// 追加时间戳后缀
    /// </summary>
    AppendTimestamp
}