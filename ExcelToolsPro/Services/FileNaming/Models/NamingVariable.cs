using System.ComponentModel;
using System.IO;

namespace ExcelToolsPro.Services.FileNaming.Models;

/// <summary>
/// 命名变量基类
/// </summary>
public abstract class NamingVariable
{
    /// <summary>
    /// 变量ID
    /// </summary>
    public string Id { get; set; } = string.Empty;
    
    /// <summary>
    /// 变量名称
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// 变量描述
    /// </summary>
    public string Description { get; set; } = string.Empty;
    
    /// <summary>
    /// 变量分类
    /// </summary>
    public string Category { get; set; } = string.Empty;
    
    /// <summary>
    /// 是否为必需变量
    /// </summary>
    public bool IsRequired { get; set; }
    
    /// <summary>
    /// 生成变量值
    /// </summary>
    /// <param name="context">命名上下文</param>
    /// <returns>生成的值</returns>
    public abstract string GenerateValue(object context);
}

/// <summary>
/// 日期时间变量
/// </summary>
public class DateTimeVariable : NamingVariable
{
    /// <summary>
    /// 日期时间格式
    /// </summary>
    public string Format { get; set; } = "yyyyMMdd_HHmmss";
    
    public override string GenerateValue(object context)
    {
        return DateTime.Now.ToString(Format);
    }
}

/// <summary>
/// 文件变量
/// </summary>
public class FileVariable : NamingVariable
{
    public override string GenerateValue(object context)
    {
        if (context is SplitNamingContext splitContext)
        {
            return Path.GetFileNameWithoutExtension(splitContext.SourceFile?.Name ?? "file");
        }
        return "file";
    }
}

/// <summary>
/// 索引变量
/// </summary>
public class IndexVariable : NamingVariable
{
    /// <summary>
    /// 索引格式（如：000表示3位数字）
    /// </summary>
    public string Format { get; set; } = "000";
    
    public override string GenerateValue(object context)
    {
        if (context is SplitNamingContext splitContext)
        {
            return splitContext.CurrentIndex.ToString(Format);
        }
        return "1";
    }
}

/// <summary>
/// 自定义变量
/// </summary>
public class CustomVariable : NamingVariable
{
    /// <summary>
    /// 默认值
    /// </summary>
    public string DefaultValue { get; set; } = string.Empty;
    
    public override string GenerateValue(object context)
    {
        return DefaultValue;
    }
}