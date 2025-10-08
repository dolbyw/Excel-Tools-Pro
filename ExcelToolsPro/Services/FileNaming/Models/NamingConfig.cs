using System.IO;

namespace ExcelToolsPro.Services.FileNaming.Models;

/// <summary>
/// 命名配置
/// </summary>
public class NamingConfig
{
    /// <summary>
    /// 命名模式
    /// </summary>
    public NamingMode Mode { get; set; }
    
    /// <summary>
    /// 默认模板
    /// </summary>
    public NamingTemplate? DefaultTemplate { get; set; }
    
    /// <summary>
    /// 模板列表
    /// </summary>
    public List<NamingTemplate> Templates { get; set; } = new();
    
    /// <summary>
    /// 全局设置
    /// </summary>
    public GlobalNamingSettings GlobalSettings { get; set; } = new();
    
    /// <summary>
    /// 配置版本
    /// </summary>
    public string Version { get; set; } = "1.0";
    
    /// <summary>
    /// 最后更新时间
    /// </summary>
    public DateTime LastUpdated { get; set; } = DateTime.Now;
}

/// <summary>
/// 全局命名设置
/// </summary>
 public class GlobalNamingSettings
{
    /// <summary>
    /// 是否启用文件名冲突检测
    /// </summary>
    public bool EnableConflictDetection { get; set; } = true;
    
    /// <summary>
    /// 冲突解决策略
    /// </summary>
    public ConflictResolutionStrategy ConflictResolution { get; set; } = ConflictResolutionStrategy.AutoIncrement;
    
    /// <summary>
    /// 最大文件名长度
    /// </summary>
    public int MaxFileNameLength { get; set; } = 255;
    
    /// <summary>
    /// 是否移除非法字符
    /// </summary>
    public bool RemoveIllegalCharacters { get; set; } = true;
    
    /// <summary>
    /// 非法字符替换字符
    /// </summary>
    public string IllegalCharacterReplacement { get; set; } = "_";
}

/// <summary>
/// 冲突解决策略
/// </summary>
public enum ConflictResolutionStrategy
{
    /// <summary>
    /// 自动递增数字
    /// </summary>
    AutoIncrement,
    
    /// <summary>
    /// 覆盖现有文件
    /// </summary>
    Overwrite,
    
    /// <summary>
    /// 跳过冲突文件
    /// </summary>
    Skip,
    
    /// <summary>
    /// 询问用户
    /// </summary>
    Ask
}

/// <summary>
/// 拆分命名上下文
/// </summary>
public class SplitNamingContext
{
    /// <summary>
    /// 命名模板
    /// </summary>
    public NamingTemplate? Template { get; set; }
    
    /// <summary>
    /// 源文件信息
    /// </summary>
    public Models.FileInfo? SourceFile { get; set; }
    
    /// <summary>
    /// 输出目录
    /// </summary>
    public string OutputDirectory { get; set; } = string.Empty;
    
    /// <summary>
    /// 文件扩展名
    /// </summary>
    public string FileExtension { get; set; } = string.Empty;
    
    /// <summary>
    /// 当前索引
    /// </summary>
    public int CurrentIndex { get; set; }
    
    /// <summary>
    /// 总部分数
    /// </summary>
    public int TotalParts { get; set; }
    
    /// <summary>
    /// 自定义变量值
    /// </summary>
    public Dictionary<string, object> CustomVariables { get; set; } = new();
}

/// <summary>
/// 预览结果
/// </summary>
public class PreviewResult
{
    /// <summary>
    /// 是否成功
    /// </summary>
    public bool IsSuccess { get; set; }
    
    /// <summary>
    /// 错误消息
    /// </summary>
    public string ErrorMessage { get; set; } = string.Empty;
    
    /// <summary>
    /// 总数量
    /// </summary>
    public int TotalCount { get; set; }
    
    /// <summary>
    /// 有效数量
    /// </summary>
    public int ValidCount { get; set; }
    
    /// <summary>
    /// 预览项目列表
    /// </summary>
    public List<PreviewItem> Items { get; set; } = new();
    
    /// <summary>
    /// 当前页码（从1开始）
    /// </summary>
    public int CurrentPage { get; set; } = 1;
    
    /// <summary>
    /// 每页大小
    /// </summary>
    public int PageSize { get; set; } = 500;
    
    /// <summary>
    /// 总页数
    /// </summary>
    public int TotalPages { get; set; } = 1;
    
    /// <summary>
    /// 是否有下一页
    /// </summary>
    public bool HasNextPage => CurrentPage < TotalPages;
    
    /// <summary>
    /// 是否有上一页
    /// </summary>
    public bool HasPreviousPage => CurrentPage > 1;
    
    /// <summary>
    /// 是否被截断（总数超过最大预览限制）
    /// </summary>
    public bool IsTruncated { get; set; }
    
    /// <summary>
    /// 截断消息
    /// </summary>
    public string TruncationMessage { get; set; } = string.Empty;
}

/// <summary>
/// 预览项目
/// </summary>
public class PreviewItem
{
    /// <summary>
    /// 索引
    /// </summary>
    public int Index { get; set; }
    
    /// <summary>
    /// 生成的文件名
    /// </summary>
    public string GeneratedName { get; set; } = string.Empty;
    
    /// <summary>
    /// 完整路径
    /// </summary>
    public string FullPath { get; set; } = string.Empty;
    
    /// <summary>
    /// 是否有效
    /// </summary>
    public bool IsValid { get; set; }
    
    /// <summary>
    /// 验证消息
    /// </summary>
    public string ValidationMessage { get; set; } = string.Empty;
}