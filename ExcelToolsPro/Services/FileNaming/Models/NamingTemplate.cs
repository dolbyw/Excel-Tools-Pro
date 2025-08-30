namespace ExcelToolsPro.Services.FileNaming.Models;

/// <summary>
/// 命名模式
/// </summary>
public enum NamingMode
{
    /// <summary>
    /// 合并模式
    /// </summary>
    Merge,
    
    /// <summary>
    /// 拆分模式
    /// </summary>
    Split
}

/// <summary>
/// 命名模板
/// </summary>
public class NamingTemplate
{
    /// <summary>
    /// 模板ID
    /// </summary>
    public string Id { get; set; } = string.Empty;
    
    /// <summary>
    /// 模板名称
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// 模板描述
    /// </summary>
    public string Description { get; set; } = string.Empty;
    
    /// <summary>
    /// 命名模式
    /// </summary>
    public NamingMode Mode { get; set; }
    
    /// <summary>
    /// 模板组件列表
    /// </summary>
    public List<NamingComponent> Components { get; set; } = new();
    
    /// <summary>
    /// 是否为默认模板
    /// </summary>
    public bool IsDefault { get; set; }
    
    /// <summary>
    /// 创建时间
    /// </summary>
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    
    /// <summary>
    /// 更新时间
    /// </summary>
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}

/// <summary>
/// 命名组件基类
/// </summary>
public abstract class NamingComponent
{
    /// <summary>
    /// 组件类型
    /// </summary>
    public abstract string ComponentType { get; }
    
    /// <summary>
    /// 生成组件值
    /// </summary>
    /// <param name="context">命名上下文</param>
    /// <returns>生成的值</returns>
    public abstract string GenerateValue(object context);
}

/// <summary>
/// 变量组件
/// </summary>
public class VariableComponent : NamingComponent
{
    public override string ComponentType => "Variable";
    
    /// <summary>
    /// 变量ID
    /// </summary>
    public string VariableId { get; set; } = string.Empty;
    
    /// <summary>
    /// 变量参数
    /// </summary>
    public Dictionary<string, object> Parameters { get; set; } = new();
    
    public override string GenerateValue(object context)
    {
        // 这里需要通过变量注册表获取变量并生成值
        return VariableId;
    }
}

/// <summary>
/// 文本组件
/// </summary>
public class TextComponent : NamingComponent
{
    public override string ComponentType => "Text";
    
    /// <summary>
    /// 文本内容
    /// </summary>
    public string Text { get; set; } = string.Empty;
    
    public override string GenerateValue(object context)
    {
        return Text;
    }
}