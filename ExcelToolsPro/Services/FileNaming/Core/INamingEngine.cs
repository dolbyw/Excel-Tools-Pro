using ExcelToolsPro.Services.FileNaming.Models;

namespace ExcelToolsPro.Services.FileNaming.Core;

/// <summary>
/// 命名引擎接口
/// </summary>
public interface INamingEngine
{
    /// <summary>
    /// 生成文件名
    /// </summary>
    /// <param name="template">命名模板</param>
    /// <param name="context">命名上下文</param>
    /// <returns>生成的文件名</returns>
    string GenerateFileName(NamingTemplate template, SplitNamingContext context);
    
    /// <summary>
    /// 验证模板
    /// </summary>
    /// <param name="template">命名模板</param>
    /// <returns>验证结果</returns>
    Models.ValidationResult ValidateTemplate(NamingTemplate template);
    
    /// <summary>
    /// 验证文件名
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <returns>验证结果</returns>
    Models.ValidationResult ValidateFileName(string fileName);
}