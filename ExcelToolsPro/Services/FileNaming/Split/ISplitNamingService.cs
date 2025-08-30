using ExcelToolsPro.Services.FileNaming.Models;

namespace ExcelToolsPro.Services.FileNaming.Split;

/// <summary>
/// 拆分文件命名服务接口
/// </summary>
public interface ISplitNamingService
{
    /// <summary>
    /// 生成文件名
    /// </summary>
    /// <param name="context">命名上下文</param>
    /// <returns>生成的文件名</returns>
    Task<string> GenerateFileNameAsync(SplitNamingContext context);
    
    /// <summary>
    /// 预览文件名
    /// </summary>
    /// <param name="context">命名上下文</param>
    /// <returns>预览结果</returns>
    Task<PreviewResult> PreviewFileNamesAsync(SplitNamingContext context);
    
    /// <summary>
    /// 获取配置
    /// </summary>
    /// <returns>命名配置</returns>
    Task<NamingConfig> GetConfigAsync();
    
    /// <summary>
    /// 保存配置
    /// </summary>
    /// <param name="config">命名配置</param>
    Task SaveConfigAsync(NamingConfig config);
    
    /// <summary>
    /// 创建模板
    /// </summary>
    /// <param name="template">命名模板</param>
    /// <returns>创建的模板</returns>
    Task<NamingTemplate> CreateTemplateAsync(NamingTemplate template);
    
    /// <summary>
    /// 更新模板
    /// </summary>
    /// <param name="template">命名模板</param>
    /// <returns>更新的模板</returns>
    Task<NamingTemplate> UpdateTemplateAsync(NamingTemplate template);
    
    /// <summary>
    /// 删除模板
    /// </summary>
    /// <param name="templateId">模板ID</param>
    /// <returns>是否成功删除</returns>
    Task<bool> DeleteTemplateAsync(string templateId);
    
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
    
    /// <summary>
    /// 获取可用变量
    /// </summary>
    /// <returns>变量列表</returns>
    IEnumerable<NamingVariable> GetAvailableVariables();
    
    /// <summary>
    /// 生成唯一文件名
    /// </summary>
    /// <param name="baseName">基础文件名</param>
    /// <param name="outputDirectory">输出目录</param>
    /// <param name="extension">文件扩展名</param>
    /// <returns>唯一文件名</returns>
    string GenerateUniqueFileName(string baseName, string outputDirectory, string extension);
}