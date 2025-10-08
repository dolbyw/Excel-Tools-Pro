using ExcelToolsPro.Services.FileNaming.Core;
using ExcelToolsPro.Services.FileNaming.Models;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelToolsPro.Services.FileNaming.Split;

/// <summary>
/// 拆分命名引擎
/// </summary>
public partial class SplitNamingEngine(IVariableRegistry variableRegistry, ILogger<SplitNamingEngine> logger) : INamingEngine
{
    private readonly IVariableRegistry _variableRegistry = variableRegistry ?? throw new ArgumentNullException(nameof(variableRegistry));
    private readonly ILogger<SplitNamingEngine> _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    private static readonly char[] InvalidFileNameChars = Path.GetInvalidFileNameChars();
    private static readonly string[] ReservedNames = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
    
    [GeneratedRegex("_{2,}")]
    private static partial Regex UnderscoreRegex();
    
    /// <summary>
    /// 生成文件名
    /// </summary>
    /// <param name="template">命名模板</param>
    /// <param name="context">命名上下文</param>
    /// <returns>生成的文件名</returns>
    public string GenerateFileName(NamingTemplate template, SplitNamingContext context)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(context);
        
        try
        {
            _logger.LogDebug("开始生成文件名，模板: {TemplateId}, 索引: {Index}", 
                template.Id, context.CurrentIndex);
            
            var fileName = new StringBuilder();
            
            foreach (var component in template.Components)
            {
                var value = GenerateComponentValue(component, context);
                fileName.Append(value);
            }
            
            var result = fileName.ToString();
            
            // 清理非法字符和处理保留名
            result = SanitizeFileName(result);
            
            _logger.LogDebug("文件名生成完成: {FileName}", result);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "生成文件名时发生错误，模板: {TemplateId}", template.Id);
            throw;
        }
    }
    
    /// <summary>
    /// 生成组件值
    /// </summary>
    /// <param name="component">命名组件</param>
    /// <param name="context">命名上下文</param>
    /// <returns>组件值</returns>
    private string GenerateComponentValue(NamingComponent component, SplitNamingContext context)
    {
        try
        {
            switch (component)
            {
                case VariableComponent variableComponent:
                    return GenerateVariableValue(variableComponent, context);
                    
                case TextComponent textComponent:
                    return textComponent.Text ?? string.Empty;
                    
                default:
                    _logger.LogWarning("未知的组件类型: {ComponentType}", component.GetType().Name);
                    return string.Empty;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "生成组件值时发生错误，组件类型: {ComponentType}", 
                component.GetType().Name);
            return string.Empty;
        }
    }
    
    /// <summary>
    /// 生成变量值
    /// </summary>
    /// <param name="variableComponent">变量组件</param>
    /// <param name="context">命名上下文</param>
    /// <returns>变量值</returns>
    private string GenerateVariableValue(VariableComponent variableComponent, SplitNamingContext context)
    {
        var variable = _variableRegistry.GetVariable(variableComponent.VariableId);
        if (variable == null)
        {
            _logger.LogWarning("未找到变量: {VariableId}", variableComponent.VariableId);
            return $"[{variableComponent.VariableId}]";
        }
        
        try
        {
            // 检查自定义变量值
            if (context.CustomVariables.TryGetValue(variableComponent.VariableId, out var customValue))
            {
                return customValue?.ToString() ?? string.Empty;
            }
            
            return variable.GenerateValue(context);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "生成变量值时发生错误，变量: {VariableId}", variableComponent.VariableId);
            return $"[ERROR:{variableComponent.VariableId}]";
        }
    }
    
    /// <summary>
    /// 清理和规范化文件名
    /// </summary>
    /// <param name="fileName">原始文件名</param>
    /// <returns>清理后的文件名</returns>
    private static string SanitizeFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return "unnamed";
        }
        
        var cleaned = fileName;
        
        // 移除或替换非法字符
        foreach (var invalidChar in InvalidFileNameChars)
        {
            cleaned = cleaned.Replace(invalidChar, '_');
        }
        
        // 移除连续的下划线
        cleaned = UnderscoreRegex().Replace(cleaned, "_");
        
        // 移除开头和结尾的下划线、空格和点
        cleaned = cleaned.Trim('_', ' ', '.');
        
        // 确保文件名不为空
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            cleaned = "unnamed";
        }
        
        // 检查并处理Windows保留名
        cleaned = HandleReservedNames(cleaned);
        
        // 限制文件名长度（为扩展名和路径留出空间）
        if (cleaned.Length > 200)
        {
            cleaned = cleaned[..200].TrimEnd('_', ' ', '.');
        }
        
        return cleaned;
    }
    
    /// <summary>
    /// 处理Windows保留文件名
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <returns>处理后的文件名</returns>
    private static string HandleReservedNames(string fileName)
    {
        var nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        
        if (ReservedNames.Contains(nameWithoutExtension.ToUpperInvariant()))
        {
            nameWithoutExtension += "_file";
        }
        
        return string.IsNullOrEmpty(extension) ? nameWithoutExtension : nameWithoutExtension + extension;
    }
    
    /// <summary>
    /// 生成唯一文件名（处理冲突）
    /// </summary>
    /// <param name="baseName">基础文件名</param>
    /// <param name="outputDirectory">输出目录</param>
    /// <param name="extension">文件扩展名</param>
    /// <param name="conflictStrategy">冲突处理策略</param>
    /// <returns>唯一的文件名</returns>
    public string GenerateUniqueFileName(string baseName, string outputDirectory, string extension, FileConflictStrategy conflictStrategy = FileConflictStrategy.AppendNumber)
    {
        if (string.IsNullOrWhiteSpace(baseName))
        {
            throw new ArgumentException("基础文件名不能为空", nameof(baseName));
        }
        
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            throw new ArgumentException("输出目录不能为空", nameof(outputDirectory));
        }
        
        try
        {
            // 规范化基础文件名
            var sanitizedBaseName = SanitizeFileName(baseName);
            
            // 确保扩展名格式正确
            if (!string.IsNullOrWhiteSpace(extension) && !extension.StartsWith('.'))
            {
                extension = "." + extension;
            }
            
            var fileName = sanitizedBaseName + extension;
            var fullPath = Path.Combine(outputDirectory, fileName);
            
            // 如果文件不存在，直接返回
            if (!File.Exists(fullPath))
            {
                return fileName;
            }
            
            // 根据策略处理冲突
            return conflictStrategy switch
            {
                FileConflictStrategy.Overwrite => fileName,
                FileConflictStrategy.Skip => throw new InvalidOperationException($"文件已存在: {fileName}"),
                FileConflictStrategy.AppendNumber => GenerateNumberedFileName(sanitizedBaseName, extension, outputDirectory),
                FileConflictStrategy.AppendTimestamp => GenerateTimestampedFileName(sanitizedBaseName, extension, outputDirectory),
                _ => throw new ArgumentException($"不支持的冲突策略: {conflictStrategy}", nameof(conflictStrategy))
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "生成唯一文件名时发生错误: {BaseName}", baseName);
            throw;
        }
    }
    
    /// <summary>
    /// 生成带数字后缀的文件名
    /// </summary>
    private static string GenerateNumberedFileName(string baseName, string extension, string outputDirectory)
    {
        for (int i = 1; i <= 9999; i++)
        {
            var numberedFileName = $"{baseName}_{i}{extension}";
            var fullPath = Path.Combine(outputDirectory, numberedFileName);
            
            if (!File.Exists(fullPath))
            {
                return numberedFileName;
            }
        }
        
        throw new InvalidOperationException($"无法生成唯一文件名，已尝试9999个数字后缀: {baseName}");
    }
    
    /// <summary>
    /// 生成带时间戳后缀的文件名
    /// </summary>
    private static string GenerateTimestampedFileName(string baseName, string extension, string outputDirectory)
    {
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var timestampedFileName = $"{baseName}_{timestamp}{extension}";
        var fullPath = Path.Combine(outputDirectory, timestampedFileName);
        
        // 如果时间戳文件名仍然冲突，添加毫秒
        if (File.Exists(fullPath))
        {
            var timestampWithMs = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            timestampedFileName = $"{baseName}_{timestampWithMs}{extension}";
        }
        
        return timestampedFileName;
    }
    
    /// <summary>
    /// 验证模板
    /// </summary>
    /// <param name="template">命名模板</param>
    /// <returns>验证结果</returns>
    public Models.ValidationResult ValidateTemplate(NamingTemplate template)
    {
        var errors = new List<string>();
        var warnings = new List<string>();
        
        try
        {
            if (template == null)
            {
                errors.Add("模板不能为空");
                return new Models.ValidationResult { IsValid = false, Errors = errors };
            }
            
            if (string.IsNullOrWhiteSpace(template.Id))
            {
                errors.Add("模板ID不能为空");
            }
            
            if (string.IsNullOrWhiteSpace(template.Name))
            {
                errors.Add("模板名称不能为空");
            }
            
            if (template.Components == null || template.Components.Count == 0)
            {
                errors.Add("模板必须包含至少一个组件");
            }
            else
            {
                // 验证组件
                for (int i = 0; i < template.Components.Count; i++)
                {
                    var component = template.Components[i];
                    var componentErrors = ValidateComponent(component, i);
                    errors.AddRange(componentErrors);
                }
                
                // 检查是否包含文件名变量
                var hasFileNameVariable = template.Components
                    .OfType<VariableComponent>()
                    .Any(vc => vc.VariableId == "filename");
                    
                if (!hasFileNameVariable)
                {
                    warnings.Add("建议在模板中包含文件名变量以确保文件名的唯一性");
                }
            }
            
            return new Models.ValidationResult 
            { 
                IsValid = errors.Count == 0, 
                Errors = errors, 
                Warnings = warnings 
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "验证模板时发生错误: {TemplateId}", template?.Id);
            errors.Add($"验证过程中发生错误: {ex.Message}");
            return new Models.ValidationResult { IsValid = false, Errors = errors };
        }
    }
    
    /// <summary>
    /// 验证组件
    /// </summary>
    /// <param name="component">命名组件</param>
    /// <param name="index">组件索引</param>
    /// <returns>错误消息列表</returns>
    private List<string> ValidateComponent(NamingComponent component, int index)
    {
        var errors = new List<string>();
        
        if (component == null)
        {
            errors.Add($"组件 {index + 1} 不能为空");
            return errors;
        }
        
        switch (component)
        {
            case VariableComponent variableComponent:
                if (string.IsNullOrWhiteSpace(variableComponent.VariableId))
                {
                    errors.Add($"组件 {index + 1} 的变量ID不能为空");
                }
                else if (!_variableRegistry.IsVariableRegistered(variableComponent.VariableId))
                {
                    errors.Add($"组件 {index + 1} 引用的变量 '{variableComponent.VariableId}' 不存在");
                }
                break;
                
            case TextComponent textComponent:
                if (string.IsNullOrEmpty(textComponent.Text))
                {
                    errors.Add($"组件 {index + 1} 的文本内容不能为空");
                }
                break;
                
            default:
                errors.Add($"组件 {index + 1} 的类型 '{component.GetType().Name}' 不受支持");
                break;
        }
        
        return errors;
    }
    
    /// <summary>
    /// 验证文件名
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <returns>验证结果</returns>
    public Models.ValidationResult ValidateFileName(string fileName)
    {
        var errors = new List<string>();
        var warnings = new List<string>();
        
        try
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                errors.Add("文件名不能为空");
                return new Models.ValidationResult { IsValid = false, Errors = errors };
            }
            
            // 检查非法字符
            var invalidChars = fileName.Where(c => InvalidFileNameChars.Contains(c)).ToList();
            if (invalidChars.Count > 0)
            {
                errors.Add($"文件名包含非法字符: {string.Join(", ", invalidChars.Distinct())}");
            }
            
            // 检查长度
            if (fileName.Length > 255)
            {
                errors.Add($"文件名过长（{fileName.Length} 字符），最大允许 255 字符");
            }
            
            // 检查保留名称
            var nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName).ToUpperInvariant();
            if (ReservedNames.Contains(nameWithoutExtension))
            {
                errors.Add($"'{nameWithoutExtension}' 是系统保留名称，不能用作文件名");
            }
            
            // 检查是否以点开头或结尾
            if (fileName.StartsWith('.') || fileName.EndsWith('.'))
            {
                warnings.Add("文件名以点开头或结尾可能在某些系统上造成问题");
            }
            
            return new Models.ValidationResult 
            { 
                IsValid = errors.Count == 0, 
                Errors = errors, 
                Warnings = warnings 
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "验证文件名时发生错误: {FileName}", fileName);
            errors.Add($"验证过程中发生错误: {ex.Message}");
            return new Models.ValidationResult { IsValid = false, Errors = errors };
        }
    }
}