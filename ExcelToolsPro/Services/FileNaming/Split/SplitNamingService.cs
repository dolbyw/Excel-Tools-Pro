using ExcelToolsPro.Services.FileNaming.Core;
using ExcelToolsPro.Services.FileNaming.Models;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using System.IO;
using System.Linq;

namespace ExcelToolsPro.Services.FileNaming.Split;

/// <summary>
/// 拆分文件命名服务实现
/// </summary>
public class SplitNamingService : ISplitNamingService
{
    private readonly SplitNamingEngine _engine;
    private readonly IVariableRegistry _variableRegistry;
    private readonly ILogger<SplitNamingService> _logger;
    private readonly string _configFilePath;
    private NamingConfig? _cachedConfig;
    
    public SplitNamingService(
        SplitNamingEngine engine, 
        IVariableRegistry variableRegistry, 
        ILogger<SplitNamingService> logger)
    {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _variableRegistry = variableRegistry ?? throw new ArgumentNullException(nameof(variableRegistry));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        
        // 设置配置文件路径
        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var configDir = Path.Combine(appDataPath, "ExcelToolsPro", "FileNaming");
        Directory.CreateDirectory(configDir);
        _configFilePath = Path.Combine(configDir, "SplitNamingConfig.json");
    }
    
    public Task<string> GenerateFileNameAsync(SplitNamingContext context)
    {
        if (context?.Template == null)
        {
            throw new ArgumentException("命名上下文和模板不能为空", nameof(context));
        }
        
        try
        {
            _logger.LogDebug("开始生成文件名，模板: {TemplateId}, 索引: {Index}", 
                context.Template.Id, context.CurrentIndex);
            
            var fileName = _engine.GenerateFileName(context.Template, context);
            
            // 添加文件扩展名
            if (!string.IsNullOrWhiteSpace(context.FileExtension))
            {
                if (!context.FileExtension.StartsWith("."))
                {
                    fileName += "." + context.FileExtension;
                }
                else
                {
                    fileName += context.FileExtension;
                }
            }
            
            _logger.LogDebug("文件名生成完成: {FileName}", fileName);
            return Task.FromResult(fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "生成文件名时发生错误");
            throw;
        }
    }
    
    public async Task<PreviewResult> PreviewFileNamesAsync(SplitNamingContext context)
    {
        if (context?.Template == null)
        {
            return new PreviewResult
            {
                IsSuccess = false,
                ErrorMessage = "命名上下文和模板不能为空"
            };
        }
        
        try
        {
            _logger.LogDebug("开始预览文件名，模板: {TemplateId}, 总数: {TotalParts}", 
                context.Template.Id, context.TotalParts);
            
            var result = new PreviewResult
            {
                IsSuccess = true,
                TotalCount = context.TotalParts
            };
            
            var validCount = 0;
            
            for (int i = 1; i <= context.TotalParts; i++)
            {
                var previewContext = new SplitNamingContext
                {
                    Template = context.Template,
                    SourceFile = context.SourceFile,
                    OutputDirectory = context.OutputDirectory,
                    FileExtension = context.FileExtension,
                    CurrentIndex = i,
                    TotalParts = context.TotalParts,
                    CustomVariables = new Dictionary<string, object>(context.CustomVariables)
                };
                
                try
                {
                    var fileName = await GenerateFileNameAsync(previewContext);
                    var fullPath = Path.Combine(context.OutputDirectory, fileName);
                    
                    var validation = ValidateFileName(fileName);
                    
                    var item = new PreviewItem
                    {
                        Index = i,
                        GeneratedName = fileName,
                        FullPath = fullPath,
                        IsValid = validation.IsValid,
                        ValidationMessage = validation.IsValid ? "有效" : string.Join("; ", validation.Errors)
                    };
                    
                    result.Items.Add(item);
                    
                    if (validation.IsValid)
                    {
                        validCount++;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "预览第 {Index} 个文件名时发生错误", i);
                    
                    result.Items.Add(new PreviewItem
                    {
                        Index = i,
                        GeneratedName = $"[错误: {ex.Message}]",
                        IsValid = false,
                        ValidationMessage = ex.Message
                    });
                }
            }
            
            result.ValidCount = validCount;
            
            _logger.LogDebug("文件名预览完成，总数: {Total}, 有效: {Valid}", 
                result.TotalCount, result.ValidCount);
            
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "预览文件名时发生错误");
            return new PreviewResult
            {
                IsSuccess = false,
                ErrorMessage = $"预览过程中发生错误: {ex.Message}"
            };
        }
    }
    
    public async Task<NamingConfig> GetConfigAsync()
    {
        try
        {
            if (_cachedConfig != null)
            {
                return _cachedConfig;
            }
            
            if (File.Exists(_configFilePath))
            {
                _logger.LogDebug("从文件加载配置: {ConfigPath}", _configFilePath);
                
                var json = await File.ReadAllTextAsync(_configFilePath);
                var config = JsonSerializer.Deserialize<NamingConfig>(json);
                
                if (config != null)
                {
                    _cachedConfig = config;
                    _logger.LogDebug("配置加载成功");
                    return config;
                }
            }
            
            _logger.LogDebug("配置文件不存在或无效，创建默认配置");
            _cachedConfig = CreateDefaultConfig();
            return _cachedConfig;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "获取配置时发生错误");
            return CreateDefaultConfig();
        }
    }
    
    public async Task SaveConfigAsync(NamingConfig config)
    {
        if (config == null)
        {
            throw new ArgumentNullException(nameof(config));
        }
        
        try
        {
            _logger.LogDebug("保存配置到文件: {ConfigPath}", _configFilePath);
            
            config.LastUpdated = DateTime.Now;
            
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };
            
            var json = JsonSerializer.Serialize(config, options);
            await File.WriteAllTextAsync(_configFilePath, json);
            
            _cachedConfig = config;
            
            _logger.LogDebug("配置保存成功");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "保存配置时发生错误");
            throw;
        }
    }
    
    public async Task<NamingTemplate> CreateTemplateAsync(NamingTemplate template)
    {
        if (template == null)
        {
            throw new ArgumentNullException(nameof(template));
        }
        
        try
        {
            var validation = ValidateTemplate(template);
            if (!validation.IsValid)
            {
                throw new ArgumentException($"模板验证失败: {string.Join("; ", validation.Errors)}");
            }
            
            var config = await GetConfigAsync();
            
            // 检查ID是否已存在
            if (config.Templates.Any(t => t.Id == template.Id))
            {
                throw new InvalidOperationException($"模板ID '{template.Id}' 已存在");
            }
            
            template.CreatedAt = DateTime.Now;
            template.UpdatedAt = DateTime.Now;
            
            config.Templates.Add(template);
            await SaveConfigAsync(config);
            
            _logger.LogInformation("模板创建成功: {TemplateId} - {TemplateName}", template.Id, template.Name);
            return template;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "创建模板时发生错误: {TemplateId}", template.Id);
            throw;
        }
    }
    
    public async Task<NamingTemplate> UpdateTemplateAsync(NamingTemplate template)
    {
        if (template == null)
        {
            throw new ArgumentNullException(nameof(template));
        }
        
        try
        {
            var validation = ValidateTemplate(template);
            if (!validation.IsValid)
            {
                throw new ArgumentException($"模板验证失败: {string.Join("; ", validation.Errors)}");
            }
            
            var config = await GetConfigAsync();
            
            var existingTemplate = config.Templates.FirstOrDefault(t => t.Id == template.Id);
            if (existingTemplate == null)
            {
                throw new InvalidOperationException($"模板 '{template.Id}' 不存在");
            }
            
            // 保留创建时间
            template.CreatedAt = existingTemplate.CreatedAt;
            template.UpdatedAt = DateTime.Now;
            
            // 替换模板
            var index = config.Templates.IndexOf(existingTemplate);
            config.Templates[index] = template;
            
            await SaveConfigAsync(config);
            
            _logger.LogInformation("模板更新成功: {TemplateId} - {TemplateName}", template.Id, template.Name);
            return template;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "更新模板时发生错误: {TemplateId}", template.Id);
            throw;
        }
    }
    
    public async Task<bool> DeleteTemplateAsync(string templateId)
    {
        if (string.IsNullOrWhiteSpace(templateId))
        {
            throw new ArgumentException("模板ID不能为空", nameof(templateId));
        }
        
        try
        {
            var config = await GetConfigAsync();
            
            var template = config.Templates.FirstOrDefault(t => t.Id == templateId);
            if (template == null)
            {
                return false;
            }
            
            // 不能删除默认模板
            if (config.DefaultTemplate?.Id == templateId)
            {
                throw new InvalidOperationException("不能删除默认模板");
            }
            
            config.Templates.Remove(template);
            await SaveConfigAsync(config);
            
            _logger.LogInformation("模板删除成功: {TemplateId} - {TemplateName}", template.Id, template.Name);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "删除模板时发生错误: {TemplateId}", templateId);
            throw;
        }
    }
    
    public ValidationResult ValidateTemplate(NamingTemplate template)
    {
        return _engine.ValidateTemplate(template);
    }
    
    public ValidationResult ValidateFileName(string fileName)
    {
        return _engine.ValidateFileName(fileName);
    }
    
    public IEnumerable<NamingVariable> GetAvailableVariables()
    {
        return _variableRegistry.GetAllVariables();
    }
    
    public string GenerateUniqueFileName(string baseName, string outputDirectory, string extension)
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
            var fileName = baseName;
            if (!string.IsNullOrWhiteSpace(extension))
            {
                if (!extension.StartsWith("."))
                {
                    fileName += "." + extension;
                }
                else
                {
                    fileName += extension;
                }
            }
            
            var fullPath = Path.Combine(outputDirectory, fileName);
            
            // 如果文件不存在，直接返回
            if (!File.Exists(fullPath))
            {
                return fileName;
            }
            
            // 生成唯一文件名
            var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            var ext = Path.GetExtension(fileName);
            
            for (int i = 1; i <= 9999; i++)
            {
                var uniqueName = $"{nameWithoutExt}_{i}{ext}";
                var uniquePath = Path.Combine(outputDirectory, uniqueName);
                
                if (!File.Exists(uniquePath))
                {
                    return uniqueName;
                }
            }
            
            // 如果还是冲突，使用GUID
            var guidName = $"{nameWithoutExt}_{Guid.NewGuid():N[..8]}{ext}";
            return guidName;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "生成唯一文件名时发生错误: {BaseName}", baseName);
            throw;
        }
    }
    
    /// <summary>
    /// 创建默认配置
    /// </summary>
    /// <returns>默认配置</returns>
    private NamingConfig CreateDefaultConfig()
    {
        var defaultTemplate = new NamingTemplate
        {
            Id = "default_split",
            Name = "默认拆分模板",
            Description = "默认的文件拆分命名模板",
            Mode = NamingMode.Split,
            IsDefault = true,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "filename" },
                new TextComponent { Text = "_part" },
                new VariableComponent { VariableId = "index" }
            }
        };
        
        return new NamingConfig
        {
            Mode = NamingMode.Split,
            DefaultTemplate = defaultTemplate,
            Templates = new List<NamingTemplate> { defaultTemplate },
            GlobalSettings = new GlobalNamingSettings()
        };
    }
}