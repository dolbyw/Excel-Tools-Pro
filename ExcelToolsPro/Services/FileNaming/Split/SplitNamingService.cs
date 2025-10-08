using ExcelToolsPro.Services.FileNaming.Core;
using ExcelToolsPro.Services.FileNaming.Models;
using ExcelToolsPro.Models;
using ExcelToolsPro.Services;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using System.IO;
using System.Linq;

namespace ExcelToolsPro.Services.FileNaming.Split;

/// <summary>
/// 拆分文件命名服务实现
/// </summary>
public class SplitNamingService : ISplitNamingService, IDisposable
{
    private readonly SplitNamingEngine _engine;
    private readonly IVariableRegistry _variableRegistry;
    private readonly ILogger<SplitNamingService> _logger;
    private readonly AppConfig _appConfig;
    private readonly string _configFilePath;
    private readonly string _backupConfigFilePath;
    private NamingConfig? _cachedConfig;
    private readonly SemaphoreSlim _configSemaphore = new(1, 1);
    private bool _disposed = false;
    
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };
    
    public SplitNamingService(
        SplitNamingEngine engine, 
        IVariableRegistry variableRegistry, 
        ILogger<SplitNamingService> logger,
        AppConfig appConfig)
    {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _variableRegistry = variableRegistry ?? throw new ArgumentNullException(nameof(variableRegistry));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _appConfig = appConfig ?? throw new ArgumentNullException(nameof(appConfig));
        
        // 设置配置文件路径
        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var configDir = Path.Combine(appDataPath, "ExcelToolsPro", "FileNaming");
        Directory.CreateDirectory(configDir);
        _configFilePath = Path.Combine(configDir, "SplitNamingConfig.json");
        _backupConfigFilePath = Path.Combine(configDir, "SplitNamingConfig.json.bak");
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
                if (!context.FileExtension.StartsWith('.'))
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
        // 添加null检查
        if (context == null)
        {
            return new PreviewResult
            {
                IsSuccess = false,
                ErrorMessage = "命名上下文不能为空"
            };
        }

        // 对于大数据量，自动使用分页预览
        if (context.TotalParts > _appConfig.MaxPreviewItems)
        {
            return await PreviewFileNamesPagedAsync(context, 1, _appConfig.PreviewPageSize);
        }
        
        return await PreviewFileNamesPagedAsync(context, 1, context.TotalParts);
    }
    
    public async Task<PreviewResult> PreviewFileNamesPagedAsync(SplitNamingContext context, int page = 1, int pageSize = 500)
    {
        using var timer = PerformanceTimerExtensions.CreateTimer(_logger, "PreviewFileNamesPaged", 
            new { TemplateId = context?.Template?.Id, Page = page, PageSize = pageSize, TotalParts = context?.TotalParts });
        
        if (context?.Template == null)
        {
            timer.LogError("命名上下文和模板不能为空");
            return new PreviewResult
            {
                IsSuccess = false,
                ErrorMessage = "命名上下文和模板不能为空"
            };
        }
        
        try
        {
            timer.Checkpoint("参数验证和计算");
            // 使用配置的分页大小
            pageSize = Math.Min(pageSize, _appConfig.PreviewPageSize);
            page = Math.Max(1, page);
            
            var totalCount = context.TotalParts;
            var maxItems = _appConfig.MaxPreviewItems;
            var isTruncated = totalCount > maxItems;
            
            // 如果总数超过最大限制，截断到最大限制
            if (isTruncated)
            {
                totalCount = maxItems;
            }
            
            var totalPages = (int)Math.Ceiling((double)totalCount / pageSize);
            page = Math.Min(page, totalPages);
            
            var startIndex = (page - 1) * pageSize + 1;
            var endIndex = Math.Min(startIndex + pageSize - 1, totalCount);
            
            _logger.LogDebug("分页预览参数 - TemplateId: {TemplateId}, TotalParts: {TotalParts}, Page: {Page}, TotalPages: {TotalPages}, Range: {Start}-{End}, IsTruncated: {IsTruncated}", 
                context.Template.Id, context.TotalParts, page, totalPages, startIndex, endIndex, isTruncated);
            
            var result = new PreviewResult
            {
                IsSuccess = true,
                TotalCount = context.TotalParts,
                CurrentPage = page,
                PageSize = pageSize,
                TotalPages = totalPages,
                IsTruncated = isTruncated
            };
            
            if (isTruncated)
            {
                result.TruncationMessage = $"预览项目过多，已截断到前 {maxItems} 项。如需查看更多，请减少拆分数量或调整预览设置。";
            }
            
            var validCount = 0;
            
            timer.Checkpoint("开始生成预览项");
            // 只生成当前页的预览项
            for (int i = startIndex; i <= endIndex; i++)
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
                    _logger.LogWarning(ex, "预览文件名错误 - Index: {Index}, Error: {Error}", i, ex.Message);
                    
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
            
            timer.Checkpoint("预览项生成完成");
            _logger.LogDebug("分页预览完成 - Page: {Page}, TotalPages: {TotalPages}, ValidCount: {Valid}, TotalItems: {Total}, TemplateId: {TemplateId}", 
                page, totalPages, validCount, result.Items.Count, context.Template.Id);
            
            return result;
        }
        catch (Exception ex)
        {
            timer.LogError($"分页预览文件名时发生错误 - Error: {ex.Message}");
            return new PreviewResult
            {
                IsSuccess = false,
                ErrorMessage = $"预览过程中发生错误: {ex.Message}"
            };
        }
    }
    
    public async Task<NamingConfig> GetConfigAsync()
    {
        await _configSemaphore.WaitAsync();
        try
        {
            if (_cachedConfig != null)
            {
                return _cachedConfig;
            }
            
            // 尝试从主配置文件加载
            var config = await TryLoadConfigFromFile(_configFilePath);
            
            // 如果主配置文件加载失败，尝试从备份文件恢复
            if (config == null && File.Exists(_backupConfigFilePath))
            {
                _logger.LogWarning("主配置文件损坏或不存在，尝试从备份恢复: {BackupPath}", _backupConfigFilePath);
                config = await TryLoadConfigFromFile(_backupConfigFilePath);
                
                if (config != null)
                {
                    // 从备份恢复成功，重新保存主配置文件
                    await SaveConfigCoreAsync(config, false);
                    _logger.LogInformation("配置已从备份成功恢复");
                }
            }
            
            if (config != null)
            {
                _cachedConfig = config;
                _logger.LogDebug("配置加载成功");
                return config;
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
        finally
        {
            _configSemaphore.Release();
        }
    }
    
    private async Task<NamingConfig?> TryLoadConfigFromFile(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            
            _logger.LogDebug("从文件加载配置: {ConfigPath}", filePath);
            
            var json = await File.ReadAllTextAsync(filePath);
            if (string.IsNullOrWhiteSpace(json))
            {
                _logger.LogWarning("配置文件为空: {ConfigPath}", filePath);
                return null;
            }
            
            var config = JsonSerializer.Deserialize<NamingConfig>(json);
            return config;
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "配置文件JSON格式错误: {ConfigPath}", filePath);
            // 保留损坏的文件供诊断
            var corruptedPath = filePath + ".corrupted." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            try
            {
                File.Copy(filePath, corruptedPath, true);
                _logger.LogInformation("损坏的配置文件已备份到: {CorruptedPath}", corruptedPath);
            }
            catch (Exception copyEx)
            {
                _logger.LogWarning(copyEx, "无法备份损坏的配置文件");
            }
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "读取配置文件时发生错误: {ConfigPath}", filePath);
            return null;
        }
    }
    
    public async Task SaveConfigAsync(NamingConfig config)
    {
        using var timer = PerformanceTimerExtensions.CreateTimer(_logger, "SaveConfig", 
            new { TemplateCount = config?.Templates?.Count ?? 0, HasGlobalSettings = config?.GlobalSettings != null });
        
        ArgumentNullException.ThrowIfNull(config);
        
        timer.Checkpoint("等待配置信号量");
        await _configSemaphore.WaitAsync();
        try
        {
            await SaveConfigCoreAsync(config, true);
        }
        finally
        {
            _configSemaphore.Release();
        }
    }
    
    private async Task SaveConfigCoreAsync(NamingConfig config, bool createBackup)
    {
        try
        {
            _logger.LogDebug("保存配置 - ConfigPath: {ConfigPath}, CreateBackup: {CreateBackup}, TemplateCount: {TemplateCount}", 
                _configFilePath, createBackup, config.Templates?.Count ?? 0);
            
            config.LastUpdated = DateTime.Now;
            
            var options = JsonOptions;
            var json = JsonSerializer.Serialize(config, options);
            
            // 原子写入：先写入临时文件，然后替换
            var tempFilePath = _configFilePath + ".tmp";
            
            // 写入临时文件
            await File.WriteAllTextAsync(tempFilePath, json);
            
            // 创建备份（如果主文件存在且需要备份）
            if (createBackup && File.Exists(_configFilePath))
            {
                try
                {
                    File.Copy(_configFilePath, _backupConfigFilePath, true);
                    _logger.LogDebug("配置备份创建成功 - BackupPath: {BackupPath}", _backupConfigFilePath);
                }
                catch (Exception backupEx)
                {
                    _logger.LogWarning(backupEx, "创建配置备份错误 - BackupPath: {BackupPath}, Error: {Error}", 
                        _backupConfigFilePath, backupEx.Message);
                }
            }
            
            // 原子替换：使用File.Replace确保原子性
            if (File.Exists(_configFilePath))
            {
                var oldFilePath = _configFilePath + ".old";
                File.Replace(tempFilePath, _configFilePath, oldFilePath);
                
                // 清理旧文件
                try
                {
                    if (File.Exists(oldFilePath))
                    {
                        File.Delete(oldFilePath);
                    }
                }
                catch (Exception cleanupEx)
                {
                    _logger.LogDebug(cleanupEx, "清理旧配置文件时发生错误，可忽略");
                }
            }
            else
            {
                // 首次创建，直接移动临时文件
                File.Move(tempFilePath, _configFilePath);
            }
            
            _cachedConfig = config;
            
            _logger.LogDebug("配置保存成功 - ConfigPath: {ConfigPath}, FileSize: {FileSize}", 
                _configFilePath, new System.IO.FileInfo(_configFilePath).Length);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "保存配置错误 - ConfigPath: {ConfigPath}, Error: {Error}", 
                _configFilePath, ex.Message);
            
            // 清理可能残留的临时文件
            var tempFilePath = _configFilePath + ".tmp";
            try
            {
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
            catch (Exception cleanupEx)
            {
                _logger.LogDebug(cleanupEx, "清理临时文件时发生错误，可忽略");
            }
            
            throw;
        }
    }
    
    public async Task<NamingTemplate> CreateTemplateAsync(NamingTemplate template)
    {
        using var timer = PerformanceTimerExtensions.CreateTimer(_logger, "CreateTemplate", 
            new { TemplateId = template?.Id, TemplateName = template?.Name });
        
        ArgumentNullException.ThrowIfNull(template);
        
        try
        {
            timer.Checkpoint("验证模板");
            var validation = ValidateTemplate(template);
            if (!validation.IsValid)
            {
                timer.LogError($"模板验证失败: {string.Join("; ", validation.Errors)}");
                throw new ArgumentException($"模板验证失败: {string.Join("; ", validation.Errors)}");
            }
            
            timer.Checkpoint("获取配置");
            var config = await GetConfigAsync();
            
            // 检查ID是否已存在
            if (config.Templates.Any(t => t.Id == template.Id))
            {
                timer.LogError($"模板ID '{template.Id}' 已存在");
                throw new InvalidOperationException($"模板ID '{template.Id}' 已存在");
            }
            
            template.CreatedAt = DateTime.Now;
            template.UpdatedAt = DateTime.Now;
            
            config.Templates.Add(template);
            
            timer.Checkpoint("保存配置");
            await SaveConfigAsync(config);
            
            _logger.LogInformation("模板创建成功 - TemplateId: {TemplateId}, TemplateName: {TemplateName}, TotalTemplates: {TotalTemplates}", 
                template.Id, template.Name, config.Templates.Count);
            return template;
        }
        catch (Exception ex)
        {
            timer.LogError($"创建模板错误 - TemplateId: {template.Id}, Error: {ex.Message}");
            throw;
        }
    }
    
    public async Task<NamingTemplate> UpdateTemplateAsync(NamingTemplate template)
    {
        using var timer = PerformanceTimerExtensions.CreateTimer(_logger, "UpdateTemplate", 
            new { TemplateId = template?.Id, TemplateName = template?.Name });
        
        ArgumentNullException.ThrowIfNull(template);
        
        try
        {
            timer.Checkpoint("验证模板");
            var validation = ValidateTemplate(template);
            if (!validation.IsValid)
            {
                timer.LogError($"模板验证失败: {string.Join("; ", validation.Errors)}");
                throw new ArgumentException($"模板验证失败: {string.Join("; ", validation.Errors)}");
            }
            
            timer.Checkpoint("获取配置");
            var config = await GetConfigAsync();
            
            var existingTemplate = config.Templates.FirstOrDefault(t => t.Id == template.Id);
            if (existingTemplate == null)
            {
                timer.LogError($"模板 '{template.Id}' 不存在");
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
            if (template is null)
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
    
    public Models.ValidationResult ValidateTemplate(NamingTemplate template)
    {
        return _engine.ValidateTemplate(template);
    }
    
    public Models.ValidationResult ValidateFileName(string fileName)
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
                if (!extension.StartsWith('.'))
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
    private static NamingConfig CreateDefaultConfig()
    {
        var defaultTemplate = new NamingTemplate
        {
            Id = "default_split",
            Name = "默认拆分模板",
            Description = "默认的文件拆分命名模板",
            Mode = NamingMode.Split,
            IsDefault = true,
            Components =
            [
                new VariableComponent { VariableId = "filename" },
                new TextComponent { Text = "_part" },
                new VariableComponent { VariableId = "index" }
            ]
        };
        
        return new NamingConfig
        {
            Mode = NamingMode.Split,
            DefaultTemplate = defaultTemplate,
            Templates = [defaultTemplate],
            GlobalSettings = new GlobalNamingSettings()
        };
    }
    
    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    
    /// <summary>
    /// 释放资源的具体实现
    /// </summary>
    /// <param name="disposing">是否正在释放托管资源</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed && disposing)
        {
            try
            {
                _configSemaphore?.Dispose();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "释放配置信号量时发生错误");
            }
            
            _disposed = true;
        }
    }
}