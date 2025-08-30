using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Text.Json;
using System.Diagnostics;
using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services;

/// <summary>
/// 配置服务实现
/// </summary>
public class ConfigurationService : IConfigurationService
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<ConfigurationService> _logger;
    private readonly string _configFilePath;
    private AppConfig? _cachedConfig;
    
    // 缓存JsonSerializerOptions实例以避免重复创建
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public ConfigurationService(IConfiguration configuration, ILogger<ConfigurationService> logger)
    {
        var stopwatch = Stopwatch.StartNew();
        
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _logger.LogDebug("=== ConfigurationService 模块初始化开始 ===");
        _logger.LogDebug("注入的配置服务状态: {ConfigStatus}", configuration != null ? "有效" : "无效");
        
        _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        
        // 初始化应用数据目录
        _logger.LogDebug("开始初始化应用数据目录...");
        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        _logger.LogDebug("系统应用数据路径: {AppDataPath}", appDataPath);
        
        var appFolder = Path.Combine(appDataPath, "ExcelToolsPro");
        _logger.LogDebug("应用程序数据目录: {AppFolder}", appFolder);
        
        try
        {
            Directory.CreateDirectory(appFolder);
            _logger.LogDebug("应用程序数据目录创建成功");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "创建应用程序数据目录失败: {AppFolder}", appFolder);
            throw;
        }
        
        _configFilePath = Path.Combine(appFolder, "config.json");
        _logger.LogDebug("配置文件路径: {ConfigFilePath}, 文件存在: {FileExists}", 
            _configFilePath, File.Exists(_configFilePath));
        
        if (File.Exists(_configFilePath))
        {
            var fileInfo = new System.IO.FileInfo(_configFilePath);
            _logger.LogDebug("配置文件信息 - 大小: {FileSize} bytes, 最后修改: {LastModified}", 
                fileInfo.Length, fileInfo.LastWriteTime);
        }
        
        _logger.LogInformation("ConfigurationService 模块初始化完成，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
    }

    public async Task<AppConfig> GetConfigurationAsync(CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        
        _logger.LogDebug("开始异步获取应用程序配置...");
        
        if (_cachedConfig != null)
        {
            _logger.LogDebug("返回缓存的配置，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
            return _cachedConfig;
        }

        try
        {
            _logger.LogDebug("缓存中无配置，开始异步加载配置文件...");
            
            if (File.Exists(_configFilePath))
            {
                _logger.LogDebug("发现用户配置文件，开始异步读取: {ConfigFile}", _configFilePath);
                var loadStart = stopwatch.ElapsedMilliseconds;
                
                var json = await File.ReadAllTextAsync(_configFilePath, cancellationToken).ConfigureAwait(false);
                _logger.LogDebug("配置文件异步读取完成，JSON长度: {JsonLength} 字符，耗时: {ReadElapsedMs}ms", 
                    json.Length, stopwatch.ElapsedMilliseconds - loadStart);
                
                var deserializeStart = stopwatch.ElapsedMilliseconds;
                _cachedConfig = JsonSerializer.Deserialize<AppConfig>(json) ?? CreateDefaultConfig();
                _logger.LogDebug("配置反序列化完成，耗时: {DeserializeElapsedMs}ms", 
                    stopwatch.ElapsedMilliseconds - deserializeStart);
                
                _logger.LogInformation("用户配置文件加载成功");
                LogConfigurationDetails(_cachedConfig, "用户配置");
            }
            else
            {
                _logger.LogDebug("用户配置文件不存在，从appsettings.json加载默认配置");
                var defaultConfigStart = stopwatch.ElapsedMilliseconds;
                _cachedConfig = CreateConfigFromAppSettings();
                _logger.LogDebug("默认配置创建完成，耗时: {DefaultConfigElapsedMs}ms", 
                    stopwatch.ElapsedMilliseconds - defaultConfigStart);
                
                _logger.LogInformation("默认配置加载成功");
                LogConfigurationDetails(_cachedConfig, "默认配置");
            }

            _logger.LogInformation("配置加载成功，总耗时: {TotalElapsedMs}ms", stopwatch.ElapsedMilliseconds);
            return _cachedConfig;
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("配置加载被取消");
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "加载配置时发生错误，错误类型: {ExceptionType}, 使用默认配置", ex.GetType().Name);
            
            var fallbackStart = stopwatch.ElapsedMilliseconds;
            _cachedConfig = CreateDefaultConfig();
            _logger.LogWarning("已创建后备默认配置，耗时: {FallbackElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds - fallbackStart);
            
            LogConfigurationDetails(_cachedConfig, "后备默认配置");
            return _cachedConfig;
        }
    }

    public async Task SaveConfigurationAsync(AppConfig config)
    {
        var stopwatch = Stopwatch.StartNew();
        
        _logger.LogDebug("开始保存配置到文件: {ConfigFile}", _configFilePath);
        
        if (config == null)
        {
            _logger.LogError("尝试保存空配置对象");
            throw new ArgumentNullException(nameof(config), "配置对象不能为空");
        }
        
        // 验证配置参数
        var (isValid, errors) = ValidateConfiguration(config);
        if (!isValid)
        {
            _logger.LogError("配置验证失败: {ValidationErrors}", string.Join(", ", errors));
            throw new ArgumentException($"配置验证失败: {string.Join(", ", errors)}");
        }
        
        try
        {
            _logger.LogDebug("配置验证通过，开始序列化...");
            var serializeStart = stopwatch.ElapsedMilliseconds;
            
            var json = JsonSerializer.Serialize(config, JsonOptions);
            _logger.LogDebug("配置序列化完成，JSON长度: {JsonLength} 字符，耗时: {SerializeElapsedMs}ms", 
                json.Length, stopwatch.ElapsedMilliseconds - serializeStart);
            
            var writeStart = stopwatch.ElapsedMilliseconds;
            await File.WriteAllTextAsync(_configFilePath, json).ConfigureAwait(false);
            _logger.LogDebug("配置文件写入完成，耗时: {WriteElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds - writeStart);
            
            _cachedConfig = config;
            LogConfigurationDetails(config, "已保存配置");
            
            _logger.LogInformation("配置保存成功，总耗时: {TotalElapsedMs}ms", stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "保存配置时发生错误，错误类型: {ExceptionType}, 文件路径: {FilePath}, 耗时: {ElapsedMs}ms", 
                ex.GetType().Name, _configFilePath, stopwatch.ElapsedMilliseconds);
            throw;
        }
    }

    public async Task ResetConfigurationAsync()
    {
        var stopwatch = Stopwatch.StartNew();
        
        _logger.LogDebug("开始重置配置为默认值...");
        
        try
        {
            // 删除现有配置文件
            if (File.Exists(_configFilePath))
            {
                _logger.LogDebug("删除现有配置文件: {ConfigFile}", _configFilePath);
                var deleteStart = stopwatch.ElapsedMilliseconds;
                File.Delete(_configFilePath);
                _logger.LogDebug("配置文件删除完成，耗时: {DeleteElapsedMs}ms", 
                    stopwatch.ElapsedMilliseconds - deleteStart);
            }
            else
            {
                _logger.LogDebug("配置文件不存在，无需删除");
            }

            // 创建默认配置
            _logger.LogDebug("创建默认配置...");
            var createStart = stopwatch.ElapsedMilliseconds;
            _cachedConfig = CreateDefaultConfig();
            _logger.LogDebug("默认配置创建完成，耗时: {CreateElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds - createStart);
            
            // 保存默认配置
            _logger.LogDebug("保存默认配置到文件...");
            var saveStart = stopwatch.ElapsedMilliseconds;
            await SaveConfigurationAsync(_cachedConfig).ConfigureAwait(false);
            _logger.LogDebug("默认配置保存完成，耗时: {SaveElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds - saveStart);
            
            LogConfigurationDetails(_cachedConfig, "重置后的默认配置");
            _logger.LogInformation("配置已重置为默认值，总耗时: {TotalElapsedMs}ms", stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "重置配置时发生错误，错误类型: {ExceptionType}, 耗时: {ElapsedMs}ms", 
                ex.GetType().Name, stopwatch.ElapsedMilliseconds);
            throw;
        }
    }

    private AppConfig CreateDefaultConfig()
    {
        _logger.LogDebug("创建默认配置对象...");
        
        var config = new AppConfig
        {
            AddHeaders = true,
            DedupeHeaders = true,
            OutputFormat = "xlsx",
            MaxMemoryMB = 512,
            ConcurrentTasks = 2,
            EnableDarkMode = false,
            Language = "zh-CN",
            MinimizeToTray = false,
            EnableAutoThrottle = true,
            ChunkSizeMB = 1,
            ProcessingTimeout = TimeSpan.FromMinutes(30)
        };
        
        _logger.LogDebug("默认配置创建完成");
        return config;
    }
    
    /// <summary>
    /// 验证配置参数
    /// </summary>
    private (bool IsValid, List<string> Errors) ValidateConfiguration(AppConfig config)
    {
        var errors = new List<string>();
        
        _logger.LogDebug("开始验证配置参数...");
        
        // 验证内存限制
        if (config.MaxMemoryMB <= 0 || config.MaxMemoryMB > 8192)
        {
            errors.Add($"最大内存限制无效: {config.MaxMemoryMB}MB (有效范围: 1-8192MB)");
        }
        
        // 验证并发任务数
        if (config.ConcurrentTasks <= 0 || config.ConcurrentTasks > Environment.ProcessorCount * 2)
        {
            errors.Add($"并发任务数无效: {config.ConcurrentTasks} (有效范围: 1-{Environment.ProcessorCount * 2})");
        }
        
        // 验证输出格式
        var validFormats = new[] { "xlsx", "xls", "csv" };
        if (string.IsNullOrWhiteSpace(config.OutputFormat) || !validFormats.Contains(config.OutputFormat.ToLower()))
        {
            errors.Add($"输出格式无效: {config.OutputFormat} (有效格式: {string.Join(", ", validFormats)})");
        }
        
        // 验证块大小
        if (config.ChunkSizeMB <= 0 || config.ChunkSizeMB > 100)
        {
            errors.Add($"块大小无效: {config.ChunkSizeMB}MB (有效范围: 1-100MB)");
        }
        
        // 验证处理超时
        if (config.ProcessingTimeout <= TimeSpan.Zero || config.ProcessingTimeout > TimeSpan.FromHours(24))
        {
            errors.Add($"处理超时时间无效: {config.ProcessingTimeout} (有效范围: 1秒-24小时)");
        }
        
        // 验证语言代码
        var validLanguages = new[] { "zh-CN", "en-US" };
        if (string.IsNullOrWhiteSpace(config.Language) || !validLanguages.Contains(config.Language))
        {
            errors.Add($"语言代码无效: {config.Language} (有效语言: {string.Join(", ", validLanguages)})");
        }
        
        var isValid = errors.Count == 0;
        _logger.LogDebug("配置验证完成，结果: {IsValid}, 错误数量: {ErrorCount}", 
            isValid ? "通过" : "失败", errors.Count);
        
        if (!isValid)
        {
            _logger.LogWarning("配置验证失败: {ValidationErrors}", string.Join("; ", errors));
        }
        
        return (isValid, errors);
    }
    
    /// <summary>
    /// 记录配置详细信息
    /// </summary>
    private void LogConfigurationDetails(AppConfig config, string configType)
    {
        _logger.LogDebug("=== {ConfigType} 详细信息 ===", configType);
        _logger.LogDebug("添加表头: {AddHeaders}", config.AddHeaders);
        _logger.LogDebug("表头去重: {DedupeHeaders}", config.DedupeHeaders);
        _logger.LogDebug("输出格式: {OutputFormat}", config.OutputFormat);
        _logger.LogDebug("最大内存: {MaxMemoryMB}MB", config.MaxMemoryMB);
        _logger.LogDebug("并发任务: {ConcurrentTasks}", config.ConcurrentTasks);
        _logger.LogDebug("最后输出目录: {LastOutputDirectory}", config.LastOutputDirectory ?? "(未设置)");
        _logger.LogDebug("深色模式: {EnableDarkMode}", config.EnableDarkMode);
        _logger.LogDebug("语言: {Language}", config.Language);
        _logger.LogDebug("最小化到托盘: {MinimizeToTray}", config.MinimizeToTray);
        _logger.LogDebug("自动节流: {EnableAutoThrottle}", config.EnableAutoThrottle);
        _logger.LogDebug("块大小: {ChunkSizeMB}MB", config.ChunkSizeMB);
        _logger.LogDebug("处理超时: {ProcessingTimeout}", config.ProcessingTimeout);
        _logger.LogDebug("=== {ConfigType} 信息结束 ===", configType);
    }

    private AppConfig CreateConfigFromAppSettings()
    {
        _logger.LogDebug("开始从appsettings.json创建配置...");
        
        var config = CreateDefaultConfig();
        _logger.LogDebug("默认配置对象创建完成");
        
        try
        {
            var appSettingsSection = _configuration.GetSection("AppSettings");
            _logger.LogDebug("AppSettings配置节状态: {SectionExists}", appSettingsSection.Exists() ? "存在" : "不存在");
            
            if (appSettingsSection.Exists())
            {
                _logger.LogDebug("开始从AppSettings节读取配置项...");
                
                // 逐项读取配置并记录
                var originalAddHeaders = config.AddHeaders;
                config.AddHeaders = appSettingsSection.GetValue<bool>("AddHeaders", config.AddHeaders);
                _logger.LogDebug("AddHeaders: {Original} -> {New}", originalAddHeaders, config.AddHeaders);
                
                var originalDedupeHeaders = config.DedupeHeaders;
                config.DedupeHeaders = appSettingsSection.GetValue<bool>("DedupeHeaders", config.DedupeHeaders);
                _logger.LogDebug("DedupeHeaders: {Original} -> {New}", originalDedupeHeaders, config.DedupeHeaders);
                
                var originalOutputFormat = config.OutputFormat;
                config.OutputFormat = appSettingsSection.GetValue<string>("OutputFormat", config.OutputFormat);
                _logger.LogDebug("OutputFormat: {Original} -> {New}", originalOutputFormat, config.OutputFormat);
                
                var originalMaxMemoryMB = config.MaxMemoryMB;
                config.MaxMemoryMB = appSettingsSection.GetValue<int>("MaxMemoryMB", config.MaxMemoryMB);
                _logger.LogDebug("MaxMemoryMB: {Original} -> {New}", originalMaxMemoryMB, config.MaxMemoryMB);
                
                var originalConcurrentTasks = config.ConcurrentTasks;
                config.ConcurrentTasks = appSettingsSection.GetValue<int>("ConcurrentTasks", config.ConcurrentTasks);
                _logger.LogDebug("ConcurrentTasks: {Original} -> {New}", originalConcurrentTasks, config.ConcurrentTasks);
                
                config.LastOutputDirectory = appSettingsSection.GetValue<string>("LastOutputDirectory", config.LastOutputDirectory ?? string.Empty);
                _logger.LogDebug("LastOutputDirectory: {Value}", config.LastOutputDirectory);
                
                var originalEnableDarkMode = config.EnableDarkMode;
                config.EnableDarkMode = appSettingsSection.GetValue<bool>("EnableDarkMode", config.EnableDarkMode);
                _logger.LogDebug("EnableDarkMode: {Original} -> {New}", originalEnableDarkMode, config.EnableDarkMode);
                
                var originalLanguage = config.Language;
                config.Language = appSettingsSection.GetValue<string>("Language", config.Language ?? "zh-CN");
                _logger.LogDebug("Language: {Original} -> {New}", originalLanguage, config.Language);
                
                var originalMinimizeToTray = config.MinimizeToTray;
                config.MinimizeToTray = appSettingsSection.GetValue<bool>("MinimizeToTray", config.MinimizeToTray);
                _logger.LogDebug("MinimizeToTray: {Original} -> {New}", originalMinimizeToTray, config.MinimizeToTray);
                
                var originalEnableAutoThrottle = config.EnableAutoThrottle;
                config.EnableAutoThrottle = appSettingsSection.GetValue<bool>("EnableAutoThrottle", config.EnableAutoThrottle);
                _logger.LogDebug("EnableAutoThrottle: {Original} -> {New}", originalEnableAutoThrottle, config.EnableAutoThrottle);
                
                var originalChunkSizeMB = config.ChunkSizeMB;
                config.ChunkSizeMB = appSettingsSection.GetValue<int>("ChunkSizeMB", config.ChunkSizeMB);
                _logger.LogDebug("ChunkSizeMB: {Original} -> {New}", originalChunkSizeMB, config.ChunkSizeMB);
                
                var timeoutMinutes = appSettingsSection.GetValue<int>("ProcessingTimeoutMinutes", 30);
                var originalTimeout = config.ProcessingTimeout;
                config.ProcessingTimeout = TimeSpan.FromMinutes(timeoutMinutes);
                _logger.LogDebug("ProcessingTimeout: {Original} -> {New} (来源: {TimeoutMinutes}分钟)", 
                    originalTimeout, config.ProcessingTimeout, timeoutMinutes);
                
                _logger.LogInformation("从AppSettings节成功读取配置项");
            }
            else
            {
                _logger.LogWarning("AppSettings配置节不存在，使用完全默认配置");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "从appsettings.json读取配置时发生错误，错误类型: {ExceptionType}，使用默认值", ex.GetType().Name);
        }
        
        // 验证最终配置
        var (isValid, errors) = ValidateConfiguration(config);
        if (isValid)
        {
            _logger.LogDebug("从AppSettings创建的配置验证通过");
        }
        else
        {
            _logger.LogWarning("从AppSettings创建的配置验证失败: {ValidationErrors}", 
                string.Join(", ", errors));
        }
        
        return config;
    }
}