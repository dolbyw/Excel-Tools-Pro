using Xunit;
using Microsoft.Extensions.Logging;
using Moq;
using ExcelToolsPro.Services.FileNaming.Split;
using ExcelToolsPro.Services.FileNaming.Core;
using ExcelToolsPro.Services.FileNaming.Models;
using System.Text.Json;

namespace ExcelToolsPro.Tests.Services.FileNaming;

/// <summary>
/// 拆分文件命名服务测试
/// </summary>
public class SplitNamingServiceTests : IDisposable
{
    private readonly Mock<IVariableRegistry> _mockVariableRegistry;
    private readonly Mock<ILogger<SplitNamingService>> _mockLogger;
    private readonly Mock<ILogger<SplitNamingEngine>> _mockEngineLogger;
    private readonly SplitNamingService _service;
    private readonly string _testConfigPath;
    
    public SplitNamingServiceTests()
    {
        _mockVariableRegistry = new Mock<IVariableRegistry>();
        _mockLogger = new Mock<ILogger<SplitNamingService>>();
        _mockEngineLogger = new Mock<ILogger<SplitNamingEngine>>();
        
        // 设置测试配置路径
        _testConfigPath = Path.Combine(Path.GetTempPath(), "ExcelToolsPro_Test", "FileNaming", "SplitNamingConfig.json");
        Directory.CreateDirectory(Path.GetDirectoryName(_testConfigPath)!);
        
        // 设置变量注册表模拟
        SetupVariableRegistry();
        
        // 创建服务实例
        var engine = new SplitNamingEngine(_mockVariableRegistry.Object, _mockEngineLogger.Object);
        _service = new SplitNamingService(engine, _mockVariableRegistry.Object, _mockLogger.Object);
    }
    
    private void SetupVariableRegistry()
    {
        var variables = new List<NamingVariable>
        {
            new DateTimeVariable { Id = "timestamp", Name = "时间戳", Format = "yyyyMMdd_HHmmss" },
            new DateTimeVariable { Id = "date", Name = "日期", Format = "yyyyMMdd" },
            new DateTimeVariable { Id = "year", Name = "年份", Format = "yyyy" },
            new DateTimeVariable { Id = "month", Name = "月份", Format = "MM" },
            new DateTimeVariable { Id = "day", Name = "日期", Format = "dd" },
            new FileVariable { Id = "filename", Name = "文件名" },
            new IndexVariable { Id = "index", Name = "索引" },
            new CustomVariable { Id = "custom", Name = "自定义文本" }
        };
        
        _mockVariableRegistry.Setup(x => x.GetAllVariables()).Returns(variables);
        
        foreach (var variable in variables)
        {
            _mockVariableRegistry.Setup(x => x.GetVariable(variable.Id)).Returns(variable);
        }
    }
    
    [Fact]
    public async Task GenerateFileNameAsync_WithValidContext_ShouldReturnValidFileName()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "test_template",
            Name = "测试模板",
            Mode = NamingMode.Split,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "filename" },
                new TextComponent { Text = "_" },
                new VariableComponent { VariableId = "index" }
            }
        };
        
        var context = new SplitNamingContext
        {
            Template = template,
            SourceFile = new Models.FileInfo
            {
                Name = "测试文件.xlsx",
                FullPath = @"C:\测试\测试文件.xlsx"
            },
            OutputDirectory = @"C:\输出",
            FileExtension = ".xlsx",
            CurrentIndex = 1,
            TotalParts = 3
        };
        
        // Act
        var result = await _service.GenerateFileNameAsync(context);
        
        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("测试文件", result);
        Assert.Contains("1", result);
    }
    
    [Fact]
    public async Task PreviewFileNamesAsync_WithValidContext_ShouldReturnPreviewResult()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "test_template",
            Name = "测试模板",
            Mode = NamingMode.Split,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "filename" },
                new TextComponent { Text = "_part" },
                new VariableComponent { VariableId = "index" }
            }
        };
        
        var context = new SplitNamingContext
        {
            Template = template,
            SourceFile = new Models.FileInfo
            {
                Name = "示例文件.xlsx",
                FullPath = @"C:\示例\示例文件.xlsx"
            },
            OutputDirectory = @"C:\输出",
            FileExtension = ".xlsx",
            TotalParts = 3
        };
        
        // Act
        var result = await _service.PreviewFileNamesAsync(context);
        
        // Assert
        Assert.NotNull(result);
        Assert.True(result.IsSuccess);
        Assert.Equal(3, result.TotalCount);
        Assert.Equal(3, result.ValidCount);
        Assert.All(result.Items, item =>
        {
            Assert.Contains("示例文件", item.GeneratedName);
            Assert.Contains("_part", item.GeneratedName);
        });
    }
    
    [Fact]
    public async Task GetConfigAsync_WhenConfigFileNotExists_ShouldReturnDefaultConfig()
    {
        // Act
        var config = await _service.GetConfigAsync();
        
        // Assert
        Assert.NotNull(config);
        Assert.Equal(NamingMode.Split, config.Mode);
        Assert.NotNull(config.DefaultTemplate);
        Assert.NotEmpty(config.Templates);
    }
    
    [Fact]
    public async Task SaveConfigAsync_WithValidConfig_ShouldSaveSuccessfully()
    {
        // Arrange
        var config = new NamingConfig
        {
            Mode = NamingMode.Split,
            DefaultTemplate = new NamingTemplate
            {
                Id = "custom_template",
                Name = "自定义模板",
                Mode = NamingMode.Split,
                Components = new List<NamingComponent>
                {
                    new VariableComponent { VariableId = "filename" },
                    new TextComponent { Text = "_custom" }
                }
            },
            Templates = new List<NamingTemplate>(),
            GlobalSettings = new GlobalNamingSettings()
        };
        
        // Act & Assert
        await _service.SaveConfigAsync(config);
        
        // 验证配置已保存
        var savedConfig = await _service.GetConfigAsync();
        Assert.Equal("custom_template", savedConfig.DefaultTemplate?.Id);
    }
    
    [Fact]
    public async Task CreateTemplateAsync_WithValidTemplate_ShouldCreateSuccessfully()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "new_template",
            Name = "新模板",
            Description = "测试创建的新模板",
            Mode = NamingMode.Split,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "filename" },
                new TextComponent { Text = "_new" },
                new VariableComponent { VariableId = "index" }
            }
        };
        
        // Act
        var result = await _service.CreateTemplateAsync(template);
        
        // Assert
        Assert.NotNull(result);
        Assert.Equal(template.Id, result.Id);
        Assert.Equal(template.Name, result.Name);
        Assert.Equal(NamingMode.Split, result.Mode);
    }
    
    [Fact]
    public async Task UpdateTemplateAsync_WithExistingTemplate_ShouldUpdateSuccessfully()
    {
        // Arrange
        var originalTemplate = new NamingTemplate
        {
            Id = "update_template",
            Name = "原始模板",
            Mode = NamingMode.Split,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "filename" }
            }
        };
        
        await _service.CreateTemplateAsync(originalTemplate);
        
        var updatedTemplate = new NamingTemplate
        {
            Id = "update_template",
            Name = "更新后的模板",
            Description = "已更新的描述",
            Mode = NamingMode.Split,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "filename" },
                new TextComponent { Text = "_updated" }
            }
        };
        
        // Act
        await _service.UpdateTemplateAsync(updatedTemplate);
        
        // Assert
        var retrievedTemplate = await _service.GetTemplateAsync("update_template");
        Assert.NotNull(retrievedTemplate);
        Assert.Equal("更新后的模板", retrievedTemplate.Name);
        Assert.Equal("已更新的描述", retrievedTemplate.Description);
        Assert.Equal(2, retrievedTemplate.Components.Count);
    }
    
    [Fact]
    public void ValidateTemplate_WithValidTemplate_ShouldReturnSuccess()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "valid_template",
            Name = "有效模板",
            Mode = NamingMode.Split,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "filename" },
                new TextComponent { Text = "_" },
                new VariableComponent { VariableId = "index" }
            }
        };
        
        // Act
        var result = _service.ValidateTemplate(template);
        
        // Assert
        Assert.True(result.IsValid);
        Assert.Null(result.ErrorMessage);
    }
    
    [Fact]
    public void ValidateTemplate_WithInvalidTemplate_ShouldReturnFailure()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "", // 无效的ID
            Name = "无效模板",
            Mode = NamingMode.Split,
            Components = new List<NamingComponent>()
        };
        
        // Act
        var result = _service.ValidateTemplate(template);
        
        // Assert
        Assert.False(result.IsValid);
        Assert.NotNull(result.ErrorMessage);
    }
    
    [Fact]
    public void ValidateFileName_WithValidFileName_ShouldReturnSuccess()
    {
        // Arrange
        var fileName = "有效的文件名_123.xlsx";
        
        // Act
        var result = _service.ValidateFileName(fileName);
        
        // Assert
        Assert.True(result.IsValid);
    }
    
    [Fact]
    public void ValidateFileName_WithInvalidFileName_ShouldReturnFailure()
    {
        // Arrange
        var fileName = "无效<文件>名*.xlsx"; // 包含非法字符
        
        // Act
        var result = _service.ValidateFileName(fileName);
        
        // Assert
        Assert.False(result.IsValid);
        Assert.NotNull(result.ErrorMessage);
    }
    
    [Fact]
    public void GetAvailableVariables_ShouldReturnVariableList()
    {
        // Act
        var variables = _service.GetAvailableVariables();
        
        // Assert
        Assert.NotNull(variables);
        Assert.NotEmpty(variables);
        Assert.Contains(variables, v => v.Id == "filename");
        Assert.Contains(variables, v => v.Id == "index");
        Assert.Contains(variables, v => v.Id == "timestamp");
    }
    
    [Fact]
    public void GenerateUniqueFileName_WithConflictingName_ShouldReturnUniqueFileName()
    {
        // Arrange
        var baseName = "测试文件";
        var outputDirectory = Path.GetTempPath();
        var extension = ".xlsx";
        
        // 创建冲突文件
        var conflictFile = Path.Combine(outputDirectory, baseName + extension);
        File.WriteAllText(conflictFile, "test");
        
        try
        {
            // Act
            var uniqueName = _service.GenerateUniqueFileName(baseName, outputDirectory, extension);
            
            // Assert
            Assert.NotEqual(baseName, uniqueName);
            Assert.Contains(baseName, uniqueName);
            Assert.False(File.Exists(Path.Combine(outputDirectory, uniqueName + extension)));
        }
        finally
        {
            // 清理
            if (File.Exists(conflictFile))
            {
                File.Delete(conflictFile);
            }
        }
    }
    
    public void Dispose()
    {
        // 清理测试配置文件
        if (File.Exists(_testConfigPath))
        {
            File.Delete(_testConfigPath);
        }
        
        var testDir = Path.GetDirectoryName(_testConfigPath);
        if (Directory.Exists(testDir) && !Directory.EnumerateFileSystemEntries(testDir).Any())
        {
            Directory.Delete(testDir, true);
        }
    }
}