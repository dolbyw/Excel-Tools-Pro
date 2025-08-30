using Xunit;
using Microsoft.Extensions.Logging;
using Moq;
using ExcelToolsPro.Services.FileNaming.Merge;
using ExcelToolsPro.Services.FileNaming.Core;
using ExcelToolsPro.Services.FileNaming.Models;

namespace ExcelToolsPro.Tests.Services.FileNaming;

/// <summary>
/// 合并文件命名服务测试
/// </summary>
public class MergeNamingServiceTests : IDisposable
{
    private readonly Mock<IVariableRegistry> _mockVariableRegistry;
    private readonly Mock<ILogger<MergeNamingService>> _mockLogger;
    private readonly Mock<ILogger<MergeNamingEngine>> _mockEngineLogger;
    private readonly MergeNamingService _service;
    private readonly string _testConfigPath;
    
    public MergeNamingServiceTests()
    {
        _mockVariableRegistry = new Mock<IVariableRegistry>();
        _mockLogger = new Mock<ILogger<MergeNamingService>>();
        _mockEngineLogger = new Mock<ILogger<MergeNamingEngine>>();
        
        // 设置测试配置路径
        _testConfigPath = Path.Combine(Path.GetTempPath(), "ExcelToolsPro_Test", "FileNaming", "MergeNamingConfig.json");
        Directory.CreateDirectory(Path.GetDirectoryName(_testConfigPath)!);
        
        // 设置变量注册表模拟
        SetupVariableRegistry();
        
        // 创建服务实例
        var engine = new MergeNamingEngine(_mockVariableRegistry.Object, _mockEngineLogger.Object);
        _service = new MergeNamingService(engine, _mockVariableRegistry.Object, _mockLogger.Object);
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
            new FileVariable { Id = "firstfilename", Name = "首个文件名" },
            new FileVariable { Id = "filecount", Name = "文件数量" },
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
            Id = "test_merge_template",
            Name = "测试合并模板",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>
            {
                new TextComponent { Text = "merged_" },
                new VariableComponent { VariableId = "timestamp" }
            }
        };
        
        var context = new MergeNamingContext
        {
            Template = template,
            SourceFiles = new List<Models.FileInfo>
            {
                new() { Name = "文件1.xlsx", FullPath = @"C:\源文件\文件1.xlsx" },
                new() { Name = "文件2.xlsx", FullPath = @"C:\源文件\文件2.xlsx" },
                new() { Name = "文件3.xlsx", FullPath = @"C:\源文件\文件3.xlsx" }
            },
            OutputDirectory = @"C:\输出",
            FileExtension = ".xlsx",
            MergeStrategy = MergeStrategy.Append
        };
        
        // Act
        var result = await _service.GenerateFileNameAsync(context);
        
        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.StartsWith("merged_", result);
    }
    
    [Fact]
    public async Task PreviewFileNameAsync_WithValidContext_ShouldReturnPreviewResult()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "test_merge_template",
            Name = "测试合并模板",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "firstfilename" },
                new TextComponent { Text = "_合并_" },
                new VariableComponent { VariableId = "filecount" },
                new TextComponent { Text = "个文件" }
            }
        };
        
        var context = new MergeNamingContext
        {
            Template = template,
            SourceFiles = new List<Models.FileInfo>
            {
                new() { Name = "报表1.xlsx", FullPath = @"C:\源文件\报表1.xlsx" },
                new() { Name = "报表2.xlsx", FullPath = @"C:\源文件\报表2.xlsx" }
            },
            OutputDirectory = @"C:\输出",
            FileExtension = ".xlsx",
            MergeStrategy = MergeStrategy.Append
        };
        
        // Act
        var result = await _service.PreviewFileNameAsync(context);
        
        // Assert
        Assert.NotNull(result);
        Assert.True(result.IsSuccess);
        Assert.Equal(1, result.TotalCount);
        Assert.Equal(1, result.ValidCount);
        Assert.Single(result.Items);
        
        var previewItem = result.Items.First();
        Assert.Contains("报表1", previewItem.GeneratedName);
        Assert.Contains("_合并_", previewItem.GeneratedName);
        Assert.Contains("2", previewItem.GeneratedName);
        Assert.Contains("个文件", previewItem.GeneratedName);
    }
    
    [Fact]
    public async Task GetConfigAsync_WhenConfigFileNotExists_ShouldReturnDefaultConfig()
    {
        // Act
        var config = await _service.GetConfigAsync();
        
        // Assert
        Assert.NotNull(config);
        Assert.Equal(NamingMode.Merge, config.Mode);
        Assert.NotNull(config.DefaultTemplate);
        Assert.NotEmpty(config.Templates);
        Assert.Equal("default_merge", config.DefaultTemplate.Id);
    }
    
    [Fact]
    public async Task CreateTemplateAsync_WithValidTemplate_ShouldCreateSuccessfully()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "new_merge_template",
            Name = "新合并模板",
            Description = "测试创建的新合并模板",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>
            {
                new TextComponent { Text = "合并结果_" },
                new VariableComponent { VariableId = "date" },
                new TextComponent { Text = "_" },
                new VariableComponent { VariableId = "filecount" },
                new TextComponent { Text = "files" }
            }
        };
        
        // Act
        var result = await _service.CreateTemplateAsync(template);
        
        // Assert
        Assert.NotNull(result);
        Assert.Equal(template.Id, result.Id);
        Assert.Equal(template.Name, result.Name);
        Assert.Equal(NamingMode.Merge, result.Mode);
        Assert.Equal(5, result.Components.Count);
    }
    
    [Fact]
    public async Task UpdateTemplateAsync_WithExistingTemplate_ShouldUpdateSuccessfully()
    {
        // Arrange
        var originalTemplate = new NamingTemplate
        {
            Id = "update_merge_template",
            Name = "原始合并模板",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>
            {
                new TextComponent { Text = "merged_" },
                new VariableComponent { VariableId = "timestamp" }
            }
        };
        
        await _service.CreateTemplateAsync(originalTemplate);
        
        var updatedTemplate = new NamingTemplate
        {
            Id = "update_merge_template",
            Name = "更新后的合并模板",
            Description = "已更新的合并模板描述",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>
            {
                new TextComponent { Text = "updated_merged_" },
                new VariableComponent { VariableId = "date" },
                new TextComponent { Text = "_final" }
            }
        };
        
        // Act
        await _service.UpdateTemplateAsync(updatedTemplate);
        
        // Assert
        var retrievedTemplate = await _service.GetTemplateAsync("update_merge_template");
        Assert.NotNull(retrievedTemplate);
        Assert.Equal("更新后的合并模板", retrievedTemplate.Name);
        Assert.Equal("已更新的合并模板描述", retrievedTemplate.Description);
        Assert.Equal(3, retrievedTemplate.Components.Count);
    }
    
    [Fact]
    public async Task DeleteTemplateAsync_WithExistingTemplate_ShouldDeleteSuccessfully()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "delete_merge_template",
            Name = "待删除的合并模板",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>
            {
                new TextComponent { Text = "to_delete_" },
                new VariableComponent { VariableId = "timestamp" }
            }
        };
        
        await _service.CreateTemplateAsync(template);
        
        // Act
        await _service.DeleteTemplateAsync("delete_merge_template");
        
        // Assert
        var retrievedTemplate = await _service.GetTemplateAsync("delete_merge_template");
        Assert.Null(retrievedTemplate);
    }
    
    [Fact]
    public async Task DeleteTemplateAsync_WithDefaultTemplate_ShouldThrowException()
    {
        // Arrange
        var config = await _service.GetConfigAsync();
        var defaultTemplateId = config.DefaultTemplate?.Id;
        
        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => _service.DeleteTemplateAsync(defaultTemplateId!));
    }
    
    [Fact]
    public void ValidateTemplate_WithValidMergeTemplate_ShouldReturnSuccess()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "valid_merge_template",
            Name = "有效合并模板",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "firstfilename" },
                new TextComponent { Text = "_merged_" },
                new VariableComponent { VariableId = "timestamp" }
            }
        };
        
        // Act
        var result = _service.ValidateTemplate(template);
        
        // Assert
        Assert.True(result.IsValid);
        Assert.Null(result.ErrorMessage);
    }
    
    [Fact]
    public void ValidateTemplate_WithEmptyComponents_ShouldReturnFailure()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "empty_merge_template",
            Name = "空组件合并模板",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>() // 空组件列表
        };
        
        // Act
        var result = _service.ValidateTemplate(template);
        
        // Assert
        Assert.False(result.IsValid);
        Assert.Contains("至少一个组件", result.ErrorMessage!);
    }
    
    [Fact]
    public void ValidateTemplate_WithInvalidVariableId_ShouldReturnFailure()
    {
        // Arrange
        var template = new NamingTemplate
        {
            Id = "invalid_var_template",
            Name = "无效变量模板",
            Mode = NamingMode.Merge,
            Components = new List<NamingComponent>
            {
                new VariableComponent { VariableId = "nonexistent_variable" } // 不存在的变量
            }
        };
        
        // Act
        var result = _service.ValidateTemplate(template);
        
        // Assert
        Assert.False(result.IsValid);
        Assert.Contains("不存在", result.ErrorMessage!);
    }
    
    [Fact]
    public void GetAvailableVariables_ShouldReturnMergeSpecificVariables()
    {
        // Act
        var variables = _service.GetAvailableVariables();
        
        // Assert
        Assert.NotNull(variables);
        Assert.NotEmpty(variables);
        Assert.Contains(variables, v => v.Id == "firstfilename");
        Assert.Contains(variables, v => v.Id == "filecount");
        Assert.Contains(variables, v => v.Id == "timestamp");
        Assert.Contains(variables, v => v.Id == "date");
    }
    
    [Fact]
    public void RegisterCustomVariable_ShouldAddToRegistry()
    {
        // Arrange
        var customVariable = new CustomVariable
        {
            Id = "project_name",
            Name = "项目名称",
            Description = "当前项目的名称",
            DefaultValue = "MyProject"
        };
        
        // Act
        _service.RegisterCustomVariable(customVariable);
        
        // Assert
        _mockVariableRegistry.Verify(x => x.RegisterVariable(It.Is<NamingVariable>(v => 
            v.Id == "project_name" && v.Category == "Custom")), Times.Once);
    }
    
    [Fact]
    public void UnregisterCustomVariable_ShouldRemoveFromRegistry()
    {
        // Arrange
        var variableId = "temp_variable";
        
        // Act
        _service.UnregisterCustomVariable(variableId);
        
        // Assert
        _mockVariableRegistry.Verify(x => x.UnregisterVariable(variableId), Times.Once);
    }
    
    [Fact]
    public void GenerateUniqueFileName_WithBaseName_ShouldReturnUniqueFileName()
    {
        // Arrange
        var baseName = "合并结果";
        var outputDirectory = Path.GetTempPath();
        var extension = ".xlsx";
        
        // Act
        var uniqueName = _service.GenerateUniqueFileName(baseName, outputDirectory, extension);
        
        // Assert
        Assert.NotNull(uniqueName);
        Assert.Contains(baseName, uniqueName);
    }
    
    [Fact]
    public void HasFileNameConflict_WithExistingFile_ShouldReturnTrue()
    {
        // Arrange
        var fileName = "conflict_test.xlsx";
        var outputDirectory = Path.GetTempPath();
        var fullPath = Path.Combine(outputDirectory, fileName);
        
        // 创建冲突文件
        File.WriteAllText(fullPath, "test content");
        
        try
        {
            // Act
            var hasConflict = _service.HasFileNameConflict(fileName, outputDirectory);
            
            // Assert
            Assert.True(hasConflict);
        }
        finally
        {
            // 清理
            if (File.Exists(fullPath))
            {
                File.Delete(fullPath);
            }
        }
    }
    
    [Fact]
    public void HasFileNameConflict_WithNonExistingFile_ShouldReturnFalse()
    {
        // Arrange
        var fileName = "non_existing_file.xlsx";
        var outputDirectory = Path.GetTempPath();
        
        // Act
        var hasConflict = _service.HasFileNameConflict(fileName, outputDirectory);
        
        // Assert
        Assert.False(hasConflict);
    }
    
    [Fact]
    public async Task ResetToDefaultAsync_ShouldRestoreDefaultConfiguration()
    {
        // Arrange
        // 先修改配置
        var customConfig = new NamingConfig
        {
            Mode = NamingMode.Merge,
            DefaultTemplate = new NamingTemplate
            {
                Id = "custom_default",
                Name = "自定义默认模板",
                Mode = NamingMode.Merge,
                Components = new List<NamingComponent>
                {
                    new TextComponent { Text = "custom_" }
                }
            },
            Templates = new List<NamingTemplate>(),
            GlobalSettings = new GlobalNamingSettings()
        };
        
        await _service.SaveConfigAsync(customConfig);
        
        // Act
        await _service.ResetToDefaultAsync();
        
        // Assert
        var resetConfig = await _service.GetConfigAsync();
        Assert.Equal("default_merge", resetConfig.DefaultTemplate?.Id);
        Assert.Contains("merged_", resetConfig.DefaultTemplate?.Components.OfType<TextComponent>().First().Text);
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