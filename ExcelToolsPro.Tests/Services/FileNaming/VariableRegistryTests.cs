using Xunit;
using ExcelToolsPro.Services.FileNaming.Core;
using ExcelToolsPro.Services.FileNaming.Models;

namespace ExcelToolsPro.Tests.Services.FileNaming;

/// <summary>
/// 变量注册表测试
/// </summary>
public class VariableRegistryTests
{
    private readonly VariableRegistry _registry;
    
    public VariableRegistryTests()
    {
        _registry = new VariableRegistry();
    }
    
    [Fact]
    public void RegisterVariable_WithValidVariable_ShouldAddToRegistry()
    {
        // Arrange
        var variable = new CustomVariable
        {
            Id = "test_var",
            Name = "测试变量",
            Description = "用于测试的变量",
            DefaultValue = "test_value"
        };
        
        // Act
        _registry.RegisterVariable(variable);
        
        // Assert
        var retrievedVariable = _registry.GetVariable("test_var");
        Assert.NotNull(retrievedVariable);
        Assert.Equal("test_var", retrievedVariable.Id);
        Assert.Equal("测试变量", retrievedVariable.Name);
    }
    
    [Fact]
    public void RegisterVariable_WithDuplicateId_ShouldOverwriteExisting()
    {
        // Arrange
        var variable1 = new CustomVariable
        {
            Id = "duplicate_var",
            Name = "第一个变量",
            DefaultValue = "value1"
        };
        
        var variable2 = new CustomVariable
        {
            Id = "duplicate_var",
            Name = "第二个变量",
            DefaultValue = "value2"
        };
        
        // Act
        _registry.RegisterVariable(variable1);
        _registry.RegisterVariable(variable2);
        
        // Assert
        var retrievedVariable = _registry.GetVariable("duplicate_var");
        Assert.NotNull(retrievedVariable);
        Assert.Equal("第二个变量", retrievedVariable.Name);
        Assert.Equal("value2", ((CustomVariable)retrievedVariable).DefaultValue);
    }
    
    [Fact]
    public void RegisterVariable_WithNullVariable_ShouldThrowArgumentNullException()
    {
        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => _registry.RegisterVariable(null!));
    }
    
    [Fact]
    public void RegisterVariable_WithEmptyId_ShouldThrowArgumentException()
    {
        // Arrange
        var variable = new CustomVariable
        {
            Id = "", // 空ID
            Name = "无效变量"
        };
        
        // Act & Assert
        Assert.Throws<ArgumentException>(() => _registry.RegisterVariable(variable));
    }
    
    [Fact]
    public void UnregisterVariable_WithExistingVariable_ShouldRemoveFromRegistry()
    {
        // Arrange
        var variable = new CustomVariable
        {
            Id = "remove_var",
            Name = "待移除变量"
        };
        
        _registry.RegisterVariable(variable);
        
        // Act
        var result = _registry.UnregisterVariable("remove_var");
        
        // Assert
        Assert.True(result);
        var retrievedVariable = _registry.GetVariable("remove_var");
        Assert.Null(retrievedVariable);
    }
    
    [Fact]
    public void UnregisterVariable_WithNonExistingVariable_ShouldReturnFalse()
    {
        // Act
        var result = _registry.UnregisterVariable("non_existing_var");
        
        // Assert
        Assert.False(result);
    }
    
    [Fact]
    public void GetVariable_WithExistingId_ShouldReturnVariable()
    {
        // Arrange
        var variable = new DateTimeVariable
        {
            Id = "current_time",
            Name = "当前时间",
            Format = "HH:mm:ss"
        };
        
        _registry.RegisterVariable(variable);
        
        // Act
        var retrievedVariable = _registry.GetVariable("current_time");
        
        // Assert
        Assert.NotNull(retrievedVariable);
        Assert.IsType<DateTimeVariable>(retrievedVariable);
        Assert.Equal("current_time", retrievedVariable.Id);
        Assert.Equal("当前时间", retrievedVariable.Name);
    }
    
    [Fact]
    public void GetVariable_WithNonExistingId_ShouldReturnNull()
    {
        // Act
        var retrievedVariable = _registry.GetVariable("non_existing_id");
        
        // Assert
        Assert.Null(retrievedVariable);
    }
    
    [Fact]
    public void GetAllVariables_ShouldReturnAllRegisteredVariables()
    {
        // Arrange
        var variables = new List<NamingVariable>
        {
            new DateTimeVariable { Id = "date1", Name = "日期1" },
            new FileVariable { Id = "file1", Name = "文件1" },
            new CustomVariable { Id = "custom1", Name = "自定义1" }
        };
        
        foreach (var variable in variables)
        {
            _registry.RegisterVariable(variable);
        }
        
        // Act
        var allVariables = _registry.GetAllVariables();
        
        // Assert
        Assert.NotNull(allVariables);
        Assert.True(allVariables.Count >= 3); // 至少包含我们注册的3个变量
        Assert.Contains(allVariables, v => v.Id == "date1");
        Assert.Contains(allVariables, v => v.Id == "file1");
        Assert.Contains(allVariables, v => v.Id == "custom1");
    }
    
    [Fact]
    public void GetVariablesByCategory_WithValidCategory_ShouldReturnFilteredVariables()
    {
        // Arrange
        var variables = new List<NamingVariable>
        {
            new DateTimeVariable { Id = "dt1", Name = "日期时间1", Category = "DateTime" },
            new DateTimeVariable { Id = "dt2", Name = "日期时间2", Category = "DateTime" },
            new FileVariable { Id = "f1", Name = "文件1", Category = "File" },
            new CustomVariable { Id = "c1", Name = "自定义1", Category = "Custom" }
        };
        
        foreach (var variable in variables)
        {
            _registry.RegisterVariable(variable);
        }
        
        // Act
        var dateTimeVariables = _registry.GetVariablesByCategory("DateTime");
        
        // Assert
        Assert.NotNull(dateTimeVariables);
        Assert.True(dateTimeVariables.Count >= 2);
        Assert.All(dateTimeVariables, v => Assert.Equal("DateTime", v.Category));
        Assert.Contains(dateTimeVariables, v => v.Id == "dt1");
        Assert.Contains(dateTimeVariables, v => v.Id == "dt2");
    }
    
    [Fact]
    public void GetVariablesByCategory_WithNonExistingCategory_ShouldReturnEmptyList()
    {
        // Act
        var variables = _registry.GetVariablesByCategory("NonExistingCategory");
        
        // Assert
        Assert.NotNull(variables);
        Assert.Empty(variables);
    }
    
    [Fact]
    public void ValidateVariable_WithValidVariable_ShouldReturnSuccess()
    {
        // Arrange
        var variable = new DateTimeVariable
        {
            Id = "valid_datetime",
            Name = "有效日期时间",
            Description = "有效的日期时间变量",
            Format = "yyyy-MM-dd",
            Category = "DateTime"
        };
        
        // Act
        var result = _registry.ValidateVariable(variable);
        
        // Assert
        Assert.True(result.IsValid);
        Assert.Null(result.ErrorMessage);
    }
    
    [Fact]
    public void ValidateVariable_WithNullVariable_ShouldReturnFailure()
    {
        // Act
        var result = _registry.ValidateVariable(null!);
        
        // Assert
        Assert.False(result.IsValid);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("不能为空", result.ErrorMessage);
    }
    
    [Fact]
    public void ValidateVariable_WithEmptyId_ShouldReturnFailure()
    {
        // Arrange
        var variable = new CustomVariable
        {
            Id = "", // 空ID
            Name = "无效变量"
        };
        
        // Act
        var result = _registry.ValidateVariable(variable);
        
        // Assert
        Assert.False(result.IsValid);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("ID不能为空", result.ErrorMessage);
    }
    
    [Fact]
    public void ValidateVariable_WithEmptyName_ShouldReturnFailure()
    {
        // Arrange
        var variable = new CustomVariable
        {
            Id = "valid_id",
            Name = "" // 空名称
        };
        
        // Act
        var result = _registry.ValidateVariable(variable);
        
        // Assert
        Assert.False(result.IsValid);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("名称不能为空", result.ErrorMessage);
    }
    
    [Fact]
    public void ValidateVariable_WithInvalidDateTimeFormat_ShouldReturnFailure()
    {
        // Arrange
        var variable = new DateTimeVariable
        {
            Id = "invalid_datetime",
            Name = "无效日期时间",
            Format = "invalid_format" // 无效格式
        };
        
        // Act
        var result = _registry.ValidateVariable(variable);
        
        // Assert
        Assert.False(result.IsValid);
        Assert.NotNull(result.ErrorMessage);
    }
    
    [Fact]
    public void IsVariableRegistered_WithExistingVariable_ShouldReturnTrue()
    {
        // Arrange
        var variable = new CustomVariable
        {
            Id = "existing_var",
            Name = "存在的变量"
        };
        
        _registry.RegisterVariable(variable);
        
        // Act
        var isRegistered = _registry.IsVariableRegistered("existing_var");
        
        // Assert
        Assert.True(isRegistered);
    }
    
    [Fact]
    public void IsVariableRegistered_WithNonExistingVariable_ShouldReturnFalse()
    {
        // Act
        var isRegistered = _registry.IsVariableRegistered("non_existing_var");
        
        // Assert
        Assert.False(isRegistered);
    }
    
    [Fact]
    public void Clear_ShouldRemoveAllVariables()
    {
        // Arrange
        var variables = new List<NamingVariable>
        {
            new CustomVariable { Id = "var1", Name = "变量1" },
            new CustomVariable { Id = "var2", Name = "变量2" },
            new CustomVariable { Id = "var3", Name = "变量3" }
        };
        
        foreach (var variable in variables)
        {
            _registry.RegisterVariable(variable);
        }
        
        // Act
        _registry.Clear();
        
        // Assert
        var allVariables = _registry.GetAllVariables();
        Assert.Empty(allVariables);
    }
    
    [Fact]
    public void GetVariableCount_ShouldReturnCorrectCount()
    {
        // Arrange
        var initialCount = _registry.GetVariableCount();
        
        var variables = new List<NamingVariable>
        {
            new CustomVariable { Id = "count1", Name = "计数1" },
            new CustomVariable { Id = "count2", Name = "计数2" }
        };
        
        foreach (var variable in variables)
        {
            _registry.RegisterVariable(variable);
        }
        
        // Act
        var finalCount = _registry.GetVariableCount();
        
        // Assert
        Assert.Equal(initialCount + 2, finalCount);
    }
    
    [Theory]
    [InlineData("DateTime")]
    [InlineData("File")]
    [InlineData("Custom")]
    [InlineData("Index")]
    public void GetVariablesByCategory_WithDifferentCategories_ShouldReturnCorrectVariables(string category)
    {
        // Arrange
        var variables = new List<NamingVariable>
        {
            new DateTimeVariable { Id = "dt", Name = "日期时间", Category = "DateTime" },
            new FileVariable { Id = "file", Name = "文件", Category = "File" },
            new CustomVariable { Id = "custom", Name = "自定义", Category = "Custom" },
            new IndexVariable { Id = "index", Name = "索引", Category = "Index" }
        };
        
        foreach (var variable in variables)
        {
            _registry.RegisterVariable(variable);
        }
        
        // Act
        var categoryVariables = _registry.GetVariablesByCategory(category);
        
        // Assert
        Assert.NotNull(categoryVariables);
        Assert.All(categoryVariables, v => Assert.Equal(category, v.Category));
    }
}