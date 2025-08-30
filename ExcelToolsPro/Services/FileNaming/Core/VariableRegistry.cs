using ExcelToolsPro.Services.FileNaming.Models;
using Microsoft.Extensions.Logging;
using System.Collections.Concurrent;
using System.Linq;

namespace ExcelToolsPro.Services.FileNaming.Core;

/// <summary>
/// 变量注册表实现
/// </summary>
public class VariableRegistry : IVariableRegistry
{
    private readonly ConcurrentDictionary<string, NamingVariable> _variables;
    private readonly ILogger<VariableRegistry> _logger;
    
    public VariableRegistry(ILogger<VariableRegistry> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _variables = new ConcurrentDictionary<string, NamingVariable>();
        
        // 注册默认变量
        RegisterDefaultVariables();
    }
    
    /// <summary>
    /// 注册默认变量
    /// </summary>
    private void RegisterDefaultVariables()
    {
        try
        {
            _logger.LogDebug("开始注册默认变量...");
            
            // 日期时间变量
            RegisterVariable(new DateTimeVariable 
            { 
                Id = "timestamp", 
                Name = "时间戳", 
                Description = "当前日期时间",
                Category = "DateTime",
                Format = "yyyyMMdd_HHmmss" 
            });
            
            RegisterVariable(new DateTimeVariable 
            { 
                Id = "date", 
                Name = "日期", 
                Description = "当前日期",
                Category = "DateTime",
                Format = "yyyyMMdd" 
            });
            
            RegisterVariable(new DateTimeVariable 
            { 
                Id = "year", 
                Name = "年份", 
                Description = "当前年份",
                Category = "DateTime",
                Format = "yyyy" 
            });
            
            RegisterVariable(new DateTimeVariable 
            { 
                Id = "month", 
                Name = "月份", 
                Description = "当前月份",
                Category = "DateTime",
                Format = "MM" 
            });
            
            RegisterVariable(new DateTimeVariable 
            { 
                Id = "day", 
                Name = "日期", 
                Description = "当前日期",
                Category = "DateTime",
                Format = "dd" 
            });
            
            // 文件变量
            RegisterVariable(new FileVariable 
            { 
                Id = "filename", 
                Name = "文件名", 
                Description = "源文件名（不含扩展名）",
                Category = "File" 
            });
            
            RegisterVariable(new FileVariable 
            { 
                Id = "firstfilename", 
                Name = "首个文件名", 
                Description = "第一个文件的文件名",
                Category = "File" 
            });
            
            RegisterVariable(new FileVariable 
            { 
                Id = "filecount", 
                Name = "文件数量", 
                Description = "处理的文件总数",
                Category = "File" 
            });
            
            // 索引变量
            RegisterVariable(new IndexVariable 
            { 
                Id = "index", 
                Name = "索引", 
                Description = "当前文件索引",
                Category = "Index",
                Format = "000" 
            });
            
            _logger.LogInformation("默认变量注册完成，共注册 {Count} 个变量", _variables.Count);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "注册默认变量时发生错误");
            throw;
        }
    }
    
    public IEnumerable<NamingVariable> GetAllVariables()
    {
        return _variables.Values.ToList();
    }
    
    public NamingVariable? GetVariable(string id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return null;
        }
        
        _variables.TryGetValue(id, out var variable);
        return variable;
    }
    
    public IEnumerable<NamingVariable> GetVariablesByCategory(string category)
    {
        if (string.IsNullOrWhiteSpace(category))
        {
            return Enumerable.Empty<NamingVariable>();
        }
        
        return _variables.Values
            .Where(v => string.Equals(v.Category, category, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }
    
    public void RegisterVariable(NamingVariable variable)
    {
        if (variable == null)
        {
            throw new ArgumentNullException(nameof(variable));
        }
        
        if (string.IsNullOrWhiteSpace(variable.Id))
        {
            throw new ArgumentException("变量ID不能为空", nameof(variable));
        }
        
        try
        {
            _variables.AddOrUpdate(variable.Id, variable, (key, oldValue) => variable);
            _logger.LogDebug("变量已注册: {VariableId} - {VariableName}", variable.Id, variable.Name);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "注册变量时发生错误: {VariableId}", variable.Id);
            throw;
        }
    }
    
    public bool UnregisterVariable(string id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return false;
        }
        
        try
        {
            var removed = _variables.TryRemove(id, out var variable);
            if (removed && variable != null)
            {
                _logger.LogDebug("变量已注销: {VariableId} - {VariableName}", id, variable.Name);
            }
            return removed;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "注销变量时发生错误: {VariableId}", id);
            return false;
        }
    }
    
    public bool IsVariableRegistered(string id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return false;
        }
        
        return _variables.ContainsKey(id);
    }
    
    public IEnumerable<string> GetCategories()
    {
        return _variables.Values
            .Select(v => v.Category)
            .Where(c => !string.IsNullOrWhiteSpace(c))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(c => c)
            .ToList();
    }
}