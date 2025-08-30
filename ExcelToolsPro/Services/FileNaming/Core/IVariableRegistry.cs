using ExcelToolsPro.Services.FileNaming.Models;

namespace ExcelToolsPro.Services.FileNaming.Core;

/// <summary>
/// 变量注册表接口
/// </summary>
public interface IVariableRegistry
{
    /// <summary>
    /// 获取所有变量
    /// </summary>
    /// <returns>变量列表</returns>
    IEnumerable<NamingVariable> GetAllVariables();
    
    /// <summary>
    /// 根据ID获取变量
    /// </summary>
    /// <param name="id">变量ID</param>
    /// <returns>变量实例，如果不存在则返回null</returns>
    NamingVariable? GetVariable(string id);
    
    /// <summary>
    /// 根据分类获取变量
    /// </summary>
    /// <param name="category">变量分类</param>
    /// <returns>变量列表</returns>
    IEnumerable<NamingVariable> GetVariablesByCategory(string category);
    
    /// <summary>
    /// 注册变量
    /// </summary>
    /// <param name="variable">变量实例</param>
    void RegisterVariable(NamingVariable variable);
    
    /// <summary>
    /// 注销变量
    /// </summary>
    /// <param name="id">变量ID</param>
    /// <returns>是否成功注销</returns>
    bool UnregisterVariable(string id);
    
    /// <summary>
    /// 检查变量是否已注册
    /// </summary>
    /// <param name="id">变量ID</param>
    /// <returns>是否已注册</returns>
    bool IsVariableRegistered(string id);
    
    /// <summary>
    /// 获取所有变量分类
    /// </summary>
    /// <returns>分类列表</returns>
    IEnumerable<string> GetCategories();
}