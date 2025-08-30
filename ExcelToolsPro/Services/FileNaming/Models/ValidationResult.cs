using System.Collections.Generic;
using System.Linq;

namespace ExcelToolsPro.Services.FileNaming.Models;

/// <summary>
/// 验证结果
/// </summary>
public class ValidationResult
{
    /// <summary>
    /// 是否有效
    /// </summary>
    public bool IsValid { get; set; }
    
    /// <summary>
    /// 错误消息列表
    /// </summary>
    public List<string> Errors { get; set; } = new();
    
    /// <summary>
    /// 警告消息列表
    /// </summary>
    public List<string> Warnings { get; set; } = new();
    
    /// <summary>
    /// 创建成功结果
    /// </summary>
    /// <returns>成功的验证结果</returns>
    public static ValidationResult Success() => new() { IsValid = true };
    
    /// <summary>
    /// 创建失败结果
    /// </summary>
    /// <param name="errors">错误消息</param>
    /// <returns>失败的验证结果</returns>
    public static ValidationResult Failure(params string[] errors) => new() 
    { 
        IsValid = false, 
        Errors = errors.ToList() 
    };
}