using ExcelToolsPro.Models;
using System.IO;

namespace ExcelToolsPro.Services;

/// <summary>
/// 低内存处理器接口
/// </summary>
public interface ILowMemoryProcessor
{
    /// <summary>
    /// 检查是否应该启用低内存模式
    /// </summary>
    /// <param name="filePaths">文件路径列表</param>
    /// <param name="config">应用配置</param>
    /// <returns>是否启用低内存模式</returns>
    bool ShouldUseLowMemoryMode(string[] filePaths, AppConfig config);
    
    /// <summary>
    /// 低内存模式处理CSV文件转换
    /// </summary>
    /// <param name="inputPath">输入文件路径</param>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="config">应用配置</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>处理结果</returns>
    Task<ProcessingResult> ProcessCsvToExcelLowMemoryAsync(
        string inputPath, 
        string outputPath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 低内存模式处理Excel文件转CSV
    /// </summary>
    /// <param name="inputPath">输入文件路径</param>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="config">应用配置</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>处理结果</returns>
    Task<ProcessingResult> ProcessExcelToCsvLowMemoryAsync(
        string inputPath, 
        string outputPath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
    
    /// <summary>
    /// 低内存模式处理HTML表格转换
    /// </summary>
    /// <param name="inputPath">输入文件路径</param>
    /// <param name="outputPath">输出文件路径</param>
    /// <param name="config">应用配置</param>
    /// <param name="progress">进度报告</param>
    /// <param name="cancellationToken">取消令牌</param>
    /// <returns>处理结果</returns>
    Task<ProcessingResult> ProcessHtmlToExcelLowMemoryAsync(
        string inputPath, 
        string outputPath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default);
}