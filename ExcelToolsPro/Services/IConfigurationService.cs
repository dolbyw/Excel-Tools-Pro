using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services;

/// <summary>
/// 配置服务接口
/// </summary>
public interface IConfigurationService
{
    /// <summary>
    /// 获取应用程序配置
    /// </summary>
    Task<AppConfig> GetConfigurationAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// 保存应用程序配置
    /// </summary>
    Task SaveConfigurationAsync(AppConfig config);

    /// <summary>
    /// 重置配置为默认值
    /// </summary>
    Task ResetConfigurationAsync();
}