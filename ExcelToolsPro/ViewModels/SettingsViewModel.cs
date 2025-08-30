using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using ExcelToolsPro.Models;
using ExcelToolsPro.Services;
using Microsoft.Extensions.Logging;

namespace ExcelToolsPro.ViewModels;

/// <summary>
/// 设置视图模型
/// </summary>
public class SettingsViewModel : INotifyPropertyChanged
{
    private readonly IConfigurationService _configurationService;
    private readonly ILogger<SettingsViewModel> _logger;
    private AppConfig _config;

    public SettingsViewModel(
        IConfigurationService configurationService,
        ILogger<SettingsViewModel> logger)
    {
        _configurationService = configurationService;
        _logger = logger;
        _config = new AppConfig();
        
        SaveCommand = new RelayCommand(async () => await SaveSettings());
        ResetCommand = new RelayCommand(async () => await ResetSettings());
        
        _ = LoadSettingsAsync();
    }

    /// <summary>
    /// 应用配置
    /// </summary>
    public AppConfig Config
    {
        get => _config;
        set => SetProperty(ref _config, value);
    }

    /// <summary>
    /// 保存命令
    /// </summary>
    public ICommand SaveCommand { get; }

    /// <summary>
    /// 重置命令
    /// </summary>
    public ICommand ResetCommand { get; }

    /// <summary>
    /// 加载设置
    /// </summary>
    private async Task LoadSettingsAsync()
    {
        try
        {
            Config = await _configurationService.GetConfigurationAsync();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "加载设置失败");
        }
    }

    /// <summary>
    /// 保存设置
    /// </summary>
    private async Task SaveSettings()
    {
        try
        {
            await _configurationService.SaveConfigurationAsync(Config);
            _logger.LogInformation("设置已保存");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "保存设置失败");
        }
    }

    /// <summary>
    /// 重置设置
    /// </summary>
    private async Task ResetSettings()
    {
        try
        {
            await _configurationService.ResetConfigurationAsync();
            Config = await _configurationService.GetConfigurationAsync();
            _logger.LogInformation("设置已重置");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "重置设置失败");
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}