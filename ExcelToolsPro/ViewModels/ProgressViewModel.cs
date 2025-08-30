using System.ComponentModel;
using System.Runtime.CompilerServices;
using Microsoft.Extensions.Logging;

namespace ExcelToolsPro.ViewModels;

/// <summary>
/// 进度视图模型
/// </summary>
public class ProgressViewModel : INotifyPropertyChanged
{
    private readonly ILogger<ProgressViewModel> _logger;
    private double _progress;
    private string _statusMessage = string.Empty;
    private bool _isIndeterminate;
    private bool _isVisible;

    public ProgressViewModel(ILogger<ProgressViewModel> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// 进度值 (0-100)
    /// </summary>
    public double Progress
    {
        get => _progress;
        set => SetProperty(ref _progress, value);
    }

    /// <summary>
    /// 状态消息
    /// </summary>
    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    /// <summary>
    /// 是否为不确定进度
    /// </summary>
    public bool IsIndeterminate
    {
        get => _isIndeterminate;
        set => SetProperty(ref _isIndeterminate, value);
    }

    /// <summary>
    /// 是否可见
    /// </summary>
    public bool IsVisible
    {
        get => _isVisible;
        set => SetProperty(ref _isVisible, value);
    }

    /// <summary>
    /// 进度文本
    /// </summary>
    public string ProgressText => $"{Progress:F1}%";

    /// <summary>
    /// 更新进度
    /// </summary>
    public void UpdateProgress(double progress, string? message = null)
    {
        Progress = Math.Max(0, Math.Min(100, progress));
        if (!string.IsNullOrEmpty(message))
        {
            StatusMessage = message;
        }
        OnPropertyChanged(nameof(ProgressText));
    }

    /// <summary>
    /// 显示进度
    /// </summary>
    public void Show(string? message = null, bool indeterminate = false)
    {
        IsVisible = true;
        IsIndeterminate = indeterminate;
        if (!string.IsNullOrEmpty(message))
        {
            StatusMessage = message;
        }
        if (!indeterminate)
        {
            Progress = 0;
        }
    }

    /// <summary>
    /// 隐藏进度
    /// </summary>
    public void Hide()
    {
        IsVisible = false;
        Progress = 0;
        StatusMessage = string.Empty;
        IsIndeterminate = false;
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