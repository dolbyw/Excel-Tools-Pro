using System.ComponentModel;
using System.Runtime.CompilerServices;
using Microsoft.Extensions.Logging;

namespace ExcelToolsPro.ViewModels;

/// <summary>
/// 主页视图模型
/// </summary>
public class HomeViewModel : INotifyPropertyChanged
{
    private readonly ILogger<HomeViewModel> _logger;
    private string _welcomeMessage = "欢迎使用 Excel Tools Pro";

    public HomeViewModel(ILogger<HomeViewModel> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// 欢迎消息
    /// </summary>
    public string WelcomeMessage
    {
        get => _welcomeMessage;
        set => SetProperty(ref _welcomeMessage, value);
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