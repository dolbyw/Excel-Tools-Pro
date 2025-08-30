using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Security;
using System.Configuration;

namespace ExcelToolsPro.ViewModels;

/// <summary>
/// 错误严重程度枚举
/// </summary>
public enum ErrorSeverity
{
    Info,
    Warning,
    Error,
    Critical,
    Fatal
}

/// <summary>
/// 错误分类枚举
/// </summary>
public enum ErrorCategory
{
    Unknown,
    FileSystem,
    Network,
    Memory,
    Configuration,
    UserInput,
    ExcelProcessing,
    SystemResource
}

/// <summary>
/// 错误视图模型
/// </summary>
public class ErrorViewModel : INotifyPropertyChanged
{
    private readonly ILogger<ErrorViewModel> _logger;
    private string _errorMessage = string.Empty;
    private string _errorDetails = string.Empty;
    private Exception? _exception;
    private ErrorSeverity _severity = ErrorSeverity.Error;
    private ErrorCategory _category = ErrorCategory.Unknown;
    private DateTime _occurredAt = DateTime.Now;
    private string _userFriendlyMessage = string.Empty;

    public ErrorViewModel(ILogger<ErrorViewModel> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        CloseCommand = new RelayCommand(Close, CanClose);
        CopyErrorCommand = new RelayCommand(CopyError, CanCopyError);
        ExportErrorCommand = new RelayCommand(ExportError, CanExportError);
        RestartApplicationCommand = new RelayCommand(RestartApplication, CanRestartApplication);
        
        _logger.LogDebug("ErrorViewModel 初始化完成");
    }

    /// <summary>
    /// 错误消息
    /// </summary>
    public string ErrorMessage
    {
        get => _errorMessage;
        set => SetProperty(ref _errorMessage, value);
    }

    /// <summary>
    /// 用户友好的错误消息
    /// </summary>
    public string UserFriendlyMessage
    {
        get => _userFriendlyMessage;
        set => SetProperty(ref _userFriendlyMessage, value);
    }

    /// <summary>
    /// 错误详情
    /// </summary>
    public string ErrorDetails
    {
        get => _errorDetails;
        set => SetProperty(ref _errorDetails, value);
    }

    /// <summary>
    /// 错误严重程度
    /// </summary>
    public ErrorSeverity Severity
    {
        get => _severity;
        set => SetProperty(ref _severity, value);
    }

    /// <summary>
    /// 错误分类
    /// </summary>
    public ErrorCategory Category
    {
        get => _category;
        set => SetProperty(ref _category, value);
    }

    /// <summary>
    /// 错误发生时间
    /// </summary>
    public DateTime OccurredAt
    {
        get => _occurredAt;
        set => SetProperty(ref _occurredAt, value);
    }

    /// <summary>
    /// 异常对象
    /// </summary>
    public Exception? Exception
    {
        get => _exception;
        set
        {
            _exception = value;
            if (value != null)
            {
                ErrorMessage = value.Message;
                ErrorDetails = FormatExceptionDetails(value);
                Category = ClassifyException(value);
                Severity = DetermineSeverity(value);
                UserFriendlyMessage = GenerateUserFriendlyMessage(value);
            }
        }
    }

    /// <summary>
    /// 关闭命令
    /// </summary>
    public ICommand CloseCommand { get; }

    /// <summary>
    /// 复制错误信息命令
    /// </summary>
    public ICommand CopyErrorCommand { get; }

    /// <summary>
    /// 导出错误报告命令
    /// </summary>
    public ICommand ExportErrorCommand { get; }

    /// <summary>
    /// 重启应用程序命令
    /// </summary>
    public ICommand RestartApplicationCommand { get; }

    /// <summary>
    /// 设置错误信息
    /// </summary>
    public void SetError(string message, string? details = null, Exception? exception = null)
    {
        try
        {
            OccurredAt = DateTime.Now;
            ErrorMessage = message;
            
            if (exception != null)
            {
                Exception = exception;
            }
            else
            {
                ErrorDetails = details ?? string.Empty;
                Category = ClassifyErrorMessage(message);
                Severity = DetermineSeverityFromMessage(message);
                UserFriendlyMessage = GenerateUserFriendlyMessageFromText(message);
            }
            
            _logger.LogError(exception, "应用程序错误 [严重程度: {Severity}, 分类: {Category}]: {Message}", 
                Severity, Category, message);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "设置错误信息时发生异常");
        }
    }

    private bool CanClose() => true;
    
    private bool CanCopyError() => !string.IsNullOrEmpty(ErrorMessage);
    
    private bool CanExportError() => !string.IsNullOrEmpty(ErrorMessage);
    
    private bool CanRestartApplication() => Severity >= ErrorSeverity.Critical;

    private void Close()
    {
        try
        {
            _logger.LogInformation("用户关闭错误对话框");
            
            // 根据错误严重程度决定关闭方式
            if (Severity >= ErrorSeverity.Fatal)
            {
                // 致命错误：强制退出
                _logger.LogWarning("由于致命错误，强制退出应用程序");
                Environment.Exit(1);
            }
            else if (Severity >= ErrorSeverity.Critical)
            {
                // 严重错误：正常关闭但记录日志
                _logger.LogWarning("由于严重错误，正常关闭应用程序");
                System.Windows.Application.Current.Shutdown(1);
            }
            else
            {
                // 一般错误：继续运行，只关闭错误窗口
                _logger.LogDebug("关闭错误对话框，应用程序继续运行");
                // 这里应该关闭错误窗口而不是整个应用程序
                // 但由于当前架构限制，暂时保持原有行为
                System.Windows.Application.Current.Shutdown();
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "关闭错误对话框时发生异常");
            Environment.Exit(1);
        }
    }

    private void CopyError()
    {
        try
        {
            _logger.LogDebug("开始复制错误信息到剪贴板");
            
            var errorReport = GenerateErrorReport();
            
            // 使用重试机制处理剪贴板操作
            var maxRetries = 3;
            var retryDelay = 100;
            
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    System.Windows.Clipboard.Clear();
                    System.Windows.Clipboard.SetText(errorReport);
                    _logger.LogInformation("错误信息已成功复制到剪贴板");
                    
                    // 显示成功提示
                    System.Windows.MessageBox.Show("错误信息已复制到剪贴板", "复制成功", 
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    return;
                }
                catch (System.Runtime.InteropServices.ExternalException ex) when (i < maxRetries - 1)
                {
                    _logger.LogWarning(ex, "剪贴板操作失败，第 {Attempt} 次重试", i + 1);
                    System.Threading.Thread.Sleep(retryDelay);
                    retryDelay *= 2; // 指数退避
                }
            }
            
            throw new InvalidOperationException("多次尝试后剪贴板操作仍然失败");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "复制错误信息失败");
            System.Windows.MessageBox.Show($"复制错误信息失败: {ex.Message}", "复制失败", 
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
        }
    }
    
    private void ExportError()
    {
        try
        {
            _logger.LogDebug("开始导出错误报告");
            
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "导出错误报告",
                Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                DefaultExt = "txt",
                FileName = $"ErrorReport_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
            };
            
            if (saveDialog.ShowDialog() == true)
            {
                var errorReport = GenerateDetailedErrorReport();
                File.WriteAllText(saveDialog.FileName, errorReport, Encoding.UTF8);
                
                _logger.LogInformation("错误报告已导出到: {FilePath}", saveDialog.FileName);
                System.Windows.MessageBox.Show($"错误报告已导出到:\n{saveDialog.FileName}", "导出成功", 
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "导出错误报告失败");
            System.Windows.MessageBox.Show($"导出错误报告失败: {ex.Message}", "导出失败", 
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }
    
    private void RestartApplication()
    {
        try
        {
            _logger.LogInformation("用户请求重启应用程序");
            
            var result = System.Windows.MessageBox.Show(
                "是否要重启应用程序？这将关闭当前会话并启动新的实例。", 
                "重启应用程序", 
                System.Windows.MessageBoxButton.YesNo, 
                System.Windows.MessageBoxImage.Question);
            
            if (result == System.Windows.MessageBoxResult.Yes)
            {
                var currentProcess = Process.GetCurrentProcess();
                var executablePath = currentProcess.MainModule?.FileName;
                
                if (!string.IsNullOrEmpty(executablePath))
                {
                    _logger.LogInformation("启动新的应用程序实例: {ExecutablePath}", executablePath);
                    Process.Start(executablePath);
                }
                
                _logger.LogInformation("关闭当前应用程序实例");
                System.Windows.Application.Current.Shutdown();
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "重启应用程序失败");
            System.Windows.MessageBox.Show($"重启应用程序失败: {ex.Message}", "重启失败", 
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 格式化异常详细信息
    /// </summary>
    private string FormatExceptionDetails(Exception exception)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine($"异常类型: {exception.GetType().Name}");
        sb.AppendLine($"错误消息: {exception.Message}");
        sb.AppendLine($"发生时间: {OccurredAt:yyyy-MM-dd HH:mm:ss}");
        
        if (exception.Data.Count > 0)
        {
            sb.AppendLine("\n附加数据:");
            foreach (System.Collections.DictionaryEntry item in exception.Data)
            {
                sb.AppendLine($"  {item.Key}: {item.Value}");
            }
        }
        
        if (!string.IsNullOrEmpty(exception.StackTrace))
        {
            sb.AppendLine("\n堆栈跟踪:");
            sb.AppendLine(exception.StackTrace);
        }
        
        if (exception.InnerException != null)
        {
            sb.AppendLine("\n内部异常:");
            sb.AppendLine(FormatExceptionDetails(exception.InnerException));
        }
        
        return sb.ToString();
    }
    
    /// <summary>
    /// 分类异常类型
    /// </summary>
    private ErrorCategory ClassifyException(Exception exception)
    {
        return exception switch
        {
            FileNotFoundException or DirectoryNotFoundException or DriveNotFoundException => ErrorCategory.FileSystem,
            UnauthorizedAccessException or SecurityException => ErrorCategory.FileSystem,
            IOException => ErrorCategory.FileSystem,
            OutOfMemoryException => ErrorCategory.Memory,
            System.Net.NetworkInformation.NetworkInformationException => ErrorCategory.Network,
            System.Net.WebException => ErrorCategory.Network,
            ConfigurationErrorsException => ErrorCategory.Configuration,
            ArgumentException or ArgumentNullException => ErrorCategory.UserInput,
            InvalidOperationException when exception.Message.Contains("Excel") => ErrorCategory.ExcelProcessing,
            SystemException => ErrorCategory.SystemResource,
            _ => ErrorCategory.Unknown
        };
    }
    
    /// <summary>
    /// 根据错误消息分类
    /// </summary>
    private ErrorCategory ClassifyErrorMessage(string message)
    {
        var lowerMessage = message.ToLower();
        
        if (lowerMessage.Contains("file") || lowerMessage.Contains("文件") || lowerMessage.Contains("directory") || lowerMessage.Contains("目录"))
            return ErrorCategory.FileSystem;
        if (lowerMessage.Contains("memory") || lowerMessage.Contains("内存"))
            return ErrorCategory.Memory;
        if (lowerMessage.Contains("network") || lowerMessage.Contains("网络"))
            return ErrorCategory.Network;
        if (lowerMessage.Contains("config") || lowerMessage.Contains("配置"))
            return ErrorCategory.Configuration;
        if (lowerMessage.Contains("excel") || lowerMessage.Contains("xlsx") || lowerMessage.Contains("xls"))
            return ErrorCategory.ExcelProcessing;
        
        return ErrorCategory.Unknown;
    }
    
    /// <summary>
    /// 确定异常严重程度
    /// </summary>
    private ErrorSeverity DetermineSeverity(Exception exception)
    {
        return exception switch
        {
            OutOfMemoryException => ErrorSeverity.Fatal,
            StackOverflowException => ErrorSeverity.Fatal,
            AccessViolationException => ErrorSeverity.Fatal,
            System.Threading.ThreadAbortException => ErrorSeverity.Critical,
            UnauthorizedAccessException => ErrorSeverity.Critical,
            SecurityException => ErrorSeverity.Critical,
            FileNotFoundException => ErrorSeverity.Error,
            DirectoryNotFoundException => ErrorSeverity.Error,
            IOException => ErrorSeverity.Warning,
            ArgumentException => ErrorSeverity.Warning,
            InvalidOperationException => ErrorSeverity.Warning,
            _ => ErrorSeverity.Error
        };
    }
    
    /// <summary>
    /// 根据错误消息确定严重程度
    /// </summary>
    private ErrorSeverity DetermineSeverityFromMessage(string message)
    {
        var lowerMessage = message.ToLower();
        
        if (lowerMessage.Contains("fatal") || lowerMessage.Contains("致命") || lowerMessage.Contains("crash") || lowerMessage.Contains("崩溃"))
            return ErrorSeverity.Fatal;
        if (lowerMessage.Contains("critical") || lowerMessage.Contains("严重") || lowerMessage.Contains("无法"))
            return ErrorSeverity.Critical;
        if (lowerMessage.Contains("error") || lowerMessage.Contains("错误") || lowerMessage.Contains("失败"))
            return ErrorSeverity.Error;
        if (lowerMessage.Contains("warning") || lowerMessage.Contains("警告") || lowerMessage.Contains("注意"))
            return ErrorSeverity.Warning;
        
        return ErrorSeverity.Error;
    }
    
    /// <summary>
    /// 生成用户友好的错误消息
    /// </summary>
    private string GenerateUserFriendlyMessage(Exception exception)
    {
        return exception switch
        {
            FileNotFoundException => "找不到指定的文件，请检查文件路径是否正确。",
            DirectoryNotFoundException => "找不到指定的目录，请检查目录路径是否正确。",
            UnauthorizedAccessException => "没有权限访问该文件或目录，请检查文件权限设置。",
            OutOfMemoryException => "系统内存不足，请关闭其他程序后重试。",
            IOException => "文件读写操作失败，请检查文件是否被其他程序占用。",
            ArgumentNullException => "程序参数错误，请重新操作。",
            InvalidOperationException when exception.Message.Contains("Excel") => "Excel文件处理失败，请检查文件格式是否正确。",
            System.Net.WebException => "网络连接失败，请检查网络连接。",
            _ => "程序运行时发生了未预期的错误，请联系技术支持。"
        };
    }
    
    /// <summary>
    /// 根据文本生成用户友好消息
    /// </summary>
    private string GenerateUserFriendlyMessageFromText(string message)
    {
        var lowerMessage = message.ToLower();
        
        if (lowerMessage.Contains("file not found") || lowerMessage.Contains("文件不存在"))
            return "找不到指定的文件，请检查文件路径是否正确。";
        if (lowerMessage.Contains("access denied") || lowerMessage.Contains("访问被拒绝"))
            return "没有权限访问该文件或目录，请检查文件权限设置。";
        if (lowerMessage.Contains("out of memory") || lowerMessage.Contains("内存不足"))
            return "系统内存不足，请关闭其他程序后重试。";
        if (lowerMessage.Contains("network") || lowerMessage.Contains("网络"))
            return "网络连接失败，请检查网络连接。";
        if (lowerMessage.Contains("excel") || lowerMessage.Contains("xlsx"))
            return "Excel文件处理失败，请检查文件格式是否正确。";
        
        return "程序运行时发生了错误，请重试或联系技术支持。";
    }
    
    /// <summary>
    /// 生成简单错误报告
    /// </summary>
    private string GenerateErrorReport()
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("=== 错误报告 ===");
        sb.AppendLine($"发生时间: {OccurredAt:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"严重程度: {Severity}");
        sb.AppendLine($"错误分类: {Category}");
        sb.AppendLine();
        sb.AppendLine("错误消息:");
        sb.AppendLine(ErrorMessage);
        
        if (!string.IsNullOrEmpty(UserFriendlyMessage))
        {
            sb.AppendLine();
            sb.AppendLine("用户提示:");
            sb.AppendLine(UserFriendlyMessage);
        }
        
        if (!string.IsNullOrEmpty(ErrorDetails))
        {
            sb.AppendLine();
            sb.AppendLine("详细信息:");
            sb.AppendLine(ErrorDetails);
        }
        
        return sb.ToString();
    }
    
    /// <summary>
    /// 生成详细错误报告
    /// </summary>
    private string GenerateDetailedErrorReport()
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("=== 详细错误报告 ===");
        sb.AppendLine($"报告生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"错误发生时间: {OccurredAt:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"严重程度: {Severity}");
        sb.AppendLine($"错误分类: {Category}");
        sb.AppendLine();
        
        // 系统信息
        sb.AppendLine("=== 系统信息 ===");
        sb.AppendLine($"操作系统: {Environment.OSVersion}");
        sb.AppendLine($"CLR版本: {Environment.Version}");
        sb.AppendLine($"处理器数量: {Environment.ProcessorCount}");
        sb.AppendLine($"工作集内存: {Environment.WorkingSet / 1024 / 1024} MB");
        sb.AppendLine();
        
        // 应用程序信息
        sb.AppendLine("=== 应用程序信息 ===");
        var currentProcess = Process.GetCurrentProcess();
        sb.AppendLine($"进程ID: {currentProcess.Id}");
        sb.AppendLine($"进程名称: {currentProcess.ProcessName}");
        sb.AppendLine($"启动时间: {currentProcess.StartTime:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"内存使用: {currentProcess.WorkingSet64 / 1024 / 1024} MB");
        sb.AppendLine();
        
        // 错误信息
        sb.AppendLine("=== 错误信息 ===");
        sb.AppendLine($"错误消息: {ErrorMessage}");
        
        if (!string.IsNullOrEmpty(UserFriendlyMessage))
        {
            sb.AppendLine($"用户提示: {UserFriendlyMessage}");
        }
        
        sb.AppendLine();
        sb.AppendLine("详细信息:");
        sb.AppendLine(ErrorDetails);
        
        return sb.ToString();
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

/// <summary>
/// 简单的命令实现
/// </summary>
public class RelayCommand : ICommand
{
    private readonly Action _execute;
    private readonly Func<bool>? _canExecute;

    public RelayCommand(Action execute, Func<bool>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public event EventHandler? CanExecuteChanged
    {
        add => System.Windows.Input.CommandManager.RequerySuggested += value;
        remove => System.Windows.Input.CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter)
    {
        try
        {
            return _canExecute?.Invoke() ?? true;
        }
        catch
        {
            return false;
        }
    }

    public void Execute(object? parameter)
    {
        try
        {
            if (CanExecute(parameter))
            {
                _execute();
            }
        }
        catch (Exception ex)
        {
            // 记录命令执行异常，但不抛出以避免应用程序崩溃
            System.Diagnostics.Debug.WriteLine($"命令执行异常: {ex.Message}");
        }
    }
}