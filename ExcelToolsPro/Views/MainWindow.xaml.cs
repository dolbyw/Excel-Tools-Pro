using Microsoft.Extensions.DependencyInjection;
using System.Windows;
using System.Windows.Controls;
using ExcelToolsPro.ViewModels;
using System.IO;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using System.Threading.Tasks;

namespace ExcelToolsPro.Views;

/// <summary>
/// MainWindow.xaml 的交互逻辑
/// </summary>
public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel;
    private readonly ILogger<MainWindow> _logger;
    private bool _isDisposed = false;
    private readonly object _disposeLock = new();

    public MainWindow(MainWindowViewModel viewModel, ILogger<MainWindow> logger)
    {
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            
            _logger.LogDebug("开始初始化主窗口组件...");
            InitializeComponent();
            DataContext = _viewModel;
            
            // 注册事件处理器
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;
            
            // 添加全局异常处理
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            
            _logger.LogInformation("主窗口初始化完成，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "主窗口初始化失败，错误类型: {ExceptionType}, 耗时: {ElapsedMs}ms", 
                ex.GetType().Name, stopwatch.ElapsedMilliseconds);
            
            // 显示用户友好的错误消息
            System.Windows.MessageBox.Show(
                $"应用程序初始化失败：{ex.Message}\n\n请重新启动应用程序。如果问题持续存在，请联系技术支持。", 
                "初始化错误", 
                System.Windows.MessageBoxButton.OK, 
                System.Windows.MessageBoxImage.Error);
            
            throw;
        }
    }

    private async void MainWindow_Loaded(object? sender, RoutedEventArgs e)
    {
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            _logger.LogInformation("主窗口已加载，开始初始化ViewModel...");
            
            // 初始化ViewModel
            await _viewModel.InitializeViewModelAsync();
            
            _logger.LogInformation("ViewModel初始化完成，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
        }
        catch (OperationCanceledException)
        {
            _logger.LogError("ViewModel初始化超时（30秒），耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
            
            var result = System.Windows.MessageBox.Show(
                "应用程序初始化超时。这可能是由于系统资源不足或配置文件损坏导致的。\n\n是否要重置配置并重试？", 
                "初始化超时", 
                System.Windows.MessageBoxButton.YesNo, 
                System.Windows.MessageBoxImage.Warning);
            
            if (result == System.Windows.MessageBoxResult.Yes)
            {
                await TryRecoverFromInitializationError();
            }
            else
            {
                Close();
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "初始化ViewModel时发生错误，错误类型: {ExceptionType}, 耗时: {ElapsedMs}ms", 
                ex.GetType().Name, stopwatch.ElapsedMilliseconds);
            
            await HandleInitializationError(ex);
        }
    }

    private async void OnFileDrop(object? sender, System.Windows.DragEventArgs e)
    {
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            if (_isDisposed)
            {
                _logger.LogWarning("窗口已释放，忽略拖拽操作");
                return;
            }
            
            if (!e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                _logger.LogDebug("拖拽数据不包含文件格式");
                return;
            }
            
            if (e.Data.GetData(System.Windows.DataFormats.FileDrop) is not string[] items || items.Length == 0)
            {
                _logger.LogWarning("拖拽的文件列表为空");
                return;
            }
            
            _logger.LogDebug("开始处理拖拽的 {ItemCount} 个项目...", items.Length);
            
            // 使用超时机制防止拖拽操作卡死
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
            
            var processedCount = 0;
            var errorCount = 0;
            
            if (_viewModel.IsMergeMode)
            {
                // 合并模式：支持文件夹和文件拖拽
                var folders = items.Where(Directory.Exists).ToArray();
                var files = items.Where(File.Exists).ToArray();
                
                _logger.LogDebug("合并模式 - 文件夹: {FolderCount}, 文件: {FileCount}", folders.Length, files.Length);
                
                // 处理文件夹
                foreach (var folder in folders)
                {
                    try
                    {
                        cts.Token.ThrowIfCancellationRequested();
                        await _viewModel.AddFolderFilesAsync(folder);
                        processedCount++;
                    }
                    catch (OperationCanceledException)
                    {
                        _logger.LogWarning("文件夹处理被取消: {Folder}", folder);
                        throw;
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        _logger.LogError(ex, "处理文件夹时发生错误: {Folder}, 错误类型: {ExceptionType}", 
                            folder, ex.GetType().Name);
                    }
                }
                
                // 处理直接拖拽的Excel文件
                if (files.Length > 0)
                {
                    var excelFiles = files.Where(file => 
                        Path.GetExtension(file).ToLower() is ".xlsx" or ".xls" or ".csv")
                        .ToArray();
                    
                    if (excelFiles.Length > 0)
                    {
                        try
                        {
                            cts.Token.ThrowIfCancellationRequested();
                            await _viewModel.AddFilesAsync(excelFiles);
                            processedCount += excelFiles.Length;
                        }
                        catch (OperationCanceledException)
                        {
                            _logger.LogWarning("文件处理被取消");
                            throw;
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            _logger.LogError(ex, "处理Excel文件时发生错误，错误类型: {ExceptionType}", ex.GetType().Name);
                        }
                    }
                    
                    var nonExcelFiles = files.Length - excelFiles.Length;
                    if (nonExcelFiles > 0)
                    {
                        _logger.LogInformation("跳过了 {NonExcelCount} 个非Excel文件", nonExcelFiles);
                    }
                }
                
                _logger.LogInformation("合并模式拖拽处理完成 - 文件夹: {FolderCount}, 文件: {FileCount}, 成功: {ProcessedCount}, 错误: {ErrorCount}, 耗时: {ElapsedMs}ms", 
                    folders.Length, files.Length, processedCount, errorCount, stopwatch.ElapsedMilliseconds);
            }
            else
            {
                // 拆分模式：只支持文件拖拽
                var files = items.Where(File.Exists).ToArray();
                var excelFiles = files.Where(file => 
                    Path.GetExtension(file).ToLower() is ".xlsx" or ".xls" or ".csv")
                    .ToArray();
                
                _logger.LogDebug("拆分模式 - Excel文件: {ExcelFileCount}, 总文件: {TotalFileCount}", excelFiles.Length, files.Length);
                
                if (excelFiles.Length > 0)
                {
                    try
                    {
                        cts.Token.ThrowIfCancellationRequested();
                        await _viewModel.AddFilesAsync(excelFiles);
                        processedCount = excelFiles.Length;
                        
                        _logger.LogInformation("拆分模式拖拽处理完成 - 添加了 {ExcelFileCount} 个Excel文件，耗时: {ElapsedMs}ms", 
                            excelFiles.Length, stopwatch.ElapsedMilliseconds);
                    }
                    catch (OperationCanceledException)
                    {
                        _logger.LogWarning("拆分模式文件处理被取消");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        _logger.LogError(ex, "拆分模式处理Excel文件时发生错误，错误类型: {ExceptionType}", ex.GetType().Name);
                    }
                }
                else
                {
                    var nonExcelCount = items.Length - excelFiles.Length;
                    _logger.LogInformation("拆分模式 - 跳过了 {NonExcelCount} 个非Excel文件", nonExcelCount);
                    
                    System.Windows.MessageBox.Show(
                        $"拆分模式下只能处理Excel文件（.xlsx, .xls, .csv）。\n\n已跳过 {nonExcelCount} 个不支持的文件。", 
                        "文件类型提示", 
                        System.Windows.MessageBoxButton.OK, 
                        System.Windows.MessageBoxImage.Information);
                }
            }
            
            // 显示处理结果摘要
            if (processedCount > 0 || errorCount > 0)
            {
                var message = $"拖拽处理完成：\n成功处理 {processedCount} 个项目";
                if (errorCount > 0)
                {
                    message += $"\n处理失败 {errorCount} 个项目";
                }
                message += $"\n耗时 {stopwatch.ElapsedMilliseconds} 毫秒";
                
                _viewModel.StatusText = message;
            }
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("拖拽操作被取消，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
            _viewModel.StatusText = "拖拽操作已取消";
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "处理拖拽文件时发生严重错误，错误类型: {ExceptionType}, 耗时: {ElapsedMs}ms", 
                ex.GetType().Name, stopwatch.ElapsedMilliseconds);
            
            var errorMessage = GetUserFriendlyErrorMessage(ex);
            System.Windows.MessageBox.Show(
                $"处理拖拽文件时发生错误：\n{errorMessage}\n\n如果问题持续存在，请重新启动应用程序。", 
                "拖拽处理错误", 
                System.Windows.MessageBoxButton.OK, 
                System.Windows.MessageBoxImage.Error);
            
            _viewModel.StatusText = "拖拽处理失败";
        }
    }

    private void OnDragOver(object sender, System.Windows.DragEventArgs e)
    {
        try
        {
            if (_isDisposed)
            {
                e.Effects = System.Windows.DragDropEffects.None;
                e.Handled = true;
                return;
            }
            
            if (!e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                e.Effects = System.Windows.DragDropEffects.None;
                e.Handled = true;
                return;
            }
            
            if (e.Data.GetData(System.Windows.DataFormats.FileDrop) is not string[] items)
            {
                e.Effects = System.Windows.DragDropEffects.None;
                e.Handled = true;
                return;
            }
            
            bool canDrop = false;
            
            if (items.Length > 0)
            {
                try
                {
                    if (_viewModel.IsMergeMode)
                    {
                        // 合并模式：接受文件夹或Excel文件
                        var hasFolders = items.Any(item => 
                        {
                            try { return Directory.Exists(item); }
                            catch { return false; }
                        });
                        
                        var hasExcelFiles = items.Where(item => 
                        {
                            try { return File.Exists(item); }
                            catch { return false; }
                        }).Any(file => 
                        {
                            try 
                            { 
                                var ext = Path.GetExtension(file).ToLower();
                                return ext is ".xlsx" or ".xls" or ".csv";
                            }
                            catch { return false; }
                        });
                        
                        canDrop = hasFolders || hasExcelFiles;
                    }
                    else
                    {
                        // 拆分模式：只接受Excel文件
                        var hasExcelFiles = items.Where(item => 
                        {
                            try { return File.Exists(item); }
                            catch { return false; }
                        }).Any(file => 
                        {
                            try 
                            { 
                                var ext = Path.GetExtension(file).ToLower();
                                return ext is ".xlsx" or ".xls" or ".csv";
                            }
                            catch { return false; }
                        });
                        
                        canDrop = hasExcelFiles;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "检查拖拽项目时发生错误，拒绝拖拽操作");
                    canDrop = false;
                }
            }
            
            e.Effects = canDrop ? System.Windows.DragDropEffects.Copy : System.Windows.DragDropEffects.None;
            e.Handled = true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "处理拖拽悬停时发生严重错误，错误类型: {ExceptionType}", ex.GetType().Name);
            e.Effects = System.Windows.DragDropEffects.None;
            e.Handled = true;
        }
    }

    private async void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            _logger.LogInformation("主窗口开始关闭流程...");
            
            // 检查是否有正在进行的操作
            if (_viewModel?.IsProcessing == true)
            {
                var result = System.Windows.MessageBox.Show(
                    "当前有文件正在处理中。强制关闭可能导致数据丢失。\n\n是否要取消当前操作并关闭应用程序？", 
                    "确认关闭", 
                    System.Windows.MessageBoxButton.YesNo, 
                    System.Windows.MessageBoxImage.Warning);
                
                if (result == System.Windows.MessageBoxResult.No)
                {
                    e.Cancel = true;
                    _logger.LogInformation("用户取消了窗口关闭操作");
                    return;
                }
                
                // 尝试优雅地取消当前操作
                try
                {
                    _logger.LogInformation("尝试取消当前处理操作...");
                    _viewModel.CancelProcessingCommand?.Execute(null);
                    
                    // 等待操作取消完成，最多等待5秒
                    var cancellationTimeout = TimeSpan.FromSeconds(5);
                    var cancellationStart = DateTime.UtcNow;
                    
                    while (_viewModel.IsProcessing && (DateTime.UtcNow - cancellationStart) < cancellationTimeout)
                    {
                        await Task.Delay(100);
                    }
                    
                    if (_viewModel.IsProcessing)
                    {
                        _logger.LogWarning("操作取消超时，强制关闭");
                    }
                    else
                    {
                        _logger.LogInformation("当前操作已成功取消");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "取消当前操作时发生错误，错误类型: {ExceptionType}", ex.GetType().Name);
                }
            }
            
            _logger.LogInformation("窗口关闭检查完成，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "窗口关闭检查时发生错误，错误类型: {ExceptionType}, 耗时: {ElapsedMs}ms", 
                ex.GetType().Name, stopwatch.ElapsedMilliseconds);
            
            // 即使检查失败，也允许窗口关闭
        }
    }
    
    protected override void OnClosed(EventArgs e)
    {
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            lock (_disposeLock)
            {
                if (_isDisposed)
                {
                    _logger.LogDebug("窗口已经释放，跳过重复释放");
                    return;
                }
                
                _isDisposed = true;
            }
            
            _logger.LogInformation("开始释放主窗口资源...");
            
            // 注销事件处理器
            try
            {
                Loaded -= MainWindow_Loaded;
                Closing -= MainWindow_Closing;
                AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
                _logger.LogDebug("事件处理器注销完成");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "注销事件处理器时发生错误");
            }
            
            // 释放ViewModel资源
            try
            {
                _viewModel?.Dispose();
                _logger.LogDebug("ViewModel资源释放完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "释放ViewModel资源时发生错误，错误类型: {ExceptionType}", ex.GetType().Name);
            }
            
            // 执行垃圾回收
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                _logger.LogDebug("垃圾回收完成");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "执行垃圾回收时发生错误");
            }
            
            _logger.LogInformation("主窗口资源释放完成，总耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            // 使用Console.WriteLine作为后备日志记录，因为logger可能已经被释放
            Console.WriteLine($"关闭主窗口时发生严重错误: {ex}");
            
            try
            {
                _logger?.LogCritical(ex, "关闭主窗口时发生严重错误，错误类型: {ExceptionType}, 耗时: {ElapsedMs}ms", 
                    ex.GetType().Name, stopwatch.ElapsedMilliseconds);
            }
            catch
            {
                // 忽略日志记录错误
            }
        }
        finally
        {
            try
            {
                base.OnClosed(e);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"调用基类OnClosed时发生错误: {ex}");
            }
        }
    }
    
    /// <summary>
    /// 处理未捕获的异常
    /// </summary>
    private void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        try
        {
            var exception = e.ExceptionObject as Exception;
            _logger.LogCritical(exception, "主窗口捕获到未处理的异常，即将终止: {IsTerminating}", e.IsTerminating);
            
            if (exception != null && !e.IsTerminating)
            {
                var errorMessage = GetUserFriendlyErrorMessage(exception);
                
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    System.Windows.MessageBox.Show(
                        $"应用程序遇到未处理的错误：\n{errorMessage}\n\n应用程序将尝试继续运行，但建议您保存工作并重新启动。", 
                        "未处理的错误", 
                        System.Windows.MessageBoxButton.OK, 
                        System.Windows.MessageBoxImage.Error);
                }));
            }
        }
        catch
        {
            // 在异常处理器中不能抛出异常
        }
    }
    
    /// <summary>
    /// 尝试从初始化错误中恢复
    /// </summary>
    private async Task TryRecoverFromInitializationError()
    {
        try
        {
            _logger.LogInformation("尝试从初始化错误中恢复...");
            
            // 重置配置并重新初始化
            // 这里可以添加具体的恢复逻辑
            
            await _viewModel.InitializeViewModelAsync();
            
            _logger.LogInformation("初始化错误恢复成功");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "初始化错误恢复失败，错误类型: {ExceptionType}", ex.GetType().Name);
            
            System.Windows.MessageBox.Show(
                "恢复失败。应用程序将关闭。\n\n请检查系统资源和配置文件，然后重新启动应用程序。", 
                "恢复失败", 
                System.Windows.MessageBoxButton.OK, 
                System.Windows.MessageBoxImage.Error);
            
            Close();
        }
    }
    
    /// <summary>
    /// 处理初始化错误
    /// </summary>
    private async Task HandleInitializationError(Exception ex)
    {
        var errorMessage = GetUserFriendlyErrorMessage(ex);
        
        var result = System.Windows.MessageBox.Show(
            $"应用程序初始化失败：\n{errorMessage}\n\n是否要尝试重置配置并重新初始化？", 
            "初始化错误", 
            System.Windows.MessageBoxButton.YesNo, 
            System.Windows.MessageBoxImage.Error);
        
        if (result == System.Windows.MessageBoxResult.Yes)
        {
            await TryRecoverFromInitializationError();
        }
        else
        {
            Close();
        }
    }
    
    /// <summary>
    /// 获取用户友好的错误消息
    /// </summary>
    private static string GetUserFriendlyErrorMessage(Exception ex)
    {
        return ex switch
        {
            UnauthorizedAccessException => "访问被拒绝。请检查文件权限或以管理员身份运行。",
            DirectoryNotFoundException => "找不到指定的目录。请检查路径是否正确。",
            FileNotFoundException => "找不到指定的文件。请检查文件是否存在。",
            IOException => "文件操作失败。请检查文件是否被其他程序占用。",
            OutOfMemoryException => "内存不足。请关闭其他应用程序或重新启动。",
            TimeoutException => "操作超时。请检查网络连接或稍后重试。",
            ArgumentException => "参数错误。请检查输入的数据是否正确。",
            InvalidOperationException => "当前操作无效。请检查应用程序状态。",
            _ => ex.Message
        };
    }
}