using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using Serilog.Context;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using System.Threading.Tasks;
using ExcelToolsPro.Services;
using ExcelToolsPro.Services.FileNaming.Core;
using ExcelToolsPro.Services.FileNaming.Split;
using ExcelToolsPro.ViewModels;
using ExcelToolsPro.Views;

namespace ExcelToolsPro;

/// <summary>
/// App.xaml 的交互逻辑
/// </summary>
public partial class App : System.Windows.Application
{
    public static ServiceProvider? ServiceProvider { get; private set; }
    private ILogger<App>? _logger;

    protected override void OnStartup(StartupEventArgs e)
    {
        var stopwatch = Stopwatch.StartNew();
        
        // 设置全局异常处理器
        SetupGlobalExceptionHandlers();
        
        try // Top-level catch-all
        {
            // Initial logging setup
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .Enrich.FromLogContext()
                .Enrich.WithThreadId()
                .WriteTo.Console(outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] [{ThreadId}] {SourceContext} - {Message:lj}{NewLine}{Exception}")
                .WriteTo.File("logs/startup-.log",
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] [{ThreadId}] {SourceContext} - {Message:lj}{NewLine}{Exception}",
                    encoding: System.Text.Encoding.UTF8)
                .CreateLogger();

        using (LogContext.PushProperty("SourceContext", "App"))
        {
            Log.Debug("=== Excel Tools Pro 应用程序启动开始 ===");
            Log.Debug("启动参数: {Args}, 进程ID: {ProcessId}, 线程ID: {ThreadId}",
                e.Args, Environment.ProcessId, Environment.CurrentManagedThreadId);
            Log.Debug("工作目录: {WorkingDirectory}, 基础目录: {BaseDirectory}",
                Environment.CurrentDirectory, AppDomain.CurrentDomain.BaseDirectory);
        }

        base.OnStartup(e);

        // Main application setup logic
        using (LogContext.PushProperty("SourceContext", "App"))
        {
            Log.Debug("开始配置服务容器...");
            var serviceConfigStart = stopwatch.ElapsedMilliseconds;
            ConfigureServices();
            Log.Debug("服务容器配置完成，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds - serviceConfigStart);

            Log.Debug("开始配置日志系统...");
            var loggingConfigStart = stopwatch.ElapsedMilliseconds;
            ConfigureLogging();
            Log.Debug("日志系统配置完成，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds - loggingConfigStart);

            _logger = ServiceProvider?.GetService<ILogger<App>>();
            _logger?.LogInformation("应用程序服务配置完成，开始创建主窗口，总耗时: {TotalElapsedMs}ms", stopwatch.ElapsedMilliseconds);
            _logger?.LogDebug("服务提供者状态: {ServiceProviderStatus}, 日志记录器状态: {LoggerStatus}",
                ServiceProvider != null ? "已初始化" : "未初始化",
                _logger != null ? "已初始化" : "未初始化");

            // Create and show the main window
            _logger?.LogDebug("开始创建主窗口...");
            var mainWindowStart = stopwatch.ElapsedMilliseconds;
            var mainWindow = ServiceProvider?.GetService<MainWindow>();
            if (mainWindow != null)
            {
                _logger?.LogDebug("主窗口创建成功，设置为主窗口并显示...");
                MainWindow = mainWindow;
                mainWindow.Loaded += (s, ev) => _logger?.LogInformation("主窗口已加载完成");
                mainWindow.Activated += (s, ev) => _logger?.LogDebug("主窗口已激活");
                mainWindow.Deactivated += (s, ev) => _logger?.LogDebug("主窗口已失去焦点");
                mainWindow.Closed += (s, ev) => _logger?.LogInformation("主窗口已关闭");
                mainWindow.Show();
                mainWindow.Activate();
                _logger?.LogInformation("主窗口显示完成，窗口创建耗时: {WindowElapsedMs}ms",
                    stopwatch.ElapsedMilliseconds - mainWindowStart);
                _logger?.LogDebug("主窗口状态 - 可见: {IsVisible}, 已加载: {IsLoaded}, 窗口状态: {WindowState}",
                    mainWindow.IsVisible, mainWindow.IsLoaded, mainWindow.WindowState);
            }
            else
            {
                _logger?.LogError("主窗口创建失败：服务容器返回null");
                throw new InvalidOperationException("无法创建主窗口实例");
            }

            _logger?.LogInformation("=== 应用程序启动流程完成，总耗时: {TotalMs}ms ===", stopwatch.ElapsedMilliseconds);
        }
        stopwatch.Stop();
    }
    catch (Exception ex)
    {
        HandleStartupException(ex);
    }
    }

    private static void ConfigureServices()
    {
        using (LogContext.PushProperty("SourceContext", "App.ConfigureServices"))
        {
            Log.Debug("开始初始化服务集合...");
            var services = new ServiceCollection();
            Log.Debug("服务集合创建完成");

            // 配置
            Log.Debug("开始构建配置...");
            var configBuilder = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory);
            Log.Debug("配置构建器基础路径设置为: {BasePath}", AppDomain.CurrentDomain.BaseDirectory);
            
            var configFile = "appsettings.json";
            var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, configFile);
            Log.Debug("检查配置文件: {ConfigPath}, 存在: {FileExists}", configPath, File.Exists(configPath));
            
            var configuration = configBuilder
                .AddJsonFile(configFile, optional: false, reloadOnChange: true)
                .Build();
            Log.Debug("配置构建完成，配置节数量: {SectionCount}", configuration.GetChildren().Count());

            services.AddSingleton<IConfiguration>(configuration);
            Log.Debug("配置服务注册完成");

            // 日志
            Log.Debug("开始配置日志服务...");
            services.AddLogging(builder =>
            {
                builder.ClearProviders();
                builder.AddSerilog();
                Log.Debug("Serilog提供程序已添加到日志构建器");
            });
            Log.Debug("日志服务配置完成");

            // 业务服务
            Log.Debug("开始注册业务服务...");
            services.AddSingleton<IExcelProcessingService, ExcelProcessingService>();
            Log.Debug("Excel处理服务已注册");
            
            services.AddSingleton<IFileSystemService, FileSystemService>();
            Log.Debug("文件系统服务已注册");
            
            services.AddSingleton<IConfigurationService, ConfigurationService>();
            Log.Debug("配置管理服务已注册");
            
            services.AddSingleton<IErrorRecoveryService, ErrorRecoveryService>();
            Log.Debug("错误恢复服务已注册");
            
            services.AddSingleton<IPerformanceMonitorService, PerformanceMonitorService>();
            Log.Debug("性能监控服务已注册");
            
            services.AddSingleton<ILowMemoryProcessor, LowMemoryProcessor>();
            Log.Debug("低内存处理器服务已注册");
            
            // FileNaming服务
            services.AddSingleton<IVariableRegistry, VariableRegistry>();
            Log.Debug("变量注册表已注册");
            
            services.AddSingleton<SplitNamingEngine>();
            Log.Debug("命名引擎已注册");
            
            services.AddSingleton<ISplitNamingService, SplitNamingService>();
            Log.Debug("文件命名服务已注册");

            // ViewModels
            Log.Debug("开始注册视图模型...");
            services.AddTransient<MainWindowViewModel>();
            services.AddTransient<HomeViewModel>();
            services.AddTransient<ProgressViewModel>();
            services.AddTransient<SettingsViewModel>();
            services.AddTransient<ErrorViewModel>();
            Log.Debug("所有视图模型注册完成");

            // Views
            Log.Debug("开始注册视图...");
            services.AddTransient<MainWindow>();
            Log.Debug("主窗口视图注册完成");

            Log.Debug("开始构建服务提供者...");
            ServiceProvider = services.BuildServiceProvider();
            Log.Debug("服务提供者构建完成，注册服务数量: {ServiceCount}", services.Count);
            
            // 验证关键服务
            Log.Debug("开始验证关键服务注册...");
            var criticalServices = new[]
            {
                typeof(IConfiguration),
                typeof(ILogger<App>),
                typeof(IExcelProcessingService),
                typeof(MainWindow)
            };
            
            foreach (var serviceType in criticalServices)
            {
                var service = ServiceProvider.GetService(serviceType);
                Log.Debug("服务验证 - {ServiceType}: {Status}", 
                    serviceType.Name, service != null ? "成功" : "失败");
            }
        }
    }

    private static void ConfigureLogging()
    {
        using (LogContext.PushProperty("SourceContext", "App.ConfigureLogging"))
        {
            Log.Debug("开始配置Serilog日志系统...");
            
            var configuration = ServiceProvider?.GetService<IConfiguration>();
            if (configuration == null)
            {
                Log.Warning("无法获取配置服务，使用默认日志配置");
            }
            else
            {
                Log.Debug("配置服务获取成功，开始读取日志配置...");
            }

            var logConfig = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .Enrich.FromLogContext()
                .Enrich.WithThreadId();

            // 从配置文件读取日志设置
            if (configuration != null)
            {
                try
                {
                    logConfig.ReadFrom.Configuration(configuration);
                    Log.Debug("从配置文件读取日志设置成功");
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "从配置文件读取日志设置失败，使用默认配置");
                    // 使用默认配置
                    logConfig
                        .WriteTo.Console(outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] [{ThreadId}] {SourceContext} - {Message:lj}{NewLine}{Exception}")
                        .WriteTo.File("logs/app-.log", 
                            rollingInterval: RollingInterval.Day,
                            outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] [{ThreadId}] {SourceContext} - {Message:lj}{NewLine}{Exception}");
                }
            }
            else
            {
                // 使用默认配置
                logConfig
                    .WriteTo.Console(outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] [{ThreadId}] {SourceContext} - {Message:lj}{NewLine}{Exception}")
                    .WriteTo.File("logs/app-.log", 
                        rollingInterval: RollingInterval.Day,
                        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] [{ThreadId}] {SourceContext} - {Message:lj}{NewLine}{Exception}");
            }

            Log.Logger = logConfig.CreateLogger();
            Log.Debug("Serilog日志系统配置完成");
            
            // 测试日志输出
            Log.Debug("日志系统测试 - Debug级别");
            Log.Information("日志系统测试 - Information级别");
            Log.Warning("日志系统测试 - Warning级别");
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            using (LogContext.PushProperty("SourceContext", "App.OnExit"))
            {
                _logger?.LogInformation("=== 应用程序开始退出流程 ===");
                _logger?.LogDebug("退出代码: {ExitCode}, 进程ID: {ProcessId}", e.ApplicationExitCode, Environment.ProcessId);
                
                // 清理服务提供者中的资源
                if (ServiceProvider != null)
                {
                    _logger?.LogDebug("开始释放服务提供者资源...");
                    var disposeStart = stopwatch.ElapsedMilliseconds;
                    
                    // 显式清理可能持有资源的服务
                    try
                    {
                        var performanceMonitor = ServiceProvider.GetService<IPerformanceMonitorService>();
                        if (performanceMonitor is IDisposable disposableMonitor)
                        {
                            disposableMonitor.Dispose();
                            _logger?.LogDebug("性能监控服务已释放");
                        }
                        
                        var excelService = ServiceProvider.GetService<IExcelProcessingService>();
                        if (excelService is IDisposable disposableExcel)
                        {
                            disposableExcel.Dispose();
                            _logger?.LogDebug("Excel处理服务已释放");
                        }
                    }
                    catch (Exception serviceEx)
                    {
                        _logger?.LogWarning(serviceEx, "清理服务时发生错误");
                    }
                    
                    ServiceProvider.Dispose();
                    ServiceProvider = null;
                    _logger?.LogDebug("服务提供者资源释放完成，耗时: {ElapsedMs}ms", 
                        stopwatch.ElapsedMilliseconds - disposeStart);
                }
                
                // 强制垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                _logger?.LogInformation("应用程序退出流程完成，总耗时: {TotalMs}ms", stopwatch.ElapsedMilliseconds);
                
                // 关闭日志系统
                Log.Information("关闭日志系统...");
                Log.CloseAndFlush();
            }
        }
        catch (Exception ex)
        {
            // 使用系统诊断作为后备日志
            System.Diagnostics.Debug.WriteLine($"应用程序退出时发生错误: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"错误堆栈: {ex.StackTrace}");
            
            // 尝试记录到Serilog
            try
            {
                Log.Fatal(ex, "应用程序退出时发生严重错误");
                Log.CloseAndFlush();
            }
            catch
            {
                // 忽略日志记录错误
            }
        }
        finally
        {
            stopwatch.Stop();
            base.OnExit(e);
            
            // 确保进程完全退出
            Environment.Exit(e.ApplicationExitCode);
        }
    }
    
    /// <summary>
    /// 设置全局异常处理器
    /// </summary>
    private void SetupGlobalExceptionHandlers()
    {
        // 处理UI线程未处理的异常
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        
        // 处理非UI线程未处理的异常
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        
        // 处理Task中未观察到的异常
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
    }
    
    /// <summary>
    /// 处理UI线程未处理的异常
    /// </summary>
    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        try
        {
            _logger?.LogError(e.Exception, "UI线程发生未处理异常: {ExceptionType}", e.Exception.GetType().Name);
            
            var errorMessage = GetUserFriendlyErrorMessage(e.Exception);
            var result = System.Windows.MessageBox.Show(
                $"{errorMessage}\n\n是否继续运行应用程序？\n\n点击'是'继续运行，点击'否'退出应用程序。",
                "应用程序错误",
                MessageBoxButton.YesNo,
                MessageBoxImage.Error);
            
            if (result == MessageBoxResult.Yes)
            {
                e.Handled = true; // 标记异常已处理，继续运行
                _logger?.LogInformation("用户选择继续运行应用程序");
            }
            else
            {
                _logger?.LogInformation("用户选择退出应用程序");
                Shutdown(1);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogCritical(ex, "处理UI线程异常时发生错误");
            Environment.Exit(1);
        }
    }
    
    /// <summary>
    /// 处理非UI线程未处理的异常
    /// </summary>
    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        try
        {
            var exception = e.ExceptionObject as Exception;
            _logger?.LogCritical(exception, "非UI线程发生未处理异常，即将终止: {IsTerminating}", e.IsTerminating);
            
            if (exception != null)
            {
                var errorMessage = GetUserFriendlyErrorMessage(exception);
                
                // 在UI线程显示错误消息
                Dispatcher.Invoke(() =>
                {
                    System.Windows.MessageBox.Show(
                        $"应用程序遇到严重错误，即将退出：\n\n{errorMessage}",
                        "严重错误",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                });
            }
        }
        catch (Exception ex)
        {
            // 最后的防线，记录到系统事件日志
            try
            {
                System.Diagnostics.EventLog.WriteEntry("ExcelToolsPro", 
                    $"Fatal error in exception handler: {ex.Message}", 
                    System.Diagnostics.EventLogEntryType.Error);
            }
            catch
            {
                // 忽略事件日志写入错误
            }
        }
        finally
        {
            if (e.IsTerminating)
            {
                try
                {
                    Log.CloseAndFlush();
                }
                catch
                {
                    // 忽略日志关闭错误
                }
            }
        }
    }
    
    /// <summary>
    /// 处理Task中未观察到的异常
    /// </summary>
    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        try
        {
            _logger?.LogError(e.Exception, "Task中发生未观察异常: {ExceptionCount} 个异常", 
                e.Exception.InnerExceptions.Count);
            
            foreach (var innerException in e.Exception.InnerExceptions)
            {
                _logger?.LogError(innerException, "Task内部异常: {ExceptionType}", 
                    innerException.GetType().Name);
            }
            
            // 标记异常已观察，防止应用程序崩溃
            e.SetObserved();
            
            // 在UI线程显示警告
            Dispatcher.BeginInvoke(() =>
            {
                var errorMessage = GetUserFriendlyErrorMessage(e.Exception.GetBaseException());
                System.Windows.MessageBox.Show(
                    $"后台任务发生错误，但应用程序将继续运行：\n\n{errorMessage}",
                    "后台任务错误",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            });
        }
        catch (Exception ex)
        {
            _logger?.LogCritical(ex, "处理Task异常时发生错误");
        }
    }
    
    /// <summary>
    /// 处理启动异常
    /// </summary>
    private void HandleStartupException(Exception ex)
    {
        var errorMessage = $"应用程序启动失败: {GetUserFriendlyErrorMessage(ex)}";
        _logger?.LogCritical(ex, "OnStartup 顶层捕获到致命错误");
        
        System.Windows.MessageBox.Show(errorMessage, "启动失败", MessageBoxButton.OK, MessageBoxImage.Error);
        
        // Try to log it, but this might fail if logging isn't set up
        try 
        {
            Log.Fatal(ex, "OnStartup 顶层捕获到致命错误");
            Log.CloseAndFlush();
        } 
        catch 
        {
            // Ignore logging errors during shutdown
        }
        
        Environment.Exit(1);
    }
    
    /// <summary>
    /// 获取用户友好的错误消息
    /// </summary>
    private static string GetUserFriendlyErrorMessage(Exception ex)
    {
        return ex switch
        {
            FileNotFoundException => "找不到所需的文件，请检查文件是否存在。",
            DirectoryNotFoundException => "找不到指定的目录，请检查路径是否正确。",
            UnauthorizedAccessException => "没有足够的权限访问文件或目录，请以管理员身份运行。",
            OutOfMemoryException => "系统内存不足，请关闭其他应用程序后重试。",
            InvalidOperationException => "操作无效，请检查当前状态是否正确。",
            ArgumentException => "参数错误，请检查输入的数据是否正确。",
            TimeoutException => "操作超时，请检查网络连接或稍后重试。",
            System.IO.IOException => "文件操作失败，请检查文件是否被其他程序占用。",
            NotSupportedException => "不支持的操作，请检查文件格式是否正确。",
            _ => $"发生了意外错误：{ex.Message}"
        };
    }
    
    protected override void OnSessionEnding(SessionEndingCancelEventArgs e)
    {
        _logger?.LogInformation("系统会话结束，原因: {Reason}", e.ReasonSessionEnding);
        
        // 执行快速清理
        try
        {
            ServiceProvider?.Dispose();
            Log.CloseAndFlush();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"会话结束清理时发生错误: {ex.Message}");
        }
        
        base.OnSessionEnding(e);
    }
}