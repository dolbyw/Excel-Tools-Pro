using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using WinForms = System.Windows.Forms;

using ExcelToolsPro.Models;
using ExcelToolsPro.Services;
using FileInfo = ExcelToolsPro.Models.FileInfo;
using TaskStatus = ExcelToolsPro.Models.TaskStatus;

namespace ExcelToolsPro.ViewModels;

/// <summary>
/// 主窗口视图模型
/// </summary>
public partial class MainWindowViewModel : ObservableObject, IDisposable
{
    private readonly IExcelProcessingService _excelService;
    private readonly IConfigurationService _configService;
    private readonly IPerformanceMonitorService _performanceService;
    private readonly ILogger<MainWindowViewModel> _logger;
    private ProcessingTask? _currentTask;
    private bool _disposed = false;
    
    [ObservableProperty]
    private string _uiResponseStatus = "响应正常";

    // 默认输出路径常量
    private const string DefaultMergeOutputPath = @"C:\Users\Administrator\Documents\ToooOutput\Merge";
    private const string DefaultSplitOutputPath = @"C:\Users\Administrator\Documents\ToooOutput\Split";

    [ObservableProperty]
    private ObservableCollection<FileInfo> _selectedFiles = [];

    [ObservableProperty]
    private bool _isMergeMode = true;

    [ObservableProperty]
    private bool _addHeaders = true;

    [ObservableProperty]
    private bool _dedupeHeaders = true;

    [ObservableProperty]
    private string _outputPath = string.Empty;

    [ObservableProperty]
    private bool _isProcessing = false;

    [ObservableProperty]
    private float _progressValue = 0f;

    [ObservableProperty]
    private string _progressText = "准备就绪";

    [ObservableProperty]
    private string _statusText = "准备就绪";

    [ObservableProperty]
    private string _fileCountText = "0 个文件";

    [ObservableProperty]
    private string _memoryUsageText = "0 MB";

    [ObservableProperty]
    private int _customSplitRowCount = 1000;

    [ObservableProperty]
    private bool _autoAddHeader = true;

    [ObservableProperty]
    private int _splitRowCount = 1000;

    [ObservableProperty]
    private string _outputFolder = string.Empty;
    
    [ObservableProperty]
    private string _title = "Excel Tools Pro";
    
    // 缓存支持的文件扩展名数组以避免重复创建
    private static readonly string[] SupportedFileExtensions = [".xlsx", ".xls", ".csv"];
    
    /// <summary>
    /// 开始处理命令
    /// </summary>
    public IRelayCommand StartCommand => StartProcessingCommand;

    /// <summary>
    /// 选择输出文件夹命令
    /// </summary>
    [RelayCommand]
    private void SelectOutputFolder()
    {
        try
        {
            var dialog = new WinForms.FolderBrowserDialog
            {
                Description = "选择输出文件夹",
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                OutputFolder = dialog.SelectedPath;
                OutputPath = dialog.SelectedPath;
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "选择输出文件夹");
        }
    }

    public MainWindowViewModel(
        IExcelProcessingService excelService,
        IConfigurationService configService,
        IPerformanceMonitorService performanceService,
        ILogger<MainWindowViewModel> logger)
    {
        _excelService = excelService;
        _configService = configService;
        _performanceService = performanceService;
        _logger = logger;

        // 订阅文件集合变化事件
        SelectedFiles.CollectionChanged += OnFilesCollectionChanged;
        
        // 设置默认输出路径 - 延迟到异步初始化阶段执行
        // SetDefaultOutputPathAsync(); // 改为在InitializeAsync中调用
        
        // 记录关键属性的初始值用于调试
        _logger.LogDebug("ViewModel初始状态 - IsMergeMode: {IsMergeMode}, HasFiles: {HasFiles}, OutputPath: {OutputPath}", 
            IsMergeMode, HasFiles, OutputPath);
        
        _logger.LogInformation("主窗口视图模型构造完成");
    }

    public async Task InitializeViewModelAsync()
    {
        try
        {
            await InitializeAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogCritical(ex, "在视图模型初始化期间发生致命错误，尝试恢复而不是关闭应用程序。");
            
            // 尝试备用初始化而不是直接关闭应用程序
            try
            {
                await FallbackInitializationAsync().ConfigureAwait(false);
                _logger.LogInformation("使用备用初始化成功恢复应用程序");
            }
            catch (Exception fallbackEx)
            {
                _logger.LogCritical(fallbackEx, "备用初始化也失败了，显示错误但不关闭应用程序");
                
                // 在UI线程上显示错误消息但不关闭应用程序
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    System.Windows.MessageBox.Show($"应用程序初始化失败，但将继续运行: {ex.Message}\n\n部分功能可能不可用。", "初始化错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                });
            }
        }
    }
    
    private async Task InitializeAsync()
    {
        var initializationTimeout = TimeSpan.FromSeconds(30);
        using var cts = new CancellationTokenSource(initializationTimeout);
        
        try
        {
            _logger.LogDebug("开始异步初始化主窗口视图模型");
            
            // 并行执行初始化任务以提高效率
            var tasks = new List<Task>
            {
                InitializePerformanceMonitoringAsync(cts.Token),
                LoadConfigurationAsync(cts.Token)
            };
            
            await Task.WhenAll(tasks).ConfigureAwait(false);
            
            _logger.LogInformation("主窗口视图模型异步初始化完成");
        }
        catch (OperationCanceledException) when (cts.Token.IsCancellationRequested)
        {
            _logger.LogWarning("主窗口视图模型初始化超时，使用默认配置");
            await FallbackInitializationAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "主窗口视图模型异步初始化失败，尝试恢复");
            await FallbackInitializationAsync().ConfigureAwait(false);
        }
    }
    
    /// <summary>
    /// 初始化性能监控 - 带超时和错误恢复
    /// </summary>
    private async Task InitializePerformanceMonitoringAsync(CancellationToken cancellationToken)
    {
        try
        {
            // 延迟订阅性能监控事件，避免启动时的性能开销
            await Task.Delay(2000, cancellationToken).ConfigureAwait(false);
            
            if (!cancellationToken.IsCancellationRequested)
            {
                // 使用异步方式订阅事件，避免UI线程阻塞
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    _performanceService.MetricsUpdated += OnMetricsUpdated;
                    _performanceService.StartMonitoring();
                }, System.Windows.Threading.DispatcherPriority.Background, cancellationToken);
                
                _logger.LogDebug("性能监控初始化完成");
            }
        }
        catch (OperationCanceledException)
        {
            _logger.LogDebug("性能监控初始化被取消");
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "性能监控初始化失败，将跳过性能监控功能");
            // 不重新抛出异常，允许应用继续运行
        }
    }
    
    /// <summary>
    /// 备用初始化 - 当主初始化失败时使用
    /// </summary>
    private async Task FallbackInitializationAsync()
    {
        try
        {
            _logger.LogInformation("执行备用初始化流程");
            
            // 设置默认配置 - 使用异步方式避免UI线程阻塞
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
            {
                AddHeaders = true;
                DedupeHeaders = true;
                AutoAddHeader = true;
                SplitRowCount = 1000;
                CustomSplitRowCount = 1000;
                StatusText = "初始化完成（使用默认配置）";
            });
            
            // 异步设置默认输出路径
            await SetDefaultOutputPathAsync();
            
            _logger.LogInformation("备用初始化完成");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "备用初始化也失败了");
            // 最后的保护措施 - 使用异步方式
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
            {
                StatusText = "初始化失败，请重启应用程序";
            });
        }
    }

    /// <summary>
    /// 是否为拆分模式
    /// </summary>
    public bool IsSplitMode
    {
        get => !IsMergeMode;
        set => IsMergeMode = !value;
    }

    /// <summary>
    /// 设置默认输出路径
    /// </summary>
    private async Task SetDefaultOutputPathAsync()
    {
        try
        {
            if (IsMergeMode)
            {
                // 合并模式：设置默认目录路径（不包含文件名）
                var defaultDir = DefaultMergeOutputPath;
                if (!Directory.Exists(defaultDir))
                {
                    await Task.Run(() => Directory.CreateDirectory(defaultDir));
                    _logger.LogInformation("创建默认输出目录: {Path}", defaultDir);
                }
                
                // 合并模式只显示目录路径，不包含文件名
                OutputPath = defaultDir;
                OutputFolder = defaultDir;
            }
            else
            {
                // 拆分模式：设置默认文件夹路径
                var defaultPath = DefaultSplitOutputPath;
                if (!Directory.Exists(defaultPath))
                {
                    await Task.Run(() => Directory.CreateDirectory(defaultPath));
                    _logger.LogInformation("创建默认输出目录: {Path}", defaultPath);
                }
                
                OutputPath = defaultPath;
                OutputFolder = defaultPath;
            }
            
            _logger.LogDebug("设置默认输出路径: {Path}", OutputPath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "设置默认输出路径失败");
            OutputPath = string.Empty;
            OutputFolder = string.Empty;
        }
    }

    /// <summary>
    /// 生成默认文件名
    /// </summary>
    private string GenerateDefaultFileName(string sourceFileName, string mode, int index = 0)
    {
        var nameWithoutExt = Path.GetFileNameWithoutExtension(sourceFileName);
        var extension = Path.GetExtension(sourceFileName);
        
        if (string.IsNullOrEmpty(extension))
            extension = ".xlsx";
            
        if (mode == "Merge")
        {
            return $"{nameWithoutExt}_Merge{extension}";
        }
        else // Split
        {
            return $"{nameWithoutExt}_Split_{index:D3}{extension}";
        }
    }

    /// <summary>
    /// 是否有文件
    /// </summary>
    public bool HasFiles => SelectedFiles.Count > 0;

    /// <summary>
    /// 是否可以开始处理
    /// </summary>
    public bool CanStartProcessing => HasFiles && !IsProcessing && !string.IsNullOrWhiteSpace(OutputPath);

    /// <summary>
    /// 选择文件命令
    /// </summary>
    [RelayCommand]
    private async Task SelectFiles()
    {
        try
        {
            if (IsMergeMode)
            {
                // 合并模式：选择文件夹
                var dialog = new WinForms.FolderBrowserDialog
                {
                    Description = "选择包含Excel文件的文件夹",
                    ShowNewFolderButton = false
                };

                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    await AddFolderFilesAsync(dialog.SelectedPath);
                }
            }
            else
            {
                // 拆分模式：选择文件
                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "选择Excel文件",
                    Filter = "Excel文件|*.xlsx;*.xls;*.csv|所有文件|*.*",
                    Multiselect = true
                };

                if (dialog.ShowDialog() == true)
                {
                    await AddFilesAsync(dialog.FileNames);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "选择文件时发生错误");
            System.Windows.MessageBox.Show($"选择文件时发生错误: {ex.Message}", "错误", 
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 清空文件列表命令
    /// </summary>
    [RelayCommand]
    private void ClearFiles()
    {
        try
        {
            SelectedFiles.Clear();
            StatusText = "文件列表已清空";
            _logger.LogInformation("文件列表已清空");
        }
        catch (Exception ex)
        {
            HandleError(ex, "清空文件列表");
        }
    }

    /// <summary>
    /// 移除文件命令
    /// </summary>
    [RelayCommand]
    private void RemoveFile(FileInfo fileInfo)
    {
        try
        {
            if (fileInfo != null && SelectedFiles.Contains(fileInfo))
            {
                SelectedFiles.Remove(fileInfo);
                StatusText = $"已移除文件: {fileInfo.Name}";
                _logger.LogInformation("已移除文件: {FilePath}", fileInfo.Path);
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "移除文件");
        }
    }

    /// <summary>
    /// 浏览输出路径命令
    /// </summary>
    [RelayCommand]
    private void BrowseOutputPath()
    {
        try
        {
            // 合并模式和拆分模式都选择输出目录
            var dialog = new WinForms.FolderBrowserDialog
            {
                Description = IsMergeMode ? "选择输出目录" : "选择输出文件夹",
                ShowNewFolderButton = true
            };
    
            if (dialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                OutputPath = dialog.SelectedPath;
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "浏览输出路径");
        }
    }

    /// <summary>
    /// 开始处理命令
    /// </summary>
    [RelayCommand]
    private async Task StartProcessing()
    {
        if (!CanStartProcessing)
            return;

        try
        {
            IsProcessing = true;
            ProgressValue = 0f;
            ProgressText = "开始处理...";
            StatusText = "正在处理文件...";

            // 创建处理任务
            _currentTask = new ProcessingTask
            {
                TaskType = IsMergeMode ? TaskType.Merge : TaskType.Split,
                OutputPath = OutputPath,
                Config = new AppConfig
                {
                    AddHeaders = AddHeaders,
                    DedupeHeaders = DedupeHeaders
                }
            };

            // 添加文件到任务
            foreach (var file in SelectedFiles)
            {
                _currentTask.InputFiles.Add(file);
            }

            _currentTask.Status = TaskStatus.Processing;
            _logger.LogInformation("开始{Mode}处理，文件数量: {Count}", IsMergeMode ? "合并" : "拆分", SelectedFiles.Count);

            // 执行处理
            if (IsMergeMode)
            {
                await ProcessMergeAsync();
            }
            else
            {
                await ProcessSplitAsync();
            }
        }
        catch (OperationCanceledException)
        {
            StatusText = "处理已取消";
            _logger.LogInformation("处理已取消");
        }
        catch (Exception ex)
        {
            HandleError(ex, "文件处理");
        }
        finally
        {
            IsProcessing = false;
            _currentTask?.Dispose();
            _currentTask = null;
        }
    }

    /// <summary>
    /// 取消处理命令
    /// </summary>
    [RelayCommand]
    private void CancelProcessing()
    {
        try
        {
            _currentTask?.Cancel();
            StatusText = "正在取消处理...";
            _logger.LogInformation("用户取消了处理操作");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "取消处理时发生错误");
        }
    }

    /// <summary>
    /// 显示设置命令
    /// </summary>
    [RelayCommand]
    private void ShowSettings()
    {
        ExecuteWithErrorHandling(() =>
        {
            // TODO: 实现设置对话框
            System.Windows.MessageBox.Show("设置功能正在开发中...", "提示", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }, "显示设置");
    }

    /// <summary>
    /// 显示关于命令
    /// </summary>
    [RelayCommand]
    private void ShowAbout()
    {
        ExecuteWithErrorHandling(() =>
        {
            System.Windows.MessageBox.Show("Excel Tools Pro v1.0\n\n专业的Excel文件处理工具\n\n© 2024 All Rights Reserved", 
                "关于", MessageBoxButton.OK, MessageBoxImage.Information);
        }, "显示关于信息");
    }

    /// <summary>
    /// 添加文件 - 异步版本，避免UI线程阻塞
    /// </summary>
    public async Task AddFilesAsync(string[] filePaths)
    {
        try
        {
            StatusText = "正在添加文件...";
            
            // 在后台线程处理文件信息创建
            var fileInfos = await Task.Run(() => 
            {
                var results = new List<FileInfo>();
                foreach (var filePath in filePaths)
                {
                    if (SelectedFiles.Any(f => f.Path.Equals(filePath, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var fileInfo = CreateFileInfo(filePath);
                    results.Add(fileInfo);
                }
                return results;
            });
            
            // 在UI线程批量添加到集合
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
            {
                foreach (var fileInfo in fileInfos)
                {
                    SelectedFiles.Add(fileInfo);
                }
            });

            StatusText = $"已添加 {fileInfos.Count} 个文件";
            _logger.LogInformation("添加了 {Count} 个文件", fileInfos.Count);
        }
        catch (Exception ex)
        {
            HandleError(ex, "添加文件");
        }
    }
    
    /// <summary>
    /// 添加文件夹中的Excel文件 - 异步版本
    /// </summary>
    public async Task AddFolderFilesAsync(string folderPath)
    {
        try
        {
            StatusText = "正在扫描文件夹...";
            
            // 在后台线程扫描文件夹中的Excel文件
            var excelFiles = await Task.Run(() => 
            {
                return Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(file => SupportedFileExtensions.Contains(Path.GetExtension(file).ToLower()))
                    .ToArray();
            });
            
            if (excelFiles.Length == 0)
            {
                StatusText = "所选文件夹中未找到Excel文件";
                System.Windows.MessageBox.Show("所选文件夹中未找到Excel文件（.xlsx, .xls, .csv）", "提示", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            // 添加找到的Excel文件
            await AddFilesAsync(excelFiles);
            
            _logger.LogInformation("从文件夹 {FolderPath} 中添加了 {FileCount} 个Excel文件", folderPath, excelFiles.Length);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "扫描文件夹时发生错误: {FolderPath}", folderPath);
            HandleError(ex, "扫描文件夹");
        }
    }
    
    /// <summary>
    /// 添加文件 - 同步版本（保持向后兼容）
    /// </summary>
    public void AddFiles(string[] filePaths)
    {
        // 使用ConfigureAwait(false)避免捕获同步上下文
        _ = Task.Run(async () => await AddFilesAsync(filePaths).ConfigureAwait(false));
    }

    private static FileInfo CreateFileInfo(string filePath)
    {
        var systemFileInfo = new System.IO.FileInfo(filePath);
        var fileInfo = new FileInfo
        {
            Path = filePath,
            Name = systemFileInfo.Name,
            Size = systemFileInfo.Length,
            LastModified = systemFileInfo.LastWriteTime,
            FileType = GetFileType(systemFileInfo.Extension),
            IsValid = true
        };

        return fileInfo;
    }

    private static FileType GetFileType(string extension)
    {
        return extension.ToLower() switch
        {
            ".xlsx" => FileType.Xlsx,
            ".xls" => FileType.Xls,
            ".csv" => FileType.Csv,
            ".html" or ".htm" => FileType.Html,
            _ => FileType.Unknown
        };
    }

    private async Task ProcessMergeAsync()
    {
        if (_currentTask == null) return;

        // 生成默认输出文件名
        var firstFileName = SelectedFiles.FirstOrDefault()?.Name ?? "merged";
        var outputFileName = GenerateDefaultFileName(firstFileName, "Merge");
        var finalOutputPath = Path.Combine(OutputPath, outputFileName);

        var request = new MergeRequest
        {
            FilePaths = SelectedFiles.Select(f => f.Path).ToArray(),
            OutputDirectory = finalOutputPath,
            AddHeaders = AddHeaders,
            DedupeHeaders = DedupeHeaders
        };

        var progress = new Progress<float>(value =>
        {
            ProgressValue = value;
            ProgressText = $"处理进度: {value:F1}%";
        });

        var result = await _excelService.MergeExcelFilesAsync(request, progress, _currentTask.CancellationTokenSource.Token);

        if (result.Success)
        {
            ProgressValue = 100f;
            ProgressText = "处理完成";
            StatusText = $"文件合并完成: {result.OutputFile}";
            _currentTask.Status = TaskStatus.Completed;
            
            System.Windows.MessageBox.Show($"文件合并完成！\n输出文件: {result.OutputFile}", "成功", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        else
        {
            _currentTask.Status = TaskStatus.Failed;
            _currentTask.ErrorMessage = result.Message;
            throw new InvalidOperationException(result.Message);
        }
    }

    private async Task ProcessSplitAsync()
    {
        if (_currentTask == null || SelectedFiles.Count == 0) return;

        var totalFiles = SelectedFiles.Count;
        var totalOutputFiles = 0;
        var processedFiles = 0;

        var progress = new Progress<float>(value =>
        {
            // 计算总体进度：当前文件进度 + 已完成文件数
            var overallProgress = (processedFiles * 100f + value) / totalFiles;
            ProgressValue = overallProgress;
            ProgressText = $"处理进度: {overallProgress:F1}% ({processedFiles + 1}/{totalFiles})";
        });

        // 遍历处理所有选中的文件
        foreach (var file in SelectedFiles)
        {
            _currentTask.CancellationTokenSource.Token.ThrowIfCancellationRequested();

            // 为每个文件创建独立的输出目录
            var fileOutputDir = Path.Combine(OutputPath, Path.GetFileNameWithoutExtension(file.Name));
            Directory.CreateDirectory(fileOutputDir);

            var request = new SplitRequest
            {
                FilePath = file.Path,
                OutputDirectory = fileOutputDir,
                SplitBy = SplitMode.ByRows,
                RowsPerFile = SplitRowCount,
                AddHeaders = AutoAddHeader
            };

            var result = await _excelService.SplitExcelFileAsync(request, progress, _currentTask.CancellationTokenSource.Token);

            if (result.Success)
            {
                totalOutputFiles += result.OutputFiles?.Count ?? 0;
                processedFiles++;
                _logger.LogInformation("文件 {FileName} 拆分完成，生成了 {OutputFileCount} 个文件", file.Name, result.OutputFiles?.Count ?? 0);
            }
            else
            {
                _currentTask.Status = TaskStatus.Failed;
                _currentTask.ErrorMessage = $"处理文件 {file.Name} 时发生错误: {result.Message}";
                throw new InvalidOperationException(_currentTask.ErrorMessage);
            }
        }

        // 所有文件处理完成
        ProgressValue = 100f;
        ProgressText = "处理完成";
        StatusText = $"文件拆分完成，共处理 {totalFiles} 个文件，生成了 {totalOutputFiles} 个文件";
        _currentTask.Status = TaskStatus.Completed;
        
        System.Windows.MessageBox.Show($"文件拆分完成！\n共处理 {totalFiles} 个文件\n生成了 {totalOutputFiles} 个文件\n输出目录: {OutputPath}", "成功", 
            MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private async Task LoadConfigurationAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            _logger.LogDebug("开始异步加载配置");
            
            var config = await _configService.GetConfigurationAsync(cancellationToken).ConfigureAwait(false);
            
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                AddHeaders = config.AddHeaders;
                DedupeHeaders = config.DedupeHeaders;
                CustomSplitRowCount = config.CustomSplitRowCount > 0 ? config.CustomSplitRowCount : 1000;
                
                // 设置默认输出路径
                if (string.IsNullOrWhiteSpace(config.LastOutputDirectory))
                {
                    await SetDefaultOutputPathAsync().ConfigureAwait(false);
                }
                else
                {
                    OutputPath = config.LastOutputDirectory;
                }
                
                StatusText = "配置加载完成";
            }, System.Windows.Threading.DispatcherPriority.Background, cancellationToken);
            
            _logger.LogDebug("配置加载完成");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            _logger.LogDebug("配置加载被取消");
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "加载配置失败，使用默认配置");
            
            await SafeUpdateUIAsync(async () =>
            {
                AddHeaders = true;
                DedupeHeaders = true;
                AutoAddHeader = true;
                CustomSplitRowCount = 1000;
                SplitRowCount = 1000;
                await SetDefaultOutputPathAsync().ConfigureAwait(false);
                StatusText = "配置加载失败，已使用默认配置";
            }, "加载默认配置").ConfigureAwait(false);
        }
    }

    /// <summary>
    /// 模式切换时的处理
    /// </summary>
    partial void OnIsMergeModeChanged(bool value)
    {
        try
        {
            _logger.LogDebug("操作模式切换为: {Mode}", value ? "合并模式" : "拆分模式");
            
            // 清空当前文件列表
            SelectedFiles.Clear();
            
            // 设置模式特定的默认配置
            if (value)
            {
                // 合并模式：默认开启表头去重
                DedupeHeaders = true;
                AddHeaders = true;
            }
            else
            {
                // 拆分模式：默认开启自动添加表头
                AutoAddHeader = true;
                AddHeaders = true;
            }
            
            // 更新默认输出路径 - 使用异步操作避免UI线程阻塞
            _ = Task.Run(async () =>
            {
                await SetDefaultOutputPathAsync();
            });
            
            // 更新状态文本
            StatusText = value ? "已切换到合并模式" : "已切换到拆分模式";
            
            // 通知相关属性变化
            OnPropertyChanged(nameof(IsSplitMode));
        }
        catch (Exception ex)
        {
            HandleError(ex, "模式切换");
        }
    }

    private void OnFilesCollectionChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        try
        {
            // 延迟更新UI以提高性能
            _ = Task.Run(async () =>
            {
                await Task.Delay(100).ConfigureAwait(false);
                
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    OnPropertyChanged(nameof(HasFiles));
                    OnPropertyChanged(nameof(CanStartProcessing));
                    FileCountText = $"{SelectedFiles.Count} 个文件";
                });
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "文件集合变化处理失败");
            // 不显示用户错误对话框，这是内部处理错误
        }
    }

    private void OnMetricsUpdated(object? sender, SystemMetrics metrics)
    {
        try
        {
            // 使用节流更新UI
            _ = Task.Run(async () =>
            {
                await Task.Delay(500).ConfigureAwait(false);
                
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    MemoryUsageText = $"{metrics.MemoryUsagePercent:F1}%";
                    
                    if (metrics.MemoryUsagePercent > 80)
                    {
                        UiResponseStatus = "内存使用较高";
                    }
                    else
                    {
                        UiResponseStatus = "响应正常";
                    }
                });
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "性能指标更新失败");
            // 不显示用户错误对话框，这是内部处理错误
        }
    }



    private void HandleError(Exception ex, string operation)
    {
        _logger.LogError(ex, "{Operation} 操作失败，异常类型: {ExceptionType}", operation, ex.GetType().Name);
        
        var userFriendlyMessage = GetUserFriendlyErrorMessage(ex, operation);
        StatusText = userFriendlyMessage;
        
        // 显示详细的错误对话框
        ShowErrorDialog(ex, operation, userFriendlyMessage);
    }

    private void ExecuteWithErrorHandling(Action action, string operation)
    {
        try
        {
            action();
        }
        catch (OperationCanceledException)
        {
            _logger.LogInformation("{Operation} 操作被取消", operation);
            StatusText = $"{operation}已取消";
        }
        catch (Exception ex)
        {
            HandleError(ex, operation);
        }
    }
    
    /// <summary>
    /// 获取用户友好的错误消息
    /// </summary>
    private static string GetUserFriendlyErrorMessage(Exception ex, string operation)
    {
        var baseMessage = ex switch
        {
            FileNotFoundException => "找不到指定的文件",
            DirectoryNotFoundException => "找不到指定的目录",
            UnauthorizedAccessException => "没有足够的权限访问文件",
            OutOfMemoryException => "系统内存不足",
            InvalidOperationException => "当前操作无效",
            ArgumentException => "参数错误",
            TimeoutException => "操作超时",
            System.IO.IOException => "文件操作失败",
            NotSupportedException => "不支持的文件格式或操作",
            InvalidDataException => "文件数据格式错误",
            _ => "发生了意外错误"
        };
        
        return $"{operation}失败: {baseMessage}";
    }
    
    /// <summary>
    /// 显示错误对话框
    /// </summary>
    private void ShowErrorDialog(Exception ex, string operation, string userFriendlyMessage)
    {
        try
        {
            var detailedMessage = $"{userFriendlyMessage}\n\n详细信息：\n{ex.Message}";
            
            if (ex.InnerException != null)
            {
                detailedMessage += $"\n\n内部异常：\n{ex.InnerException.Message}";
            }
            
            // 在UI线程显示错误对话框
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                var result = System.Windows.MessageBox.Show(
                    detailedMessage,
                    $"{operation}错误",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            });
        }
        catch (Exception dialogEx)
        {
            _logger.LogError(dialogEx, "显示错误对话框时发生异常");
        }
    }
    
    /// <summary>
    /// 安全执行UI更新操作
    /// </summary>
    private void SafeUpdateUI(Action uiAction, string operationName = "UI更新")
    {
        try
        {
            if (System.Windows.Application.Current?.Dispatcher.CheckAccess() == true)
            {
                uiAction();
            }
            else
            {
                System.Windows.Application.Current?.Dispatcher.Invoke(uiAction);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "{OperationName} 操作失败", operationName);
        }
    }

    /// <summary>
    /// 安全执行UI更新操作（异步版本）
    /// </summary>
    private async Task SafeUpdateUIAsync(Func<Task> uiAction, string operationName = "UI更新")
    {
        try
        {
            if (System.Windows.Application.Current?.Dispatcher.CheckAccess() == true)
            {
                await uiAction().ConfigureAwait(false);
            }
            else
            {
                var dispatcher = System.Windows.Application.Current?.Dispatcher;
                if (dispatcher != null)
                {
                    await dispatcher.InvokeAsync(async () => await uiAction().ConfigureAwait(false));
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "{Operation} 失败", operationName);
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _performanceService.MetricsUpdated -= OnMetricsUpdated;
            _currentTask?.Dispose();
            
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}