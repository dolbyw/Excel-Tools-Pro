using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;
using ExcelToolsPro.Models;
using ExcelToolsPro.Services;

namespace ExcelToolsPro.Services;

/// <summary>
/// Excel处理服务实现
/// </summary>
public partial class ExcelProcessingService : IExcelProcessingService, IDisposable
{
    private readonly ILogger<ExcelProcessingService> _logger;
    private readonly SemaphoreSlim _semaphore;
    private readonly IPerformanceMonitorService _performanceMonitor;
    private readonly IFileSystemService _fileSystemService;
    private readonly ILowMemoryProcessor _lowMemoryProcessor;
    private readonly IConfigurationService _configService;
    private bool _disposed = false;

    public ExcelProcessingService(
        ILogger<ExcelProcessingService> logger,
        IPerformanceMonitorService performanceMonitor,
        IFileSystemService fileSystemService,
        ILowMemoryProcessor lowMemoryProcessor,
        IConfigurationService configService)
    {
        var stopwatch = Stopwatch.StartNew();
        
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _performanceMonitor = performanceMonitor ?? throw new ArgumentNullException(nameof(performanceMonitor));
        _fileSystemService = fileSystemService ?? throw new ArgumentNullException(nameof(fileSystemService));
        _lowMemoryProcessor = lowMemoryProcessor ?? throw new ArgumentNullException(nameof(lowMemoryProcessor));
        _configService = configService ?? throw new ArgumentNullException(nameof(configService));
        
        _logger.LogDebug("=== ExcelProcessingService 模块初始化开始 ===");
        _logger.LogDebug("注入的日志服务状态: {LoggerStatus}", logger != null ? "有效" : "无效");
        
        // 初始化信号量
        _logger.LogDebug("初始化并发控制信号量...");
        _semaphore = new SemaphoreSlim(4, 4); // 默认并发度，将在实际使用时获取配置
        _logger.LogDebug("信号量初始化完成，使用默认并发数: 4");
        
        // 检查ClosedXML库可用性
        _logger.LogDebug("检查ClosedXML库可用性...");
        try
        {
            using var testWorkbook = new XLWorkbook();
            _logger.LogDebug("ClosedXML库检查通过");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ClosedXML库初始化失败，错误类型: {ExceptionType}", ex.GetType().Name);
            throw new InvalidOperationException("Excel处理库初始化失败", ex);
        }
        
        _logger.LogInformation("ExcelProcessingService 模块初始化完成，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
    }
    
    /// <summary>
    /// 低内存模式下的文件合并处理
    /// </summary>
    private async Task<ProcessingResult> ProcessMergeWithLowMemoryMode(
        MergeRequest request, 
        IProgress<float>? progress, 
        CancellationToken cancellationToken,
        string operationId,
        string outputPath,
        PerformanceTimer timer)
    {
        try
        {
            _logger.LogInformation("开始低内存模式文件合并 - OperationId: {OperationId}, FileCount: {FileCount}", 
                operationId, request.FilePaths.Length);
            
            // 异步获取配置
            var config = await _configService.GetConfigurationAsync(cancellationToken).ConfigureAwait(false);
            
            // 在低内存模式下，限制并发度
            var lowMemoryConcurrency = Math.Min(config.MaxDegreeOfParallelism, 2);
            using var lowMemorySemaphore = new SemaphoreSlim(lowMemoryConcurrency, lowMemoryConcurrency);
            
            timer.Checkpoint("创建低内存模式工作簿");
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("合并数据");
            
            var currentRow = 1;
            var processedFiles = 0;
            var failedFiles = new List<string>();
            var throttler = new ProgressThrottler(progress, config.ProgressThrottleMs, 1f, _logger);
            
            // 分批处理文件，避免同时打开太多文件
            var batchSize = Math.Max(1, config.MaxDegreeOfParallelism / 2);
            
            for (int batchStart = 0; batchStart < request.FilePaths.Length; batchStart += batchSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                var batchEnd = Math.Min(batchStart + batchSize, request.FilePaths.Length);
                var batchFiles = request.FilePaths.Skip(batchStart).Take(batchEnd - batchStart).ToArray();
                
                _logger.LogDebug("处理文件批次 - OperationId: {OperationId}, BatchStart: {BatchStart}, BatchSize: {BatchSize}", 
                    operationId, batchStart, batchFiles.Length);
                
                // 并行处理批次内的文件
                var tasks = batchFiles.Select(async (filePath, index) =>
                {
                    await lowMemorySemaphore.WaitAsync(cancellationToken);
                    try
                    {
                        var isFirstFile = batchStart == 0 && index == 0;
                        var fileProgress = new Progress<float>(p =>
                        {
                            var overallProgress = ((float)(batchStart + index) / request.FilePaths.Length + p / 100f / request.FilePaths.Length) * 100f;
                            throttler.Report(overallProgress);
                        });
                        
                        return await ProcessSingleFileForMergeLowMemory(
                            filePath, 
                            worksheet, 
                            currentRow + (index * 10000), // 预留空间避免冲突
                            isFirstFile && request.AddHeaders,
                            request.DedupeHeaders,
                            fileProgress,
                            cancellationToken);
                    }
                    finally
                    {
                        lowMemorySemaphore.Release();
                    }
                });
                
                var results = await Task.WhenAll(tasks);
                
                // 合并结果
                foreach (var result in results.Where(r => r.Success))
                {
                    processedFiles++;
                }
                
                foreach (var result in results.Where(r => !r.Success))
                {
                    failedFiles.Add(result.Message ?? "未知文件");
                }
                
                // 更新当前行位置
                currentRow += results.Where(r => r.Success).Sum(r => r.RowsProcessed);
                
                // 批次间让出CPU时间
                await Task.Yield();
            }
            
            timer.Checkpoint("保存低内存模式合并文件");
            
            // 确保输出目录存在
            if (!Directory.Exists(request.OutputDirectory))
            {
                Directory.CreateDirectory(request.OutputDirectory);
            }
            
            await Task.Run(() => workbook.SaveAs(outputPath), cancellationToken);
            
            throttler.Report(100f, true);
            
            _logger.LogInformation("低内存模式文件合并完成 - OperationId: {OperationId}, ProcessedFiles: {ProcessedFiles}, FailedFiles: {FailedFiles}", 
                operationId, processedFiles, failedFiles.Count);
            
            return new ProcessingResult
            {
                Success = true,
                Message = failedFiles.Count > 0 
                    ? $"文件合并完成（低内存模式），但有 {failedFiles.Count} 个文件处理失败" 
                    : "文件合并完成（低内存模式）",
                OutputFile = outputPath
            };
        }
        catch (Exception ex)
        {
            timer.LogError($"低内存模式合并失败 - ExceptionType: {ex.GetType().Name}, Message: {ex.Message}");
            return new ProcessingResult
            {
                Success = false,
                Message = $"低内存模式合并失败: {ex.Message}"
            };
        }
    }
    
    /// <summary>
    /// 低内存模式下处理单个文件的合并
    /// </summary>
    private async Task<(bool Success, int RowsProcessed, string Message)> ProcessSingleFileForMergeLowMemory(
        string filePath, 
        IXLWorksheet targetWorksheet, 
        int startRow,
        bool includeHeaders,
        bool dedupeHeaders,
        IProgress<float> progress,
        CancellationToken cancellationToken)
    {
        try
        {
            var extension = Path.GetExtension(filePath).ToLower();
            
            // 对于CSV文件，使用低内存处理器
            if (extension == ".csv")
            {
                var tempOutputPath = Path.GetTempFileName() + ".xlsx";
                try
                {
                    var config = await _configService.GetConfigurationAsync(cancellationToken).ConfigureAwait(false);
                    var result = await _lowMemoryProcessor.ProcessCsvToExcelLowMemoryAsync(
                        filePath, tempOutputPath, config, progress, cancellationToken);
                    
                    if (result.Success)
                    {
                        // 将临时Excel文件的内容复制到目标工作表
                        using var tempWorkbook = new XLWorkbook(tempOutputPath);
                        var tempWorksheet = tempWorkbook.Worksheets.First();
                        var usedRange = tempWorksheet.RangeUsed();
                        
                        if (usedRange != null)
                        {
                            var rowCount = usedRange.RowCount();
                            var startRowToUse = includeHeaders ? startRow : startRow;
                            var sourceStartRow = includeHeaders ? 1 : 2;
                            
                            for (int row = sourceStartRow; row <= rowCount; row++)
                            {
                                var sourceRow = tempWorksheet.Row(row);
                                var targetRow = targetWorksheet.Row(startRowToUse + row - sourceStartRow);
                                sourceRow.CopyTo(targetRow);
                            }
                            
                            return (true, rowCount - (includeHeaders ? 0 : 1), filePath);
                        }
                    }
                    
                    return (false, 0, filePath);
                }
                finally
                {
                    if (File.Exists(tempOutputPath))
                    {
                        File.Delete(tempOutputPath);
                    }
                }
            }
            
            // 对于其他文件类型，使用标准处理但限制内存使用
            var rowsProcessed = await ProcessSingleFileForMerge(
                filePath, targetWorksheet, startRow, includeHeaders, dedupeHeaders, progress, cancellationToken);
            
            return (true, rowsProcessed - startRow, filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "低内存模式处理单个文件失败: {FilePath}", filePath);
            return (false, 0, filePath);
        }
    }

    public async Task<ProcessingResult> MergeExcelFilesAsync(
        MergeRequest request, 
        IProgress<float>? progress = null, 
        CancellationToken cancellationToken = default)
    {
        var operationId = Guid.NewGuid().ToString("N")[..8];
        
        using var timer = PerformanceTimerExtensions.CreateTimer(_logger, "MergeExcelFiles", new { OperationId = operationId });
        
        _logger.LogInformation("开始Excel文件合并操作 - OperationId: {OperationId}, FileCount: {FileCount}", 
            operationId, request?.FilePaths?.Length ?? 0);
        
        // 输入验证
        timer.Checkpoint("开始输入参数验证");
        
        if (request == null)
        {
            _logger.LogError("[{OperationId}] 合并请求对象为空", operationId);
            throw new ArgumentNullException(nameof(request));
        }
        
        _logger.LogDebug("请求参数验证 - OperationId: {OperationId}, FileCount: {FileCount}, OutputDirectory: {OutputDirectory}, AddHeaders: {AddHeaders}, DedupeHeaders: {DedupeHeaders}", 
            operationId, request.FilePaths?.Length ?? 0, request.OutputDirectory, request.AddHeaders, request.DedupeHeaders);
        
        if (request.FilePaths == null || request.FilePaths.Length == 0)
        {
            timer.LogError("没有选择要合并的文件");
            return new ProcessingResult
            {
                Success = false,
                Message = "没有选择要合并的文件"
            };
        }
        
        if (string.IsNullOrWhiteSpace(request.OutputDirectory))
        {
            timer.LogError("输出目录为空");
            return new ProcessingResult
            {
                Success = false,
                Message = "输出目录不能为空"
            };
        }
        
        // 生成输出文件名：源文件名+Merge.xlsx
        timer.Checkpoint("生成输出文件名");
        var firstFileName = Path.GetFileNameWithoutExtension(request.FilePaths[0]);
        if (string.IsNullOrEmpty(firstFileName))
        {
            firstFileName = "合并文件";
        }
        var outputFileName = $"{firstFileName}_Merge.xlsx";
        var outputPath = Path.Combine(request.OutputDirectory, outputFileName);
        
        _logger.LogDebug("生成输出文件名 - OperationId: {OperationId}, OutputFileName: {OutputFileName}, OutputPath: {OutputPath}", 
            operationId, outputFileName, outputPath);
        
        // 验证输入文件
        timer.Checkpoint("验证输入文件存在性");
        var missingFiles = request.FilePaths.Where(f => !File.Exists(f)).ToArray();
        if (missingFiles.Length > 0)
        {
            timer.LogError($"发现 {missingFiles.Length} 个不存在的文件: {string.Join(", ", missingFiles)}");
            return new ProcessingResult
            {
                Success = false,
                Message = $"以下文件不存在: {string.Join(", ", missingFiles.Select(Path.GetFileName))}"
            };
        }
        
        // 检查是否需要使用低内存模式
        timer.Checkpoint("检查低内存模式");
        var config = await _configService.GetConfigurationAsync(cancellationToken).ConfigureAwait(false);
        var shouldUseLowMemory = _lowMemoryProcessor.ShouldUseLowMemoryMode(request.FilePaths, config);
        if (shouldUseLowMemory)
        {
            _logger.LogInformation("启用低内存模式进行文件合并 - OperationId: {OperationId}", operationId);
            return await ProcessMergeWithLowMemoryMode(request, progress, cancellationToken, operationId, outputPath, timer);
        }
        
        timer.Checkpoint("等待并发控制信号量");
        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        timer.Checkpoint("获得并发控制权限");
        
        try
        {
            _logger.LogInformation("开始合并文件 - OperationId: {OperationId}, FileCount: {FileCount}, OutputPath: {OutputPath}", 
                operationId, request.FilePaths.Length, outputPath);
            
            // 创建工作簿
            timer.Checkpoint("创建Excel工作簿");
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("合并数据");
            timer.Checkpoint("工作簿创建完成");
            
            int currentRow = 1;
            bool isFirstFile = true;
            var processedFiles = 0;
            var failedFiles = new List<string>();

            // 主进度节流器 (基于文件数量)
            var mainProgressThrottler = new ProgressThrottler(progress, 200, 1f, _logger);

            // 降低循环内日志频率
            const int logEveryN = 10;
            
            for (int i = 0; i < request.FilePaths.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                var filePath = request.FilePaths[i];
                
                bool shouldLogThisIteration = (i == 0) || (i == request.FilePaths.Length - 1) || (i % logEveryN == 0);
                if (shouldLogThisIteration)
                {
                    _logger.LogDebug("处理文件 - OperationId: {OperationId}, FileIndex: {FileIndex}, TotalFiles: {TotalFiles}, FilePath: {FilePath}", 
                        operationId, i + 1, request.FilePaths.Length, filePath);
                }
                
                try
                {
                    var fileInfo = new System.IO.FileInfo(filePath);
                    if (shouldLogThisIteration)
                    {
                        _logger.LogDebug("文件信息 - OperationId: {OperationId}, FileSize: {FileSize}, Extension: {Extension}, FilePath: {FilePath}", 
                            operationId, fileInfo.Length, fileInfo.Extension, filePath);
                    }

                    // 为每个文件创建分层进度报告
                    var fileProgress = new Progress<float>(p =>
                    {
                        // p 的范围是 0-100 (来自子任务)
                        // 将其转换为在总进度中的一小部分
                        float baseProgress = (float)i / request.FilePaths.Length * 100;
                        float weightedProgress = p / request.FilePaths.Length;
                        mainProgressThrottler.Report(baseProgress + weightedProgress);
                    });
                    
                    var originalRow = currentRow;
                    currentRow = await ProcessSingleFileForMerge(
                        filePath, 
                        worksheet, 
                        currentRow, 
                        isFirstFile && request.AddHeaders,
                        request.DedupeHeaders,
                        fileProgress, // 传递行级进度报告器
                        cancellationToken).ConfigureAwait(false);
                    
                    var rowsAdded = currentRow - originalRow;
                    processedFiles++;
                    
                    if (shouldLogThisIteration)
                    {
                        _logger.LogDebug("文件处理完成 - OperationId: {OperationId}, RowsAdded: {RowsAdded}, ProcessedFiles: {ProcessedFiles}, FilePath: {FilePath}", 
                            operationId, rowsAdded, processedFiles, filePath);
                    }
                    
                    isFirstFile = false;
                }
                catch (Exception ex)
                {
                    failedFiles.Add(filePath);
                    timer.LogError($"处理文件时发生错误 - FilePath: {filePath}, ExceptionType: {ex.GetType().Name}, Message: {ex.Message}");
                    
                    // 尝试恢复处理
                    if (await TryRecoverFromFileError(filePath, ex, operationId).ConfigureAwait(false))
                    {
                        _logger.LogInformation("文件错误恢复成功 - OperationId: {OperationId}, FilePath: {FilePath}", operationId, filePath);
                        failedFiles.Remove(filePath); // 从失败列表中移除
                        processedFiles++;
                    }
                    // 继续处理其他文件，不中断整个流程
                }
                
                // 更新主进度 (基于文件)
                mainProgressThrottler.Report((float)(i + 1) / request.FilePaths.Length * 100);
                if (shouldLogThisIteration)
                {
                    _logger.LogDebug("主进度更新 - OperationId: {OperationId}, Progress: {Progress:F1}%, ProcessedFiles: {ProcessedFiles}", 
                        operationId, (float)(i + 1) / request.FilePaths.Length * 100, processedFiles);
                }
            }
            
            // 确保最终进度为100%
            mainProgressThrottler.Report(100f, true);

            // 确保输出目录存在
            timer.Checkpoint("检查输出目录");
            if (!Directory.Exists(request.OutputDirectory))
            {
                _logger.LogDebug("创建输出目录 - OperationId: {OperationId}, OutputDir: {OutputDir}", operationId, request.OutputDirectory);
                Directory.CreateDirectory(request.OutputDirectory);
            }
            
            // 保存合并后的文件
            timer.Checkpoint("保存合并文件");
            await Task.Run(() => workbook.SaveAs(outputPath), cancellationToken).ConfigureAwait(false);
            timer.Checkpoint("文件保存完成");
            
            var outputFileInfo = new System.IO.FileInfo(outputPath);
            _logger.LogInformation("文件合并完成 - OperationId: {OperationId}, OutputPath: {OutputPath}, FileSize: {FileSize}, ProcessedFiles: {ProcessedFiles}, TotalFiles: {TotalFiles}, FailedFiles: {FailedFiles}", 
                operationId, outputPath, outputFileInfo.Length, processedFiles, request.FilePaths.Length, failedFiles.Count);
            
            if (failedFiles.Count > 0)
            {
                timer.LogWarning($"部分文件处理失败: {string.Join(", ", failedFiles.Select(Path.GetFileName))}");
            }
            
            return new ProcessingResult
            {
                Success = true,
                Message = failedFiles.Count > 0 
                    ? $"文件合并完成，但有 {failedFiles.Count} 个文件处理失败" 
                    : "文件合并完成",
                OutputFile = outputPath
            };
        }
        catch (OperationCanceledException)
        {
            timer.LogWarning("文件合并操作被取消");
            throw;
        }
        catch (Exception ex)
        {
            timer.LogError($"合并文件时发生严重错误 - ExceptionType: {ex.GetType().Name}, Message: {ex.Message}");
            return new ProcessingResult
            {
                Success = false,
                Message = $"合并文件时发生错误: {ex.Message}"
            };
        }
        finally
        {
            _semaphore.Release();
            _logger.LogDebug("释放并发控制权限 - OperationId: {OperationId}", operationId);
        }
    }

    public async Task<ProcessingResult> SplitExcelFileAsync(
        SplitRequest request, 
        IProgress<float>? progress = null, 
        CancellationToken cancellationToken = default)
    {
        // 输入验证
        ArgumentNullException.ThrowIfNull(request);
        
        if (string.IsNullOrWhiteSpace(request.FilePath))
        {
            return new ProcessingResult
            {
                Success = false,
                Message = "文件路径不能为空"
            };
        }
        
        if (!File.Exists(request.FilePath))
        {
            return new ProcessingResult
            {
                Success = false,
                Message = "要拆分的文件不存在"
            };
        }
        
        if (string.IsNullOrWhiteSpace(request.OutputDirectory))
        {
            return new ProcessingResult
            {
                Success = false,
                Message = "输出目录不能为空"
            };
        }

        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        
        try
        {
            _logger.LogInformation("开始拆分文件: {FilePath}", request.FilePath);
            
            // 确保输出目录存在
            if (!Directory.Exists(request.OutputDirectory))
            {
                Directory.CreateDirectory(request.OutputDirectory);
            }
            
            var outputFiles = new List<string>();
            
            // 检查是否为HTML格式的XLS文件
            var extension = Path.GetExtension(request.FilePath).ToLower();
            if (extension == ".xls" && await IsHtmlDisguisedFile(request.FilePath, cancellationToken).ConfigureAwait(false))
            {
                // 处理HTML格式的XLS文件
                outputFiles = await SplitHtmlXlsFile(request.FilePath, request.OutputDirectory, request.RowsPerFile ?? 1000, progress, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                // 在后台线程打开工作簿，避免在UI线程上进行重负载I/O与解析
                using var workbook = await Task.Run(() => new XLWorkbook(request.FilePath), cancellationToken).ConfigureAwait(false);
                var worksheets = workbook.Worksheets.ToList();
                
                if (request.SplitBy == SplitMode.BySheet)
                {
                    outputFiles = await SplitBySheets(workbook, request.FilePath, request.OutputDirectory, progress, cancellationToken).ConfigureAwait(false);
                }
                else
                {
                    outputFiles = await SplitByRows(workbook, request.FilePath, request.OutputDirectory, request.RowsPerFile ?? 1000, progress, cancellationToken).ConfigureAwait(false);
                }
             }
            
            _logger.LogInformation("文件拆分完成，生成 {FileCount} 个文件", outputFiles.Count);
            
            return new ProcessingResult
            {
                Success = true,
                Message = $"文件拆分完成，生成 {outputFiles.Count} 个文件",
                OutputFiles = outputFiles
            };
        }
        catch (OperationCanceledException)
        {
            _logger.LogInformation("文件拆分操作被取消");
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "拆分文件时发生错误");
            return new ProcessingResult
            {
                Success = false,
                Message = $"拆分文件时发生错误: {ex.Message}"
            };
        }
        finally
        {
            _semaphore.Release();
        }
    }

    public async Task<ValidationResult> ValidateExcelFilesAsync(
        string[] filePaths, 
        CancellationToken cancellationToken = default)
    {
        var result = new ValidationResult();
        var validFiles = new List<string>();
        var invalidFiles = new List<string>();
        var htmlFiles = new List<string>();
        var fileErrors = new Dictionary<string, string>();
        
        foreach (var filePath in filePaths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            try
            {
                if (!File.Exists(filePath))
                {
                    invalidFiles.Add(filePath);
                    fileErrors[filePath] = "文件不存在";
                    continue;
                }
                
                var extension = Path.GetExtension(filePath).ToLower();
                
                // 检查是否为HTML伪装文件
                if (extension == ".xls" && await IsHtmlDisguisedFile(filePath, cancellationToken).ConfigureAwait(false))
                {
                    htmlFiles.Add(filePath);
                    continue;
                }
                
                // 尝试打开文件验证
                if (extension is ".xlsx" or ".xls")
                {
                    await Task.Run(() =>
                    {
                        using var workbook = new XLWorkbook(filePath);
                        // 如果能成功打开，说明文件有效
                    }, cancellationToken).ConfigureAwait(false);
                }
                
                validFiles.Add(filePath);
            }
            catch (Exception ex)
            {
                invalidFiles.Add(filePath);
                fileErrors[filePath] = ex.Message;
                _logger.LogWarning(ex, "验证文件 {FilePath} 时发生错误", filePath);
            }
        }
        
        result.ValidFiles = [.. validFiles];
        result.InvalidFiles = [.. invalidFiles];
        result.HtmlFiles = [.. htmlFiles];
        result.FileErrors = fileErrors;
        
        return result;
    }

    private async Task<int> ProcessSingleFileForMerge(
        string filePath, 
        IXLWorksheet targetWorksheet, 
        int currentRow,
        bool includeHeaders,
        bool _,  // dedupeHeaders parameter not used in this method
        IProgress<float> progress, // 新增进度参数
        CancellationToken cancellationToken)
    {
        var extension = Path.GetExtension(filePath).ToLower();
        
        if (extension == ".csv")
        {
            return await ProcessCsvFile(filePath, targetWorksheet, currentRow, includeHeaders, progress, cancellationToken).ConfigureAwait(false);
        }
        else if (extension == ".xls" && await IsHtmlDisguisedFile(filePath, cancellationToken).ConfigureAwait(false))
        {
            // 处理HTML格式的XLS文件
            return await ProcessHtmlXlsFile(filePath, targetWorksheet, currentRow, includeHeaders, progress, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            return await ProcessExcelFile(filePath, targetWorksheet, currentRow, includeHeaders, progress, cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task<int> ProcessExcelFile(
        string filePath, 
        IXLWorksheet targetWorksheet, 
        int currentRow,
        bool includeHeaders,
        IProgress<float> progress, // 新增进度参数
        CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            using var sourceWorkbook = new XLWorkbook(filePath);
            var sourceWorksheet = sourceWorkbook.Worksheets.First();
            
            var usedRange = sourceWorksheet.RangeUsed();
            if (usedRange == null) return currentRow;
            
            int startRow = includeHeaders ? 1 : 2;
            int newCurrentRow = currentRow;
            int totalRowsToProcess = usedRange.RowCount() - startRow + 1;
            
            var throttler = new ProgressThrottler(progress, 100, 5f, _logger); // 报告间隔100ms或进度变化5%

            for (int i = 0; i < totalRowsToProcess; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var sourceRowNumber = startRow + i;
                
                var sourceRow = sourceWorksheet.Row(sourceRowNumber);
                var targetRow = targetWorksheet.Row(newCurrentRow + i);
                
                // 性能优化：整行复制
                sourceRow.CopyTo(targetRow);

                // 报告进度
                throttler.Report((float)(i + 1) / totalRowsToProcess * 100);
            }
            
            newCurrentRow += totalRowsToProcess;
            throttler.Report(100f, true); // 确保最后报告100%

            return newCurrentRow;
        }, cancellationToken).ConfigureAwait(false);
    }

    private async Task<int> ProcessCsvFile(
        string filePath, 
        IXLWorksheet targetWorksheet, 
        int currentRow,
        bool includeHeaders,
        IProgress<float> progress, // 新增进度参数
        CancellationToken cancellationToken)
    {
        // 异步获取配置
        var config = await _configService.GetConfigurationAsync(cancellationToken).ConfigureAwait(false);
        
        // 使用异步I/O和配置化的缓冲区大小
        var newCurrentRow = currentRow;
        var bufferSize = config.IOBufferSizeKB * 1024;
        var encoding = GetCsvEncodingFromConfig(config);
        
        // 为了计算进度，需要先获取总行数
        long totalLines = 0;
        using (var stream = await _fileSystemService.CreateFileStreamAsync(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, cancellationToken).ConfigureAwait(false))
        using (var preReader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize))
        {
            while (await preReader.ReadLineAsync().ConfigureAwait(false) != null)
            {
                totalLines++;
            }
        }
        
        var throttler = new ProgressThrottler(progress, config.ProgressThrottleMs, 5f, _logger);

        using var fileStream = await _fileSystemService.CreateFileStreamAsync(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, cancellationToken).ConfigureAwait(false);
        using var reader = new StreamReader(fileStream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize);
        
        string? line;
        bool isFirstLine = true;
        long linesProcessed = 0;
        
        while ((line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
        {
            cancellationToken.ThrowIfCancellationRequested();
            linesProcessed++;
            
            if (!includeHeaders && isFirstLine)
            {
                isFirstLine = false;
                continue; // 跳过表头
            }
            isFirstLine = false;

            var values = ParseCsvLine(line);
            for (int col = 0; col < values.Length; col++)
            {
                targetWorksheet.Cell(newCurrentRow, col + 1).Value = values[col];
            }
            newCurrentRow++;
            
            // 报告进度
            throttler.Report((float)linesProcessed / totalLines * 100);
        }

        throttler.Report(100f, true); // 确保最后报告100%
        return newCurrentRow;
    }

    private async Task<List<string>> SplitBySheets(
        XLWorkbook workbook,
        string sourceFilePath,
        string outputDirectory, 
        IProgress<float>? progress, 
        CancellationToken cancellationToken)
    {
        var outputFiles = new List<string>();
        var worksheets = workbook.Worksheets.ToList();
        
        var throttler = new ProgressThrottler(progress, 200, 1f, _logger);
        
        for (int i = 0; i < worksheets.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            var worksheet = worksheets[i];
            var sourceFileName = Path.GetFileNameWithoutExtension(sourceFilePath);
            if (string.IsNullOrEmpty(sourceFileName))
            {
                sourceFileName = "Document";
            }
            var outputFileName = $"{sourceFileName}_Split_{i + 1:D3}.xlsx";
            var outputPath = Path.Combine(outputDirectory, outputFileName);
            
            await Task.Run(() =>
            {
                using var newWorkbook = new XLWorkbook();
                worksheet.CopyTo(newWorkbook, worksheet.Name);
                newWorkbook.SaveAs(outputPath);
            }, cancellationToken).ConfigureAwait(false);
            
            outputFiles.Add(outputPath);
            
            // 更新进度（节流）
            throttler.Report((float)(i + 1) / worksheets.Count * 100);
        }
        
        throttler.Report(100f, true);
        return outputFiles;
    }

    private async Task<List<string>> SplitByRows(
        XLWorkbook workbook,
        string sourceFilePath,
        string outputDirectory, 
        int rowsPerFile, 
        IProgress<float>? progress, 
        CancellationToken cancellationToken)
    {
        var outputFiles = new List<string>();
        var sourceWorksheet = workbook.Worksheets.First();
        var usedRange = sourceWorksheet.RangeUsed();
        
        if (usedRange == null) return outputFiles;
        
        var header = sourceWorksheet.Row(1);
        int totalRows = usedRange.RowCount() - 1; // 减去表头
        if (totalRows <= 0) return outputFiles;

        int fileCount = (int)Math.Ceiling((double)totalRows / rowsPerFile);
        
        var throttler = new ProgressThrottler(progress, 200, 1f, _logger);
        
        for (int fileIndex = 0; fileIndex < fileCount; fileIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            var startRow = fileIndex * rowsPerFile + 2; // +2 是因为跳过表头行
            var endRow = Math.Min(startRow + rowsPerFile - 1, totalRows + 1);
            
            var sourceFileName = Path.GetFileNameWithoutExtension(sourceFilePath);
            if (string.IsNullOrEmpty(sourceFileName))
            {
                sourceFileName = "Document";
            }
            var outputFileName = $"{sourceFileName}_Split_{fileIndex + 1:D3}.xlsx";
            var outputPath = Path.Combine(outputDirectory, outputFileName);
            
            await Task.Run(() =>
            {
                using var newWorkbook = new XLWorkbook();
                var newWorksheet = newWorkbook.Worksheets.Add("数据");
                
                // 复制表头
                header.CopyTo(newWorksheet.Row(1));

                // 复制数据行
                for (int i = 0; i < (endRow - startRow + 1); i++)
                {
                    var sourceRow = sourceWorksheet.Row(startRow + i);
                    var targetRow = newWorksheet.Row(i + 2);
                    sourceRow.CopyTo(targetRow);
                }
                
                newWorkbook.SaveAs(outputPath);
            }, cancellationToken).ConfigureAwait(false);
            
            outputFiles.Add(outputPath);
            
            // 更新进度（节流）
            throttler.Report((float)(fileIndex + 1) / fileCount * 100);
        }
        
        throttler.Report(100f, true);
        return outputFiles;
    }

    private static async Task<bool> IsHtmlDisguisedFile(string filePath, CancellationToken cancellationToken)
    {
        try
        {
            var firstBytes = new byte[1024];
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var bytesRead = await fileStream.ReadAsync(firstBytes.AsMemory(0, firstBytes.Length), cancellationToken).ConfigureAwait(false);
            
            var content = Encoding.UTF8.GetString(firstBytes, 0, bytesRead).ToLower();
            return content.Contains("<html") || content.Contains("<!doctype html") || content.Contains("<table");
        }
        catch
        {
            return false;
        }
    }

    private async Task<int> ProcessHtmlXlsFile(
        string filePath, 
        IXLWorksheet targetWorksheet, 
        int currentRow,
        bool includeHeaders,
        IProgress<float> progress,
        CancellationToken cancellationToken)
    {
        return await Task.Run(async () =>
        {
            try
            {
                var htmlContent = File.ReadAllText(filePath, Encoding.UTF8);
                var rows = await ParseHtmlTableAsync(htmlContent, cancellationToken).ConfigureAwait(false);
                
                if (rows.Count == 0) return currentRow;
                
                var throttler = new ProgressThrottler(progress, 100, 5f, _logger);
                int startRowIndex = includeHeaders ? 0 : 1;
                int newCurrentRow = currentRow;
                
                for (int i = startRowIndex; i < rows.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    
                    var row = rows[i];
                    for (int col = 0; col < row.Count; col++)
                    {
                        targetWorksheet.Cell(newCurrentRow, col + 1).Value = row[col];
                    }
                    newCurrentRow++;
                    
                    // 报告进度
                    throttler.Report((float)(i - startRowIndex + 1) / (rows.Count - startRowIndex) * 100);
                }
                
                throttler.Report(100f, true);
                return newCurrentRow;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "处理HTML格式XLS文件时发生错误: {FilePath}", filePath);
                throw new InvalidOperationException($"无法处理HTML格式的XLS文件: {ex.Message}", ex);
            }
        }, cancellationToken).ConfigureAwait(false);
    }

    private async Task<List<string>> SplitHtmlXlsFile(
        string filePath, 
        string outputDirectory, 
        int rowsPerFile, 
        IProgress<float>? progress, 
        CancellationToken cancellationToken)
    {
        var outputFiles = new List<string>();
        
        try
        {
            var htmlContent = await _fileSystemService.ReadAllTextAsync(filePath, Encoding.UTF8, cancellationToken).ConfigureAwait(false);
            var rows = await ParseHtmlTableAsync(htmlContent, cancellationToken).ConfigureAwait(false);
            
            if (rows.Count == 0) return outputFiles;
            
            var header = rows.FirstOrDefault();
            var dataRows = rows.Skip(1).ToList();
            
            if (dataRows.Count == 0) return outputFiles;
            
            var throttler = new ProgressThrottler(progress, 200, 1f, _logger);
            var fileIndex = 1;
            
            for (int i = 0; i < dataRows.Count; i += rowsPerFile)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                var endIndex = Math.Min(i + rowsPerFile, dataRows.Count);
                var currentRows = dataRows.Skip(i).Take(endIndex - i).ToList();
                
                // 创建新的Excel文件
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Sheet1");
                
                // 添加表头
                if (header != null)
                {
                    for (int col = 0; col < header.Count; col++)
                    {
                        worksheet.Cell(1, col + 1).Value = header[col];
                    }
                }
                
                // 添加数据行
                for (int row = 0; row < currentRows.Count; row++)
                {
                    var currentRow = currentRows[row];
                    for (int col = 0; col < currentRow.Count; col++)
                    {
                        worksheet.Cell(row + 2, col + 1).Value = currentRow[col];
                    }
                }
                
                // 保存文件
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var outputPath = Path.Combine(outputDirectory, $"{fileName}_Split_{fileIndex:D3}.xlsx");
                
                await Task.Run(() => workbook.SaveAs(outputPath), cancellationToken).ConfigureAwait(false);
                outputFiles.Add(outputPath);
                
                fileIndex++;
                
                // 更新进度
                var progressValue = (float)(endIndex) / dataRows.Count * 100f;
                throttler.Report(progressValue);
            }
            
            // 确保最终进度为100%
            throttler.Report(100f, true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "拆分HTML格式XLS文件时发生错误: {FilePath}", filePath);
            throw;
        }
        
        return outputFiles;
    }

    private async Task<List<List<string>>> ParseHtmlTableAsync(string htmlContent, CancellationToken cancellationToken)
    {
        var rows = new List<List<string>>();
        
        try
        {
            // 异步获取配置
            var config = await _configService.GetConfigurationAsync(cancellationToken).ConfigureAwait(false);
            
            // 检查内容大小限制
            var maxSizeBytes = config.HtmlContentMaxSizeKB * 1024;
            if (htmlContent.Length > maxSizeBytes)
            {
                _logger.LogWarning("HTML内容过大，截断处理: {ActualSize}KB > {MaxSize}KB", 
                    htmlContent.Length / 1024, config.HtmlContentMaxSizeKB);
                htmlContent = htmlContent.Substring(0, maxSizeBytes);
            }
            
            // 使用配置的超时时间
            var timeoutMs = config.HtmlParseTimeoutMs;
            using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(timeoutMs));
            
            if (config.EnableHtmlParseOptimization)
            {
                var parseTask = Task.Run(() => ParseHtmlTableOptimized(htmlContent), cts.Token);
                
                try
                {
                    rows = await parseTask.ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    _logger.LogWarning("HTML表格解析超时({TimeoutMs}ms)，内容长度: {ContentLength}，尝试快速解析", 
                        timeoutMs, htmlContent.Length);
                    // 超时时尝试快速解析
                    rows = ParseHtmlTableFast(htmlContent);
                }
            }
            else
            {
                var parseTask = Task.Run(() => ParseHtmlTableCore(htmlContent), cts.Token);
                
                try
                {
                    rows = await parseTask.ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    _logger.LogWarning("HTML表格解析超时({TimeoutMs}ms)，内容长度: {ContentLength}", 
                        timeoutMs, htmlContent.Length);
                    return rows;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "解析HTML表格时发生错误");
        }
        
        return rows;
    }
    
    /// <summary>
    /// 优化的HTML表格解析（分段处理，减少正则回溯风险）
    /// </summary>
    private List<List<string>> ParseHtmlTableOptimized(string htmlContent)
    {
        var rows = new List<List<string>>();
        
        try
        {
            // 预处理：移除注释和脚本标签，减少干扰
            htmlContent = RemoveHtmlNoise(htmlContent);
            
            // 查找表格边界
            var tableStart = htmlContent.IndexOf("<table", StringComparison.OrdinalIgnoreCase);
            var tableEnd = htmlContent.LastIndexOf("</table>", StringComparison.OrdinalIgnoreCase);
            
            if (tableStart == -1 || tableEnd == -1 || tableEnd <= tableStart)
            {
                _logger.LogDebug("未找到有效的HTML表格结构");
                return rows;
            }
            
            var tableContent = htmlContent.Substring(tableStart, tableEnd - tableStart + 8);
            
            // 分段提取行，避免一次性处理大量数据
            var currentPos = 0;
            var maxRowsPerBatch = 100; // 每批处理的最大行数
            var processedRows = 0;
            
            while (currentPos < tableContent.Length && processedRows < 10000) // 最大行数限制
            {
                var trStart = tableContent.IndexOf("<tr", currentPos, StringComparison.OrdinalIgnoreCase);
                if (trStart == -1) break;
                
                var trEnd = tableContent.IndexOf("</tr>", trStart, StringComparison.OrdinalIgnoreCase);
                if (trEnd == -1) break;
                
                var rowContent = tableContent.Substring(trStart, trEnd - trStart + 5);
                var cells = ExtractCellsOptimized(rowContent);
                
                if (cells.Count > 0)
                {
                    rows.Add(cells);
                    processedRows++;
                }
                
                currentPos = trEnd + 5;
                
                // 每处理一批后检查是否需要中断
                if (processedRows % maxRowsPerBatch == 0)
                {
                    // 给其他任务让出CPU时间
                    Thread.Yield();
                }
            }
            
            if (processedRows >= 10000)
            {
                _logger.LogWarning("HTML表格行数过多，已截断到 {MaxRows} 行", 10000);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "优化HTML解析时发生错误，回退到基础解析");
            return ParseHtmlTableCore(htmlContent);
        }
        
        return rows;
    }
    
    /// <summary>
    /// 快速HTML解析（简化版，用于超时情况）
    /// </summary>
    private List<List<string>> ParseHtmlTableFast(string htmlContent)
    {
        var rows = new List<List<string>>();
        
        try
        {
            // 只处理前面的内容，避免超时
            var maxLength = Math.Min(htmlContent.Length, 50000); // 限制处理长度
            var content = htmlContent.Substring(0, maxLength);
            
            // 简单的行分割
            var lines = content.Split(new[] { "<tr", "</tr>" }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var line in lines.Take(100)) // 最多处理100行
            {
                if (line.Contains("<td") || line.Contains("<th"))
                {
                    var cells = ExtractCellsFast(line);
                    if (cells.Count > 0)
                    {
                        rows.Add(cells);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "快速HTML解析失败");
        }
        
        return rows;
    }
    
    /// <summary>
    /// 移除HTML中的噪声内容
    /// </summary>
    private static string RemoveHtmlNoise(string htmlContent)
    {
        // 移除注释
        var commentStart = 0;
        while ((commentStart = htmlContent.IndexOf("<!--", commentStart)) != -1)
        {
            var commentEnd = htmlContent.IndexOf("-->", commentStart);
            if (commentEnd == -1) break;
            
            htmlContent = htmlContent.Remove(commentStart, commentEnd - commentStart + 3);
        }
        
        // 移除脚本和样式标签
        htmlContent = System.Text.RegularExpressions.Regex.Replace(htmlContent, 
            @"<script[^>]*>.*?</script>", "", 
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
            
        htmlContent = System.Text.RegularExpressions.Regex.Replace(htmlContent, 
            @"<style[^>]*>.*?</style>", "", 
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
        
        return htmlContent;
    }
    
    /// <summary>
    /// 优化的单元格提取
    /// </summary>
    private List<string> ExtractCellsOptimized(string rowContent)
    {
        var cells = new List<string>();
        
        var currentPos = 0;
        while (currentPos < rowContent.Length)
        {
            // 查找下一个单元格开始标签
            var cellStart = -1;
            var tdStart = rowContent.IndexOf("<td", currentPos, StringComparison.OrdinalIgnoreCase);
            var thStart = rowContent.IndexOf("<th", currentPos, StringComparison.OrdinalIgnoreCase);
            
            if (tdStart != -1 && (thStart == -1 || tdStart < thStart))
            {
                cellStart = tdStart;
            }
            else if (thStart != -1)
            {
                cellStart = thStart;
            }
            
            if (cellStart == -1) break;
            
            // 查找单元格内容开始位置
            var contentStart = rowContent.IndexOf('>', cellStart);
            if (contentStart == -1) break;
            contentStart++;
            
            // 查找单元格结束标签
            var cellEnd = rowContent.IndexOf("</t", contentStart, StringComparison.OrdinalIgnoreCase);
            if (cellEnd == -1) break;
            
            // 提取并清理单元格内容
            var cellContent = rowContent.Substring(contentStart, cellEnd - contentStart);
            cellContent = CleanCellContent(cellContent);
            cells.Add(cellContent);
            
            currentPos = cellEnd + 4;
        }
        
        return cells;
    }
    
    /// <summary>
    /// 快速单元格提取
    /// </summary>
    private List<string> ExtractCellsFast(string rowContent)
    {
        var cells = new List<string>();
        
        // 简单的分割方式
        var parts = rowContent.Split(new[] { "<td", "<th", "</td>", "</th>" }, StringSplitOptions.RemoveEmptyEntries);
        
        foreach (var part in parts)
        {
            if (part.Contains(">"))
            {
                var contentStart = part.IndexOf('>') + 1;
                if (contentStart < part.Length)
                {
                    var content = part.Substring(contentStart);
                    content = CleanCellContent(content);
                    if (!string.IsNullOrWhiteSpace(content))
                    {
                        cells.Add(content);
                    }
                }
            }
        }
        
        return cells;
    }
    
    /// <summary>
    /// 清理单元格内容
    /// </summary>
    private static string CleanCellContent(string content)
    {
        // 移除HTML标签
        content = System.Text.RegularExpressions.Regex.Replace(content, @"<[^>]+>", "");
        
        // 解码HTML实体
        content = System.Net.WebUtility.HtmlDecode(content);
        
        // 清理空白字符
        content = content.Trim().Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        
        // 压缩多个空格
        while (content.Contains("  "))
        {
            content = content.Replace("  ", " ");
        }
        
        return content;
    }
    
    private List<List<string>> ParseHtmlTableCore(string htmlContent)
    {
        var rows = new List<List<string>>();
        
        // 简单的HTML表格解析
        var tableStart = htmlContent.IndexOf("<table", StringComparison.OrdinalIgnoreCase);
        var tableEnd = htmlContent.IndexOf("</table>", StringComparison.OrdinalIgnoreCase);
        
        if (tableStart == -1 || tableEnd == -1) return rows;
        
        var tableContent = htmlContent.Substring(tableStart, tableEnd - tableStart + 8);
        
        // 提取所有行
        var trMatches = TrRegex().Matches(tableContent);
        
        foreach (System.Text.RegularExpressions.Match trMatch in trMatches)
        {
            var rowContent = trMatch.Groups[1].Value;
            var cells = new List<string>();
            
            // 提取单元格
            var tdMatches = TdRegex().Matches(rowContent);
            
            foreach (System.Text.RegularExpressions.Match tdMatch in tdMatches)
            {
                var cellContent = tdMatch.Groups[1].Value;
                // 清理HTML标签和解码HTML实体
                cellContent = HtmlTagRegex().Replace(cellContent, "");
                cellContent = System.Net.WebUtility.HtmlDecode(cellContent);
                cellContent = cellContent.Trim();
                cells.Add(cellContent);
            }
            
            if (cells.Count > 0)
            {
                rows.Add(cells);
            }
        }
        
        return rows;
    }

    private static string[] ParseCsvLine(string line)
    {
        var values = new List<string>();
        var current = new StringBuilder();
        bool inQuotes = false;
        
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            
            if (c == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (c == ',' && !inQuotes)
            {
                values.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }
        
        values.Add(current.ToString());
        return [.. values];
    }

    /// <summary>
    /// 尝试从文件处理错误中恢复
    /// </summary>
    private async Task<bool> TryRecoverFromFileError(string filePath, Exception ex, string operationId)
    {
        try
        {
            _logger.LogDebug("[{OperationId}] 尝试恢复文件处理错误: {FilePath}, 异常类型: {ExceptionType}", 
                operationId, filePath, ex.GetType().Name);
            
            // 根据异常类型尝试不同的恢复策略
            switch (ex)
            {
                case UnauthorizedAccessException:
                    // 等待一段时间后重试，可能是文件被临时锁定
                    await Task.Delay(1000).ConfigureAwait(false);
                    return await RetryFileOperation(filePath, operationId).ConfigureAwait(false);
                    
                case IOException when ex.Message.Contains("being used by another process"):
                    // 文件被其他进程占用，等待后重试
                    await Task.Delay(2000).ConfigureAwait(false);
                    return await RetryFileOperation(filePath, operationId).ConfigureAwait(false);
                    
                case OutOfMemoryException:
                    // 内存不足，强制垃圾回收后重试
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    await Task.Delay(500).ConfigureAwait(false);
                    return await RetryFileOperation(filePath, operationId).ConfigureAwait(false);
                    
                case InvalidDataException:
                case NotSupportedException:
                    // 数据格式问题，尝试用不同的方式读取
                    return await TryAlternativeFileReading(filePath, operationId).ConfigureAwait(false);
                    
                default:
                    _logger.LogDebug("[{OperationId}] 无法恢复的异常类型: {ExceptionType}", operationId, ex.GetType().Name);
                    return false;
            }
        }
        catch (Exception recoveryEx)
        {
            _logger.LogWarning(recoveryEx, "[{OperationId}] 错误恢复过程中发生异常: {FilePath}", operationId, filePath);
            return false;
        }
    }
    
    /// <summary>
    /// 重试文件操作
    /// </summary>
    private Task<bool> RetryFileOperation(string filePath, string operationId)
    {
        try
        {
            _logger.LogDebug("[{OperationId}] 重试文件操作: {FilePath}", operationId, filePath);
            
            // 检查文件是否仍然存在
            if (!File.Exists(filePath))
            {
                _logger.LogWarning("[{OperationId}] 重试时发现文件不存在: {FilePath}", operationId, filePath);
                return Task.FromResult(false);
            }
            
            // 尝试简单的文件访问测试
            using (var testStream = File.OpenRead(filePath))
            {
                // 文件可以正常打开，测试通过
            }
            
            _logger.LogDebug("[{OperationId}] 文件访问测试成功: {FilePath}", operationId, filePath);
            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "[{OperationId}] 文件操作重试失败: {FilePath}", operationId, filePath);
            return Task.FromResult(false);
        }
    }
    
    /// <summary>
    /// 尝试替代的文件读取方式
    /// </summary>
    private async Task<bool> TryAlternativeFileReading(string filePath, string operationId)
    {
        try
        {
            _logger.LogDebug("[{OperationId}] 尝试替代文件读取方式: {FilePath}", operationId, filePath);
            
            var extension = Path.GetExtension(filePath).ToLower();
            
            // 对于.xls文件，可能是HTML格式，尝试不同的处理方式
            if (extension == ".xls")
            {
                var fileContent = await File.ReadAllTextAsync(filePath).ConfigureAwait(false);
                if (fileContent.TrimStart().StartsWith("<html", StringComparison.OrdinalIgnoreCase) ||
                    fileContent.TrimStart().StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogDebug("[{OperationId}] 检测到HTML格式的XLS文件: {FilePath}", operationId, filePath);
                    // 这种情况下，我们可以标记为已处理但跳过
                    return true;
                }
            }
            
            // 尝试以只读模式打开文件
            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            
            if (worksheet != null && worksheet.RangeUsed() != null)
            {
                _logger.LogDebug("[{OperationId}] 替代读取方式成功: {FilePath}", operationId, filePath);
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "[{OperationId}] 替代文件读取方式失败: {FilePath}", operationId, filePath);
            return false;
        }
    }
    
    // 生成的正则表达式
    [GeneratedRegex(@"<tr[^>]*>(.*?)</tr>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex TrRegex();
    
    [GeneratedRegex(@"<t[dh][^>]*>(.*?)</t[dh]>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex TdRegex();
    
    [GeneratedRegex(@"<[^>]+>", RegexOptions.IgnoreCase)]
    private static partial Regex HtmlTagRegex();
    
    /// <summary>
    /// 根据配置获取CSV编码
    /// </summary>
    private Encoding GetCsvEncodingFromConfig(AppConfig config)
    {
        return config.CsvEncoding.ToUpperInvariant() switch
        {
            "UTF-8" => config.CsvIncludeBom ? new UTF8Encoding(true) : new UTF8Encoding(false),
            "UTF-16" => Encoding.Unicode,
            "UTF-32" => Encoding.UTF32,
            "ASCII" => Encoding.ASCII,
            "GB2312" => Encoding.GetEncoding("GB2312"),
            "GBK" => Encoding.GetEncoding("GBK"),
            _ => new UTF8Encoding(config.CsvIncludeBom)
        };
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            try
            {
                _semaphore?.Dispose();
                _logger.LogDebug("ExcelProcessingService 资源释放完成");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "ExcelProcessingService 释放资源时发生异常");
            }
            finally
            {
                _disposed = true;
                GC.SuppressFinalize(this);
            }
        }
    }
}