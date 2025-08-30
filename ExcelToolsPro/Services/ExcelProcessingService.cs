using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;
using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services;

/// <summary>
/// Excel处理服务实现
/// </summary>
public partial class ExcelProcessingService : IExcelProcessingService, IDisposable
{
    private readonly ILogger<ExcelProcessingService> _logger;
    private readonly SemaphoreSlim _semaphore;
    private bool _disposed = false;

    public ExcelProcessingService(ILogger<ExcelProcessingService> logger)
    {
        var stopwatch = Stopwatch.StartNew();
        
        _logger = logger;
        _logger.LogDebug("=== ExcelProcessingService 模块初始化开始 ===");
        _logger.LogDebug("注入的日志服务状态: {LoggerStatus}", logger != null ? "有效" : "无效");
        
        // 初始化信号量
        _logger.LogDebug("初始化并发控制信号量...");
        var maxConcurrency = Math.Max(1, Environment.ProcessorCount / 2);
        _semaphore = new SemaphoreSlim(maxConcurrency, maxConcurrency);
        _logger.LogDebug("信号量初始化完成，最大并发数: {MaxConcurrency}", maxConcurrency);
        
        // 检查ClosedXML库可用性
        _logger.LogDebug("检查ClosedXML库可用性...");
        try
        {
            var testWorkbook = new XLWorkbook();
            testWorkbook.Dispose();
            _logger.LogDebug("ClosedXML库检查通过");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "ClosedXML库初始化失败，错误类型: {ExceptionType}", ex.GetType().Name);
            throw new InvalidOperationException("Excel处理库初始化失败", ex);
        }
        
        _logger.LogInformation("ExcelProcessingService 模块初始化完成，耗时: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
    }

    public async Task<ProcessingResult> MergeExcelFilesAsync(
        MergeRequest request, 
        IProgress<float>? progress = null, 
        CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        var operationId = Guid.NewGuid().ToString("N")[..8];
        
        _logger.LogDebug("=== 开始Excel文件合并操作 [ID: {OperationId}] ===", operationId);
        
        // 输入验证
        _logger.LogDebug("[{OperationId}] 开始输入参数验证...", operationId);
        
        if (request == null)
        {
            _logger.LogError("[{OperationId}] 合并请求对象为空", operationId);
            throw new ArgumentNullException(nameof(request));
        }
        
        _logger.LogDebug("[{OperationId}] 请求参数 - 文件数量: {FileCount}, 输出目录: {OutputDirectory}, 添加表头: {AddHeaders}, 表头去重: {DedupeHeaders}", 
            operationId, request.FilePaths?.Length ?? 0, request.OutputDirectory, request.AddHeaders, request.DedupeHeaders);
        
        if (request.FilePaths == null || request.FilePaths.Length == 0)
        {
            _logger.LogWarning("[{OperationId}] 没有选择要合并的文件", operationId);
            return new ProcessingResult
            {
                Success = false,
                Message = "没有选择要合并的文件"
            };
        }
        
        if (string.IsNullOrWhiteSpace(request.OutputDirectory))
        {
            _logger.LogWarning("[{OperationId}] 输出目录为空", operationId);
            return new ProcessingResult
            {
                Success = false,
                Message = "输出目录不能为空"
            };
        }
        
        // 生成输出文件名：源文件名+Merge.xlsx
        var firstFileName = Path.GetFileNameWithoutExtension(request.FilePaths[0]);
        if (string.IsNullOrEmpty(firstFileName))
        {
            firstFileName = "合并文件";
        }
        var outputFileName = $"{firstFileName}_Merge.xlsx";
        var outputPath = Path.Combine(request.OutputDirectory, outputFileName);
        
        _logger.LogDebug("[{OperationId}] 生成输出文件名: {OutputFileName}", operationId, outputFileName);
        
        // 验证输入文件
        _logger.LogDebug("[{OperationId}] 验证输入文件存在性...", operationId);
        var missingFiles = request.FilePaths.Where(f => !File.Exists(f)).ToArray();
        if (missingFiles.Length > 0)
        {
            _logger.LogError("[{OperationId}] 发现 {MissingCount} 个不存在的文件: {MissingFiles}", 
                operationId, missingFiles.Length, string.Join(", ", missingFiles));
            return new ProcessingResult
            {
                Success = false,
                Message = $"以下文件不存在: {string.Join(", ", missingFiles.Select(Path.GetFileName))}"
            };
        }
        
        _logger.LogDebug("[{OperationId}] 等待并发控制信号量...", operationId);
        var semaphoreWaitStart = stopwatch.ElapsedMilliseconds;
        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        _logger.LogDebug("[{OperationId}] 获得并发控制权限，等待耗时: {WaitMs}ms", 
            operationId, stopwatch.ElapsedMilliseconds - semaphoreWaitStart);
        
        try
        {
            _logger.LogInformation("[{OperationId}] 开始合并 {FileCount} 个文件到 {OutputPath}", 
                operationId, request.FilePaths.Length, outputPath);
            
            // 创建工作簿
            _logger.LogDebug("[{OperationId}] 创建Excel工作簿...", operationId);
            var workbookCreateStart = stopwatch.ElapsedMilliseconds;
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("合并数据");
            _logger.LogDebug("[{OperationId}] 工作簿创建完成，耗时: {CreateMs}ms", 
                operationId, stopwatch.ElapsedMilliseconds - workbookCreateStart);
            
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
                var fileProcessStart = stopwatch.ElapsedMilliseconds;
                
                bool shouldLogThisIteration = (i == 0) || (i == request.FilePaths.Length - 1) || (i % logEveryN == 0);
                if (shouldLogThisIteration)
                {
                    _logger.LogDebug("[{OperationId}] 处理文件 {FileIndex}/{TotalFiles}: {FilePath}", 
                        operationId, i + 1, request.FilePaths.Length, filePath);
                }
                
                try
                {
                    var fileInfo = new System.IO.FileInfo(filePath);
                    if (shouldLogThisIteration)
                    {
                        _logger.LogDebug("[{OperationId}] 文件信息 - 大小: {FileSize} bytes, 扩展名: {Extension}", 
                            operationId, fileInfo.Length, fileInfo.Extension);
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
                        _logger.LogDebug("[{OperationId}] 文件处理完成，添加行数: {RowsAdded}, 耗时: {ProcessMs}ms", 
                            operationId, rowsAdded, stopwatch.ElapsedMilliseconds - fileProcessStart);
                    }
                    
                    isFirstFile = false;
                }
                catch (Exception ex)
                {
                    failedFiles.Add(filePath);
                    _logger.LogError(ex, "[{OperationId}] 处理文件 {FilePath} 时发生错误，错误类型: {ExceptionType}, 耗时: {ProcessMs}ms", 
                        operationId, filePath, ex.GetType().Name, stopwatch.ElapsedMilliseconds - fileProcessStart);
                    
                    // 尝试恢复处理
                    if (await TryRecoverFromFileError(filePath, ex, operationId).ConfigureAwait(false))
                    {
                        _logger.LogInformation("[{OperationId}] 文件 {FilePath} 错误恢复成功", operationId, filePath);
                        failedFiles.Remove(filePath); // 从失败列表中移除
                        processedFiles++;
                    }
                    // 继续处理其他文件，不中断整个流程
                }
                
                // 更新主进度 (基于文件)
                mainProgressThrottler.Report((float)(i + 1) / request.FilePaths.Length * 100);
                if (shouldLogThisIteration)
                {
                    _logger.LogDebug("[{OperationId}] 主进度更新(节流): {Progress:F1}%", operationId, (float)(i + 1) / request.FilePaths.Length * 100);
                }
            }
            
            // 确保最终进度为100%
            mainProgressThrottler.Report(100f, true);

            // 确保输出目录存在
            _logger.LogDebug("[{OperationId}] 检查输出目录...", operationId);
            if (!Directory.Exists(request.OutputDirectory))
            {
                _logger.LogDebug("[{OperationId}] 创建输出目录: {OutputDir}", operationId, request.OutputDirectory);
                Directory.CreateDirectory(request.OutputDirectory);
            }
            
            // 保存合并后的文件
            _logger.LogDebug("[{OperationId}] 保存合并文件...", operationId);
            var saveStart = stopwatch.ElapsedMilliseconds;
            await Task.Run(() => workbook.SaveAs(outputPath), cancellationToken).ConfigureAwait(false);
            _logger.LogDebug("[{OperationId}] 文件保存完成，耗时: {SaveMs}ms", 
                operationId, stopwatch.ElapsedMilliseconds - saveStart);
            
            var outputFileInfo = new System.IO.FileInfo(outputPath);
            _logger.LogInformation("[{OperationId}] 文件合并完成 - 输出文件: {OutputPath}, 大小: {FileSize} bytes, 成功: {ProcessedFiles}/{TotalFiles}, 失败: {FailedFiles}, 总耗时: {TotalMs}ms", 
                operationId, outputPath, outputFileInfo.Length, processedFiles, request.FilePaths.Length, failedFiles.Count, stopwatch.ElapsedMilliseconds);
            
            if (failedFiles.Count > 0)
            {
                _logger.LogWarning("[{OperationId}] 部分文件处理失败: {FailedFiles}", 
                    operationId, string.Join(", ", failedFiles.Select(Path.GetFileName)));
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
            _logger.LogInformation("[{OperationId}] 文件合并操作被取消，耗时: {ElapsedMs}ms", 
                operationId, stopwatch.ElapsedMilliseconds);
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[{OperationId}] 合并文件时发生严重错误，错误类型: {ExceptionType}, 耗时: {ElapsedMs}ms", 
                operationId, ex.GetType().Name, stopwatch.ElapsedMilliseconds);
            return new ProcessingResult
            {
                Success = false,
                Message = $"合并文件时发生错误: {ex.Message}"
            };
        }
        finally
        {
            _semaphore.Release();
            _logger.LogDebug("[{OperationId}] 释放并发控制权限", operationId);
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
        // 在后台线程流式读取CSV，避免一次性将整文件加载到内存并阻塞UI线程
        return await Task.Run(() =>
        {
            var newCurrentRow = currentRow;
            
            // 为了计算进度，需要先获取总行数
            long totalLines = 0;
            using (var preReader = new StreamReader(filePath, Encoding.UTF8))
            {
                while (preReader.ReadLine() != null)
                {
                    totalLines++;
                }
            }
            
            var throttler = new ProgressThrottler(progress, 100, 5f, _logger);

            using var reader = new StreamReader(filePath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            string? line;
            bool isFirstLine = true;
            long linesProcessed = 0;
            
            while ((line = reader.ReadLine()) != null)
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
        }, cancellationToken).ConfigureAwait(false);
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
        return await Task.Run(() =>
        {
            try
            {
                var htmlContent = File.ReadAllText(filePath, Encoding.UTF8);
                var rows = ParseHtmlTable(htmlContent);
                
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
            var htmlContent = await File.ReadAllTextAsync(filePath, Encoding.UTF8, cancellationToken).ConfigureAwait(false);
            var rows = ParseHtmlTable(htmlContent);
            
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

    private List<List<string>> ParseHtmlTable(string htmlContent)
    {
        var rows = new List<List<string>>();
        
        try
        {
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
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "解析HTML表格时发生错误");
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
            using var testStream = File.OpenRead(filePath);
            testStream.Close();
            
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
    
    [GeneratedRegex(@"<td[^>]*>(.*?)</td>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex TdRegex();
    
    [GeneratedRegex(@"<[^>]+>")]
    private static partial Regex HtmlTagRegex();

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