using ClosedXML.Excel;
using ExcelToolsPro.Models;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Text;
using System.Diagnostics;

namespace ExcelToolsPro.Services;

/// <summary>
/// 低内存处理器实现
/// </summary>
public class LowMemoryProcessor : ILowMemoryProcessor
{
    private readonly ILogger<LowMemoryProcessor> _logger;
    private readonly IFileSystemService _fileSystemService;
    
    public LowMemoryProcessor(
        ILogger<LowMemoryProcessor> logger,
        IFileSystemService fileSystemService)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _fileSystemService = fileSystemService ?? throw new ArgumentNullException(nameof(fileSystemService));
    }
    
    public bool ShouldUseLowMemoryMode(string[] filePaths, AppConfig config)
    {
        if (config.LargeFileMode)
        {
            _logger.LogDebug("低内存模式已手动启用");
            return true;
        }
        
        var thresholdBytes = config.LargeFileSizeThresholdMB * 1024L * 1024L;
        var totalSize = 0L;
        var largeFileCount = 0;
        
        foreach (var filePath in filePaths)
        {
            try
            {
                var fileInfo = new System.IO.FileInfo(filePath);
                var fileSize = fileInfo.Length;
                totalSize += fileSize;
                
                if (fileSize > thresholdBytes)
                {
                    largeFileCount++;
                    _logger.LogDebug("检测到大文件: {FilePath}, 大小: {FileSizeMB}MB", 
                        filePath, fileSize / 1024.0 / 1024.0);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "无法获取文件大小: {FilePath}", filePath);
            }
        }
        
        var shouldUse = largeFileCount > 0 || totalSize > thresholdBytes * 2;
        
        if (shouldUse)
        {
            _logger.LogInformation("自动启用低内存模式 - 大文件数: {LargeFileCount}, 总大小: {TotalSizeMB}MB, 阈值: {ThresholdMB}MB",
                largeFileCount, totalSize / 1024.0 / 1024.0, config.LargeFileSizeThresholdMB);
        }
        
        return shouldUse;
    }
    
    public async Task<ProcessingResult> ProcessCsvToExcelLowMemoryAsync(
        string inputPath, 
        string outputPath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default)
    {
        using var timer = PerformanceTimerExtensions.CreateTimer(_logger, "ProcessCsvToExcelLowMemory", new { InputPath = inputPath });
        
        try
        {
            _logger.LogInformation("开始低内存模式CSV转Excel: {InputPath} -> {OutputPath}", inputPath, outputPath);
            
            var bufferSize = config.IOBufferSizeKB * 1024;
            var encoding = GetEncodingFromConfig(config);
            var chunkSize = Math.Max(config.ChunkSizeMB * 1024 * 1024, 1024 * 1024); // 最小1MB
            
            // 计算总行数用于进度报告
            timer.Checkpoint("计算总行数");
            var totalLines = await CountLinesAsync(inputPath, encoding, bufferSize, cancellationToken);
            _logger.LogDebug("CSV文件总行数: {TotalLines}", totalLines);
            
            var throttler = new ProgressThrottler(progress, config.ProgressThrottleMs, 5f, _logger);
            
            timer.Checkpoint("开始流式处理");
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("数据");
            
            using var fileStream = await _fileSystemService.CreateFileStreamAsync(
                inputPath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, cancellationToken);
            using var reader = new StreamReader(fileStream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize);
            
            var currentRow = 1;
            var linesProcessed = 0L;
            var batchSize = 1000; // 每批处理1000行
            var batch = new List<string[]>();
            
            string? line;
            while ((line = await reader.ReadLineAsync()) != null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                var values = ParseCsvLine(line);
                batch.Add(values);
                linesProcessed++;
                
                // 批量写入Excel
                if (batch.Count >= batchSize)
                {
                    WriteBatchToWorksheet(worksheet, batch, currentRow);
                    currentRow += batch.Count;
                    batch.Clear();
                    
                    // 报告进度
                    throttler.Report((float)linesProcessed / totalLines * 100);
                    
                    // 让出CPU时间
                    if (linesProcessed % (batchSize * 10) == 0)
                    {
                        await Task.Yield();
                    }
                }
            }
            
            // 处理剩余的批次
            if (batch.Count > 0)
            {
                WriteBatchToWorksheet(worksheet, batch, currentRow);
                batch.Clear();
            }
            
            timer.Checkpoint("保存Excel文件");
            await Task.Run(() => workbook.SaveAs(outputPath), cancellationToken);
            
            throttler.Report(100f, true);
            
            _logger.LogInformation("低内存模式CSV转Excel完成，处理行数: {LinesProcessed}, 耗时: {ElapsedMs}ms", 
                linesProcessed, timer.ElapsedMilliseconds);
            
            return new ProcessingResult
            {
                Success = true,
                Message = $"成功转换 {linesProcessed} 行数据"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "低内存模式CSV转Excel失败: {InputPath}", inputPath);
            return new ProcessingResult
            {
                Success = false,
                Message = $"转换失败: {ex.Message}"
            };
        }
    }
    
    public async Task<ProcessingResult> ProcessExcelToCsvLowMemoryAsync(
        string inputPath, 
        string outputPath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default)
    {
        using var timer = PerformanceTimerExtensions.CreateTimer(_logger, "ProcessExcelToCsvLowMemory", new { InputPath = inputPath });
        
        try
        {
            _logger.LogInformation("开始低内存模式Excel转CSV: {InputPath} -> {OutputPath}", inputPath, outputPath);
            
            var bufferSize = config.IOBufferSizeKB * 1024;
            var encoding = GetEncodingFromConfig(config);
            var batchSize = 1000; // 每批处理1000行
            
            timer.Checkpoint("打开Excel文件");
            using var workbook = new XLWorkbook(inputPath);
            var worksheet = workbook.Worksheets.First();
            var usedRange = worksheet.RangeUsed();
            
            if (usedRange == null)
            {
                return new ProcessingResult
                {
                    Success = false,
                    Message = "Excel文件中没有数据"
                };
            }
            
            var totalRows = usedRange.RowCount();
            _logger.LogDebug("Excel文件总行数: {TotalRows}", totalRows);
            
            var throttler = new ProgressThrottler(progress, config.ProgressThrottleMs, 5f, _logger);
            
            timer.Checkpoint("开始流式写入CSV");
            using var fileStream = await _fileSystemService.CreateFileStreamAsync(
                outputPath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize, cancellationToken);
            using var writer = new StreamWriter(fileStream, encoding, bufferSize);
            
            var processedRows = 0;
            
            // 分批处理行数据
            for (int startRow = 1; startRow <= totalRows; startRow += batchSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                var endRow = Math.Min(startRow + batchSize - 1, totalRows);
                var batch = new List<string>();
                
                // 读取一批行数据
                for (int row = startRow; row <= endRow; row++)
                {
                    var rowData = new List<string>();
                    var worksheetRow = worksheet.Row(row);
                    
                    // 获取行中的所有单元格值
                    for (int col = 1; col <= usedRange.ColumnCount(); col++)
                    {
                        var cell = worksheetRow.Cell(col);
                        var value = cell.GetValue<string>() ?? string.Empty;
                        rowData.Add(EscapeCsvValue(value));
                    }
                    
                    batch.Add(string.Join(",", rowData));
                }
                
                // 批量写入CSV
                foreach (var line in batch)
                {
                    await writer.WriteLineAsync(line);
                }
                
                processedRows += batch.Count;
                
                // 报告进度
                throttler.Report((float)processedRows / totalRows * 100);
                
                // 让出CPU时间
                if (processedRows % (batchSize * 10) == 0)
                {
                    await Task.Yield();
                }
            }
            
            await writer.FlushAsync();
            throttler.Report(100f, true);
            
            _logger.LogInformation("低内存模式Excel转CSV完成，处理行数: {ProcessedRows}, 耗时: {ElapsedMs}ms", 
                processedRows, timer.ElapsedMilliseconds);
            
            return new ProcessingResult
            {
                Success = true,
                Message = $"成功转换 {processedRows} 行数据"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "低内存模式Excel转CSV失败: {InputPath}", inputPath);
            return new ProcessingResult
            {
                Success = false,
                Message = $"转换失败: {ex.Message}"
            };
        }
    }
    
    public async Task<ProcessingResult> ProcessHtmlToExcelLowMemoryAsync(
        string inputPath, 
        string outputPath, 
        AppConfig config,
        IProgress<float>? progress = null,
        CancellationToken cancellationToken = default)
    {
        using var timer = PerformanceTimerExtensions.CreateTimer(_logger, "ProcessHtmlToExcelLowMemory", new { InputPath = inputPath });
        
        try
        {
            _logger.LogInformation("开始低内存模式HTML转Excel: {InputPath} -> {OutputPath}", inputPath, outputPath);
            
            var bufferSize = config.IOBufferSizeKB * 1024;
            var maxContentSize = config.HtmlContentMaxSizeKB * 1024;
            
            timer.Checkpoint("读取HTML内容");
            string htmlContent;
            using (var fileStream = await _fileSystemService.CreateFileStreamAsync(
                inputPath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, cancellationToken))
            using (var reader = new StreamReader(fileStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize))
            {
                var content = await reader.ReadToEndAsync();
                
                // 限制内容大小
                if (content.Length > maxContentSize)
                {
                    _logger.LogWarning("HTML内容过大，截断处理: {ActualSize}KB > {MaxSize}KB", 
                        content.Length / 1024, config.HtmlContentMaxSizeKB);
                    content = content.Substring(0, maxContentSize);
                }
                
                htmlContent = content;
            }
            
            timer.Checkpoint("解析HTML表格");
            var rows = await ParseHtmlTableStreamAsync(htmlContent, config, cancellationToken);
            
            if (rows.Count == 0)
            {
                return new ProcessingResult
                {
                    Success = false,
                    Message = "HTML文件中没有找到有效的表格数据"
                };
            }
            
            timer.Checkpoint("创建Excel文件");
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("数据");
            
            var throttler = new ProgressThrottler(progress, config.ProgressThrottleMs, 5f, _logger);
            var batchSize = 500; // HTML表格行数通常较少，使用较小的批次
            
            // 分批写入Excel
            for (int i = 0; i < rows.Count; i += batchSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                var endIndex = Math.Min(i + batchSize, rows.Count);
                var batch = rows.Skip(i).Take(endIndex - i).ToList();
                
                WriteBatchToWorksheet(worksheet, batch.Select(r => r.ToArray()).ToList(), i + 1);
                
                // 报告进度
                throttler.Report((float)endIndex / rows.Count * 100);
                
                // 让出CPU时间
                if (i % (batchSize * 4) == 0)
                {
                    await Task.Yield();
                }
            }
            
            timer.Checkpoint("保存Excel文件");
            await Task.Run(() => workbook.SaveAs(outputPath), cancellationToken);
            
            throttler.Report(100f, true);
            
            _logger.LogInformation("低内存模式HTML转Excel完成，处理行数: {RowCount}, 耗时: {ElapsedMs}ms", 
                rows.Count, timer.ElapsedMilliseconds);
            
            return new ProcessingResult
            {
                Success = true,
                Message = $"成功转换 {rows.Count} 行数据"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "低内存模式HTML转Excel失败: {InputPath}", inputPath);
            return new ProcessingResult
            {
                Success = false,
                Message = $"转换失败: {ex.Message}"
            };
        }
    }
    
    #region 私有辅助方法
    
    private async Task<long> CountLinesAsync(string filePath, Encoding encoding, int bufferSize, CancellationToken cancellationToken)
    {
        var lineCount = 0L;
        using var fileStream = await _fileSystemService.CreateFileStreamAsync(
            filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, cancellationToken);
        using var reader = new StreamReader(fileStream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize);
        
        while (await reader.ReadLineAsync() != null)
        {
            lineCount++;
            
            // 每1000行检查一次取消
            if (lineCount % 1000 == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }
        }
        
        return lineCount;
    }
    
    private void WriteBatchToWorksheet(IXLWorksheet worksheet, List<string[]> batch, int startRow)
    {
        for (int i = 0; i < batch.Count; i++)
        {
            var row = batch[i];
            for (int col = 0; col < row.Length; col++)
            {
                worksheet.Cell(startRow + i, col + 1).Value = row[col];
            }
        }
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
        return values.ToArray();
    }
    
    private static string EscapeCsvValue(string value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;
        
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\""; // 转义双引号
        }
        
        return value;
    }
    
    private async Task<List<List<string>>> ParseHtmlTableStreamAsync(string htmlContent, AppConfig config, CancellationToken cancellationToken)
    {
        var rows = new List<List<string>>();
        
        try
        {
            // 使用配置的超时时间
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(TimeSpan.FromMilliseconds(config.HtmlParseTimeoutMs));
            
            var parseTask = Task.Run(() => ParseHtmlTableOptimized(htmlContent), cts.Token);
            rows = await parseTask;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            _logger.LogWarning("HTML表格解析超时({TimeoutMs}ms)，使用快速解析", config.HtmlParseTimeoutMs);
            rows = ParseHtmlTableFast(htmlContent);
        }
        
        return rows;
    }
    
    private List<List<string>> ParseHtmlTableOptimized(string htmlContent)
    {
        var rows = new List<List<string>>();
        
        // 预处理：移除注释和脚本标签
        htmlContent = RemoveHtmlNoise(htmlContent);
        
        // 查找表格边界
        var tableStart = htmlContent.IndexOf("<table", StringComparison.OrdinalIgnoreCase);
        var tableEnd = htmlContent.LastIndexOf("</table>", StringComparison.OrdinalIgnoreCase);
        
        if (tableStart == -1 || tableEnd == -1 || tableEnd <= tableStart)
        {
            return rows;
        }
        
        var tableContent = htmlContent.Substring(tableStart, tableEnd - tableStart + 8);
        
        // 分段提取行
        var currentPos = 0;
        var processedRows = 0;
        
        while (currentPos < tableContent.Length && processedRows < 10000)
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
            
            // 每处理100行让出CPU时间
            if (processedRows % 100 == 0)
            {
                Thread.Yield();
            }
        }
        
        return rows;
    }
    
    private List<List<string>> ParseHtmlTableFast(string htmlContent)
    {
        var rows = new List<List<string>>();
        
        // 简化的快速解析
        var maxLength = Math.Min(htmlContent.Length, 50000);
        var content = htmlContent.Substring(0, maxLength);
        
        var lines = content.Split(new[] { "<tr", "</tr>" }, StringSplitOptions.RemoveEmptyEntries);
        
        foreach (var line in lines.Take(100))
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
        
        return rows;
    }
    
    private static string RemoveHtmlNoise(string htmlContent)
    {
        // 移除注释
        htmlContent = System.Text.RegularExpressions.Regex.Replace(htmlContent, "<!--.*?-->", "", 
            System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.Compiled);
        
        // 移除脚本和样式标签
        htmlContent = System.Text.RegularExpressions.Regex.Replace(htmlContent, "<script.*?</script>", "", 
            System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
        htmlContent = System.Text.RegularExpressions.Regex.Replace(htmlContent, "<style.*?</style>", "", 
            System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
        
        return htmlContent;
    }
    
    private List<string> ExtractCellsOptimized(string rowContent)
    {
        var cells = new List<string>();
        var currentPos = 0;
        
        while (currentPos < rowContent.Length)
        {
            var tdStart = FindNextCellStart(rowContent, currentPos);
            if (tdStart == -1) break;
            
            var contentStart = rowContent.IndexOf('>', tdStart) + 1;
            if (contentStart == 0) break;
            
            var tdEnd = FindCellEnd(rowContent, contentStart);
            if (tdEnd == -1) break;
            
            var cellContent = rowContent.Substring(contentStart, tdEnd - contentStart);
            var cleanContent = CleanHtmlContent(cellContent);
            cells.Add(cleanContent);
            
            currentPos = tdEnd + 1;
        }
        
        return cells;
    }
    
    private List<string> ExtractCellsFast(string rowContent)
    {
        var cells = new List<string>();
        
        // 简化的单元格提取
        var cellTags = new[] { "<td", "<th" };
        
        foreach (var tag in cellTags)
        {
            var parts = rowContent.Split(new[] { tag }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts.Skip(1).Take(50)) // 最多50个单元格
            {
                var contentStart = part.IndexOf('>') + 1;
                if (contentStart > 0)
                {
                    var endTag = part.IndexOf("</t", contentStart);
                    if (endTag > contentStart)
                    {
                        var content = part.Substring(contentStart, endTag - contentStart);
                        cells.Add(CleanHtmlContent(content));
                    }
                }
            }
        }
        
        return cells;
    }
    
    private static int FindNextCellStart(string content, int startPos)
    {
        var tdPos = content.IndexOf("<td", startPos, StringComparison.OrdinalIgnoreCase);
        var thPos = content.IndexOf("<th", startPos, StringComparison.OrdinalIgnoreCase);
        
        if (tdPos == -1) return thPos;
        if (thPos == -1) return tdPos;
        return Math.Min(tdPos, thPos);
    }
    
    private static int FindCellEnd(string content, int startPos)
    {
        var tdEnd = content.IndexOf("</td>", startPos, StringComparison.OrdinalIgnoreCase);
        var thEnd = content.IndexOf("</th>", startPos, StringComparison.OrdinalIgnoreCase);
        
        if (tdEnd == -1) return thEnd;
        if (thEnd == -1) return tdEnd;
        return Math.Min(tdEnd, thEnd);
    }
    
    private static string CleanHtmlContent(string content)
    {
        if (string.IsNullOrEmpty(content))
            return string.Empty;
        
        // 移除HTML标签
        content = System.Text.RegularExpressions.Regex.Replace(content, "<.*?>", "", 
            System.Text.RegularExpressions.RegexOptions.Compiled);
        
        // 解码HTML实体
        content = System.Net.WebUtility.HtmlDecode(content);
        
        // 清理空白字符
        content = content.Trim().Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        while (content.Contains("  "))
        {
            content = content.Replace("  ", " ");
        }
        
        return content;
    }
    
    private static Encoding GetEncodingFromConfig(AppConfig config)
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
    
    #endregion
}