using Microsoft.Extensions.Logging;
using System.IO;
using System.Text;
using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services;

/// <summary>
/// 文件系统服务实现
/// </summary>
public class FileSystemService : IFileSystemService
{
    private readonly ILogger<FileSystemService> _logger;
    private readonly string _tempDirectory;
    private readonly IConfigurationService _configService;

    public FileSystemService(ILogger<FileSystemService> logger, IConfigurationService configService)
    {
        _logger = logger;
        _configService = configService;
        
        // 创建应用程序专用的临时目录
        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        _tempDirectory = Path.Combine(appDataPath, "ExcelToolsPro", "Temp");
        
        try
        {
            Directory.CreateDirectory(_tempDirectory);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "创建临时目录失败: {TempDirectory}", _tempDirectory);
        }
    }

    public bool FileExists(string filePath)
    {
        try
        {
            return File.Exists(filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "检查文件是否存在时发生错误: {FilePath}", filePath);
            return false;
        }
    }

    public bool DirectoryExists(string directoryPath)
    {
        try
        {
            return Directory.Exists(directoryPath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "检查目录是否存在时发生错误: {DirectoryPath}", directoryPath);
            return false;
        }
    }

    public void CreateDirectory(string directoryPath)
    {
        try
        {
            Directory.CreateDirectory(directoryPath);
            _logger.LogInformation("目录创建成功: {DirectoryPath}", directoryPath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "创建目录时发生错误: {DirectoryPath}", directoryPath);
            throw;
        }
    }

    public void DeleteFile(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                _logger.LogInformation("文件删除成功: {FilePath}", filePath);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "删除文件时发生错误: {FilePath}", filePath);
            throw;
        }
    }

    public void DeleteDirectory(string directoryPath, bool recursive = false)
    {
        try
        {
            if (Directory.Exists(directoryPath))
            {
                Directory.Delete(directoryPath, recursive);
                _logger.LogInformation("目录删除成功: {DirectoryPath}", directoryPath);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "删除目录时发生错误: {DirectoryPath}", directoryPath);
            throw;
        }
    }

    public long GetFileSize(string filePath)
    {
        try
        {
            var fileInfo = new System.IO.FileInfo(filePath);
            return fileInfo.Length;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "获取文件大小时发生错误: {FilePath}", filePath);
            return 0;
        }
    }

    public DateTime GetLastWriteTime(string filePath)
    {
        try
        {
            return File.GetLastWriteTime(filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "获取文件最后修改时间时发生错误: {FilePath}", filePath);
            return DateTime.MinValue;
        }
    }

    public void CopyFile(string sourceFilePath, string destinationFilePath, bool overwrite = false)
    {
        try
        {
            File.Copy(sourceFilePath, destinationFilePath, overwrite);
            _logger.LogInformation("文件复制成功: {Source} -> {Destination}", sourceFilePath, destinationFilePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "复制文件时发生错误: {Source} -> {Destination}", sourceFilePath, destinationFilePath);
            throw;
        }
    }

    public void MoveFile(string sourceFilePath, string destinationFilePath)
    {
        try
        {
            File.Move(sourceFilePath, destinationFilePath);
            _logger.LogInformation("文件移动成功: {Source} -> {Destination}", sourceFilePath, destinationFilePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "移动文件时发生错误: {Source} -> {Destination}", sourceFilePath, destinationFilePath);
            throw;
        }
    }

    public string GetTempFilePath(string extension = ".tmp")
    {
        try
        {
            var fileName = $"{Guid.NewGuid()}{extension}";
            return Path.Combine(_tempDirectory, fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "生成临时文件路径时发生错误");
            return Path.GetTempFileName();
        }
    }

    public async Task CleanupTempFilesAsync(TimeSpan olderThan)
    {
        try
        {
            if (!Directory.Exists(_tempDirectory))
                return;

            var cutoffTime = DateTime.Now - olderThan;
            var tempFiles = Directory.GetFiles(_tempDirectory, "*", SearchOption.AllDirectories);
            
            var deletedCount = 0;
            var totalSize = 0L;

            foreach (var file in tempFiles)
            {
                try
                {
                    var fileInfo = new System.IO.FileInfo(file);
                    if (fileInfo.LastWriteTime < cutoffTime)
                    {
                        totalSize += fileInfo.Length;
                        await Task.Run(() => File.Delete(file)).ConfigureAwait(false);
                        deletedCount++;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "删除临时文件时发生错误: {FilePath}", file);
                }
            }

            if (deletedCount > 0)
            {
                _logger.LogInformation("临时文件清理完成，删除 {Count} 个文件，释放 {Size} MB 空间", 
                    deletedCount, totalSize / (1024.0 * 1024.0));
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "清理临时文件时发生错误");
        }
    }

    public async Task CopyFileAsync(string sourceFilePath, string destinationFilePath, bool overwrite = false, CancellationToken cancellationToken = default)
    {
        try
        {
            var config = await _configService.GetConfigurationAsync().ConfigureAwait(false);
            
            if (!config.UseAsyncIO)
            {
                await Task.Run(() => File.Copy(sourceFilePath, destinationFilePath, overwrite), cancellationToken).ConfigureAwait(false);
            }
            else
            {
                var bufferSize = config.IOBufferSizeKB * 1024;
                using var sourceStream = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, FileOptions.SequentialScan);
                using var destinationStream = new FileStream(destinationFilePath, overwrite ? FileMode.Create : FileMode.CreateNew, FileAccess.Write, FileShare.None, bufferSize, FileOptions.SequentialScan);
                await sourceStream.CopyToAsync(destinationStream, bufferSize, cancellationToken).ConfigureAwait(false);
            }
            
            _logger.LogInformation("文件异步复制成功: {Source} -> {Destination}", sourceFilePath, destinationFilePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "异步复制文件时发生错误: {Source} -> {Destination}", sourceFilePath, destinationFilePath);
            throw;
        }
    }

    public async Task MoveFileAsync(string sourceFilePath, string destinationFilePath, CancellationToken cancellationToken = default)
    {
        try
        {
            var config = await _configService.GetConfigurationAsync().ConfigureAwait(false);
            
            if (!config.UseAsyncIO)
            {
                await Task.Run(() => File.Move(sourceFilePath, destinationFilePath), cancellationToken).ConfigureAwait(false);
            }
            else
            {
                // 对于跨驱动器移动，需要先复制再删除
                var sourceInfo = new System.IO.FileInfo(sourceFilePath);
                var destInfo = new System.IO.FileInfo(destinationFilePath);
                
                if (sourceInfo.Directory?.Root.FullName == destInfo.Directory?.Root.FullName)
                {
                    // 同驱动器，直接移动
                    await Task.Run(() => File.Move(sourceFilePath, destinationFilePath), cancellationToken).ConfigureAwait(false);
                }
                else
                {
                    // 跨驱动器，复制后删除
                    await CopyFileAsync(sourceFilePath, destinationFilePath, false, cancellationToken).ConfigureAwait(false);
                    await Task.Run(() => File.Delete(sourceFilePath), cancellationToken).ConfigureAwait(false);
                }
            }
            
            _logger.LogInformation("文件异步移动成功: {Source} -> {Destination}", sourceFilePath, destinationFilePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "异步移动文件时发生错误: {Source} -> {Destination}", sourceFilePath, destinationFilePath);
            throw;
        }
    }

    public async Task<string> ReadAllTextAsync(string filePath, Encoding? encoding = null, CancellationToken cancellationToken = default)
    {
        try
        {
            var config = await _configService.GetConfigurationAsync().ConfigureAwait(false);
            encoding ??= GetEncodingFromConfig(config);
            
            if (!config.UseAsyncIO)
            {
                return await Task.Run(() => File.ReadAllText(filePath, encoding), cancellationToken).ConfigureAwait(false);
            }
            
            var bufferSize = config.IOBufferSizeKB * 1024;
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, FileOptions.SequentialScan);
            using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize);
            return await reader.ReadToEndAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "异步读取文件时发生错误: {FilePath}", filePath);
            throw;
        }
    }

    public async Task WriteAllTextAsync(string filePath, string content, Encoding? encoding = null, CancellationToken cancellationToken = default)
    {
        try
        {
            var config = await _configService.GetConfigurationAsync().ConfigureAwait(false);
            encoding ??= GetEncodingFromConfig(config);
            
            if (!config.UseAsyncIO)
            {
                await Task.Run(() => File.WriteAllText(filePath, content, encoding), cancellationToken).ConfigureAwait(false);
                return;
            }
            
            var bufferSize = config.IOBufferSizeKB * 1024;
            using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize, FileOptions.SequentialScan);
            using var writer = new StreamWriter(stream, encoding, bufferSize);
            await writer.WriteAsync(content).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "异步写入文件时发生错误: {FilePath}", filePath);
            throw;
        }
    }

    public async Task<byte[]> ReadAllBytesAsync(string filePath, CancellationToken cancellationToken = default)
    {
        try
        {
            var config = await _configService.GetConfigurationAsync().ConfigureAwait(false);
            
            if (!config.UseAsyncIO)
            {
                return await Task.Run(() => File.ReadAllBytes(filePath), cancellationToken).ConfigureAwait(false);
            }
            
            var bufferSize = config.IOBufferSizeKB * 1024;
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, FileOptions.SequentialScan);
            using var memoryStream = new MemoryStream();
            await stream.CopyToAsync(memoryStream, bufferSize, cancellationToken).ConfigureAwait(false);
            return memoryStream.ToArray();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "异步读取文件字节时发生错误: {FilePath}", filePath);
            throw;
        }
    }

    public async Task WriteAllBytesAsync(string filePath, byte[] bytes, CancellationToken cancellationToken = default)
    {
        try
        {
            var config = await _configService.GetConfigurationAsync().ConfigureAwait(false);
            
            if (!config.UseAsyncIO)
            {
                await Task.Run(() => File.WriteAllBytes(filePath, bytes), cancellationToken).ConfigureAwait(false);
                return;
            }
            
            var bufferSize = config.IOBufferSizeKB * 1024;
            using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize, FileOptions.SequentialScan);
            await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "异步写入文件字节时发生错误: {FilePath}", filePath);
            throw;
        }
    }

    public async Task<FileStream> CreateFileStreamAsync(string filePath, FileMode mode, FileAccess access, FileShare share = FileShare.Read, int bufferSize = 4096, CancellationToken cancellationToken = default)
    {
        try
        {
            var config = await _configService.GetConfigurationAsync().ConfigureAwait(false);
            var actualBufferSize = bufferSize > 0 ? bufferSize : config.IOBufferSizeKB * 1024;
            return await Task.Run(() => new FileStream(filePath, mode, access, share, actualBufferSize, FileOptions.SequentialScan), cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "创建异步文件流时发生错误: {FilePath}", filePath);
            throw;
        }
    }

    /// <summary>
    /// 根据配置获取编码
    /// </summary>
    private Encoding GetEncodingFromConfig(AppConfig config)
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
}