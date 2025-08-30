using Microsoft.Extensions.Logging;
using System.IO;

namespace ExcelToolsPro.Services;

/// <summary>
/// 文件系统服务实现
/// </summary>
public class FileSystemService : IFileSystemService
{
    private readonly ILogger<FileSystemService> _logger;
    private readonly string _tempDirectory;

    public FileSystemService(ILogger<FileSystemService> logger)
    {
        _logger = logger;
        
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
            var fileInfo = new FileInfo(filePath);
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
                    var fileInfo = new FileInfo(file);
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
}