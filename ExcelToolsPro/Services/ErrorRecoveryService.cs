using Microsoft.Extensions.Logging;
using System.IO;
using System.Text.Json;
using ExcelToolsPro.Models;

namespace ExcelToolsPro.Services;

/// <summary>
/// 错误恢复服务实现
/// </summary>
public class ErrorRecoveryService : IErrorRecoveryService
{
    private readonly ILogger<ErrorRecoveryService> _logger;
    private readonly IFileSystemService _fileSystemService;
    private readonly string _checkpointDirectory;

    public ErrorRecoveryService(
        ILogger<ErrorRecoveryService> logger,
        IFileSystemService fileSystemService)
    {
        _logger = logger;
        _fileSystemService = fileSystemService;
        
        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        _checkpointDirectory = Path.Combine(appDataPath, "ExcelToolsPro", "Checkpoints");
        
        try
        {
            Directory.CreateDirectory(_checkpointDirectory);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "创建检查点目录失败: {CheckpointDirectory}", _checkpointDirectory);
        }
    }

    public async Task<RetryResult> RetryFailedOperationAsync(string taskId, RetryOptions? options = null)
    {
        options ??= new RetryOptions();
        
        try
        {
            _logger.LogInformation("开始重试失败的操作，任务ID: {TaskId}", taskId);
            
            // 加载检查点
            var checkpointInfo = await GetProcessingCheckpointAsync(taskId).ConfigureAwait(false);
            if (!checkpointInfo.CanResume || checkpointInfo.Checkpoint == null)
            {
                return new RetryResult
                {
                    Success = false,
                    Message = "无法找到有效的检查点或任务无法恢复"
                };
            }
            
            var checkpoint = checkpointInfo.Checkpoint;
            
            // 生成新的任务ID
            var newTaskId = Guid.NewGuid().ToString();
            
            // 创建新的检查点
            var newCheckpoint = new ProcessingCheckpoint
            {
                TaskId = newTaskId,
                CompletedFiles = new List<string>(checkpoint.CompletedFiles),
                CurrentFile = checkpoint.CurrentFile,
                ProgressPercent = checkpoint.ProgressPercent,
                Timestamp = DateTime.UtcNow,
                TempFiles = new List<string>()
            };
            
            await SaveCheckpointAsync(newCheckpoint).ConfigureAwait(false);
            
            _logger.LogInformation("重试操作准备完成，新任务ID: {NewTaskId}", newTaskId);
            
            return new RetryResult
            {
                Success = true,
                NewTaskId = newTaskId,
                Message = "重试操作已准备就绪"
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "重试失败操作时发生错误，任务ID: {TaskId}", taskId);
            return new RetryResult
            {
                Success = false,
                Message = $"重试操作时发生错误: {ex.Message}"
            };
        }
    }

    public async Task<CheckpointInfo> GetProcessingCheckpointAsync(string taskId)
    {
        try
        {
            var checkpointPath = GetCheckpointPath(taskId);
            
            if (!_fileSystemService.FileExists(checkpointPath))
            {
                return new CheckpointInfo
                {
                    Checkpoint = null,
                    CanResume = false
                };
            }
            
            var json = await File.ReadAllTextAsync(checkpointPath).ConfigureAwait(false);
            var checkpoint = JsonSerializer.Deserialize<ProcessingCheckpoint>(json);
            
            if (checkpoint == null)
            {
                return new CheckpointInfo
                {
                    Checkpoint = null,
                    CanResume = false
                };
            }
            
            // 检查检查点是否有效（例如，检查临时文件是否仍然存在）
            var canResume = ValidateCheckpoint(checkpoint);
            
            return new CheckpointInfo
            {
                Checkpoint = checkpoint,
                CanResume = canResume
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "获取处理检查点时发生错误，任务ID: {TaskId}", taskId);
            return new CheckpointInfo
            {
                Checkpoint = null,
                CanResume = false
            };
        }
    }

    public async Task<CleanupResult> CleanupTempFilesAsync(string? taskId = null)
    {
        var cleanedFiles = new List<string>();
        long totalSize = 0;
        
        try
        {
            if (!string.IsNullOrEmpty(taskId))
            {
                // 清理特定任务的临时文件
                var checkpointInfo = await GetProcessingCheckpointAsync(taskId).ConfigureAwait(false);
                if (checkpointInfo.Checkpoint != null)
                {
                    foreach (var tempFile in checkpointInfo.Checkpoint.TempFiles)
                    {
                        if (_fileSystemService.FileExists(tempFile))
                        {
                            try
                            {
                                var size = _fileSystemService.GetFileSize(tempFile);
                                _fileSystemService.DeleteFile(tempFile);
                                cleanedFiles.Add(tempFile);
                                totalSize += size;
                            }
                            catch (Exception ex)
                            {
                                _logger.LogWarning(ex, "删除临时文件失败: {TempFile}", tempFile);
                            }
                        }
                    }
                    
                    // 删除检查点文件
                    var checkpointPath = GetCheckpointPath(taskId);
                    if (_fileSystemService.FileExists(checkpointPath))
                    {
                        _fileSystemService.DeleteFile(checkpointPath);
                    }
                }
            }
            else
            {
                // 清理所有过期的检查点和临时文件
                await CleanupExpiredCheckpoints(TimeSpan.FromDays(7)).ConfigureAwait(false);
                
                // 清理应用程序临时文件
                await _fileSystemService.CleanupTempFilesAsync(TimeSpan.FromHours(24)).ConfigureAwait(false);
            }
            
            _logger.LogInformation("临时文件清理完成，删除 {Count} 个文件，释放 {Size} MB 空间", 
                cleanedFiles.Count, totalSize / (1024.0 * 1024.0));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "清理临时文件时发生错误");
        }
        
        return new CleanupResult
        {
            CleanedFiles = cleanedFiles.ToArray(),
            FreedSpaceMB = totalSize / (1024 * 1024)
        };
    }

    /// <summary>
    /// 保存检查点
    /// </summary>
    public async Task SaveCheckpointAsync(ProcessingCheckpoint checkpoint)
    {
        try
        {
            var checkpointPath = GetCheckpointPath(checkpoint.TaskId);
            
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };
            
            var json = JsonSerializer.Serialize(checkpoint, options);
            await File.WriteAllTextAsync(checkpointPath, json).ConfigureAwait(false);
            
            _logger.LogDebug("检查点保存成功，任务ID: {TaskId}", checkpoint.TaskId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "保存检查点时发生错误，任务ID: {TaskId}", checkpoint.TaskId);
            throw;
        }
    }

    private string GetCheckpointPath(string taskId)
    {
        return Path.Combine(_checkpointDirectory, $"{taskId}.checkpoint.json");
    }

    private bool ValidateCheckpoint(ProcessingCheckpoint checkpoint)
    {
        try
        {
            // 检查检查点是否过期（超过24小时）
            if (DateTime.UtcNow - checkpoint.Timestamp > TimeSpan.FromHours(24))
            {
                _logger.LogWarning("检查点已过期，任务ID: {TaskId}", checkpoint.TaskId);
                return false;
            }
            
            // 检查已完成的文件是否仍然存在
            foreach (var completedFile in checkpoint.CompletedFiles)
            {
                if (!_fileSystemService.FileExists(completedFile))
                {
                    _logger.LogWarning("已完成的文件不存在: {FilePath}", completedFile);
                    return false;
                }
            }
            
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "验证检查点时发生错误，任务ID: {TaskId}", checkpoint.TaskId);
            return false;
        }
    }

    private async Task CleanupExpiredCheckpoints(TimeSpan maxAge)
    {
        try
        {
            if (!Directory.Exists(_checkpointDirectory))
                return;
                
            var checkpointFiles = Directory.GetFiles(_checkpointDirectory, "*.checkpoint.json");
            var cutoffTime = DateTime.UtcNow - maxAge;
            
            foreach (var checkpointFile in checkpointFiles)
            {
                try
                {
                    var lastWriteTime = _fileSystemService.GetLastWriteTime(checkpointFile);
                    if (lastWriteTime < cutoffTime)
                    {
                        // 尝试加载检查点并清理其临时文件
                        var json = await File.ReadAllTextAsync(checkpointFile).ConfigureAwait(false);
                        var checkpoint = JsonSerializer.Deserialize<ProcessingCheckpoint>(json);
                        
                        if (checkpoint != null)
                        {
                            foreach (var tempFile in checkpoint.TempFiles)
                            {
                                if (_fileSystemService.FileExists(tempFile))
                                {
                                    _fileSystemService.DeleteFile(tempFile);
                                }
                            }
                        }
                        
                        // 删除检查点文件
                        _fileSystemService.DeleteFile(checkpointFile);
                        
                        _logger.LogInformation("已清理过期检查点: {CheckpointFile}", checkpointFile);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "清理检查点文件时发生错误: {CheckpointFile}", checkpointFile);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "清理过期检查点时发生错误");
        }
    }
}