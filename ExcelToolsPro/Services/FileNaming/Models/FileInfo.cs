using System.IO;

namespace ExcelToolsPro.Services.FileNaming.Models;

/// <summary>
/// 文件信息
/// </summary>
public class FileInfo
{
    /// <summary>
    /// 文件名（包含扩展名）
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// 完整路径
    /// </summary>
    public string FullPath { get; set; } = string.Empty;
    
    /// <summary>
    /// 文件大小（字节）
    /// </summary>
    public long Size { get; set; }
    
    /// <summary>
    /// 创建时间
    /// </summary>
    public DateTime CreatedAt { get; set; }
    
    /// <summary>
    /// 修改时间
    /// </summary>
    public DateTime ModifiedAt { get; set; }
    
    /// <summary>
    /// 文件扩展名
    /// </summary>
    public string Extension => Path.GetExtension(Name);
    
    /// <summary>
    /// 不含扩展名的文件名
    /// </summary>
    public string NameWithoutExtension => Path.GetFileNameWithoutExtension(Name);
    
    /// <summary>
    /// 目录路径
    /// </summary>
    public string DirectoryPath => Path.GetDirectoryName(FullPath) ?? string.Empty;
}