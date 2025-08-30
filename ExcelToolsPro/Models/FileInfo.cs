using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelToolsPro.Models;

/// <summary>
/// 文件信息模型
/// </summary>
public class FileInfo : INotifyPropertyChanged
{
    private string _path = string.Empty;
    private string _name = string.Empty;
    private long _size;
    private FileType _fileType;
    private bool _isValid;
    private bool _isHtmlDisguised;
    private int? _sheetCount;
    private int? _rowCount;
    private DateTime _lastModified;
    private string? _errorMessage;

    /// <summary>
    /// 文件路径
    /// </summary>
    public string Path
    {
        get => _path;
        set => SetProperty(ref _path, value);
    }

    /// <summary>
    /// 文件名
    /// </summary>
    public string Name
    {
        get => _name;
        set => SetProperty(ref _name, value);
    }

    /// <summary>
    /// 文件大小（字节）
    /// </summary>
    public long Size
    {
        get => _size;
        set => SetProperty(ref _size, value);
    }

    /// <summary>
    /// 文件类型
    /// </summary>
    public FileType FileType
    {
        get => _fileType;
        set => SetProperty(ref _fileType, value);
    }

    /// <summary>
    /// 是否为有效文件
    /// </summary>
    public bool IsValid
    {
        get => _isValid;
        set => SetProperty(ref _isValid, value);
    }

    /// <summary>
    /// 是否为HTML伪装文件
    /// </summary>
    public bool IsHtmlDisguised
    {
        get => _isHtmlDisguised;
        set => SetProperty(ref _isHtmlDisguised, value);
    }

    /// <summary>
    /// 工作表数量
    /// </summary>
    public int? SheetCount
    {
        get => _sheetCount;
        set => SetProperty(ref _sheetCount, value);
    }

    /// <summary>
    /// 行数
    /// </summary>
    public int? RowCount
    {
        get => _rowCount;
        set => SetProperty(ref _rowCount, value);
    }

    /// <summary>
    /// 最后修改时间
    /// </summary>
    public DateTime LastModified
    {
        get => _lastModified;
        set => SetProperty(ref _lastModified, value);
    }

    /// <summary>
    /// 错误消息
    /// </summary>
    public string? ErrorMessage
    {
        get => _errorMessage;
        set => SetProperty(ref _errorMessage, value);
    }

    /// <summary>
    /// 格式化的文件大小文本
    /// </summary>
    public string SizeText
    {
        get
        {
            if (Size < 1024)
                return $"{Size} B";
            if (Size < 1024 * 1024)
                return $"{Size / 1024.0:F1} KB";
            if (Size < 1024 * 1024 * 1024)
                return $"{Size / (1024.0 * 1024.0):F1} MB";
            return $"{Size / (1024.0 * 1024.0 * 1024.0):F1} GB";
        }
    }

    /// <summary>
    /// 文件状态文本
    /// </summary>
    public string StatusText
    {
        get
        {
            if (!IsValid)
                return "无效";
            if (IsHtmlDisguised)
                return "HTML伪装";
            return "正常";
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;

        field = value;
        OnPropertyChanged(propertyName);
        
        // 当Size改变时，同时通知SizeText属性
        if (propertyName == nameof(Size))
        {
            OnPropertyChanged(nameof(SizeText));
        }
        
        // 当IsValid或IsHtmlDisguised改变时，同时通知StatusText属性
        if (propertyName is nameof(IsValid) or nameof(IsHtmlDisguised))
        {
            OnPropertyChanged(nameof(StatusText));
        }
        
        return true;
    }
}

/// <summary>
/// 文件类型枚举
/// </summary>
public enum FileType
{
    /// <summary>
    /// Excel 2007+ 格式
    /// </summary>
    Xlsx,
    
    /// <summary>
    /// Excel 97-2003 格式
    /// </summary>
    Xls,
    
    /// <summary>
    /// CSV 格式
    /// </summary>
    Csv,
    
    /// <summary>
    /// HTML 格式
    /// </summary>
    Html,
    
    /// <summary>
    /// 未知格式
    /// </summary>
    Unknown
}