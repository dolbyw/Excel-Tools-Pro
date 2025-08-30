using CommunityToolkit.Mvvm.ComponentModel;
using System.IO;

namespace ExcelToolsPro.Models
{
    public partial class FileListItem : ObservableObject
    {
        [ObservableProperty]
        private string _originalName;

        [ObservableProperty]
        private string _newName;

        public FileInfo FileInfo { get; }

        public FileListItem(FileInfo fileInfo)
        {
            FileInfo = fileInfo;
            _originalName = fileInfo.Name;
            _newName = fileInfo.Name;
        }
    }
}