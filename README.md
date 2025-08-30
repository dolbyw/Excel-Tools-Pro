# Excel Tools Pro

一个功能强大的Excel文件处理工具，支持文件合并和拆分操作。

## 功能特性

- **文件合并**: 将多个Excel文件合并为一个文件
- **文件拆分**: 将大型Excel文件按指定规则拆分为多个小文件
- **智能命名**: 支持自定义文件命名模板和变量
- **批量处理**: 支持拖拽操作和批量文件处理
- **实时预览**: 提供文件名生成预览功能
- **错误处理**: 完善的错误恢复和日志记录机制

## 技术栈

- **.NET 8**: 现代化的.NET平台
- **WPF**: Windows桌面应用程序框架
- **MVVM模式**: 清晰的架构设计
- **依赖注入**: 使用Microsoft.Extensions.DependencyInjection
- **日志记录**: 集成Serilog日志框架

## 项目结构

```
ExcelToolsPro/
├── Models/              # 数据模型
├── ViewModels/          # 视图模型
├── Views/               # 用户界面
├── Services/            # 业务服务
│   ├── FileNaming/      # 文件命名服务
│   └── ...
├── Converters/          # 数据转换器
└── Styles/              # 样式资源
```

## 开始使用

### 系统要求

- Windows 10 或更高版本
- .NET 8 Runtime

### 构建项目

```bash
# 克隆仓库
git clone <repository-url>
cd Excel-Tools-Pro

# 还原依赖
dotnet restore

# 构建项目
dotnet build

# 运行应用程序
dotnet run --project ExcelToolsPro
```

### 运行测试

```bash
dotnet test
```

## 使用说明

1. **文件选择**: 拖拽Excel文件到应用程序窗口或点击"选择文件"按钮
2. **操作模式**: 选择"合并文件"或"拆分文件"模式
3. **配置设置**: 根据需要配置合并或拆分参数
4. **开始处理**: 点击"开始处理"按钮执行操作
5. **查看结果**: 处理完成后在指定输出目录查看结果文件

## 贡献指南

欢迎提交Issue和Pull Request来改进这个项目。

## 许可证

本项目采用MIT许可证。详情请参阅LICENSE文件。