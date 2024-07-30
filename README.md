
# Image Metadata Extractor

## 项目简介

Image Metadata Extractor 是一个用于提取和处理图像元数据的工具。该工具允许用户从图像文件中读取隐藏的元数据，并将这些数据保存到 Excel 文件中.

## 功能

- **图像拖放支持**：用户可以将图像文件直接拖放到应用程序窗口，自动进行处理。
- **多种图像格式支持**：支持 PNG、JPEG 等常见图像格式。
- **元数据提取**：使用嵌入式 Python 脚本从图像中提取隐藏的元数据。
- **图像压缩**：将处理过的图像进行压缩，生成 JPEG 和 PNG 格式的压缩图像。
- **数据导出**：将提取的元数据导出到 Excel 文件中，并在其中包含压缩后的图像预览。
- **自定义右键菜单**：包含设置窗口置顶和查看关于信息的选项。
- **快捷打开文件**：点击消息框中的文件路径可以直接打开该文件或所在目录。

## 使用方法

### 图像处理

1. 启动应用程序。
2. 将图像文件拖放到应用程序窗口。
3. 图像文件将自动进行处理，提取隐藏的元数据，并生成压缩图像。
4. 提取的数据将保存到 Excel 文件中，并在处理完成后弹出包含文件路径的消息框。

### 菜单选项

- **设置置顶/取消置顶**：右键点击应用程序窗口，选择 "设置置顶" 或 "取消置顶" 来切换窗口的置顶状态。
- **关于**：右键点击应用程序窗口，选择 "About Me" 查看程序介绍、GitHub 和afdian链接。

## 项目结构

```
├── .gitignore              # Git忽略文件
├── .gitattributes          # Git属性文件
├── App.config              # 应用程序配置文件
├── extract_data_win.csproj # 项目文件
├── extract_data_win.sln    # 解决方案文件
├── MainForm.cs             # 主窗体代码文件
├── MainForm.Designer.cs    # 主窗体设计器文件
├── MainForm.resx           # 主窗体资源文件
├── myicon.ico              # 应用程序图标
├── packages.config         # NuGet包配置文件
├── Program.cs              # 应用程序入口文件
├── LICENSE                 # 许可证文件
├── README.md               # 项目介绍文档
└── /py                     # Python脚本文件夹
	├── extract_data.py     # 元数据提取脚本
	└── extract_data.spec   # PyInstaller打包配置文件
└── /Properties             # 项目属性文件夹
	├── AssemblyInfo.cs     # 程序集信息文件
	└── Resources.resx      # 资源文件
```

## 开发环境

- **开发工具**：Visual Studio
- **编程语言**：C#
- **框架**：.NET Framework
- **依赖项**：请查看 `packages.config` 文件中的 NuGet 包列表。

## 安装和运行

1. 克隆仓库到本地：
    ```sh
    git clone https://github.com/YILING0013/extract_data_win.git
    ```
2. 打开 Visual Studio 并加载解决方案文件 (`extract_data_win.sln`)。
3. 还原 NuGet 包：
    ```sh
    nuget restore
    ```
4. 编译并运行项目。

## 贡献

欢迎任何形式的贡献！如果您有任何建议或改进，请提交 Pull Request 或在 Issue 区留言.

## 其他

- [NovelAI绘图](https://nai3.xianyun.cool/)https://nai3.xianyun.cool/
- [爱发电](https://afdian.com/a/lingyunfei)https://afdian.com/a/lingyunfei
