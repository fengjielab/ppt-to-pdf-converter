好的，这是一个专业且详细的 `README.md` 文件模板。您可以将以下内容复制到一个名为 `README.md` 的新文件中，并将其与您的 Python 脚本放在同一个项目文件夹下。

---

# PPT to PDF 转换器 (PPT to PDF Converter)

这是一个使用 Python 编写的脚本，能够将 Microsoft PowerPoint 演示文稿（`.ppt` 或 `.pptx` 格式）转换为 PDF 文件。

该脚本主要利用 `comtypes` 库来调用 Windows 系统上安装的 Microsoft PowerPoint 应用程序，从而实现高质量的格式转换，确保生成的 PDF 文件与原始 PPT 的布局、动画和样式保持高度一致。

## ✨ 特性

- **高质量转换**: 直接调用 PowerPoint 内核进行转换，最大限度保留原始格式。
- **简单易用**: 只需配置输入和输出文件名即可运行。
- **自动化**: 可轻松集成到其他自动化工作流程中，用于批量处理文件。
- **演示功能**: 当找不到指定的输入文件时，会自动创建一个简单的 PPTX 文件用于功能演示。

## ⚙️ 先决条件

在运行此脚本之前，请确保您的系统满足以下条件：

1.  **操作系统**: **Windows** 操作系统 (因为脚本需要调用 Windows COM 接口来控制 PowerPoint)。
2.  **软件**: 已完整安装 **Microsoft PowerPoint** 应用程序。
3.  **Python**: 已安装 Python 3.x 环境。
4.  **Python 库**:
    - `comtypes`: 用于与 COM 对象（如 PowerPoint 应用程序）进行交互。
    - `python-pptx` (可选): 用于在找不到输入文件时，自动创建一个演示用的 PPTX 文件。

## 🚀 安装指南

1.  克隆或下载本项目到您的本地计算机。

2.  打开您的命令行工具 (例如 CMD, PowerShell 或 Windows Terminal)。

3.  使用 `pip` 安装所需的 Python 库：

    ```bash
    pip install comtypes
    pip install python-pptx
    ```

## 📝 如何使用

1.  **准备文件**:
    - 将 `ppttopdf.py` 脚本文件和您想要转换的 PowerPoint 文件（例如 `MyPresentation.pptx`）放在同一个文件夹下。

2.  **修改脚本**:
    - 用代码编辑器 (如 VS Code) 打开 `ppttopdf.py` 文件。
    - 找到文件末尾的 `if __name__ == '__main__':` 代码块。
    - 修改 `input_ppt` 和 `output_pdf` 变量的值，以匹配您的文件名。

    ```python
    # --- 使用示例 ---
    if __name__ == '__main__':
        # ...
        
        # 将 "Your_Presentation.pptx" 替换为您的输入文件名
        input_ppt = "MyPresentation.pptx" 
        
        # （可选）设置您想要的输出文件名
        output_pdf = "MyPresentation_Converted.pdf"
        
        # 调用转换函数
        ppt_to_pdf(input_ppt, output_pdf)
    ```

3.  **运行脚本**:
    - 在文件所在的文件夹中打开命令行。
    - 运行以下命令：

    ```bash
    python ppttopdf.py
    ```

4.  **查看结果**:
    - 脚本运行成功后，终端会显示成功信息。
    - 您会在同一个文件夹下看到新生成的 PDF 文件（例如 `MyPresentation_Converted.pdf`）。

## 🔧 故障排除

- **错误: `Import "pptx" could not be resolved` 或 `reportMissingImports`**
  - **原因**: 缺少 `python-pptx` 库。
  - **解决方案**: 在命令行中运行 `pip install python-pptx`。

- **错误: `OSError: [WinError -2147221005] 无效的类字符串` 或其他 `comtypes` 相关错误**
  - **原因**: 脚本无法启动 PowerPoint 应用程序。这通常是因为您的系统没有安装 Microsoft PowerPoint，或者安装不完整。
  - **解决方案**: 确保您已在 Windows 系统上正确安装了 Microsoft PowerPoint 并且可以正常打开它。

- **错误: `输入文件 '...' 未找到`**
  - **原因**: 脚本中指定的 `input_ppt` 文件名不正确，或者该文件不在脚本所在的目录中。
  - **解决方案**: 请仔细检查文件名是否拼写正确（包括扩展名 `.ppt` 或 `.pptx`），并确保文件与脚本在同一路径下。

## 🐧 对于非 Windows 用户

本脚本的核心方法依赖于 Windows 和 Microsoft PowerPoint。如果您在 macOS 或 Linux 上，可以考虑使用以下替代方案：

- **商业库**: 如 [Aspose.Slides for Python](https://products.aspose.com/slides/python-net/)，它功能强大，无需安装 PowerPoint。
- **在线转换服务 API**: 如 [ConvertAPI](https://www.convertapi.com/)，它提供了 Python SDK，可以通过网络请求完成文件转换。

## 📜 许可证

本项目采用 [MIT 许可证](https://opensource.org/licenses/MIT)。请随意使用和修改。
