# PDF to PPTX Converter

## 简介

`pdf2ppt.py` 是一个基于Python的工具，旨在将PDF文件转换为PPTX格式的演示文稿。它提供了一个图形用户界面(GUI)，通过简单的操作即可完成文件格式的转换。该工具特别适合需要将PDF文档内容快速转换为可编辑PPTX格式的用户。

## 功能特点

- **图形用户界面**：提供了一个直观的图形界面，用户可以通过几次点击完成转换过程。
- **选择PDF文件**：用户可以通过文件对话框选择需要转换的PDF文件。
- **DPI设置**：支持自定义DPI（每英寸点数），以调整输出PPTX文件的图像质量。
- **进度显示**：转换过程中，进度条会显示当前的转换进度。
- **路径显示与复制**：转换完成后，会显示输出的PPTX文件路径，用户可以直接复制此路径。
- **直接打开文件**：双击输出路径即可直接打开生成的PPTX文件，方便快捷。
- **错误处理**：转换过程中如果发生错误，会通过界面提示用户错误信息。

## 使用方法

1. 运行`pdf2ppt.py`启动程序。
2. 点击“Select PDF”按钮选择一个PDF文件。
3. （可选）在“Settings”选项卡中调整DPI设置。
4. 点击“Convert”按钮开始转换。转换进度会在进度条中显示。
5. 转换完成后，会在界面上显示输出的PPTX文件路径。双击路径可以直接打开文件。

## 环境要求

- Python 3
- PyQt5
- PyMuPDF (fitz)
- python-pptx

## 安装依赖

在运行`pdf2ppt.py`之前，需要安装必要的Python库。可以通过以下命令安装：
```bash
pip install PyQt5 PyMuPDF python-pptx
```

## 开发背景

该工具的开发旨在提供一个简单、高效的方式，帮助用户将PDF文档转换为PPTX格式，以便在演示文稿中使用PDF内容。

## 许可证

本工具是开源的，欢迎在遵守原许可证的基础上进行使用和修改。

---