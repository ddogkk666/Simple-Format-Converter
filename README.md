# Simple Format Converter

一个基于 PowerShell + WPF 的 Windows 桌面格式转换工具，主打拖拽文件、选择目标格式、点击转换。

## 当前支持

- `DOC/DOCX -> PDF`
- `PPT/PPTX -> PDF`
- `PDF -> DOCX`
- `PDF -> PPTX`
- `PDF -> PNG/JPG`
- `PPT/PPTX -> PNG/JPG`
- `PNG/JPG/JPEG/BMP/GIF/TIFF -> PDF`
- `MP4 -> MP3`
- 常见视频格式互转：`MP4 / MOV / AVI / MKV / WEBM`
- 常见音频格式互转：`MP3 / WAV / AAC / M4A / FLAC / OGG`

## 运行方式

双击：

- `文件转换器.exe`

或者手动运行：

```powershell
powershell -ExecutionPolicy Bypass -STA -File .\FormatConverter.ps1
```

## 输出位置

转换后的文件会统一写到项目目录下的：

- `converted`

并按源文件名自动创建子文件夹，例如：

```text
C:\Users\...\文件转换\converted\example\example.pdf
```

## 引擎说明

本项目会根据本机可用环境自动选择转换引擎：

- `FFmpeg`：音视频转换
- `Microsoft Office`：部分 Word / PowerPoint 转换
- `LibreOffice`：文档与图片类转换兜底
- `pdfslides`：`PDF -> PPTX`
- `pdf2docx`：`PDF -> DOCX`

## 仓库说明

这个仓库默认只放源码和说明，不提交以下大体积运行时文件：

- `ffmpeg.exe`
- `LibreOfficePortable`
- `PythonRuntime`
- `converted`

原因：

- GitHub 对大文件有限制
- 部分第三方运行时体积过大
- 这些内容更适合本地打包或做 Release 附件

## 本地开发建议

如果你想在另一台电脑完整复现当前功能，通常需要准备：

- `FFmpeg`
- `LibreOffice`
- `Python 3.12`
- Python 包：`pdfslides`、`pdf2docx`

## 注意

- `PDF -> PPTX` 当前更偏“适合展示和编辑的幻灯片重建”，不保证和原 PDF 完全一致
- `PDF -> DOCX` 的排版结果也可能和原文件有差异
