# Simple Format Converter

一个面向 Windows 的便携式文件格式转换工具，支持拖拽文件、选择目标格式，然后一键转换。

## 下载安装

便携版下载：

- [前往 Releases 下载](https://github.com/ddogkk666/Simple-Format-Converter/releases)

使用方法：

1. 下载 `Simple-Format-Converter-portable.zip`
2. 解压到任意文件夹
3. 双击 `文件转换器.exe`

## 支持的转换

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

## 输出位置

转换后的文件会统一写到程序目录下的：

- `converted`

并按源文件名自动创建子文件夹，例如：

```text
converted\example\example.pdf
```

## 仓库说明

这个仓库主要用于存放源码和说明。

体积较大的运行时文件会放在 Releases 便携包里提供下载，不直接提交到仓库文件列表中，例如：

- `ffmpeg.exe`
- `LibreOfficePortable`
- `PythonRuntime`

## 开发运行

如果你想直接从源码运行：

```powershell
powershell -ExecutionPolicy Bypass -STA -File .\FormatConverter.ps1
```

## 注意

- `PDF -> PPTX` 更适合生成可展示、可继续编辑的幻灯片，不保证与原 PDF 完全一致。
- `PDF -> DOCX` 的排版结果也可能与原文档存在差异。
