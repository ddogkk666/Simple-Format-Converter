[CmdletBinding()]
param(
    [switch]$SelfTest,
    [switch]$WorkerMode,
    [string]$InputListPath,
    [string]$TargetExtensionArg,
    [string]$ImageModeArg,
    [string]$StatusPath
)

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Runtime.WindowsRuntime

function Get-ToolPath {
    param(
        [Parameter(Mandatory)]
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        if ($candidate.Contains("*")) {
            $matched = Get-ChildItem -Path $candidate -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($matched) {
                return $matched.FullName
            }
            continue
        }

        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }

        $command = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($command) {
            return $command.Source
        }
    }

    return $null
}

function Get-InstalledTools {
    return [PSCustomObject]@{
        Ffmpeg = Get-ToolPath -Candidates @(
            (Join-Path $PSScriptRoot "ffmpeg.exe"),
            (Join-Path $PSScriptRoot "tools\ffmpeg\ffmpeg.exe"),
            (Join-Path $PSScriptRoot "tools\ffmpeg\bin\ffmpeg.exe"),
            "ffmpeg",
            "C:\ffmpeg\bin\ffmpeg.exe",
            "C:\Program Files\ffmpeg\bin\ffmpeg.exe",
            "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-*\bin\ffmpeg.exe"
        )
        LibreOffice = Get-ToolPath -Candidates @(
            "soffice",
            (Join-Path $PSScriptRoot "LibreOfficePortable\App\libreoffice\program\soffice.exe"),
            (Join-Path $PSScriptRoot "LibreOfficePortable\App\libreoffice\program\soffice.com"),
            (Join-Path $PSScriptRoot "LibreOfficePortable\LibreOffice\App\libreoffice\program\soffice.exe"),
            "C:\Program Files\LibreOffice\program\soffice.exe",
            "C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        )
        Python = Get-ToolPath -Candidates @(
            (Join-Path $PSScriptRoot "PythonRuntime\python.exe"),
            (Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"),
            (Join-Path $env:LOCALAPPDATA "Programs\Python\Python311\python.exe"),
            "python"
        )
    }
}

function Test-PdfRenderAvailable {
    try {
        $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.DataReader, Windows.Storage.Streams, ContentType = WindowsRuntime]
        return $true
    }
    catch {
        return $false
    }
}

function Get-PdfWinRtTypes {
    return [PSCustomObject]@{
        StorageFile = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        PdfDocument = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        InMemoryRandomAccessStream = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        DataReader = [Windows.Storage.Streams.DataReader, Windows.Storage.Streams, ContentType = WindowsRuntime]
    }
}

function Await-WinRtTask {
    param(
        [Parameter(Mandatory)]
        [object]$AsyncOperation
    )

    if ($AsyncOperation -is [Windows.Foundation.IAsyncAction]) {
        $task = [System.WindowsRuntimeSystemExtensions]::AsTask([Windows.Foundation.IAsyncAction]$AsyncOperation)
        $task.GetAwaiter().GetResult()
        return $null
    }

    $asyncInterface = $AsyncOperation.GetType().GetInterfaces() | Where-Object {
        $_.IsGenericType -and $_.Namespace -eq "Windows.Foundation" -and $_.Name -in @("IAsyncOperation`1", "IAsyncOperationWithProgress`2")
    } | Select-Object -First 1

    if (-not $asyncInterface) {
        throw "Unsupported WinRT async operation type: $($AsyncOperation.GetType().FullName)"
    }

    $asyncInterfaceDefinition = $asyncInterface.GetGenericTypeDefinition().FullName
    $candidateMethods = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq "AsTask" -and
        $_.IsGenericMethodDefinition -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.IsGenericType -and
        $_.GetParameters()[0].ParameterType.GetGenericTypeDefinition().FullName -eq $asyncInterfaceDefinition
    }

    $genericMethod = $candidateMethods | Where-Object {
        $_.GetGenericArguments().Count -eq $asyncInterface.GetGenericArguments().Count
    } | Select-Object -First 1

    if (-not $genericMethod) {
        throw "Could not locate a compatible AsTask overload for $asyncInterfaceDefinition."
    }

    $task = $genericMethod.MakeGenericMethod($asyncInterface.GetGenericArguments()).Invoke($null, @($AsyncOperation))
    $task.GetAwaiter().GetResult() | Out-Null

    $resultProperty = $task.GetType().GetProperty("Result")
    if ($resultProperty) {
        return $resultProperty.GetValue($task)
    }

    return $null
}

function Get-FileTypeProfile {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $ext = [System.IO.Path]::GetExtension($Path)
    if ([string]::IsNullOrWhiteSpace($ext)) {
        $ext = ""
    }
    $ext = $ext.ToLowerInvariant()

    switch ($ext) {
        ".pdf" {
            return [PSCustomObject]@{
                Extension = $ext
                Kind = "document"
                Targets = @("docx", "pptx", "png", "jpg")
            }
        }
        ".doc" {
            return [PSCustomObject]@{
                Extension = $ext
                Kind = "document"
                Targets = @("pdf")
            }
        }
        ".docx" {
            return [PSCustomObject]@{
                Extension = $ext
                Kind = "document"
                Targets = @("pdf")
            }
        }
        ".ppt" {
            return [PSCustomObject]@{
                Extension = $ext
                Kind = "presentation"
                Targets = @("pdf", "png", "jpg")
            }
        }
        ".pptx" {
            return [PSCustomObject]@{
                Extension = $ext
                Kind = "presentation"
                Targets = @("pdf", "png", "jpg")
            }
        }
    }

    $videoExts = @(".mp4", ".mov", ".avi", ".mkv", ".wmv", ".webm", ".flv", ".m4v")
    $audioExts = @(".mp3", ".wav", ".aac", ".flac", ".m4a", ".ogg", ".wma")
    $imageExts = @(".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff")

    if ($imageExts -contains $ext) {
        return [PSCustomObject]@{
            Extension = $ext
            Kind = "image"
            Targets = @("pdf")
        }
    }

    if ($videoExts -contains $ext) {
        return [PSCustomObject]@{
            Extension = $ext
            Kind = "media"
            Targets = @("mp3", "wav", "aac", "m4a", "flac", "mp4", "mov", "avi", "mkv", "webm") |
                Where-Object { ("." + $_) -ne $ext }
        }
    }
    