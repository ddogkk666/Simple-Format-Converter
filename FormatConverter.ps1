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

    if ($audioExts -contains $ext) {
        return [PSCustomObject]@{
            Extension = $ext
            Kind = "media"
            Targets = @("mp3", "wav", "aac", "m4a", "flac", "ogg") |
                Where-Object { ("." + $_) -ne $ext }
        }
    }

    return $null
}

function Get-OutputPath {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension
    )

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $outputDirectory = Get-OutputDirectory -InputPath $InputPath
    return (Join-Path $outputDirectory ($fileName + "." + $TargetExtension))
}

function Get-OutputDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath
    )

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $safeFolderName = $fileName
    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safeFolderName = $safeFolderName.Replace([string]$invalidChar, "_")
    }

    $outputRoot = Join-Path $PSScriptRoot "converted"
    $outputDirectory = Join-Path $outputRoot $safeFolderName

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    return $outputDirectory
}

function Get-PagedOutputPath {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension,
        [Parameter(Mandatory)]
        [int]$PageNumber
    )

    $outputDirectory = Get-OutputDirectory -InputPath $InputPath
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    return (Join-Path $outputDirectory ($fileName + "_p" + $PageNumber.ToString("000") + "." + $TargetExtension))
}

function Ensure-ComObjectReleased {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Test-WordAvailable {
    try {
        return ($null -ne [type]::GetTypeFromProgID("Word.Application"))
    }
    catch {
        return $false
    }
}

function Test-PowerPointAvailable {
    try {
        return ($null -ne [type]::GetTypeFromProgID("PowerPoint.Application"))
    }
    catch {
        return $false
    }
}

function Convert-WithLibreOffice {
    param(
        [Parameter(Mandatory)]
        [string]$LibreOfficePath,
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension
    )

    $outputPath = Get-OutputPath -InputPath $InputPath -TargetExtension $TargetExtension
    $outputDirectory = [System.IO.Path]::GetDirectoryName($outputPath)

    $filterMap = @{
        "pdf"  = "pdf"
        "docx" = "docx"
        "pptx" = "pptx"
        "png"  = "png"
        "jpg"  = "jpg"
    }

    if (-not $filterMap.ContainsKey($TargetExtension)) {
        throw "LibreOffice does not support .$TargetExtension output."
    }

    $quotedOutputDirectory = '"' + $outputDirectory.Replace('"', '""') + '"'
    $quotedInputPath = '"' + $InputPath.Replace('"', '""') + '"'
    $arguments = '--headless --convert-to ' + $filterMap[$TargetExtension] + ' --outdir ' + $quotedOutputDirectory + ' ' + $quotedInputPath

    $process = Start-Process -FilePath $LibreOfficePath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) {
        throw "LibreOffice conversion failed with exit code $($process.ExitCode)."
    }

    if (Test-Path -LiteralPath $outputPath) {
        return $outputPath
    }

    $converted = Get-ChildItem -LiteralPath $outputDirectory -File |
        Where-Object {
            $_.BaseName -eq [System.IO.Path]::GetFileNameWithoutExtension($InputPath) -and
            $_.Extension -eq ("." + $TargetExtension)
        } |
        Select-Object -First 1

    if ($converted) {
        return $converted.FullName
    }

    throw "LibreOffice did not create the expected output file."
}

function Convert-PdfToPptxWithPdfslides {
    param(
        [Parameter(Mandatory)]
        [string]$PythonPath,
        [Parameter(Mandatory)]
        [string]$InputPath
    )

    $outputPath = Get-OutputPath -InputPath $InputPath -TargetExtension "pptx"
    if (Test-Path -LiteralPath $outputPath) {
        Remove-Item -LiteralPath $outputPath -Force -ErrorAction SilentlyContinue
    }

    & $PythonPath -m pdfslides --pdf $InputPath --output $outputPath --dpi 200 | Out-Null
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "pdfslides conversion failed with exit code $exitCode."
    }

    if (-not (Test-Path -LiteralPath $outputPath)) {
        throw "pdfslides did not create the expected PPTX file."
    }

    return $outputPath
}

function Convert-PdfToDocxWithPdf2docx {
    param(
        [Parameter(Mandatory)]
        [string]$PythonPath,
        [Parameter(Mandatory)]
        [string]$InputPath
    )

    $outputPath = Get-OutputPath -InputPath $InputPath -TargetExtension "docx"
    if (Test-Path -LiteralPath $outputPath) {
        Remove-Item -LiteralPath $outputPath -Force -ErrorAction SilentlyContinue
    }

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("converter_pdf2docx_" + [guid]::NewGuid().ToString("N"))
    $tempInputPath = Join-Path $tempRoot "input.pdf"
    $tempOutputPath = Join-Path $tempRoot "output.docx"
    $scriptPath = Join-Path ([System.IO.Path]::GetTempPath()) ("converter_pdf2docx_" + [guid]::NewGuid().ToString("N") + ".py")
    @'
from pdf2docx import Converter
import sys

pdf_path = sys.argv[1]
docx_path = sys.argv[2]

converter = Converter(pdf_path)
try:
    converter.convert(docx_path)
finally:
    converter.close()
'@ | Set-Content -LiteralPath $scriptPath -Encoding UTF8

    try {
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        Copy-Item -LiteralPath $InputPath -Destination $tempInputPath -Force

        & $PythonPath $scriptPath $tempInputPath $tempOutputPath | Out-Null
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            throw "pdf2docx conversion failed with exit code $exitCode."
        }

        if (-not (Test-Path -LiteralPath $tempOutputPath)) {
            throw "pdf2docx did not create the expected DOCX file."
        }

        Move-Item -LiteralPath $tempOutputPath -Destination $outputPath -Force
        return $outputPath
    }
    finally {
        if (Test-Path -LiteralPath $scriptPath) {
            Remove-Item -LiteralPath $scriptPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Convert-WordCom {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension
    )

    $word = $null
    $document = $null
    $workingInputPath = $InputPath
    $tempWorkingDirectory = $null

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.ScreenUpdating = $false
        $word.Options.ConfirmConversions = $false
        $word.Options.SaveNormalPrompt = $false

        if ([System.IO.Path]::GetExtension($InputPath).ToLowerInvariant() -eq ".pdf" -and $TargetExtension -eq "docx") {
            $tempWorkingDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("converter_word_" + [guid]::NewGuid().ToString("N"))
            New-Item -ItemType Directory -Path $tempWorkingDirectory -Force | Out-Null
            $workingInputPath = Join-Path $tempWorkingDirectory "input.pdf"
            Copy-Item -LiteralPath $InputPath -Destination $workingInputPath -Force
        }

        $document = $word.Documents.OpenNoRepairDialog($workingInputPath)
        $outputPath = Get-OutputPath -InputPath $InputPath -TargetExtension $TargetExtension

        switch ($TargetExtension) {
            "pdf" { $document.ExportAsFixedFormat($outputPath, 17) }
            "docx" { $document.SaveAs2($outputPath, 16) }
            default { throw "Word COM does not support .$TargetExtension output." }
        }

        return $outputPath
    }
    finally {
        if ($document) {
            $document.Close(0)
        }
        if ($word) {
            $word.Quit()
        }
        Ensure-ComObjectReleased -ComObject $document
        Ensure-ComObjectReleased -ComObject $word
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        if ($tempWorkingDirectory -and (Test-Path -LiteralPath $tempWorkingDirectory)) {
            Remove-Item -LiteralPath $tempWorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Convert-PowerPointCom {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension
    )

    if ($TargetExtension -ne "pdf") {
        throw "PowerPoint COM only supports PDF output."
    }

    $powerPoint = $null
    $presentation = $null

    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $presentation = $powerPoint.Presentations.Open($InputPath, $true, $false, $false)
        $outputPath = Get-OutputPath -InputPath $InputPath -TargetExtension $TargetExtension
        $presentation.SaveAs($outputPath, 32)
        return $outputPath
    }
    finally {
        if ($presentation) {
            $presentation.Close()
        }
        if ($powerPoint) {
            $powerPoint.Quit()
        }
        Ensure-ComObjectReleased -ComObject $presentation
        Ensure-ComObjectReleased -ComObject $powerPoint
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Merge-Images {
    param(
        [Parameter(Mandatory)]
        [string[]]$ImagePaths,
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    if ($ImagePaths.Count -eq 0) {
        throw "No images were provided for merging."
    }

    $bitmaps = New-Object System.Collections.Generic.List[System.Drawing.Bitmap]
    try {
        foreach ($imagePath in $ImagePaths) {
            $bitmaps.Add([System.Drawing.Bitmap]::FromFile($imagePath))
        }

        $maxWidth = ($bitmaps | Measure-Object -Property Width -Maximum).Maximum
        $totalHeight = ($bitmaps | Measure-Object -Property Height -Sum).Sum

        $canvas = New-Object System.Drawing.Bitmap($maxWidth, $totalHeight)
        try {
            $graphics = [System.Drawing.Graphics]::FromImage($canvas)
            try {
                $graphics.Clear([System.Drawing.Color]::White)
                $offsetY = 0
                foreach ($bitmap in $bitmaps) {
                    $graphics.DrawImage($bitmap, 0, $offsetY, $bitmap.Width, $bitmap.Height)
                    $offsetY += $bitmap.Height
                }
            }
            finally {
                $graphics.Dispose()
            }

            $format = if ($OutputPath.ToLowerInvariant().EndsWith(".jpg")) {
                [System.Drawing.Imaging.ImageFormat]::Jpeg
            }
            else {
                [System.Drawing.Imaging.ImageFormat]::Png
            }
            $canvas.Save($OutputPath, $format)
        }
        finally {
            $canvas.Dispose()
        }
    }
    finally {
        foreach ($bitmap in $bitmaps) {
            $bitmap.Dispose()
        }
    }

    return $OutputPath
}

function Convert-ImagesToPdf {
    param(
        [Parameter(Mandatory)]
        [string[]]$InputPaths
    )

    if (-not $wordAvailable) {
        throw "Image to PDF currently requires Microsoft Word on this computer."
    }

    $firstInput = $InputPaths[0]
    $outputPath = Get-OutputPath -InputPath $firstInput -TargetExtension "pdf"
    $word = $null
    $document = $null

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Add()
        $selection = $word.Selection

        for ($index = 0; $index -lt $InputPaths.Count; $index++) {
            $imagePath = $InputPaths[$index]
            $selection.InlineShapes.AddPicture($imagePath) | Out-Null
            if ($index -lt ($InputPaths.Count - 1)) {
                $selection.InsertBreak(7)
            }
        }

        $document.ExportAsFixedFormat($outputPath, 17)
        return $outputPath
    }
    finally {
        if ($document) {
            $document.Close($false)
        }
        if ($word) {
            $word.Quit()
        }
        Ensure-ComObjectReleased -ComObject $document
        Ensure-ComObjectReleased -ComObject $word
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Convert-MediaWithFfmpeg {
    param(
        [Parameter(Mandatory)]
        [string]$FfmpegPath,
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension
    )

    $outputPath = Get-OutputPath -InputPath $InputPath -TargetExtension $TargetExtension

    if ($TargetExtension -eq "mp3") {
        $arguments = @("-y", "-i", $InputPath, "-vn", "-codec:a", "libmp3lame", "-q:a", "2", $outputPath)
    }
    elseif ($TargetExtension -in @("wav", "flac", "aac", "m4a", "ogg")) {
        $arguments = @("-y", "-i", $InputPath, $outputPath)
    }
    else {
        $arguments = @("-y", "-i", $InputPath, "-c:v", "libx264", "-c:a", "aac", $outputPath)
    }

    $process = Start-Process -FilePath $FfmpegPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) {
        throw "FFmpeg conversion failed with exit code $($process.ExitCode)."
    }

    if (-not (Test-Path -LiteralPath $outputPath)) {
        throw "FFmpeg did not create the expected output file."
    }

    return $outputPath
}

function Convert-PowerPointToImages {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension,
        [Parameter(Mandatory)]
        [ValidateSet("separate")]
        [string]$ImageMode
    )

    if (-not $powerPointAvailable) {
        throw "PowerPoint image export is not available on this computer."
    }

    $powerPoint = $null
    $presentation = $null
    $renderDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("converter_ppt_render_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $renderDirectory -Force | Out-Null

    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $presentation = $powerPoint.Presentations.Open($InputPath, $true, $false, $false)
        $saveFormat = if ($TargetExtension -eq "png") { 18 } else { 17 }
        $presentation.SaveAs($renderDirectory, $saveFormat)

        $generated = Get-ChildItem -LiteralPath $renderDirectory -File |
            Where-Object { $_.Extension -eq ("." + $TargetExtension) } |
            Sort-Object Name

        if ($generated.Count -eq 0) {
            throw "PowerPoint did not generate any slide images."
        }

        $outputPaths = @()
        $pageNumber = 1
        foreach ($item in $generated) {
            $pageOutputPath = Get-PagedOutputPath -InputPath $InputPath -TargetExtension $TargetExtension -PageNumber $pageNumber
            Copy-Item -LiteralPath $item.FullName -Destination $pageOutputPath -Force
            $outputPaths += $pageOutputPath
            $pageNumber++
        }

        return $outputPaths
    }
    finally {
        if ($presentation) {
            $presentation.Close()
        }
        if ($powerPoint) {
            $powerPoint.Quit()
        }
        Ensure-ComObjectReleased -ComObject $presentation
        Ensure-ComObjectReleased -ComObject $powerPoint
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        if (Test-Path -LiteralPath $renderDirectory) {
            Remove-Item -LiteralPath $renderDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Convert-PdfToImages {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension,
        [Parameter(Mandatory)]
        [ValidateSet("separate")]
        [string]$ImageMode
    )

    if (-not $pdfRenderAvailable) {
        throw "PDF rendering is not available on this computer."
    }

    $winRtTypes = Get-PdfWinRtTypes
    $storageFile = Await-WinRtTask -AsyncOperation ($winRtTypes.StorageFile::GetFileFromPathAsync($InputPath))
    $pdfDocument = Await-WinRtTask -AsyncOperation ($winRtTypes.PdfDocument::LoadFromFileAsync($storageFile))

    $generatedImages = @()
    for ($pageIndex = 0; $pageIndex -lt $pdfDocument.PageCount; $pageIndex++) {
        $page = $pdfDocument.GetPage($pageIndex)
        try {
            $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            try {
                Await-WinRtTask -AsyncOperation ($page.RenderToStreamAsync($stream)) | Out-Null
                $stream.Seek(0)

                $reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0))
                try {
                    Await-WinRtTask -AsyncOperation ($reader.LoadAsync([uint32]$stream.Size)) | Out-Null
                    $bytes = New-Object byte[] ([int]$stream.Size)
                    $reader.ReadBytes($bytes)
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $stream.Dispose()
            }

            $memory = New-Object System.IO.MemoryStream(, $bytes)
            try {
                $bitmap = [System.Drawing.Bitmap]::FromStream($memory)
                try {
                    $pageOutputPath = Get-PagedOutputPath -InputPath $InputPath -TargetExtension $TargetExtension -PageNumber ($pageIndex + 1)
                    $format = if ($TargetExtension -eq "jpg") { [System.Drawing.Imaging.ImageFormat]::Jpeg } else { [System.Drawing.Imaging.ImageFormat]::Png }
                    $bitmap.Save($pageOutputPath, $format)
                    $generatedImages += $pageOutputPath
                }
                finally {
                    $bitmap.Dispose()
                }
            }
            finally {
                $memory.Dispose()
            }
        }
        finally {
            $page.Dispose()
        }
    }

    if ($generatedImages.Count -eq 0) {
        throw "PDF export did not generate any page images."
    }

    return $generatedImages
}

function Get-FriendlyErrorMessage {
    param(
        [Parameter(Mandatory)]
        [System.Exception]$Exception
    )

    $message = $Exception.Message
    if ($Exception.InnerException) {
        $message += "`r`nInner: " + $Exception.InnerException.Message
    }
    return $message
}

function Write-WorkerStatus {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [hashtable]$Data
    )

    $json = $Data | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
}

function Convert-SingleFile {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [string]$TargetExtension,
        [Parameter(Mandatory)]
        [pscustomobject]$Tools,
        [Parameter(Mandatory)]
        [string]$ImageMode
    )

    $profile = Get-FileTypeProfile -Path $InputPath
    if (-not $profile) {
        throw "Unsupported file type: $InputPath"
    }

    $sourceExt = $profile.Extension.TrimStart(".")

    if ($sourceExt -eq $TargetExtension) {
        throw "Source and target formats are the same."
    }

    if ($sourceExt -in @("ppt", "pptx") -and $TargetExtension -eq "pdf") {
        if ($Tools.LibreOffice) {
            try {
                return Convert-WithLibreOffice -LibreOfficePath $Tools.LibreOffice -InputPath $InputPath -TargetExtension $TargetExtension
            }
            catch {
                if (-not $powerPointAvailable) {
                    throw
                }
            }
        }

        if ($powerPointAvailable) {
            return Convert-PowerPointCom -InputPath $InputPath -TargetExtension $TargetExtension
        }

        throw "No presentation conversion engine was found."
    }

    if ($sourceExt -in @("ppt", "pptx") -and $TargetExtension -in @("png", "jpg")) {
        if ($powerPointAvailable) {
            try {
                return Convert-PowerPointToImages -InputPath $InputPath -TargetExtension $TargetExtension -ImageMode $ImageMode
            }
            catch {
                if (-not $Tools.LibreOffice) {
                    throw
                }
            }
        }

        if ($Tools.LibreOffice) {
            return Convert-WithLibreOffice -LibreOfficePath $Tools.LibreOffice -InputPath $InputPath -TargetExtension $TargetExtension
        }

        throw "No presentation image export engine was found."
    }

    if ($sourceExt -eq "pdf" -and $TargetExtension -eq "docx") {
        $conversionErrors = [System.Collections.Generic.List[string]]::new()

        if ($Tools.Python) {
            try {
                return Convert-PdfToDocxWithPdf2docx -PythonPath $Tools.Python -InputPath $InputPath
            }
            catch {
                $conversionErrors.Add("pdf2docx: " + (Get-FriendlyErrorMessage -Exception $_.Exception))
            }
        }

        if ($wordAvailable) {
            try {
                return Convert-WordCom -InputPath $InputPath -TargetExtension $TargetExtension
            }
            catch {
                $conversionErrors.Add("Word: " + (Get-FriendlyErrorMessage -Exception $_.Exception))
            }
        }
        if ($Tools.LibreOffice) {
            try {
                return Convert-WithLibreOffice -LibreOfficePath $Tools.LibreOffice -InputPath $InputPath -TargetExtension $TargetExtension
            }
            catch {
                $conversionErrors.Add("LibreOffice: " + (Get-FriendlyErrorMessage -Exception $_.Exception))
            }
        }

        if ($conversionErrors.Count -gt 0) {
            throw ("This PDF could not be converted to DOCX on this computer.`r`n" + ($conversionErrors -join "`r`n"))
        }

        throw "No document conversion engine was found."
    }

    if ($sourceExt -eq "pdf" -and $TargetExtension -eq "pptx") {
        if ($Tools.Python) {
            return Convert-PdfToPptxWithPdfslides -PythonPath $Tools.Python -InputPath $InputPath
        }
        throw "PDF to PPTX requires the lightweight pdfslides engine, but Python was not found."
    }

    if ($sourceExt -eq "pdf" -and $TargetExtension -in @("png", "jpg")) {
        if ($Tools.LibreOffice) {
            try {
                return Convert-WithLibreOffice -LibreOfficePath $Tools.LibreOffice -InputPath $InputPath -TargetExtension $TargetExtension
            }
            catch {
                if (-not $pdfRenderAvailable) {
                    throw
                }
            }
        }

        return Convert-PdfToImages -InputPath $InputPath -TargetExtension $TargetExtension -ImageMode $ImageMode
    }

    if ($profile.Kind -eq "image" -and $TargetExtension -eq "pdf") {
        return Convert-ImagesToPdf -InputPaths @($InputPath)
    }

    if ($sourceExt -in @("doc", "docx") -and $TargetExtension -eq "pdf") {
        if ($wordAvailable) {
            return Convert-WordCom -InputPath $InputPath -TargetExtension $TargetExtension
        }

        if ($Tools.LibreOffice) {
            return Convert-WithLibreOffice -LibreOfficePath $Tools.LibreOffice -InputPath $InputPath -TargetExtension $TargetExtension
        }

        throw "No document conversion engine was found."
    }

    if ($profile.Kind -eq "media") {
        if (-not $Tools.Ffmpeg) {
            throw "FFmpeg was not found. Place ffmpeg.exe next to the app to enable media conversion."
        }
        return Convert-MediaWithFfmpeg -FfmpegPath $Tools.Ffmpeg -InputPath $InputPath -TargetExtension $TargetExtension
    }

    if ($Tools.LibreOffice) {
        return Convert-WithLibreOffice -LibreOfficePath $Tools.LibreOffice -InputPath $InputPath -TargetExtension $TargetExtension
    }

    throw "No compatible conversion engine was available for this file."
}

function Convert-Files {
    param(
        [Parameter(Mandatory)]
        [string[]]$InputPaths,
        [Parameter(Mandatory)]
        [string]$TargetExtension,
        [Parameter(Mandatory)]
        [pscustomobject]$Tools,
        [Parameter(Mandatory)]
        [string]$ImageMode
    )

    $profiles = @($InputPaths | ForEach-Object { Get-FileTypeProfile -Path $_ })
    if ($profiles.Count -gt 0 -and ($profiles | Where-Object { $_.Kind -ne "image" }).Count -eq 0 -and $TargetExtension -eq "pdf") {
        return ,(Convert-ImagesToPdf -InputPaths $InputPaths)
    }

    $outputs = New-Object System.Collections.Generic.List[string]
    foreach ($input in $InputPaths) {
        $result = Convert-SingleFile -InputPath $input -TargetExtension $TargetExtension -Tools $Tools -ImageMode $ImageMode
        foreach ($path in @($result)) {
            [void]$outputs.Add($path)
        }
    }

    return ,$outputs.ToArray()
}

function Add-LogLine {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Controls.TextBox]$LogBox,
        [Parameter(Mandatory)]
        [string]$Message
    )

    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $line = "[{0}] {1}" -f $timestamp, $Message
    if ([string]::IsNullOrWhiteSpace($LogBox.Text)) {
        $LogBox.Text = $line
    }
    else {
        $LogBox.AppendText("`r`n" + $line)
    }
    $LogBox.ScrollToEnd()
}

function Update-ConversionProgress {
    param(
        [Parameter(Mandatory)]
        [int]$CompletedCount,
        [Parameter(Mandatory)]
        [int]$TotalCount,
        [Parameter(Mandatory)]
        [string]$StatusText
    )

    $progressStatusText.Text = $StatusText
    if ($TotalCount -le 0) {
        $conversionProgressBar.Value = 0
        return
    }

    $value = [Math]::Min(100, [Math]::Max(0, ($CompletedCount / $TotalCount) * 100))
    $conversionProgressBar.Value = $value
}

function Get-AvailableTargetsForFiles {
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[string]]$Files
    )

    $profiles = @($Files | ForEach-Object { Get-FileTypeProfile -Path $_ })
    if ($profiles.Count -eq 0) {
        return @()
    }

    $commonTargets = $profiles[0].Targets
    foreach ($profile in $profiles | Select-Object -Skip 1) {
        $commonTargets = @($commonTargets | Where-Object { $profile.Targets -contains $_ })
    }

    $availableTargets = @()
    foreach ($target in $commonTargets) {
        if (Test-TargetSupportedForFile -Profile $profiles[0] -TargetExtension $target -tools $tools) {
            $availableTargets += $target
        }
    }

    return $availableTargets
}

function Test-TargetSupportedForFile {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Profile,
        [Parameter(Mandatory)]
        [string]$TargetExtension,
        [Parameter(Mandatory)]
        [pscustomobject]$tools
    )

    $sourceExt = $Profile.Extension.TrimStart(".")

    if ($Profile.Kind -eq "media") {
        return [bool]$tools.Ffmpeg
    }

    if ($Profile.Kind -eq "image" -and $TargetExtension -eq "pdf") {
        return $wordAvailable
    }

    if ($sourceExt -in @("ppt", "pptx") -and $TargetExtension -eq "pdf") {
        return ([bool]$tools.LibreOffice -or $powerPointAvailable)
    }

    if ($sourceExt -eq "pdf" -and $TargetExtension -eq "docx") {
        return ([bool]$tools.Python -or $wordAvailable -or [bool]$tools.LibreOffice)
    }

    if ($sourceExt -eq "pdf" -and $TargetExtension -eq "pptx") {
        return [bool]$tools.Python
    }

    if ($sourceExt -eq "pdf" -and $TargetExtension -in @("png", "jpg")) {
        return ([bool]$tools.LibreOffice -or $pdfRenderAvailable)
    }

    if ($sourceExt -in @("doc", "docx") -and $TargetExtension -eq "pdf") {
        return ($wordAvailable -or [bool]$tools.LibreOffice)
    }

    if ($sourceExt -in @("ppt", "pptx") -and $TargetExtension -in @("png", "jpg")) {
        return ($powerPointAvailable -or [bool]$tools.LibreOffice)
    }

    return [bool]$tools.LibreOffice
}

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Simple File Converter"
        Height="860"
        Width="1280"
        MinHeight="760"
        MinWidth="1100"
        WindowStartupLocation="CenterScreen"
        Background="#FAF6EF"
        AllowsTransparency="False"
        FontFamily="Segoe UI">
    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <Border Grid.Row="0"
                CornerRadius="26"
                Background="#246B5A"
                Padding="26"
                Margin="0,0,0,18">
            <StackPanel>
                <TextBlock Text="Simple File Converter"
                           Foreground="White"
                           FontSize="32"
                           FontWeight="Bold" />
                <TextBlock Margin="0,10,0,0"
                           Text="Drop files into the window, choose a target format, and start converting."
                           Foreground="#F1F7F4"
                           FontSize="14" />
            </StackPanel>
        </Border>
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="3*" />
                <ColumnDefinition Width="1.7*" />
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0"
                    x:Name="DropZone"
                    AllowDrop="True"
                    CornerRadius="22"
                    BorderBrush="#C8E1D8"
                    BorderThickness="2"
                    Background="#FFFEFB"
                    Padding="24"
                    Margin="0,0,16,0">
                <StackPanel>
                    <TextBlock Text="Drop files here"
                               FontSize="24"
                               FontWeight="Bold"
                               Foreground="#183B4A" />
                    <TextBlock x:Name="DropHint"
                               Margin="0,10,0,0"
                               Text="You can also browse manually. Batch conversion is supported when all files share a common target format."
                               TextWrapping="Wrap"
                               Foreground="#50636E"
                               FontSize="13" />
                    <Button x:Name="BrowseButton"
                            Content="Choose Files"
                            Width="170"
                            Height="44"
                            Margin="0,18,0,0"
                            Background="#E4F2EC"
                            Foreground="#173B49"
                            BorderThickness="0"
                            FontWeight="SemiBold"
                            Cursor="Hand" />
                </StackPanel>
            </Border>
            <Border Grid.Column="1"
                    CornerRadius="22"
                    BorderBrush="#E8E1D5"
                    BorderThickness="1"
                    Background="#FFFEFB"
                    Padding="20">
                <StackPanel>
                    <TextBlock Text="Conversion"
                               FontSize="24"
                               FontWeight="Bold"
                               Foreground="#183B4A" />
                    <TextBlock Margin="0,14,0,4"
                               Text="Target format"
                               Foreground="#50636E"
                               FontSize="13" />
                    <ComboBox x:Name="TargetFormatCombo"
                              Height="40"
                              Background="#F7F7F7"
                              BorderBrush="#D7D7D7"
                              Padding="10,4" />
                    <TextBlock x:Name="ImageModeLabel"
                               Margin="0,14,0,4"
                               Text="Multi-page image output"
                               Foreground="#50636E"
                               FontSize="13"
                               Visibility="Collapsed" />
                    <ComboBox x:Name="ImageModeCombo"
                              Height="40"
                              Background="#F7F7F7"
                              BorderBrush="#D7D7D7"
                              Padding="10,4"
                              Visibility="Collapsed" />
                    <TextBlock x:Name="EngineInfo"
                               Margin="0,16,0,0"
                               Foreground="#3C5B68"
                               FontSize="12"
                               TextWrapping="Wrap" />
                    <Button x:Name="ConvertButton"
                            Content="Start Conversion"
                            Height="44"
                            Margin="0,18,0,0"
                            Background="#C96F33"
                            Foreground="White"
                            BorderThickness="0"
                            FontWeight="Bold"
                            Cursor="Hand" />
                    <Button x:Name="OpenLastOutputButton"
                            Content="Open Output File"
                            Height="40"
                            Margin="0,10,0,0"
                            Background="#E9F2ED"
                            Foreground="#173B49"
                            BorderThickness="0"
                            FontWeight="SemiBold"
                            Cursor="Hand"
                            IsEnabled="False" />
                    <Button x:Name="OpenConvertedFolderButton"
                            Content="Open Converted Folder"
                            Height="40"
                            Margin="0,8,0,0"
                            Background="#F2E7D8"
                            Foreground="#173B49"
                            BorderThickness="0"
                            FontWeight="SemiBold"
                            Cursor="Hand"
                            IsEnabled="False" />
                    <TextBlock x:Name="ProgressStatusText"
                               Margin="0,16,0,0"
                               Foreground="#3C5B68"
                               FontSize="12"
                               Text="Ready to convert" />
                    <ProgressBar x:Name="ConversionProgressBar"
                                 Height="16"
                                 Margin="0,10,0,0"
                                 Minimum="0"
                                 Maximum="100"
                                 Value="0"
                                 Background="#EFE4D8"
                                 Foreground="#C96F33" />
                </StackPanel>
            </Border>
        </Grid>
        <Grid Grid.Row="2" Margin="0,20,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="*" />
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.3*" />
                <ColumnDefinition Width="1*" />
            </Grid.ColumnDefinitions>
            <Border Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="2"
                    CornerRadius="20"
                    BorderBrush="#E8E1D5"
                    BorderThickness="1"
                    Background="#FFFEFB"
                    Padding="20"
                    Margin="0,0,0,16">
                <TextBlock Text="Files"
                           FontSize="18"
                           FontWeight="Bold"
                           Foreground="#183B4A" />
            </Border>
            <Border Grid.Row="1" Grid.Column="0"
                    CornerRadius="20"
                    BorderBrush="#E8E1D5"
                    BorderThickness="1"
                    Background="#FFFEFB"
                    Padding="20"
                    Margin="0,0,16,0">
                <ListBox x:Name="FileListBox"
                         BorderThickness="0"
                         Background="Transparent"
                         FontSize="12" />
            </Border>
            <Border Grid.Row="1" Grid.Column="1"
                    CornerRadius="20"
                    BorderBrush="#E8E1D5"
                    BorderThickness="1"
                    Background="#FFFEFB"
                    Padding="20">
                <DockPanel>
                    <TextBlock DockPanel.Dock="Top"
                               Text="Log"
                               FontSize="18"
                               FontWeight="Bold"
                               Foreground="#183B4A"
                               Margin="0,0,0,10" />
                    <TextBox x:Name="LogTextBox"
                             VerticalScrollBarVisibility="Auto"
                             HorizontalScrollBarVisibility="Disabled"
                             TextWrapping="Wrap"
                             IsReadOnly="True"
                             BorderThickness="0"
                             Background="Transparent" />
                </DockPanel>
            </Border>
        </Grid>
        <TextBlock Grid.Row="3"
                   Margin="6,16,0,0"
                   Foreground="#6E7B77"
                   Text="Converted files are written to the project's converted folder, grouped by source file name." />
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

$dropZone = $window.FindName("DropZone")
$browseButton = $window.FindName("BrowseButton")
$dropHint = $window.FindName("DropHint")
$targetFormatCombo = $window.FindName("TargetFormatCombo")
$imageModeLabel = $window.FindName("ImageModeLabel")
$imageModeCombo = $window.FindName("ImageModeCombo")
$engineInfo = $window.FindName("EngineInfo")
$convertButton = $window.FindName("ConvertButton")
$openLastOutputButton = $window.FindName("OpenLastOutputButton")
$openConvertedFolderButton = $window.FindName("OpenConvertedFolderButton")
$progressStatusText = $window.FindName("ProgressStatusText")
$conversionProgressBar = $window.FindName("ConversionProgressBar")
$fileListBox = $window.FindName("FileListBox")
$logTextBox = $window.FindName("LogTextBox")

$selectedFiles = [System.Collections.Generic.List[string]]::new()
$tools = Get-InstalledTools
$wordAvailable = Test-WordAvailable
$powerPointAvailable = Test-PowerPointAvailable
$pdfRenderAvailable = Test-PdfRenderAvailable
$script:currentConversionProcess = $null
$script:currentStatusPath = $null
$script:currentInputListPath = $null
$script:currentTimeoutTimer = $null
$script:currentJobId = $null
$script:lastOutputPaths = @()

[void]$imageModeCombo.Items.Add("Separate images")
$imageModeCombo.SelectedIndex = 0

function Refresh-EngineInfo {
    $lines = @()
    $lines += ("LibreOffice: " + $(if ($tools.LibreOffice) { "detected" } else { "not found" }))
    $lines += ("Microsoft Word: " + $(if ($wordAvailable) { "available" } else { "not found" }))
    $lines += ("Microsoft PowerPoint: " + $(if ($powerPointAvailable) { "available" } else { "not found" }))
    $lines += ("Windows PDF renderer: " + $(if ($pdfRenderAvailable) { "available" } else { "not found" }))
    $lines += ("FFmpeg: " + $(if ($tools.Ffmpeg) { "detected" } else { "not found" }))
    $lines += ("pdfslides: " + $(if ($tools.Python) { "detected" } else { "not found" }))
    $lines += ("pdf2docx: " + $(if ($tools.Python) { "detected" } else { "not found" }))
    if (-not $tools.Ffmpeg) {
        $lines += "Tip: place ffmpeg.exe next to this script, or in tools\ffmpeg\bin\ffmpeg.exe"
    }
    $engineInfo.Text = ($lines -join "`r`n")
}

function Update-ImageModeVisibility {
    $imageModeLabel.Visibility = "Collapsed"
    $imageModeCombo.Visibility = "Collapsed"
    $imageModeCombo.IsEnabled = $false
}

function Get-ConvertedRootPath {
    return (Join-Path $PSScriptRoot "converted")
}

function Update-OutputButtons {
    $hasLastOutput = $false
    foreach ($path in @($script:lastOutputPaths)) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            $hasLastOutput = $true
            break
        }
    }

    $openLastOutputButton.IsEnabled = $hasLastOutput
    $openConvertedFolderButton.IsEnabled = (Test-Path -LiteralPath (Get-ConvertedRootPath))
}

function Finish-ConversionSession {
    param(
        [psobject]$Status = $null,
        [string]$FallbackErrorMessage = "Unknown conversion error.",
        [switch]$TerminateProcess
    )

    $conversionTimer.Stop()

    if ($script:currentTimeoutTimer) {
        $script:currentTimeoutTimer.Stop()
        $script:currentTimeoutTimer = $null
    }

    $successCount = 0
    $failedCount = $selectedFiles.Count
    $errorMessage = $FallbackErrorMessage
    $outputs = @()

    if (-not $Status -and $script:currentStatusPath -and (Test-Path -LiteralPath $script:currentStatusPath)) {
        try {
            $Status = [System.IO.File]::ReadAllText($script:currentStatusPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        }
        catch {
            $errorMessage = $_.Exception.Message
        }
    }

    if ($Status) {
        if ($Status.state -eq "completed") {
            $successCount = [int]$Status.successCount
            $failedCount = [int]$Status.failedCount
            $outputs = @($Status.outputs)
            $script:lastOutputPaths = @($outputs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }
        elseif ($Status.state -eq "failed") {
            $errorMessage = [string]$Status.message
        }
    }

    foreach ($output in @($outputs)) {
        if (-not [string]::IsNullOrWhiteSpace($output)) {
            Add-LogLine -LogBox $logTextBox -Message ("Output: " + $output)
        }
    }

    Update-ConversionProgress -CompletedCount ([Math]::Max($successCount, $selectedFiles.Count)) -TotalCount ([Math]::Max($selectedFiles.Count, 1)) -StatusText ("Finished. Success: {0}, Failed: {1}" -f $successCount, $failedCount)
    Add-LogLine -LogBox $logTextBox -Message ("Finished. Success: " + $successCount + ", Failed: " + $failedCount)

    if ($failedCount -gt 0) {
        Add-LogLine -LogBox $logTextBox -Message ("Failed: " + $errorMessage)
        [System.Windows.MessageBox]::Show("Finished. Success: $successCount, Failed: $failedCount.`r`n`r`n$errorMessage", "Conversion Failed") | Out-Null
    }
    else {
        [System.Windows.MessageBox]::Show("Finished. Success: $successCount, Failed: $failedCount. Output files are in the converted folder.", "Conversion Completed") | Out-Null
    }

    Set-ConversionUiState -IsBusy $false
    Update-OutputButtons
    Update-UiForFiles

    if ($TerminateProcess -and $script:currentConversionProcess -and -not $script:currentConversionProcess.HasExited) {
        try {
            Stop-Process -Id $script:currentConversionProcess.Id -Force -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    if ($TerminateProcess) {
        Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue |
            Where-Object { [string]::IsNullOrWhiteSpace($_.MainWindowTitle) } |
            ForEach-Object {
                try {
                    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                }
                catch {
                }
            }
    }

    if ($script:currentStatusPath -and (Test-Path -LiteralPath $script:currentStatusPath)) {
        Remove-Item -LiteralPath $script:currentStatusPath -Force -ErrorAction SilentlyContinue
    }
    if ($script:currentInputListPath -and (Test-Path -LiteralPath $script:currentInputListPath)) {
        Remove-Item -LiteralPath $script:currentInputListPath -Force -ErrorAction SilentlyContinue
    }

    $script:currentConversionProcess = $null
    $script:currentStatusPath = $null
    $script:currentInputListPath = $null
    $script:currentJobId = $null
}

function Read-CurrentStatusFile {
    if (-not $script:currentStatusPath -or -not (Test-Path -LiteralPath $script:currentStatusPath)) {
        return $null
    }

    try {
        return ([System.IO.File]::ReadAllText($script:currentStatusPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json)
    }
    catch {
        return $null
    }
}

function Convert-PresentationToPdfViaTemp {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,
        [Parameter(Mandatory)]
        [pscustomobject]$Tools
    )

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("converter_pdf_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $tempPdfPath = Join-Path $tempRoot ([System.IO.Path]::GetFileNameWithoutExtension($InputPath) + ".pdf")

    try {
        if ($Tools.LibreOffice) {
            $quotedOutputDirectory = '"' + $tempRoot.Replace('"', '""') + '"'
            $quotedInputPath = '"' + $InputPath.Replace('"', '""') + '"'
            $arguments = '--headless --convert-to pdf --outdir ' + $quotedOutputDirectory + ' ' + $quotedInputPath
            $process = Start-Process -FilePath $Tools.LibreOffice -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
            if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $tempPdfPath)) {
                throw "LibreOffice failed to export presentation to PDF."
            }
            return $tempPdfPath
        }

        if ($powerPointAvailable) {
            $powerPoint = $null
            $presentation = $null
            try {
                $powerPoint = New-Object -ComObject PowerPoint.Application
                $presentation = $powerPoint.Presentations.Open($InputPath, $true, $false, $false)
                $presentation.SaveAs($tempPdfPath, 32)
                if (-not (Test-Path -LiteralPath $tempPdfPath)) {
                    throw "PowerPoint failed to export presentation to PDF."
                }
                return $tempPdfPath
            }
            finally {
                if ($presentation) { $presentation.Close() }
                if ($powerPoint) { $powerPoint.Quit() }
                Ensure-ComObjectReleased -ComObject $presentation
                Ensure-ComObjectReleased -ComObject $powerPoint
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
        }

        throw "No presentation engine is available for image export."
    }
    catch {
        if (Test-Path -LiteralPath $tempPdfPath) {
            Remove-Item -LiteralPath $tempPdfPath -Force -ErrorAction SilentlyContinue
        }
        throw
    }
}

if ($SelfTest) {
    $tools = Get-InstalledTools
    Write-Output "SelfTest OK"
    Write-Output ("FFmpeg=" + $tools.Ffmpeg)
    Write-Output ("LibreOffice=" + $tools.LibreOffice)
    exit 0
}

if ($WorkerMode) {
    $script:tools = Get-InstalledTools
    $script:wordAvailable = Test-WordAvailable
    $script:powerPointAvailable = Test-PowerPointAvailable
    $script:pdfRenderAvailable = Test-PdfRenderAvailable

    try {
        $inputs = @()
        if (Test-Path -LiteralPath $InputListPath) {
            $inputs = @([System.IO.File]::ReadAllLines($InputListPath, [System.Text.Encoding]::UTF8) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }

        Write-WorkerStatus -Path $StatusPath -Data @{
            state = "running"
            message = "Conversion in progress"
            successCount = 0
            failedCount = 0
            outputs = @()
        }

        $outputs = Convert-Files -InputPaths $inputs -TargetExtension $TargetExtensionArg -Tools $tools -ImageMode $ImageModeArg
        Write-WorkerStatus -Path $StatusPath -Data @{
            state = "completed"
            message = "Conversion completed"
            successCount = $inputs.Count
            failedCount = 0
            outputs = @($outputs)
        }
        exit 0
    }
    catch {
        Write-WorkerStatus -Path $StatusPath -Data @{
            state = "failed"
            message = (Get-FriendlyErrorMessage -Exception $_.Exception)
            successCount = 0
            failedCount = 1
            outputs = @()
        }
        exit 1
    }
}

function Set-ConversionUiState {
    param(
        [Parameter(Mandatory)]
        [bool]$IsBusy
    )

    $convertButton.IsEnabled = (-not $IsBusy -and $selectedFiles.Count -gt 0 -and $targetFormatCombo.SelectedItem)
    $browseButton.IsEnabled = -not $IsBusy
    $dropZone.IsEnabled = -not $IsBusy
    $targetFormatCombo.IsEnabled = -not $IsBusy
    if ($imageModeCombo.Visibility -eq "Visible") {
        $imageModeCombo.IsEnabled = -not $IsBusy
    }
    $openLastOutputButton.IsEnabled = (-not $IsBusy -and ($script:lastOutputPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_) } | Measure-Object).Count -gt 0)
    $openConvertedFolderButton.IsEnabled = (-not $IsBusy -and (Test-Path -LiteralPath (Get-ConvertedRootPath)))
}

function Update-UiForFiles {
    $fileListBox.Items.Clear()
    $targetFormatCombo.Items.Clear()

    foreach ($file in $selectedFiles) {
        [void]$fileListBox.Items.Add($file)
    }

    if ($selectedFiles.Count -eq 0) {
        $dropHint.Text = "No files selected"
        $targetFormatCombo.IsEnabled = $false
        $convertButton.IsEnabled = $false
        Update-ConversionProgress -CompletedCount 0 -TotalCount 1 -StatusText "Ready to convert"
        Update-ImageModeVisibility
        return
    }

    $dropHint.Text = "Selected files: $($selectedFiles.Count)"
    $targets = Get-AvailableTargetsForFiles -Files $selectedFiles

    if ($targets.Count -eq 0) {
        $targetFormatCombo.IsEnabled = $false
        $convertButton.IsEnabled = $false
        Update-ConversionProgress -CompletedCount 0 -TotalCount $selectedFiles.Count -StatusText "No available target format"
        Add-LogLine -LogBox $logTextBox -Message "No supported target format is available for these files on this computer. Install FFmpeg, LibreOffice, or Office to unlock more conversions."
        Update-ImageModeVisibility
        return
    }

    foreach ($target in @($targets)) {
        [void]$targetFormatCombo.Items.Add([string]$target)
    }

    $targetFormatCombo.SelectedIndex = 0
    $targetFormatCombo.IsEnabled = $true
    $convertButton.IsEnabled = $true
    Update-ConversionProgress -CompletedCount 0 -TotalCount $selectedFiles.Count -StatusText ("Ready: 0/{0} files converted" -f $selectedFiles.Count)
    Update-ImageModeVisibility
}

function Set-SelectedFiles {
    param(
        [Parameter(Mandatory)]
        [string[]]$Files
    )

    $selectedFiles.Clear()

    foreach ($file in $Files) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
            continue
        }

        $profile = Get-FileTypeProfile -Path $file
        if ($profile) {
            $selectedFiles.Add($file)
        }
        else {
            Add-LogLine -LogBox $logTextBox -Message ("Skipped unsupported file: " + $file)
        }
    }

    Update-UiForFiles
}

$dropZone.Add_DragOver({
    if ($_.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $_.Effects = [System.Windows.DragDropEffects]::Copy
    }
    else {
        $_.Effects = [System.Windows.DragDropEffects]::None
    }
    $_.Handled = $true
})

$dropZone.Add_Drop({
    $files = $_.Data.GetData([System.Windows.DataFormats]::FileDrop)
    if ($files) {
        Set-SelectedFiles -Files $files
        Add-LogLine -LogBox $logTextBox -Message ("Loaded dropped files: " + $files.Count)
    }
})

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Multiselect = $true
    $dialog.Title = "Choose files to convert"
    $dialog.Filter = "Supported files|*.pdf;*.doc;*.docx;*.ppt;*.pptx;*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.mp4;*.mov;*.avi;*.mkv;*.wmv;*.webm;*.flv;*.m4v;*.mp3;*.wav;*.aac;*.flac;*.m4a;*.ogg;*.wma|All files|*.*"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Set-SelectedFiles -Files $dialog.FileNames
        Add-LogLine -LogBox $logTextBox -Message ("Loaded selected files: " + $dialog.FileNames.Count)
    }
})

$openLastOutputButton.Add_Click({
    $targetPath = $script:lastOutputPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
    if (-not $targetPath) {
        return
    }

    Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$targetPath`""
})

$openConvertedFolderButton.Add_Click({
    $convertedRoot = Get-ConvertedRootPath
    if (-not (Test-Path -LiteralPath $convertedRoot)) {
        New-Item -ItemType Directory -Path $convertedRoot -Force | Out-Null
    }

    Start-Process -FilePath "explorer.exe" -ArgumentList "`"$convertedRoot`""
})

$targetFormatCombo.Add_SelectionChanged({
    Update-ImageModeVisibility
})

$conversionTimer = New-Object System.Windows.Threading.DispatcherTimer
$conversionTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$conversionTimer.Add_Tick({
    if (-not $script:currentConversionProcess) {
        $conversionTimer.Stop()
        return
    }

    $status = $null
    $statusReady = $false

    $status = Read-CurrentStatusFile
    if ($status) {
        if ($status.state -eq "running") {
            Update-ConversionProgress -CompletedCount 0 -TotalCount ([Math]::Max($selectedFiles.Count, 1)) -StatusText "Conversion in progress"
        }
        elseif ($status.state -in @("completed", "failed")) {
            $statusReady = $true
        }
    }

    if ($statusReady) {
        Finish-ConversionSession -Status $status -TerminateProcess
        return
    }

    if ($script:currentConversionProcess.HasExited) {
        Finish-ConversionSession -Status $status
    }
})

$convertButton.Add_Click({
    if ($selectedFiles.Count -eq 0 -or -not $targetFormatCombo.SelectedItem) {
        return
    }

    $targetExtension = [string]$targetFormatCombo.SelectedItem
    $imageMode = "separate"
    $totalCount = $selectedFiles.Count
    Set-ConversionUiState -IsBusy $true
    Update-ConversionProgress -CompletedCount 0 -TotalCount $totalCount -StatusText ("Starting conversion: 0/{0}" -f $totalCount)

    Add-LogLine -LogBox $logTextBox -Message ("Converting to ." + $targetExtension + " for " + $selectedFiles.Count + " selected file(s)")
    if ($targetExtension -in @("png", "jpg")) {
        Add-LogLine -LogBox $logTextBox -Message "Image export mode: separate images"
    }

    $jobId = [guid]::NewGuid().ToString("N")
    $script:currentJobId = $jobId
    $script:currentInputListPath = Join-Path ([System.IO.Path]::GetTempPath()) ("converter_inputs_" + $jobId + ".txt")
    $script:currentStatusPath = Join-Path ([System.IO.Path]::GetTempPath()) ("converter_status_" + $jobId + ".json")
    [System.IO.File]::WriteAllLines($script:currentInputListPath, @($selectedFiles.ToArray()), [System.Text.UTF8Encoding]::new($true))

    $scriptPath = $PSCommandPath
    $args = @(
        "-ExecutionPolicy", "Bypass",
        "-STA",
        "-WindowStyle", "Hidden",
        "-File", $scriptPath,
        "-WorkerMode",
        "-InputListPath", $script:currentInputListPath,
        "-TargetExtensionArg", $targetExtension,
        "-ImageModeArg", $imageMode,
        "-StatusPath", $script:currentStatusPath
    )

    $script:currentConversionProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $args -WindowStyle Hidden -PassThru
    $conversionTimer.Start()
    if ($script:currentTimeoutTimer) {
        $script:currentTimeoutTimer.Stop()
    }
    $timeoutTimer = New-Object System.Windows.Threading.DispatcherTimer
    $timeoutTimer.Interval = [TimeSpan]::FromSeconds(120)
    $script:currentTimeoutTimer = $timeoutTimer
    $jobIdForTimer = $jobId
    $timeoutTimer.Add_Tick({
        if ($timeoutTimer) {
            $timeoutTimer.Stop()
        }
        if ($script:currentJobId -ne $jobIdForTimer) {
            return
        }

        $script:currentTimeoutTimer = $null
        $status = Read-CurrentStatusFile
        if ($status -and $status.state -in @("completed", "failed")) {
            Finish-ConversionSession -Status $status -TerminateProcess
            return
        }
        Add-LogLine -LogBox $logTextBox -Message "Failed: conversion timed out after 120 seconds."
        Finish-ConversionSession -FallbackErrorMessage "Conversion timed out and was stopped automatically after 120 seconds.`r`nThis usually means the backend converter is stuck in the background." -TerminateProcess
    }.GetNewClosure())
    $timeoutTimer.Start()
})

Refresh-EngineInfo
Update-OutputButtons
Add-LogLine -LogBox $logTextBox -Message "Application started. Drop files into the window or use Choose Files."
[void]$window.ShowDialog()
