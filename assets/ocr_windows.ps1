param(
    [Parameter(Mandatory = $true)][string]$ImagePath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Write-JsonResult {
    param($Payload)
    $directory = [System.IO.Path]::GetDirectoryName($OutputPath)
    if ($directory -and -not [System.IO.Directory]::Exists($directory)) {
        [System.IO.Directory]::CreateDirectory($directory) | Out-Null
    }
    $json = $Payload | ConvertTo-Json -Depth 5 -Compress
    [System.IO.File]::WriteAllText(
        $OutputPath,
        $json,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

try {
    if (-not [System.IO.File]::Exists($ImagePath)) {
        throw "找不到待识别的图片：$ImagePath"
    }

    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
            $_.Name -eq 'AsTask' -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
        })[0]

    function Await($WinRtTask, $ResultType) {
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(-1) | Out-Null
        $netTask.Result
    }

    [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
    [Windows.Storage.Streams.DataWriter, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime] | Out-Null
    [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
    [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null

    $stream = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
    $writer = [Windows.Storage.Streams.DataWriter]::new($stream)
    $bytes = [System.IO.File]::ReadAllBytes($ImagePath)
    $writer.WriteBytes($bytes)
    $null = Await ($writer.StoreAsync()) ([uint32])
    $writer.DetachStream() | Out-Null
    $stream.Seek(0) | Out-Null

    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

    $available = @(
        [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages |
            ForEach-Object { $_.LanguageTag }
    )
    $engine = $null
    foreach ($tag in @('zh-Hans', 'zh-CN', 'zh-Hant', 'en-US')) {
        try {
            $language = [Windows.Globalization.Language]::new($tag)
            $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
        }
        catch {
            $engine = $null
        }
        if ($null -ne $engine) { break }
    }
    if ($null -eq $engine) {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    }
    if ($null -eq $engine) {
        throw "系统没有可用的 OCR 语言包。请在 Windows 设置中安装 中文（简体） OCR 组件。"
    }

    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $lines = @()
    foreach ($line in $result.Lines) {
        if ($line.Text) {
            $lines += [string]$line.Text
        }
    }

    Write-JsonResult ([ordered]@{
            ok                  = $true
            language            = [string]$engine.RecognizerLanguage.LanguageTag
            available_languages = $available
            text                = [string]$result.Text
            lines               = $lines
        })
}
catch {
    $message = $_.Exception.Message
    if ($_.Exception.InnerException) {
        $message = $_.Exception.InnerException.Message
    }
    try {
        Write-JsonResult ([ordered]@{
                ok        = $false
                error     = [string]$message
                lines     = @()
                text      = ''
                language  = ''
            })
    }
    catch {
        # 输出文件都写不了时，把错误留给调用方读取 stderr。
        [Console]::Error.WriteLine($message)
    }
    exit 1
}
