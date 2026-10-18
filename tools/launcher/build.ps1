$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path -LiteralPath $csc)) {
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path -LiteralPath $csc)) {
    throw "未找到 .NET Framework C# 编译器 (csc.exe)。"
}

$launcher = Join-Path $root "识文寻表.exe"
$icon = Join-Path $root "assets\app.ico"
$source = Join-Path $PSScriptRoot "launcher.cs"

& $csc /nologo /target:winexe /platform:anycpu /optimize+ `
    "/out:$launcher" "/win32icon:$icon" `
    /reference:System.Windows.Forms.dll $source

if ($LASTEXITCODE -ne 0) {
    throw "启动器编译失败，退出码：$LASTEXITCODE"
}

Write-Host "已生成：$launcher"
