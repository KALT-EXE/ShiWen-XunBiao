param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][string]$Address
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null

try {
    if (-not [System.IO.File]::Exists($WorkbookPath)) {
        throw "找不到表格文件：$WorkbookPath"
    }

    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
    }

    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    $target = [System.IO.Path]::GetFullPath($WorkbookPath)
    foreach ($book in $excel.Workbooks) {
        try {
            if ([System.IO.Path]::GetFullPath($book.FullName) -eq $target) {
                $workbook = $book
                break
            }
        }
        catch {
            continue
        }
    }
    if ($null -eq $workbook) {
        $workbook = $excel.Workbooks.Open($target)
    }

    $sheet = $workbook.Worksheets.Item($SheetName)
    if ($null -eq $sheet) {
        throw "工作表中不存在：$SheetName"
    }
    $sheet.Activate()
    $range = $sheet.Range($Address)
    $excel.Goto($range, $true) | Out-Null
    $range.Select() | Out-Null
    $window = $excel.ActiveWindow
    $window.Activate()
    $excel.Visible = $true

    [Console]::Out.Write("OK")
    exit 0
}
catch {
    $message = $_.Exception.Message
    if ($_.Exception.InnerException) {
        $message = $_.Exception.InnerException.Message
    }
    [Console]::Error.Write($message)
    exit 1
}
