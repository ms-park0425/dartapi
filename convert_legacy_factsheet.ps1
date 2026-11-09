param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = 'Stop'
$target = (Resolve-Path -LiteralPath $Path).Path
$repo = [IO.Path]::GetFullPath($PSScriptRoot) + [IO.Path]::DirectorySeparatorChar
$permitted = @((Join-Path $repo 'runs'), (Join-Path $repo 'data'))
$safe = $false
foreach ($directory in $permitted) {
    $prefix = [IO.Path]::GetFullPath($directory) + [IO.Path]::DirectorySeparatorChar
    if ($target.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { $safe = $true }
}
if (-not $safe -or [IO.Path]::GetExtension($target) -ne '.xls') { throw 'Only cached XLS source copies may be converted.' }
$output = [IO.Path]::ChangeExtension($target, '.xlsx')
if (Test-Path -LiteralPath $output) { throw 'Existing conversion must not be overwritten.' }
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($target, 0, $true)
    $book.SaveAs($output, 51)
    Write-Output 'converted'
} finally {
    if ($null -ne $book) {
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
