param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = 'Stop'
$target = (Resolve-Path -LiteralPath $Path).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'runs')) + [IO.Path]::DirectorySeparatorChar
if (-not $target.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Only execution copies under runs/ can be recalculated.'
}
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($target, 0, $false)
    $excel.CalculateFullRebuild()
    $book.Save()
    Write-Output 'recalculated'
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
