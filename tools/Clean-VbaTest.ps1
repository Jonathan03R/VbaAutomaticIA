[CmdletBinding()]
param(
    [string] $Workbook = 'C:\Mi_D\VbaAutomaticIA\test.xlsm',
    [string] $SourceRoot = 'C:\Mi_D\VbaAutomaticIA\src'
)

$ErrorActionPreference = 'Stop'
$keepModules = @('ModuleEjemplo', 'Módulo1')
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = "$Workbook.before-cleanup-$stamp"
Copy-Item -LiteralPath $Workbook -Destination $backup -Force
Write-Host "Backup creado: $backup" -ForegroundColor DarkGray

$excel = $null
$book = $null
$ownedExcel = $false
try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        $excel = New-Object -ComObject Excel.Application
        $ownedExcel = $true
        $excel.Visible = $true
    }
    $excel.DisplayAlerts = $false
    $fullBook = [IO.Path]::GetFullPath($Workbook)
    foreach ($candidate in @($excel.Workbooks)) {
        if ([IO.Path]::GetFullPath($candidate.FullName) -ieq $fullBook) {
            $book = $candidate
            break
        }
    }
    if (-not $book) { $book = $excel.Workbooks.Open($fullBook) }

    $removed = @()
    foreach ($component in @($book.VBProject.VBComponents)) {
        if ([int]$component.Type -eq 1 -and $component.Name -notin $keepModules) {
            $removed += $component.Name
            $book.VBProject.VBComponents.Remove($component)
        }
    }
    $book.Save()
    Write-Host "Módulos estándar eliminados del libro: $($removed.Count)" -ForegroundColor Green
    if ($removed.Count -gt 0) { $removed | ForEach-Object { Write-Host "  $_" } }
} finally {
    if ($book) { [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
    if ($ownedExcel -and $excel) { $excel.Quit() }
    if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$deletedFiles = Get-ChildItem -LiteralPath $SourceRoot -Filter '*.bas' -File | Where-Object {
    $_.BaseName -notin $keepModules
}
foreach ($file in @($deletedFiles)) {
    Remove-Item -LiteralPath $file.FullName -Force
}
Write-Host "Archivos .bas generados eliminados de src: $($deletedFiles.Count)" -ForegroundColor Green
