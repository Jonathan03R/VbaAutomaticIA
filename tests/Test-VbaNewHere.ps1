$ErrorActionPreference = 'Stop'
$entry = Join-Path (Split-Path -Parent $PSScriptRoot) 'vba.ps1'
Import-Module (Join-Path $PSScriptRoot '..\tools\VbaDev.psm1') -Force
$module = Get-Module VbaDev
$temp = & $module { New-VbaTemp }
try {
    Push-Location -LiteralPath $temp
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $entry new . -Once
        if ($LASTEXITCODE -ne 0) { throw 'new . failed' }
        $name = Split-Path -Leaf $temp
        if (-not (Test-Path -LiteralPath (Join-Path $temp "excel\$name.xlsm"))) { throw 'Workbook not created in current folder' }
        if (-not (Test-Path -LiteralPath (Join-Path $temp 'src\modules\Main.bas'))) { throw 'Source not created in current folder' }
        Write-Host 'PASS: new . creates workbook and source in the existing current folder'
        $before = (Get-FileHash -LiteralPath (Join-Path $temp "excel\$name.xlsm")).Hash
        $rejected = $false
        try { New-VbaProject -Path '.' } catch { $rejected = $_.Exception.Message -like '*no esta vacia*' }
        if (-not $rejected -or $before -ne (Get-FileHash -LiteralPath (Join-Path $temp "excel\$name.xlsm")).Hash) {
            throw 'Existing project was not protected'
        }
        Write-Host 'PASS: repeated new . refuses to overwrite the existing project'
    } finally { Pop-Location }
} finally { & $module { param($Path) Remove-VbaTemp $Path } $temp }
