$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\tools\VbaDev.psm1') -Force
& (Get-Module VbaDev) {
    $root = Join-Path ([IO.Path]::GetTempPath()) ('vba-recovery-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $root
    $path = Join-Path $root 'book.xlsm'
    [IO.File]::WriteAllText($path, 'original bytes')
    $components = [pscustomobject]@{ Count = 1 }
    $components | Add-Member ScriptMethod Item { param($name) return [pscustomobject]@{ Type = 100 } }
    $book = [pscustomobject]@{
        ReadOnly = $false; Saved = $true; CodeName = 'ThisWorkbook'; Sheets = @()
        VBProject = [pscustomobject]@{ Protection = 0; VBComponents = $components }
        Path = $path; Closed = $false
    }
    $book | Add-Member ScriptMethod Save {
        [IO.File]::WriteAllText($this.Path, 'partial failed save')
        throw 'Simulated disk error'
    }
    $book | Add-Member ScriptMethod Close { param($save) $this.Closed = $true }
    $books = [pscustomobject]@{ Book = $book; Owner = $null }
    $books | Add-Member ScriptMethod Open {
        param($path, $links, $readOnly)
        if ($this.Owner.EnableEvents -or $this.Owner.AutomationSecurity -ne 3) { throw 'Unsafe repair opening' }
        return $this.Book
    }
    $excel = [pscustomobject]@{ EnableEvents = $true; AutomationSecurity = 2; Workbooks = $books }
    $books.Owner = $excel
    $failed = $false
    try { Repair-VbaDocumentModules $excel $path @('HojaNueva') } catch { $failed = $true }
    if (-not $failed -or -not $book.Closed -or [IO.File]::ReadAllText($path) -ne 'original bytes') {
        throw 'FAIL: failed save did not restore original bytes after closing the workbook'
    }
    if (-not $excel.EnableEvents -or $excel.AutomationSecurity -ne 2) { throw 'FAIL: application settings were not restored' }
    $backups = @(Get-ChildItem (Join-Path $root 'backup') -Filter '*.xlsm')
    if ($backups.Count -ne 1 -or [IO.File]::ReadAllText($backups[0].FullName) -ne 'original bytes') { throw 'FAIL: backup not preserved' }
    Write-Host 'PASS: failed save restores original bytes and retains backup'
    Write-Host 'PASS: repair opens with macros and events disabled; restores prior settings'
    $book.ReadOnly = $true
    $failed = $false
    try { Repair-VbaDocumentModules $excel $path @('HojaNueva') } catch { $failed = $true }
    if (-not $failed -or @(Get-ChildItem (Join-Path $root 'backup')).Count -ne 1) { throw 'FAIL: read-only workbook modified' }
    Write-Host 'PASS: read-only workbook rejected without another backup or save'
    $failed = $false
    try { Get-VbaStoredStreamNames (New-Object byte[] 512) } catch { $failed = $true }
    if (-not $failed) { throw 'FAIL: corrupt VBA accepted' }
    Write-Host 'PASS: corrupt compound-file header rejected'
}
