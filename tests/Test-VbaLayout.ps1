$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\tools\VbaDev.psm1') -Force
& (Get-Module VbaDev) {
    $temp = New-VbaTemp
    try {
        $book = Join-Path $temp 'Example.xlsm'
        $meta = Join-Path $temp '.vba'
        New-Item -ItemType Directory -Path $meta | Out-Null
        $config = Join-Path $meta 'project.json'
        Write-VbaText $book 'fixture bytes'
        Write-VbaText $config '{"Workbook":"Example.xlsm","Source":"src","Flat":false}'
        $next = '{"Workbook":"excel\\Example.xlsm","Source":"src","Flat":false}'
        $hash = (Get-FileHash -LiteralPath $book).Hash
        $lock = [IO.File]::Open($book, 'Open', 'ReadWrite', 'None')
        try {
            $blocked = $false
            try { Move-VbaWorkbookToExcel $book $temp $config $next } catch { $blocked = $_.Exception.Message -like '*Cierra Excel*' }
            if (-not $blocked) { throw 'Open workbook was not rejected' }
        } finally { $lock.Dispose() }
        $moved = Move-VbaWorkbookToExcel $book $temp $config $next
        if ((Test-Path -LiteralPath $book) -or (Get-FileHash -LiteralPath $moved).Hash -ne $hash) { throw 'Move changed or duplicated workbook' }
        if ((Get-Content -LiteralPath $config -Raw | ConvertFrom-Json).Workbook -ne 'excel\Example.xlsm') { throw 'Configuration not migrated' }
        if ((Move-VbaWorkbookToExcel $moved $temp $config $next) -ne $moved) { throw 'Repeated migration changed location' }
        Write-Host 'PASS: locked workbook rejected; closed workbook moved unchanged and configuration updated; repeat is stable'
        Write-VbaText $book 'second workbook'
        $blocked = $false
        try { Move-VbaWorkbookToExcel $book $temp $config $next } catch { $blocked = $_.Exception.Message -like '*Ya existe*' }
        if (-not $blocked -or (Get-FileHash -LiteralPath $moved).Hash -ne $hash -or -not (Test-Path -LiteralPath $book)) { throw 'Collision did not preserve both files' }
        Write-Host 'PASS: destination collision preserves both workbooks'
        $rollbackBook = Join-Path $temp 'Rollback.xlsm'
        Write-VbaText $rollbackBook 'rollback fixture'
        $blocked = $false
        try { Move-VbaWorkbookToExcel $rollbackBook $temp (Join-Path $temp 'missing\project.json') $next } catch { $blocked = $true }
        if (-not $blocked -or -not (Test-Path -LiteralPath $rollbackBook) -or (Test-Path -LiteralPath (Join-Path $temp 'excel\Rollback.xlsm'))) { throw 'Failed configuration write did not roll back move' }
        Write-Host 'PASS: failed configuration write restores original workbook location'
    } finally { Remove-VbaTemp $temp }
}
