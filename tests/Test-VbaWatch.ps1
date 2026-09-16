$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'tools\VbaDev.psm1') -Force
$module = Get-Module VbaDev
& $module {
    param($Repo)
    function Read-Log([string] $Path) {
        if (-not (Test-Path -LiteralPath $Path)) { return '' }
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $reader = New-Object IO.StreamReader($stream)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    $temp = New-VbaTemp
    $process = $null
    $connection = $null
    try {
        # Existing-workbook case: copy only the fixture, preserving the repository workbook.
        $bookPath = Join-Path $temp 'Existing.xlsm'
        Copy-Item -LiteralPath (Join-Path $Repo 'test.xlsm') -Destination $bookPath
        $log = Join-Path $temp 'watch.log'
        $errorLog = Join-Path $temp 'watch.err'
        $entry = Join-Path $Repo 'vba.ps1'
        $process = Start-Process powershell.exe -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $entry + '"'),
            'dev', ('"' + $bookPath + '"')
        ) -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError $errorLog -PassThru
        $null = $process.Handle # Retain the handle so ExitCode remains available after exit.
        $deadline = [DateTime]::UtcNow.AddSeconds(30)
        do {
            Start-Sleep -Milliseconds 300
            if ($process.HasExited) { throw ((Read-Log $log) + (Read-Log $errorLog)) }
            $ready = (Read-Log $log) -match 'Ctrl\+C'
        } until ($ready -or [DateTime]::UtcNow -gt $deadline)
        if (-not $ready) { throw 'Watcher did not start within 30 seconds.' }
        Write-Host 'PASS: CLI extracts existing workbook and starts watching'
        $bookPath = Join-Path $temp 'excel\Existing.xlsm'
        if (-not (Test-Path -LiteralPath $bookPath)) { throw 'Existing workbook was not organized into excel.' }
        $source = Join-Path $temp 'src'
        $files = @(Get-ChildItem -LiteralPath (Join-Path $source 'modules') -Filter '*.bas')
        $file = $files[0]
        $marker = "' watcher-integration-marker"
        Write-VbaText $file.FullName ([IO.File]::ReadAllText($file.FullName, $script:Utf8) + "`r`n$marker`r`n")
        $expected = (Get-VbaSnapshot $source)[$file.BaseName].Hash
        $deadline = [DateTime]::UtcNow.AddSeconds(20)
        do {
            Start-Sleep -Milliseconds 300
            $state = Read-VbaState (Join-Path $temp '.vba\state.json')
            $synced = $state.Source[$file.BaseName].Hash -eq $expected
        } until ($synced -or [DateTime]::UtcNow -gt $deadline)
        if (-not $synced) { throw ('Watcher did not synchronize: ' + (Read-Log $log) + (Read-Log $errorLog)) }
        Write-Host 'PASS: watcher detects a saved file and persists synchronization'
        # Bind the exact temporary workbook even when another Excel instance is active.
        $watchedBook = [Runtime.InteropServices.Marshal]::BindToMoniker($bookPath)
        $connection = @{ Book = $watchedBook; Excel = $watchedBook.Application; Project = $watchedBook.VBProject; Owned = $false }
        # This Excel instance was created by this test's watcher for a unique temporary file.
        $connection.Owned = $true
        $component = $connection.Project.VBComponents.Item($file.BaseName)
        if (-not $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines).Contains($marker)) {
            throw 'Open Excel does not contain the saved source edit.'
        }
        Write-Host 'PASS: same open Excel workbook contains the live edit'
        $excelMarker = "' live-edit-from-Excel"
        $component.CodeModule.AddFromString($excelMarker)
        $connection.Book.Save()
        $deadline = [DateTime]::UtcNow.AddSeconds(20)
        do {
            Start-Sleep -Milliseconds 300
            $exported = [IO.File]::ReadAllText($file.FullName, $script:Utf8).Contains($excelMarker)
        } until ($exported -or [DateTime]::UtcNow -gt $deadline)
        if (-not $exported) { throw ('Excel edit was not exported: ' + (Read-Log $log) + (Read-Log $errorLog)) }
        Write-Host 'PASS: saving an edit in the Excel editor updates the source file'
        $returnMarker = "' source-edit-after-Excel"
        Write-VbaText $file.FullName ([IO.File]::ReadAllText($file.FullName, $script:Utf8) + "`r`n$returnMarker`r`n")
        $expected = (Get-VbaSnapshot $source)[$file.BaseName].Hash
        $deadline = [DateTime]::UtcNow.AddSeconds(20)
        do {
            Start-Sleep -Milliseconds 300
            $state = Read-VbaState (Join-Path $temp '.vba\state.json')
            $synced = $state.Source[$file.BaseName].Hash -eq $expected
        } until ($synced -or [DateTime]::UtcNow -gt $deadline)
        if (-not $synced) { throw 'Source edit after Excel export was not imported.' }
        $component = $connection.Project.VBComponents.Item($file.BaseName)
        $code = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
        if (-not $code.Contains($excelMarker) -or -not $code.Contains($returnMarker)) { throw 'Round trip lost one of the edits.' }
        Write-Host 'PASS: editing source after Excel preserves both edits in the workbook'
        $connection.Project.VBComponents.Remove($component)
        $connection.Book.Save()
        $deadline = [DateTime]::UtcNow.AddSeconds(20)
        do {
            Start-Sleep -Milliseconds 300
            $removed = -not (Test-Path -LiteralPath $file.FullName)
        } until ($removed -or [DateTime]::UtcNow -gt $deadline)
        if (-not $removed) { throw ('Watcher did not propagate Excel deletion: ' + (Read-Log $log) + (Read-Log $errorLog)) }
        Write-Host 'PASS: deleting a module in Excel and saving removes its source file'
        # Closing Excel's workbook must release the watcher, even if its save changes mtime.
        $connection.Book.Close($false)
        if (-not $process.WaitForExit(10000)) { throw 'Watcher did not exit after closing the workbook.' }
        $process.Refresh()
        if ($process.ExitCode -ne 0) { throw ("Watcher exit code '$($process.ExitCode)': " + (Read-Log $log) + (Read-Log $errorLog)) }
        if ((Read-Log $log) -notmatch 'Libro cerrado') { throw 'Watcher did not report workbook closure.' }
        Write-Host 'PASS: workbook closes while watcher is active and watcher exits cleanly'
        $connection.Excel.Quit()
        $connection.Owned = $false
    } finally {
        if ($process -and -not $process.HasExited) { Stop-Process -Id $process.Id; $process.WaitForExit() }
        Disconnect-VbaWorkbook $connection -Close
        try { Remove-VbaTemp $temp } catch { Write-Warning "Temporary test files retained: $temp. $($_.Exception.Message)" }
    }
} $repo
