Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$script:Utf8 = New-Object Text.UTF8Encoding($false, $true)
$script:Ansi = [Text.Encoding]::GetEncoding(
    [Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage,
    (New-Object Text.EncoderExceptionFallback), (New-Object Text.DecoderExceptionFallback))

function Write-VbaText([string] $Path, [string] $Text) {
    [IO.File]::WriteAllText($Path, $Text, $script:Utf8)
}

function New-VbaTemp {
    $path = Join-Path ([IO.Path]::GetTempPath()) ('vba-dev-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $path | Out-Null
    return $path
}

function Remove-VbaTemp([string] $Path) {
    $full = [IO.Path]::GetFullPath($Path)
    $parent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
    if ([IO.Path]::GetDirectoryName($full) -ine $parent -or [IO.Path]::GetFileName($full) -notlike 'vba-dev-*') {
        throw "Unsafe temporary directory: $Path"
    }
    if (Test-Path -LiteralPath $full) { Remove-Item -LiteralPath $full -Recurse -Force }
}

function Get-VbaLayout([int] $Type) {
    switch ($Type) {
        1 { return @('modules', '.bas') }
        2 { return @('classes', '.cls') }
        3 { return @('forms', '.frm') }
        100 { return @('documents', '.cls') }
        default { throw "Unsupported VBA component type: $Type" }
    }
}

function Export-VbaProject($Project, [string] $Root, [switch] $Flat) {
    # Root is always a fresh staging directory; Export cannot overwrite files.
    foreach ($component in @($Project.VBComponents)) {
        $layout = Get-VbaLayout ([int]$component.Type)
        $folder = $Root
        if (-not $Flat) { $folder = Join-Path $Root $layout[0] }
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $target = Join-Path $folder ($component.Name + $layout[1])
        $component.Export($target)
        $text = [IO.File]::ReadAllText($target, $script:Ansi)
        Write-VbaText $target $text
    }
}

function Get-VbaSnapshot([string] $Root) {
    $map = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $Root -File -Recurse | Where-Object { $_.Extension -in '.bas', '.cls', '.frm' })) {
        $text = [IO.File]::ReadAllText($file.FullName, $script:Utf8)
        $match = [regex]::Match($text, '(?m)^Attribute VB_Name = "([^"]+)"\s*$')
        if (-not $match.Success) { throw "File '$($file.Name)' must contain Attribute VB_Name." }
        $name = $match.Groups[1].Value
        if ($map.ContainsKey($name)) { throw "Duplicate VBA component name: $name" }
        $relative = $file.FullName.Substring($Root.TrimEnd('\').Length + 1)
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $textHash = $hash
        $binaryHash = ''
        $binary = [IO.Path]::ChangeExtension($file.FullName, '.frx')
        if ($file.Extension -eq '.frm') {
            if ($text -match 'OleObjectBlob\s*=\s*"([^"]+)"' -and $Matches[1] -cne ($file.BaseName + '.frx')) {
                throw "Form '$($file.Name)' must reference '$($file.BaseName).frx' in the same folder."
            }
            if ($text -match 'OleObjectBlob' -and -not (Test-Path -LiteralPath $binary)) {
                throw "Missing form resource: $binary"
            }
            if (Test-Path -LiteralPath $binary) {
                $binaryHash = (Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash
                $hash += ':' + $binaryHash
            }
        }
        $map[$name] = @{ Path = $relative; Hash = $hash; TextHash = $textHash; BinaryHash = $binaryHash }
    }
    return $map
}

function Read-VbaState([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $json = [IO.File]::ReadAllText($Path, $script:Utf8) | ConvertFrom-Json
    $state = @{ Source = @{}; Book = @{} }
    foreach ($side in @('Source', 'Book')) {
        foreach ($entry in $json.$side.PSObject.Properties) {
            $state[$side][$entry.Name] = @{
                Path = $entry.Value.Path; Hash = $entry.Value.Hash
                TextHash = $entry.Value.TextHash; BinaryHash = $entry.Value.BinaryHash
            }
        }
    }
    return $state
}

function Get-VbaFingerprint($Snapshot) {
    return ((@($Snapshot.Keys | Sort-Object) | ForEach-Object {
        $_ + '|' + $Snapshot[$_].Path + '|' + $Snapshot[$_].Hash
    }) -join ';')
}

function Save-VbaState([string] $Path, $Source, $Book) {
    $temp = $Path + '.tmp'
    Write-VbaText $temp (@{ Source = $Source; Book = $Book } | ConvertTo-Json -Depth 6)
    Move-Item -LiteralPath $temp -Destination $Path -Force
}

function Get-VbaChangePlan($Source, $Book, $State, [switch] $Bidirectional) {
    $actions = @()
    $names = @(@($Source.Keys) + @($Book.Keys) + @($State.Source.Keys) + @($State.Book.Keys) | Sort-Object -Unique)
    foreach ($name in $names) {
        $sourceHash = if ($Source.ContainsKey($name)) { $Source[$name].Hash } else { '' }
        # FRX exports contain changing timestamps and padding. Compare Excel's text,
        # and preserve its current binary resources on external code-only edits.
        $bookHash = if ($Book.ContainsKey($name)) { $Book[$name].TextHash } else { '' }
        $baseSource = if ($State.Source.ContainsKey($name)) { $State.Source[$name].Hash } else { '' }
        $baseBook = if ($State.Book.ContainsKey($name)) { $State.Book[$name].TextHash } else { '' }
        $sourceChanged = $sourceHash -ne $baseSource
        $bookChanged = $bookHash -ne $baseBook
        if ($sourceChanged) { $actions += @{ Name = $name; Direction = 'Push' } }
        elseif ($Source.ContainsKey($name)) {
            if (-not $Book.ContainsKey($name)) {
                # An Excel deletion propagates only when no local edits are pending.
                $direction = if ($State.Book.ContainsKey($name)) { 'Pull' } else { 'Push' }
                $actions += @{ Name = $name; Direction = $direction }
            }
            elseif ($bookChanged) {
                $direction = if ($Bidirectional) { 'Pull' } else { 'Push' }
                $actions += @{ Name = $name; Direction = $direction }
            }
        }
        elseif ($Book.ContainsKey($name)) {
            # First appearance of a component created inside Excel.
            $actions += @{ Name = $name; Direction = 'Pull' }
        }
    }
    return $actions
}

function Get-VbaDocumentCode([string] $Text) {
    # Strip the export preamble only before VB_Name. Preserve every End Sub/Function.
    $lines = $Text -split '\r?\n'
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^Attribute VB_Name = ') { $start = $i + 1; break }
    }
    if ($start -lt 0) { throw 'Missing Attribute VB_Name in document module.' }
    $code = @()
    for ($i = $start; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -notmatch '^Attribute ') { $code += $lines[$i] }
    }
    return ($code -join "`r`n")
}

function Copy-VbaComponent([string] $From, [string] $To, $Entry, $OldEntry) {
    if ($OldEntry) {
        $old = Join-Path $To $OldEntry.Path
        if (-not $Entry -or $OldEntry.Path -ine $Entry.Path) {
            if (Test-Path -LiteralPath $old) { Remove-Item -LiteralPath $old -Force }
            if ([IO.Path]::GetExtension($old) -eq '.frm') {
                $oldBinary = [IO.Path]::ChangeExtension($old, '.frx')
                if (Test-Path -LiteralPath $oldBinary) { Remove-Item -LiteralPath $oldBinary -Force }
            }
        }
    }
    if (-not $Entry) { return }
    $target = Join-Path $To $Entry.Path
    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $From $Entry.Path) -Destination $target -Force
    if ([IO.Path]::GetExtension($target) -eq '.frm') {
        $binary = [IO.Path]::ChangeExtension((Join-Path $From $Entry.Path), '.frx')
        $targetBinary = [IO.Path]::ChangeExtension($target, '.frx')
        if (Test-Path -LiteralPath $binary) { Copy-Item -LiteralPath $binary -Destination $targetBinary -Force }
        elseif (Test-Path -LiteralPath $targetBinary) { Remove-Item -LiteralPath $targetBinary -Force }
    }
}

function Copy-VbaEntryWithSourcePath($BookEntry, $SourceEntry) {
    if (-not $BookEntry) { return $null }
    $entry = @{}
    foreach ($key in $BookEntry.Keys) { $entry[$key] = $BookEntry[$key] }
    # Excel knows the component type, but not the user's folder architecture.
    if ($SourceEntry) { $entry.Path = $SourceEntry.Path }
    return $entry
}

function Copy-VbaExportToSourcePath([string] $ExportRoot, [string] $SourceRoot, $BookEntry, $SourceEntry) {
    if (-not $BookEntry) {
        Copy-VbaComponent $ExportRoot $SourceRoot $null $SourceEntry
        return $null
    }
    $entry = Copy-VbaEntryWithSourcePath $BookEntry $SourceEntry
    if ($SourceEntry -and $SourceEntry.Path -ine $entry.Path) {
        Copy-VbaComponent $ExportRoot $SourceRoot $null $SourceEntry
    }
    $origin = Join-Path $ExportRoot $BookEntry.Path
    $target = Join-Path $SourceRoot $entry.Path
    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    Copy-Item -LiteralPath $origin -Destination $target -Force
    if ([IO.Path]::GetExtension($origin) -ieq '.frm') {
        $originBinary = [IO.Path]::ChangeExtension($origin, '.frx')
        $targetBinary = [IO.Path]::ChangeExtension($target, '.frx')
        if (Test-Path -LiteralPath $originBinary) { Copy-Item -LiteralPath $originBinary -Destination $targetBinary -Force }
        elseif (Test-Path -LiteralPath $targetBinary) { Remove-Item -LiteralPath $targetBinary -Force }
    }
    return $entry
}

function Import-VbaComponent($Project, [string] $Name, [string] $Root, $Entry) {
    $existing = @($Project.VBComponents) | Where-Object { $_.Name -ieq $Name } | Select-Object -First 1
    if (-not $Entry) {
        if ($existing -and [int]$existing.Type -eq 100) { throw "Cannot delete document '$Name'. Clear its code instead." }
        if ($existing) { $Project.VBComponents.Remove($existing) }
        return
    }
    $path = Join-Path $Root $Entry.Path
    $text = [IO.File]::ReadAllText($path, $script:Utf8)
    if ($existing -and [int]$existing.Type -eq 100) {
        $code = Get-VbaDocumentCode $text
        $module = $existing.CodeModule
        if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
        if ($code.Trim().Length -gt 0) { $module.AddFromString($code) }
        return
    }
    if ($Entry.Path -match '(^|[\\/])documents[\\/]') { throw "Document '$Name' does not exist in this workbook." }
    $temp = New-VbaTemp
    try {
        $importPath = Join-Path $temp ([IO.Path]::GetFileName($path))
        # Convert before removing the existing component; unsupported characters fail safely.
        [IO.File]::WriteAllText($importPath, $text, $script:Ansi)
        $binary = [IO.Path]::ChangeExtension($path, '.frx')
        if ([IO.Path]::GetExtension($path) -eq '.frm' -and (Test-Path -LiteralPath $binary)) {
            Copy-Item -LiteralPath $binary -Destination ([IO.Path]::ChangeExtension($importPath, '.frx'))
        }
        if ($existing) { $Project.VBComponents.Remove($existing) }
        $imported = $Project.VBComponents.Import($importPath)
        if ($imported.Name -cne $Name) { throw "Excel imported '$Name' as '$($imported.Name)'." }
    } finally { Remove-VbaTemp $temp }
}

function Wait-VbaProjectAccess($Project,
    [ValidateRange(1, 3600)] [int] $TimeoutSeconds = 300,
    [ValidateRange(10, 1000)] [int] $PollMilliseconds = 250) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $protection = $Project.Protection
            if ($null -ne $protection -and [int]$protection -eq 0) {
                # Office can remain busy after accepting the password. Wait until
                # the project is actually readable, not just marked unlocked.
                $count = $Project.VBComponents.Count
                if ($null -ne $count) { return }
            }
        } catch {
            # Modal Office dialogs can reject COM calls while the user is typing.
        }
        Start-Sleep -Milliseconds $PollMilliseconds
    }
    throw "Se agoto la espera de $TimeoutSeconds segundos para acceder al VBA. Excel permanece abierto. Termina el desbloqueo, cierra sus cuadros de dialogo y vuelve a ejecutar el comando."
}

function Unlock-VbaProject($Excel, $Project, [int] $TimeoutSeconds = 300) {
    if ([int]$Project.Protection -eq 0) { return }
    $vbe = $Excel.VBE
    $wasVisible = $vbe.MainWindow.Visible
    $Excel.Visible = $true
    $vbe.MainWindow.Visible = $true
    $vbe.ActiveVBProject = $Project
    # Project Properties invokes Office's own password dialog for a locked project.
    # VBIDE has no supported Unprotect(password) method. Keep secrets inside Office.
    $command = $vbe.CommandBars.FindControl(1, 2578)
    if (-not $command -or -not $command.Enabled) {
        throw 'No se pudo abrir el dialogo VBA. En Excel pulsa Alt+F11, abre tu proyecto, introduce su contrasena y vuelve a ejecutar el comando.'
    }
    Write-Host 'Proyecto VBA protegido. Introduce su contrasena en la ventana de Excel.' -ForegroundColor Yellow
    Write-Host 'Si aparecen las propiedades del proyecto, cierra ese cuadro para continuar.' -ForegroundColor Yellow
    Write-Host 'Esperando acceso al codigo. Si cancelas en Excel, pulsa Ctrl+C aqui para detener la espera.' -ForegroundColor Yellow
    $command.Execute()
    # Execute may return while the password dialog is still open.
    Wait-VbaProjectAccess $Project -TimeoutSeconds $TimeoutSeconds
    $vbe.MainWindow.Visible = $wasVisible
    Write-Host 'Proyecto VBA desbloqueado. Continuando sincronizacion.' -ForegroundColor Green
}

function Get-VbaAccessibleProject($Excel, $Book) {
    try {
        $project = $Book.VBProject
        $null = $project.Protection
    } catch {
        throw 'Excel no permite acceder al proyecto VBA. Activa: Archivo > Opciones > Centro de confianza > Configuracion del Centro de confianza > Configuracion de macros > Confiar en el acceso al modelo de objetos de proyectos de VBA. Este permiso es distinto de la contrasena.'
    }
    if ([int]$project.Protection -ne 0) {
        $Book.Activate()
        Unlock-VbaProject $Excel $project
    }
    try { $null = $project.VBComponents.Count }
    catch { throw 'No se pudieron leer los modulos VBA. Comprueba que el proyecto este desbloqueado y que Excel permita acceder al modelo de objetos VBA.' }
    return $project
}

function Connect-VbaWorkbook([string] $Path, [switch] $Visible) {
    $excel = $null
    $book = $null
    $owned = $false
    $keepOpenOnError = $false
    try {
        try {
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
            foreach ($candidate in @($excel.Workbooks)) {
                if ($candidate.FullName -ieq $Path) { $book = $candidate; break }
            }
        } catch { $excel = $null }
        if (-not $book) {
            if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
            $excel = New-Object -ComObject Excel.Application
            $owned = $true
            # Display Excel BEFORE Open: Office requests opening/write passwords
            # natively when their optional arguments are omitted.
            $excel.Visible = $true
            $events = $excel.EnableEvents
            $security = $excel.AutomationSecurity
            $alerts = $excel.DisplayAlerts
            try {
                $excel.EnableEvents = $false
                $excel.DisplayAlerts = $true
                $excel.AutomationSecurity = 2 # Respect the user's Trust Center settings.
                $book = $excel.Workbooks.Open($Path, 0, $false)
            } finally {
                $excel.EnableEvents = $events
                $excel.AutomationSecurity = $security
                $excel.DisplayAlerts = $alerts
            }
        }
        if ($book.ReadOnly) { throw 'Workbook is read-only or open in another Excel instance. Close that copy and retry.' }
        $keepOpenOnError = $true
        $project = Get-VbaAccessibleProject $excel $book
        if ($Visible) { $excel.Visible = $true }
        elseif ($owned) { $excel.Visible = $false }
        return @{ Excel = $excel; Book = $book; Project = $project; Owned = $owned }
    } catch {
        if ($owned -and $excel -and -not $keepOpenOnError) {
            if ($book) { $book.Close($false) }
            $excel.Quit()
        }
        if ($excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
        throw
    }
}

function Get-VbaWorkbookStatus($Connection, [string] $Path) {
    try {
        foreach ($candidate in @($Connection.Excel.Workbooks)) {
            if ($candidate.FullName -ieq $Path) {
                if (-not $Connection.Excel.Ready) { return 'Busy' }
                return 'Open'
            }
        }
        return 'Closed'
    } catch {
        $failure = $_.Exception
        while ($failure) {
            if ($failure.HResult.ToString('X8') -in '80010108', '800706BA', '800706BE', '800401FD') {
                return 'Closed'
            }
            $failure = $failure.InnerException
        }
        # A save prompt, cell edit or another modal may temporarily reject COM.
        return 'Busy'
    }
}

function Disconnect-VbaWorkbook($Connection, [switch] $Close) {
    if (-not $Connection) { return }
    if ($Close -and $Connection.Owned) {
        try { $Connection.Book.Close($false) }
        catch { Write-Warning 'No se pudo cerrar el libro temporalmente; revisa Excel.' }
        try { $Connection.Excel.Quit() }
        catch { Write-Warning 'Excel sigue abierto. Puedes cerrarlo desde su ventana.' }
    }
    foreach ($key in @('Project', 'Book', 'Excel')) {
        try {
            if ($Connection[$key] -and [Runtime.InteropServices.Marshal]::IsComObject($Connection[$key])) {
                [Runtime.InteropServices.Marshal]::ReleaseComObject($Connection[$key]) | Out-Null
            }
        } catch { Write-Warning "No se pudo liberar referencia COM: $key" }
        finally { $Connection[$key] = $null }
    }
    [GC]::Collect()
    # Do not wait for COM finalizers: a pending Excel dialog can block them.
}

function Sync-VbaProject($Connection, [string] $SourceRoot, [string] $StatePath, [switch] $Flat,
    [ValidateSet('Auto', 'Push', 'Pull')] [string] $Direction = 'Auto', [switch] $Bidirectional) {
    $stage = New-VbaTemp
    $sourceStage = New-VbaTemp
    $events = $Connection.Excel.EnableEvents
    try {
        Export-VbaProject $Connection.Project $stage -Flat:$Flat
        $bookMap = Get-VbaSnapshot $stage
        $sourceMap = Get-VbaSnapshot $SourceRoot
        foreach ($name in $sourceMap.Keys) { Copy-VbaComponent $SourceRoot $sourceStage $sourceMap[$name] $null }
        $sourceMap = Get-VbaSnapshot $sourceStage
        $state = Read-VbaState $StatePath
        if ($Direction -eq 'Pull') {
            $names = @(@($sourceMap.Keys) + @($bookMap.Keys) | Sort-Object -Unique)
            foreach ($name in $names) {
                $entry = Copy-VbaExportToSourcePath $stage $SourceRoot $bookMap[$name] $sourceMap[$name]
                if ($entry) { $sourceMap[$name] = $entry } else { $sourceMap.Remove($name) }
            }
            Save-VbaState $StatePath $sourceMap $bookMap
            Write-Host "Extracted $($bookMap.Count) components -> $SourceRoot" -ForegroundColor Cyan
            return
        }
        if ($Direction -eq 'Push') {
            if (-not $state) { $state = @{ Source = @{}; Book = @{} } }
            $state.Book = $bookMap
            foreach ($name in $sourceMap.Keys) { $state.Source[$name] = @{ Path = $sourceMap[$name].Path; Hash = '' } }
        }
        if (-not $state) {
            if ($sourceMap.Count -gt 0) { throw "Source folder is not empty and has no sync state: $SourceRoot. Choose an empty project folder for the first extraction." }
            foreach ($name in $bookMap.Keys) { Copy-VbaComponent $stage $SourceRoot $bookMap[$name] $null }
            Save-VbaState $StatePath $bookMap $bookMap
            Write-Host "Extracted $($bookMap.Count) components -> $SourceRoot" -ForegroundColor Cyan
            return
        }
        $actions = @(Get-VbaChangePlan $sourceMap $bookMap $state -Bidirectional:$Bidirectional)
        $push = @($actions | Where-Object { $_.Direction -eq 'Push' })
        if ($push.Count -gt 0) {
            # Validate all document deletions before changing any component.
            foreach ($action in $push) {
                $component = @($Connection.Project.VBComponents) | Where-Object { $_.Name -ieq $action.Name } | Select-Object -First 1
                if ($component -and [int]$component.Type -eq 100 -and -not $sourceMap.ContainsKey($action.Name)) {
                    throw "Cannot delete document '$($action.Name)'. Restore its file and clear its code instead."
                }
            }
            $backupRoot = Join-Path (Split-Path -Parent $StatePath) 'backups'
            New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
            $backup = Join-Path $backupRoot ((Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + [IO.Path]::GetExtension($Connection.Book.FullName))
            $Connection.Excel.EnableEvents = $false
            $Connection.Book.SaveCopyAs($backup)
            try {
                foreach ($action in $push) {
                    $name = $action.Name
                    if ($Direction -eq 'Auto' -and $sourceMap.ContainsKey($name) -and $bookMap.ContainsKey($name) -and
                        $state.Source.ContainsKey($name) -and $sourceMap[$name].Path -like '*.frm' -and
                        $sourceMap[$name].BinaryHash -eq $state.Source[$name].BinaryHash) {
                        # The user edited code, not FRX. Retain the latest Excel controls/design.
                        $currentBinary = [IO.Path]::ChangeExtension((Join-Path $stage $bookMap[$name].Path), '.frx')
                        $stagedBinary = [IO.Path]::ChangeExtension((Join-Path $sourceStage $sourceMap[$name].Path), '.frx')
                        if (Test-Path -LiteralPath $currentBinary) { Copy-Item -LiteralPath $currentBinary -Destination $stagedBinary -Force }
                    }
                    Import-VbaComponent $Connection.Project $action.Name $sourceStage $sourceMap[$action.Name]
                }
                $Connection.Book.Save()
                if (-not $Connection.Book.Saved) { throw 'Excel did not save the workbook.' }
            } catch {
                $failure = $_
                foreach ($action in $push) {
                    try { Import-VbaComponent $Connection.Project $action.Name $stage $bookMap[$action.Name] }
                    catch { Write-Warning "Could not restore '$($action.Name)'. Backup: $backup" }
                }
                throw "Sync failed: $($failure.Exception.Message). Previous VBA restored in memory where possible. Backup: $backup"
            }
            foreach ($action in $push) { Write-Host "Excel <- $($action.Name)" -ForegroundColor Green }
        }
        foreach ($action in @($actions | Where-Object { $_.Direction -eq 'Pull' })) {
            $liveMap = Get-VbaSnapshot $SourceRoot
            $liveHash = if ($liveMap.ContainsKey($action.Name)) { $liveMap[$action.Name].Hash } else { '' }
            $expectedHash = if ($sourceMap.ContainsKey($action.Name)) { $sourceMap[$action.Name].Hash } else { '' }
            if ($liveHash -ne $expectedHash) { throw "Source changed during export: $($action.Name). Retrying without overwriting it." }
            $entry = Copy-VbaExportToSourcePath $stage $SourceRoot $bookMap[$action.Name] $sourceMap[$action.Name]
            if ($entry) { $sourceMap[$action.Name] = $entry }
            else { $sourceMap.Remove($action.Name) }
            Write-Host "src <- $($action.Name)" -ForegroundColor Cyan
        }
        if ($push.Count -gt 0) {
            $finalStage = New-VbaTemp
            try {
                Export-VbaProject $Connection.Project $finalStage -Flat:$Flat
                $bookMap = Get-VbaSnapshot $finalStage
            } finally { Remove-VbaTemp $finalStage }
        }
        Save-VbaState $StatePath $sourceMap $bookMap
    } finally {
        $Connection.Excel.EnableEvents = $events
        Remove-VbaTemp $stage
        Remove-VbaTemp $sourceStage
    }
}

function New-VbaProject([string] $Path) {
    $root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (Test-Path -LiteralPath $root) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            throw "El destino debe ser una carpeta: $root"
        }
        if (@(Get-ChildItem -LiteralPath $root -Force).Count -gt 0) {
            throw "La carpeta no esta vacia: $root. Para continuar un proyecto existente, usa dev ."
        }
    }
    $excel = $null
    $book = $null
    try {
        New-Item -ItemType Directory -Path (Join-Path $root 'excel') -Force | Out-Null
        $excel = New-Object -ComObject Excel.Application
        $book = $excel.Workbooks.Add()
        try {
            $module = $book.VBProject.VBComponents.Add(1)
            $module.Name = 'Main'
            $module.CodeModule.AddFromString("Option Explicit`r`n`r`nPublic Sub Hola()`r`n    MsgBox `"Hola desde VBA Dev`"`r`nEnd Sub")
        } catch { throw 'Cannot create VBA module. Enable Trust access to the VBA project object model in Excel Trust Center.' }
        $target = Join-Path $root ('excel\' + (Split-Path -Leaf $root) + '.xlsm')
        $book.SaveAs($target, 52)
        return $root
    } finally {
        if ($book) { $book.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
        if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    }
}

function Move-VbaWorkbookToExcel([string] $Workbook, [string] $Root, [string] $ConfigPath, [string] $ConfigText) {
    $rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $original = [IO.Path]::GetFullPath($Workbook)
    if ((Split-Path -Parent $original) -ine $rootPath) { return $original }
    $directory = Join-Path $rootPath 'excel'
    $target = Join-Path $directory (Split-Path -Leaf $original)
    if (Test-Path -LiteralPath $target) { throw "Ya existe un archivo en el destino: $target. No se reemplazo ninguno." }
    # Reject open workbooks before moving; never close the user's Excel instance.
    try {
        $handle = [IO.File]::Open($original, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $handle.Dispose()
    } catch { throw "No se puede mover el libro. Cierra Excel y vuelve a ejecutar dev .: $original. $($_.Exception.Message)" }
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    Move-Item -LiteralPath $original -Destination $target -ErrorAction Stop
    try {
        # Persist the new location before synchronization, including on unlock failures.
        Write-VbaText ($ConfigPath + '.tmp') $ConfigText
        Move-Item -LiteralPath ($ConfigPath + '.tmp') -Destination $ConfigPath -Force -ErrorAction Stop
    } catch {
        Move-Item -LiteralPath $target -Destination $original -ErrorAction Stop
        throw
    }
    Write-Host "Libro organizado: $target" -ForegroundColor Cyan
    return $target
}

function Start-VbaProject([string] $Path, [switch] $Once, [string] $SourceRoot, [switch] $Flat,
    [ValidateSet('Auto', 'Push', 'Pull')] [string] $Direction = 'Auto') {
    $pathInfo = Get-Item -LiteralPath $Path
    if ($pathInfo.PSIsContainer) {
        $root = $pathInfo.FullName
        $configPath = Join-Path $root '.vba\project.json'
        if (Test-Path -LiteralPath $configPath) {
            $config = [IO.File]::ReadAllText($configPath, $script:Utf8) | ConvertFrom-Json
            $workbook = [IO.Path]::GetFullPath((Join-Path $root $config.Workbook))
        } else {
            $books = @(Get-ChildItem -LiteralPath $root -File | Where-Object { $_.Extension -in '.xlsm', '.xlsb' -and $_.Name -notlike '~$*' })
            if (Test-Path -LiteralPath (Join-Path $root 'excel')) {
                $books += @(Get-ChildItem -LiteralPath (Join-Path $root 'excel') -File | Where-Object { $_.Extension -in '.xlsm', '.xlsb' -and $_.Name -notlike '~$*' })
            }
            if ($books.Count -ne 1) { throw 'Folder must contain exactly one .xlsm/.xlsb workbook, or pass the workbook path.' }
            $workbook = $books[0].FullName
        }
    } else {
        $workbook = $pathInfo.FullName
        $root = $pathInfo.DirectoryName
        $parent = Split-Path -Parent $root
        $parentConfig = Join-Path $parent '.vba\project.json'
        if (Test-Path -LiteralPath $parentConfig) {
            $saved = [IO.File]::ReadAllText($parentConfig, $script:Utf8) | ConvertFrom-Json
            if ([IO.Path]::GetFullPath((Join-Path $parent $saved.Workbook)) -ieq $workbook) { $root = $parent }
        }
    }
    if ([IO.Path]::GetExtension($workbook) -notin '.xlsm', '.xlsb') { throw 'Use an .xlsm or .xlsb workbook.' }
    $meta = Join-Path $root '.vba'
    $configPath = Join-Path $meta 'project.json'
    if (Test-Path -LiteralPath $configPath) {
        $config = [IO.File]::ReadAllText($configPath, $script:Utf8) | ConvertFrom-Json
        if ([IO.Path]::GetFullPath((Join-Path $root $config.Workbook)) -ine $workbook) { throw 'This folder is already linked to another workbook. Use one folder per project.' }
        $configuredSource = [IO.Path]::GetFullPath((Join-Path $root $config.Source))
        if ($SourceRoot -and $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourceRoot) -ine $configuredSource) {
            throw "This project already uses source folder: $configuredSource"
        }
        $SourceRoot = $configuredSource
        $Flat = [bool]$config.Flat
    }
    if (-not $SourceRoot) { $SourceRoot = Join-Path $root 'src' }
    $SourceRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourceRoot)
    if (-not $SourceRoot.StartsWith($root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Source folder must be inside the workbook project folder.'
    }
    New-Item -ItemType Directory -Path $SourceRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $meta -Force | Out-Null
    if (-not $Flat) {
        foreach ($folder in @('modules', 'classes', 'forms', 'documents')) { New-Item -ItemType Directory -Path (Join-Path $SourceRoot $folder) -Force | Out-Null }
    }
    $statePath = Join-Path $meta 'state.json'
    $layoutConfig = @{
        Workbook = 'excel\' + (Split-Path -Leaf $workbook)
        Source = $SourceRoot.Substring($root.TrimEnd('\').Length + 1)
        Flat = [bool]$Flat
    } | ConvertTo-Json
    $workbook = Move-VbaWorkbookToExcel $workbook $root $configPath $layoutConfig
    $connection = $null
    $started = $false
    try {
        $connection = Connect-VbaWorkbook $workbook -Visible:(-not $Once)
        Sync-VbaProject $connection $SourceRoot $statePath -Flat:$Flat -Direction $Direction
        Write-VbaText $configPath (@{
            Workbook = $workbook.Substring($root.TrimEnd('\').Length + 1)
            Source = $SourceRoot.Substring($root.TrimEnd('\').Length + 1)
            Flat = [bool]$Flat
        } | ConvertTo-Json)
        $started = $true
        if ($Once) { return }
        Write-Host "Editing: $SourceRoot`nSave files to update Excel. Ctrl+C stops watching; Excel stays open.`nCerrar el libro tambien detiene la sincronizacion." -ForegroundColor Yellow
        $lastSource = Get-VbaFingerprint (Read-VbaState $statePath).Source
        $lastBookWrite = (Get-Item -LiteralPath $workbook).LastWriteTimeUtc.Ticks
        $lastError = ''
        while ($true) {
            Start-Sleep -Milliseconds 500
            try {
                $bookStatus = Get-VbaWorkbookStatus $connection $workbook
                if ($bookStatus -eq 'Closed') {
                    Write-Host 'Libro cerrado. Sincronizacion detenida.' -ForegroundColor Yellow
                    break
                }
                if ($bookStatus -eq 'Busy') { continue }
                $currentSource = Get-VbaFingerprint (Get-VbaSnapshot $SourceRoot)
                $currentBookWrite = (Get-Item -LiteralPath $workbook).LastWriteTimeUtc.Ticks
                if ($currentSource -ne $lastSource -or $currentBookWrite -ne $lastBookWrite) {
                    # Wait for the editor's save/rename sequence to settle.
                    Start-Sleep -Milliseconds 250
                    $stable = Get-VbaFingerprint (Get-VbaSnapshot $SourceRoot)
                    if ($stable -ne $currentSource) { continue }
                    Sync-VbaProject $connection $SourceRoot $statePath -Flat:$Flat -Bidirectional
                    $lastSource = Get-VbaFingerprint (Read-VbaState $statePath).Source
                    $lastBookWrite = (Get-Item -LiteralPath $workbook).LastWriteTimeUtc.Ticks
                    $lastError = ''
                }
            } catch {
                $message = $_.Exception.Message
                if ($message -ne $lastError) { Write-Warning $message; $lastError = $message }
                # Transient COM errors or partial saves must not terminate the watcher.
            }
        }
    } finally { Disconnect-VbaWorkbook $connection -Close:($Once -or -not $started) }
}

Export-ModuleMember -Function Start-VbaProject, New-VbaProject
