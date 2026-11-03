# Read the compound-file directory only; never execute or rewrite VBA binaries.
function Get-VbaStoredStreamNames([byte[]] $Data) {
    if ($Data.Length -lt 512 -or [BitConverter]::ToString($Data, 0, 8) -ne 'D0-CF-11-E0-A1-B1-1A-E1') {
        throw 'El proyecto VBA no contiene un archivo compuesto valido.'
    }
    $shift = [BitConverter]::ToUInt16($Data, 30)
    if ($shift -notin @(9, 12) -or [BitConverter]::ToUInt16($Data, 28) -ne 65534) {
        throw 'Formato de sectores VBA no compatible.'
    }
    $sectorSize = 1 -shl $shift
    $sectorCount = [int][Math]::Floor($Data.Length / $sectorSize) - 1
    $fatIds = New-Object 'Collections.Generic.List[uint32]'
    $fatCount = [BitConverter]::ToUInt32($Data, 44)
    if ($fatCount -gt $sectorCount) { throw 'Cantidad de sectores FAT invalida.' }
    for ($i = 0; $i -lt 109; $i++) {
        $id = [BitConverter]::ToUInt32($Data, 76 + 4 * $i)
        if ($id -lt 4294967292) { $fatIds.Add($id) }
    }
    $difat = [BitConverter]::ToUInt32($Data, 68)
    $difatCount = [BitConverter]::ToUInt32($Data, 72)
    if ($difatCount -gt $sectorCount) { throw 'Cantidad de sectores DIFAT invalida.' }
    $visited = New-Object 'Collections.Generic.HashSet[uint32]'
    for ($i = 0; $i -lt $difatCount; $i++) {
        if ($difat -ge $sectorCount -or -not $visited.Add($difat)) { throw 'Cadena DIFAT invalida.' }
        $offset = ([long]$difat + 1) * $sectorSize
        for ($j = 0; $j -lt $sectorSize / 4 - 1; $j++) {
            $id = [BitConverter]::ToUInt32($Data, [int]($offset + 4 * $j))
            if ($id -lt 4294967292) { $fatIds.Add($id) }
        }
        $difat = [BitConverter]::ToUInt32($Data, [int]($offset + $sectorSize - 4))
    }
    if ($fatIds.Count -ne $fatCount) { throw 'Directorio FAT incompleto.' }
    $fat = New-Object 'Collections.Generic.List[uint32]'
    $visited.Clear()
    foreach ($id in $fatIds) {
        if ($id -ge $sectorCount -or -not $visited.Add($id)) { throw 'Sector FAT invalido.' }
        $offset = ([long]$id + 1) * $sectorSize
        for ($j = 0; $j -lt $sectorSize / 4; $j++) {
            $fat.Add([BitConverter]::ToUInt32($Data, [int]($offset + 4 * $j)))
        }
    }
    $directory = New-Object IO.MemoryStream
    try {
        $visited.Clear()
        $id = [BitConverter]::ToUInt32($Data, 48)
        while ($id -ne 4294967294) {
            if ($id -ge $sectorCount -or $id -ge $fat.Count -or -not $visited.Add($id)) {
                throw 'Cadena del directorio VBA invalida.'
            }
            $directory.Write($Data, [int](([long]$id + 1) * $sectorSize), $sectorSize)
            $id = $fat[[int]$id]
        }
        $entries = $directory.ToArray()
    } finally { $directory.Dispose() }
    if ($entries.Length -lt 128 -or $entries[66] -ne 5) { throw 'Directorio raiz VBA ausente.' }
    $pending = New-Object Collections.Stack
    $pending.Push(@{ Id = [BitConverter]::ToUInt32($entries, 76); Parent = '' })
    $visited.Clear()
    while ($pending.Count) {
        $item = $pending.Pop()
        $id = [uint32]$item.Id
        if ($id -eq 4294967295) { continue }
        if ($id -ge $entries.Length / 128 -or -not $visited.Add($id)) { throw 'Arbol del directorio VBA invalido.' }
        $offset = [int]$id * 128
        $nameLength = [BitConverter]::ToUInt16($entries, $offset + 64)
        if ($nameLength -lt 2 -or $nameLength -gt 64 -or $nameLength % 2) { throw 'Nombre de modulo VBA invalido.' }
        $name = [Text.Encoding]::Unicode.GetString($entries, $offset, $nameLength - 2)
        $path = $item.Parent + $name
        $type = $entries[$offset + 66]
        foreach ($position in @(68, 72)) {
            $pending.Push(@{ Id = [BitConverter]::ToUInt32($entries, $offset + $position); Parent = $item.Parent })
        }
        if ($type -eq 1) {
            $pending.Push(@{ Id = [BitConverter]::ToUInt32($entries, $offset + 76); Parent = $path + '/' })
        } elseif ($type -eq 2) { $path }
        else { throw 'Entrada VBA inesperada.' }
    }
}

function Get-VbaMissingDocumentModules([string] $Path) {
    # XLSB document metadata is binary; retain its existing opening behaviour.
    if ([IO.Path]::GetExtension($Path) -ine '.xlsm') { return }
    Import-CustomUiAssemblies
    $file = [IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try {
        # Encrypted Office files are compound files, not ZIP. Let Excel request
        # their password normally; do not mistake encryption for missing modules.
        $signature = New-Object byte[] 8
        $null = $file.Read($signature, 0, 8)
        $file.Position = 0
        if ([BitConverter]::ToString($signature) -eq 'D0-CF-11-E0-A1-B1-1A-E1') { return }
        $archive = New-Object IO.Compression.ZipArchive($file, 0, $true)
        try {
            $vba = $archive.GetEntry('xl/vbaProject.bin')
            if (-not $vba) { return }
            $stream = $vba.Open()
            $bytes = New-Object IO.MemoryStream
            try { $stream.CopyTo($bytes); $stored = @(Get-VbaStoredStreamNames $bytes.ToArray()) }
            finally { $stream.Dispose(); $bytes.Dispose() }
            $names = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            foreach ($part in $archive.Entries) {
                if ($part.FullName -ne 'xl/workbook.xml' -and $part.FullName -notmatch '^xl/(worksheets|chartsheets)/[^/]+\.xml$') { continue }
                [xml]$xml = Read-ZipEntryText $archive $part.FullName
                $nodes = @($xml.SelectNodes('/*/*[local-name()="sheetPr" or local-name()="workbookPr"][@codeName]'))
                if (-not $nodes.Count -or -not $nodes[0].GetAttribute('codeName')) {
                    # A sheet injected through ZIP/XML may not have a codeName
                    # at all. Excel must create and persist its document module.
                    '@' + $part.FullName
                }
                foreach ($node in $nodes) {
                    $name = $node.GetAttribute('codeName')
                    if ($name -and $stored -inotcontains ('VBA/' + $name) -and $names.Add($name)) { $name }
                }
            }
        } finally { $archive.Dispose() }
    } finally { $file.Dispose() }
}

function Repair-VbaDocumentModules($Excel, [string] $Path, [string[]] $Missing) {
    $book = $null
    $backup = $null
    $saving = $false
    $closed = $false
    $events = $Excel.EnableEvents
    $security = $Excel.AutomationSecurity
    try {
        $Excel.EnableEvents = $false
        $Excel.AutomationSecurity = 3 # Disable macros for this repair opening only.
        $book = $Excel.Workbooks.Open($Path, 0, $false)
        if ($book.ReadOnly) { throw 'El libro esta abierto en otra instancia o es de solo lectura. Cierralo y vuelve a ejecutar vba dev .' }
        $project = Get-VbaAccessibleProject $Excel $book
        $documentNames = @($book.CodeName) + @($book.Sheets | ForEach-Object { $_.CodeName })
        foreach ($name in $documentNames) {
            if ([string]::IsNullOrWhiteSpace($name)) { throw 'Excel no pudo asignar nombre interno a una hoja.' }
            $component = $project.VBComponents.Item($name)
            if ([int]$component.Type -ne 100) { throw "Excel no pudo reconstruir el modulo de documento '$name'." }
        }
        $backupRoot = Join-Path (Split-Path -Parent $Path) 'backup'
        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
        $backup = Join-Path $backupRoot ([IO.Path]::GetFileNameWithoutExtension($Path) + '-antes-reparacion-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsm')
        [IO.File]::Copy($Path, $backup, $false)
        Write-Host "Respaldo de reparacion: $backup" -ForegroundColor Cyan
        # Excel synthesizes missing document modules on load, yet may report
        # Saved=True. Force persistence even when no source code was edited.
        $saving = $true
        $book.Saved = $false
        $book.Save()
        if (-not $book.Saved) { throw 'Excel no pudo guardar la reparacion.' }
        $book.Close($false)
        $closed = $true
        if (@(Get-VbaMissingDocumentModules $Path).Count) { throw 'El proyecto VBA sigue incompleto despues de guardar.' }
        Write-Host ("Proyecto VBA reparado; documentos registrados: " + ($documentNames -join ', ')) -ForegroundColor Green
    } catch {
        $failure = $_
        if ($book -and -not $closed) {
            try { $book.Close($false); $closed = $true } catch { }
        }
        if ($saving -and $backup -and $closed) {
            try { [IO.File]::Copy($backup, $Path, $true) }
            catch { throw "No se pudo restaurar automaticamente el libro. Respaldo intacto: $backup. $($failure.Exception.Message)" }
        }
        throw "No se pudo reparar el proyecto VBA. Respaldo: $backup. $($failure.Exception.Message)"
    } finally {
        $Excel.EnableEvents = $events
        $Excel.AutomationSecurity = $security
    }
}
