# Windows PowerShell 5.1 + desktop Excel. Uses only disposable local workbooks.
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\tools\VbaDev.psm1') -Force
$module = Get-Module VbaDev
$root = Join-Path ([IO.Path]::GetTempPath()) ('vba-integrity-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path (Join-Path $root 'excel')
$path = Join-Path $root 'excel\Integrity.xlsm'
function Assert($Condition, [string] $Message) {
    if (-not $Condition) { throw "FAIL: $Message (fixture retained at $root)" }
    Write-Host "PASS: $Message"
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$xl = New-Object -ComObject Excel.Application
try {
    $xl.EnableEvents = $false
    $xl.AutomationSecurity = 3
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $wb.Worksheets.Item(1).Range('A1').Value2 = 'Original data'
    $wb.Worksheets.Item(1).Range('A2').Value2 = 123.5
    $wb.Worksheets.Item(1).Range('A3').Formula = '=A2*2'
    $project = $wb.VBProject
    $project.VBComponents.Item($wb.CodeName).CodeModule.AddFromString("Option Explicit`r`nPublic Function IntegrityProbe() As Long`r`nIntegrityProbe = 42`r`nEnd Function")
    $wb.SaveAs($path, 52)
    $file = [IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
    $zip = New-Object IO.Compression.ZipArchive($file, 0, $false)
    $memory = New-Object IO.MemoryStream
    try {
        $input = $zip.GetEntry('xl/vbaProject.bin').Open()
        try { $input.CopyTo($memory); $oldVba = $memory.ToArray() } finally { $input.Dispose() }
    } finally { $zip.Dispose(); $memory.Dispose() }
    $sheet = $wb.Worksheets.Add()
    $sheet.Name = 'New sheet'
    $sheet.Range('B2').Value2 = 'Keep this too'
    $newCodeName = $sheet.CodeName
    $wb.Save()
    $wb.Close($false)
} finally { $xl.Quit() }

# Reproduce the actual defect: sheet XML exists, but its VBA project is stale.
& $module {
    param($path, $oldVba, $newCodeName)
    $zip = [IO.Compression.ZipFile]::Open($path, 'Update')
    try {
        foreach ($entry in @($zip.Entries)) {
            if ($entry.FullName -notmatch '^xl/worksheets/[^/]+\.xml$') { continue }
            [xml]$xml = Read-ZipEntryText $zip $entry.FullName
            $pr = $xml.SelectSingleNode('/*/*[local-name()="sheetPr"]')
            if ($pr -and $pr.GetAttribute('codeName') -eq $newCodeName) {
                $pr.RemoveAttribute('codeName')
                Write-ZipEntryText $zip $entry.FullName $xml.OuterXml
            }
        }
        $zip.GetEntry('xl/vbaProject.bin').Delete()
        $output = $zip.CreateEntry('xl/vbaProject.bin').Open()
        try { $output.Write($oldVba, 0, $oldVba.Length) } finally { $output.Dispose() }
    } finally { $zip.Dispose() }
} $path $oldVba $newCodeName
$missing = @(& $module { param($p) Get-VbaMissingDocumentModules $p } $path)
Assert ($missing.Count -eq 1 -and $missing[0] -like '@xl/worksheets/*') 'detects sheet without stored module or codeName'
$originalHash = (Get-FileHash $path -Algorithm SHA256).Hash
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot '..\vba.ps1') dev $root -Once
Assert ($LASTEXITCODE -eq 0) 'vba dev repairs and completes initial sync'
Assert (@(& $module { param($p) Get-VbaMissingDocumentModules $p } $path).Count -eq 0) 'repaired file has every document module on disk'
$backups = @(Get-ChildItem (Join-Path $root 'excel\backup') -Filter '*.xlsm')
Assert ($backups.Count -eq 1 -and (Get-FileHash $backups[0].FullName).Hash -eq $originalHash) 'backup is byte-identical to original damaged file'

$xl = New-Object -ComObject Excel.Application
try {
    $xl.AutomationSecurity = 3
    $xl.EnableEvents = $false
    $wb = $xl.Workbooks.Open($path, 0, $false)
    $newCodeName = $wb.Worksheets.Item('New sheet').CodeName
    Assert (Test-Path (Join-Path $root ('src\documents\' + $newCodeName + '.cls'))) 'new document is synchronized to source'
    $original = @($wb.Worksheets | Where-Object { $_.Name -ne 'New sheet' })[0]
    Assert ($original.Range('A1').Value2 -eq 'Original data' -and $original.Range('A2').Value2 -eq 123.5) 'original text and number survive repair'
    Assert ($original.Range('A3').Formula -eq '=A2*2') 'formula survives repair'
    Assert ($wb.Worksheets.Item('New sheet').Range('B2').Value2 -eq 'Keep this too') 'added sheet data survives repair'
    $wb.Close($false)
} finally { $xl.Quit() }
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot '..\vba.ps1') dev $root -Once
Assert ($LASTEXITCODE -eq 0) 'healthy workbook still synchronizes'
Assert (@(Get-ChildItem (Join-Path $root 'excel\backup') -Filter '*.xlsm').Count -eq 1) 'healthy reopening does not create another repair or backup'
Write-Host "Disposable regression project: $root"
