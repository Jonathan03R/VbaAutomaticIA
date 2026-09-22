param([switch] $Integration)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\tools\VbaDev.psm1') -Force
$module = Get-Module VbaDev

& $module {
    function Assert($Condition, [string] $Message) {
        if (-not $Condition) { throw "FAIL: $Message" }
        Write-Host "PASS: $Message"
    }
    function New-CustomUiWorkbook([string] $Path, [string] $Xml) {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew)
        try {
            $archive = New-Object IO.Compression.ZipArchive($stream, 1, $false)
            try {
                $entry = $archive.CreateEntry('customUI/customUI.xml')
                $writer = New-Object IO.StreamWriter($entry.Open(), $script:Utf8)
                try { $writer.Write($Xml) } finally { $writer.Dispose() }
                $relations = $archive.CreateEntry('customUI/_rels/customUI.xml.rels')
                $relationWriter = New-Object IO.StreamWriter($relations.Open(), $script:Utf8)
                try { $relationWriter.Write('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>') } finally { $relationWriter.Dispose() }
                $image = $archive.CreateEntry('customUI/images/prueba/iniciar.png')
                $imageStream = $image.Open()
                try { $imageStream.Write([byte[]]@(1, 2, 3), 0, 3) } finally { $imageStream.Dispose() }
            } finally { $archive.Dispose() }
        } finally { $stream.Dispose() }
    }
    $header = @'
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_PredeclaredId = True
Option Explicit
Private Sub Workbook_Open()
End Sub
Public Function Value() As Long
    Value = 1
End Function
'@
    $code = Get-VbaDocumentCode $header
    Assert ($code.Contains('End Sub') -and $code.Contains('End Function')) 'document procedure endings survive import'
    Assert ($code -notmatch 'MultiUse|VERSION|Attribute') 'document export metadata removed'
    $base = @{ A = @{ Path = 'modules\A.bas'; Hash = 'old'; TextHash = 'old'; BinaryHash = '' } }
    $changed = @{ A = @{ Path = 'modules\A.bas'; Hash = 'new'; TextHash = 'new'; BinaryHash = '' } }
    $state = @{ Source = $base; Book = $base }
    $plan = @(Get-VbaChangePlan $changed $base $state)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Push') 'offline source edit is pushed on restart'
    $plan = @(Get-VbaChangePlan $base $changed $state)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Push') 'source restores Excel-only code edits'
    $plan = @(Get-VbaChangePlan $base $changed $state -Bidirectional)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Pull') 'live Excel-only edit is extracted'
    $plan = @(Get-VbaChangePlan $changed $base $state -Bidirectional)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Push') 'live source-only edit is imported'
    $plan = @(Get-VbaChangePlan $changed @{ A = @{ Hash = 'different'; TextHash = 'different' } } $state -Bidirectional)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Push') 'source wins simultaneous live edits'
    $plan = @(Get-VbaChangePlan $changed @{ A = @{ Hash = 'different'; TextHash = 'different'; Path = 'modules\A.bas' } } $state)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Push') 'source wins concurrent edits'
    $plan = @(Get-VbaChangePlan $base @{} $state)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Pull') 'Excel deletion removes unchanged source'
    $plan = @(Get-VbaChangePlan $changed @{} $state)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Push') 'pending source edit wins Excel deletion'
    $plan = @(Get-VbaChangePlan @{} $base @{ Source = @{}; Book = @{} })
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Pull') 'new Excel component is extracted'
    $plan = @(Get-VbaChangePlan @{} $base $state)
    Assert ($plan.Count -eq 1 -and $plan[0].Direction -eq 'Push') 'source deletion tracked'
    $busyError = [Runtime.InteropServices.COMException]::new('Excel is busy', -2147417848)
    Assert ((Get-VbaWorkbookErrorStatus $busyError) -eq 'Busy') 'transient Excel COM state does not stop watcher'
    Assert (-not (Should-StopForClosedWorkbook 1)) 'one missing workbook check does not stop watcher'
    Assert (Should-StopForClosedWorkbook 3) 'repeated missing workbook checks stop watcher'
    $editing = @{ Excel = [pscustomobject]@{ Ready = $false; Workbooks = @() } }
    Assert ((Get-VbaWorkbookStatus $editing 'C:\book.xlsm') -eq 'Busy') 'editing with empty workbook enumeration must not signal closure'
    $unavailable = @{ Excel = [pscustomobject]@{ Ready = $true; Workbooks = $null } }
    Assert ((Get-VbaWorkbookStatus $unavailable 'C:\book.xlsm') -eq 'Busy') 'unavailable workbook collection must not signal closure'
    $closed = @{ Excel = [pscustomobject]@{ Ready = $true; Workbooks = @() } }
    Assert ((Get-VbaWorkbookStatus $closed 'C:\book.xlsm') -eq 'Closed') 'responsive Excel with zero workbooks confirms closure'
    $indexedBooks = [pscustomobject]@{ Count = 1; Path = 'C:\book.xlsm' }
    $indexedBooks | Add-Member ScriptMethod Item { param($index) if ($index -ne 1) { throw 'Invalid index' }; return [pscustomobject]@{ FullName = $this.Path } }
    $openBook = @{ Excel = [pscustomobject]@{ Ready = $true; Workbooks = $indexedBooks } }
    Assert ((Get-VbaWorkbookStatus $openBook 'C:\book.xlsm') -eq 'Open') 'indexed collection finds open workbook'
    $indexedBooks.Path = ''
    Assert ((Get-VbaWorkbookStatus $openBook 'C:\book.xlsm') -eq 'Busy') 'unavailable workbook path is not closure'
    $temp = New-VbaTemp
    try {
        Write-VbaText (Join-Path $temp 'NombreLegible.bas') "Attribute VB_Name = `"NombreInterno`"`r`nOption Explicit"
        $named = Get-VbaSnapshot $temp
        Assert ($named.ContainsKey('NombreInterno') -and $named.NombreInterno.Path -eq 'NombreLegible.bas') 'friendly source filename preserves VBA component identity'
        Remove-Item -LiteralPath (Join-Path $temp 'NombreLegible.bas')
        Write-VbaText (Join-Path $temp 'Form1.frm') "VERSION 5.00`r`nAttribute VB_Name = `"Form1`"`r`nOleObjectBlob = `"Form1.frx`":0000"
        [IO.File]::WriteAllBytes((Join-Path $temp 'Form1.frx'), [byte[]]@(1, 2, 3))
        $before = Get-VbaSnapshot $temp
        [IO.File]::WriteAllBytes((Join-Path $temp 'Form1.frx'), [byte[]]@(1, 2, 4))
        $after = Get-VbaSnapshot $temp
        Assert ($before.Form1.Hash -ne $after.Form1.Hash) 'FRX-only edit detected'
        $statePath = Join-Path $temp 'state.json'
        Save-VbaState $statePath @{} @{}
        $empty = Read-VbaState $statePath
        Assert ($empty.Source.Count -eq 0) 'empty state can be resumed'

        $uiBook = Join-Path $temp 'Ribbon.xlsm'
        $uiSource = Join-Path $temp 'ui-src'
        $uiXml = '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon /></customUI>'
        New-CustomUiWorkbook $uiBook $uiXml
        Export-CustomUi $uiBook $uiSource
        $uiFile = Join-Path $uiSource 'custom-ui\customUI.xml'
        Assert ([IO.File]::ReadAllText($uiFile, $script:Utf8) -eq $uiXml) 'Custom UI is extracted from XLSM'
        $uiImage = Join-Path $uiSource 'custom-ui\images\prueba\iniciar.png'
        Assert ((Test-Path -LiteralPath $uiImage) -and ((Get-Item $uiImage).Length -eq 3)) 'Custom UI nested image is extracted from XLSM'
        $uiFingerprintBeforeImageDelete = Get-CustomUiFingerprint $uiSource
        Remove-Item -LiteralPath $uiImage
        Assert ((Get-CustomUiFingerprint $uiSource) -ne $uiFingerprintBeforeImageDelete) 'Custom UI watcher detects nested image deletion'
        Import-CustomUi $uiBook $uiSource
        $afterDelete = Join-Path $temp 'ui-after-delete'
        Export-CustomUi $uiBook $afterDelete
        Assert (-not (Test-Path -LiteralPath (Join-Path $afterDelete 'custom-ui\images\prueba\iniciar.png'))) 'deleting source Custom UI image deletes it from XLSM'
        $changedUiXml = '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs /></ribbon></customUI>'
        Write-VbaText $uiFile $changedUiXml
        Import-CustomUi $uiBook $uiSource
        $verifySource = Join-Path $temp 'ui-verify'
        Export-CustomUi $uiBook $verifySource
        Assert ([IO.File]::ReadAllText((Join-Path $verifySource 'custom-ui\customUI.xml'), $script:Utf8) -eq $changedUiXml) 'Custom UI source change is saved into XLSM'
        Write-VbaText $uiFile '<customUI>'
        try { Import-CustomUi $uiBook $uiSource; throw 'Expected invalid XML error.' }
        catch { Assert ($_.Exception.Message -match 'XML.*válido') 'invalid Custom UI XML reports Spanish error' }
        $uiLockRoot = Join-Path $temp 'ui-lock'
        New-Item -ItemType Directory -Path $uiLockRoot | Out-Null
        $lock = Acquire-VbaSessionLock $uiLockRoot 'dev'
        try {
            try { Acquire-VbaSessionLock $uiLockRoot 'ui'; throw 'Expected active lock error.' }
            catch { Assert ($_.Exception.Message -match 'vba dev') 'active dev lock blocks Custom UI with Spanish error' }
        } finally { Release-VbaSessionLock $lock }
        $uiLock = Acquire-VbaSessionLock $uiLockRoot 'ui'
        Release-VbaSessionLock $uiLock
        Assert (-not (Test-Path -LiteralPath (Join-Path $uiLockRoot 'session.lock'))) 'session lock is released'

        $uiProject = Join-Path $temp 'UiProject'
        New-Item -ItemType Directory -Path $uiProject | Out-Null
        $uiProjectBook = Join-Path $uiProject 'Ribbon.xlsm'
        New-CustomUiWorkbook $uiProjectBook $uiXml
        Start-CustomUiProject $uiProject -Once
        $uiProjectBook = Join-Path $uiProject 'excel\Ribbon.xlsm'
        $uiProjectFile = Join-Path $uiProject 'src\custom-ui\customUI.xml'
        Assert (Test-Path -LiteralPath $uiProjectFile) 'vba ui exports Custom UI into project source'
        Write-VbaText $uiProjectFile $changedUiXml
        Start-CustomUiProject $uiProject -Once
        $uiProjectVerify = Join-Path $temp 'ui-project-verify'
        Export-CustomUi $uiProjectBook $uiProjectVerify
        Assert ([IO.File]::ReadAllText((Join-Path $uiProjectVerify 'custom-ui\customUI.xml'), $script:Utf8) -eq $changedUiXml) 'vba ui once imports edited Custom UI'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot '..\vba.ps1') ui $uiProject -Once | Out-Null
        Assert ($LASTEXITCODE -eq 0) 'CLI accepts vba ui command'
        $devLock = Acquire-VbaSessionLock (Join-Path $uiProject '.vba') 'dev'
        try { Start-CustomUiProject $uiProject -Once; throw 'Expected vba ui lock error.' }
        catch { Assert ($_.Exception.Message -match 'vba dev') 'vba ui refuses project while vba dev is active' }
        finally { Release-VbaSessionLock $devLock }
        $uiModeLock = Acquire-VbaSessionLock (Join-Path $uiProject '.vba') 'ui'
        try { Start-VbaProject $uiProject -Once; throw 'Expected vba dev lock error.' }
        catch { Assert ($_.Exception.Message -match 'vba ui') 'vba dev refuses project while vba ui is active' }
        finally { Release-VbaSessionLock $uiModeLock }
        Assert (Should-CloseVbaOnExit $true $true) 'stopping active vba dev closes managed Excel'
    } finally { Remove-VbaTemp $temp }
}

if ($Integration) {
    & $module {
        function Assert($Condition, [string] $Message) {
            if (-not $Condition) { throw "FAIL: $Message" }
            Write-Host "PASS: $Message"
        }
        $temp = New-VbaTemp
        $connection = $null
        try {
            $root = New-VbaProject (Join-Path $temp 'Example')
            Start-VbaProject $root -Once
            $bookPath = Join-Path $root 'excel\Example.xlsm'
            $source = Join-Path $root 'src'
            $statePath = Join-Path $root '.vba\state.json'
            $main = Join-Path $source 'modules\Main.bas'
            Assert ((Test-Path -LiteralPath $bookPath) -and (Test-Path -LiteralPath $main)) 'new project creates Excel and source tree'
            $text = [IO.File]::ReadAllText($main, $script:Utf8).Replace('Hola desde VBA Dev', ('Edici' + [char]0xF3 + 'n externa'))
            Write-VbaText $main $text
            Start-VbaProject $root -Once
            $connection = Connect-VbaWorkbook $bookPath
            $component = $connection.Project.VBComponents.Item('Main')
            $code = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
            Assert ($code.Contains('Edici' + [char]0xF3 + 'n externa')) 'saved source edit imported with accents on restart'
            $alerts = $connection.Excel.DisplayAlerts
            $events = $connection.Excel.EnableEvents
            $documents = @(Get-ChildItem -LiteralPath (Join-Path $source 'documents') -Filter '*.cls')
            $doc = $documents[0]
            $docText = [IO.File]::ReadAllText($doc.FullName, $script:Utf8) + "`r`nPublic Sub Example()`r`n    Debug.Print 42`r`nEnd Sub`r`n"
            Write-VbaText $doc.FullName $docText
            Sync-VbaProject $connection $source $statePath
            $docComponent = $connection.Project.VBComponents.Item($doc.BaseName)
            $docCode = $docComponent.CodeModule.Lines(1, $docComponent.CodeModule.CountOfLines)
            Assert ($docCode.Contains('End Sub') -and $docCode -notmatch 'MultiUse') 'document module imports valid procedure text'
            Assert ($connection.Excel.DisplayAlerts -eq $alerts -and $connection.Excel.EnableEvents -eq $events) 'Excel application settings restored'
            $component = $connection.Project.VBComponents.Item('Main')
            $component.CodeModule.AddFromString("' edited in Excel")
            Sync-VbaProject $connection $source $statePath
            $component = $connection.Project.VBComponents.Item('Main')
            Assert (-not $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines).Contains('edited in Excel')) 'Excel code restored from source'
            Assert (-not [IO.File]::ReadAllText($main, $script:Utf8).Contains("' edited in Excel")) 'source preserved after Excel edit'
            Write-VbaText $main ([IO.File]::ReadAllText($main, $script:Utf8) + "`r`n' local conflict")
            $component.CodeModule.AddFromString("' Excel conflict")
            Sync-VbaProject $connection $source $statePath
            $component = $connection.Project.VBComponents.Item('Main')
            $code = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
            Assert ($code.Contains('local conflict') -and -not $code.Contains('Excel conflict')) 'source wins real concurrent changes'
            Write-VbaText $main ([IO.File]::ReadAllText($main, $script:Utf8) + "`r`n' pending while deleted")
            $connection.Project.VBComponents.Remove($component)
            Sync-VbaProject $connection $source $statePath
            $component = $connection.Project.VBComponents.Item('Main')
            Assert ($component.CodeModule.Lines(1, $component.CodeModule.CountOfLines).Contains('pending while deleted')) 'pending edit recreates deleted Excel module'
            Sync-VbaProject $connection $source $statePath -Direction Pull
            Remove-Item -LiteralPath $main
            Sync-VbaProject $connection $source $statePath
            Assert (@($connection.Project.VBComponents | Where-Object Name -eq 'Main').Count -eq 0) 'deleted source module removed from Excel'
            $backups = @(Get-ChildItem -LiteralPath (Join-Path $root '.vba\backups') -File)
            Assert ($backups.Count -ge 3) 'workbook backups created before modifications'
            # Import a real exported UserForm with its binary resource.
            $fixtureRoot = Join-Path $PSScriptRoot '..\src'
            $formSource = Join-Path $source 'forms\UserForm1.frm'
            $formText = [IO.File]::ReadAllText((Join-Path $fixtureRoot 'UserForm1.frm'), $script:Ansi)
            Write-VbaText $formSource $formText
            Copy-Item -LiteralPath (Join-Path $fixtureRoot 'UserForm1.frx') -Destination (Join-Path $source 'forms\UserForm1.frx')
            Sync-VbaProject $connection $source $statePath
            Write-VbaText $formSource ($formText.Replace('"hola"', '"external"'))
            Sync-VbaProject $connection $source $statePath
            $form = $connection.Project.VBComponents.Item('UserForm1')
            Assert ($form.Designer.Controls.Count -gt 0) 'UserForm controls survive reimport with FRX'
            Assert ($form.CodeModule.Lines(1, $form.CodeModule.CountOfLines).Contains('external')) 'UserForm edited code reaches Excel'
            $connection.Project.VBComponents.Remove($form)
            $connection.Book.Save()
            Sync-VbaProject $connection $source $statePath
            Assert (-not (Test-Path -LiteralPath $formSource) -and -not (Test-Path -LiteralPath ([IO.Path]::ChangeExtension($formSource, '.frx')))) 'Excel form deletion removes source and binary resource'
            Sync-VbaProject $connection $source $statePath
            Assert (@($connection.Project.VBComponents | Where-Object Name -eq 'UserForm1').Count -eq 0) 'deleted component stays deleted on next sync'
            $newModule = $connection.Project.VBComponents.Add(1)
            $newModule.Name = 'CreatedInExcel'
            $newModule.CodeModule.AddFromString("' new Excel module")
            Sync-VbaProject $connection $source $statePath
            Assert (Test-Path -LiteralPath (Join-Path $source 'modules\CreatedInExcel.bas')) 'new Excel module exported to source'
            $connection.Book.Saved = $true
        } finally {
            Disconnect-VbaWorkbook $connection -Close
            Remove-VbaTemp $temp
        }
    }
}
