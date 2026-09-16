param([switch] $Integration)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\tools\VbaDev.psm1') -Force
$module = Get-Module VbaDev

& $module {
    function Assert($Condition, [string] $Message) {
        if (-not $Condition) { throw "FAIL: $Message" }
        Write-Host "PASS: $Message"
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
    $temp = New-VbaTemp
    try {
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
