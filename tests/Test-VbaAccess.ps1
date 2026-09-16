param([switch] $Integration)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\tools\VbaDev.psm1') -Force
$module = Get-Module VbaDev
& $module {
    function Assert($Condition, [string] $Message) {
        if (-not $Condition) { throw "FAIL: $Message" }
        Write-Host "PASS: $Message"
    }
    function New-FakeExcel($Project, [bool] $AcceptPassword) {
        $command = [pscustomobject]@{ Enabled = $true; Calls = 0; Project = $Project; Accept = $AcceptPassword }
        $command | Add-Member ScriptMethod Execute {
            $this.Calls++
            if ($this.Accept) { $this.Project.Protection = 0 }
        }
        $bars = [pscustomobject]@{ Command = $command }
        $bars | Add-Member ScriptMethod FindControl {
            param($Type, $Id)
            if ($Type -ne 1 -or $Id -ne 2578) { throw 'Incorrect command' }
            return $this.Command
        }
        return [pscustomobject]@{
            Visible = $false
            VBE = [pscustomobject]@{
                MainWindow = [pscustomobject]@{ Visible = $false }
                ActiveVBProject = $null
                CommandBars = $bars
            }
        }
    }
    $project = [pscustomobject]@{ Protection = 0; VBComponents = @('Main') }
    $excel = New-FakeExcel $project $true
    Unlock-VbaProject $excel $project
    Assert ($excel.VBE.CommandBars.Command.Calls -eq 0) 'unprotected VBA never prompts'
    $project.Protection = 1
    Unlock-VbaProject $excel $project
    Assert ($excel.VBE.CommandBars.Command.Calls -eq 1) 'locked VBA invokes the native password dialog'
    Assert ($project.Protection -eq 0 -and -not $excel.VBE.MainWindow.Visible) 'accepted password resumes and restores editor visibility'
    $project.Protection = 1
    $excel = New-FakeExcel $project $false
    $cancelled = $false
    try { Unlock-VbaProject $excel $project -TimeoutSeconds 1 }
    catch { $cancelled = $_.Exception.Message -like '*Se agoto la espera*' }
    Assert $cancelled 'a still-locked project times out without assuming the password was rejected'
    Assert $excel.VBE.MainWindow.Visible 'password window remains available after timeout'
    $delayed = [pscustomobject]@{ Reads = 0; ComponentReads = 0; Started = $false }
    $delayed | Add-Member ScriptProperty Protection {
        if (-not $this.Started) { return 1 }
        $this.Reads++
        if ($this.Reads -le 3) { return 1 }
        if ($this.Reads -le 5) { throw (New-Object Runtime.InteropServices.COMException('Office is busy', -2147418111)) }
        return 0
    }
    $delayed | Add-Member ScriptProperty VBComponents {
        $this.ComponentReads++
        if ($this.ComponentReads -lt 3) { throw 'Properties dialog still open' }
        return [pscustomobject]@{ Count = 1 }
    }
    $excel = New-FakeExcel $delayed $false
    $excel.VBE.CommandBars.Command | Add-Member -Force ScriptMethod Execute {
        $this.Calls++
        $this.Project.Started = $true
    }
    Unlock-VbaProject $excel $delayed -TimeoutSeconds 5
    Assert ($delayed.Reads -ge 8 -and $delayed.ComponentReads -eq 3) 'async dialog waits through locked and busy states until modules are accessible'
    Assert ($excel.VBE.CommandBars.Command.Calls -eq 1) 'waiting never opens duplicate password dialogs'
    $denied = [pscustomobject]@{}
    $denied | Add-Member ScriptProperty VBProject { throw 'Programmatic access denied' }
    $trustError = $false
    try { Get-VbaAccessibleProject $excel $denied }
    catch { $trustError = $_.Exception.Message -like '*Este permiso es distinto*' }
    Assert $trustError 'trust-center denial is distinguished from a VBA password'
}
if ($Integration) {
    $excel = $null
    $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $book = $excel.Workbooks.Add()
        $command = $excel.VBE.CommandBars.FindControl(1, 2578)
        if (-not $command -or -not $command.Enabled) { throw 'Native VBA properties command unavailable.' }
        Write-Host "PASS: native Excel command 2578 available: $($command.Caption)"
    } finally {
        if ($book) { $book.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
        if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    }
}
