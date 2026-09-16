[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $Workbook,
    [Parameter(Mandatory = $true)] [string] $SourceRoot,
    [ValidateSet('Push', 'Pull')] [string] $Direction = 'Push',
    [switch] $Watch,
    [switch] $Visible
)
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath,
        '-Workbook', $Workbook, '-SourceRoot', $SourceRoot, '-Direction', $Direction)
    if ($Watch) { $arguments += '-Watch' }
    if ($Visible) { $arguments += '-Visible' }
    & powershell.exe @arguments
    exit $LASTEXITCODE
}
Import-Module (Join-Path $PSScriptRoot 'VbaDev.psm1') -Force
Start-VbaProject -Path $Workbook -SourceRoot $SourceRoot -Flat -Direction $Direction -Once:(-not $Watch)
