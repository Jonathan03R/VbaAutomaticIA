[CmdletBinding()]
param([string] $Workbook = '.')
& (Join-Path (Split-Path -Parent $PSScriptRoot) 'vba.ps1') dev $Workbook
