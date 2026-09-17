[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('dev', 'new', 'ui')]
    [string] $Command = 'dev',
    [Parameter(Position = 1)]
    [string] $Path = '.',
    [switch] $Once
)

$ErrorActionPreference = 'Stop'
# Excel's active-object API is available in Windows PowerShell's .NET Framework.
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, $Command, $Path)
    if ($Once) { $arguments += '-Once' }
    & powershell.exe @arguments
    exit $LASTEXITCODE
}

try {
    Import-Module (Join-Path $PSScriptRoot 'tools\VbaDev.psm1') -Force
    if ($Command -eq 'new') {
        $Path = New-VbaProject -Path $Path
    }
    if ($Command -eq 'ui') {
        Start-CustomUiProject -Path $Path -Once:$Once
    } else {
        Start-VbaProject -Path $Path -Once:$Once
    }
} catch {
    Write-Host ("ERROR: " + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
