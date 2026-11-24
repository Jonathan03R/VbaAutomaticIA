[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('dev', 'new', 'ui')]
    [string] $Command = 'dev',
    [Parameter(Position = 1)]
    [string] $Path = '.',
    [switch] $Once,
    [string] $Clave,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $ExtraArgs
)

$ErrorActionPreference = 'Stop'
if ($Path -match '^clave:(.*)$') {
    if ($PSBoundParameters.ContainsKey('Clave')) { throw 'Indica la clave una sola vez.' }
    $Clave = $Matches[1]
    if ($ExtraArgs.Count -ne 1) { throw 'Uso: vba dev clave:TU_CLAVE RUTA' }
    $Path = $ExtraArgs[0]
} elseif ($ExtraArgs.Count) {
    throw 'Argumentos adicionales no reconocidos.'
}
if ($Command -eq 'ui' -and -not [string]::IsNullOrEmpty($Clave)) { throw 'La clave solo se admite con vba dev o vba new.' }
if ($PSBoundParameters.ContainsKey('Clave') -and [string]::IsNullOrEmpty($Clave)) { throw 'La clave no puede estar vacia.' }
# Excel's active-object API is available in Windows PowerShell's .NET Framework.
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, $Command, $Path)
    if ($Once) { $arguments += '-Once' }
    if (-not [string]::IsNullOrEmpty($Clave)) { $arguments += @('-Clave', $Clave) }
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
        Start-VbaProject -Path $Path -Once:$Once -Clave $Clave
    }
} catch {
    Write-Host ("ERROR: " + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
