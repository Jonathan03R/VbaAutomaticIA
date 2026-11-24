$ErrorActionPreference = 'Stop'
$temp = Join-Path ([IO.Path]::GetTempPath()) ('vba-password-ui-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
$handleFile = Join-Path $temp 'window.txt'
$resultFile = Join-Path $temp 'result.txt'
$fixture = Join-Path $PSScriptRoot 'VbaPasswordUiFixture.ps1'
$process = $null
try {
    $process = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $fixture + '"'), ('"' + $handleFile + '"'), ('"' + $resultFile + '"')) -WindowStyle Hidden -PassThru
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    while (-not (Test-Path -LiteralPath $handleFile) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    if (-not (Test-Path -LiteralPath $handleFile)) { throw 'No se abrio el cuadro de prueba.' }
    $windowHandle = [long]([IO.File]::ReadAllText($handleFile))
    & (Join-Path $PSScriptRoot '..\tools\Enter-VbaPassword.ps1') $windowHandle 'test-secret' | Out-Null
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while (-not (Test-Path -LiteralPath $resultFile) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    if (-not (Test-Path -LiteralPath $resultFile)) { throw 'No se confirmo el cuadro de prueba.' }
    if ([IO.File]::ReadAllText($resultFile) -ne 'test-secret') { throw 'La clave no llego al cuadro protegido.' }
    Write-Host 'PASS: la clave se introduce en un cuadro protegido y se confirma.'
} finally {
    if ($process -and -not $process.HasExited) { $process.Kill() }
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}
