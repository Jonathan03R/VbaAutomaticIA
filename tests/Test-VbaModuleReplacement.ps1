# Integration regression: requires Windows PowerShell and desktop Excel.
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\tools\VbaDev.psm1') -Force
& (Get-Module VbaDev) {
    $excel = $null
    $book = $null
    $root = New-VbaTemp
    try {
        $text = @'
Attribute VB_Name = "ReplaceProbe"
Option Explicit
Private Declare PtrSafe Function NativeProbe Lib "kernel32" _
    Alias "GetTickCount" () As Long
Public Function ProbeValue() As Long
    ProbeValue = 42
End Function
'@
        Write-VbaText (Join-Path $root 'ReplaceProbe.bas') $text
        $excel = New-Object -ComObject Excel.Application
        $excel.EnableEvents = $false
        $excel.DisplayAlerts = $false
        $book = $excel.Workbooks.Add()
        $project = $book.VBProject
        $map = Get-VbaSnapshot $root
        for ($pass = 0; $pass -lt 3; $pass++) {
            Import-VbaComponent $project 'ReplaceProbe' $root $map['ReplaceProbe']
            $code = $project.VBComponents.Item('ReplaceProbe').CodeModule
            $actual = $code.Lines(1, $code.CountOfLines)
            if ($actual.Trim() -cne (Get-VbaDocumentCode $text).Trim()) {
                throw "FAIL: import $pass changed source text: $actual"
            }
            Complete-VbaProjectCompilation $excel $project
            $result = $excel.Run("'$($book.Name)'!ReplaceProbe.ProbeValue")
            if ($result -ne 42) { throw "FAIL: import $pass returned $result" }
        }
        Write-Host 'PASS: initial import and repeated replacements preserve code, compile and execute.'
    } finally {
        if ($book) { $book.Close($false) }
        if ($excel) {
            $excel.Quit()
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
        }
        Remove-VbaTemp $root
    }
}
