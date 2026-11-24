param([string] $HandleFile, [string] $ResultFile)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$excel = New-Object Windows.Forms.Form
$excel.Text = 'Excel de prueba'
$excel.StartPosition = 'Manual'
$excel.Location = New-Object Drawing.Point(100, 100)
$dialog = New-Object Windows.Forms.Form
$dialog.Text = 'VBAProject Password'
$dialog.StartPosition = 'Manual'
$dialog.Location = New-Object Drawing.Point(120, 120)
$passwordBox = New-Object Windows.Forms.TextBox
$passwordBox.UseSystemPasswordChar = $true
$passwordBox.Location = New-Object Drawing.Point(10, 10)
$ok = New-Object Windows.Forms.Button
$ok.Text = 'OK'
$ok.Location = New-Object Drawing.Point(10, 40)
$dialog.AcceptButton = $ok
$dialog.Controls.Add($passwordBox)
$dialog.Controls.Add($ok)
$ok.Add_Click({
    $script:CapturedPassword = $passwordBox.Text
    $dialog.Close()
    $script:PropertiesForm = New-Object Windows.Forms.Form
    $script:PropertiesForm.Text = 'Project Properties'
    $script:PropertiesForm.StartPosition = 'Manual'
    $script:PropertiesForm.Location = New-Object Drawing.Point(120, 120)
    $script:PropertiesForm.Add_FormClosed({
        [IO.File]::WriteAllText($ResultFile, $script:CapturedPassword)
        $excel.Close()
    })
    $script:PropertiesForm.Show($excel)
})
$excel.Show()
$dialog.Show($excel)
[IO.File]::WriteAllText($HandleFile, [string]([long]$excel.Handle))
[Windows.Forms.Application]::Run($excel)
