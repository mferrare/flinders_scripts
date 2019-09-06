Add-Type -AssemblyName System.Windows.Forms

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$null = $FileBrowser.ShowDialog()

Get-FileHash $FileBrowser.FileName
