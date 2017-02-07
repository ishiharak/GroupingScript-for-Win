Set-StrictMode -version Latest
Add-Type -assemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = 'CSV Files|*.csv|All Files|*.*'
$dialog.InitialDirectory = '.'
 
if ($dialog.ShowDialog() -ne "OK") {
    exit 1
}

Import-Csv $dialog.FileName -Header "id","group","alias" | ForEach-Object {
    $groupPath = Join-Path $PSScriptRoot  $_.group
    if (!(Test-Path $groupPath)) {
        ni $groupPath -itemType Directory
    }
    $shortcutName = $_.alias + ".lnk"
    $shortcutPath = Join-Path $groupPath $shortcutName

    if (Test-Path $shortcutPath) {
        rm $shortcutPath -Force
    }

    $FolderPath = Join-Path $PSScriptRoot $_.id

    $wsh = New-Object -comObject WScript.Shell
    $link = $wsh.CreateShortcut($shortcutPath)
    $link.TargetPath = $FolderPath
    $link.Save()
}