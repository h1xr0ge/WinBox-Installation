$packageName = 'winbox'
$appName = 'WinBox'
$url32 = 'https://mt.lv/winbox'
$url64 = 'https://mt.lv/winbox64'
$localPrograms = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Programs'
$dir = Join-Path -Path $localPrograms -ChildPath $appName
$is64 = [environment]::Is64BitOperatingSystem -and [environment]::Is64BitProcess
$processName32 = 'winbox'
$processName64 = 'winbox64'
$processName = If ($is64) { $processName64 } Else { $processName32 }
$exe = $processName + '.exe'
$fullPath = Join-Path -Path $dir -ChildPath $exe
$shortcutPath = Join-Path -Path ([environment]::GetFolderPath('Programs')) -ChildPath ($appName + '.lnk')

Write-Host 'WinBox installation start'

if(-not (Test-Path -Path $dir))
{
    New-Item -Path $dir -ItemType Directory
}
 
(New-Object System.Net.WebClient).DownloadFile($(if ($is64) {$url64} else {$url32}), $fullPath)


$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $fullPath
$Shortcut.Save()

Write-Host 'WinBox installation is complete.'