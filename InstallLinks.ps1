$WshShell = New-Object -comObject WScript.Shell
$where = $WshShell.SpecialFolders("Startup") 
$Shortcut = $WshShell.CreateShortcut($where + "\StopClock4Work.lnk")
$TargetPath = Join-Path $PSScriptRoot "StopClock4Work.ps1" 
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-windowstyle hidden " + $TargetPath 
$IconPath = Join-Path $PSScriptRoot "StopClock4Work.ico"
$Shortcut.IconLocation = $IconPath
$Shortcut.Save()

$where = $WshShell.SpecialFolders("Programs") 
$Shortcut = $WshShell.CreateShortcut($where + "\StopClock4Work.lnk")
$TargetPath = Join-Path $PSScriptRoot "StopClock4Work.ps1" 
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-windowstyle hidden " + $TargetPath 
$IconPath = Join-Path $PSScriptRoot "StopClock4Work.ico"
$Shortcut.IconLocation = $IconPath
$Shortcut.Save()