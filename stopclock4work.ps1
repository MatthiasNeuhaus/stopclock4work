# Die ersten beiden Befehle holen sich die .NET-Erweiterungen (sog. Assemblies) für die grafische Gestaltung in den RAM.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


$StartTime = Get-Date #- (New-TimeSpan -Minute 1)  # margin for checkin/boot Time
#$Pause_StartTime = 0;
#$Pause_StartTime = New-TimeSpan -Seconds 0;
#$Stop = 0;
$Pause = 0;
$NormalTime = New-TimeSpan -Seconds 0;
$TotalPauseTime = New-TimeSpan -Seconds 0;
$PauseTime = New-TimeSpan -Seconds 0;


# Die nächste Zeile erstellt aus der Formsbibliothek das Fensterobjekt.
$MainWindow = New-Object System.Windows.Forms.Form

# Hintergrundfarbe für das Fenster festlegen
$MainWindow.Backcolor = “white“

# Icon in die Titelleiste setzen
# $MainWindow.Icon="C:\Powershell\XXX.ico"  #kann selbst definiert werden

# Hintergrundbild mit Formatierung Zentral = 2
#$MainWindow.BackgroundImageLayout = 2
#$MainWindow.BackgroundImage = [System.Drawing.Image]::FromFile('C:\Powershell\xxxx.jpg')  #kann selbst definiert werden

# Position des Fensters festlegen
$MainWindow.StartPosition = "CenterScreen"

# Button size
$Height = 32
$Widht = 147

# Fenstergröße festlegen
$MainWindow.Size = New-Object System.Drawing.Size((2 * $Widht), (5 * $Height))

# Titelleiste festlegen
$MainWindow.Text = "Attendance"

# Countup
$Countup = New-Object System.Windows.Forms.Label
$Countup.Location = New-Object System.Drawing.Size(0, 0)
#$Countup.Size = New-Object System.Drawing.Size($Widht,$Height)
$Countup.Text = "Counter"
$MainWindow.Controls.Add($Countup)

# Countup Pause
$CountupPause = New-Object System.Windows.Forms.Label
$CountupPause.Location = New-Object System.Drawing.Size($Widht, 0)
$CountupPause.AutoSize = $True
$CountupPause.Text = "00:00:00"
$CountupPause.ForeColor = "red"
$CountupPause.TextAlign = "MiddleCenter"
$MainWindow.Controls.Add($CountupPause)


function GetLogonStatus () {
    try {
        $user = $null
        $user = Get-WmiObject -Class win32_computersystem -ComputerName localhost | Select-Object -ExpandProperty username -ErrorAction Stop
    }
    catch { return -1 }
    try {
        if ((Get-Process logonui -ComputerName localhost -ErrorAction Stop) -and ($user)) {
            return 1
        }
    }
    catch { if ($user) { return 0 } }
}


$stopclock = {
    if ($script:Pause -eq 0) {
        $Time = Get-Date;
        $script:NormalTime = New-TimeSpan –Start $script:StartTime –End $Time;
        $script:NormalTime = $script:NormalTime - $script:TotalPauseTime;
        $Countup.Text = $script:NormalTime.ToString("hh\:mm\:ss");
    }
    else {
        $Time = Get-Date;
        $script:PauseTime = New-TimeSpan –Start $script:PauseStartTime –End $Time;
        $script:PauseTime = $script:PauseTime + $script:TotalPauseTime;
        $CountupPause.Text = $script:PauseTime.ToString("hh\:mm\:ss");
        
        if ($script:Pause -eq 2) {
            $retval = GetLogonStatus
            if ( $retval -eq 0) {
                $script:Pause = 0
                $script:TotalPauseTime = $script:PauseTime
            }
        }
    }
}


$timer = New-Object 'System.Windows.Forms.Timer'
$timer.Enabled = $True 
$timer.Interval = 1000
$timer.add_Tick($stopclock)



# Buttons
$StopClose = New-Object System.Windows.Forms.Button
$StopClose.Text = "Stop and Close" 
$StopClose.Width = $Widht
$StopClose.Height = $Height
$StopClose.Location = New-Object System.Drawing.Size(0, $Height)
$StopClose.Add_Click( { $timer.Enabled = $False })
$MainWindow.Controls.Add($StopClose)

$StopHybernate = New-Object System.Windows.Forms.Button
$StopHybernate.Text = "Stop and Hybernate" 
$StopHybernate.Width = $Widht
$StopHybernate.Height = $Height
$StopHybernate.Location = New-Object System.Drawing.Size(0, (2 * $Height))
$StopHybernate.Add_Click( { $script:Pause = 2 })
$MainWindow.Controls.Add($StopHybernate)

$StopShutdown = New-Object System.Windows.Forms.Button
$StopShutdown.Text = "Stop and Shutdown" 
$StopShutdown.Width = $Widht
$StopShutdown.Height = $Height
$StopShutdown.Location = New-Object System.Drawing.Size(0, (3 * $Height))
$StopShutdown.Add_Click( { [void] $NewWindow.ShowDialog() })
$MainWindow.Controls.Add($StopShutdown)

$PauseButton = New-Object System.Windows.Forms.Button
$PauseButton.Text = "Pause" 
$PauseButton.Width = $Widht
$PauseButton.Height = $Height
$PauseButton.Location = New-Object System.Drawing.Size($Widht, $Height)
$PauseButton.Add_Click( { 
        $script:Pause = 1
        $script:PauseStartTime = Get-Date })
$MainWindow.Controls.Add($PauseButton)

$Resume = New-Object System.Windows.Forms.Button
$Resume.Text = "Resume" 
$Resume.Width = $Widht
$Resume.Height = $Height
$Resume.Location = New-Object System.Drawing.Size($Widht, (2 * $Height))
$Resume.Add_Click( { 
        $script:Pause = 0
        $script:TotalPauseTime = $script:PauseTime })
$MainWindow.Controls.Add($Resume)

$PauseLock = New-Object System.Windows.Forms.Button
$PauseLock.Text = "Pause and Lock" 
$PauseLock.Width = $Widht
$PauseLock.Height = $Height
$PauseLock.Location = New-Object System.Drawing.Size($Widht, (3 * $Height))
$PauseLock.Add_Click( { 
        
        $script:PauseStartTime = Get-Date
        rundll32.exe user32.dll, LockWorkStation
        Start-Sleep -Seconds 10;
        $script:Pause = 2
     })
$MainWindow.Controls.Add($PauseLock)


# Die letzte Zeile sorgt dafür, dass unser Fensterobjekt auf dem Bildschirm angezeigt wird.
[void] $MainWindow.ShowDialog()