# Die ersten beiden Befehle holen sich die .NET-Erweiterungen (sog. Assemblies) für die grafische Gestaltung in den RAM.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


enum PauseTypes {
    running
    stopped
    stoppedByLock
}

$StartTime = Get-Date #- (New-TimeSpan -Minute 1)  # margin for checkin/boot Time
[PauseTypes]$Pause = [PauseTypes]::running;
$WorkTime = New-TimeSpan -Seconds 0;
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
$MainWindow.Text = "WT: 00:00:00"

# Countup
$Countup = New-Object System.Windows.Forms.Label
$Countup.Location = New-Object System.Drawing.Size(0, 0)
#$Countup.Size = New-Object System.Drawing.Size($Widht,$Height)
$Countup.Text = "00:00:00"
$MainWindow.Controls.Add($Countup)

# Countup Pause
$CountupPause = New-Object System.Windows.Forms.Label
$CountupPause.Location = New-Object System.Drawing.Size($Widht, 0)
$CountupPause.AutoSize = $True
$CountupPause.Text = "00:00:00"
$CountupPause.ForeColor = "red"
$CountupPause.TextAlign = "MiddleCenter"
$MainWindow.Controls.Add($CountupPause)

function WriteToCsv () {
    $PauseTimes = Import-Csv -Path .\PauseTimes.csv -Delimiter ";"

    ForEach ($Index in 0 .. ($PauseTimes.AboveHrs.Count - 1))
    {
        $AboveHr = [Timespan]::Parse($PauseTimes.AboveHrs[$Index])
        if ($script:WorkTime -ge $AboveHr) 
        { 
            $PauseTimeCalc = [Timespan]::Parse($PauseTimes.PauseTime[$Index])
        }
    }    
    $WorkEndTime        = $script:StartTime + $script:WorkTime;
    $WorkEndTimeCalc    = $script:StartTime + $script:WorkTime + $PauseTimeCalc;
    $writeStart = [PSCustomObject]@{
        DayOfWeek       = $script:StartTime.DayOfWeek.ToString() 
        Date            = $script:StartTime.Date.ToString('dd/MM/yyyy')
        WorkStartTime   = $script:StartTime.TimeOfDay.ToString('hh\:mm\:ss')
        WorkTime        = $script:WorkTime.ToString('hh\:mm\:ss')
        PauseTime       = $script:TotalPauseTime.ToString('hh\:mm\:ss')
        WorkEndTime     = $WorkEndTime.TimeOfDay.ToString('hh\:mm\:ss')
        WorkEndTimeCalc = $WorkEndTimeCalc.TimeOfDay.ToString('hh\:mm\:ss')
        PauseTimeCalc   = $PauseTimeCalc.ToString('hh\:mm\:ss')
    }
    $writeStart | Export-Csv -UseCulture -Path .\timesheet.csv -Append -NoTypeInformation -Force
}

function GetLogonStatus () {

    #get current user name
    try {
        $user = $null
        $user = Get-WmiObject -Class win32_computersystem -ComputerName localhost | Select-Object -ExpandProperty username -ErrorAction Stop
    }
    catch { return -1 }
    #check if workstation is locked
    try {
        if ((Get-Process logonui -ComputerName localhost -ErrorAction Stop) -and ($user)) {
            return 1
        }
    }
    catch { if ($user) { return 0 } } #user is logged on 
}


$stopclock = {

    $logonStatus = GetLogonStatus

    switch ($script:Pause) {
        
        ([PauseTypes]::running) {
            if ( $logonStatus -eq 1)
            {
                $script:Pause = [PauseTypes]::stoppedByLock
                $script:PauseStartTime = Get-Date
            }
            else {
                $Time = Get-Date;
                $script:WorkTime    = New-TimeSpan –Start $script:StartTime –End $Time;
                $script:WorkTime    = $script:WorkTime - $script:TotalPauseTime;
                $Countup.Text       = $script:WorkTime.ToString("hh\:mm\:ss");
                $MainWindow.Text    = "WT: " + $Countup.Text;
            }
            break  
        }

        ([PauseTypes]::stoppedByLock) {

            if ( $logonStatus -eq 0) {
                $Time = Get-Date;
                $script:PauseTime = New-TimeSpan –Start $script:PauseStartTime –End $Time;
                $script:PauseTime = $script:PauseTime + $script:TotalPauseTime;
                $CountupPause.Text = $script:PauseTime.ToString("hh\:mm\:ss");

                $script:Pause = [PauseTypes]::running
                $script:TotalPauseTime = $script:PauseTime
            }
            break
        }

        ([PauseTypes]::stopped) {
            if ( $logonStatus -eq 1)
            {
                $script:Pause = [PauseTypes]::stoppedByLock
            }
            else {            
                $Time = Get-Date;
                $script:PauseTime = New-TimeSpan –Start $script:PauseStartTime –End $Time;
                $script:PauseTime = $script:PauseTime + $script:TotalPauseTime;
                $CountupPause.Text = $script:PauseTime.ToString("hh\:mm\:ss");
            }
            break
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
$StopClose.Add_Click( { 
        $timer.Enabled = $False 
        WriteToCsv
        [void] $MainWindow.Close()
    })
$MainWindow.Controls.Add($StopClose)

$StopHybernate = New-Object System.Windows.Forms.Button
$StopHybernate.Text = "Stop and Hybernate" 
$StopHybernate.Width = $Widht
$StopHybernate.Height = $Height
$StopHybernate.Location = New-Object System.Drawing.Size(0, (2 * $Height))
$StopHybernate.Add_Click( { 
        $timer.Enabled = $False
        WriteToCsv
        rundll32.exe powrprof.dll, SetSuspendState 0, 1, 0
    })
$MainWindow.Controls.Add($StopHybernate)

$StopShutdown = New-Object System.Windows.Forms.Button
$StopShutdown.Text = "Stop and Shutdown" 
$StopShutdown.Width = $Widht
$StopShutdown.Height = $Height
$StopShutdown.Location = New-Object System.Drawing.Size(0, (3 * $Height))
$StopShutdown.Add_Click( { 
        $timer.Enabled = $False
        WriteToCsv
        shutdown /s /hybrid
        [void] $MainWindow.Close()
    })
$MainWindow.Controls.Add($StopShutdown)

$PauseButton = New-Object System.Windows.Forms.Button
$PauseButton.Text = "Pause" 
$PauseButton.Width = $Widht
$PauseButton.Height = $Height
$PauseButton.Location = New-Object System.Drawing.Size($Widht, $Height)
$PauseButton.Add_Click( { 
        $script:Pause = [PauseTypes]::stopped
        $script:PauseStartTime = Get-Date 
    })
$MainWindow.Controls.Add($PauseButton)

$Resume = New-Object System.Windows.Forms.Button
$Resume.Text = "Resume" 
$Resume.Width = $Widht
$Resume.Height = $Height
$Resume.Location = New-Object System.Drawing.Size($Widht, (2 * $Height))
$Resume.Add_Click( { 
        $script:Pause = [PauseTypes]::running
        $script:TotalPauseTime = $script:PauseTime 
    })
$MainWindow.Controls.Add($Resume)

$PauseLock = New-Object System.Windows.Forms.Button
$PauseLock.Text = "Pause and Lock" 
$PauseLock.Width = $Widht
$PauseLock.Height = $Height
$PauseLock.Location = New-Object System.Drawing.Size($Widht, (3 * $Height))
$PauseLock.Add_Click( { 
        $script:PauseStartTime = Get-Date
        rundll32.exe user32.dll, LockWorkStation
        Start-Sleep -Seconds 5; # wait till user is really loged of
        $script:Pause = [PauseTypes]::stoppedByLock
    })
$MainWindow.Controls.Add($PauseLock)


# Die letzte Zeile sorgt dafür, dass unser Fensterobjekt auf dem Bildschirm angezeigt wird.
[void] $MainWindow.ShowDialog()