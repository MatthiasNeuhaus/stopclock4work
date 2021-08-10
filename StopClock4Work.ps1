﻿# Die ersten beiden Befehle holen sich die .NET-Erweiterungen (sog. Assemblies) für die grafische Gestaltung in den RAM.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


enum BreakTypes {
    running
    stopped
    stoppedByLock
}

# Create (gitignored) Data folder if it not exists
$DataFolder = Join-Path $PSScriptRoot "Data\"
if (!(Test-Path $DataFolder ))
{
    mkdir $DataFolder
}
# Copy (generic) Data Source Elements into Data folder, so they can be adapted at will without being under version controll
$DataSourceFolder = Join-Path $PSScriptRoot "DataSource\"
Get-ChildItem $DataSourceFolder |
ForEach-Object {
    $DataTargetElement = Join-Path $DataFolder $_.Name
    if (!(Test-Path $DataTargetElement -PathType leaf))
    {
        Copy-Item $_.FullName $DataFolder
    }
}

# Variables for calculation
$StartTime = Get-Date
[BreakTypes]$BreakType = [BreakTypes]::running;
$WorkTime = New-TimeSpan -Seconds 0;
$TotalBreakTime = New-TimeSpan -Seconds 0;
$BreakTime = New-TimeSpan -Seconds 0;
$BreakStartTime = Get-Date
$SecondsToWriteTmp = 300 # -> means every 5min
$TempFileTimer = $SecondsToWriteTmp

# Write tmp file for persistence during restart
$TmpFileName = $StartTime.Date.ToString('dd/MM/yyyy') + ".csv"
$TmpFile = Join-Path $DataFolder $TmpFileName

if (Test-Path $TmpFile -PathType leaf) # Temp file does exist - read it!
{
    $ReadStart = Import-Csv -UseCulture -Path $TmpFile
    
    $StartTime              = [DateTime]::Parse($ReadStart.StartTime)
    $WorkTime               = [Timespan]::Parse($ReadStart.WorkTime)
    $TotalBreakTime         = [Timespan]::Parse($ReadStart.BreakTime)
    $BreakStartTime         = [DateTime]::Parse($ReadStart.BreakStartTime)
    $BreakType              = $ReadStart.BreakType
}

# Read target work times for countdown
$TargetWorkTimesPath = Join-Path $DataFolder "TargetWorkTimesV0.csv"
$TargetWorkTimes = Import-Csv -Path $TargetWorkTimesPath -Delimiter ";"
$Today = $StartTime.DayOfWeek
$TargetWorktime = [Timespan]::Parse($TargetWorkTimes.$Today)
$WorkTimeBalance = $WorkTime - $TargetWorktime

# Die nächste Zeile erstellt aus der Formsbibliothek das Fensterobjekt.
$MainWindow = New-Object System.Windows.Forms.Form

# Remove close buttons
$MainWindow.ControlBox = $false

# Hintergrundfarbe für das Fenster festlegen
$MainWindow.Backcolor = “white“

# Icon in die Titelleiste setzen
$IconPath = Join-Path $PSScriptRoot "StopClock4Work.ico"
$MainWindow.Icon = $IconPath

# Hintergrundbild mit Formatierung Zentral = 2
#$MainWindow.BackgroundImageLayout = 2
#$MainWindow.BackgroundImage = [System.Drawing.Image]::FromFile('C:\Powershell\xxxx.jpg')  #kann selbst definiert werden

# Position des Fensters festlegen
$MainWindow.StartPosition = "CenterScreen"

# Button size
$Height = 32
$Widht = 147

# Fenstergröße festlegen
$MainWindow.Size = New-Object System.Drawing.Size(((2 * $Widht)+15), ((5 * $Height)+5))

# Countdown in title
$MainWindow.Text = "WTB: $(if($WorkTimeBalance -lt [TimeSpan]::Zero ){"-"})$($WorkTimeBalance.ToString('hh\:mm\:ss'))"

# Countup
$Countup = New-Object System.Windows.Forms.Label
$Countup.Location = New-Object System.Drawing.Size(0, 0)
#$Countup.Size = New-Object System.Drawing.Size($Widht,$Height)
$Countup.Text = "WT: " + $script:WorkTime.ToString("hh\:mm\:ss")
$MainWindow.Controls.Add($Countup)

# Countup Break
$CountupBreak = New-Object System.Windows.Forms.Label
$CountupBreak.Location = New-Object System.Drawing.Size($Widht, 0)
$CountupBreak.AutoSize = $True
$CountupBreak.Text = "BT: " + $script:TotalBreakTime.ToString("hh\:mm\:ss")
$CountupBreak.ForeColor = "red"
$CountupBreak.TextAlign = "MiddleCenter"
$MainWindow.Controls.Add($CountupBreak)

function WriteTmpFile ()
{
    $writeTmp = [PSCustomObject]@{
        StartTime       = $script:StartTime
        WorkTime        = $script:WorkTime
        BreakTime       = $script:TotalBreakTime
        BreakStartTime  = $script:BreakStartTime
        BreakType       = $script:BreakType
    }
    $writeTmp | Export-Csv -UseCulture -Path $script:TmpFile
}

function WriteToCsv () {
	
    $BreakTimesPath = Join-Path $DataFolder "BreakTimesV0.csv"
    $BreakTimes = Import-Csv -Path $BreakTimesPath -Delimiter ";"

    ForEach ($Index in 0 .. ($BreakTimes.AboveHrs.Count - 1))
    {
        $AboveHr = [Timespan]::Parse($BreakTimes.AboveHrs[$Index])
        if ($script:WorkTime -ge $AboveHr) 
        { 
            $BreakTimeCalc = [Timespan]::Parse($BreakTimes.BreakTime[$Index])
        }
    }    
    $WorkEndTime        = $script:StartTime + $script:WorkTime + $script:TotalBreakTime;
    $WorkEndTimeCalc    = $script:StartTime + $script:WorkTime + $BreakTimeCalc;

    $TimeSheetPath = Join-Path $DataFolder "WorkTimesV0.csv"
    $writeOutput = [PSCustomObject]@{
        DayOfWeek       = $script:StartTime.DayOfWeek.ToString() 
        Date            = $script:StartTime.Date.ToString('dd/MM/yyyy')
        WorkStartTime   = $script:StartTime.TimeOfDay.ToString('hh\:mm\:ss')
        WorkTime        = $script:WorkTime.ToString('hh\:mm\:ss')
        BreakTime       = $script:TotalBreakTime.ToString('hh\:mm\:ss')
        WorkEndTime     = $WorkEndTime.TimeOfDay.ToString('hh\:mm\:ss')
        WorkEndTimeCalc = $WorkEndTimeCalc.TimeOfDay.ToString('hh\:mm\:ss')
        BreakTimeCalc   = $BreakTimeCalc.ToString('hh\:mm\:ss')
        WorkTimeBalance = "$(if($script:WorkTimeBalance -lt [TimeSpan]::Zero ){"-"})$($script:WorkTimeBalance.ToString('hh\:mm\:ss'))"
    }
    $writeOutput | Export-Csv -UseCulture -Path $TimeSheetPath -Append -NoTypeInformation -Force

    if (Test-Path $TmpFile -PathType leaf)
    {
        Remove-Item $TmpFile
    }
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

    if ($script:TempFileTimer -ge $SecondsToWriteTmp)
    {
        WriteTmpFile
        $script:TempFileTimer = 0
    }
    else 
    {
        $script:TempFileTimer++
    }

    switch ($script:BreakType) {
        
        ([BreakTypes]::running) {
            if ( $logonStatus -eq 1)
            {
                $script:BreakType = [BreakTypes]::stoppedByLock
                $script:BreakStartTime = Get-Date
                WriteTmpFile
            }
            else {
                $Time = Get-Date;
                $script:WorkTime    = New-TimeSpan –Start $script:StartTime –End $Time
                $script:WorkTime    = $script:WorkTime - $script:TotalBreakTime
                $Countup.Text       = "WT: " + $script:WorkTime.ToString("hh\:mm\:ss")
                $script:WorkTimeBalance    = $script:WorkTime - $script:TargetWorktime
                $MainWindow.Text    = "WTB: $(if($script:WorkTimeBalance -lt [TimeSpan]::Zero ){"-"})$($script:WorkTimeBalance.ToString('hh\:mm\:ss'))"
            }
            break  
        }

        ([BreakTypes]::stoppedByLock) {

            if ( $logonStatus -eq 0) {
                $Time = Get-Date;
                $script:BreakTime = New-TimeSpan –Start $script:BreakStartTime –End $Time
                $script:BreakTime = $script:BreakTime + $script:TotalBreakTime
                $CountupBreak.Text = "BT: " + $script:BreakTime.ToString("hh\:mm\:ss")

                $script:BreakType = [BreakTypes]::running
                $script:TotalBreakTime = $script:BreakTime
                WriteTmpFile
            }
            break
        }

        ([BreakTypes]::stopped) {
            if ( $logonStatus -eq 1)
            {
                $script:BreakType = [BreakTypes]::stoppedByLock
                WriteTmpFile
            }
            else {            
                $Time = Get-Date;
                $script:BreakTime = New-TimeSpan –Start $script:BreakStartTime –End $Time
                $script:BreakTime = $script:BreakTime + $script:TotalBreakTime
                $CountupBreak.Text = "BT: " + $script:BreakTime.ToString("hh\:mm\:ss")
            }
            break
        }
    }
}


$timer = New-Object 'System.Windows.Forms.Timer'
$timer.Enabled = $True 
$timer.Interval = 1000 # ms -> 1 s
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
    [void] $MainWindow.Close()
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

$BreakButton = New-Object System.Windows.Forms.Button
$BreakButton.Text = "Break" 
$BreakButton.Width = $Widht
$BreakButton.Height = $Height
$BreakButton.Location = New-Object System.Drawing.Size($Widht, $Height)
$BreakButton.Add_Click( { 
    $script:BreakType = [BreakTypes]::stopped
    $script:BreakStartTime = Get-Date
    WriteTmpFile 
    })
$MainWindow.Controls.Add($BreakButton)

$Resume = New-Object System.Windows.Forms.Button
$Resume.Text = "Resume" 
$Resume.Width = $Widht
$Resume.Height = $Height
$Resume.Location = New-Object System.Drawing.Size($Widht, (2 * $Height))
$Resume.Add_Click( { 
    $script:BreakType = [BreakTypes]::running
    $script:TotalBreakTime = $script:BreakTime 
    WriteTmpFile
    })
$MainWindow.Controls.Add($Resume)

$BreakLock = New-Object System.Windows.Forms.Button
$BreakLock.Text = "Break and Lock" 
$BreakLock.Width = $Widht
$BreakLock.Height = $Height
$BreakLock.Location = New-Object System.Drawing.Size($Widht, (3 * $Height))
$BreakLock.Add_Click( { 
    if ($script:BreakType -eq [BreakTypes]::running)
    {
        $script:BreakStartTime = Get-Date
    }
    rundll32.exe user32.dll, LockWorkStation
    Start-Sleep -Seconds 5; # wait till user is really loged of
    $script:BreakType = [BreakTypes]::stoppedByLock
    WriteTmpFile
    })
$MainWindow.Controls.Add($BreakLock)


# Die letzte Zeile sorgt dafür, dass unser Fensterobjekt auf dem Bildschirm angezeigt wird.
[void] $MainWindow.ShowDialog()