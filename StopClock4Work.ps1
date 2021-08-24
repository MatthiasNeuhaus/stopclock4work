# Die ersten beiden Befehle holen sich die .NET-Erweiterungen (sog. Assemblies) für die grafische Gestaltung in den RAM.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


enum BreakTypes {
    running
    breaked
    breakedByLock
}

enum FailreTypes {
    retry
    noFailure
    abortFailure
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

# Path for output file
$TimeSheetPath = Join-Path $DataFolder "WorkTimesV0.csv"

# Import User Options
Import-Module "$($DataFolder)\OptionsV0.ps1" 

# Variables for calculation
$SecondsToWriteTmp = 300 # -> means every 5min
function init() {
    $script:StartTime = Get-Date
    [BreakTypes]$script:BreakType = [BreakTypes]::running;
    $script:WorkTime = New-TimeSpan -Seconds 0;
    $script:TotalBreakTime = New-TimeSpan -Seconds 0;
    $script:BreakTime = New-TimeSpan -Seconds 0;
    $script:BreakStartTime = Get-Date
    $script:TempFileTimer = $SecondsToWriteTmp
    
    # Write tmp file for persistence during restart
    $script:TmpFileName = "Tmp" + $StartTime.Date.ToString('dd/MM/yyyy') + ".csv"
    $script:TmpFile = Join-Path $DataFolder $TmpFileName
    
    if (Test-Path $TmpFile -PathType leaf) # Temp file does exist - read it!
    {
        $ReadStart = Import-Csv -UseCulture -Path $TmpFile
        
        $script:StartTime              = [DateTime]::Parse($ReadStart.StartTime)
        $script:WorkTime               = [Timespan]::Parse($ReadStart.WorkTime)
        $script:TotalBreakTime         = [Timespan]::Parse($ReadStart.BreakTime)
        $script:BreakStartTime         = [DateTime]::Parse($ReadStart.BreakStartTime)
        $script:BreakType              = $ReadStart.BreakType
    }
    
    # Read target work times for countdown
    $TargetWorkTimesPath = Join-Path $DataFolder "TargetWorkTimesV0.csv"
    $TargetWorkTimes = Import-Csv -Path $TargetWorkTimesPath -Delimiter ";"
    $Today = $StartTime.DayOfWeek
    $script:TargetWorktime = [Timespan]::Parse($TargetWorkTimes.$Today)
    $script:WorkTimeBalance = $WorkTime - $TargetWorktime
}
init 

# Die nächste Zeile erstellt aus der Formsbibliothek das Fensterobjekt.
$MainWindow = New-Object System.Windows.Forms.Form

# Create Tooltips 
$Tooltip = New-Object System.Windows.Forms.ToolTip

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
$Height = 25
$Widht = 115

# Fenstergröße festlegen
$MainWindow.Size = New-Object System.Drawing.Size(((3 * $Widht)+16), ((6 * $Height)+14))

# Countdown in title
$MainWindow.Text = "WTB: $(if($WorkTimeBalance -lt [TimeSpan]::Zero ){"-"})$($WorkTimeBalance.ToString('hh\:mm\:ss'))"
# Countup
$Countup = New-Object System.Windows.Forms.Label
$Countup.Location = New-Object System.Drawing.Size(20, 5)
$Countup.AutoSize = $True
$Countup.Text = "WT: " + $script:WorkTime.ToString("hh\:mm\:ss")
$MainWindow.Controls.Add($Countup)
$Tooltip.SetToolTip($Countup, "Todays work time counted")

# Countup Break
$CountupBreak = New-Object System.Windows.Forms.Label
$CountupBreak.Location = New-Object System.Drawing.Size(($Widht + 20), 5)
$CountupBreak.ForeColor = "red"
$CountupBreak.AutoSize = $True
$CountupBreak.Text = "BT: " + $script:TotalBreakTime.ToString("hh\:mm\:ss")
$MainWindow.Controls.Add($CountupBreak)
$Tooltip.SetToolTip($CountupBreak, "Todays break time counted")

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

    
    $writeOutput = [PSCustomObject]@{
        DayOfWeek       = $script:StartTime.DayOfWeek.ToString() 
        Date            = $script:StartTime.Date.ToString('dd/MM/yyyy')
        DateAgain       = $script:StartTime.Date.ToString('dd/MM/yyyy')
        WorkStartTime   = $script:StartTime.TimeOfDay.ToString('hh\:mm\:ss')
        WorkTime        = $script:WorkTime.ToString('hh\:mm\:ss')
        BreakTime       = $script:TotalBreakTime.ToString('hh\:mm\:ss')
        WorkEndTime     = $WorkEndTime.TimeOfDay.ToString('hh\:mm\:ss')
        WorkEndTimeCalc = $WorkEndTimeCalc.TimeOfDay.ToString('hh\:mm\:ss')
        BreakTimeCalc   = $BreakTimeCalc.ToString('hh\:mm\:ss')
        WorkTimeBalance = "$(if($script:WorkTimeBalance -lt [TimeSpan]::Zero ){"-"})$($script:WorkTimeBalance.ToString('hh\:mm\:ss'))"
    }

    $Failure = [FailreTypes]::noFailure
    do 
    {
        $Failure = [FailreTypes]::noFailure
        try 
        {
            $writeOutput | Export-Csv -UseCulture -Path $TimeSheetPath -Append -NoTypeInformation
        }
        catch 
        {
            $wshell     = New-Object -ComObject Wscript.Shell
            $PoUpReturn = $wshell.Popup("You seem to have a write protection on $($TimeSheetPath). Please close it 😉. If you abort, you can still find the worktime a seperate file for today",0,"Excel sucks!",0x5)

            if ($PoUpReturn -eq 4) #means retry
            {
                $Failure = [FailreTypes]::retry
            }
            elseif ($PoUpReturn -eq 2) #means abort
            {
                $Failure = [FailreTypes]::abortFailure
            }
        }
    } while ($Failure -eq [FailreTypes]::retry)
    
    if ($Failure -eq [FailreTypes]::abortFailure)
    {
        $OneTimeSheetPath = Join-Path $DataFolder "WorkTime$($StartTime.Date.ToString('dd/MM/yyyy')).csv"
        $writeOutput | Export-Csv -UseCulture -Path $OneTimeSheetPath
    }

    # Remove tmp file - information stored in the correct format.
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
            if (  ( $logonStatus -eq 1) -and ($AutoStopByLock)  )
            {
                $script:BreakType = [BreakTypes]::breakedByLock
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
        
        ([BreakTypes]::breakedByLock) {

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

        ([BreakTypes]::breaked) {
            if ( $logonStatus -eq 1)
            {
                $script:BreakType = [BreakTypes]::breakedByLock
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
$Stop = New-Object System.Windows.Forms.Button
$Stop.Text = "Stop" 
$Stop.Width = $Widht
$Stop.Height = $Height
$Stop.Location = New-Object System.Drawing.Size(0, $Height)
$Stop.Add_Click( { 
    WriteToCsv
    $MainWindow.Text    = "WTB: -" + $TargetWorktime.ToString('hh\:mm\:ss')
    $CountupBreak.Text  = "BT: " + [TimeSpan]::Zero.ToString("hh\:mm\:ss")
    $Countup.Text       = "WT: " + [TimeSpan]::Zero.ToString("hh\:mm\:ss")
    $timer.Enabled = $False
    
    })
$MainWindow.Controls.Add($Stop)
$Tooltip.SetToolTip($Stop, "Stop counting todays work and write a line to output file")

$StopHybernate = New-Object System.Windows.Forms.Button
$StopHybernate.Text = "Stop and Hybernate" 
$StopHybernate.Width = $Widht
$StopHybernate.Height = $Height
$StopHybernate.Location = New-Object System.Drawing.Size(0, (2 * $Height))
$StopHybernate.Add_Click( { 
    WriteToCsv
    $timer.Enabled = $False
    rundll32.exe powrprof.dll, SetSuspendState 0, 1, 0
    })
$MainWindow.Controls.Add($StopHybernate)
$Tooltip.SetToolTip($StopHybernate, "Stop counting todays work, write a line to output file and hybernate the workstation")

$StopClose = New-Object System.Windows.Forms.Button
$StopClose.Text = "Stop and Close" 
$StopClose.Width = $Widht
$StopClose.Height = $Height
$StopClose.Location = New-Object System.Drawing.Size(0, (3 * $Height))
$StopClose.Add_Click( { 
    WriteToCsv
    $timer.Enabled = $False
    [void] $MainWindow.Close()
    })
$MainWindow.Controls.Add($StopClose)
$Tooltip.SetToolTip($StopClose, "Stop counting todays work, write a line to output file and close this tool")

$StopShutdown = New-Object System.Windows.Forms.Button
$StopShutdown.Text = "Stop and Shutdown" 
$StopShutdown.Width = $Widht
$StopShutdown.Height = $Height
$StopShutdown.Location = New-Object System.Drawing.Size(0, (4 * $Height))
$StopShutdown.Add_Click( { 
    WriteToCsv
    $timer.Enabled = $False
    shutdown /s /hybrid
    [void] $MainWindow.Close()
    })
$MainWindow.Controls.Add($StopShutdown)
$Tooltip.SetToolTip($StopShutdown, "Stop counting todays work, write a line to output file and shut down the workstation")

$Start = New-Object System.Windows.Forms.Button
$Start.Text = "Start" 
$Start.Width = $Widht
$Start.Height = $Height
$Start.Location = New-Object System.Drawing.Size($Widht, $Height)
$Start.Add_Click( { 
    init
    $timer.Enabled = $True
    WriteTmpFile
    })
$MainWindow.Controls.Add($Start)
$Tooltip.SetToolTip($Start, "Start counting todays work (if not allready started automatically)")

$BreakButton = New-Object System.Windows.Forms.Button
$BreakButton.Text = "Break" 
$BreakButton.Width = $Widht
$BreakButton.Height = $Height
$BreakButton.Location = New-Object System.Drawing.Size($Widht, (2 * $Height))
$BreakButton.Add_Click( { 
    $script:BreakType = [BreakTypes]::breaked
    $script:BreakStartTime = Get-Date
    WriteTmpFile 
    })
$MainWindow.Controls.Add($BreakButton)
$Tooltip.SetToolTip($BreakButton, "Pause counting todays work, don't forget to resume counting manually if you do not lock the workstation")

$Resume = New-Object System.Windows.Forms.Button
$Resume.Text = "Resume" 
$Resume.Width = $Widht
$Resume.Height = $Height
$Resume.Location = New-Object System.Drawing.Size($Widht, (3 * $Height))
$Resume.Add_Click( { 
    $script:BreakType = [BreakTypes]::running
    $script:TotalBreakTime = $script:BreakTime 
    WriteTmpFile
    })
$MainWindow.Controls.Add($Resume)
$Tooltip.SetToolTip($Resume, "Resume counting todays work")

$BreakLock = New-Object System.Windows.Forms.Button
$BreakLock.Text = "Break and Lock" 
$BreakLock.Width = $Widht
$BreakLock.Height = $Height
$BreakLock.Location = New-Object System.Drawing.Size($Widht, (4 * $Height))
$BreakLock.Add_Click( { 
    if ($script:BreakType -eq [BreakTypes]::running)
    {
        $script:BreakStartTime = Get-Date
    }
    rundll32.exe user32.dll, LockWorkStation
    Start-Sleep -Seconds 5; # wait till user is really loged of
    $script:BreakType = [BreakTypes]::breakedByLock
    WriteTmpFile
    })
$MainWindow.Controls.Add($BreakLock)
$Tooltip.SetToolTip($BreakLock, "Pause counting todays work and lock the workstation - tool will resume counting automatically")

$OutputFile = New-Object System.Windows.Forms.Button
$OutputFile.Text = "View Output File" 
$OutputFile.Width = $Widht
$OutputFile.Height = $Height
$OutputFile.Location = New-Object System.Drawing.Size((2 * $Widht), (4 * $Height))
$OutputFile.Add_Click( { 
    Start-Process excel $TimeSheetPath
    })
$MainWindow.Controls.Add($OutputFile)
$Tooltip.SetToolTip($OutputFile, "Open output CSV File with Excel")

# Die letzte Zeile sorgt dafür, dass unser Fensterobjekt auf dem Bildschirm angezeigt wird.
[void] $MainWindow.ShowDialog()