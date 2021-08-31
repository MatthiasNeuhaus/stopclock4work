# StopClock4Work
## Description
**Automatically** logs your worktime. 
- A "log of" from your computer is detected as a break. It will resume automatically afterwards.
- If you start your computer after hibernation the next day, counting of time will resume automatically.
- If you restart your computer while working, counting will be continued.

## How to start
Right click on InstallLinks.ps1 and run with PowerShell to add start links to autostart and the start program menu.
The tool es equipped with tooltips, so therefore (hopefully) self declaring. 
Only parts you might need to additionally understand are:

### Data folder
On first start, content of DataSource folder will be copied to (gitignored) Data Folder. So, your Data is safe even you pull the tool for an update (when I update content on DataSource folder, revision of file will be increased and a new file will be copied - your content won't get lost!)

### Content of Data folder
1. During counting you'll find a temporary file inside. This Tmp\<date>.csv file is also your information backup if anything went wrong (it will only be removed, if a line is sucessfully written to the output file. Information in the file is updated every 5 min).
2. WorkTimes file is **the** output file. You may want to reorder the column or remove unwanted. But if you add new, the script will fail while writing!
3. TargetWorkTimes is the file you may want to edit if you do not work every day 8 hours. These times are used for the countdown of today’s time (Work Time Balance). Also for the last column in the output file. If you sum this up (and maybe add an start entry of current balance) you get your actual work time balance. (Maybe this will be displayed in the tool later too)
4. The Options file let’s you configure the script. Currently you can only deactivate the automatically counting of a break in case of workstation is locked.
5. BreakTimes is the input for the calculated work time in the output file. This is the break time, that will be subtracted from your reported work time based on the work time. Therefore the script will add it for you.

Enjoy! And don’t hesitate to report enhancements via PR 😉