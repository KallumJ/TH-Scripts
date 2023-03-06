# Generate Inactive Account Report
A script to generate a report of users who have no activity in the last 90 days.

# Running the script
The script can be run from a Powershell terminal like:
```
.\GenerateInactiveAccountReport.ps1
```
or by right clicking the file and selecting "Run with PowerShell".

The generated report can be found in the same folder the script was run from. The report will be titled "inactive-users.csv"

# Notes
- The script will prompt for login twice, once for Azure, once for MS365. Sometimes the prompts can open behind active windows, so if you don't see it, check behind your foreground windows.

- Make sure any previous "inactive-users.csv" is deleted before rerunning the script.