# System-Health-Reports
This is a Windows health reporting system built to be automated with Task Scheduler and ran on an IIS server. The script runs a full report on any number of Windows systems you point it to. Once complete, the script compiles an HTML file with the current Date of that report and an HTML file to navigate the reports.

# Screenshots:
![Imgur](https://i.imgur.com/uzProRG.png)
![Imgur](https://i.imgur.com/mvlim5J.png)

# Features:
- Lightweight, Configurable, and Expandable.
- Automatically creates HTML pages for navigating and viewing reports.
- Fully written in Powershell. No need to install anything!
- Built for Task Scheduler.
- Emails you a link to view reports when finished.
- Holds logs for as long as you want.

# What's in the report:
System Information
- Computer Name
- Operating System
- Build + Service Pack Version
- Last Boot Time
- Up Time
- CPU usage at time of report.
- Memory usage at time of report.
- Internet Connectivity by making the system ping an external address and returning the result.

Network Information
- IPv4 Address
- IPv6 Address
- MAC Address
- Address Family

Hardware Information
- Chassis: Manufacturer, Model, BIOS Version, and Service Tag
- Processor: Name, and Socket Designation
- Memory: Slot ID, Part Number, Speed, and Capacity
- Storage Controller: Manufacturer, Name, DriverName
- Disk: Drive Letter, Volume Name, Size, Free Space, and Free Space Percentage

Application Event Log

System Event Log

Running Services

Installed Programs

Hotfixes

# To Test:
- Download all files and store in same folder either locally or on a system with IIS installed.
- Update "servers.txt" with the Computer Names or IP Adresses you wish to get a report on. 
> #### Important: Remote Shell Access must be enabled and correctly configured for elevated access only.
- In "MainExec.ps1" update Lines 33 to 46 with your Email and Exchange Server information or remove completely.
- Execute MainExec.ps1 to test.

# To Deploy:
At download, the scripts are configured to be ran locally for testing. When it's time to configure the script for Task Scheduler:
- Open each ps1 file and edit the top lines at "$fpath" with the folder path to where you've stored the scripts.
> #### If you skip this step, Task Scheduler will think the base folder is System32.
- Create a New Task in Task Scheduler
- Use > Action: Start a program
- Program\script: powershell
- Add arguments: -File C:\Path\To\MainExec.ps1
- Set your time to run the script. I use 7am and Daily.

 That's it!

# Notes:
- MainExec.ps1 will produce an error if the system's Event Logs are cleared. In the security realm, this is considered a feature!
- Most runtime errors are related to the $fpath location. Be sure to check that.
- When using IIS make sure the folder containing the files is not "Read-Only" and that the user has rights to the folder.
- Feel free to update the CSS in the CreateHTML and GetSystemReport scripts.
- Keep an eye on your Reports folder. There's no code setup to delete these files. I recommend clearing anything older than 3months.

