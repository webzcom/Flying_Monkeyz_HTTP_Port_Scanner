# Flying_Monkeyz_HTTP_Port_Scanner
  - Flying Monkeyz Port Scanner
  - Author: CyberAbyss (Rick Cable)
  - Version 7.1 Alpha
  - Website: https://lostinthecyberabyss.com/
  - Released for educational purposes without warranty

IMPORTANT!:
  - This is a VBScript that runs on Windows.
  - DO NOT run this on external target without using a VPN or remote Cloud Server.

3 files that make up the Flying Monkeyz HTTP Port Scanner Suite:
  - FM.vbs:
    - This is the HTTP Port Scanner. Comes configured with most basic settings.
  - Monkey_Launcher.vbs:
    - Allows you to configure a work list of IP ranges and lauch waves of Flying Monkeyz in batches until all the work is done.
  - Kill_All_Flying_Monkeyz.vbs:
    - Kills all running instances of Flying Monkeyz by killing all running VBS scripts via Windows.

FM.vbs HTTP Port Scaner:
  - FM.vbs is the HTTP Port Scanner. 
  - Make sure you have a "scans" and "scans_completed" folder setup before running the script like I have it in the repo.
  - I recommend trying to scan your local network first like 192.168.1.0

## Configuring Flying Monkeyz (FM.vbs)

You can view a quick start guide here: 
- https://www.youtube.com/watch?v=Zm88z5FdKwo&ab_channel=CyberAbyss

Before You Start:
  - Make sure you have a "scans" and "scans_completed" folder setup before running the script like I have it in the repo.
  - I recommend trying to scan your local network first like 192.168.1.0 and use this time to setup your VPN Killswtch.

DEFINE GLOBAL VARIABLES:
- rootPath = "C:\scripts\01-Monkeyz"
- scrapePath = "C:\scripts\01-Monkeyz\scrape\"
- scanPath = "C:\scripts\01-Monkeyz\scans\"
- completedPath = "C:\scripts\01-Monkeyz\scans_completed\"
 
MODES / SETTINGS: 
- Modes & Settings are set by true/false values in the configuration 

SHORT SCAN SETTINGS:
- Short Scan: Scans the most common ports from a common ports list of which there are several you can choose from.
- There are other scan settings but they are not ready for prime time. Just use the Short Scan.
- Long Scan: Scans all the ports from 1 to 65536
- Mass Scan: Runs a short scan on all IP addresses in the target IP's subnet
- Note: If you run a mass scan, you'll need to set both Long and Short Scans to False until I can fix the code.
- LONG SCAN SETTINGS:
- Step = 1: Controls the increment value of the loop (Use values like 2 or 3 for Even & Odd ports or 1000)
- iStartPort = 1: Lowest Value is 1
- iEndPort = 65536: Max HTTP port value is 65536

Common Ports Lists:
- There are several common ports lists that are available in the code like the Web Cam Ports list.
- Enable a ports list by uncommenting out the line of code for the list you want to use, comment out any other ports list lines.
- Be carefule when selecting large port lists. Each ports adds an additional loop and can affect your Monkey Launcher timing calculations
- See the Monkey Launcher documentation for more information on impact of ports lists on batch launching of Monkeyz.
- For example, a short common ports list could be just port 80 if you like, by default, ports 80, 81, 8080 are typically selecte

TARGET INFORMATION:
- This information is captured by the user input when running FM.vbs
- This information is passed via command line parameter when using Monkey_Launcher.vbs to launch FM.vbs
- Example: targetIP = "localhost" OR "192.168.1.0"
- Variables:
  - target = "http://" & targetIP
  - sTarget = "https://" & targetIP

DETECTION OF COMMON TARGET TYPES:
- Added the abililty to detect various types of targets we've gotten a response from.
- The target type column will show you what values were detected including some values in Persian and Manarin.
- Example of Target Types by string values:
  - defaultwebpage.cgi,.asp?,index.js,Synology,IIS,Apache,webcam,webcamXP,Webmail,redirect_suffix,NextFiber Monitoring,nginx,router configuration,Network Security Appliance,Login,تلگرام
-  Note: redirect_suffix is for QNAP NAS redirect page found 4/3/2024

FILE READING / WRITING
- strNewLine = Chr(13) & Chr(10)
- outFile: Output file path
- errorLogFile: Error log file path

SCRAPER CONFIGURATION: True / False
- doWeScrapeContent: If set to True, make sure you have the scraper folder path created and configured.

RECON OPTIONS:
- showFoundMessage: Display VBScript Message Box showing found content (True / False)
- logCalls: True for verbose mode to log all target HTTP calls (True / False)
- logOnEvery: If logCalls is True, you can set this value to only log on the 10th, 100th or 1000th call. Good for diagnostics info.
- httpTimeout: Set at 500 for best performance.

HOW TO RUN:
  - To run, double click the .vbs file.
  - Use Windows Task Manger to kill the wscript.exe manually

SCRAPER:
  - Added in version 2.0, you can now configure the Flying Monkeyz scraper with a path to save scraped data for all found pages.
  - You might have to manually create the scraper folder that you configure until I get chance to test and refactor if needed.

LOG FILES:
  - The first time you run it, you'll see the log and error log files created. 
  - In this Alpha version, you will have to delete the log file if you want to clear it. 
  - Log file will try to create new file every time you run it but will use the existing file and append to it. 
    Note: This if helpful if you have to puase then restart later or you can run more than one copy and have all the results go to a single output file.
  - The log file records each HTTP call to a port # and any HTTP result for it, if one was returned. If not the request will eventually timeout and script will             move on to the next port number. 


