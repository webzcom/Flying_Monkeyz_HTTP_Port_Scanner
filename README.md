Before You Start:
  - Make sure you have a "scans" and "scans_completed" folder setup before running the script like I have it in the repo.
  - I recommend trying to scan your local network first like 192.168.1.0

FM.vbs HTTP Port Scaner:
  - Make sure you have a "scans" and "scans_completed" folder setup before running the script like I have it in the repo.
  - I recommend trying to scan your local network first like 192.168.1.0 and use this time to setup your VPN Killswtch.

DEFINE GLOBAL VARIABLES:
- rootPath = "C:\scripts\01-Monkeyz"
- scrapePath = "C:\scripts\01-Monkeyz\scrape\"
- scanPath = "C:\scripts\01-Monkeyz\scans\"
- completedPath = "C:\scripts\01-Monkeyz\scans_completed\"
 
MODES: Modes are set by true / false values in the configuration 

- Short Scan: Scans the most common ports from list of which there are several.
SHORT SCAN SETTINGS:
- commonPortsList: I've left two lines for the common ports list. There maybe more.
- A short list w/ port 80, 81, 8080 only for use with mass scan to speed it up then a longer list "80,443,5000,8080,32400,554,88,81,555,7447,8554,7070,10554,6667,8081,8090"
- I've added addtional common ports lists like web cameras which you can uncomment to enable.
- arrCommonPorts: Holds the array of common ports created from the list.

  
- Long Scan: Scans all the ports from 1 to 65536
- Mass Scan: Runs a short scan on all IP addresses in the target IP's subnet
- Note: If you run a mass scan, you'll need to set both Long and Short Scans to False until I can fix the code.
- LONG SCAN SETTINGS:
- Step = 1: Controls the increment value of the loop (Use values like 2 or 3 for Even & Odd ports or 1000)
- iStartPort = 1: Lowest Value is 1
- iEndPort = 65536: Max HTTP port value is 65536

TARGET INFORMATION
- targetIP = "localhost" OR "192.168.1.0"
- target = "http://" & targetIP
- sTarget = "https://" & targetIP

DETECTION OF COMMON TARGET TYPES (V. 2.1)
- I've added the abililty to detect various types of targets we've scraped
- The comments will show you what values are detected including some values in Persian
- redirect_suffix is for QNAP NAS redirect page found 4/3/2024
Target Types by these string values: defaultwebpage.cgi,.asp?,index.js,Synology,IIS,Apache,webcam,webcamXP,Webmail,redirect_suffix,NextFiber Monitoring,nginx,router configuration,Network Security Appliance,Login,تلگرام

FILE PATHS:
- rootPath = "C:\scripts\01-Monkeyz"
- scrapePath = "C:\scripts\01-Monkeyz\scrape\"

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


