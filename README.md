# Flying_Monkeyz_HTTP_Port_Scanner
  - Flying Monkeyz Port Scanner
  - Author: CyberAbyss
  - Version 2.0 Alpha
  - Released for educational purposes without warranty

IMPORTANT!
  - This is a VBScript that runs on Windows.
  - DO NOT run this on external target without having VPN or remote Cloud Server.
  - By default, script will scan http://localhost. You can figure out how to change it in the variables.

HOW TO CONFIGURE:
  - Use the script configuration variables to configure Flying Monkeyz.
  - Below are details of how to configure the script.
 
MODES: Modes are set by true / false values in the configuration 
- Short Scan: Scans the most common ports from list
- Long Scan: Scans all the ports from 1 t0 65000
- Mass Scan: Runs a short scan on all IP addresses in the target IP's subnet
- Note: If you run a mass scan, you'll need to set both Long and Short Scans to False until I can fix the code.

HOW TO RUN:
  - To run, double click the .vbs file.
  - Use Windows Task Manger to kill the wscript.exe manually

Pop-Up Messages on Port Found:
  -  When script hits a web page on a port #, the HTML will be displayed in a pop-up window and log the URL and HTTP status for that URL.

Srape Found Content:
  - Added in version 2.0, you can now configure a scraper path to save scraped data for all found pages.
  - You might have to manually create the scraper folder that you configure until I get chance to test and refactor if needed.

LOG FILES:
  - The first time you run it, you'll see the log and error log files created. 
  - In this Alpha version, you will have to delete the log file if you want to clear it. 
  - Log file will try to create new file every time you run it but will use the existing file and append to it. 
    Note: This if helpful if you have to puase then restart later or you can run more than one copy and have all the results go to a single output file.
  - The log file records each HTTP call to a port # and any HTTP result for it, if one was returned. If not the request will eventually timeout and script will             move on to the next port number. 


