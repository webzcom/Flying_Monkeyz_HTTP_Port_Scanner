# Flying_Monkeyz_HTTP_Port_Scanner

Flying Monkeyz Port Scanner by CyberAbyss
Version 0.1 Alpha
Released for educational purposes without warranty

This is a VBScript that runs on Windows. Do run this on external target without having VPN or remote Cloud Server.

By default, script will scan http://localhost. You can figure out how to change the variables.

If you can read VBScript, you'll be able to figure this one out.

To run, double click the .vbs file to run. 

The first time you run it, you'll see the log and error log files created. In this Alpha version, you will have to delete the log file if you want to clear it. Log file will try to create new file every time you run it. If the file already exists, it will just append to it. This if helpful if you have to puase then restart later or you can run more than one copy and have all the results go to a single output file.

Creates a log file and records each HTTP call to a port # and records a HTTP result for it. 

When script hits a web page on a port #, the HTML will be displayed in a pop-up window and log the URL and HTTP status for that URL.
