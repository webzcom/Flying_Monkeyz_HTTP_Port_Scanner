'Flying Monkeyz Port Scanner
'Author: Rick Cable (CyberAbyss)
'Version 1.0 Beta
'Released for educational purposes without warranty

'Ask for User Input
strUserSelectedIP = InputBox("Target IP address:", "Target IP")
'strUserSelectedScanType = InputBox("Type of scan? (Short, Long or Mass)", "Scan Type - S, L or M")

'Define global variables

rootPath = "C:\scripts\01-Monkeyz"
scrapePath = "C:\scripts\01-Monkeyz\scrape\"
scanPath = "C:\scripts\01-Monkeyz\scans\"

targetIP = "192.168.1.0"

if strUserSelectedIP <> "" Then
	targetIP = strUserSelectedIP
end if

target = "http://" & targetIP
sTarget = "https://" & targetIP
strNewLine = Chr(13) & Chr(10)
outFile = scanPath & "log-" & targetIP & ".csv" 		'File Path
errorLogFile = "error-log.csv"	'Error Log File Path
showFoundMessage = false
logCalls = False
logOnEvery = 100
'Example shows 10 * 10000 form miliseconds to seconds
httpTimeout = 500
commonPortsList = "80,81,88,443,1337,5000,8080,32400,554,555,1024,1337,4840,7447,8554,7070,10554,6667,8081,8090,9100,19999"
'commonPortsList = "80"
arrCommonPorts = split(commonPortsList,",")
'Common target types
'تلگرام is Telegram in Persian
'Пустая страница is Empty Page in Russian
'торрент трекер is torrent tracker in Russian
'luxteb is Iranian Medical Software Company. Found 4/4/2023
'redirect_suffix is for QNAP NAS redirect page found 4/3/2023
strTargetTypes = "Tor Exit Server,SCADA,Swagger UI,SmarterMail,OctoPos,phpMyAdmin,Looking Glass Point,Plesk,OoklaServer,Nagios,HTTP Parrot,Welcome to CentOS,Index of,payment method,listing:,Client sent an HTTP request to an HTTPS server,Ruby on Rails,Tor Exit Router,The Shadowserver Foundation,Georgia Institute of Technology,CentOS-WebPanel,PHP Version,luxteb,popper.js,Nexus Repository Minecraft Server,Manager,hospital,ISPmanager,defaultwebpage.cgi,.asp?,index.js,Synology,IIS,Apache,Node Exporter,Plone,webcam,webcamXP,Webmail,redirect_suffix,NextFiber Monitoring,Nexcess,nginx,router configuration,Network Security Appliance, NAS,Admin Panel,IKCard Web Mail,Amazon ECS,Unknown Domain,Lucee,ZITADEL • Console,OpenResty,NETSurveillance,WEB SERVICE,Bootstrap Theme,Blog,Coming Soon,Droplet,Your new web server,تلگرام,ASP.NET,Video Collection,You need to enable JavaScript,Пустая страница,торрент трекер,CTF platform,qBittorrent,Shared IP,没有找到站点,webui,Login,content is to be added"
arrTargetTypes = Split(strTargetTypes,",")
doWeScrapeContent = "true"
strDoNotScrapeList = "nginx,qBittorrent,Apache,Node Exporter,Shared IP,Coming Soon,defaultwebpage.cgi,Plesk,Unknown Domain,Welcome to CentOS,没有找到站点"
arrDoNotScrapeList = Split(strDoNotScrapeList,",")
currentTargetType = ""
currentHTTPStatus = ""

iStep = 1
iStartPort = 444	'Max Value is 65535
iEndPort = 65536	'Max Value is 65536
isShortScan = False
isLongScan = False
isMassScan = True	'Mass Scan runs a short scan on all IP addresses in the target IP's subnet

'Create log file in CSV format and seed header row
Sub CreateLogFile(strFileName)
	'Create the File System Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'Setup file to write
	if NOT objFSO.FileExists(strFileName) then
		Set objFile = objFSO.CreateTextFile(strFileName,True)
		'Write out the header line for each of the two types of log files
		if strFileName = "error-log.csv" then
			objFile.WriteLine("Date,Module,Error")
		Else
			objFile.WriteLine("Date,URL,Event/Repsonse,Target_Type")
		end if
		
		objFile.Close
	End If
	ReportError("CreateLogFile")
End Sub

'Called from each function, checks to see if error was gernerated and report it to the screen then clear it.
Sub ReportError(strFunction)
	If Err.Number <> 0 Then
		LogErrorEventCSV Now,strFunction,Err.Description
		'WScript.Echo "Error in " & strFunction & ": " &  Err.Description
		Err.Clear
	End If
End Sub

'LogEventCSV: Called by MonitorUptime: Logs events in the csv file
Sub LogEventCSV(strDate,strURL,strStatus)
	'Create the File System Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(outfile, 8)	'8 = ForAppending https://technet.microsoft.com/en-us/library/ee198716.aspx
	'Don't check the HD space on every pass
	'Only check HD space once a day
	'if DatePart("h", Now) = 10 AND (DatePart("n", Now) > 42 AND DatePart("n", NOW) < 50)  then	'@10:05 Run this code
	temp = strDate & "," & strURL & "," & strStatus
	currentTargetType = ""

	objFile.WriteLine(temp)
	'objFile.WriteLine(temp2)
	
	objFile.Close
	ReportError("LogEventCSV")
End Sub

'http://dwarf1711.blogspot.com/2007/10/vbscript-urlencode-function.html
Function URLEncode(ByVal str)
	Dim strTemp, strChar
	Dim intPos, intASCII
	strTemp = ""
	strChar = ""
	For intPos = 1 To Len(str)
		intASCII = Asc(Mid(str, intPos, 1))
		If intASCII = 32 Then
			strTemp = strTemp & "+"
		ElseIf ((intASCII < 123) And (intASCII > 96)) Then
			strTemp = strTemp & Chr(intASCII)
		ElseIf ((intASCII < 91) And (intASCII > 64)) Then
			strTemp = strTemp & Chr(intASCII)
		ElseIf ((intASCII < 58) And (intASCII > 47)) Then
			strTemp = strTemp & Chr(intASCII)
		Else
			strChar = Trim(Hex(intASCII))
			If intASCII < 16 Then
				strTemp = strTemp & "%0" & strChar
			Else
				strTemp = strTemp & "%" & strChar
			End If
		End If
	Next
	URLEncode = strTemp
End Function

'isWebsiteOffline: Takes String URL 
'isPortalOffiline Returns (True / False)
Function isWebsiteOffline(strURL)
	On Error Resume Next
	
	if InStr(strURL, ":443") > 0 OR InStr(strURL, ":1337") Then
		strURL = Replace(strURL,"http","https")
	end if
	'Set WshShell = WScript.CreateObject("WScript.Shell")

	'Note 3/12/2023: From testing results... 
	'The timeout settings: 
	'1000: 1 HTTP call every 3 seconds
	'500: 1 HTTP call every 2 seconds
	'150: 1 HTTP call every 1 second
	'You can try to set them lower but I'm not sure what the minimum values could be w/o making it malfunction
	'Also: If you want to try and avoid detection, slow it down!
	
	'Reference: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms760403(v=vs.85)
	'Set http = CreateObject("MSXML2.ServerXMLHTTP")
	'Adding new version ServerXMLHTTP.6.0 so we can use the timeout options (only in version 3 & 6) 
	Set http = CreateObject("Msxml2.ServerXMLHTTP.6.0")
	currentTargetType = ""
	lResolve = 5 * httpTimeout 
	lConnect = 5 * httpTimeout  
	lSend = 15 * httpTimeout 
	lReceive = 15 * httpTimeout  
	http.setTimeouts lResolve, lConnect, lSend, lReceive
 	'Set http = CreateObject("Microsoft.XmlHttp")
	http.open "GET", strURL, False
	http.send ""

	'Only check for error of the HTTP Get request for 200 or 404 code returned. If any status is returned then the server is up
	if err.number <> 0 Then	'Site is down
		'msgbox("I've got an error")
		isWebsiteOffline = True
		SweepHasErrors = True
	Else
		'Wscript.Echo "error" & Err.Number & ": " & Err.Description
		isWebsiteOffline = False
	End If
	
	if isWebsiteOffline = False then
		if showFoundMessage then
			msgbox("Found!")
		end if	

		'Check to see if we can determine what type of site this is
		For each item in arrTargetTypes
			'msgbox(item)
			if currentTargetType <> "" Then
				Exit For
			End if
			if InStr(1,http.responseText,item,1) > 0 Then
				currentTargetType = item
			end if
		Next		
		
		'do not scrape these types
		for each item in arrDoNotScrapeList
			if currentTargetType = "" then
				doWeScrapeContent = true
			end if
			if item = currentTargetType Then
				doWeScrapeContent = false
			end if
		next	
			
		if doWeScrapeContent Then
			'msgbox("Downloading " & strURL)		
			arrTempURL = Split(strURL,":")
			if currentTargetType <> "" then
				DownLoadFile strURL, scrapePath & arrTempURL(1) & "-" & arrTempURL(2) & "-" & currentTargetType & ".html"				
			Else
				DownLoadFile strURL, scrapePath & arrTempURL(1) & "-" & arrTempURL(2) & ".html"
			end if
			
		end if
		
		'LogEventCSV Now(),strURL,http.status
        LogEventCSV Now(),strURL,http.status & " found," & currentTargetType & "!"
		currentTargetType = ""
	end if
		

	'set WshShell = Nothing
	Set http = Nothing	

	if isWebsiteOffline then
        'msgbox("Heads Up! Site is now Offline " & isWebsiteOffline & ": " & Now())
	end if
	err.clear
End Function

Function runShortScan(target)
	'Short Scan from List
	For each item in arrCommonPorts
		'msgbox("Scanning " & target1 & iLastOctet & ":" & i)
		if logCalls then
			LogEventCSV Now(),target & ":" & item,"Calling"
		end if
		call isWebsiteOffline(target & ":" & item)
	Next
	currentTargetType = ""
End Function

Function runLongScan(target)
	'Long Scan Loop
		For i = iStartPort to iEndPort Step iStep
			'msgbox("Scanning " & target1 & iLastOctet & ":" & i)
			if logCalls then
				'msgbox(i Mod logOnEvery)
				if i Mod logOnEvery = 0 then
					LogEventCSV Now(),target & ":" & i,"Calling"
				end if
			end if
			call isWebsiteOffline(target & ":" & i)
		Next
		currentTargetType = ""
End Function

Function runMassScan(target)		
	For iLastOctet = 0 to 256
		'MsgBox(target & "." & iLastOctet)	
		runShortScan(target & "." & iLastOctet)
		currentTargetType = ""
	Next	
End Function


Sub DownloadFile(url,filePath)
    Dim WinHttpReq, attempts
    attempts = 3
    'On Error GoTo TryAgain
'TryAgain:
    attempts = attempts - 1
    Err.Clear
    If attempts > 0 Then
        Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
        WinHttpReq.Open "GET", url, False
        WinHttpReq.send
		currentHTTPStatus = WinHttpReq.Status
		'if currentHTTPStatus = 400 Then
			'Try with HTTPS
		'	tempURL = Replace(url,"http","https")
		'	DownloadFile tempURL,filePath
		'end if
		
        If WinHttpReq.Status = 200 Then
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write WinHttpReq.responseBody
            oStream.SaveToFile filePath, 2 ' 1 = no overwrite, 2 = overwrite
            oStream.Close
        End If
    End If
	currentTargetType = ""
End Sub

'Start processing commands here!
CreateLogFile("error-log.csv")
CreateLogFile(outfile)

	call CreateLogFile(outFile)
	call CreateLogFile(errorLogFile)

	if isShortScan then
		runShortScan(target)
	end if
	
	if isLongScan then
		runLongScan(target)
	end if
	
	if isMassScan then
		arrTemp = Split(target, ".")
		target = arrTemp(0) & "." & arrTemp(1) & "." & arrTemp(2)
		runMassScan(target)
	end if

msgbox("Completed port scan on " & target & " with port # " & i & ".")
