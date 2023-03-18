
'Flying Monkeyz Port Scanner by CyberAbyss
'Version 0.3 Alpha
'Released for educational purposes without warranty

'Define global variables
targetIP = "localhost"
target = "http://" & targetIP
sTarget = "https://" & targetIP
strNewLine = Chr(13) & Chr(10)
outFile = "log-" & targetIP & ".csv" 		'File Path
errorLogFile = "error-log.csv"	'Error Log File Path
showFoundMessage = True
logCalls = True
logOnEvery = 100
'Example shows 10 * 10000 form miliseconds to seconds
httpTimeout = 500
commonPortsList = "80,443,5000,8080,32400,554,88,81,555,7447,8554,7070,10554,6667,8081,8090"
arrCommonPorts = split(commonPortsList,",")
iStep = 1
iStartPort = 30001	'Max Value is 65535
iEndPort = 65536	'Max Value is 65536
runShortScan = False
runLongScan = True

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
			objFile.WriteLine("Date,URL,Event/Repsonse")
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
			msgbox("Flying Monkeys by CyberAbyss! v1.0 Beta" & strNewLine & strNewLine & strURL & " / Port #" & i & " Found! " & strNewLine & strNewLine & http.responseText)
		end if
		'LogEventCSV Now(),strURL,http.status
        LogEventCSV Now(),strURL,http.status & " found!"
	end if

	'set WshShell = Nothing
	Set http = Nothing	

	if isWebsiteOffline then
        'msgbox("Heads Up! Site is now Offline " & isWebsiteOffline & ": " & Now())
	end if
	err.clear
End Function


'Start processing commands here!
CreateLogFile("error-log.csv")
CreateLogFile(outfile)

	call CreateLogFile(outFile)
	call CreateLogFile(errorLogFile)

'Short Scan from List
'Outside IP loop for last octet
'For iLastOctet = 0 to 255
if runShortScan then
	For each item in arrCommonPorts
		'msgbox("Scanning " & target1 & iLastOctet & ":" & i)
		LogEventCSV Now(),target1 & ":" & item,"Calling"
		call isWebsiteOffline(target & ":" & item)
	Next
end if
'Next


'Long Scan Loop
if runLongScan then
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
end if

msgbox("Completed port scan on " & target & " with port # " & i & ".")


