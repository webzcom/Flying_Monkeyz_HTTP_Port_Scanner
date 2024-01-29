'Outside Loop

	Dim objShell
	Set objShell = Wscript.CreateObject("WScript.Shell")

'Separate IP subnets with a pipe character. This is the list we will loop thru
'strIPList = "192.167.|192.168."

arrIPList = Split(strIPList,"|")

'Monkey Launcher Sleep Timer Setting Estimates with Hard Wired Network Connection: 
'With common ports set your minutes to:
'Web Cams, set your sleep minutes to 180
iSleepTimer = 180 

For each item in arrIPList
	strIP = item
	For i = 0 to 255
		tempIP = strIP & i & ".0"
		objShell.Run "FM.vbs " & tempIP		
	Next
		'Sleep for 40 minutes before running the next one until we're on the last one.
	'	1000 = 1 second * 60 = 1 minute * 30 minutes
		'Check to see if we are on the last IP range, if yes then don't wait to close script
		if item <> UBound(arrIPList) then
			wscript.sleep(60*1000*iSleepTimer)
		end if	
Next
' Using Set is mandatory
Set objShell = Nothing
	
