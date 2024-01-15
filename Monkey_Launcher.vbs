'Monkeyz Launcher for Flying Monkeyz HTTP Port Scanner
'Author: Rick Cable (CyberAbyss)
'Version 3.0 Beta
'Released for educational purposes without warranty

'Outside Loop for 3rd Octet of IP Address
	Dim objShell
	Set objShell = Wscript.CreateObject("WScript.Shell")
	
	'Web Cam IPs
	'24.255.72.234:8080/multi.html
	'46.243.108.21:8080
	'24.255.72.234:8080
	'109.233.191.228:8090
	'86.57.137.159:8082
	'75.149.26.30:1024
	'2.40.45.90
	'strIP = "2.40."

'Be careful here not to launch too many Monkeyz at once!
'This loop is for 3rd Octet of IP Address. FM21 will loop through the 4th Octect
For i = 49 to 69
	tempIP = strIP & i & ".0"
	objShell.Run "FM30.vbs " & tempIP	
Next	
	' Using Set is mandatory
	Set objShell = Nothing
