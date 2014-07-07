'On Error Resume Next

Set objShell = CreateObject("WScript.Shell")
Set WshSysEnv = objShell.Environment("Process") 

'strHostname = "10.1.40.100"
'strDomain = "uptime-demo"
'strUser = "administrator"
'strPassword = ""

strHostname = WshSysEnv("UPTIME_HOSTNAME")
strDomain = WshSysEnv("UPTIME_USER_DOMAIN")
strUser = WshSysEnv("UPTIME_USERNAME")
strPassword = WshSysEnv("UPTIME_PASSWORD")

'WScript.Echo strHostname
'WScript.Echo strDomain
'WScript.Echo strUser
'WScript.Echo strPassword

strCommandNoPass = "cmd /c systeminfo /s " & strHostname & " /u " & strDomain & "\" & strUser & " /p "
strCommand = strCommandNoPass & strPassword
'strCommand = "cmd /c systeminfo"

'WScript.Echo strCommand

Set objCmdExec = objShell.Exec(strCommand)
strCommandOutput = objCmdExec.StdOut.ReadAll

'WScript.Echo strCommandOutput

If strCommandOutput = "" Then
	WScript.Echo "Error: Check user credentials. Confirm you can run the following from the monitoring station command line: " & strCommandNoPass & "PASSWORD"
	Wscript.Quit 2
Else
	strArray = Split(strCommandOutput,VbCrLf)	' break up the output by newlines
	'WScript.Echo UBound(strArray)

	processor = False

	For Each line In strArray
		If Not processor Then
			If (InStr(line, "OS Name") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "OS_NAME " & Trim(strTemp(1))
			ElseIf (InStr(line, "OS Version") <> 0) AND (InStr(line, "BIOS Version") = 0) Then	'get OS Version and not BIOS Version
				strTemp = Split(line,": ")
				WScript.Echo "OS_VERSION " & Trim(strTemp(1))
			ElseIf (InStr(line, "OS Manufacturer") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "OS_MANUFACTURER " & Trim(strTemp(1))
			ElseIf (InStr(line, "OS Build Type") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "OS_BUILD_TYPE " & Trim(strTemp(1))
			ElseIf (InStr(line, "Original Install Date") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "ORIGINAL_INSTALL_DATE " & Trim(strTemp(1))
			ElseIf (InStr(line, "System Boot Time") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "SYSTEM_BOOT_TIME " & Trim(strTemp(1))
			ElseIf (InStr(line, "System Manufacturer") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "SYSTEM_MANUFACTURER " & Trim(strTemp(1))
			ElseIf (InStr(line, "System Model") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "SYSTEM_MODEL " & Trim(strTemp(1))
			ElseIf (InStr(line, "System Type") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "SYSTEM_TYPE " & Trim(strTemp(1))
			ElseIf (InStr(line, "Processor(s)") <> 0) Then
				processor = True
			ElseIf (InStr(line, "Time Zone") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "TIME_ZONE " & Trim(strTemp(1))
			ElseIf (InStr(line, "Total Physical Memory") <> 0) Then
				strTemp = Split(line,": ")
				strTemp2 = Split(Trim(strTemp(1))," ")
				WScript.Echo "TOTAL_PHYSICAL_MEMORY " & Replace(Trim(strTemp2(0)),",","")
			ElseIf (InStr(line, "Available Physical Memory") <> 0) Then
				strTemp = Split(line,": ")
				strTemp2 = Split(Trim(strTemp(1))," ")
				WScript.Echo "AVAILABLE_PHYSICAL_MEMORY " & Replace(Trim(strTemp2(0)),",","")
			ElseIf (InStr(line, "Virtual Memory: Max Size") <> 0) Then
				strTemp = Split(line,": ")
				strTemp2 = Split(Trim(strTemp(2))," ")
				WScript.Echo "VIRTUAL_MEMORY_MAX_SIZE " & Replace(Trim(strTemp2(0)),",","")
			ElseIf (InStr(line, "Virtual Memory: Available") <> 0) Then
				strTemp = Split(line,": ")
				strTemp2 = Split(Trim(strTemp(2))," ")
				WScript.Echo "VIRTUAL_MEMORY_AVAILABLE " & Replace(Trim(strTemp2(0)),",","")
			ElseIf (InStr(line, "Virtual Memory: In Use") <> 0) Then
				strTemp = Split(line,": ")
				strTemp2 = Split(Trim(strTemp(2))," ")
				WScript.Echo "VIRTUAL_MEMORY_IN_USE " & Replace(Trim(strTemp2(0)),",","")
			ElseIf (InStr(line, "Domain") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "DOMAIN " & Trim(strTemp(1))
			ElseIf (InStr(line, "Logon Server") <> 0) Then
				strTemp = Split(line,": ")
				WScript.Echo "LOGON_SERVER " & Trim(strTemp(1))
			End If
		ElseIf processor Then
			If (InStr(line, "BIOS Version") <> 0) Then
				processor = False
			Else
				strTemp = Split(line,": ")
				strTemp2 = Replace(strTemp(0),"[","")	' get rid of square brackets around processor number
				strTemp3 = Replace(strTemp2,"]","")
				WScript.Echo Trim(strTemp3) & ".PROCESSOR " & Trim(strTemp(1))
			End If
		End If
	Next
End If