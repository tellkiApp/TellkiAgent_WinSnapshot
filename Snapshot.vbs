'###################################################################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution								              ##
'##																													              ##
'## December, 2014																									              ##
'##																													              ##
'## Version 1.0																										              ##
'##																													              ##
'## DESCRIPTION: Collect server snapshot (processes running, CPU and memory usage,...). See widget "Snapshot Windows".			  ##
'##																													              ##
'## SYNTAX: cscript "//Nologo" "//E:vbscript" "//T:90" "Snapshot.vbs" <HOST> <USERNAME> <PASSWORD> <DOMAIN>             		  ##
'##																													              ##
'## EXAMPLE: cscript "//Nologo" "//E:vbscript" "//T:90" "Snapshot.vbs" "10.10.10.1" "user" "pwd" "domain"	              		  ##
'##																													              ##
'## README:	<USERNAME>, <PASSWORD> and <DOMAIN> are only required if you want to monitor a remote server. If you want to use this ##
'##			script to monitor the local server where agent is installed, leave this parameters empty ("") but you still need to   ##
'##			pass them to the script.																						      ##
'## 																												              ##
'###################################################################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 4 Then 
	CALL ShowError(3,0) 
End If
'Set Culture - en-us
SetLocale(1033)

'INPUTS
Dim Host, Username, Password, Domain, Handles
Host = WScript.Arguments(0)
Username = WScript.Arguments(1)
Password = WScript.Arguments(2)
Domain = WScript.Arguments(3)


Dim objSWbemLocator
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

Dim Counter, FullUserName
Counter = 0

	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If
	if Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
		WScript.Quit (222)
	End If
	if Err.Number = -2147024891 Then
		CALL ShowError(2, Host)
	End If
	If Err Then CALL ShowError(1, Host)

Dim objSWbemServices, objItem, colItems, output, colIds, objItem2, objItem3, colMemory, INFO_TYPE, info, result, Processes, Threads, committedbytes
Dim FreePhysicalMemory, FreeVirtualMemory, Physical_Memory_Used, Virtual_Memory_Used, colMemory2
Dim objItem4, colCPU, PercentProcessorTime
Dim strList, processid,output_process, objItem5, memory_status, objItem6, colUptime, Uptime
Dim arrProcesses, strProcesses, PID
handlecount = 0
Processes = 0 
Threads = 0
processid = 0


INFO_TYPE="4"
Set objSWbemServices = GetObject("winmgmts:\\" & Host & "\root\cimv2")

Set colItems = objSWbemServices.ExecQuery("Select ThreadCount,HandleCount from Win32_Process where Name<>'System Idle Process' and Name<>'_Total'",,16)
If colItems.Count <> 0 Then
For Each objItem in colItems
	processid = objItem.processid
	Handles = Handles + objItem.HandleCount
	Threads = Threads + objItem.ThreadCount
	Processes = Processes + 1
Next
Else
	'If there is no response in WMI query
	CALL ShowError(5, Host)
End If
Dim nTotalSeconds,nHours,nMinutes,nSeconds,Creationtime
Set colIds = objSWbemServices.ExecQuery("Select IDProcess,ElapsedTime,Name,VirtualBytes,PercentProcessorTime,WorkingSet,IOReadBytesPerSec,IOWriteBytesPerSec,PageFaultsPerSec from Win32_PerfFormattedData_PerfProc_Process where Name <> '_Total' and Name <> 'Idle' ",,16)
	If colIds.Count <> 0 Then
	For Each objItem2 In colIds
	   output_process = output_process & objItem2.Name & ";" & objItem2.IDProcess & ";" & objItem2.elapsedtime & ";" & objItem2.PercentProcessorTime & ";" & objItem2.WorkingSet/1024 & ";" & objItem2.VirtualBytes/1024 & ";" & objItem2.IOReadBytesPerSec & ";" & objItem2.IOWriteBytesPerSec & ";" & objItem2.PageFaultsPerSec & "^"
	Next
	Else
		'If there is no response in WMI query
		CALL ShowError(5, Host)
	End If
	'TOTALS
	'Memory
	
	Set colMemory = objSWbemServices.ExecQuery("select CommittedBytes from Win32_PerfFormattedData_PerfOS_Memory",,16)	
	If colMemory.Count <> 0 Then
	For Each objItem3 in colMemory
			committedbytes = Round(objItem3.CommittedBytes/1048576)
	Next
	Else
		'If there is no response in WMI query
		CALL ShowError(5, Host)
	End If
	Set colMemory2 = objSWbemServices.ExecQuery("select FreePhysicalMemory,FreeVirtualMemory,TotalVirtualMemorySize,TotalVisibleMemorySize from Win32_OperatingSystem",,16) 
		If colMemory2.Count <> 0 Then
		For Each objItem5 in colMemory2
			'FreePhysicalMemory
			FreePhysicalMemory = FormatNumber(objItem5.FreePhysicalMemory/1024)
			'FreeVirtualMemory
			FreeVirtualMemory = FormatNumber(objItem5.FreeVirtualMemory/1024)
			'%Physical Memory Used
			Physical_Memory_Used = FormatNumber((objItem5.TotalVisibleMemorySize-objItem5.FreePhysicalMemory)/1024)
			'%Virtual Memory Used
			Virtual_Memory_Used = FormatNumber((objItem5.TotalVirtualMemorySize-objItem5.FreeVirtualMemory)/1024)
		Next
		Else
			'If there is no response in WMI query
			CALL ShowError(5, Host)
		End If
	memory_status= committedbytes & ";" & FreePhysicalMemory & ";" & FreeVirtualMemory & ";" & Physical_Memory_Used & ";" & Virtual_Memory_Used
	
	'CPU
	
	Dim PercentPrivilegedTime, PercentUserTime, Processor
	Set colCPU = objSWbemServices.ExecQuery("select PercentPrivilegedTime, PercentUserTime from Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'",,16) 
		If colCPU.Count <> 0 Then
		For Each objItem4 in colCPU
			PercentPrivilegedTime = objItem4.PercentPrivilegedTime
			PercentUserTime = objItem4.PercentUserTime
			Processor = PercentUserTime & ";" & PercentPrivilegedTime
		
	Next
	Else
		'If there is no response in WMI query
		CALL ShowError(5, Host)
	End If
	
	'Uptime
	Set colUptime = objSWbemServices.ExecQuery("Select SystemUpTime from Win32_PerfFormattedData_PerfOS_System",,16) 
		If colUptime.Count <> 0 Then
		For Each objItem6 in colUptime
			Uptime = objItem6.SystemUpTime
		Next
		Else
			'If there is no response in WMI query
			CALL ShowError(5, Host)
		End If

	result = output_process & "?" & Handles & "?" & Threads & "?" & Processes & "?" & memory_status & "?" & Processor & "?" & Uptime
	If result = "" Then
		CALL ShowError(5, Host)
	Else
		info = ToUTC() & "|" & INFO_TYPE & "|"
		output = info & result
		WScript.Echo output
	End If


If Err.number <> 0 Then
	Err.Clear
End If
	
If Err Then 
	CALL ShowError(1, 0)
Else
	WScript.Quit(0)
End If	
	
Function ToUTC()
	Dim dtmDateValue, dtmAdjusted
	Dim objShell, lngBiasKey, lngBias, k, UTC
	dtmDateValue = Now()
	'Obtain local Time Zone bias from machine registry.
	Set objShell = CreateObject("Wscript.Shell")
	lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
	If (UCase(TypeName(lngBiasKey)) = "LONG") Then
		lngBias = lngBiasKey
		ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
			lngBias = 0
		For k = 0 To UBound(lngBiasKey)
			lngBias = lngBias + (lngBiasKey(k) * 256^k)
		Next
	End If
	'Convert datetime value to UTC.
	UTC = DateAdd("n", lngBias, dtmDateValue)
	ToUTC =  FormatDateTime(UTC,2) & " " & FormatDateTime(UTC,3)
End Function

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg 
	WScript.Quit(ErrorCode)
End Sub



