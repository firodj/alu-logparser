#$language = "VBScript"
#$interface = "1.0"

' SecureCRT Script

' author: Fadhil Mandaga <firodj@gmail.com>
' update: 09 des 2013
' update: 20 nov 2013
' update: 25 apr 2013
' update: 27 jun 2014

' NOTE:
' TODO: Please set up g_user and g_password

crt.Screen.Synchronous = True
' Set objTab = crt.GetScriptTab
Set g_fso   = CreateObject("Scripting.FileSystemObject")

Const g_user     = ""
Const g_password = "!"

' TODO: Please specify exact prompt (usually hostname prompt on linux machine)
' Before the ssh connection take over !
Dim g_BasePrompts(3)
g_BasePrompts(0) = "#" 
g_BasePrompts(1) = ">"
g_BasePrompts(2) = "bash-3.00$ "
g_BasePrompts(3) = " # "
	
Const ForReading = 1
Const ICON_QUESTION = 32 
Const BUTTON_YESNO = 4
Const IDYES = 6             ' Yes button clicked
Const IDNO = 7              ' No button clicked

Class RouterBox
	Public ipAddr
	Public hostName
	Public chassisType
	Public chassisSerial
End Class

Function IsEmptyArray(arr)
	IsEmptyArray = True
    If Not IsArray(arr) Then : Exit Function

	'' Test the bounds
    On Error Resume Next

        ub = UBound(arr)
        If (Err.Number <> 0) Then
			Err.Clear
		Else
			IsEmptyArray = False
		End If

    On Error Goto 0
End Function                  ' IsDimmedArray(arrParam)

Function LoginProvider(ipaddr)
	LoginProvider = False
	crt.Screen.Send "ssh " & g_user & "@" & ipaddr & chr(13)
	str = crt.Screen.ReadString("Unable to connect", "No route to destination", "password: ", "connecting (yes/no)?", 10) 
	If crt.Screen.MatchIndex <= 2 Then
		crt.Screen.SendKeys "^c"
		crt.Screen.Send "# skip... " & ipaddr & chr(13)
		Exit Function
	End If
	If crt.Screen.MatchIndex = 4 Then
		crt.Screen.Send "yes" & chr(13)
		crt.Screen.WaitForString "password: "
	End If
	crt.Screen.Send g_password & chr(13)
	str = crt.Screen.ReadString("password: ", "A:", "B:", 10) 
	if crt.Screen.MatchIndex <= 1 Then
		crt.Screen.SendKeys "^z"
		MsgBox "Incorrect Password"
		Exit Function
	End If
	LoginProvider = True
End Function

Function ConnectRouter(ipaddr, ByRef router)
	ConnectRouter = False
	
	result = LoginProvider(ipaddr)
	If Not result Then: Exit Function
	
	ConnectRouter = True
	router.ipAddr = ipaddr
	
	nodename = crt.Screen.ReadString("#", 10)
	' nodename = crt.Dialog.Prompt("Node Name:", "Node System Name", nodename)
	For i = 1 to Len(nodename)
		c = Mid(nodename,i,1)
		Do While True
			if Asc(UCase(c)) >= Asc("A") and Asc(UCase(c)) =< Asc("Z") Then Exit Do
			if Asc(c) >= Asc("0") and Asc(UCase(c)) <= Asc("9") Then Exit Do
			if c = "-" or c ="_" Then Exit Do
			MsgBox "Invalid Node Name Character: " & c & " for " & nodename
			Exit For
		Loop
	Next
	
	''' environment no more
	crt.Screen.Send "environment no more" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	''' show time
	crt.Screen.Send "show time" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"

	''' show chassis
	crt.Screen.Send "show chassis" & chr(13)
	
	crt.Screen.WaitForString chr(10) & "Chassis Information"
	crt.Screen.WaitForString "Name"
	crt.Screen.WaitForString ": "
	router.hostName = Trim(crt.Screen.ReadString( chr(13) ))
	
	crt.Screen.WaitForString "Type"
	crt.Screen.WaitForString ": "
	router.chassisType = Trim( crt.Screen.ReadString( chr(13) ) )
	
	crt.Screen.WaitForStrings chr(10) & "Hardware Data", chr(10) & "  Hardware Data"
	crt.Screen.WaitForString "Serial number"
	crt.Screen.WaitForString ": "
	router.chassisSerial = Trim( crt.Screen.ReadString( chr(13) ) )
	
	crt.Screen.WaitForString ":" & nodename & "#"
	
	ConnectRouter = True
End Function

Function DoLogout()
	crt.Screen.Send "logout" & chr(13)
	crt.Screen.WaitForStrings "foreign host.", "closed.", 20
	crt.Screen.WaitForStrings "#", "$", ">"
End Function
	
Function DoShowPort(ByRef router)
	DoShowPort = False
	nodename = router.hostName
	
	''' show port
	crt.Screen.Send "show port" & chr(13)
	Dim arMatching(10)
	For i = 0 to 9
		arMatching(i) = chr(10) & CStr(i+1) & "/"
	Next
	arMatching(10) = ":" & nodename & "#"

	Dim arPorts()
	j = 0
	
	Do While True
		str = crt.Screen.ReadString( arMatching, 10)
		If crt.Screen.MatchIndex = 11 Then
			Exit Do
		ElseIf crt.Screen.MatchIndex > 0 Then
			str1 = Mid(arMatching(crt.Screen.MatchIndex-1),2)
			str2 = crt.Screen.ReadString( chr(13), 10 )
			str3 = str1 & str2
			If Len(str3) < 50 Then
				portId = Trim(str3)
				crt.Screen.ReadString chr(10)
				str3 = crt.Screen.ReadString( chr(13), 10 )
			Else
				portId = Trim(Mid(str3, 1, 11))
			End If
			
			adminState = Trim(Mid(str3, 13, 5))
			operState = Trim(Mid(str3, 24, 7))
			portMode = Trim(Mid(str3, 47, 4))
			portType = Trim(Mid(str3, 57, 5))
			sfpModel = Trim(Mid(str3, 64))
			
			sfpModel4 = Left(sfpModel, 4)
			If (sfpModel4 = "GIGE" Or sfpModel4 = "10GB" Or sfpModel4 = "OC3-") _
				And adminState = "Up" And portMode <> "accs" Then
				Redim Preserve arPorts(j)
				'MsgBox portId & sfpModel & portMode
				arPorts(j) = portId
				j = j + 1
			End if
		Else
			MsgBox "Fail to parsing ports"
			Exit Do
		End If
	Loop
	
	''' show port description
	crt.Screen.Send "show port description" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	'' show each filtered ports
	If not IsEmptyArray(arPorts) Then
		For j = 0 to Ubound(arPorts)
			crt.Screen.Send "show port " &  arPorts(j) & chr(13)
			crt.Screen.WaitForString ":" & nodename & "#"
		Next
	End if
	
	DoShowPort = True
End Function

Function DoPreventiveMaintenance(ByRef router)
	DoPreventiveMaintenance = False
	
	nodename = router.hostName
	
	crt.Screen.Send "show time" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show users" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	''' Version
	crt.Screen.Send "show version" & chr(13)	
	crt.Screen.WaitForString "TiMOS-"
	str = crt.Screen.ReadString(" ", 10)
	pos = InStr(str, "-")
	If pos > 0 Then
		timos = "TiMOS-" & Left(str,pos-1)
		str = Right(str, Len(str)-pos)
	End If
	pos = InStr(str, ".")
	if pos > 0 Then
		major = CInt(Left(str, pos-1))
		str = Right(str, Len(str)-pos)
	End If
	pos = InStr(str, ".R")
	if pos > 0 Then
		minor = CInt(Left(str, pos-1))
		release = CInt(Right(str, Len(str)-(pos+1)))
	End If
	
	'result = crt.Dialog.MessageBox(timos & ";" & major & ";" & minor & ";" & release, "TiMOS Version", BUTTON_YESNO)
	'DoPreventiveMaintenance  = True
	'Exit Function
	
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show bof" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show system information" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show redundancy synchronization" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show system cpu" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show system ntp all" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show system security authentication" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "file dir cf3-a:" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "file dir cf3-b:" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	If major <= 3 Then
		crt.Screen.Send "show system sync-if-timing summary" & chr(13)
	Else
		crt.Screen.Send "show system sync-if-timing" & chr(13)
	End If
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "tools dump system-resources" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	''' Config
	crt.Screen.Send "admin display-config" & chr(13)
	crt.Screen.WaitForString "# Finished"
	crt.Screen.WaitForString ":" & nodename & "#"
	
	''' Log
	crt.Screen.Send "show log syslog" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show log log-id 99" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show log log-id 100" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show log snmp-trap-group" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	''' Hardware
	crt.Screen.Send "show card state" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show card detail" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show mda detail" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	DoShowPort(router)
	
	''' Network
	crt.Screen.Send "show router interface" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router route-table summary" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router static-route" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router ospf neighbor" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router ospf interface" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router mpls interface" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router ldp interface" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router ldp session" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router rsvp session" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show router bgp summary" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	''' Service
	crt.Screen.Send "show service customer" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show service service-using" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show service sdp" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show service sdp-using" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show service sap-using" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show service fdb-mac" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	''' Environment
	crt.Screen.Send "show system threshold" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	DoPreventiveMaintenance = True
End Function

Function CreateTechSupport(ipaddr)
	crt.Screen.Send "telnet " & ipaddr & chr(13)
	str = crt.Screen.ReadString("Unable to connect", "No route to destination", "Login: ", 30) 
	If crt.Screen.MatchIndex <= 2 Then
		crt.Screen.Send "# skip... " & ipaddr & chr(13)
		CreateTechSupport = True
		Exit Function
	End If
    
	crt.Screen.Send g_user & chr(13)
	crt.Screen.WaitForString "Password: "
	crt.Screen.Send g_password & chr(13)
    
	str = crt.Screen.ReadString("Login: ", "A:", "B:", 10) 
	if crt.Screen.MatchIndex <= 1 Then
		MsgBox "Incorrect Password"
		CreateTechSupport = False
		Exit Function
	End If
	
	nodename = crt.Screen.ReadString("#", 10)
	' nodename = crt.Dialog.Prompt("Node Name:", "Node System Name", nodename)
	For i = 1 to Len(nodename)
		c = Mid(nodename,i,1)
		Do While True
			if Asc(UCase(c)) >= Asc("A") and Asc(UCase(c)) =< Asc("Z") Then Exit Do
			if Asc(c) >= Asc("0") and Asc(UCase(c)) <= Asc("9") Then Exit Do
			if c = "-" or c ="_" Then Exit Do
			MsgBox "Invalid Node Name Character: " & c & " for " & nodename
			Exit For
		Loop
	Next
	
	crt.Screen.Send "environment no more" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "file dir *.tec" & chr(13)
	number = 1
	Do While True
		str = crt.Screen.ReadString(" " & nodename & "_", ":" & nodename & "#", 10)
		If crt.Screen.MatchIndex = 1 Then
			sequence = crt.Screen.ReadString(".tec", 10)
			pos = InStrRev(sequence, "-")
			If pos > 0 Then
				new_number = CInt(Mid(sequence, pos+1)) + 1
				If new_number > number Then number = new_number
			End If
		ElseIf crt.Screen.MatchIndex = 2 Then
			Exit Do
		Else
			MsgBox "Aborted, Failed To Parsing CF Contents"
			CreateTechSupport = False
			Exit Function
		End If
	Loop
	
	crt.Screen.Send "show users" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	crt.Screen.Send "show time" & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	dd = Day(Now())
	If Len(dd) < 2 Then dd = "0" & dd
	mm = Month(Now())
	If Len(mm) < 2 Then mm = "0" & mm
	yy = Mid(Year(Now()), 3)
	
	today = dd & mm & yy
	command = "admin tech-support cf3:\" & nodename & "_" & today & "-" & number & ".tec"
	
	crt.Screen.Send command & chr(13)
	crt.Screen.WaitForString ":" & nodename & "#"
	
	DoLogout()
	
	CreateTechSupport = True
End Function

Sub Main

	Do While True
		cliprompt = crt.Screen.Get( crt.Screen.CurrentRow, 1, crt.Screen.CurrentRow, crt.Screen.CurrentColumn )
		
		For i = 0 to UBound(g_BasePrompts)
			If InStr(cliprompt, g_BasePrompts(i)) > 0 Then Exit Do
		Next
		
		crt.Screen.SendKeys "^z"
		crt.Screen.WaitForString "#"
		
		' Should I use DoLogout
		crt.Screen.Send "logout" & chr(13)
		crt.Screen.WaitForStrings "#", "$", ">"
	Loop
	
	Dim arrIpAddress()
	msg = "List of Nodes:" & Chr(13)
	
    ' TODO: Please select what method to connect to many routers
	If False Then
		Set router = New RouterBox
        ' TODO: Specify IP address
		result = ConnectRouter("", router)
		
		'if result Then: DoPreventiveMaintenance(router)
		if result Then 
			DoShowPort(router)
			DoLogout()
		End if
		
		MsgBox router.hostName
		Exit Sub
	ElseIf False Then
		ipaddr = crt.Dialog.Prompt("Enter node loopback IP:", "Node IP Address")
		If ipaddr = "" Then Exit Sub
		
		Redim Preserve arrIpAddress(0)
		arrIpAddress(0) = ipaddr
		msg = msg & ipaddr & " " & Chr(13)
	Else
        ' TODO: fill the ip adress of routers on ip_adress.txt file
		filelist = crt.Dialog.Prompt("Enter path\file to open:", "File List", g_fso.GetFolder(".") & "\ip_address.txt")
		if filelist = "" Then Exit Sub
		
		i = 0
		Set objFile = g_fso.OpenTextFile(filelist, ForReading)
		
		Do Until objFile.AtEndOfStream
			ipaddr = Trim(objFile.ReadLine)
			pos = InStr(ipaddr, "#")
			If pos > 0 Then
				ipaddr = Trim(Left(ipaddr, pos-1))
			End If
			If ipaddr <> "" Then				
				Redim Preserve arrIpAddress(i)
				arrIpAddress(i) = ipaddr
				i = i + 1
				msg = msg & i & ") " & ipaddr & Chr(13)
			End If
		Loop
		objFile.Close
	End If
	
	result = crt.Dialog.MessageBox(msg, "IP Address List", ICON_QUESTION or BUTTON_YESNO)
	If result = IDYES Then
		For i = 0 to UBound(arrIpAddress)
			Set router = New RouterBox
			ipaddr = arrIpAddress(i)
			result = ConnectRouter(ipaddr, router)
			dontbreak = True
			if result Then
				dontbreak = DoPreventiveMaintenance(router)
				'dontbreak = DoShowPort(router) ' Only
				DoLogout()
			End if
			if not dontbreak Then
				MsgBox "Stop at " & ipaddr
				Exit Sub
			End if
		Next
	End If
	
	MsgBox "Done."
End Sub
