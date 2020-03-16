' Name: PatchMSIProductVersion
' Parameters: sPathToMSI = The path to the msi file
' Comment: Strip the dot build part of the 'ProductVersion' property
'

Const sQuery = "SELECT Value FROM Property WHERE Property='ProductVersion'"
Const msiOpenDatabaseModeDirect = 2
Const msiViewModifyUpdate = 2

Dim sPathToMSI, sOldVersion, sNewVersion, retVal, objWI, objDB, objView, objRecord

Main

Sub Main

	Dim objArgs, ArgCount, cArgument, objFS, sArgument

	retVal = 1
	'Get the command line parameters.
	Set objArgs	= WScript.Arguments
	ArgCount	= objArgs.Count
	
	If ArgCount = 0 Then
		Call ShowUsage
	End If
	
	Select Case UCase(CStr(WScript.Arguments(0)))
		Case "?", "/?", "-?", "H", "/H", "-H", "HELP", "/HELP"
			Call ShowUsage
	End Select

	sPathToMSI = CStr(WScript.Arguments(0))

	Set objFS = CreateObject("Scripting.FileSystemObject")
	If Not objFS.FileExists(sPathToMSI) Then
		Fail "File: '" & sPathToMSI & "' doesn't exist!"
	End If

	Set objWI = CreateObject("WindowsInstaller.Installer")
	CheckError
	Set objDB = objWI.OpenDatabase(sPathToMSI, msiOpenDatabaseModeDirect)
	CheckError

	Set objView = objDB.OpenView(sQuery)
	CheckError

	objView.Execute
	Set objRecord = objView.Fetch
	If Not objRecord Is Nothing Then
		sOldVersion = objRecord.StringData(1)

		If Not IsValidVersion(sOldVersion) = 1 Then
			Fail "Product Version '" & sOldVersion & "' is not valid! The script only works for Exchange 2007 SP3 RTM!"
		End If
		
		Wscript.Echo "Current ""ProductVersion"" is " & sOldVersion & vbCRLF

		sNewVersion = StripDotBuild(sOldVersion)

		WScript.Echo "Replace old version '" & sOldVersion & "' with new version '" & sNewVersion & "' [Y/N] "
		sArgument = WScript.StdIn.ReadLine

		Select Case UCase(sArgument)
			Case "N", "No", "n", "no", "nO"
				WScript.Quit
		End Select

		objRecord.StringData(1) = sNewVersion
		objView.Modify msiViewModifyUpdate, objRecord
		CheckError

		objView.Close
		objDB.Commit
	Else
		Fail """ProductVersion"" is not in the property table."
	End If

	WScript.Quit retVal
End Sub

Sub ShowUsage
	Dim strMessage
	
	strMessage = ""
	strMessage = strMessage & vbCRLF & "Syntax"
	strMessage = strMessage & vbCRLF & "--------------------------------------------------------------------------------"
	strMessage = strMessage & vbCRLF & "CScript " & WScript.ScriptName & " file"
	strMessage = strMessage & vbCRLF & ""
	strMessage = strMessage & vbCRLF & "Parameters"
	strMessage = strMessage & vbCRLF & "--------------------------------------------------------------------------------"
	strMessage = strMessage & vbCRLF & "file : The path to the MSI file to be patched" & vbCRLF
	strMessage = strMessage & vbCRLF & "Example : CScript " & WScript.ScriptName & " ""D:\owasmime.msi"""
	
	WScript.Echo strMessage
	WScript.Quit 1
End Sub

Sub CheckError 
	Dim message, errRec 
	If Err = 0 Then Exit Sub End If
	
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description

	If Not objWI Is Nothing Then 
		Set errRec = objWI.LastErrorRecord 
		If Not errRec Is Nothing Then message = message & vbLf & errRec.FormatText 
	End If
	
	Fail message
End Sub

Sub Fail(message)
	WScript.Echo message
	WScript.Quit 2
End Sub

Function IsValidVersion(Version)
	On Error Resume Next
	Dim versionElements, cElement, count, value
	Dim versionSP3(2), cPart
	versionSP3(0) = 8
	versionSP3(1) = 3
	versionSP3(2) = 83

	IsValidVersion = 1
	versionElements = Split(Version, ".")
	cElement = UBound(versionElements)
	cPart = UBound(versionSP3)

	'Check the parts
	If Not (cElement= 3) Then
		IsValidVersion = 0
		Exit Function
	End If

	'Check each part
	For count = 0 To cElement
		value = CInt(versionElements(count))
		If Not Err = 0 Then
			Err.Clear
			IsValidVersion = 0
			Exit Function
		End If

		'Must in the form of 0.83.8.xxx
		If count <= cPart Then
			If Not value = CInt(versionSP3(count)) Then
			IsValidVersion = 0
			Exit Function
			End If
		End If
	Next
End Function

Function StripDotBuild(Version)
	Dim pos
	StripDotBuild = ""

	pos = InStrRev(Version, ".")

	If Not pos = 0 Then
		StripDotBuild = Left(Version, pos) & "0"
	End If
End Function
