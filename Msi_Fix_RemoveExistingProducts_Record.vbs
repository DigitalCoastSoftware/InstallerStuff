Option Explicit
'
' Purpose: 
'  Work-around for a Microsoft bug in Visual Studio Installer projects plugin.
'  Incorrect msi file gets generated when you set RemovePreviousVersion=True in 
'  the installer project properties. This script can fix it.
'
' Usage: Call it with a msi file argument.
' Example:
'  cscript //nologo "Msi_Fix_RemoveExistingProducts_Record.vbs" "SomeCoolInstaller.msi"
'
' When to use it:
'  If you want your installer to automatically uninstall older versions of your product.
'  In other words, when you want your new version to replace an existing install.
'  (BTW When creating an installer for a new revision of your product, change the 
'  ProductCode and PackageCode guids but _keep_ the same UpgradeCode.)
'
' The Microsoft bug:
'  Microsoft's plugin for Visual Studio 2017 Installer projects has a bug that places the 
'  uninstall operation too late in the sequence. First it installs the new files, then 
'  invokes uninstall for the previous version. This can cause new files to get incorrectly 
'  deleted if they have the same name as old files.
'  The fix: Do uninstall of an older version _before_ installing new items.
'  This function moves the uninstall operation right after InstallValidate, and just before 
'  InstallInitialize.
'
' Tip: 
'   You can add this as a post build operation in your Visual Studio Installer project. 
'   I typically put the call in a bat file and then call the script from there. 
'   For example: Select the installer project in solution explorer and view properties (F4)
'   Set these two props:
'     PostBuildEvent:    "$(ProjectDir)installer_postbuild.bat" "$(BuiltOuputPath)"
'     RunPostBuildEvent: On successful build
'   
' Additional information:
'  I hacked this together by looking at some of the sample scripts provided by Microsoft.
'  I recommend checking that stuff out. See Windows SDK example scripts...
'   https://docs.microsoft.com/en-us/windows/desktop/msi/windows-installer-scripting-examples
'
'

Const msiOpenDatabaseModeReadOnly = 0
Const msiOpenDatabaseModeTransact = 1

If (Wscript.Arguments.Count < 1) Then
	Wscript.Echo "usage: cscript thisfile.vbs filename.msi"
	Wscript.Quit 1
End If

TheMainThing Wscript.Arguments(0)

'  This function takes an msi filename and updates the 'RemoveExistingProducts' record in 
'  the InstallExecuteSequence table.
'
Sub TheMainThing(msiFile)
    On Error Resume Next
    If Len(msiFile) = 0 Then Exit Sub

    Dim installer : Set installer = Nothing
    Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError
    Dim database : Set database = installer.OpenDatabase(msiFile, msiOpenDatabaseModeTransact) : CheckError

    '  Get the sequence number for InstallValidate and InstallInitialize
    Dim iValidate: iValidate = GetSequenceId(database, "InstallValidate")
    If iValidate < 0 Then Fail "InstallValidate"

    Dim iInitialize: iInitialize = GetSequenceId(database, "InstallInitialize")
    If iInitialize < 0 Then Fail "InstallInitialize"

    '  Now get the midpoint between the two sequence ids and use that for
    '  the 'RemoveExistingProducts' sequence number.
    Dim iMidpoint: iMidpoint = CLng((iValidate + iInitialize) / 2)

    If Not UpdateSequenceId(database, "RemoveExistingProducts", iMidpoint) Then Fail "Failed to update sequence id of RemoveExistingProducts"
    database.Commit
    Set database = Nothing
    Wscript.Echo "msi database updated: " & msifile
End Sub

Function UpdateSequenceId(db, name, value)
    UpdateSequenceId = False
    If db is Nothing Then Exit function
    If Len(name) = 0 Then Exit function
    Dim query1, query2, query3
    Dim query
    ' "UPDATE `InstallExecuteSequence` SET `InstallExecuteSequence`.`Sequence`=1450 WHERE `InstallExecuteSequence`.`Action`='RemoveExistingProducts'"
    query1 = "UPDATE `InstallExecuteSequence` SET `InstallExecuteSequence`.`Sequence`=" 'number
    query2 = " WHERE `InstallExecuteSequence`.`Action`='"
    query3 = "'"
    query = query1 & CStr(value) & query2 & name & query3
    Dim view
    Set view = db.OpenView(query) : CheckError
	view.Execute : CheckError
    UpdateSequenceId = True
End Function

Function GetSequenceId(db, name)
    GetSequenceId = -1
    If db is Nothing Then Exit function
    If Len(name) = 0 Then Exit function
    'query = "SELECT `InstallExecuteSequence`.`Sequence` FROM `InstallExecuteSequence` WHERE `InstallExecuteSequence`.`Action`='InstallValidate'"
    Dim query1
    Dim query2
    Dim query
    query1 = "SELECT `InstallExecuteSequence`.`Sequence` FROM `InstallExecuteSequence` WHERE `InstallExecuteSequence`.`Action`='"
    query2 = "'"
    query = query1 & name & query2
    Dim view
    Set view = db.OpenView(query) : CheckError
	view.Execute : CheckError
    Dim record
	Set record = view.Fetch : CheckError
	Set view = Nothing
	If record Is Nothing Then Fail "no record: " & name
    Dim value
    value = CLng(record.StringData(1))
    GetSequenceId = value
End Function

Sub CheckError()
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbLf & errRec.FormatText
	End If
	Fail message
End Sub

Sub Fail(message)
	Wscript.Echo message
	Wscript.Quit 2
End Sub

