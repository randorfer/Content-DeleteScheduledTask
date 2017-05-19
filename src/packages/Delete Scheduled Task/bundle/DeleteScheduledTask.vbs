'========================================
'Package VBScript to call SCHTASK /DELETE 
'leveraging a sensor parameter
'========================================
OPTION EXPLICIT

'Controls logic on how we handle duplicate
'task names
const SKIP_DUPS = TRUE

dim strName, strData
strName = Wscript.Arguments.Item(0)
strData=UCASE(Trim(Unescape(strName)))
wscript.echo "Scheduled TaskName passed: " &strData

dim objShell
Set objShell = CreateObject("WScript.Shell")

dim objExec
Set objExec = objShell.Exec("schtasks /query /FO list")

dim dictTasks
Set dictTasks = CreateObject("Scripting.Dictionary")

dim strLine
Do Until objExec.StdOut.AtEndOfStream
 strLine = UCASE(objExec.StdOut.ReadLine)
 If InStr(strLine, strData) Then
    
    strLine = trim(replace(strLine,"TASKNAME:",""))

    If dictTasks.Exists(strLine) Then
        'wscript.echo strLine & " is already in the dictionary"
    Else
        'wscript.echo strLine & " will be added to dictionary"
        dictTasks.Add strLine, strLine
    End If
 End If
Loop


dim intCount : intCount = 0

dim colKeys,strKey
colKeys = dictTasks.Keys
For Each strKey in colKeys
    If InStr(strKey, strData) Then
	intCount = intCount + 1
    End If
Next

If intCount > 1 and SKIP_DUPS Then
	wscript.echo "*** Manual remediation is required ***"
	wscript.echo
	wscript.echo strData & " was found " & intCount & " time(s)."
	wscript.echo "Alternatively, You can modify the package script to allow duplicate deletion."
Else
    For Each strKey in colKeys
       If InStr(strKey, strData) Then
	  wscript.echo "Deleting " & strKey
	  DeleteTask(strKey)
       End If
    Next
End If

Sub DeleteTask (strTaskName)

	dim strCMD
	strCMD = "cmd.exe /c schtasks /DELETE /TN """ & strTaskName & """ /F"

	wscript.echo "Executing: " & strCMD
	objShell.Run strCMD,0,FALSE	

End Sub
