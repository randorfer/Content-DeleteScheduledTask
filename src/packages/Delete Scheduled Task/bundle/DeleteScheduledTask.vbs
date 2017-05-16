'========================================
'Package VBScript to call SCHTASK /DELETE 
'leveraging a sensor parameter
'========================================

Option Explicit
Const BAIL_IF_MORE_THAN_ONE = TRUE
Const TASK_DIR = "Tasks"

dim strName, strData
strName = Wscript.Arguments.Item(0)
strData=Trim(Unescape(strName))

dim objShell
Set objShell = CreateObject("wscript.shell")

'ssfSYSTEM32 = 0x25
dim strSystemRoot
strSystemRoot = CreateObject("Shell.Application").Namespace(&H25).Self.Path

dim objFSO, strDir
Set objFSO = CreateObject("Scripting.FileSystemObject")
strDir = strSystemRoot & "\" & TASK_DIR


If objFSO.FolderExists(strDir) Then
	dim strTasks
	strTasks = FindFile(strData,objFSO.GetFolder(strDir))
	If Len(strTasks) > 0 Then
		If InStr(1, strTasks, ",",1) Then
			'========================================
            'Quit the script if we find more than one 
            'scheduled task with the same name
            'leverages the BAIL_IF_MORE_THAN_ONE 
            'constant variable
            '========================================

            
            If BAIL_IF_MORE_THAN_ONE Then
				wscript.echo "Found multiple tasks with name of """ & strData & """ - bailing" 
				wscript.quit
			Else
				dim arrTask
				arrTask = Split(strTasks,",")
				dim i
				For i=0 To UBound(arrTask)
					'wscript.echo arrTask(i)  & " | " & TrimForSchtasks(arrTask(i))
					DeleteTask (TrimForSchtasks(arrTask(i)))
				Next
			End If
		Else
			'wscript.echo strTasks & " | " & TrimForSchtasks(strTasks)
			DeleteTask (TrimForSchtasks(strTasks))
		End If
	End If	
End If


WScript.Quit

Sub DeleteTask (strTaskName)

	dim strCMD
	strCMD = "cmd.exe /c schtasks /DELETE /TN """ & strTaskName & """ /F"

	wscript.echo "Executing: " & strCMD
	objShell.Run strCMD,0,FALSE	

End Sub


Function TrimForSchtasks(strJob)

	Dim strPath
	strPath = objFSO.GetAbsolutePathName(strJob)
	
	dim arrPath
	arrPath = Split(strPath,"Tasks")
	
	TrimForSchtasks = arrPath(Ubound(arrPath))

End Function


Function FindFile(strName,objFolder) 'As String
  dim strRet : strRet = ""
  
  FindFile = ""
  dim objFile
  For Each objfile In objFolder.Files
    If LCase(objfile.Name) = strName Then
      strRet = LCase(objfile.Path)
    End If
  Next 'file
  
  dim objDir
  For Each objDir In objFolder.SubFolders
	dim strFuncRet : strFuncRet = ""
	strFuncRet = FindFile(strName, objDir)
    If Len(strFuncRet) Then
		If strRet = "" Then
			strRet = strFuncRet
		Else
			strRet = strRet & "," & strFuncRet
		End If
	End If
  Next 'dir
  
  FindFile = strRet
  
End Function