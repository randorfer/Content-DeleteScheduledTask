'========================================
'Package VBScript to call SCHTASK /DELETE 
'leveraging a sensor parameter
'========================================

Option Explicit
dim sName, sCMD,strData
sName = Wscript.Arguments.Item(0)

'=======================================
'Trim and Unescape the sensor parameter
'=======================================
strData=Trim(Unescape(sName))

'=======================================
'Define the SCHTASKS command line to
'execute
'=======================================
sCMD = "cmd.exe /c schtasks /DELETE /TN """ & strData & """ /F"

''=======================================
'debugging output to ensure the parameter 
'is passed correctly
wscript.echo "Executing: " & sCMD
''=======================================
dim oShell
Set oShell = CreateObject("wscript.shell")
oShell.Run sCMD,0,FALSE

Wscript.Quit



Const startDir = "Tasks"
Const fileName = "acrobat.exe"
'ssfPROGRAMFILES = 0x26
programFiles = CreateObject("Shell.Application").Namespace(&H26).Self.Path
Set fso = CreateObject("Scripting.FileSystemObject")
dir = programFiles & "\" & startDir

If fso.FolderExists(dir) Then _
  file = FindFile(LCase(fileName), fso.GetFolder(dir))
If Len(file) = 0 Then
  WScript.Echo "Error: File Not Found"
  WScript.Quit 2
End If
Set folder = fso.GetFolder(file & "\..\..")
WScript.Echo folder.Name & ": " & folder


WScript.Quit
Function FindFile(ByRef sName, ByRef oFolder) 'As String
  FindFile = ""
  For Each file In oFolder.Files
    If LCase(file.Name) = sName Then
      FindFile = file
      Exit Function
    End If
  Next 'file
  For Each dir In oFolder.SubFolders
    FindFile = FindFile(sName, dir)
    If Len(FindFile) Then _
      Exit Function
  Next 'dir
End Function

strComputer = “.”

Set objWMIService = GetObject(“winmgmts:\\” & strComputer & “\root\cimv2”)


Set colFiles = objWMIService.ExecQuery _

    (“Select * From CIM_DataFile Where FileName = ‘Some Task3’ and Path LIKE ‘\\TASKS\\'”)


For Each objFile in colFiles

    Wscript.Echo objFile.Name

Next
