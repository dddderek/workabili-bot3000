Option Explicit

Dim WshShell, fso, scriptDir, pyw, cmd, logFile

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

pyw = "C:\Users\dddde\AppData\Local\Programs\Python\Python311\pythonw.exe"
logFile = """" & scriptDir & "\output\logs\launcher.log" & """"

cmd = "cmd /c cd /d " & """" & scriptDir & """" & _
      " && " & """" & pyw & """" & " -m app.gui" & _
      " >> " & logFile & " 2>&1"

WshShell.Run cmd, 0, False

Set fso = Nothing
Set WshShell = Nothing
