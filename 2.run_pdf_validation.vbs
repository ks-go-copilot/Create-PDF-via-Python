Option Explicit

Dim shell
Dim fso
Dim folder

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

folder = fso.GetParentFolderName(WScript.ScriptFullName)

shell.CurrentDirectory = folder

shell.Run "py pdf_checker.py", 0, True

MsgBox "PDF QA Finished!", vbInformation