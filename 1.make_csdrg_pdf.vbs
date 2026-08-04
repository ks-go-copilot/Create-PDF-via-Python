Set objShell = CreateObject("WScript.Shell")
objShell.Run "python.exe """ & CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\make_csdrg_pdf.py""", 1, True
