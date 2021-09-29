set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
directory = objShell.CurrentDirectory
MsgBox("Terminate bot")
objFile.CreateTextFile(directory & "\terminate.txt", true).WriteLine("1")
call objShell.run("taskkill /f /im wscript.exe",,True)
