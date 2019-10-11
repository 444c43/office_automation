Option Explicit

Dim strCurDir
Dim fs, logfile, folder, file

strCurDir = CreateObject("WScript.Shell").CurrentDirectory

Set fs = CreateObject("Scripting.FileSystemObject")
Set logFile = fs.OpenTextFile("fileNameLogs.txt", 2, True)
Set folder = fs.GetFolder(strCurDir)

For Each file in folder.Files
  logFile.writeline(file.name)
Next

logFile.close
wscript.echo "File names written!"
