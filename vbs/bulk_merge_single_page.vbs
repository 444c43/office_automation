Option Explicit

Dim OriginDirectory, DestinationDirectory
Dim FileToMerge
Dim DocumentsPath
Dim WshShell, fs
Dim folder, file, files, currentFilePath

Call AssignVariables()

For Each file in files
  currentFilePath = Lcase(file.Path)

  If (InStr(1,currentFilePath, "proposal") > 0) Then call Proposal(file)
Next

MsgBox "Merge Complete!"



Private Sub AssignVariables
  Set WshShell = CreateObject("WScript.Shell")

  DocumentsPath = "C:\Users\" & GetCurrentUser & "\Documents\"

  OriginDirectory = SelectFolder( DocumentsPath, "Please select origin directory" )
  FileToMerge = SelectFile()
  DestinationDirectory = SelectFolder( DocumentsPath, "Please select destination directory" )

  Set fs = CreateObject("Scripting.FileSystemObject")
  Set folder = fs.GetFolder(OriginDirectory)
  Set files = folder.Files
End Sub

Private Sub Proposal(file)
  Dim oMainDoc, oTempDoc

  Set oMainDoc = CreateObject("AcroExch.PDDoc")
  oMainDoc.Open file.Path

  Set oTempDoc = CreateObject("AcroExch.PDDoc")
  oTempDoc.Open FileToMerge

  oMainDoc.InsertPages -1, oTempDoc, 0, oTempDoc.GetNumPages, False
  oMainDoc.Save 1, DestinationDirectory & "\" & file.Name & ".pdf"
  oTempDoc.Close
  oMainDoc.Close
End Sub

Private Function SelectFolder( myStartFolder, dialogText )
    ' Standard housekeeping
    Dim objFolder, objItem, objShell
    
    ' Custom error handling
    On Error Resume Next
    SelectFolder = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, dialogText, 0, myStartFolder )

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function

Private Function GetCurrentUser()
  GetCurrentUser = CreateObject("WScript.Network").UserName
End Function

Private Function SelectFile()
  Dim wShell, oExec

  Set wShell=CreateObject("WScript.Shell")
  Set oExec = wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
  SelectFile = oExec.StdOut.ReadLine
End Function
