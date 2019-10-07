Set WshShell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Set logFile = fs.OpenTextFile("fileNameLogs.txt", 2, True)

strCurDir    = WshShell.CurrentDirectory

Set folder = fs.GetFolder(strCurDir)
Set files = folder.Files

For Each file in files
  currentFilePath = Lcase(file.Path)

  If (InStr(1,currentFilePath, "<INSERT UNIQUE TEXT TO SEARCH FOR>") > 0) Then call move_file(file)
Next

wscript.echo "Files moved!"
logFile.close

Sub move_file(storage_dir, file)
  Dim destination_base

  Set currentFile = CreateObject("Scripting.FileSystemObject")
  filename = file.Name
  destination_base = "C:\Users\..CHANGE TO MATCH DESIRED LOCATION"

  Select Case True
    Case InStr(1,filename, "AL") > 0
      currentFile.MoveFile file.Path, destination_base & "Alabama\"
    Case InStr(1,filename, "AK") > 0
      currentFile.MoveFile file.Path, destination_base & "Alaska\"
    Case InStr(1,filename, "AZ") > 0
      currentFile.MoveFile file.Path, destination_base & "Arizona\"
    Case InStr(1,filename, "AR") > 0
      currentFile.MoveFile file.Path, destination_base & "Arkansas\"
    Case InStr(1,filename, "CA") > 0
      currentFile.MoveFile file.Path, destination_base & "California\"
    Case InStr(1,filename, "CZ") > 0
      currentFile.MoveFile file.Path, destination_base & "Canal Zone\"
    Case InStr(1,filename, "CO") > 0
      currentFile.MoveFile file.Path, destination_base & "Colorado\"
    Case InStr(1,filename, "CT") > 0
      currentFile.MoveFile file.Path, destination_base & "Connecticut\"
    Case InStr(1,filename, "DE") > 0
      currentFile.MoveFile file.Path, destination_base & "Delaware\"
    Case InStr(1,filename, "DC") > 0
      currentFile.MoveFile file.Path, destination_base & "District of Columbia\"
    Case InStr(1,filename, "FL") > 0
      currentFile.MoveFile file.Path, destination_base & "Florida\"
    Case InStr(1,filename, "GA") > 0
      currentFile.MoveFile file.Path, destination_base & "Georgia\"
    Case InStr(1,filename, "GU") > 0
      currentFile.MoveFile file.Path, destination_base & "Guam\"
    Case InStr(1,filename, "HI") > 0
      currentFile.MoveFile file.Path, destination_base & "Hawaii\"
    Case InStr(1,filename, "ID") > 0
      currentFile.MoveFile file.Path, destination_base & "Idaho\"
    Case InStr(1,filename, "IL") > 0
      currentFile.MoveFile file.Path, destination_base & "Illinois\"
    Case InStr(1,filename, "IN") > 0
      currentFile.MoveFile file.Path, destination_base & "Indiana\"
    Case InStr(1,filename, "IA") > 0
      currentFile.MoveFile file.Path, destination_base & "Iowa\"
    Case InStr(1,filename, "KS") > 0
      currentFile.MoveFile file.Path, destination_base & "Kansas\"
    Case InStr(1,filename, "KY") > 0
      currentFile.MoveFile file.Path, destination_base & "Kentucky\"
    Case InStr(1,filename, "LA") > 0
      currentFile.MoveFile file.Path, destination_base & "Louisiana\"
    Case InStr(1,filename, "ME") > 0
      currentFile.MoveFile file.Path, destination_base & "Maine\"
    Case InStr(1,filename, "MD") > 0
      currentFile.MoveFile file.Path, destination_base & "Maryland\"
    Case InStr(1,filename, "MA") > 0
      currentFile.MoveFile file.Path, destination_base & "Massachusetts\"
    Case InStr(1,filename, "MI") > 0
      currentFile.MoveFile file.Path, destination_base & "Michigan\"
    Case InStr(1,filename, "MN") > 0
      currentFile.MoveFile file.Path, destination_base & "Minnesota\"
    Case InStr(1,filename, "MS") > 0
      currentFile.MoveFile file.Path, destination_base & "Mississippi\"
    Case InStr(1,filename, "MO") > 0
      currentFile.MoveFile file.Path, destination_base & "Missouri\"
    Case InStr(1,filename, "MT") > 0
      currentFile.MoveFile file.Path, destination_base & "Montana\"
    Case InStr(1,filename, "NE") > 0
      currentFile.MoveFile file.Path, destination_base & "Nebraska\"
    Case InStr(1,filename, "NV") > 0
      currentFile.MoveFile file.Path, destination_base & "Nevada\"
    Case InStr(1,filename, "NH") > 0
      currentFile.MoveFile file.Path, destination_base & "New Hampshire\"
    Case InStr(1,filename, "NJ") > 0
      currentFile.MoveFile file.Path, destination_base & "New Jersey\"
    Case InStr(1,filename, "NM") > 0
      currentFile.MoveFile file.Path, destination_base & "New Mexico\"
    Case InStr(1,filename, "NY") > 0
      currentFile.MoveFile file.Path, destination_base & "New York\"
    Case InStr(1,filename, "NC") > 0
      currentFile.MoveFile file.Path, destination_base & "North Carolina\"
    Case InStr(1,filename, "ND") > 0
      currentFile.MoveFile file.Path, destination_base & "North Dakota\"
    Case InStr(1,filename, "OH") > 0
      currentFile.MoveFile file.Path, destination_base & "Ohio\"
    Case InStr(1,filename, "OK") > 0
      currentFile.MoveFile file.Path, destination_base & "Oklahoma\"
    Case InStr(1,filename, "OR") > 0
      currentFile.MoveFile file.Path, destination_base & "Oregon\"
    Case InStr(1,filename, "PA") > 0
      currentFile.MoveFile file.Path, destination_base & "Pennsylvania\"
    Case InStr(1,filename, "PR") > 0
      currentFile.MoveFile file.Path, destination_base & "Puerto Rico\"
    Case InStr(1,filename, "RI") > 0
      currentFile.MoveFile file.Path, destination_base & "Rhode Island\"
    Case InStr(1,filename, "SC") > 0
      currentFile.MoveFile file.Path, destination_base & "South Carolina\"
    Case InStr(1,filename, "SD") > 0
      currentFile.MoveFile file.Path, destination_base & "South Dakota\"
    Case InStr(1,filename, "TN") > 0
      currentFile.MoveFile file.Path, destination_base & "Tennessee\"
    Case InStr(1,filename, "TX") > 0
      currentFile.MoveFile file.Path, destination_base & "Texas\"
    Case InStr(1,filename, "UT") > 0
      currentFile.MoveFile file.Path, destination_base & "Utah\"
    Case InStr(1,filename, "VT") > 0
      currentFile.MoveFile file.Path, destination_base & "Vermont\"
    Case InStr(1,filename, "VI") > 0
      currentFile.MoveFile file.Path, destination_base & "Virgin Islands\"
    Case InStr(1,filename, "VA") > 0
      currentFile.MoveFile file.Path, destination_base & "Virginia\"
    Case InStr(1,filename, "WA") > 0
      currentFile.MoveFile file.Path, destination_base & "Washington\"
    Case InStr(1,filename, "WV") > 0
      currentFile.MoveFile file.Path, destination_base & "West Virginia\"
    Case InStr(1,filename, "WI") > 0
      currentFile.MoveFile file.Path, destination_base & "Wisconsin\"
    Case InStr(1,filename, "WY") > 0
      currentFile.MoveFile file.Path, destination_base & "Wyoming\"
    Case Else
      ' DO NOTHING
  End Select
End Sub
