dim fs,objTextFile
set fs=CreateObject("Scripting.FileSystemObject")
dim arrStr
set objTextFile = fs.OpenTextFile("temp.csv")


'Create new text file in csv format
Const ForWriting = 2
Set objCSVFile = fs.OpenTextFile("result.csv", ForWriting, True)
csvColumns = "Col1,Col2,Col3"
objCSVFile.Write csvColumns
objCSVFile.Writeline

' Write test values as comma-separated in new CSV file.
' For i = 0 to 5 
'    objCSVFile.Write chr(34) & "Value1" & chr(34) & ","
'    objCSVFile.Write chr(34) & "Value2" & chr(34) & ","
'    objCSVFile.Write chr(34) & "Value23" & chr(34) & ""
'    objCSVFile.Writeline
' Next

lineNumber = 1
Do while NOT objTextFile.AtEndOfStream
  arrStr = split(objTextFile.ReadLine,",")

' arrStr is now an array that has each of your fields
' process them, whatever.....
If lineNumber = 1 Then
  lineNumber = lineNumber + 1  
Else
  objCSVFile.Write arrStr(0) & ","
  objCSVFile.Write arrStr(1) & ","
  objCSVFile.Write arrStr(2) & ","
  objCSVFile.Writeline
  lineNumber = lineNumber + 1
End If
Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing