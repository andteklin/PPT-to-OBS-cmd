VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Opens and reads a CSV file. Each row is a key-value pair, mapping
' PPT speaker notes to file names of scripts that run OBS commands.
' The default file path is the Documents folder.
' Returns a Map of possible PPT speaker notes mapped to file names of OBSCommand scripts.
Public Function csvToArray(filePath As String) As Object

Dim Map As Object
Set Map = CreateObject("Scripting.Dictionary")

Open filePath For Input As #1

' Loops through all lines in file.
' Invariant: each element of column 1 is mapped to an element of column 2.
Do Until EOF(1)
    Line Input #1, lineText
    LineArr = Split(lineText, ",")
    Map.Add LineArr(0), LineArr(1)
Loop

Close #1

Set csvToArray = Map

End Function

