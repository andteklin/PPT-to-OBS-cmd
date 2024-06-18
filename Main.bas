Attribute VB_Name = "Main"
' The event object associated with the listener class
Dim E As New EventClassModule

' The object associated with the class that parses CSV files.
Dim CSVr As New CSVReader

' Initialize event listener and generate dictionary of PowerPoint speaker notes
' mapped to values when the add-in loads.
Sub Auto_Open()

' Set the active presentation as the application for which to listen for events.
Set E.App = Application

' Set SceneMap to directory of CSV file.
Set E.SceneMap = CSVr.csvToArray("C:\Program Files\PPT-to-OBS-cmd\SceneMap.csv")

' Set file path of where the commands are stored.
E.scriptFilePath = "C:\Program Files\OBSCommand\scenes\"

End Sub
