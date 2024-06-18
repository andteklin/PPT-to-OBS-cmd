VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class represents the events and methods associated with the active presentation.
' When an object of this class is instantiated, the event listeners are
' activated.

' Event listener for the PowerPoint application
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' A dictionary that maps keys (PowerPoint speaker notes as String)
' to values (Strings of script files used to run OBS commands)
Public SceneMap As Object

' File path of the OBSCommand scripts
Public scriptFilePath As String

' Event handler which is triggered immediately after moving to the next slide.
' Finds if there is any speaker note in the next slide that matches a specified key in SceneMap.
' If so, the corresponding script is executed which changes the scene in OBS
' to the one that corresponds with the speaker note. Otherwise, does nothing.
Private Sub App_SlideShowNextSlide(ByVal Wn As SlideShowWindow)
    On Error Resume Next
    Dim showPos As Integer
    Dim sNotes As String
    Dim sKey As String
    
    showPos = Wn.View.CurrentShowPosition
    sNotes = getSpeakerNotes(showPos)
    sKey = getFirstLine(sNotes)
    
    If SceneMap.Exists(sKey) Then
        SendCmd (SceneMap.Item(sKey))
    End If
    
    If showPos = 1 Then
        AppActivate "PowerPoint Presenter View"
    End If
    
End Sub

' Takes in the current show position and returns the speaker notes for that slide.
' Invariant: position must be in range of the number of slides of the active presentation.
Private Function getSpeakerNotes(position As Integer) As String
    Dim sl As Slide
    Dim notesText As String
    Set sl = ActivePresentation.Slides(position)
    notesText = NotesTextPlaceholder(sl).TextFrame.TextRange.Text
    getSpeakerNotes = notesText
End Function

' Helper function to identify the notes text placeholder if present on the notes page.
' Returns a reference to the notes text placeholder on the slide.
' Original code from https://www.rdpslides.com/pptfaq/FAQ00963_Show_me_the_Notes_Text.htm
Private Function NotesTextPlaceholder(oSl As Slide) As shape

Dim osh As shape

For Each osh In oSl.NotesPage.Shapes

    If osh.Type = msoPlaceholder Then
        If osh.PlaceholderFormat.Type = ppPlaceholderBody Then
            ' we found it
            Set NotesTextPlaceholder = osh
            Exit Function
        End If
    End If

Next

End Function

' Helper function that takes in the text from the speaker note shape
' and returns the first line of text.
' Invariant: the first line must have a string that is contained
' as a PPT speaker note key in SceneMap.
Private Function getFirstLine(sNotes As String) As String
sNotes = Split(sNotes, vbCr)(0)
sNotes = LTrim(sNotes)
sNotes = RTrim(sNotes)
sNotes = UCase(sNotes)
getFirstLine = sNotes
End Function

' Calls the script file in C:\ProgramFiles\OBSCommand\scenes that matches the string Scene. The script file changes the scene in an active session of OBS.
' Precondition: Scene must be a valid string matching the list of script file names in C:\ProgramFiles\OBSCommand\scenes\.
Private Sub SendCmd(Scene As String)
    On Error Resume Next
    DoEvents
    Run = Shell(scriptFilePath & Scene & ".bat", 6)
    AppActivate "PowerPoint"
End Sub

' Prints the SceneMap in CSV format to the debugger.
Public Sub printSceneMap()
    For Each Key In SceneMap.Keys
        Debug.Print Key + "," + SceneMap.Item(Key)
    Next Key
End Sub
