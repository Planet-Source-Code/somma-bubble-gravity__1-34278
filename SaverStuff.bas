Attribute VB_Name = "SaverStuff"
Option Explicit

'Name of the settings file. Used in frmsettings as well
Public Const SetName = "bubblegrav.som"

Sub Main()

'Stops windows from opening the screensaver multiple times
If App.PrevInstance Then
    End
End If

'Reads any parameters given by windows
Dim Param As String
Param = Left(Command, 2)

Call ReadSettings

If Param <> "" Then
    SaverMode = True
    Select Case Param
        Case "/p"
            End
        Case "/c"
            frmSettings.Show
        Case "/s"
            frmGravity.Show
    End Select
Else
    SaverMode = False
    frmGravity.Show
End If


End Sub

Public Sub ReadSettings()

'Reads settings from file
Dim Free As Byte
Dim ChkFile As String
Dim Temp As String
Dim i As Byte

ChkFile = App.Path & "/" & SetName

If FileExists(ChkFile) And SetName <> "" Then
    
    Free = FreeFile
    Open ChkFile For Input As Free
    
    Input #Free, BubbleCount
    Input #Free, ChangeTime
    Input #Free, ResetTime
    Input #Free, MaxMass
    
    Input #Free, Temp
    ChGrav = Int(Temp)
    Input #Free, Temp
    ChWind = Int(Temp)
    Input #Free, GameSpeed
    
    For i = 0 To UBound(SetColours)
        Input #Free, SetColours(i)
    Next i
    
    Close #Free
    
Else
    Call SetDefaults
End If
    
End Sub

Sub SetDefaults()

'Gives default settings if settings file not found
BubbleCount = 99
ChangeTime = 5
ResetTime = 60
MaxMass = 60
ChGrav = True
ChWind = True
GameSpeed = 40
SetColours(0) = vbRed
SetColours(1) = vbWhite
SetColours(2) = &H808000

End Sub

Public Function FileExists(FileName As String) As Boolean
'Checks if a file exists. There *has* to be
'a better way of doing this...
If FileName <> "" Then
    Dim CheckThis As String
    On Error Resume Next
    
    CheckThis = Dir(FileName)
    If CheckThis = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End If
End Function

Public Function SomRand(ByVal High As Double, ByVal Low As Double, Optional DecPlaces As Integer = 0, Optional NoRandomize As Boolean) As Single

'Makes generating random numbers easier
If Not NoRandomize Then
    Randomize
End If

SomRand = Round(((High - Low) * Rnd) + Low, DecPlaces)

End Function

