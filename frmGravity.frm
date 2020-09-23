VERSION 5.00
Begin VB.Form frmGravity 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Bubble-Gravity©"
   ClientHeight    =   8475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmGravity.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label lblSaverMode 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto-Screensaver"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   5655
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label lblWindX 
      BackStyle       =   0  'Transparent
      Caption         =   "WindX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblWindY 
      BackStyle       =   0  'Transparent
      Caption         =   "WindY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblGravY 
      BackStyle       =   0  'Transparent
      Caption         =   "GravityY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblForce 
      BackStyle       =   0  'Transparent
      Caption         =   "Force"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblMass 
      BackStyle       =   0  'Transparent
      Caption         =   "Mass"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblGravX 
      BackStyle       =   0  'Transparent
      Caption         =   "GravityX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblBubbles 
      BackStyle       =   0  'Transparent
      Caption         =   "Bubbles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblfps 
      BackStyle       =   0  'Transparent
      Caption         =   "fps"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "[Bubble-Gravity©: by SOM] Click to exit. F1 for Help"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "frmGravity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'Private Enum Sets
'    GRAVY = 0
'    GRAVX = 1
'    GAMESPEED = 2
'End Enum
'
'Dim MainSets() As String
Const LabelShowTime = 10
Dim GravX As Single
Dim GravY As Single
Dim WindX As Single
Dim WindY As Single

Dim CurObj As Integer
Dim Thrust As Single

Dim Objects() As New clsMassiveObject
Dim Colours() As Long

Dim GameOver As Boolean

Dim ShowlblTime As Integer
Dim LastlblShow As Single

Dim Activated As Boolean

Private Sub Form_Activate()

If Not Activated Then

    ReDim Colours(0)
    
    Activated = True
    Call BigReset
    
    'Puts stuff into the help label
    AddHelp "Use [Arrow Keys] to apply thrust"
    AddHelp ""
    AddHelp "Press F2 to toggle Auto-Screensaver mode"
    AddHelp ""
    AddHelp "[Escape] or any [Mouse-Click] exits."
    AddHelp ""
    AddHelp "[+] Add Bubble"
    AddHelp "[-] Pop Bubble"
    AddHelp "[Home] Increase Thrust"
    AddHelp "[End] Decrease Thrust"
    AddHelp "[Insert] Increase Bubble Mass"
    AddHelp "[Delete] Decrease Bubble Mass"
    AddHelp "[Page Up] Select Next Bubble"
    AddHelp "[Page Down] Select Prev Bubble"
    AddHelp ""
    AddHelp "With Numpad:"
    AddHelp "[1]/[0] Increase Gravity Upwards/Downwards"
    AddHelp "[2]/[3] Increase Gravity Left/Right"
    AddHelp "[7]/[4] Increase Wind Upwards/Downwards"
    AddHelp "[8]/[9] Increase Wind Left/Right"
    AddHelp ""
    AddHelp "Remember that thrust and wind affect lighter objects more than heavier ones."
    
    GravY = -0.5
    Thrust = 50
    
    'This will be part of the screen saver settings when implemented
    'ChangeTime = 5
    'BubbleCount = 149
    'SaverMode = True
    'ResetTime = 60
    'ChWind = True
    
    Call MainLoop
    
End If

End Sub

Sub AddHelp(What As String)

lblHelp.Caption = lblHelp.Caption & What & Chr(13)

End Sub

Sub ResetObj(Index As Integer)

If Index > UBound(Objects) Then
    ReDim Preserve Objects(Index)
End If

With Objects(Index)

    .Pos.X = Me.ScaleWidth / 2
    .Pos.Y = Me.ScaleHeight / 2
    
    .Pos.YSpeed = SomRand(15, -15, 1)
    .Pos.XSpeed = SomRand(15, -15, 1)
    .Mass = SomRand(MaxMass, MassMin)
    
End With


ReDim Preserve Colours(UBound(Objects))

'Set a random colour
'If SomRand(1, 0) = 1 Then 'Int(2 * Rnd) = 1 Then
'    Colours(Index) = vbBlue
'Else
'    Colours(Index) = vbGreen
'End If

Colours(Index) = SetColours(SomRand(2, 1))

End Sub

Sub BigReset()

Dim i As Integer
ReDim Objects(0)
Call ResetObj(0)

GravX = 0
GravY = 0
WindX = 0
WindY = 0
'For i = 0 To UBound(Objects)
'    Call ResetObj(i)
'Next i


End Sub

Sub MainLoop()

Const MassIncr = 0.5
Const ForceIncr = 0.5
Const GravIncr = 0.1
Const RanRange = 0.6
Const RanWindRange = 10
Const MaxGrav = 1.5
Const SaverName = "Auto-ScreenSaver"

Dim i As Integer
Dim ObjAmount As Integer
Dim StartTime As Single
Dim fpsCount As Integer
Dim fpsStart As Single
Dim UseThisCol As Long

Dim RanChangeTime As Integer
Dim LastChange As Single
Dim LastReset As Single

'This is where everything happens
Do

    'Limits the frame-rate
    If Timer - StartTime >= GameSpeed / 1000 Then
        
        'Count framerate
        fpsCount = fpsCount + 1
        If Timer - fpsStart >= 1 Then
            lblfps.Caption = fpsCount & "fps"
            fpsStart = Timer
            fpsCount = 0
        End If
        
        StartTime = Timer
        
        'Hides the labels after a few seconds of inactivity
        
        Call HideLabels(Timer - LastlblShow < LabelShowTime)

        'Auto-screensaver procedures
        If SaverMode Then
                    
            If Timer - LastChange > RanChangeTime Then
                RanChangeTime = ChangeTime 'Int(((2 * ChangeTime) / 4) * Rnd) + ChangeTime / 4
                LastChange = Timer
                
                'Randomly increases/decreases gravity
                If ChGrav Then
                    If SomRand(1, 0) = 1 Then
                        GravX = Round(GravX + SomRand(RanRange, -RanRange, 1), 1)
                    Else
                        GravY = Round(GravY + SomRand(RanRange, -RanRange, 1), 1)
                    End If
                    
                    If Abs(GravY) >= MaxGrav Then
                        GravY = 0
                    End If
                    
                    If Abs(GravX) >= MaxGrav Then
                        GravX = 0
                    End If
                End If
                
                'Randomly increases/decreases wind
                If ChWind Then
                    If SomRand(1, 0) = 1 Then
                        WindX = RoundClosest(WindX + SomRand(RanWindRange, -RanWindRange, 1), ForceIncr)
                    Else
                        WindY = RoundClosest(WindY + SomRand(RanWindRange, -RanWindRange, 1), ForceIncr)
                    End If
                    
                End If
                    
                
            End If
            
            'Checks if its time to reset
            If Timer - LastReset > ResetTime Then
                LastReset = Timer
                Call BigReset
            End If
            
            'Creates bubbles until bubblecount is reached
            If UBound(Objects) < BubbleCount Then
                Call ResetObj(UBound(Objects) + 1)
            End If
        
        End If
            
        DoEvents
        
        Me.Cls
        
        'Loop through each bubble
        For i = 0 To UBound(Objects)
        
            'Makes the selected bubble red
            If i <> CurObj Then
                UseThisCol = Colours(i)
            Else
                UseThisCol = SetColours(0)
            End If
            
            
            'Force due to wind. Related to mass
            With Objects(i)
                Call .Force(WindY, 270)
                Call .Force(WindX, 0)
            End With
            
            'Acceleration due to gravity. Unrelated to mass
            With Objects(i).Pos
                Call .Accelerate(270, GravY)
                Call .Accelerate(0, GravX)
                Call .MoveIt(Me.ScaleWidth, Me.ScaleHeight)
                
                'Draws the bubble
                Me.Circle (.X, ConvY(.Y, Me.ScaleHeight)), (Objects(i).Mass / 4), UseThisCol
            End With
            
        Next i
        
        ObjAmount = UBound(Objects)
        
        'The minus key
        If KeyPressed(109) Then
            If ObjAmount > 0 Then
                'Decrease size of array
                ReDim Preserve Objects(ObjAmount - 1)
            End If
        End If
        
        'The plus key
        If KeyPressed(107) Then
            If ObjAmount < MaxBubbles Then
                
                'Increases size of array
                ReDim Preserve Objects(ObjAmount + 1)
                Call ResetObj(UBound(Objects))
                
            End If
        End If
        
        'Cycle through bubbles
        If KeyPressed(33) Then
            CurObj = CurObj + 1
        End If
        
        If KeyPressed(34) Then
            CurObj = CurObj - 1
        End If
        
        'Checks to see if curobj is within the number of bubbles
        Call CheckBound
        
        'Change the thrust with home and end
        If KeyPressed(36) Then
            Thrust = Thrust + ForceIncr
        End If
        
        If KeyPressed(35) Then
            If Thrust > 0 Then
                Thrust = Thrust - ForceIncr
            End If
        End If
        
        'Change gravity
        If KeyPressed(97) Then
            GravY = Round(GravY - GravIncr, 1)
        End If
        
        If KeyPressed(96) Then
            GravY = Round(GravY + GravIncr, 1)
        End If
        
        If KeyPressed(98) Then
            GravX = Round(GravX - GravIncr, 1)
        End If
        
        If KeyPressed(99) Then
            GravX = Round(GravX + GravIncr, 1)
        End If
        
        'Change Wind
        'Numpad 7 and 4 to change vertical wind
        If KeyPressed(103) Then
            WindY = Round(WindY - ForceIncr, 1)
        End If
        
        If KeyPressed(100) Then
            WindY = Round(WindY + ForceIncr, 1)
        End If
        
        'Numpad 8 and 9 to change hor wind
        If KeyPressed(104) Then
            WindX = Round(WindX - ForceIncr, 1)
        End If
        
        If KeyPressed(105) Then
            WindX = Round(WindX + ForceIncr, 1)
        End If
            
        
        With Objects(CurObj)
            'Gives thrust to objects according to which key is pressed
            If KeyPressed(vbKeyUp) Then
                Call .Force(Thrust, 90)
            End If
            
            If KeyPressed(vbKeyDown) Then
                Call .Force(Thrust, 270)
            End If
            
            If KeyPressed(vbKeyLeft) Then
                Call .Force(Thrust, 180)
            End If
            
            If KeyPressed(vbKeyRight) Then
                Call .Force(Thrust, 0)
            End If
            
            'Change mass of object with ins and del
            If KeyPressed(45) Then
                .Mass = .Mass + MassIncr
            End If
            
            If KeyPressed(46) Then
                If .Mass > MassMin Then
                    .Mass = .Mass - MassIncr
                End If
            End If
            
                        
            'Displays stuff
            lblMass.Caption = "Mass: " & .Mass
            lblGravX.Caption = "X Gravity: " & GravX
            lblGravY.Caption = "Y Gravity: " & GravY
            'lblGravSum.Caption = Round(FindAngle(GravX, GravY), 1) & "º, " & Round(Sqr(GravX ^ 2 + GravY ^ 2), 1)
            
            lblWindX.Caption = "X Wind: " & WindX
            lblWindY.Caption = "Y Wind: " & WindY
            'lblWindSum.Caption = Round(FindAngle(WindX, WindY), 1) & "º, " & Round(Sqr(WindX ^ 2 + WindY ^ 2), 1)
            
            lblForce.Caption = "Thrust: " & Thrust
            lblBubbles.Caption = UBound(Objects) + 1 & " Bubble(s)"
                        
        End With
        
        'The escape key
        If KeyPressed(27) Then
            Unload Me
        End If
        
        'The F1 key for help
        lblHelp.Visible = KeyPressed(112)
        
        'The F2 key to toggle auto-screensaver
        If KeyPressed(113) Then
            SaverMode = Not SaverMode
        End If
        
        'Shows whether auto-screensaver is on
        If SaverMode Then
            lblSaverMode.Caption = SaverName & " On"
        Else
            lblSaverMode.Caption = SaverName & " Off"
        End If

    End If
        
Loop Until GameOver
        
End Sub

Function KeyPressed(KeyCode As Long) As Boolean

'Checks if a key is pressed
If GetKeyState(KeyCode) < -125 Then
    KeyPressed = True
End If

End Function

Sub CheckBound()

Dim Max As Integer
Max = UBound(Objects)

'Allows looping around when selecting
If CurObj > Max Then
    CurObj = CurObj - Max - 1
ElseIf CurObj < 0 Then
    CurObj = CurObj + Max + 1
End If

'In case curobj is still outside boundaries
If CurObj < 0 Then
    CurObj = 0
ElseIf CurObj > Max Then
    CurObj = Max
End If

End Sub

Sub HideLabels(Optional Show As Boolean)

'Hides/shows the labels
'WHY CAN'T WE HAVE TRANSPARENT CONTAINERS?!?! Damn you VB!

lblGravX.Visible = Show
lblGravY.Visible = Show
lblMass.Visible = Show
lblForce.Visible = Show
lblBubbles.Visible = Show
lblfps.Visible = Show
lblTitle.Visible = Show
lblWindX.Visible = Show
lblWindY.Visible = Show
lblSaverMode.Visible = Show

End Sub


Public Function RoundClosest(ByVal Number As Double, ByVal ClosestNum As Single) As Single

'Rounds number to closest whatever.
'e.g round to closest 360 for some angle thing

Number = Number / ClosestNum
Number = Round(Number)
RoundClosest = Number * ClosestNum

End Function


Private Sub Form_Click()

Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

ShowlblTime = LabelShowTime
LastlblShow = Timer

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

ShowlblTime = LabelShowTime
LastlblShow = Timer

End Sub

Private Sub Form_Unload(Cancel As Integer)

GameOver = True
End

End Sub

