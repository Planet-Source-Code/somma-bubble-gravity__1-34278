VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bubble-Gravity© by SOM"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   ForeColor       =   &H00404040&
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   5520
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraColours 
      BackColor       =   &H00000000&
      Caption         =   "Colours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   1695
      Left            =   3720
      TabIndex        =   20
      Top             =   2160
      Width           =   2055
      Begin VB.Label lblColour 
         BackStyle       =   0  'Transparent
         Caption         =   "Bubble Colour2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1700
      End
      Begin VB.Label lblColour 
         BackStyle       =   0  'Transparent
         Caption         =   "Bubble Colour1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblColour 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Bubble"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtGameSpeed 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   3720
      TabIndex        =   13
      Top             =   960
      Width           =   2055
      Begin VB.CheckBox chkChGrav 
         BackColor       =   &H00000000&
         Caption         =   "Change Gravity"
         ForeColor       =   &H00E2F4FC&
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   1800
      End
      Begin VB.CheckBox chkChWind 
         BackColor       =   &H00000000&
         Caption         =   "Change Wind"
         ForeColor       =   &H00E2F4FC&
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   600
         Value           =   1  'Checked
         Width           =   1800
      End
   End
   Begin VB.TextBox txtChangeTime 
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtResetTime 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "0"
      Top             =   1995
      Width           =   615
   End
   Begin VB.HScrollBar hsbBubbleSize 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.HScrollBar hsbBubbles 
      Height          =   255
      LargeChange     =   20
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Simulate every"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   2565
      Width           =   1455
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "milliseconds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   18
      Top             =   2565
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<<-- Lighter Heavier->>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   12
      Top             =   1590
      Width           =   2415
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "seconds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Gravity/Wind every"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   495
      Index           =   4
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "seconds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Reset every"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Bubble Mass"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblBubbles 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   3015
      TabIndex        =   2
      Top             =   600
      Width           =   510
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Bubbles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdApply_Click()

Call ChkSave
Call MakeBubbles

End Sub

Private Sub cmdOK_Click()

Dim IsSaved As Boolean
Call ChkSave(IsSaved)

If IsSaved Then
    Unload Me
End If

End Sub

Private Sub Form_Activate()

Dim i As Integer
Call MakeBubbles

For i = 0 To UBound(SetColours)
    Call ShowColours(i)
Next i

End Sub

Private Sub Form_Load()

Dim TempText As String
'Shows all the settings
With hsbBubbles
    .Max = MaxBubbles
    .Value = BubbleCount
End With

With hsbBubbleSize
    .Max = MassLimit
    .Min = MassMin
    .Value = MaxMass
End With

With txtChangeTime
    .Text = ChangeTime
'    TempText = "Sets how often the Auto-Screensaver will change the gravity/wind speed."
'    .ToolTipText = TempText
End With
    
With txtResetTime
    .Text = ResetTime
'    TempText = "Sets how often the Auto-Screensaver will reset." & Chr(13) & "On each reset, all the bubbles will be popped, and gravity and wind set back to 0. The Auto-Screensaver will generate the bubbles again."
'    .ToolTipText = TempText
End With
    
With txtGameSpeed
    .Text = GameSpeed
'    .ToolTipText = "Sets the speed of the simulation. The lower the setting, the faster it will run." & Chr(13) & Chr(10) & "In non-Windows NT systems, setting this to lower than 30 will have no effect."
End With

chkChGrav.Value = Abs(ChGrav)
chkChWind.Value = Abs(ChWind)

End Sub

Private Sub hsbBubbles_Change()

lblBubbles.Caption = hsbBubbles.Value + 1

End Sub

Sub MakeBubbles()

'This shows random bubbles in the settings form according to give
'a preview of bubble density and mass
Dim i As Integer
Dim ThisColour As Long
Dim NumBubbles As Integer

'Displays approx the same *density* of bubbles on the settings form
NumBubbles = Round(BubbleCount * ((Me.Width + Me.Height) / 2) / ((Screen.Width + Screen.Height) / 2))

Me.Cls

For i = 0 To NumBubbles
    DoEvents
    
    If i / 20 = Int(i / 20) Then
        Randomize
    End If
    
'    If SomRand(1, 0, , True) = 1 Then
'        ThisColour = vbBlue
'    Else
'        ThisColour = vbGreen
'    End If
    ThisColour = SetColours(SomRand(2, 1))
    
    Me.Circle (SomRand(Me.ScaleWidth, 0, , True), SomRand(Me.ScaleHeight, 0, , True)), SomRand(MaxMass, MassMin, , True) / 4, ThisColour
Next i

End Sub

Sub ChkSave(Optional ByRef Saved As Boolean)

If IsNumeric(txtResetTime.Text) And IsNumeric(txtChangeTime.Text) And IsNumeric(txtGameSpeed.Text) Then
    If txtResetTime.Text > 0 And txtChangeTime.Text > 0 And txtGameSpeed.Text > 0 Then
        Call SaveSettings
        Saved = True
    Else
        MsgBox "Enter a number more than 0", vbInformation
    End If
Else
    MsgBox "Please enter valid settings", vbInformation
End If

End Sub

Sub SaveSettings()

'Saves the settings

Dim Free As Byte
Dim i As Byte

BubbleCount = hsbBubbles.Value
ChangeTime = txtChangeTime.Text
ResetTime = txtResetTime.Text
MaxMass = hsbBubbleSize.Value
ChGrav = chkChGrav.Value
ChWind = chkChWind.Value
GameSpeed = txtGameSpeed.Text

Free = FreeFile

Open App.Path & "/" & SetName For Output As Free

Print #Free, BubbleCount
Print #Free, ChangeTime
Print #Free, ResetTime
Print #Free, MaxMass
Print #Free, Int(ChGrav)
Print #Free, Int(ChWind)
Print #Free, GameSpeed

For i = 0 To UBound(SetColours)
    Print #Free, SetColours(i)
Next i

Close Free

End Sub

Sub ShowColours(Index As Integer)

lblColour(Index).ForeColor = SetColours(Index)

End Sub

Private Sub lblColour_Click(Index As Integer)

cdlg.Color = SetColours(Index)
cdlg.ShowColor
SetColours(Index) = cdlg.Color
Call ShowColours(Index)

End Sub
