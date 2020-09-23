VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsMassiveObject"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim XPos As Single
Dim YPos As Single
Dim XVel As Single
Dim YVel As Single

Property Let X(Value As Single)

XPos = Value

End Property

Property Get X() As Single

X = XPos

End Property

Property Let Y(Value As Single)

YPos = Value

End Property

Property Get Y() As Single

Y = YPos

End Property

Property Let XSpeed(Value As Single)

XVel = Value

End Property

Property Get XSpeed() As Single

XSpeed = XVel

End Property

Property Let YSpeed(Value As Single)

YVel = Value

End Property

Property Get YSpeed() As Single

YSpeed = YVel

End Property

Sub Accelerate(ByVal Degrees As Single, ByVal Amount As Single)

XVel = XVel + (Amount * Cos(DegToRad(Degrees)))
YVel = YVel + (Amount * Sin(DegToRad(Degrees)))

End Sub

Sub MoveIt(ByVal FormWidth As Single, ByVal FormHeight As Single)

XPos = XPos + XVel
YPos = YPos + YVel

'Collision detection with walls. Reflects bubbles
'As far as I can see, this should result in a perfect reflection,
'BUT IT DOESN'T. Why?
'No, of course it doesn't! Gravity is pulling it in the same direction in the
'unit of time, when the object actually reflects during that unit of time.
If XPos < 0 Then
    XPos = -XPos
    XVel = -XVel
End If

If XPos > FormWidth Then
    XPos = FormWidth - (XPos - FormWidth)
    XVel = -XVel
End If

If YPos < 0 Then
    YPos = -YPos
    YVel = -YVel
End If

If YPos > FormHeight Then
    YPos = FormHeight - (YPos - FormHeight)
    YVel = -YVel
End If

End Sub
