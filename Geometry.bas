Attribute VB_Name = "Geometry"
Option Explicit

Public Const PI = 3.1415927

Public Function DegToRad(Degrees As Single) As Single

DegToRad = Degrees * (PI / 180)

End Function

Public Function RadToDeg(Radians As Single) As Single

RadToDeg = (Radians * 180) / PI

End Function

Public Function ConvY(ByVal Value As Single, WHeight As Single) As Single

ConvY = WHeight - Value

End Function

'Public Function FindAngle(X As Single, Y As Single) As Single
'
'If X <> 0 Then
'    FindAngle = RadToDeg(Tan(Y / X))
'ElseIf Y > 0 Then
'    FindAngle = 180
'ElseIf Y < 0 Then
'    FindAngle = 90
'Else
'    FindAngle = 0
'End If
'
'End Function
