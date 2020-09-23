VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMassiveObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'How does inheritence with classes work in VB? Does it work at all?
'Somebody help me!!!

Dim cobj As New clsObject
Dim TheMass As Single

Property Let Mass(ByVal Value As Single)

TheMass = Value

End Property

Property Get Mass() As Single

Mass = TheMass

End Property

Property Let Pos(Obj As clsObject)

cobj = Obj

End Property

Property Get Pos() As clsObject

Set Pos = cobj

End Property

Sub Force(ByVal N As Single, ByVal Degrees As Single)

'Use f=ma, a=f/m to find acceleration
Call cobj.Accelerate(Degrees, N / TheMass)

End Sub
