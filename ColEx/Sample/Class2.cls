VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<dir .\Sample /dir>
Option Explicit

Private def_


Function Init(a) As Class2
    def_ = a
    Set Init = Me
End Function


Property Get Def()
    Def = def_
End Property

Public Property Get MySelf() As Class2
    Set MySelf = Me
End Property


Public Function Equals(cls2 As Class2) As Boolean
    Equals = Me.Def = cls2.Def
End Function

Public Function ToString() As String
    ToString = Format(def_)
End Function
