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

Property Get def()
    def = def_
End Property

Function Init(a) As Class2
    def_ = a
    Set Init = Me
End Function
Function Equals(cls2 As Class2) As Boolean
    Equals = Me.def = cls2.def
End Function
