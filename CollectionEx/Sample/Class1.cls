VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<dir .\Sample /dir>
Option Explicit

Private abc_
Private def_ As Class2

Property Get abc()
    abc = abc_
End Property
Property Get def()
    Set def = def_
End Property
Property Get Defs()
    Dim col As New Collection
    Dim c21 As New Class2, c22 As New Class2, c23 As New Class2
    
    Call col.Add(c21.Init(abc_))
    Call col.Add(c22.Init(abc_ + 1))
    Call col.Add(c23.Init(abc_ + 2))
    
    Set Defs = col
End Property
Function Init(a) As Class1
    abc_ = a
    Set def_ = New Class2
    def_.Init (a + 1)
    Set Init = Me
End Function

Function Equals(cls1 As Class1) As Boolean
    Equals = Me.abc = cls1.abc
End Function

Function DebugPrint()
    Debug.Print abc_
End Function

Public Function Create(a) As Class1
    Dim cls1 As New Class1
    Set Create = cls1.Init(a)
End Function
