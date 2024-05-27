VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private abc_
Private def_ As Class2

Property Get abc()
    abc = abc_
End Property
Property Get def()
    Set def = def_
End Property
Property Get parg(arg)
    parg = arg
End Property
Property Get ReCls2(arg) As Class2
    Dim cls2 As New Class2
    Set ReCls2 = cls2.Init(arg)
End Property
Property Get IsCls2(cls2 As Class2)
    IsCls2 = cls2.def
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
