Attribute VB_Name = "ColExTests"
Option Explicit

Private col_ As Collection

Private Sub TestInitialize()
    
    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(5))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))

    Set col_ = col

End Sub

'[Fact]
Public Sub testsss()

    Dim cce As New ColEx
    Dim cls As New Class1
    
    
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(5))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))

    
    Call cce.Initialize(col)
    
    Dim cce1 As ColEx
    Set cce1 = cce.Where("abc", Equal, 2).SelectBy("def").Where("def", Equal, 3)
    
    Dim cce3 As ColEx
    Set cce3 = cce.Where(, Equal, cls.Create(3)).SelectBy("def")
        
End Sub





Sub aaaaaaaaaa()
Dim i As Long
For i = 1 To 1000
    testsss
Next i
Debug.Print "Done"
End Sub


Sub aaaa()

    Dim cex As ColEx
    Set cex = ColEx(Array(1, 2, 3, 4, 5))
    
    Dim aa As Variant
    For Each aa In cex
        Debug.Print aa
    Next
    
End Sub

'[Fact]
Public Sub Test_Initialize_Clone_Enum()

    Dim cex As New ColEx
    Dim cls As New Class1
        
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(5))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))

    
    Dim cce1 As ColEx
    Set cce1 = ColEx(col).Where("abc", Equal, 2).SelectBy("def").Where("def", Equal, 3)
    
    Dim cce3 As ColEx
    Set cce3 = ColEx(col).Where(, Equal, cls.Create(3)).SelectBy("def")
    
    
End Sub

'[Fact]
Sub Test_TakeSkip()
    TestInitialize
    
    Dim col As Collection
    Set col = col_

    Dim c As Class1
    For Each c In ColEx(col).Skip(3).Take(3)
        Debug.Print c.abc
    Next

    With UnitTest
        Call .AssertEqual(col.Count, ColEx(col).Take(100).Count)
        Call .AssertEqual(0, ColEx(col).Skip(100).Count)
    End With
End Sub
