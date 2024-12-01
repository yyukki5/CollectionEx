Attribute VB_Name = "ColExTests"
'<dir .\Tests /dir>

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

Private Sub CreateTests()
    Call UnitTest.CreateRunTests("ColExTests")
End Sub

Sub RunTests()
   Dim test As New UnitTest

    test.RegisterTest "Test_Initialize_Create_Enum"
    test.RegisterTest "Test_Add"
    test.RegisterTest "Test_Where"
    test.RegisterTest "Test_SelectBy"
    test.RegisterTest "Test_AnyAll"
    test.RegisterTest "Test_TakeSkip"
    test.RegisterTest "Test_FirstLast"
    test.RegisterTest "Test_Order"

    test.RunTests UnitTest
End Sub


Sub Test_SpeedTest()
    Call TestInitialize
    Dim i As Long, n
    Dim cls As New Class1
    n = Timer
    For i = 1 To 10000
        Call ColEx(col_).Where("abc", cexEqual, 2)
    Next i
    Debug.Print "Done. " & Format((Timer - n), "0.00") & "[s]"
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
Sub Test_Initialize_Create_Enum()

    Dim cex As New ColEx
    Dim cls As New Class1
        
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(5))
    
    With UnitTest.NameOf("Initialize and Create and Enum")
        Call .AssertEqual("ColEx", TypeName(ColEx(col)))
        Call .AssertEqual(col.Count, ColEx(col).Count)
        
        Dim cce1 As ColEx
        Set cce1 = ColEx(col)
        Call .AssertEqual(col.Count, cce1.Count)
        
        Call .AssertEqual(5, ColEx(Array(1, 2, 3, 4, 5)).Items.Count)
        
        Dim v As Variant, added As Long
        For Each v In ColEx(col)
           added = added + v.abc
        Next
        Call .AssertHasNoError
        Call .AssertEqual(15, added)
    End With
        
    With UnitTest.NameOf("collection instance is not same (copied)")
        Call .AssertFalse(ColEx(col).Items Is col)
    End With
End Sub


'[Fact]
Sub Test_Add()
    TestInitialize
    
    With UnitTest
        Call .NameOf("Add/AddRange")
        Call .AssertEqual(10, ColEx(col_).Add(col_(1)).Add(col_(2)).Count)
        Call .AssertEqual(16, ColEx(col_).AddRange(col_).Count)
    End With
End Sub

'[Fact]
Sub Test_Where()
    TestInitialize
    
    With UnitTest
        Call .AssertEqual(1, ColEx(col_).Where("abc", cexEqual, 1).Count)
        Call .AssertEqual(3, ColEx(col_).Where("abc", cexEqual, 2).Count)
        Call .AssertEqual(2, ColEx(col_).Where("abc", cexGreaterThan, 3).Count)
        Call .AssertEqual(4, ColEx(col_).Where("abc", cexGreaterThanOrEqualTo, 3).Count)
        Call .AssertEqual(4, ColEx(col_).Where("abc", cexLessThan, 3).Count)
        Call .AssertEqual(6, ColEx(col_).Where("abc", cexLessThanOrEqualTo, 3).Count)
    End With
End Sub

'[Fact]
Sub Test_SelectBy()
    TestInitialize
    
    With UnitTest
        Call .AssertEqual(8, ColEx(col_).SelectBy("abc").Count)
        Call .AssertEqual(1, ColEx(col_).SelectBy("abc").Items(1))
        Call .AssertEqual(2, ColEx(col_).SelectBy("def").Items(1).def)
    End With
End Sub

'[Fact]
Sub Test_AnyAll()
    TestInitialize
    
    With UnitTest
        Call .NameOf("Any")
        Call .AssertTrue(ColEx(col_).AnyBy("abc", cexGreaterThan, 3))
        Call .AssertFalse(ColEx(col_).AnyBy("abc", cexGreaterThan, 1000))
    
        Call .NameOf("All")
        Call .AssertTrue(ColEx(col_).AllBy("abc", cexGreaterThan, 0))
        Call .AssertFalse(ColEx(col_).AllBy("abc", cexGreaterThan, 4))
    End With
End Sub


'[Fact]
Sub Test_TakeSkip()
    TestInitialize
    
    Dim col As Collection
    Set col = col_

    With UnitTest
        Call .AssertEqual(3, ColEx(col).Take(3).Count)
        Call .AssertEqual(col.Count - 3, ColEx(col).Skip(3).Count)
        Call .AssertEqual(col.Count, ColEx(col).Take(100).Count)
        Call .AssertEqual(0, ColEx(col).Skip(100).Count)
    End With
End Sub

'[Fact]
Sub Test_FirstLast()
    TestInitialize
    
    With UnitTest
        Call .AssertEqual(col_(2), ColEx(col_).First("abc", cexEqual, 2))
        Call .AssertEqual(col_(1), ColEx(col_).First())
        Call .AssertEqual(col_(2), ColEx(col_).FirstOrDefault("abc", cexEqual, 2))
        Call .AssertEqual(col_(1), ColEx(col_).FirstOrDefault())
        Call .AssertEqual(0, ColEx(col_).FirstOrDefault("abc", cexEqual, 1000, 0))
        Call .AssertEqual(Null, ColEx(col_).FirstOrDefault("abc", cexEqual, 1000))
        Call .AssertTrue(ColEx(col_).FirstOrDefault("abc", cexEqual, 1000, Nothing) Is Nothing)
        
        Call .AssertEqual(col_(2), ColEx(col_).Last("abc", cexEqual, 2))
        Call .AssertEqual(col_(col_.Count), ColEx(col_).Last())
        Call .AssertEqual(col_(2), ColEx(col_).LastOrDefault("abc", cexEqual, 2))
        Call .AssertEqual(col_(3), ColEx(col_).LastOrDefault())
        Call .AssertEqual(0, ColEx(col_).LastOrDefault("abc", cexEqual, 1000, 0))
        Call .AssertEqual(Null, ColEx(col_).LastOrDefault("abc", cexEqual, 1000))
        Call .AssertTrue(ColEx(col_).LastOrDefault("abc", cexEqual, 1000, Nothing) Is Nothing)
    End With
End Sub


'[Fact]
Sub Test_Order()
    
    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(5))
    
    Dim res As Collection
    With UnitTest
        Set res = ColEx(col).OrderByDescending("abc").Items
        Call .AssertTrue(res(1).abc >= res(2).abc)
        Call .AssertTrue(res(2).abc >= res(3).abc)
        Call .AssertTrue(res(3).abc >= res(4).abc)
        Call .AssertTrue(res(4).abc >= res(5).abc)
        
        Set res = ColEx(col).OrderByDescending("abc").OrderBy("abc").Items
        Call .AssertEqual(col(1).abc, res(1).abc)
        Call .AssertEqual(col(col.Count).abc, res(res.Count).abc)
    End With
        
End Sub
