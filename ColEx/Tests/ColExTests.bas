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
    test.RegisterTest "Test_SelectBy_NotGet"
    test.RegisterTest "Test_SelectManyBy"
    test.RegisterTest "Test_AnyAll"
    test.RegisterTest "Test_TakeSkip"
    test.RegisterTest "Test_FirstLast"
    test.RegisterTest "Test_Order"
    test.RegisterTest "Test_Order_ByValue"
    test.RegisterTest "Test_Contains"
    test.RegisterTest "Test_Distinct"
    test.RegisterTest "Test_ToArray"
    test.RegisterTest "Test_MinMax_Value"
    test.RegisterTest "Test_MinMax_Object"
    test.RegisterTest "Test_MinByMaxBy_Object"

    test.RunTests UnitTest
End Sub


Sub Test_SpeedTest()
    Call TestInitialize
    Dim i As Long, n
    Dim cls As New Class1
    n = Timer
'    Call ColEx(GetClass1CollectionN(10000)).OrderBy("abc")
    Call ColEx(GetClass1CollectionN(100000)).Where("abc", cexEqual, 2)
    Debug.Print "Done. " & Format((Timer - n), "0.00") & "[s]"
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
        Call .AssertEqual(5, ColEx(col_).Where("abc", cexDoesNotEqual, 2).Count)
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
Sub Test_SelectBy_NotGet()
    TestInitialize
                
    Dim res As ColEx
    With UnitTest.NameOf("SelectBy for Method/Let/Set")
        Call .AssertEqual(8, ColEx(col_).SelectBy("Create", VbMethod, 1).Count)
        Call .AssertEqual(1, ColEx(col_).SelectBy("Create", VbMethod, 1).Items(1).abc)
    
        Call .AssertEqual(8, ColEx(col_).SelectBy("abc", VbLet, 5).Count)
        Call .AssertEqual(5, col_(1).abc)
    End With
End Sub

'[Fact]
Sub Test_SelectManyBy()
    TestInitialize
        
    With UnitTest
        Call .AssertEqual(24, ColEx(col_).SelectManyBy("Defs").Count)
        Call .AssertEqual("Class2", TypeName(ColEx(col_).SelectManyBy("Defs").Items(1)))
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
    
    Dim cls As New Class1
    Dim col As New Collection
        
    With UnitTest
        .NameOf ("First/FirstOrDefault")
        Call .AssertEqual(col_(2), ColEx(col_).First("abc", cexEqual, 2))
        Call .AssertEqual(col_(1), ColEx(col_).First())
        Call .AssertEqual(col_(2), ColEx(col_).FirstOrDefault("abc", cexEqual, 2))
        Call .AssertEqual(col_(1), ColEx(col_).FirstOrDefault())
        Call .AssertEqual(0, ColEx(col_).FirstOrDefault("abc", cexEqual, 1000, 0))
        Call .AssertEqual(Null, ColEx(col_).FirstOrDefault("abc", cexEqual, 1000))
        Call .AssertTrue(ColEx(col_).FirstOrDefault("abc", cexEqual, 1000, Nothing) Is Nothing)
        On Error Resume Next
        Call ColEx(col_).First("abc", cexEqual, 100)
        Call .AssertHasError
        Call ColEx(col).First
        Call .AssertHasError
        Call .AssertTrue(ColEx(col).FirstOrDefault(, , , Nothing) Is Nothing)
        
        .NameOf ("Last/LastOrDefault")
        Call .AssertEqual(col_(2), ColEx(col_).Last("abc", cexEqual, 2))
        Call .AssertEqual(col_(col_.Count), ColEx(col_).Last())
        Call .AssertEqual(col_(2), ColEx(col_).LastOrDefault("abc", cexEqual, 2))
        Call .AssertEqual(col_(3), ColEx(col_).LastOrDefault())
        Call .AssertEqual(0, ColEx(col_).LastOrDefault("abc", cexEqual, 1000, 0))
        Call .AssertEqual(Null, ColEx(col_).LastOrDefault("abc", cexEqual, 1000))
        Call .AssertTrue(ColEx(col_).LastOrDefault("abc", cexEqual, 1000, Nothing) Is Nothing)
        On Error Resume Next
        Call ColEx(col_).Last("abc", cexEqual, 100)
        Call .AssertHasError
        Call ColEx(col).Last
        Call .AssertHasError
        Call .AssertTrue(ColEx(col).LastOrDefault(, , , Nothing) Is Nothing)
        
        .NameOf ("Single/SingleOrDefault")
        Call .AssertEqual(col_(5), ColEx(col_).SingleBy("abc", cexEqual, 5))
        Call .AssertEqual(col_(5), ColEx(col_).SingleOrDefaultBy("abc", cexEqual, 5))
        Call .AssertTrue(ColEx(col_).SingleOrDefaultBy("abc", cexEqual, 1000, Nothing) Is Nothing)
        On Error Resume Next
        Call ColEx(col_).SingleBy("abc", cexEqual, 100)
        Call .AssertHasError
        Call ColEx(col_).SingleBy("abc", cexEqual, 2)
        Call .AssertHasError
        Call ColEx(col_).SingleOrDefaultBy("abc", cexEqual, 2)
        Call .AssertHasError
        Call ColEx(col_).SingleBy
        Call .AssertHasError
        Call ColEx(col_).SingleOrDefaultBy
        Call .AssertHasError
        Call ColEx(col).SingleBy
        Call .AssertHasError
        Call .AssertTrue(ColEx(col).SingleOrDefaultBy(, , , Nothing) Is Nothing)
                    
        Set col = New Collection
        Call col.Add(cls.Create(1))
        Call .AssertEqual(cls.Create(1), ColEx(col).SingleBy)
        Call .AssertEqual(cls.Create(1), ColEx(col).SingleOrDefaultBy)

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
    With UnitTest.NameOf("Order Class1 ")
        Set res = ColEx(col).OrderBy("abc").Items
        Call .AssertTrue(res(1).abc <= res(2).abc)
        Call .AssertTrue(res(2).abc <= res(3).abc)
        Call .AssertTrue(res(3).abc <= res(4).abc)
        Call .AssertTrue(res(4).abc <= res(5).abc)
        Call .AssertTrue(res(5).abc <= res(6).abc)
        
        Set res = ColEx(col).OrderByDescending("abc").Items
        Call .AssertTrue(res(1).abc >= res(2).abc)
        Call .AssertTrue(res(2).abc >= res(3).abc)
        Call .AssertTrue(res(3).abc >= res(4).abc)
        Call .AssertTrue(res(4).abc >= res(5).abc)
        Call .AssertTrue(res(5).abc >= res(6).abc)
        
        Set res = ColEx(col).OrderByDescending("abc").OrderBy("abc").Items
        Call .AssertEqual(col(1).abc, res(1).abc)
        Call .AssertEqual(col(col.Count).abc, res(res.Count).abc)
    End With
        
End Sub

'[Fact]
Sub Test_Order_ByValue()
    Dim col As Collection
    Dim res As Collection
    With UnitTest.NameOf("Order asc/desc of Int even")
        Set res = ColEx(GetIntCollection).OrderBy().Items
        Call .AssertTrue(res(1) <= res(2))
        Call .AssertTrue(res(2) <= res(3))
        Call .AssertTrue(res(3) <= res(4))
        Call .AssertTrue(res(4) <= res(5))
        Call .AssertTrue(res(5) <= res(6))
        Call .AssertTrue(res(6) <= res(7))
        Call .AssertTrue(res(7) <= res(8))
        
        Set res = ColEx(GetIntCollection).OrderByDescending().Items
        Call .AssertTrue(res(1) >= res(2))
        Call .AssertTrue(res(2) >= res(3))
        Call .AssertTrue(res(3) >= res(4))
        Call .AssertTrue(res(4) >= res(5))
        Call .AssertTrue(res(5) >= res(6))
        Call .AssertTrue(res(6) >= res(7))
        Call .AssertTrue(res(7) >= res(8))
    End With
        
    
    Set col = GetIntCollection: Call col.Add(9)
    With UnitTest.NameOf("Order asc/desc of Int odd")
        Set res = ColEx(col).OrderBy().Items
        Call .AssertTrue(res(1) <= res(2))
        Call .AssertTrue(res(2) <= res(3))
        Call .AssertTrue(res(3) <= res(4))
        Call .AssertTrue(res(4) <= res(5))
        Call .AssertTrue(res(5) <= res(6))
        Call .AssertTrue(res(6) <= res(7))
        Call .AssertTrue(res(7) <= res(8))
        Call .AssertTrue(res(8) <= res(9))
        
        Set res = ColEx(col).OrderByDescending().Items
        Call .AssertTrue(res(1) >= res(2))
        Call .AssertTrue(res(2) >= res(3))
        Call .AssertTrue(res(3) >= res(4))
        Call .AssertTrue(res(4) >= res(5))
        Call .AssertTrue(res(5) >= res(6))
        Call .AssertTrue(res(6) >= res(7))
        Call .AssertTrue(res(7) >= res(8))
        Call .AssertTrue(res(8) >= res(9))
    End With
    
    
    With UnitTest.NameOf("Order asc/desc of String even")
        Set res = ColEx(GetStringCollection).OrderBy().Items
        Call .AssertTrue(res(1) <= res(2))
        Call .AssertTrue(res(2) <= res(3))
        Call .AssertTrue(res(3) <= res(4))
        Call .AssertTrue(res(4) <= res(5))
        Call .AssertTrue(res(5) <= res(6))
        Call .AssertTrue(res(6) <= res(7))
        Call .AssertTrue(res(7) <= res(8))
        
        Set res = ColEx(GetStringCollection).OrderByDescending().Items
        Call .AssertTrue(res(1) >= res(2))
        Call .AssertTrue(res(2) >= res(3))
        Call .AssertTrue(res(3) >= res(4))
        Call .AssertTrue(res(4) >= res(5))
        Call .AssertTrue(res(5) >= res(6))
        Call .AssertTrue(res(6) >= res(7))
        Call .AssertTrue(res(7) >= res(8))
    End With
        
    Set col = GetStringCollection: Call col.Add("ccc")
    With UnitTest.NameOf("Order asc/desc of String odd")
        Set res = ColEx(col).OrderBy().Items
        Call .AssertTrue(res(1) <= res(2))
        Call .AssertTrue(res(2) <= res(3))
        Call .AssertTrue(res(3) <= res(4))
        Call .AssertTrue(res(4) <= res(5))
        Call .AssertTrue(res(5) <= res(6))
        Call .AssertTrue(res(6) <= res(7))
        Call .AssertTrue(res(7) <= res(8))
        Call .AssertTrue(res(8) <= res(9))
        
        Set res = ColEx(col).OrderByDescending().Items
        Call .AssertTrue(res(1) >= res(2))
        Call .AssertTrue(res(2) >= res(3))
        Call .AssertTrue(res(3) >= res(4))
        Call .AssertTrue(res(4) >= res(5))
        Call .AssertTrue(res(5) >= res(6))
        Call .AssertTrue(res(6) >= res(7))
        Call .AssertTrue(res(7) >= res(8))
        Call .AssertTrue(res(8) >= res(9))
    End With
        
    
    With UnitTest.NameOf("Order asc/desc, 2 elements order")
        Set col = New Collection
        Call col.Add(2)
        Call col.Add(1)
        Set res = ColEx(col).OrderBy().Items
        Call .AssertTrue(res(1) <= res(2))
    End With
    With UnitTest.NameOf("Order asc/desc, 1 elements return self")
        Set col = New Collection
        Call col.Add(1)
        Set res = ColEx(col).OrderBy().Items
        Call .AssertEqual(1, res.Count)
    End With
    With UnitTest.NameOf("Order asc/desc, 0 elements return empty collection")
        Set col = New Collection
        Set res = ColEx(col).OrderBy().Items
        Call .AssertEqual(0, res.Count)
    End With
        
End Sub

'[Fact]
Sub Test_Contains()
    
    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    
    With UnitTest.NameOf("Contains")
        Call .AssertTrue(ColEx(col).Contains(cls.Create(2)))
        Call .AssertFalse(ColEx(col).Contains(cls.Create(7)))
    End With
End Sub


'[Fact]
Sub Test_Distinct()
    
    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(5))
    
    Dim res As Collection
    With UnitTest
        Call .NameOf("Distinct")
        Set res = ColEx(col).Distinct().Items
        Call .AssertEqual(4, res.Count)
        Call .AssertEqual(1, ColEx(col).Distinct().Where("abc", cexEqual, 1).Count)
        
        Call .NameOf("DistinctBy")
        Set res = ColEx(col).DistinctBy("def").Items
        Call .AssertEqual(4, res.Count)
        Call .AssertEqual(7, ColEx(col).SelectManyBy("Defs").DistinctBy("def").Count)
    End With
        
End Sub

'[Fact]
Sub Test_ToArray()
    
    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(5))
    
    Dim res As Variant
    With UnitTest
        Call .NameOf("ToArray")
        res = ColEx(col).ToArray()
        Call .AssertTrue(IsArray(res))
        Call .AssertEqual(4, UBound(res))
        Call .AssertEqual(0, LBound(res))

        Call .NameOf("ToArray2D, param array")
        res = ColEx(col).ToArray2D("abc", "def")
        Call .AssertTrue(IsArray(res))
        Call .AssertEqual(4, UBound(res, 1))
        Call .AssertEqual(0, LBound(res, 1))
        Call .AssertEqual(1, UBound(res, 2))
        Call .AssertEqual(0, LBound(res, 2))
        Call .AssertEqual(1, res(0, 0))
        Call .AssertEqual(5, res(4, 0))
                
        Call .NameOf("ToArray2D, by collection")
        Dim col_names As New Collection
        Call col_names.Add("abc")
        Call col_names.Add("def")
        res = ColEx(col).ToArray2D(col_names)
        Call .AssertTrue(IsArray(res))
        Call .AssertEqual(4, UBound(res, 1))
        Call .AssertEqual(0, LBound(res, 1))
        Call .AssertEqual(1, UBound(res, 2))
        Call .AssertEqual(0, LBound(res, 2))
        Call .AssertEqual(1, res(0, 0))
        Call .AssertEqual(5, res(4, 0))
        
        Call .NameOf("ToArray2D, by array")
        Dim arr_names As Variant
        arr_names = Array("abc", "def")
        res = ColEx(col).ToArray2D(arr_names)
        Call .AssertTrue(IsArray(res))
        Call .AssertEqual(4, UBound(res, 1))
        Call .AssertEqual(0, LBound(res, 1))
        Call .AssertEqual(1, UBound(res, 2))
        Call .AssertEqual(0, LBound(res, 2))
        Call .AssertEqual(1, res(0, 0))
        Call .AssertEqual(5, res(4, 0))

    End With
        
End Sub

'[Fact]
Sub Test_MinMax_Value()
    
    ' Arrange
    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(2)
    Call col.Add(1)
    Call col.Add(5)
    Call col.Add(4)
    Call col.Add(3)
        
    ' Act/Assert
    With UnitTest.NameOf("Min/Max (Value)")
        Call .AssertEqual(1, ColEx(col).Min())
        Call .AssertEqual(5, ColEx(col).Max())
    End With
        
End Sub

'[Fact]
Sub Test_MinMax_Object()
    
    ' Arrange
    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(5))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(3))
    
    ' Act/Assert
    With UnitTest.NameOf("Min/Max (Object)")
        Call .AssertEqual(1, ColEx(col).Min("abc"))
        Call .AssertEqual(5, ColEx(col).Max("abc"))
    End With
        
End Sub

'[Fact]
Sub Test_MinByMaxBy_Object()
    
    ' Arrange
    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(5))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(3))
    
    ' Act/Assert
    With UnitTest.NameOf("MinBy/MaxBy")
        Call .AssertTrue(TypeOf ColEx(col).MinBy("abc") Is Class1)
        Call .AssertEqual(1, ColEx(col).MinBy("abc").abc)
        Call .AssertTrue(TypeOf ColEx(col).MaxBy("abc") Is Class1)
        Call .AssertEqual(5, ColEx(col).MaxBy("abc").abc)
    End With
        
End Sub

Private Function GetClass1Collection()
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

    Set GetClass1Collection = col
End Function

Private Function GetClass1CollectionN(Optional n As Long = 10)
    Dim cls As New Class1, i As Long
    Dim col As New Collection
    
    For i = 1 To n
        Call col.Add(cls.Create(i))
    Next i
    Set GetClass1CollectionN = col
End Function

Private Function GetIntCollection()
    Dim col As New Collection
    Call col.Add(1)
    Call col.Add(2)
    Call col.Add(3)
    Call col.Add(4)
    Call col.Add(5)
    Call col.Add(2)
    Call col.Add(2)
    Call col.Add(3)

    Set GetIntCollection = col
End Function

Private Function GetStringCollection()
    Dim col As New Collection
    Call col.Add("aaa")
    Call col.Add("aab")
    Call col.Add("aac")
    Call col.Add("aba")
    Call col.Add("caa")
    Call col.Add("cba")
    Call col.Add("aaa")
    Call col.Add("aab")

    Set GetStringCollection = col
End Function
