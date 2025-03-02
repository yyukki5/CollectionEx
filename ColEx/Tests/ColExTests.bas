Attribute VB_Name = "ColExTests"
'<dir .\Tests /dir>

Option Explicit

Private Sub CreateTests()
    Call UnitTest.CreateRunTests("ColExTests")
End Sub

Sub RunTests()
   Dim test As New UnitTest

    test.RegisterTest "Test_SpeedTests"
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

'[Fact]
Sub Test_SpeedTests()
    Dim col As Collection, not_using_time As Double, using_time As Double
    ' Should be less than 100000 elements, by the avoiding the upper limit of garbage collection of "Collection".
    Set col = GetClass1CollectionN(50000)
    
    With UnitTest
        Call .NameOf("Where method Time")
        not_using_time = Test_SpeedTest("Where 1", ColEx(col))
        using_time = Test_SpeedTest("Where 1 layer", ColEx(col))
        Call .AssertTrue(using_time < 0.0001 Or not_using_time * 10 > using_time)
        
        not_using_time = Test_SpeedTest("Where 2", ColEx(col))
        using_time = Test_SpeedTest("Where 2 layer", ColEx(col))
        Call .AssertTrue(using_time < 0.0001 Or not_using_time * 10 > using_time)

    End With
    
    Set col = Nothing
End Sub

Private Function Test_SpeedTest(test_name As String, cex As ColEx) As Double
    Dim col As New Collection, c As Class1
    Dim n:    n = Timer
    
    Select Case test_name
        Case "Not using Where 1 layer"
            For Each c In cex
                If c.Abc = 2 Then Call col.Add(c)
            Next
        Case "Not using Where 2 layer"
            For Each c In cex
                If c.Def.Def = 2 Then Call col.Add(c)
            Next
        Case "Using Where 1 layer":       Call cex.Where("Abc", cexEqual, 2)
        Case "Using Where 2 layer":       Call cex.Where("Def.Def", cexEqual, 2)
    End Select
    
    
    Test_SpeedTest = (Timer - n)
'    Debug.Print test_name & ":" & Format(Test_SpeedTest, "0.0000") & "[s]"
End Function

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
           added = added + v.Abc
        Next
        Call .AssertHasNoError
        Call .AssertEqual(15, added)
    End With
        
    With UnitTest.NameOf("collection instance is not same (copied)")
        Call .AssertNotSame(ColEx(col).Items, col)
    End With
End Sub

'[Fact]
Sub Test_Add()
    Dim col As Collection: Set col = GetClass1Collection()
    
    With UnitTest
        Call .NameOf("Add/AddRange")
        Call .AssertEqual(10, ColEx(col).Add(col(1)).Add(col(2)).Count)
        Call .AssertEqual(16, ColEx(col).AddRange(col).Count)
    End With
End Sub

'[Fact]
Sub Test_Where()
    Dim col As Collection: Set col = GetClass1Collection()
    
    With UnitTest
        Call .NameOf("Where")
        Call .AssertEqual(1, ColEx(col).Where("Abc", cexEqual, 1).Count)
        Call .AssertEqual(3, ColEx(col).Where("Abc", cexEqual, 2).Count)
        Call .AssertEqual(5, ColEx(col).Where("Abc", cexDoesNotEqual, 2).Count)
        Call .AssertEqual(2, ColEx(col).Where("Abc", cexGreaterThan, 3).Count)
        Call .AssertEqual(4, ColEx(col).Where("Abc", cexGreaterThanOrEqualTo, 3).Count)
        Call .AssertEqual(4, ColEx(col).Where("Abc", cexLessThan, 3).Count)
        Call .AssertEqual(6, ColEx(col).Where("Abc", cexLessThanOrEqualTo, 3).Count)
        
        Call .NameOf("Where 2 layer")
        Call .AssertEqual(1, ColEx(col).Where("Def.Def", cexEqual, 2).Count)
        Call .AssertEqual(3, ColEx(col).Where("Def.Def", cexEqual, 3).Count)
        Call .AssertEqual(5, ColEx(col).Where("Def.Def", cexDoesNotEqual, 3).Count)
        Call .AssertEqual(2, ColEx(col).Where("Def.Def", cexGreaterThan, 4).Count)
        Call .AssertEqual(4, ColEx(col).Where("Def.Def", cexGreaterThanOrEqualTo, 4).Count)
        Call .AssertEqual(4, ColEx(col).Where("Def.Def", cexLessThan, 4).Count)
        Call .AssertEqual(6, ColEx(col).Where("Def.Def", cexLessThanOrEqualTo, 4).Count)
        
    End With
End Sub


'[Fact]
Sub Test_SelectBy()
    Dim col As Collection:  Set col = GetClass1Collection()
    
    With UnitTest
        Call .NameOf("SelectBy")
        Call .AssertEqual(8, ColEx(col).SelectBy("Abc").Count)
        Call .AssertEqual(1, ColEx(col).SelectBy("Abc").Items(1))
        Call .AssertEqual(2, ColEx(col).SelectBy("Def").Items(1).Def)
    
        Call .NameOf("SelectBy 2 layer")
        Call .AssertEqual(8, ColEx(col).SelectBy("Def.Def").Count)
        Call .AssertEqual(2, ColEx(col).SelectBy("Def.Def").Items(1))
    End With
End Sub

'[Fact]
Sub Test_SelectBy_NotGet()
    Dim col As Collection:  Set col = GetClass1Collection()
                
    Dim res As ColEx
    With UnitTest
        Call .NameOf("SelectBy for Method/Let/Set")
        Call .AssertEqual(8, ColEx(col).SelectBy("Create", VbMethod, 1).Count)
        Call .AssertEqual(1, ColEx(col).SelectBy("Create", VbMethod, 1).Items(1).Abc)
    
        Call .AssertEqual(8, ColEx(col).SelectBy("Abc", VbLet, 5).Count)
        Call .AssertEqual(5, col(1).Abc)
        
        Call .NameOf("SelectBy for Method/Let/Set , 2nd layer")
        Call .AssertEqual(8, ColEx(col).SelectBy("Def.ToString", VbMethod).Count)
        Call .AssertEqual("2", ColEx(col).SelectBy("Def.ToString", VbMethod).Items(1))
    End With
End Sub

'[Fact]
Sub Test_SelectManyBy()
    Dim col As Collection:  Set col = GetClass1Collection()
        
    With UnitTest
        Call .NameOf("SelectManyBy")
        Call .AssertEqual(24, ColEx(col).SelectManyBy("Defs").Count)
        Call .AssertEqual("Class2", TypeName(ColEx(col).SelectManyBy("Defs").Items(1)))
    End With
End Sub


'[Fact]
Sub Test_AnyAll()
    Dim col As Collection:  Set col = GetClass1Collection()
    
    With UnitTest
        Call .NameOf("Any")
        Call .AssertTrue(ColEx(col).AnyBy("Abc", cexGreaterThan, 3))
        Call .AssertFalse(ColEx(col).AnyBy("Abc", cexGreaterThan, 1000))
    
        Call .NameOf("Any, 2nd layer")
        Call .AssertTrue(ColEx(col).AnyBy("Def.Def", cexGreaterThan, 3))
        Call .AssertFalse(ColEx(col).AnyBy("Def.Def", cexGreaterThan, 1000))
    
        Call .NameOf("All")
        Call .AssertTrue(ColEx(col).AllBy("Abc", cexGreaterThan, 0))
        Call .AssertFalse(ColEx(col).AllBy("Abc", cexGreaterThan, 4))
    
        Call .NameOf("All, 2nd layer")
        Call .AssertTrue(ColEx(col).AllBy("Def.Def", cexGreaterThan, 0))
        Call .AssertFalse(ColEx(col).AllBy("Def.Def", cexGreaterThan, 4))
    End With
End Sub


'[Fact]
Sub Test_TakeSkip()
    Dim col As Collection:  Set col = GetClass1Collection()

    With UnitTest
        Call .NameOf("Take, Skip")
        Call .AssertEqual(3, ColEx(col).Take(3).Count)
        Call .AssertEqual(col.Count - 3, ColEx(col).Skip(3).Count)
        Call .AssertEqual(col.Count, ColEx(col).Take(100).Count)
        Call .AssertEqual(0, ColEx(col).Skip(100).Count)
    End With
End Sub

'[Fact]
Sub Test_FirstLast()
    Dim col As Collection:  Set col = GetClass1Collection()
    Dim cls As New Class1
    Dim emptyCollection As New Collection
        
    With UnitTest
        .NameOf ("First/FirstOrDefault")
        Call .AssertEqual(col(2), ColEx(col).First("Abc", cexEqual, 2))
        Call .AssertEqual(col(1), ColEx(col).First("Def.Def", cexEqual, 2))
        
        Call .AssertEqual(col(1), ColEx(col).First())
        Call .AssertEqual(col(2), ColEx(col).FirstOrDefault("Abc", cexEqual, 2))
        Call .AssertEqual(col(1), ColEx(col).FirstOrDefault())
        Call .AssertEqual(0, ColEx(col).FirstOrDefault("Abc", cexEqual, 1000, 0))
        Call .AssertEqual(Null, ColEx(col).FirstOrDefault("Abc", cexEqual, 1000))
        Call .AssertTrue(ColEx(col).FirstOrDefault("Abc", cexEqual, 1000, Nothing) Is Nothing)
        
        On Error Resume Next
        Call ColEx(col).First("Abc", cexEqual, 100)
        Call .AssertHasError
        Call Err.Clear
        
        Call ColEx(emptyCollection).First
        Call .AssertHasError
        Call Err.Clear
        On Error GoTo 0
        
        Call .AssertNothing(ColEx(emptyCollection).FirstOrDefault(, , , Nothing))
        
        
        .NameOf ("Last/LastOrDefault")
        Call .AssertEqual(col(7), ColEx(col).Last("Abc", cexEqual, 2))
        Call .AssertEqual(col(1), ColEx(col).Last("Def.Def", cexEqual, 2))
        Call .AssertEqual(col(col.Count), ColEx(col).Last())
        Call .AssertEqual(col(2), ColEx(col).LastOrDefault("Abc", cexEqual, 2))
        Call .AssertEqual(col(3), ColEx(col).LastOrDefault())
        Call .AssertEqual(0, ColEx(col).LastOrDefault("Abc", cexEqual, 1000, 0))
        Call .AssertEqual(Null, ColEx(col).LastOrDefault("Abc", cexEqual, 1000))
        Call .AssertTrue(ColEx(col).LastOrDefault("Abc", cexEqual, 1000, Nothing) Is Nothing)
        
        On Error Resume Next
        Call ColEx(col).Last("Abc", cexEqual, 100)
        Call .AssertHasError
        Call Err.Clear
        
        Call ColEx(emptyCollection).Last
        Call .AssertHasError
        Call Err.Clear
        On Error GoTo 0
                
        Call .AssertNothing(ColEx(emptyCollection).LastOrDefault(, , , Nothing))
        
        
        .NameOf ("Single/SingleOrDefault")
        Call .AssertEqual(col(5), ColEx(col).SingleBy("Abc", cexEqual, 5))
        Call .AssertEqual(col(5), ColEx(col).SingleBy("Def.Def", cexEqual, 6))
        Call .AssertEqual(col(5), ColEx(col).SingleOrDefaultBy("Abc", cexEqual, 5))
        Call .AssertTrue(ColEx(col).SingleOrDefaultBy("Abc", cexEqual, 1000, Nothing) Is Nothing)
        
        On Error Resume Next
        Call ColEx(col).SingleBy("Abc", cexEqual, 100)
        Call .AssertHasError
        Call Err.Clear
        
        Call ColEx(col).SingleBy("Abc", cexEqual, 2)
        Call .AssertHasError
        Call Err.Clear
        
        Call ColEx(col).SingleOrDefaultBy("Abc", cexEqual, 2)
        Call .AssertHasError
        Call Err.Clear
        
        Call ColEx(col).SingleBy
        Call .AssertHasError
        Call Err.Clear
        
        Call ColEx(col).SingleOrDefaultBy
        Call .AssertHasError
        Call Err.Clear
        
        Call ColEx(emptyCollection).SingleBy
        Call .AssertHasError
        Call Err.Clear
        
        Call .AssertNothing(ColEx(emptyCollection).SingleOrDefaultBy(, , , Nothing))
                    
        Set emptyCollection = New Collection
        Call emptyCollection.Add(cls.Create(1))
        Call .AssertEqual(cls.Create(1), ColEx(emptyCollection).SingleBy)
        Call .AssertEqual(cls.Create(1), ColEx(emptyCollection).SingleOrDefaultBy)

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
        .NameOf ("Order Class1 ")
        Set res = ColEx(col).OrderBy("Abc").Items
        Call .AssertTrue(res(1).Abc <= res(2).Abc)
        Call .AssertTrue(res(2).Abc <= res(3).Abc)
        Call .AssertTrue(res(3).Abc <= res(4).Abc)
        Call .AssertTrue(res(4).Abc <= res(5).Abc)
        Call .AssertTrue(res(5).Abc <= res(6).Abc)
        
        Set res = ColEx(col).OrderByDescending("Abc").Items
        Call .AssertTrue(res(1).Abc >= res(2).Abc)
        Call .AssertTrue(res(2).Abc >= res(3).Abc)
        Call .AssertTrue(res(3).Abc >= res(4).Abc)
        Call .AssertTrue(res(4).Abc >= res(5).Abc)
        Call .AssertTrue(res(5).Abc >= res(6).Abc)
        
        Set res = ColEx(col).OrderByDescending("Abc").OrderBy("Abc").Items
        Call .AssertEqual(col(1).Abc, res(1).Abc)
        Call .AssertEqual(col(col.Count).Abc, res(res.Count).Abc)
        
        
        .NameOf ("Order Class1, 2nd Layer")
        Set res = ColEx(col).OrderBy("Def.Def").Items
        Call .AssertTrue(res(1).Def.Def <= res(2).Def.Def)
        Call .AssertTrue(res(2).Def.Def <= res(3).Def.Def)
        Call .AssertTrue(res(3).Def.Def <= res(4).Def.Def)
        Call .AssertTrue(res(4).Def.Def <= res(5).Def.Def)
        Call .AssertTrue(res(5).Def.Def <= res(6).Def.Def)
        
        Set res = ColEx(col).OrderByDescending("Def.Def").Items
        Call .AssertTrue(res(1).Def.Def >= res(2).Def.Def)
        Call .AssertTrue(res(2).Def.Def >= res(3).Def.Def)
        Call .AssertTrue(res(3).Def.Def >= res(4).Def.Def)
        Call .AssertTrue(res(4).Def.Def >= res(5).Def.Def)
        Call .AssertTrue(res(5).Def.Def >= res(6).Def.Def)
                
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
        Call .AssertEqual(1, ColEx(col).Distinct().Where("Abc", cexEqual, 1).Count)
        
        Call .NameOf("DistinctBy")
        Set res = ColEx(col).DistinctBy("Def").Items
        Call .AssertEqual(4, res.Count)
        Call .AssertEqual(7, ColEx(col).SelectManyBy("Defs").DistinctBy("Def").Count)
    
        Call .NameOf("DistinctBy 2nd Layer")
        Set res = ColEx(col).DistinctBy("Def.Def").Items
        Call .AssertEqual(4, res.Count)
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
        res = ColEx(col).ToArray2D("Abc", "Def")
        Call .AssertTrue(IsArray(res))
        Call .AssertEqual(4, UBound(res, 1))
        Call .AssertEqual(0, LBound(res, 1))
        Call .AssertEqual(1, UBound(res, 2))
        Call .AssertEqual(0, LBound(res, 2))
        Call .AssertEqual(1, res(0, 0))
        Call .AssertEqual(5, res(4, 0))
                
        Call .NameOf("ToArray2D, by collection")
        Dim col_names As New Collection
        Call col_names.Add("Abc")
        Call col_names.Add("Def")
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
        arr_names = Array("Abc", "Def")
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
    With UnitTest
        Call .NameOf("Min/Max (Object)")
        Call .AssertEqual(1, ColEx(col).Min("Abc"))
        Call .AssertEqual(5, ColEx(col).Max("Abc"))
    
        Call .NameOf("Min/Max (Object), 2nd layer")
        Call .AssertEqual(2, ColEx(col).Min("Def.Def"))
        Call .AssertEqual(6, ColEx(col).Max("Def.Def"))
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
    With UnitTest
        Call .NameOf("MinBy/MaxBy")
        Call .AssertTrue(TypeOf ColEx(col).MinBy("Abc") Is Class1)
        Call .AssertEqual(1, ColEx(col).MinBy("Abc").Abc)
        Call .AssertTrue(TypeOf ColEx(col).MaxBy("Abc") Is Class1)
        Call .AssertEqual(5, ColEx(col).MaxBy("Abc").Abc)
    
        Call .NameOf("MinBy/MaxBy, 2nd layer")
        Call .AssertTrue(TypeOf ColEx(col).MinBy("Def.Def") Is Class1)
        Call .AssertEqual(1, ColEx(col).MinBy("Def.Def").Abc)
        Call .AssertTrue(TypeOf ColEx(col).MaxBy("Def.Def") Is Class1)
        Call .AssertEqual(5, ColEx(col).MaxBy("Def.Def").Abc)
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
