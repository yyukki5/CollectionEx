Attribute VB_Name = "CollectionExTests"
'<dir .\Tests /dir>

' Note --------------------------------------
' This is unit tests of CollectionEx.cls
' Depend on: UnitTest.cls
' -------------------------------------------

Option Explicit

Private colVal_ As Collection
Private ex_ As CollectionEx
Private exVal_ As CollectionEx
Private exStr_ As CollectionEx
Private exEmpty_ As CollectionEx

Sub CreateTests()
    Call UnitTest.CreateRunTests("CollectionExTests")
End Sub

Sub TestInit()
    Dim cls1 As Class1, col1 As Collection
    Dim i As Long
    Set col1 = New Collection
    For i = 1 To 10
        Set cls1 = New Class1
        Call col1.Add(cls1.Init(i))
    Next
    
    Set colVal_ = New Collection
    For i = 1 To 10
        Call colVal_.Add(i)
    Next
    For i = 1 To 10
        Call colVal_.Add(i)
    Next
    
    Set ex_ = New CollectionEx
    Call ex_.Initialize(col1)
    Set exVal_ = New CollectionEx
    Call exVal_.Initialize(colVal_)
    
    Dim colStr As New Collection
    For i = 1 To 10
        Call colStr.Add("str" & i)
    Next
    Set exStr_ = New CollectionEx
    Call exStr_.Initialize(colStr)
    
    Dim colEmpty As New Collection
    Set exEmpty_ = New CollectionEx
    Call exEmpty_.Initialize(colEmpty)
End Sub

Sub RunTests()
   Dim test As New UnitTest

    test.RegisterTest "Test_Initialize_Create_Enum"
    test.RegisterTest "Test_AddRemoveCount"
    test.RegisterTest "Test_Where"
    test.RegisterTest "Test_SelectBy"
    test.RegisterTest "Test_SelectManyBy"
    test.RegisterTest "Test_FirstLast"
    test.RegisterTest "Test_AnyAndAlls"
    test.RegisterTest "Test_Contains"
    test.RegisterTest "Test_SumAverageMaxMin"
    test.RegisterTest "Test_MinMax_Value"
    test.RegisterTest "Test_MinMax_Object"
    test.RegisterTest "Test_MinByMaxBy_Object"
    test.RegisterTest "Test_Order"
    test.RegisterTest "Test_SkipTake"
    test.RegisterTest "Test_Distinct"
    test.RegisterTest "Test_ToArray"
    test.RegisterTest "Test_UserMemberID0"
    test.RegisterTest "Test_UnCollection"

    test.RunTests UnitTest
End Sub


'[Fact]
Sub Test_Initialize_Create_Enum()

    Dim cex As New CollectionEx
    Dim cls As New Class1
        
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    Call col.Add(cls.Create(4))
    Call col.Add(cls.Create(5))
    
    With UnitTest
        Call .NameOf("Initialize and Create and Enum")
        Call .AssertEqual("CollectionEx", TypeName(CollectionEx(col)))
        Call .AssertEqual(col.Count, CollectionEx(col).Count)
        Call .AssertEqual(5, CollectionEx(Array(1, 2, 3, 4, 5)).Items.Count)
        
        Dim v As Variant, added As Long
        For Each v In CollectionEx(col)
           added = added + v.abc
        Next
        Call .AssertHasNoError
        Call .AssertEqual(15, added)
        
        Call .NameOf("collection instance is not same (copied)")
        Call .AssertNotSame(CollectionEx(col).Items, col)
        
        Call .NameOf("Empty/Null, should return empty collection")
        Call .AssertEqual(0, CollectionEx(Empty).Count)
        Call .AssertEqual(0, CollectionEx(Null).Count)
        
        Call .NameOf("Nothing, Should raise error")
        On Error Resume Next
        Call .AssertEqual(0, CollectionEx(Nothing).Count)
        Call .AssertHasError
        Err.Clear
        On Error GoTo 0
        
    End With
End Sub

'[Fact]
Sub Test_AddRemoveCount()
    TestInit
    Dim recoll As Collection
    
    With UnitTest.NameOf("Add/AddRange")
        Set recoll = New Collection
        .AssertEqual 11, CollectionEx(recoll).Add(1).AddRange(GetClass1CollectionN(10)).Count
    End With
    With UnitTest.NameOf("Remove")
        Set recoll = New Collection
        .AssertEqual 0, CollectionEx(recoll).Add(1).Remove(1).Count
    End With
    
End Sub


'[Fact]
Sub Test_Where()
    TestInit
    Dim res As CollectionEx
    
    With UnitTest.NameOf("Where()")
        Set res = ex_.Where("x=>x.abc = 1")
        .AssertEqual 1, res.Items.Count
        .AssertEqual 1, res.Items(1).abc
    
        Set res = ex_.Where("x=>x.abc < 5")
        .AssertEqual 4, res.Count("")
        .AssertEqual 1, res.Items(1).abc
    
        Set res = ex_.Where("x=>3")
        .AssertEqual 0, res.Count
        
        On Error Resume Next
        Call exVal_.Where("x=>x.abc=1")
        .AssertHasError
        On Error GoTo 0
        
        .AssertEqual 3, ex_.Where("x=> x.abc<7 And x.abc>=4").Count
        
        Set res = exVal_.Where("x=>x=1")
        .AssertEqual 2, res.Count
        .AssertEqual 1, res.Items(2)
    
        Set res = ex_.Where("x=>x.def.def +  x.abc > 5 ")
        .AssertEqual 8, res.Count
        .AssertEqual 3, res.Items(1).abc
        .AssertEqual 11, res.Items(8).def.def
        
    End With
End Sub

'[Fact]
Sub Test_SelectBy()
    TestInit
    Dim res1, res2, res3
    
    With UnitTest.NameOf("Test_SelectBy()")
    
        Set res1 = ex_.SelectBy(" lm => lm.abc")
        Set res2 = ex_.SelectBy(" lm => lm.def")
        .AssertEqual 10, res1.Count
        .AssertEqual 1, res1.Items(1)
        .AssertEqual "Class2", TypeName(res2.Items(1))

        Set res3 = exVal_.SelectBy("lm => lm")
        .AssertEqual 10, res1.Count
        .AssertEqual 1, res1.Items(1)
        .AssertEqual "Class2", TypeName(res2.Items(1))
            
    End With
End Sub

'[Fact]
Sub Test_SelectManyBy()
    TestInit
    
    Dim res1, res2, res3
    Set res1 = ex_.SelectManyBy("x=>x.Defs")
    
    
    With UnitTest.NameOf("Test_SelectBy()")
        .AssertEqual 30, ex_.SelectManyBy("x=>x.Defs").Count
        .AssertEqual "Class2", TypeName(ex_.SelectManyBy("x=>x.Defs").Items(1))
    End With
End Sub


'[Fact]
Sub Test_FirstLast()
    With UnitTest.NameOf("Test First and Last")
        TestInit
        Dim cls As New Class1
        Dim col As New Collection
        
        .NameOf ("First")
        .AssertEqual 1, ex_.First().abc
        .AssertEqual 1, exVal_.First()
        On Error Resume Next
            exVal_.First ("x=>x > 100")
            .AssertHasError
        On Error GoTo 0
        
        .AssertEqual 1, ex_.FirstOrDefault().abc
        .AssertEqual 3, ex_.FirstOrDefault("x=>x.abc>2").abc
        .AssertNull ex_.Where("x=>x.abc>100").FirstOrDefault
        .AssertNothing ex_.FirstOrDefault("x=>x.abc>100", Nothing)
        .AssertNull exVal_.FirstOrDefault("x=>x > 100")
        On Error Resume Next
        Call CollectionEx(col).First
        Call .AssertHasError
        Call .AssertTrue(CollectionEx(col).FirstOrDefault(, Nothing) Is Nothing)
        On Error GoTo 0
        
        .NameOf ("Last")
        Call .AssertEqual(10, ex_.Last().abc)
        Call .AssertEqual(10, exVal_.Last())
        On Error Resume Next
            exVal_.Last ("x=>x > 100")
            .AssertHasError
        On Error GoTo 0
        Call .AssertEqual(10, ex_.LastOrDefault().abc)
        .AssertEqual 6, ex_.LastOrDefault("x=>x.abc<7").abc
        .AssertNull ex_.Where("x=>x.abc>100").LastOrDefault
        .AssertNothing ex_.LastOrDefault("x=>x.abc>100", Nothing)
        .AssertNull exVal_.LastOrDefault("x=>x>100")
        On Error Resume Next
        Call CollectionEx(col).Last
        Call .AssertHasError
        Call .AssertTrue(CollectionEx(col).LastOrDefault(, Nothing) Is Nothing)
        On Error GoTo 0
        
        
        .NameOf ("Single/SingleOrDefault")
        Set col = New Collection
        Call col.Add(cls.Create(1))
        Call col.Add(cls.Create(2))
        Call col.Add(cls.Create(2))
        Call col.Add(cls.Create(4))
        Call col.Add(cls.Create(5))
        
        Call .AssertEqual(cls.Create(5), CollectionEx(col).SingleBy("x=>x.abc = 5"))
        Call .AssertEqual(cls.Create(5), CollectionEx(col).SingleOrDefaultBy("x=>x.abc = 5"))
        Call .AssertTrue(CollectionEx(col).SingleOrDefaultBy("x=>x.abc = 1000", Nothing) Is Nothing)
        On Error Resume Next
        Call CollectionEx(col).SingleBy("x=>x.abc = 100")
        Call .AssertHasError
        Call CollectionEx(col).SingleBy("x=>x.abc = 2")
        Call .AssertHasError
        Call CollectionEx(col).SingleOrDefaultBy("x=>x.abc = 2")
        Call .AssertHasError
        Call CollectionEx(col).SingleBy
        Call .AssertHasError
        Call CollectionEx(col).SingleOrDefaultBy
        Call .AssertHasError
        
        Set col = New Collection
        Call CollectionEx(col).SingleBy
        Call .AssertHasError
        Call .AssertTrue(CollectionEx(col).SingleOrDefaultBy(, Nothing) Is Nothing)
        On Error GoTo 0
        
        Set col = New Collection
        Call col.Add(cls.Create(1))
        Call .AssertEqual(cls.Create(1), CollectionEx(col).SingleBy)
        Call .AssertEqual(cls.Create(1), CollectionEx(col).SingleOrDefaultBy)
    End With

End Sub

'[Fact]
Sub Test_AnyAndAlls()
    TestInit
    
    With UnitTest.NameOf("AnyAndAlls()")
        .NameOf ("Any")
        .AssertTrue ex_.AnyBy("x=>x.abc > 6")
        .AssertFalse ex_.AnyBy("x=>x.abc < 0")
        
        .NameOf ("All")
        .AssertTrue ex_.AllBy("x=>x.abc > 0")
        .AssertFalse ex_.AllBy("x=>x.abc > 1")
    End With

End Sub

'[Fact]
Sub Test_Contains()
    TestInit
    
    With UnitTest.NameOf("Contains()")
        .AssertTrue exVal_.Contains(6)
        .AssertFalse exVal_.Contains(11)
        
        .NameOf "for string"
        .AssertTrue exStr_.Contains("str1")
        .AssertFalse exStr_.Contains("str11")
        .AssertFalse exStr_.Contains("str")
    End With


    Dim cls As New Class1
    Dim col As New Collection
    Call col.Add(cls.Create(1))
    Call col.Add(cls.Create(2))
    Call col.Add(cls.Create(3))
    
    With UnitTest.NameOf("Contains")
        Call .AssertTrue(CollectionEx(col).Contains(cls.Create(2)))
        Call .AssertFalse(CollectionEx(col).Contains(cls.Create(7)))
    End With

End Sub

'[Fact]
Sub Test_SumAverageMaxMin()
    TestInit
    
    With UnitTest
        .NameOf "Sum"
        .AssertEqual 110, exVal_.Sum
        .AssertEqual 0, exEmpty_.Sum
        
        .NameOf "Sum by lambda"
        .AssertEqual 55, CollectionEx(GetClass1CollectionN(10)).Sum("x=>x.abc")
        
        .NameOf "Average"
        .AssertEqual 5.5, exVal_.Average
        .AssertNull exEmpty_.Average
    
        .NameOf ("Max")
        .AssertEqual 10, exVal_.Max()
        .AssertNull exVal_.Where("x=>x>100").Max
    
        .NameOf ("Min")
        .AssertEqual 1, exVal_.Min()
        .AssertNull exVal_.Where("x=>x < 0").Min
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
        Call .AssertEqual(1, CollectionEx(col).Min())
        Call .AssertEqual(5, CollectionEx(col).Max())
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
        Call .AssertEqual(1, CollectionEx(col).Min("x => x.abc"))
        Call .AssertEqual(5, CollectionEx(col).Max("x => x.abc"))
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
        Call .AssertTrue(TypeOf CollectionEx(col).MinBy("x=>x.abc") Is Class1)
        Call .AssertEqual(1, CollectionEx(col).MinBy("x=>x.abc").abc)
        Call .AssertTrue(TypeOf CollectionEx(col).MaxBy("x=>x.abc") Is Class1)
        Call .AssertEqual(5, CollectionEx(col).MaxBy("x=>x.abc").abc)
    End With
        
End Sub

'[Fact]
Sub Test_Order()
    TestInit
    
    Dim res As CollectionEx
    Dim col As New Collection
    Call col.Add(1)
    Call col.Add(2)
    Call col.Add(3)
    Call col.Add(4)
    Call col.Add(5)
    Call col.Add(2)
    Call col.Add(2)
    Call col.Add(3)
    
    With UnitTest.NameOf("Order of value")
        Set res = CollectionEx(col).Orderby("x=>x")
        Call .AssertTrue(res.Items(1) <= res.Items(2))
        Call .AssertTrue(res.Items(2) <= res.Items(3))
        Call .AssertTrue(res.Items(3) <= res.Items(4))
        Call .AssertTrue(res.Items(4) <= res.Items(5))
        Call .AssertTrue(res.Items(5) <= res.Items(6))
        Call .AssertTrue(res.Items(6) <= res.Items(7))
        Call .AssertTrue(res.Items(7) <= res.Items(8))
    End With
    
    With UnitTest.NameOf("Order Descending of value")
        Set res = res.OrderByDescending("x=>x")
        Call .AssertTrue(res.Items(1) >= res.Items(2))
        Call .AssertTrue(res.Items(2) >= res.Items(3))
        Call .AssertTrue(res.Items(3) >= res.Items(4))
        Call .AssertTrue(res.Items(4) >= res.Items(5))
        Call .AssertTrue(res.Items(5) >= res.Items(6))
        Call .AssertTrue(res.Items(6) >= res.Items(7))
        Call .AssertTrue(res.Items(7) >= res.Items(8))
    End With

    With UnitTest.NameOf("Order asc/desc by property")
        Set res = ex_.OrderByDescending("x=>x.abc").Orderby("x=>x.abc")
        Call .AssertEqual(ex_.Items(1), res.Items(1))
        Call .AssertEqual(ex_.Items(ex_.Count), res.Items(res.Count))
        
        Set res = ex_.OrderByDescending("x=>x.def.def").Orderby("x=>x.def.def")
        Call .AssertEqual(ex_.Items(1), res.Items(1))
        Call .AssertEqual(ex_.Items(ex_.Count), res.Items(res.Count))
    End With
    
    
    
    Dim re As Collection
    With UnitTest.NameOf("Order asc/desc, 2 elements order")
        Set col = New Collection
        Call col.Add(2)
        Call col.Add(1)
        Set re = CollectionEx(col).Orderby("x=>x").Items
        Call .AssertTrue(re(1) <= re(2))
    End With
    With UnitTest.NameOf("Order asc/desc, 1 elements return self")
        Set col = New Collection
        Call col.Add(1)
        Set re = CollectionEx(col).Orderby("x=>x").Items
        Call .AssertEqual(1, re.Count)
    End With
    With UnitTest.NameOf("Order asc/desc, 0 elements return empty collection")
        Set col = New Collection
        Set re = CollectionEx(col).Orderby("x=>x").Items
        Call .AssertEqual(0, re.Count)
    End With

        
End Sub

'[Fact]
Sub Test_SkipTake()
    With UnitTest.NameOf("Skip and Take")
        TestInit
        Dim res As CollectionEx
        
        .AssertEqual 3, ex_.Take(3).Count
        .AssertEqual 1, ex_.Take(4).Items(1).abc
        .AssertEqual 4, ex_.Take(4).Items(4).abc
        .AssertEqual 7, ex_.Skip(3).Count
        .AssertEqual 4, ex_.Skip(3).Items(1).abc
        .AssertEqual 10, ex_.Skip(3).Items(7).abc
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
        Set res = CollectionEx(col).Distinct().Items
        Call .AssertEqual(4, res.Count)
        Call .AssertEqual(1, CollectionEx(col).Distinct().Where("x=>x.abc = 1").Count)
        
        Call .NameOf("DistinctBy")
        Set res = CollectionEx(col).DistinctBy("x=>x.def").Items
        Call .AssertEqual(4, res.Count)
    End With
        
End Sub


'[Fact]
Sub Test_ToArray()
    With UnitTest.NameOf("ToArray()")
        TestInit
        Dim res() As Variant
        
        res = exVal_.ToArray()
        .AssertEqual 0, LBound(res)
        .AssertEqual 19, UBound(res)
        .AssertEqual exVal_.Items(1), res(LBound(res))
        .AssertEqual exVal_.Items(20), res(UBound(res))
    
    End With
End Sub

'[Fact]
Sub Test_UserMemberID0()
    TestInit
    With UnitTest.NameOf("UserMemberID=0, no decleared")
       .AssertEqual "CollectionEx", TypeName(CollectionEx(GetClass1CollectionN(10)))
       .AssertEqual 6, CollectionEx(GetClass1CollectionN(10)).Where("x=>x.abc>4").Count()
    End With
End Sub

'[Fact]
Sub Test_UnCollection()
    
    Dim ws As Worksheet
    Dim col, col2
    Set ws = CollectionEx(ThisWorkbook.Worksheets).Where("x=>x.Name = Sheet1").Items(1)
    Set col = CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10")).Where("x=>x.value>6").Items
    Set col2 = CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Where("x=>x>6").Items
    
    With UnitTest.NameOf("Enumlable, but un collection object, ")
        .AssertEqual "Sheet1", ws.Name
        .AssertEqual 7, col.Count
        .AssertIsType "Range", col.Item(1)
        .AssertEqual 7, col2.Count
        .AssertEqual 7, col2.Item(1)
        
        .AssertEqual "Sheet1", CollectionEx(ThisWorkbook.Worksheets).SelectBy("x=>x.Name").Items(1)
        .AssertEqual 3, CollectionEx(ThisWorkbook.Worksheets).Count
        .AssertEqual 1, CollectionEx(ThisWorkbook.Worksheets).Count("x=>x.Name = Sheet1")
        .AssertEqual 30, CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Count
        .AssertEqual "Sheet1", CollectionEx(ThisWorkbook.Worksheets).First.Name
        .AssertEqual "Sheet1", CollectionEx(ThisWorkbook.Worksheets).FirstOrDefault.Name
        .AssertEqual "Sheet3", CollectionEx(ThisWorkbook.Worksheets).Last.Name
        .AssertEqual "Sheet3", CollectionEx(ThisWorkbook.Worksheets).LastOrDefault.Name
        .AssertEqual "Sheet2", CollectionEx(ThisWorkbook.Worksheets).Skip(1).Items(1).Name
        .AssertEqual "Sheet2", CollectionEx(ThisWorkbook.Worksheets).Take(2).Items(2).Name
        .AssertFalse CollectionEx(ThisWorkbook.Worksheets).AllBy("x=>x.Name = Sheet2")
        .AssertTrue CollectionEx(ThisWorkbook.Worksheets).AnyBy("x=>x.Name = Sheet2")
        .AssertTrue CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Contains(10)
        .AssertFalse CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Contains(-1)
        .AssertEqual 93, CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Sum
        .AssertEqual 3.1, CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Average
        .AssertEqual 15, CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Max
        .AssertEqual 0, CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Min
        .AssertEqual "Sheet1", CollectionEx(ThisWorkbook.Worksheets).Orderby("x=>x.Index").Items(1).Name
        .AssertEqual "Sheet3", CollectionEx(ThisWorkbook.Worksheets).OrderByDescending("x=>x.Index").Items(1).Name
        .AssertEqual 1, CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).Orderby("x=>x").Items(18)
        .AssertEqual 15, CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).OrderByDescending("x=>x").Items(1)
        .AssertEqual "Sheet1", CollectionEx(ThisWorkbook.Worksheets).ToArray(0).Name
        .AssertEqual 1, CollectionEx(ThisWorkbook.Worksheets(1).Range("A1:C10").Value).ToArray(0)
    End With
    
End Sub


' Reference test
Sub Test_SpeedTests()
    Dim col As Collection, not_using_time As Double, using_time As Double
    ' Should be less than 100000 elements, by the avoiding the upper limit of garbage collection of "Collection".
    Set col = GetClass1CollectionN(100)
    
    With UnitTest
        Call .NameOf("Where method Time")
        not_using_time = Test_SpeedTest("Not using Where 1 layer", CollectionEx(col))
        using_time = Test_SpeedTest("Using Where 1 layer", CollectionEx(col))
        Call .AssertTrue(using_time < 0.1 Or not_using_time * 100 > using_time)
        
        not_using_time = Test_SpeedTest("Not using Where 2 layer", CollectionEx(col))
        using_time = Test_SpeedTest("Using Where 2 layer", CollectionEx(col))
        Call .AssertTrue(using_time < 0.1 Or not_using_time * 200 > using_time)

    End With
    
    Set col = Nothing
End Sub




Private Function Test_SpeedTest(test_name As String, cex As CollectionEx) As Double
    Dim col As New Collection, c As Class1
    Dim n:    n = (Timer)
    
    Select Case test_name
        Case "Not using Where 1 layer"
            For Each c In cex
                If c.abc = 2 Then Call col.Add(c)
            Next
        Case "Not using Where 2 layer"
            For Each c In cex
                If c.def.def = 2 Then Call col.Add(c)
            Next
        Case "Using Where 1 layer":       Call cex.Where("x=>x.Abc = 2")
        Case "Using Where 2 layer":       Call cex.Where("x=>x.Def.Def = 2")
    End Select
    
    
    Test_SpeedTest = ((Timer) - n)
    Debug.Print test_name & ":" & Format(Test_SpeedTest, "0.0000000") & "[s]"
End Function


Private Function GetClass1CollectionN(Optional n As Long = 10)
    Dim cls As New Class1, i As Long
    Dim col As New Collection
    
    For i = 1 To n
        Call col.Add(cls.Create(i))
    Next i
    Set GetClass1CollectionN = col
End Function

