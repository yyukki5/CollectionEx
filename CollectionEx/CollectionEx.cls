VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Expanding Collection's function."
Option Explicit

Private items_ As Collection

' Note --------------------------------------
' Project: https://github.com/yyukki5/CollectionEx
' Version: 1.0.0
' Depend on: Lambda.cls
' -------------------------------------------

Private Enum ErrNumber
    Unknown = 9100
    SequenceIsEmpty
    NoElementSatisfied
    NotSingleElement
    ElementsAreNotObject
    FailedToEvaluateEquals
    NotExpectedType
End Enum


Public Function Initialize(col As Collection) As CollectionEx
    Set items_ = col
    Set Initialize = Me
End Function

Public Function Create(enum_variant) As CollectionEx
Attribute Create.VB_Description = "Return new instance of CollectionEx"
Attribute Create.VB_UserMemId = 0
    Dim cex As CollectionEx, col As Collection, c
    Set col = New Collection
    Set cex = New CollectionEx

    If IsEmpty(enum_variant) Or IsNull(enum_variant) Then Set Create = cex.Initialize(col): Exit Function

    For Each c In enum_variant
        Call col.Add(c)
    Next

    Set Create = cex.Initialize(col)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = items_.[_NewEnum]
End Property

Public Property Get Items() As Collection
    Set Items = items_
End Property


Public Function Where(lambda_string As String) As CollectionEx
    Dim col As New Collection, recol As New CollectionEx, c
    
    For Each c In items_
        If Lambda(c, lambda_string).Evaluate = True Then Call col.Add(c)
    Next
    
    Set Where = recol.Initialize(col)
End Function

Public Function SelectBy(lambda_string As String) As CollectionEx
    Dim col As New Collection, recol As New CollectionEx, c
    
    For Each c In items_
        Call col.Add(Lambda(c, lambda_string).Evaluate)
    Next
    
    Set SelectBy = recol.Initialize(col)
End Function

Public Function SelectManyBy(lambda_string As String) As CollectionEx
    Dim c As Variant, c1 As Variant, col As New Collection, recol As New CollectionEx
    
    For Each c In items_
        For Each c1 In Lambda(c, lambda_string).Evaluate
            Call col.Add(c1)
        Next
    Next

    Set SelectManyBy = recol.Initialize(col)
End Function


Public Function Count(Optional lambda_string As String) As Long
    Dim col As New Collection, c
    
    If lambda_string = "" Then Count = items_.Count: Exit Function
    For Each c In items_
        If Lambda(c, lambda_string).Evaluate Then Call col.Add(c)
    Next
    
    Count = col.Count
End Function

Public Function Add(an_item, Optional a_key As String) As CollectionEx
    If a_key = "" Then Call items_.Add(an_item) Else Call items_.Add(an_item, a_key)
    Set Add = Me
End Function

Public Function AddRange(any_items As Collection) As CollectionEx
    Dim c
    For Each c In any_items
        Call items_.Add(c)
    Next
    Set AddRange = Me
End Function

Public Function Remove(index_or_key) As CollectionEx
    Call items_.Remove(index_or_key)
    Set Remove = Me
End Function

Public Function First(Optional lambda_string As String)
    Dim i As Long
    If items_.Count = 0 Then Call Err.Raise(ErrNumber.SequenceIsEmpty, "CollectionEx.First()", "The sequence is empty")
    If lambda_string = "" Then Call SetObjectAsBuffer(items_(1), First): Exit Function
    
    For i = 1 To items_.Count
        If Lambda(items_(i), lambda_string).Evaluate Then Call SetObjectAsBuffer(items_(i), First): Exit Function
    Next i
    
    Call Err.Raise(ErrNumber.NoElementSatisfied, "CollectionEx.First()", "No element satisfies the comparing condition.")
End Function
    
Public Function FirstOrDefault(Optional lambda_string As String, Optional default_value As Variant = Null)
    Dim i As Long
    If items_.Count = 0 Then Call SetObjectAsBuffer(default_value, FirstOrDefault): Exit Function
    If lambda_string = "" Then Call SetObjectAsBuffer(items_(1), FirstOrDefault): Exit Function
    
    For i = 1 To items_.Count
        If Lambda(items_(i), lambda_string).Evaluate Then Call SetObjectAsBuffer(items_(i), FirstOrDefault): Exit Function
    Next i
    
    Call SetObjectAsBuffer(default_value, FirstOrDefault)
End Function

Public Function Last(Optional lambda_string As String)
    Dim i As Long
    If items_.Count = 0 Then Call Err.Raise(ErrNumber.SequenceIsEmpty, "CollectionEx.Last()", "The sequence is empty")
    If lambda_string = "" Then Call SetObjectAsBuffer(items_(items_.Count), Last): Exit Function
    
    For i = items_.Count To 1 Step -1
        If Lambda(items_(i), lambda_string).Evaluate Then Call SetObjectAsBuffer(items_(i), Last): Exit Function
    Next i
    
    Call Err.Raise(ErrNumber.NoElementSatisfied, "CollectionEx.Last()", "No element satisfies the comparing condition.")
End Function

Public Function LastOrDefault(Optional lambda_string As String, Optional default_value As Variant = Null)
    Dim i As Long
    If items_.Count = 0 Then Call SetObjectAsBuffer(default_value, LastOrDefault): Exit Function
    If lambda_string = "" Then Call SetObjectAsBuffer(items_(items_.Count), LastOrDefault): Exit Function
    
    For i = items_.Count To 1 Step -1
        If Lambda(items_(i), lambda_string).Evaluate Then Call SetObjectAsBuffer(items_(i), LastOrDefault): Exit Function
    Next i
    
    Call SetObjectAsBuffer(default_value, LastOrDefault)
End Function

Public Function SingleBy(Optional lambda_string As String)
    Dim i As Long, f As Boolean
    If lambda_string = "" Then
        If items_.Count = 0 Then
            Call Err.Raise(ErrNumber.SequenceIsEmpty, "CollectionEx.SingleBy()", "The sequence is empty")
        ElseIf items_.Count = 1 Then
            Call SetObjectAsBuffer(items_(1), SingleBy): Exit Function
        Else
            Call Err.Raise(ErrNumber.NotSingleElement, "CollectionEx.SingleBy()", "The input sequence contains more than one element.")
        End If
    End If
    
    For i = 1 To items_.Count
        If Lambda(items_(i), lambda_string).Evaluate Then
            If f = False Then f = True Else Call Err.Raise(ErrNumber.NotSingleElement, "CollectionEx.SingleBy()", "The input sequence contains more than one element.")
            Call SetObjectAsBuffer(items_(i), SingleBy)
        End If
    Next i
        
    If f = False Then Call Err.Raise(ErrNumber.NoElementSatisfied, "CollectionEx.SingleBy()", "No element satisfies the comparing condition.")
End Function

Public Function SingleOrDefaultBy(Optional lambda_string As String, Optional default_value As Variant = Null)
    Dim i As Long, f As Boolean
    If lambda_string = "" Then
        If items_.Count = 0 Then
            Call SetObjectAsBuffer(default_value, SingleOrDefaultBy): Exit Function
        ElseIf items_.Count = 1 Then
            Call SetObjectAsBuffer(items_(1), SingleOrDefaultBy): Exit Function
        Else
            Call Err.Raise(ErrNumber.NotSingleElement, "CollectionEx.SingleOrDefaultBy()", "The input sequence contains more than one element.")
        End If
    End If
    
    For i = 1 To items_.Count
        If Lambda(items_(i), lambda_string).Evaluate Then
            If f = False Then f = True Else Call Err.Raise(ErrNumber.NotSingleElement, "CollectionEx.SingleOrDefaultBy()", "The input sequence contains more than one element.")
            Call SetObjectAsBuffer(items_(i), SingleOrDefaultBy)
        End If
    Next i
        
    If f = False Then Call SetObjectAsBuffer(default_value, SingleOrDefaultBy)
End Function


Public Function Skip(skip_index As Long) As CollectionEx
    Dim col As New Collection, c, i As Long
    For Each c In items_
        i = i + 1
        If i > skip_index Then Call col.Add(c)
    Next
    Set Skip = CollectionEx(col)
End Function

Public Function Take(take_index As Long) As CollectionEx
    Dim col As New Collection, c, i As Long
    For Each c In items_
        i = i + 1
        If i <= take_index Then Call col.Add(c)
    Next
    Set Take = CollectionEx(col)
End Function

Public Function AllBy(lambda_string As String) As Boolean
    Dim col As New Collection, c
    AllBy = False
    For Each c In items_
         If Lambda(c, lambda_string).Evaluate = False Then Exit Function
    Next
    AllBy = True
End Function

Public Function AnyBy(lambda_string As String) As Boolean
    Dim col As New Collection, c
    AnyBy = True
    For Each c In items_
         If Lambda(c, lambda_string).Evaluate = True Then Exit Function
    Next
    AnyBy = False
End Function

Public Function Contains(var) As Boolean
    Dim c
    
    For Each c In items_
        If Equals(var, c) Then Contains = True: Exit Function
    Next
    
    Contains = False
End Function

Public Function Sum(Optional lambda_string As String) As Double
    Dim c, v, val As Double
    
    If items_.Count = 0 Then Sum = 0: Exit Function
    
    If lambda_string = "" Then
        For Each c In items_
            If IsNumeric(c) = False Then Call Err.Raise(ErrNumber.NotExpectedType, "CollectionEx.Sum()", "Element should be numeric value")
            val = val + c
        Next
    Else
        For Each c In items_
            v = Lambda(c, lambda_string).Evaluate
            If IsNumeric(v) = False Then Call Err.Raise(ErrNumber.NotExpectedType, "CollectionEx.Sum()", "Element should be numeric value")
            val = val + v
        Next
    End If
    
    Sum = val
End Function

Public Function Average() As Variant
    If items_.Count = 0 Then Average = Null: Exit Function

    Average = Sum / items_.Count
End Function

Public Function Max(Optional lambda_string As String) As Variant
    Dim v As Variant, tmp As Variant, i As Long
    
    If items_.Count = 0 Then Max = Null: Exit Function
    
    If lambda_string = "" Then
        tmp = items_(1)
        For i = 2 To items_.Count
            If items_(i) > tmp Then tmp = items_(i)
        Next
    Else
        tmp = Lambda(items_(1), lambda_string).Evaluate
        For i = 2 To items_.Count
            v = Lambda(items_(i), lambda_string).Evaluate
            If v > tmp Then tmp = v
        Next
    End If
    
    Max = tmp
End Function

Public Function Min(Optional lambda_string As String) As Variant
    Dim v As Variant, tmp As Variant, i As Long
    
    If items_.Count = 0 Then Min = Null: Exit Function
    
    If lambda_string = "" Then
        tmp = items_(1)
        For i = 2 To items_.Count
            If items_(i) < tmp Then tmp = items_(i)
        Next
    Else
        tmp = Lambda(items_(1), lambda_string).Evaluate
        For i = 2 To items_.Count
            v = Lambda(items_(i), lambda_string).Evaluate
            If v < tmp Then tmp = v
        Next
    End If
    
    Min = tmp
End Function

Public Function MaxBy(lambda_string As String) As Object
    Dim v As Variant, temp_value As Variant, temp_item As Variant, i As Long
    
    If items_.Count = 0 Then Set MaxBy = Nothing: Exit Function
    If IsObject(items_(1)) = False Then Call Err.Raise(ErrNumber.ElementsAreNotObject, "CollectionEx.MaxBy()", "Elements should be object.")
        
    temp_value = Lambda(items_(1), lambda_string).Evaluate
    Set temp_item = items_(1)
    
    For i = 2 To items_.Count
        v = Lambda(items_(i), lambda_string).Evaluate
        If v > temp_value Then
            temp_value = v
            Set temp_item = items_(i)
        End If
    Next
    
    Set MaxBy = temp_item
End Function

Public Function MinBy(lambda_string As String) As Object
    Dim v As Variant, temp_value As Variant, temp_item As Variant, i As Long
    
    If items_.Count = 0 Then Set MinBy = Nothing: Exit Function
    If IsObject(items_(1)) = False Then Call Err.Raise(ErrNumber.ElementsAreNotObject, "CollectionEx.MinBy()", "Elements should be object.")
        
    temp_value = Lambda(items_(1), lambda_string).Evaluate
    Set temp_item = items_(1)
    
    For i = 2 To items_.Count
        v = Lambda(items_(i), lambda_string).Evaluate
        If v < temp_value Then
            temp_value = v
            Set temp_item = items_(i)
        End If
    Next
    
    Set MinBy = temp_item
End Function


Public Function Orderby(lambda_string As String) As CollectionEx
    Dim col As New Collection, p As New Collection, c, recol As New CollectionEx, sort_col As New Collection
    
    For Each c In items_
        Set p = New Collection
        Call p.Add(Lambda(c, lambda_string).Evaluate, "Value")
        Call p.Add(c, "Item")
        Call col.Add(p)
    Next
    For Each c In QuickSort(col, 1, items_.Count, True)
        Call sort_col.Add(c.Item("Item"))
    Next
    
    Set Orderby = recol.Initialize(sort_col)
End Function

Public Function OrderByDescending(lambda_string As String) As CollectionEx
    Dim col As New Collection, p As New Collection, c, recol As New CollectionEx, sort_col As New Collection
    
    For Each c In items_
        Set p = New Collection
        Call p.Add(Lambda(c, lambda_string).Evaluate, "Value")
        Call p.Add(c, "Item")
        Call col.Add(p)
    Next
    For Each c In QuickSort(col, 1, items_.Count, False)
        Call sort_col.Add(c.Item("Item"))
    Next
    
    Set OrderByDescending = recol.Initialize(sort_col)
End Function


Public Function Distinct() As CollectionEx
    Dim c As Variant, compare_item As Variant, col As Collection, recol As New CollectionEx, f As Boolean
    Set col = New Collection
    
    For Each c In items_
        f = False
        For Each compare_item In col
            If Equals(compare_item, c) Then f = True: Exit For
        Next
        If f = False Then Call col.Add(c)
    Next
    
    Set Distinct = recol.Initialize(col)
End Function

Public Function DistinctBy(lambda_string As String) As CollectionEx
    Dim c As Variant, compare_item As Variant, compare_items As Collection, col As Collection, recol As New CollectionEx, f As Boolean, v As Variant
    Set col = New Collection: Set compare_items = New Collection
    
    For Each c In items_
        f = False
        Call SetObjectAsBuffer(Lambda(c, lambda_string).Evaluate, v)
        For Each compare_item In compare_items
            If Equals(compare_item, v) Then f = True: Exit For
        Next
        If f = False Then
            Call col.Add(c)
            Call compare_items.Add(v)
        End If
    Next
    
    Set DistinctBy = recol.Initialize(col)
End Function

Public Function ToArray()
    Dim arr, i As Long
    
    If items_.Count = 0 Then ToArray = Array(): Exit Function
    ReDim arr(0 To items_.Count - 1)
    
    For i = LBound(arr) To UBound(arr)
        Call SetObjectAsBuffer(Items.Item(i + 1), arr(i))
    Next i
    ToArray = arr
End Function


' private functions --------------------------------
Private Function QuickSort(ByRef any_pairs As Collection, ByVal low_boundary As Long, ByVal up_boundary As Long, is_asc As Boolean) As Collection
    Dim i As Long, j As Long, pivot As Variant, temp, property_any_pairs As New Collection
        
    If any_pairs.Count <= 1 Then Set QuickSort = any_pairs: Exit Function
    
    i = low_boundary
    j = up_boundary
    pivot = any_pairs((low_boundary + up_boundary) \ 2).Item("Value")
    
    Do While i <= j
        If is_asc Then
            Do While any_pairs(i).Item("Value") < pivot: i = i + 1: If i > up_boundary Then Exit Do:
            Loop
            Do While any_pairs(j).Item("Value") > pivot: j = j - 1: If j < low_boundary Then Exit Do:
            Loop
        Else
            Do While any_pairs(i).Item("Value") > pivot: i = i + 1: If i > up_boundary Then Exit Do:
            Loop
            Do While any_pairs(j).Item("Value") < pivot: j = j - 1: If j < low_boundary Then Exit Do:
            Loop
        End If
        If i <= j Then
            Set temp = any_pairs(i)
            Call any_pairs.Add(any_pairs(j), Before:=i)
            Call any_pairs.Remove(i + 1)
            Call any_pairs.Add(temp, Before:=j)
            Call any_pairs.Remove(j + 1)
            i = i + 1
            j = j - 1
        End If
    Loop

    If low_boundary < j Then Call QuickSort(any_pairs, low_boundary, j, is_asc)
    If i < up_boundary Then Call QuickSort(any_pairs, i, up_boundary, is_asc)
    
    Set QuickSort = any_pairs
End Function

Private Function Equals(comparing_item, compared_item) As Boolean
    If IsObject(comparing_item) Then
        On Error GoTo errUnImplementEquals
        ' Should the "compare_item" has Equals() function. if not has, raise error
        If comparing_item.Equals((compared_item)) Then Equals = True
    Else
        If comparing_item = compared_item Then Equals = True
    End If
    Exit Function
    
errUnImplementEquals:
    Call Err.Raise(ErrNumber.FailedToEvaluateEquals, "CollectionEx.Equals()", "Equality evaluation failed. Should """ & TypeName(comparing_item) & """ object has ""Equals()"" function.")
End Function

Private Function SetObjectAsBuffer(ByRef unverified_value, ByRef buffer_value) As Boolean
    If IsObject(unverified_value) Then
        Set buffer_value = unverified_value
        SetObjectAsBuffer = True
    Else
        Let buffer_value = unverified_value
        SetObjectAsBuffer = False
    End If
End Function
