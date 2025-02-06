VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Note --------------------------------------
' Project: https://github.com/yyukki5/CollectionEx
' Version: 0.5.0 (developing)
' Specification may be changed.
' -------------------------------------------

Private items_ As Collection

Public Enum cexOperators
    cexEqual
    cexDoesNotEqual
    cexIs
    cexIsNot
    cexGreaterThan
    cexGreaterThanOrEqualTo
    cexLessThan
    cexLessThanOrEqualTo
End Enum


Private Sub Class_Initialize()
    Set items_ = New Collection
End Sub

Public Function Initialize(col As Collection) As ColEx
    Set items_ = col
    Set Initialize = Me
End Function

Public Function Create(enum_variant) As ColEx
Attribute Create.VB_UserMemId = 0
    Dim col As Collection, c As Variant
    
    Set col = New Collection
    For Each c In enum_variant
        Call col.Add(c)
    Next
    
    Dim cex As New ColEx
    Set Create = cex.Initialize(col)
End Function


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = items_.[_NewEnum]
End Property
Public Property Get Items() As Collection
    Set Items = items_
End Property

Public Property Get Count() As Long
    Count = items_.Count
End Property



Public Function Add(class_item As Variant) As ColEx
    Call items_.Add(class_item)
    Set Add = Me
End Function
Public Function AddRange(col As Collection) As ColEx
    Dim c
    For Each c In col
        items_.Add c
    Next
    Set AddRange = Me
End Function
Public Function Remove(item_index As Long) As ColEx
    Call items_.Remove(item_index)
    Set Remove = Me
End Function


Public Function Where(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant) As ColEx
    Dim c As Variant, col As Collection, recol As New ColEx
    Set col = New Collection
    
    For Each c In items_
        If Compare2(c, property_name, compare_operator, compare_value) Then Call col.Add(c)
    Next
    
    Set Where = recol.Initialize(col)
End Function


Public Function SelectBy(property_name As String, Optional call_type As VbCallType = VbGet, Optional any_parameters) As ColEx
    Dim c As Variant, col As Collection, recol As New ColEx
    Set col = New Collection
    
    For Each c In items_
        If IsMissing(any_parameters) Then
            Call col.Add(CallByName(c, property_name, call_type))
        Else
            Call col.Add(CallByName(c, property_name, call_type, any_parameters))
        End If
    Next
    
    Set SelectBy = recol.Initialize(col)
End Function

Public Function SelectManyBy(property_name As String) As ColEx
    Dim c As Variant, c1 As Variant, col As Collection, recol As New ColEx
    Set col = New Collection
    
    For Each c In items_
        For Each c1 In CallByName(c, property_name, VbGet)
            Call col.Add(c1)
        Next
    Next

    Set SelectManyBy = recol.Initialize(col)
End Function


Public Function First(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant)
    Dim i As Long
    If items_.Count = 0 Then Call Err.Raise(9001, "ColEx.First()", "The sequence is empty")
    If IsMissing(compare_value) Then Call SetObjectAsBuffer(items_(1), First): Exit Function
    
    For i = 1 To items_.Count
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then Call SetObjectAsBuffer(items_(i), First): Exit Function
    Next i
    
    Call Err.Raise(9002, "ColEx.First()", "No element satisfies the comparing condition.")
End Function

Public Function FirstOrDefault(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant, _
                      Optional default_value As Variant = Null)
    Dim i As Long
    If items_.Count = 0 Then Call SetObjectAsBuffer(default_value, FirstOrDefault): Exit Function
    If IsMissing(compare_value) Then Call SetObjectAsBuffer(items_(1), FirstOrDefault): Exit Function
    
    For i = 1 To items_.Count
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then Call SetObjectAsBuffer(items_(i), FirstOrDefault): Exit Function
    Next i
    
    Call SetObjectAsBuffer(default_value, FirstOrDefault)
End Function

Public Function Last(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant)
    Dim i As Long
    If items_.Count = 0 Then Call Err.Raise(9001, "ColEx.Last()", "The sequence is empty")
    If IsMissing(compare_value) Then Call SetObjectAsBuffer(items_(items_.Count), Last): Exit Function
        
    For i = items_.Count To 1 Step -1
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then Call SetObjectAsBuffer(items_(i), Last): Exit Function
    Next i
    
    Call Err.Raise(9002, "ColEx.Last()", "No element satisfies the comparing condition.")
End Function

Public Function LastOrDefault(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant, _
                      Optional default_value As Variant = Null)
    Dim i As Long
    If items_.Count = 0 Then Call SetObjectAsBuffer(default_value, LastOrDefault): Exit Function
    If IsMissing(compare_value) Then Call SetObjectAsBuffer(items_(items_.Count), LastOrDefault): Exit Function
        
    For i = items_.Count To 1 Step -1
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then Call SetObjectAsBuffer(items_(i), LastOrDefault): Exit Function
    Next i
    
    Call SetObjectAsBuffer(default_value, LastOrDefault)
End Function

Public Function SingleBy(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant)
    Dim i As Long, f As Boolean
    If IsMissing(compare_value) Then
        If items_.Count = 0 Then
            Call Err.Raise(9001, "ColEx.SingleBy()", "The sequence is empty")
        ElseIf items_.Count = 1 Then
            Call SetObjectAsBuffer(items_(1), SingleBy): Exit Function
        Else
            Call Err.Raise(9003, "ColEx.SingleBy()", "The input sequence contains more than one element.")
        End If
    End If
    
    For i = 1 To items_.Count
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then
            If f = False Then f = True Else Call Err.Raise(9002, "ColEx.SingleBy()", "The input sequence contains more than one element.")
            Call SetObjectAsBuffer(items_(i), SingleBy)
        End If
    Next i
        
    If f = False Then Call Err.Raise(9002, "ColEx.SingleBy()", "No element satisfies the comparing condition.")
End Function

Public Function SingleOrDefaultBy(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant, _
                      Optional default_value As Variant = Null)
    Dim i As Long, f As Boolean
    If IsMissing(compare_value) Then
        If items_.Count = 0 Then
            Call SetObjectAsBuffer(default_value, SingleOrDefaultBy): Exit Function
        ElseIf items_.Count = 1 Then
            Call SetObjectAsBuffer(items_(1), SingleOrDefaultBy): Exit Function
        Else
            Call Err.Raise(9003, "ColEx.SingleOrDefaultBy()", "The input sequence contains more than one element.")
        End If
    End If
    
    For i = 1 To items_.Count
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then
            If f = False Then f = True Else Call Err.Raise(9003, "ColEx.SingleOrDefaultBy()", "The input sequence contains more than one element.")
            Call SetObjectAsBuffer(items_(i), SingleOrDefaultBy)
        End If
    Next i
        
    If f = False Then Call SetObjectAsBuffer(default_value, SingleOrDefaultBy)
End Function


Public Function Take(cnt As Long) As ColEx
    Dim c As Variant, col As New Collection, i As Long
    
    For Each c In items_
        i = i + 1
        If i <= cnt Then Call col.Add(c)
    Next
    
    Set Take = ColEx(col)
End Function

Public Function Skip(cnt As Long) As ColEx
    Dim c As Variant, col As New Collection, i As Long
    
    For Each c In items_
        i = i + 1
        If i > cnt Then Call col.Add(c)
    Next
    
    Set Skip = ColEx(col)
End Function


Public Function AllBy(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant) As Boolean
    Dim c
    
    AllBy = False
    For Each c In items_
         If Compare2(c, property_name, compare_operator, compare_value) = False Then Exit Function
    Next
    
    AllBy = True
End Function

Public Function AnyBy(Optional property_name As String, _
                      Optional compare_operator As cexOperators, _
                      Optional compare_value As Variant) As Boolean
    Dim c
    
    AnyBy = True
    For Each c In items_
         If Compare2(c, property_name, compare_operator, compare_value) = True Then Exit Function
    Next
    
    AnyBy = False
End Function

Public Function Contains(compare_item As Variant) As Boolean
    Dim c
    
    For Each c In items_
        If Equals(compare_item, c) Then Contains = True: Exit Function
    Next
    
    Contains = False
End Function


Public Function OrderBy(Optional property_name As String) As ColEx
    Dim col As New Collection, p As New Collection, c, sort_col As New Collection
    
    For Each c In items_
        Set p = New Collection
        If property_name = "" Then Call p.Add(c, "Value") Else Call p.Add(CallByName(c, property_name, VbGet), "Value")
        Call p.Add(c, "Item")
        Call col.Add(p)
    Next
    For Each c In QuickSort(col, 1, items_.Count, True)
        Call sort_col.Add(c.Item("Item"))
    Next

    Set OrderBy = New ColEx
    Call OrderBy.Initialize(sort_col)
End Function

Public Function OrderByDescending(Optional property_name As String) As ColEx
    Dim col As New Collection, p As New Collection, c, sort_col As New Collection
    
    For Each c In items_
        Set p = New Collection
        If property_name = "" Then Call p.Add(c, "Value") Else Call p.Add(CallByName(c, property_name, VbGet), "Value")
        Call p.Add(c, "Item")
        Call col.Add(p)
    Next
    For Each c In QuickSort(col, 1, items_.Count, False)
        Call sort_col.Add(c.Item("Item"))
    Next
    
    Set OrderByDescending = New ColEx
    Call OrderByDescending.Initialize(sort_col)
End Function

Public Function Distinct() As ColEx
    Dim c As Variant, compare_item As Variant, col As Collection, recol As New ColEx, f As Boolean
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

Public Function DistinctBy(property_name As String) As ColEx
    Dim c As Variant, compare_item As Variant, compare_items As Collection, col As Collection, recol As New ColEx, f As Boolean
    Set col = New Collection: Set compare_items = New Collection
    
    For Each c In items_
        f = False
        For Each compare_item In compare_items
            If Equals(compare_item, CallByName(c, property_name, VbGet)) Then f = True: Exit For
        Next
        If f = False Then
            Call col.Add(c)
            Call compare_items.Add(CallByName(c, property_name, VbGet))
        End If
    Next
    
    Set DistinctBy = recol.Initialize(col)
End Function

Public Function Max(Optional property_name As String) As Variant
    Dim v As Variant, tmp As Variant, i As Long
    
    If items_.Count = 0 Then Max = Null: Exit Function
    
    If property_name = "" Then
        tmp = items_(1)
        For i = 2 To items_.Count
            If items_(i) > tmp Then tmp = items_(i)
        Next
    Else
        tmp = CallByName(items_(1), property_name, VbGet)
        For i = 2 To items_.Count
            v = CallByName(items_(i), property_name, VbGet)
            If v > tmp Then tmp = v
        Next
    End If
    
    Max = tmp
End Function

Public Function Min(Optional property_name As String) As Variant
    Dim v As Variant, tmp As Variant, i As Long
    
    If items_.Count = 0 Then Min = Null: Exit Function
    
    If property_name = "" Then
        tmp = items_(1)
        For i = 2 To items_.Count
            If items_(i) < tmp Then tmp = items_(i)
        Next
    Else
        tmp = CallByName(items_(1), property_name, VbGet)
        For i = 2 To items_.Count
            v = CallByName(items_(i), property_name, VbGet)
            If v < tmp Then tmp = v
        Next
    End If
    
    Min = tmp
End Function

Public Function MaxBy(property_name As String) As Variant
    Dim v As Variant, temp_value As Variant, temp_item As Variant, i As Long
    
    If items_.Count = 0 Then Set MaxBy = Nothing: Exit Function
    If IsObject(items_(1)) = False Then Call Err.Raise(9001, "ColEx.MaxBy()", "Elements should be object.")
    
    temp_value = CallByName(items_(1), property_name, VbGet)
    Set temp_item = items_(1)
    
    For i = 2 To items_.Count
        v = CallByName(items_(i), property_name, VbGet)
        If v > temp_value Then
            temp_value = v
            Set temp_item = items_(i)
        End If
    Next
    
    Set MaxBy = temp_item
End Function

Public Function MinBy(property_name As String) As Variant
    Dim v As Variant, temp_value As Variant, temp_item As Variant, i As Long
    
    If items_.Count = 0 Then Set MinBy = Nothing: Exit Function
    If IsObject(items_(1)) = False Then Call Err.Raise(9001, "ColEx.MinBy()", "Elements should be object.")
        
    temp_value = CallByName(items_(1), property_name, VbGet)
    Set temp_item = items_(1)
    
    For i = 2 To items_.Count
        v = CallByName(items_(i), property_name, VbGet)
        If v < temp_value Then
            temp_value = v
            Set temp_item = items_(i)
        End If
    Next
    
    Set MinBy = temp_item
End Function

Public Function ToArray()
    Dim arr, i As Long
    ReDim arr(0 To items_.Count - 1)
    
    For i = LBound(arr) To UBound(arr)
        Call SetObjectAsBuffer(Items.Item(i + 1), arr(i))
    Next i
    
    ToArray = arr
End Function

Public Function ToArray2D(ParamArray property_names())
    Dim arr, i As Long, j As Long, cnt As Long, p As Variant, ps As Variant, l As Long, u As Long
    
    ps = property_names
    cnt = UBound(property_names) - LBound(property_names)
    
    If UBound(property_names()) = 0 Then
        If IsArray(property_names(0)) Then
            ps = property_names(0)
            cnt = UBound(property_names(0)) - LBound(property_names(0))
        ElseIf TypeOf property_names(0) Is Collection Then
            Set ps = property_names(0)
            cnt = ps.Count - 1
        End If
    End If
        
    ReDim arr(0 To items_.Count - 1, 0 To cnt)
    For i = LBound(arr, 1) To UBound(arr, 1)
        j = 0
        For Each p In ps
            Call SetObjectAsBuffer(CallByName(items_.Item(i + 1), p, VbGet), arr(i, j))
            j = j + 1
        Next
    Next i
    
    ToArray2D = arr
End Function


' private functions --------------------------------
Private Function Compare(compared_value As Variant, compare_operator As cexOperators, Optional compare_value As Variant = Null)
    If IsNull(compare_value) Then
    
    Else
        Select Case compare_operator
            Case cexOperators.cexEqual
                Compare = Equals(compare_value, compared_value)
            Case cexOperators.cexDoesNotEqual
                Compare = (Equals(compare_value, compared_value) = False)
            Case cexOperators.cexGreaterThan
                Compare = (compared_value > compare_value)
            Case cexOperators.cexGreaterThanOrEqualTo
                Compare = (compared_value >= compare_value)
            Case cexOperators.cexLessThan
                Compare = (compared_value < compare_value)
            Case cexOperators.cexLessThanOrEqualTo
                Compare = (compared_value <= compare_value)
            
            Case cexOperators.cexIs
                Compare = (compared_value Is compare_value)
            Case cexOperators.cexIsNot
                Compare = (compared_value Is Not compare_value)
        End Select
    End If
End Function

Private Function Compare2(c As Variant, property_name As String, compare_operator As cexOperators, compare_value As Variant) As Boolean
    If property_name = "" Then
        If Compare(c, compare_operator, compare_value) Then Compare2 = True
    Else
        If Compare(CallByName(c, property_name, VbGet), compare_operator, compare_value) Then Compare2 = True
    End If
End Function

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
    Call Err.Raise(9001, "ColEx.Equals()", "Equality evaluation failed. Should """ & TypeName(comparing_item) & """ object has ""Equals()"" function.")
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
