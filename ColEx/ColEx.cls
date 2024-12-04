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
' Version: 0.2.0 (developing)
' Specification may be changed.
' -------------------------------------------

Private items_ As Collection

Public Enum Operators
    cexEqual
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
                      Optional compare_operator As Operators, _
                      Optional compare_value As Variant) As ColEx
    Dim c As Variant, col As Collection, recol As New ColEx
    Set col = New Collection
    
    For Each c In items_
        If Compare2(c, property_name, compare_operator, compare_value) Then Call col.Add(c)
    Next
    
    Set Where = recol.Initialize(col)
End Function


Public Function SelectBy(property_name As String) As ColEx
    Dim c As Variant, col As Collection, recol As New ColEx
    Set col = New Collection
    
    For Each c In items_
        Call col.Add(CallByName(c, property_name, VbGet))
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
                      Optional compare_operator As Operators, _
                      Optional compare_value As Variant)
    Dim i As Long
    If items_.Count > 0 And IsMissing(compare_value) Then Call SetObjectAsBuffer(items_(1), First): Exit Function
    
    For i = 1 To items_.Count
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then Call SetObjectAsBuffer(items_(i), First): Exit Function
    Next i
    
    Call Err.Raise(9001, "ColEx.First()", "no items")
End Function

Public Function FirstOrDefault(Optional property_name As String, _
                      Optional compare_operator As Operators, _
                      Optional compare_value As Variant, _
                      Optional default_value As Variant = Null)
    Dim i As Long
    If items_.Count > 0 And IsMissing(compare_value) Then Call SetObjectAsBuffer(items_(1), FirstOrDefault): Exit Function
    
    For i = 1 To items_.Count
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then Call SetObjectAsBuffer(items_(i), FirstOrDefault): Exit Function
    Next i
    
    Call SetObjectAsBuffer(default_value, FirstOrDefault)
End Function

Public Function Last(Optional property_name As String, _
                      Optional compare_operator As Operators, _
                      Optional compare_value As Variant)
    Dim i As Long
    If items_.Count > 0 And IsMissing(compare_value) Then Call SetObjectAsBuffer(items_(items_.Count), Last): Exit Function
        
    For i = items_.Count To 1 Step -1
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then Call SetObjectAsBuffer(items_(i), Last): Exit Function
    Next i
    
    Call Err.Raise(9001, "ColEx.Last()", "no items")
End Function

Public Function LastOrDefault(Optional property_name As String, _
                      Optional compare_operator As Operators, _
                      Optional compare_value As Variant, _
                      Optional default_value As Variant = Null)
    Dim i As Long
    If items_.Count > 0 And IsMissing(compare_value) Then Call SetObjectAsBuffer(items_(items_.Count), LastOrDefault): Exit Function
        
    For i = items_.Count To 1 Step -1
        If Compare2(items_(i), property_name, compare_operator, compare_value) Then Call SetObjectAsBuffer(items_(i), LastOrDefault): Exit Function
    Next i
    
    Call SetObjectAsBuffer(default_value, LastOrDefault)
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
                      Optional compare_operator As Operators, _
                      Optional compare_value As Variant) As Boolean
    Dim c
    
    AllBy = False
    For Each c In items_
         If Compare2(c, property_name, compare_operator, compare_value) = False Then Exit Function
    Next
    
    AllBy = True
End Function

Public Function AnyBy(Optional property_name As String, _
                      Optional compare_operator As Operators, _
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
        If IsObject(compare_item) Then
            ' Should the "compare_item" has Equals() function. if not has, raise error
            If compare_item.Equals(c) Then Contains = True: Exit Function
        Else
            If compare_item = c Then Contains = True: Exit Function
        End If
    Next
    
    Contains = False
End Function


Public Function OrderBy(property_name As String) As ColEx
    Set OrderBy = New ColEx
    Call OrderBy.Initialize(QuickSort(items_, 1, items_.Count, True, property_name))
End Function

Public Function OrderByDescending(property_name As String) As ColEx
    Set OrderByDescending = New ColEx
    Call OrderByDescending.Initialize(QuickSort(items_, 1, items_.Count, False, property_name))
End Function


Public Function ToArray()
    Dim arr, i As Long
    ReDim arr(0 To items_.Count - 1)
    
    For i = LBound(arr) To UBound(arr)
        Call SetObjectAsBuffer(Items.Item(i + 1), arr(i))
    Next i
    
    ToArray = arr
End Function


' private functions --------------------------------
Private Function Compare(compared_value As Variant, compare_operator As Operators, Optional compare_value As Variant = Null)
    If IsNull(compare_value) Then
    
    Else
        Select Case compare_operator
            Case Operators.cexEqual
                If IsObject(compared_value) Then
                    Compare = compared_value.Equals((compare_value))
                Else
                    Compare = (compared_value = compare_value)
                End If
            
            Case Operators.cexGreaterThan
                Compare = (compared_value > compare_value)
            Case Operators.cexGreaterThanOrEqualTo
                Compare = (compared_value >= compare_value)
            Case Operators.cexLessThan
                Compare = (compared_value < compare_value)
            Case Operators.cexLessThanOrEqualTo
                Compare = (compared_value <= compare_value)
            
            Case Operators.cexIs
                Compare = (compared_value Is compare_value)
            Case Operators.cexIsNot
                Compare = (compared_value Is Not compare_value)
        End Select
    End If
End Function

Private Function Compare2(c As Variant, property_name As String, compare_operator As Operators, compare_value As Variant) As Boolean
    If property_name = "" Then
        If Compare(c, compare_operator, compare_value) Then Compare2 = True
    Else
        If Compare(CallByName(c, property_name, VbGet), compare_operator, compare_value) Then Compare2 = True
    End If
End Function

Private Function QuickSort(ByRef col As Collection, _
                           ByVal low_boundary As Long, _
                           ByVal up_boundary As Long, _
                           is_asc As Boolean, _
                           property_name As String _
                           ) As Collection
    Dim i As Long, j As Long, pivot As Variant, temp As Variant
    Dim property_col As Collection, sorting_col As Collection
    
    Set property_col = ColEx(col).SelectBy(property_name).Items
    Set sorting_col = ColEx(col).Items
    i = low_boundary
    j = up_boundary
    pivot = property_col(Round((low_boundary + up_boundary) \ 2, 0))
    
    Do While i <= j
        If is_asc Then
            Do While property_col(i) < pivot: i = i + 1: If i > up_boundary Then Exit Do:
            Loop
            Do While property_col(j) > pivot: j = j - 1: If j < low_boundary Then Exit Do:
            Loop
        Else
            Do While property_col(i) > pivot: i = i + 1: If i > up_boundary Then Exit Do:
            Loop
            Do While property_col(j) < pivot: j = j - 1: If j < low_boundary Then Exit Do:
            Loop
        End If
        If i <= j Then
            If IsObject(col(i)) Then Set temp = col(i) Else temp = col(i)
            Call sorting_col.Add(sorting_col(j), Before:=i)
            Call sorting_col.Remove(i + 1)
            Call sorting_col.Add(temp, Before:=j)
            Call sorting_col.Remove(j + 1)
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If low_boundary < j Then Set sorting_col = QuickSort(sorting_col, low_boundary, j, is_asc, property_name)
    If i < up_boundary Then Set sorting_col = QuickSort(sorting_col, i, up_boundary, is_asc, property_name)
    
    Set QuickSort = sorting_col
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
