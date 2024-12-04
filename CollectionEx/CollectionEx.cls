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
' Version: 0.8.0 (developing)
' Depend on: Lambda.cls
' Specification may be changed.
' -------------------------------------------


Public Function Initialize(col As Collection) As CollectionEx
    Set items_ = col
    Set Initialize = Me
End Function

Public Function Create(enum_variant) As CollectionEx
Attribute Create.VB_Description = "Return new instance of CollectionEx"
Attribute Create.VB_UserMemId = 0
    Dim col As Collection, c
    
    Set col = New Collection
    For Each c In enum_variant
        Call col.Add(c)
    Next

    Dim cex As New CollectionEx
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
    Dim col As New Collection, c, recol As New CollectionEx
    For Each c In items_
        If Lambda(c, lambda_string).Evaluate = True Then col.Add c
    Next
    Set Where = recol.Initialize(col)
End Function

Public Function SelectBy(lambda_string As String) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx
    For Each c In items_
        col.Add Lambda(c, lambda_string).Evaluate
    Next
    Set SelectBy = recol.Initialize(col)
End Function

Public Function SelectManyBy(lambda_string As String) As CollectionEx
    Dim c As Variant, c1 As Variant, col As Collection, recol As New CollectionEx
    Set col = New Collection
    
    For Each c In items_
        For Each c1 In Lambda(c, lambda_string).Evaluate
            Call col.Add(c1)
        Next
    Next

    Set SelectManyBy = recol.Initialize(col)
End Function


Public Function Count(Optional lambda_string As String) As Long
    If lambda_string = "" Then Count = items_.Count: Exit Function
    Dim col As New Collection, c
    For Each c In items_
        If Lambda(c, lambda_string).Evaluate Then col.Add c
    Next
    Count = col.Count
End Function

Public Function Add(addedItem, Optional Key As String) As CollectionEx
    If Key = "" Then items_.Add addedItem Else items_.Add addedItem, Key
    Set Add = Me
End Function

Public Function AddRange(coll As Collection) As CollectionEx
    Dim c
    For Each c In coll
        items_.Add c
    Next
    Set AddRange = Me
End Function

Public Function Remove(index) As CollectionEx
    items_.Remove (index)
    Set Remove = Me
End Function

Public Function First(Optional lambda_string As String)
    Dim col As Collection
    If items_.Count = 0 Then Call Err.Raise(9001, "CollectionEx.First()", "no items")
    
    Set col = FirstLastCore(lambda_string)
    If col.Count = 0 Then Call Err.Raise(9001, "CollectionEx.First()", "no items") Else Call SetObjectAsBuffer(col.Item(1), First)
End Function
    
Public Function FirstOrDefault(Optional lambda_string As String, Optional default_value As Variant = Null)
    Dim col As Collection
    If items_.Count = 0 Then Call SetObjectAsBuffer(default_value, FirstOrDefault): Exit Function
    Set col = FirstLastCore(lambda_string)
    If col.Count = 0 Then Call SetObjectAsBuffer(default_value, FirstOrDefault) Else Call SetObjectAsBuffer(col.Item(1), FirstOrDefault)
End Function

Public Function Last(Optional lambda_string As String)
    Dim col As Collection
    If items_.Count = 0 Then Call Err.Raise(9001, "CollectionEx.Last()", "no items")
    Set col = FirstLastCore(lambda_string)
    If col.Count = 0 Then Call Err.Raise(9001, "CollectionEx.Last()", "no items") Else Call SetObjectAsBuffer(col.Item(col.Count), Last)
End Function

Public Function LastOrDefault(Optional lambda_string As String, Optional default_value As Variant = Null)
    Dim col As Collection
    If items_.Count = 0 Then Call SetObjectAsBuffer(default_value, LastOrDefault): Exit Function
    Set col = FirstLastCore(lambda_string)
    If col.Count = 0 Then Call SetObjectAsBuffer(default_value, LastOrDefault) Else Call SetObjectAsBuffer(col.Item(col.Count), LastOrDefault)
End Function

Public Function Skip(skip_index As Long) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, i As Long
    For Each c In items_
        i = i + 1
        If i > skip_index Then col.Add c
    Next
    Set Skip = recol.Initialize(col)
End Function

Public Function Take(take_index As Long) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, i As Long
    For Each c In items_
        i = i + 1
        If i <= take_index Then col.Add c
    Next
    Set Take = recol.Initialize(col)
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
    If IsObject(var) Then Err.Raise 9001, "CollectionEx.Contains()", "Unsupported object"
    For Each c In items_
        If var = c Then Contains = True: Exit Function
    Next
    Contains = False
End Function

Public Function Sum() As Double
    Dim c, val As Double
    If items_.Count = 0 Then Sum = 0: Exit Function
    If IsNumeric(Item) = False Then Err.Raise 9001, , Err.Description
    For Each c In items_
         val = val + c
    Next
    Sum = val
End Function

Public Function Average() As Variant
    If items_.Count = 0 Then Average = Null: Exit Function

    Average = Sum / items_.Count
End Function

Public Function Max() As Variant
    Dim c, tmp
    If items_.Count = 0 Then Max = Null: Exit Function
    tmp = Item
    For Each c In items_
        tmp = IIf(c > tmp, c, tmp)
    Next
    Max = tmp
End Function

Public Function Min() As Variant
    Dim c, tmp
    If items_.Count = 0 Then Min = Null: Exit Function
    tmp = Item
    For Each c In items_
        tmp = IIf(c < tmp, c, tmp)
    Next
    Min = tmp
End Function

Public Function Orderby(lambda_string As String) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, sortcol As New CollectionEx
    For Each c In items_:         col.Add c:    Next
    Set Orderby = recol.Initialize(QuickSort(sortcol.Initialize(col), 1, items_.Count, True, lambda_string))
End Function

Public Function OrderByDescending(lambda_string As String) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, sortcol As New CollectionEx
    For Each c In items_:         col.Add c:    Next
    Set OrderByDescending = recol.Initialize(QuickSort(sortcol.Initialize(col), 1, items_.Count, False, lambda_string))
End Function

Public Function ToArray()
    Dim arr, i As Long
    ReDim arr(0 To items_.Count - 1)
    
    For i = LBound(arr) To UBound(arr)
        SetObjectAsBuffer Items.Item(i + 1), arr(i)
    Next i
    ToArray = arr
End Function



' private functions --------------------------------
Private Function QuickSort(ByRef cex As CollectionEx, ByVal low As Long, ByVal up As Long, isAsc As Boolean, lambda_string As String) As Collection
    Dim i As Long, j As Long, pivot As Variant, temp, Lam As Collection, col As Collection
    
    Set Lam = cex.SelectBy(lambda_string).Items
    Set col = cex.Items
    i = low
    j = up
    pivot = Lam(Round((low + up) \ 2, 0))
    
    Do While i <= j
        If isAsc Then
            Do While Lam(i) < pivot: i = i + 1: If i > up Then Exit Do:
            Loop
            Do While Lam(j) > pivot: j = j - 1: If j < low Then Exit Do:
            Loop
        Else
            Do While Lam(i) > pivot: i = i + 1: If i > up Then Exit Do:
            Loop
            Do While Lam(j) < pivot: j = j - 1: If j < low Then Exit Do:
            Loop
        End If
        If i <= j Then

            If IsObject(col(i)) Then Set temp = col(i) Else temp = col(i)
            col.Add col(j), Before:=i
            col.Remove i + 1
            col.Add temp, Before:=j
            col.Remove j + 1
            i = i + 1
            j = j - 1
        End If
    Loop
    Call cex.Initialize(col)
    If low < j Then Call QuickSort(cex, low, j, isAsc, lambda_string)
    If i < up Then Call QuickSort(cex, i, up, isAsc, lambda_string)
    Set QuickSort = cex.Items
End Function

Private Function FirstLastCore(lambda_string As String)
    Dim col As Collection
    If lambda_string <> "" Then Set col = Where(lambda_string).Items Else Set col = items_
    Set FirstLastCore = col
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

Private Function GetDefaultValue()
    If items_.Count = 0 Then
        GetDefaultValue = Null
    ElseIf IsObject(Item) Then
        Set GetDefaultValue = Nothing
    Else
        GetDefaultValue = Null
    End If
End Function
