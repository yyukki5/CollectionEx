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
' Version: 0.5.0 (developping)
' Depend on: Lambda.cls
' Specification may be changed.
' -------------------------------------------


Public Function Init(col) As CollectionEx
Attribute Init.VB_Description = "Default Function"
Attribute Init.VB_UserMemId = 0
    Set items_ = New Collection
    Set items_ = col
    Set Init = Me
End Function

Public Property Get Items() As Collection
    Set Items = items_
End Property

Public Property Get Item(Optional index As Long = 1)
    SetObjectAsBuffer items_.Item(index), Item
End Property


Public Function Where(lambdaString As String) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, v, evaluatedFormula
    For Each c In items_
        If Lambda(c, lambdaString).Evaluate Then col.Add c
    Next
    Set Where = recol.Init(col)
End Function

Public Function SelectBy(lambdaString As String) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx
    For Each c In items_
        col.Add Lambda(c, lambdaString).Evaluate
    Next
    Set SelectBy = recol.Init(col)
End Function

Public Function Count(Optional lambdaString As String) As Long
    If lambdaString = "" Then Count = items_.Count: Exit Function
    Dim col As New Collection, c
    For Each c In items_
        If Lambda(c, lambdaString).Evaluate Then col.Add c
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

Public Function First(Optional lambdaString As String)
    Dim col As Collection
    If items_.Count = 0 Then Call Err.Raise(9001, "CollectionEx.First()", "no items")
    
    Set col = FirstLastCore(lambdaString)
    If col.Count = 0 Then Call Err.Raise(9001, "CollectionEx.First()", "no items") Else Call SetObjectAsBuffer(col.Item(1), First)
End Function
    
Public Function FirstOrDefault(Optional lambdaString As String)
    Dim col As Collection
    If items_.Count = 0 Then FirstOrDefault = Null: Exit Function
    Set col = FirstLastCore(lambdaString)
    If col.Count = 0 Then Call SetObjectAsBuffer(GetDefaultValue, FirstOrDefault) Else Call SetObjectAsBuffer(col.Item(1), FirstOrDefault)
End Function

Public Function Last(Optional lambdaString As String)
    Dim col As Collection
    If items_.Count = 0 Then Call Err.Raise(9001, "CollectionEx.Last()", "no items")
    Set col = FirstLastCore(lambdaString)
    If col.Count = 0 Then Call Err.Raise(9001, "CollectionEx.Last()", "no items") Else Call SetObjectAsBuffer(col.Item(col.Count), Last)
End Function

Public Function LastOrDefault(Optional lambdaString As String)
    Dim col As Collection
    If items_.Count = 0 Then LastOrDefault = Null: Exit Function
    Set col = FirstLastCore(lambdaString)
    If col.Count = 0 Then Call SetObjectAsBuffer(GetDefaultValue, LastOrDefault) Else Call SetObjectAsBuffer(col.Item(col.Count), LastOrDefault)
End Function

Public Function Skip(skipIndex As Long) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, i As Long
    For Each c In items_
        i = i + 1
        If i > skipIndex Then col.Add c
    Next
    Set Skip = recol.Init(col)
End Function

Public Function Take(takeIndex As Long) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, i As Long
    For Each c In items_
        i = i + 1
        If i <= takeIndex Then col.Add c
    Next
    Set Take = recol.Init(col)
End Function

Public Function AllBy(lambdaString As String) As Boolean
    Dim col As New Collection, c
    AllBy = False
    For Each c In items_
         If Lambda(c, lambdaString).Evaluate = False Then Exit Function
    Next
    AllBy = True
End Function

Public Function AnyBy(lambdaString As String) As Boolean
    Dim col As New Collection, c
    AnyBy = True
    For Each c In items_
         If Lambda(c, lambdaString).Evaluate = True Then Exit Function
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
    Dim col As New Collection, c, val
    If items_.Count = 0 Then Sum = 0: Exit Function
    If IsNumeric(Item) = False Then Err.Raise 9001, , Err.Description
    For Each c In items_
         val = val + c
    Next
    Sum = val
End Function

Public Function Average()
    If items_.Count = 0 Then Average = Null: Exit Function

    Average = Sum / items_.Count
End Function

Public Function Max() As Double
    Dim c, tmp
    If items_.Count = 0 Then Max = Null: Exit Function
    tmp = Item
    For Each c In items_
        tmp = IIf(c > tmp, c, tmp)
    Next
    Max = tmp
End Function

Public Function Min() As Double
    Dim c, tmp
    If items_.Count = 0 Then Min = Null: Exit Function
    tmp = Item
    For Each c In items_
        tmp = IIf(c < tmp, c, tmp)
    Next
    Min = tmp
End Function

Public Function Orderby(lambdaString As String) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, sortcol As New CollectionEx
    For Each c In items_:         col.Add c:    Next
    Set Orderby = recol.Init(QuickSort(sortcol.Init(col), 1, items_.Count, True, lambdaString))
End Function

Public Function OrderByDescending(lambdaString As String) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, sortcol As New CollectionEx
    For Each c In items_:         col.Add c:    Next
    Set OrderByDescending = recol.Init(QuickSort(sortcol.Init(col), 1, items_.Count, False, lambdaString))
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
Private Function QuickSort(ByRef colEx As CollectionEx, ByVal low As Long, ByVal up As Long, isAsc As Boolean, lambdaString As String) As Collection
    Dim i As Long, j As Long, pivot As Variant, temp, Lam As Collection, col As Collection
    
    Set Lam = colEx.SelectBy(lambdaString).Items
    Set col = colEx.Items
    i = low
    j = up
    pivot = Lam(Round((low + up) \ 2, 0))
    
    Do While i <= j
        If isAsc Then
            Do While Lam(i) < pivot: i = i + 1: Loop
            Do While Lam(j) > pivot: j = j - 1: Loop
        Else
            Do While Lam(i) > pivot: i = i + 1: Loop
            Do While Lam(j) < pivot: j = j - 1: Loop
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
    Call colEx.Init(col)
    If low < j Then QuickSort colEx, low, j, isAsc, lambdaString
    If i < up Then QuickSort colEx, i, up, isAsc, lambdaString
    Set QuickSort = col
End Function

Private Function FirstLastCore(lambdaString As String)
    Dim col As New Collection
    If lambdaString <> "" Then Set col = Where(lambdaString).Items Else Set col = items_
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
