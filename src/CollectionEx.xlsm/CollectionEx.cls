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
' Version: 0.3.0 (developping)
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
    Dim col As New Collection, c, recol As New CollectionEx, v, evaluatedFormula, Lambda As New Lambda
    For Each c In items_
        If Lambda(c, lambdaString) Then col.Add c
    Next
    Set Where = recol.Init(col)
End Function

Public Function SelectBy(lambdaString As String) As CollectionEx
    Dim col As New Collection, c, recol As New CollectionEx, Lambda As New Lambda
    For Each c In items_
        col.Add Lambda(c, lambdaString)
    Next
    Set SelectBy = recol.Init(col)
End Function

Public Function Count(Optional lambdaString As String) As Long
    If lambdaString = "" Then Count = items_.Count: Exit Function
    Dim col As New Collection, c, Lambda As New Lambda
    For Each c In items_
        If Lambda(c, lambdaString) Then col.Add c
    Next
    Count = col.Count
End Function

Public Function Add(collectionItem, Optional Key As String) As CollectionEx
    If Key = "" Then items_.Add collectionItem Else items_.Add collectionItem, Key
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

Public Function FirstOrDefault()
    If items_.Count = 0 Then Set FirstOrDefault = Nothing Else If IsItemTypeObject Then Set FirstOrDefault = items_.Item(1) Else FirstOrDefault = items_.Item(1)
End Function

Public Function LastOrDefault()
    If items_.Count = 0 Then Set LastOrDefault = Nothing Else If IsItemTypeObject Then Set LastOrDefault = items_.Item(items_.Count) Else LastOrDefault = items_.Item(items_.Count)
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

Public Function AllEvaluate(lambdaString As String) As Boolean
    AllEvaluate = False
    Dim col As New Collection, c, recol As New CollectionEx, Lambda As New Lambda
    For Each c In items_
         If Lambda(c, lambdaString) = False Then Exit Function
    Next
    AllEvaluate = True
End Function

Public Function AnyEvaluate(lambdaString As String) As Boolean
    AnyEvaluate = True
    Dim col As New Collection, c, Lambda As New Lambda
    For Each c In items_
         If Lambda(c, lambdaString) = True Then Exit Function
    Next
    AnyEvaluate = False
End Function

Public Function Sum() As Double
    Dim col As New Collection, c, val
    If IsItemTypeObject Then Err.Raise 9001, , Err.Description
    For Each c In items_
         val = val + c
    Next
    Sum = val
End Function

Public Function Average() As Double
    Average = Sum / items_.Count
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
        If IsItemTypeObject Then Set arr(i) = items_.Item(i + 1) Else Let arr(i) = items_.Item(i + 1)
    Next i
    ToArray = arr
End Function


Private Function QuickSort(ByRef colEx As CollectionEx, ByVal low As Long, ByVal up As Long, isAsc As Boolean, Lambda As String) As Collection
    Dim i As Long, j As Long, pivot As Variant, temp, Lam As Collection, col As Collection
    
    Set Lam = colEx.SelectByLambda(Lambda).Items
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
    If low < j Then QuickSort colEx, low, j, isAsc, Lambda
    If i < up Then QuickSort colEx, i, up, isAsc, Lambda
    Set QuickSort = col
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

