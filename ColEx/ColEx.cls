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
' Version: 0.1.0 (developing)
' Specification may be changed.
' -------------------------------------------

Private items_ As Collection

Public Enum Operators
    Equal
    Plus
    Minus
End Enum
'Public Enum cexCompareOperators
'    Equal
'End Enum


Private Sub Class_Initialize()
    Set items_ = New Collection
End Sub

Public Function Initialize(col As Collection) As ColEx
    Set items_ = col
    Set Initialize = Me
End Function

Public Function Create(enum_variant) As ColEx
Attribute Create.VB_UserMemId = 0
    Dim col As Collection
    If TypeName(enum_variant) = "Collection" Then
        Set col = enum_variant
    Else
        Dim c As Variant
        Set col = New Collection
        For Each c In enum_variant
            Call col.Add(c)
        Next
    End If
    
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

Public Sub Add(class_item As Variant)
    Call items_.Add(class_item)
End Sub

Public Sub Remove(iii As Long)
    Call items_.Remove(iii)
End Sub

Public Function Where(Optional property_name As String, Optional compare_operator As Operators, Optional compare_value As Variant) As ColEx
    Dim c As Variant, col As New Collection, recol As New ColEx: Set col = New Collection
    For Each c In items_
        If property_name = "" Then
            If Compare(c, compare_operator, compare_value) Then Call col.Add(c)
        Else
            If Compare(CallByName(c, property_name, VbGet), compare_operator, compare_value) Then Call col.Add(c)
        End If
    Next
    Set Where = recol.Initialize(col)
End Function

Public Function SelectBy(property_name As String) As ColEx
    Dim c As Variant, col As New Collection, recol As New ColEx: Set col = New Collection
    For Each c In items_
        Call col.Add(CallByName(c, property_name, VbGet))
    Next
    Set SelectBy = recol.Initialize(col)
End Function


Public Function Order() As ColEx
        
End Function
Public Function OrderByDecending() As ColEx

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


Public Function AllBy() As Boolean

End Function
Public Function AnyBy() As Boolean

End Function
Public Function Contains() As Boolean

End Function


Private Function Compare(compared_value As Variant, compare_operator As Operators, Optional compare_value As Variant = Null)
    If IsNull(compare_value) Then
    
    Else
        Select Case compare_operator
            Case Operators.Equal
                If IsObject(compared_value) Then
                    Compare = compared_value.Equals((compare_value))
                Else
                    Compare = (compared_value = compare_value)
                End If
            Case Operators.Plus
                Compare = (compared_value + compare_value)
        End Select
    End If
End Function
