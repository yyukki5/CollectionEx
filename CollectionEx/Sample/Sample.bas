Attribute VB_Name = "Sample"
'<dir .\Sample /dir>
Option Explicit

Sub SampleCode()
    Dim cls1 As Class1
    Dim col As New Collection
    Dim i As Long
    For i = 1 To 10
        Set cls1 = New Class1
        col.Add cls1.Init(i)
    Next
    
    Dim res
    res = CollectionEx(col) _
        .Where("x => x.abc < 7") _
        .OrderByDescending("x => x.abc") _
        .Take(3) _
        .SelectBy("x => x.abc") _
        .ToArray
        
    For i = LBound(res) To UBound(res)
        Debug.Print res(i)
    Next i
End Sub



