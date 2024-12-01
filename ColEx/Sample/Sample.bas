Attribute VB_Name = "Sample"
'<dir .\Sample /dir>
Option Explicit

Sub SampleCode()
    Dim cls1 As Class1
    Dim col As New Collection
    Dim i As Long
    For i = 1 To 10
        Set cls1 = New Class1
        Call col.Add(cls1.Init(i))
    Next
        
    Dim res
    res = ColEx(col) _
        .Where("abc", cexLessThan, 7) _
        .OrderByDescending("abc") _
        .Take(3) _
        .SelectBy("abc") _
        .ToArray
            
    For i = LBound(res) To UBound(res)
        Debug.Print res(i)
    Next i
End Sub



