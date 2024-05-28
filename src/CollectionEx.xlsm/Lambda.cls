VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
'try to simple calculation by not using Evaluate().

Private Sub Class_Initialize()
End Sub

Public Function EvaluateLambda(val, str)
    Dim varstr, formulastr, res, reval, calculatedBracketsString As String
    
    Call SetLambdaValueFormula(str, varstr, formulastr)

    formulastr = ProcessParentheses(val, varstr, CStr(formulastr))
    Call CalcInnerBrackets(val, varstr, formulastr)
    res = RunEvaluate(formulastr)
    If VarType(res) <> vbBoolean Then Err.Raise 9001, "Lambda.EvaluateLambda()", "Unsupport error: can not evaluate result as boolen"
    EvaluateLambda = res
End Function

Public Function SelectLambdaValue(val, str)
    Dim varstr, formulastr
    Dim sta, lst, cd, ln, reval
    
    Call SetLambdaValueFormula(str, varstr, formulastr)
    formulastr = ProcessParentheses(val, varstr, CStr(formulastr))
    Call SetOneValue(val, formulastr, reval)
    If IsObject(reval) Then Set SelectLambdaValue = reval Else SelectLambdaValue = reval
End Function

Private Function SetLambdaValueFormula(str, varstr, formulastr)
    str = Replace(str, " ", "")
    
    If InStr(str, "=>") = 0 Then Err.Raise 9001, "Lambda.SetLambdaValueFormula()", "Lambda string has error. need ""=>"""
    varstr = Mid(str, 1, InStr(str, "=>") - 1)
    formulastr = Mid(str, InStr(str, "=>") + 2)
    
    If Len(varstr) = 0 Then Err.Raise 9002, "Lambda.SetLambdaValueFormula()", "Lambda string has error."
    If Len(formulastr) = 0 Then Err.Raise 9002, "Lambda.SetLambdaValueFormula()", "Lambda string has error."
End Function

Private Function CalcInnerBrackets(val, varstr, formulastr) As String
    Dim sta, lst, cd, reval
    Do While InStr(formulastr, varstr) > 0
    
        sta = InStr(formulastr, varstr)
        lst = IIf(InStrOperator(formulastr, sta) = 0, Len(formulastr), InStrOperator(formulastr, sta) - 1)
        cd = Mid(formulastr, sta, lst - sta + 1)
        
        Call SetOneValue(val, cd, reval)
        If IsObject(reval) Then Err.Raise 9001, "Lambda.CalcInnerBrackets()", "Unsupport error: can not evaluate object "
        If VarType(reval) > vbArray Then Err.Raise 9001, "Lambda.CalcInnerBrackets()", "Unsupport array for evaluate"
        formulastr = Replace(formulastr, cd, reval)
    Loop
    CalcInnerBrackets = formulastr
End Function

Private Function SetOneValue(val, cd, ByRef reval) As Boolean
    Dim sp, i, indBrktOpen As Long
    sp = Split(cd, ".")
    
    If UBound(sp) = 0 Then
        If InStrOperator(CStr(cd), 1) > 0 Then Err.Raise 9004, "Lambda.SetOneValue()", "Lambda string has error."
        If IsObject(val) Then Set reval = val Else reval = val
    Else
        Set reval = val
        For i = 1 To UBound(sp)
            If sp(i) = "" Then Err.Raise 9004, "Lambda.SetOneValue()", "Lambda string has error."
            
            indBrktOpen = InStr(sp(i), "(")
            If indBrktOpen > 0 Then
                Call SetCallByNameValue(reval, Left(sp(i), indBrktOpen - 1), reval, Mid(sp(i), indBrktOpen + 1, InStr(sp(i), ")") - indBrktOpen - 1))
            Else
                Call SetCallByNameValue(reval, CStr(sp(i)), reval)
            End If
        Next i
    End If
    
    If IsObject(reval) Then SetOneValue = True Else SetOneValue = False
End Function

Private Function InStrOperator(str, sta) As Long
    Dim i As Long
    For i = 0 To Len(str) - 1
        If IsOperator(Mid(str, sta + i)) Then InStrOperator = sta + i: Exit Function
    Next i
    InStrOperator = 0
End Function

Private Function IsOperator(str) As Boolean
    IsOperator = True
    If Left(str, 1) = "<" Then Exit Function
    If Left(str, 2) = "<=" Then Exit Function
    If Left(str, 1) = ">" Then Exit Function
    If Left(str, 2) = ">=" Then Exit Function
    If Left(str, 1) = "=" Then Exit Function
    If Left(str, 1) = "+" Then Exit Function
    If Left(str, 1) = "-" Then Exit Function
    If Left(str, 1) = "*" Then Exit Function
    If Left(str, 1) = "/" Then Exit Function
    If Left(str, 1) = "^" Then Exit Function
    If Left(str, 1) = "&" Then Exit Function
    IsOperator = False
End Function

Private Function SetCallByNameValue(cls, propName As String, ByRef reval, Optional args As Variant) As Boolean
    If IsMissing(args) Then
        On Error GoTo errGetWithoutArg
        Call SetObjBuf(CallByName(cls, propName, VbGet), reval)
        SetCallByNameValue = IsObject(reval)
        Exit Function
errGetWithoutArg:
        If Err.Number = 450 Then
            Call SetObjBuf(CallByName(cls, propName, VbMethod), reval)
            SetCallByNameValue = IsObject(reval)
        ElseIf Err.Number = 438 Then
            Err.Raise 9001, "Lambda.SetCallByNameValue()", "Not implemented property: " & TypeName(cls) & "." & propName
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Else
        On Error GoTo errGetWitArg
        Call SetObjBuf(CallByName(cls, propName, VbGet, args), reval)
        SetCallByNameValue = IsObject(reval)
        Exit Function
errGetWitArg:
        If Err.Number = 450 Then
            Call SetObjBuf(CallByName(cls, propName, VbMethod, args), reval)
            SetCallByNameValue = IsObject(reval)
        ElseIf Err.Number = 438 Then
            Err.Raise 9001, "Lambda.SetCallByNameValue()", "Not implemented property: " & TypeName(cls) & "." & propName
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If
    Exit Function
End Function

Private Function SetObjBuf(ByRef val, ByRef reval)
    Dim isObj As Boolean
    isObj = IsObject(val)
    If isObj Then Set reval = val Else reval = val
End Function

Private Function ProcessParentheses(var, varstr, inputStr As String) As String
    Dim stack As Collection, currentContent As String, char As String, i As Long, calclatedContent, calculatedString

    Set stack = New Collection
    currentContent = ""
    calclatedContent = ""
    calculatedString = ""

    For i = 1 To Len(inputStr)
        char = Mid(inputStr, i, 1)

        If char = "(" Then
            If currentContent <> "" Then stack.Add currentContent
            currentContent = ""
        ElseIf char = ")" Then
            If InStr(currentContent, varstr) > 0 Then
                calculatedString = CalcInnerBrackets(var, varstr, currentContent)
                calclatedContent = RunEvaluate(calculatedString)
                If IsError(calclatedContent) Then Err.Raise 9006, "Lambda.ProcessParentheses()", "Unsupported lambda string. can not evaluate formula"
            Else
                calclatedContent = RunEvaluate(currentContent)
            End If
            ' after close bracket ----------
            If stack.Count > 0 Then
                currentContent = stack.Item(stack.Count) & "(" & calclatedContent & ")"
                stack.Remove stack.Count
                
                If Mid(inputStr, i + 1, 1) <> "." Then Call CalcInnerBrackets(var, varstr, currentContent)
            Else
                currentContent = ""
            End If
            ' -----------------------------
        Else
            currentContent = currentContent & char
        End If
    Next i
    ProcessParentheses = currentContent
End Function

Private Function RunEvaluate(nonBracketFormulaStr) As Variant
    Dim res, str
    str = Replace(nonBracketFormulaStr, " ", "")
    res = SplitByOperator(str)
    RunEvaluate = res
End Function

Private Function SplitByOperator(str)
    Dim c, i As Long, ope, o, arrOpeLfRg(0 To 2)
    
    ope = Array("<=", ">=", ">", "<", "=", "&", "+", "-", "*", "/")
    
    For Each o In ope
        For i = 1 To Len(str)
            If o = Mid(str, i, 1) Or o = Mid(str, i, 2) Then
                arrOpeLfRg(0) = o
                arrOpeLfRg(1) = Left(str, i - 1)
                arrOpeLfRg(2) = Mid(str, i + Len(o))
                
                SplitByOperator = Calc(arrOpeLfRg(0), SplitByOperator(arrOpeLfRg(1)), SplitByOperator(arrOpeLfRg(2)))
                Exit Function
            End If
        Next i
    Next
    If IsNumeric(str) Then SplitByOperator = val(str) Else SplitByOperator = str
End Function

Private Function Calc(ope, lf, rg)
    If IsNumeric(lf) Then lf = val(lf)
    If IsNumeric(rg) Then rg = val(rg)

    Select Case ope
        Case "<=": Calc = lf <= rg
        Case ">=": Calc = lf >= rg
        Case ">": Calc = lf > rg
        Case "<": Calc = lf < rg
        Case "=": Calc = lf = rg
        Case "&": Calc = lf & rg
        Case "+": Calc = lf + rg
        Case "-": Calc = lf - rg
        Case "*": Calc = lf * rg
        Case "/": Calc = lf / rg
    End Select

End Function



