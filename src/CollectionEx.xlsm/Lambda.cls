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

Private Const APP_NAME_EXCEL As String = "Microsoft Excel"
Private Const APP_NAME_ACCESS As String = "Microsoft Access"

Private app_ As Object
    
    
Private Sub Class_Initialize()
    If Application.name = APP_NAME_EXCEL Or Application.name = APP_NAME_ACCESS Then Else Debug.Print "This application is not supported Evaluate()/Eval()"
    Set app_ = Application
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
    On Error GoTo errNotImplement
    If IsMissing(args) Then
        If IsObject(CallByName(cls, propName, VbGet)) Then
            Set reval = CallByName(cls, propName, VbGet)
            SetCallByNameValue = True
        Else
            reval = CallByName(cls, propName, VbGet)
            SetCallByNameValue = False
        End If
    Else
        If IsObject(CallByName(cls, propName, VbGet, args)) Then
            Set reval = CallByName(cls, propName, VbGet, args)
            SetCallByNameValue = True
        Else
            reval = CallByName(cls, propName, VbGet, args)
            SetCallByNameValue = False
        End If
    End If
    Exit Function
errNotImplement:
    If Err.Number = 438 Then Err.Raise 9001, "Lambda.SetCallByNameValue()", "Not implemented property: " & TypeName(cls) & "." & propName _
        Else Err.Raise Err.Number, Err.Source, Err.Description
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

Private Function RunEvaluate(formulastr) As Variant
    Dim res
    If Application.name = APP_NAME_EXCEL Then
        res = app_.Evaluate(formulastr)
    ElseIf Application.name = APP_NAME_ACCESS Then
        res = app_.Eval(formulastr)
    Else
        Err.Raise 9900, "Lambda.RunEvaluate()", "This application is not supported Evaluate()/Eval()"
    End If
    RunEvaluate = res
End Function




