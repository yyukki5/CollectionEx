VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Implement simple anonymous function by string"
Option Explicit

' ----------------------------------------------------
' Project   : https://github.com/yyukki5/Lambda
' Version   : 0.2.0
' ---
' try to simply calculate by not using Evaluate().
' ----------------------------------------------------

Private Const TAG_STRING_OPEN As String = "<string "
Private Const TAG_STRING_CLOSE As String = " string/>"


Public Function Evaluate(evaluating_variable, lambda_string As String)
Attribute Evaluate.VB_UserMemId = 0
    Dim eval_var, var_string As String, formula_string As String, parsed_string As String, calculated_value, evaluated_value
        
On Error GoTo errEvaluate
'    If IsObject(evaluating_variable) Then Set eval_var = evaluating_variable Else Let eval_var = evaluating_variable
    Call AddTag(evaluating_variable, eval_var)

    Call ValidateLambdaString(lambda_string)
    var_string = GetVarString(lambda_string)
    formula_string = GetFormulaString(lambda_string)
    
    parsed_string = ParseParentheses(eval_var, var_string, formula_string)
    Call TryConvertPolynominalToValue(eval_var, var_string, parsed_string, evaluated_value)
    Call RemoveTags(evaluated_value, Evaluate)
    Exit Function
    
errEvaluate:
    Err.Raise Err.Number, Err.Source & vbCrLf & "at Lambda.Evaluate()", Err.Description
End Function

'Private ---------------------------------------------
Private Sub ValidateLambdaString(lambda_string As String)
    If InStr(lambda_string, "=>") = 0 Then Err.Raise 9001, "Lambda.SetLambdaValueFormula()", "Lambda string has error. need ""=>"""
End Sub


Private Function AddTag(unverfied_value, re_value)
    If IsObject(unverfied_value) Then
        Set re_value = unverfied_value
    Else
        Select Case VarType(unverfied_value)
            Case vbString: re_value = ConvertAsString(unverfied_value)
            Case Else:     re_value = unverfied_value
        End Select
    End If
End Function
Private Function RemoveTags(unverfied_value, re_value)
    If IsObject(unverfied_value) Then
        Set re_value = unverfied_value
    Else
        Select Case VarType(unverfied_value)
            Case vbString: re_value = DeconvertAsString(unverfied_value)
            Case Else:     re_value = unverfied_value
        End Select
    End If
End Function



Private Function GetVarString(lambda_string As String) As String
    Dim var_string As String
    
    var_string = Mid(lambda_string, 1, InStr(lambda_string, "=>") - 1)
    If Len(var_string) = 0 Then Err.Raise 9002, "Lambda.SetLambdaValueFormula()", "Lambda string has error."
    
    GetVarString = var_string
End Function

Private Function GetFormulaString(lambda_string As String) As String
    Dim formula_string As String
    
    formula_string = Mid(lambda_string, InStr(lambda_string, "=>") + 2)
    If Len(formula_string) = 0 Then Err.Raise 9002, "Lambda.SetLambdaValueFormula()", "Lambda string has error."
    
    GetFormulaString = formula_string
End Function

Private Function ParseParentheses(eval_var, var_string As String, formula_string As String) As String
    Dim formula_stack As Collection, current_str As String, char As String, i As Long, calculated_value
    Set formula_stack = New Collection

    For i = 1 To Len(formula_string)
        char = Mid(formula_string, i, 1)

        If char = "(" Then
            If current_str <> "" Then formula_stack.Add current_str
            current_str = ""
        
        ElseIf char = ")" Then
            ' after close parenthese ----------
            calculated_value = ConvertPolynomialToValue(eval_var, var_string, current_str)
            
            If formula_stack.Count > 0 Then
                current_str = formula_stack.Item(formula_stack.Count) & "(" & calculated_value & ")"
                formula_stack.Remove formula_stack.Count
            Else
                current_str = ""
            End If
            ' -----------------------------
        Else
            current_str = current_str & char
            
        End If
    Next i
    ParseParentheses = current_str
End Function


' example) "x.abc" -> x.abc or "x.abc.def" -> 3,   "x.abc + x.abc.def" -> x.abc + 3 (error)
Private Function TryConvertPolynominalToValue(eval_var, var_str As String, formula_str As String, ByRef return_value) As Boolean

    If InStrOperator(formula_str, 1) > 0 Then
        ' polynominal
        return_value = ConvertPolynomialToValue(eval_var, var_str, formula_str)
        TryConvertPolynominalToValue = False
    Else
        ' monominal
        TryConvertPolynominalToValue = TryConvertMonomialToValue(eval_var, formula_str, return_value)
    End If
    
End Function

' example) "x.abc.def + x.abc.efg + 1"  ->  "3 + 5 + 1" -> 9
Private Function ConvertPolynomialToValue(eval_var, var_str As String, formula_str As String) As Variant
    Dim calculated_value, calculated_string
    
    If InStr(formula_str, var_str) > 0 Then
        calculated_string = ConvertToPolynomialWithoutObject(eval_var, var_str, formula_str)
    Else
        calculated_string = formula_str
    End If
    
    calculated_value = RunEvaluate(calculated_string)
    If IsError(calculated_value) Then Err.Raise 9006, "Lambda.ConvertPolynomialToValue()", "Unsupported lambda string. can not evaluate formula"
    
    ConvertPolynomialToValue = calculated_value
End Function

' example) "x.abc.def + x.abc.efg + 1"  ->  "3 + 5 + 1"
Private Function ConvertToPolynomialWithoutObject(eval_var, var_str As String, formula_str As String) As String
    Dim start_index, last_index, term_string As String, buf_val
    
    Do While InStr(formula_str, var_str) > 0
    
        start_index = InStr(formula_str, var_str)
        last_index = IIf(InStrOperator(formula_str, start_index) = 0, Len(formula_str), InStrOperator(formula_str, start_index) - 1)
        term_string = Mid(formula_str, start_index, last_index - start_index + 1)
    
        Call TryConvertMonomialToValue(eval_var, term_string, buf_val)
    
        If IsObject(buf_val) Then Err.Raise 9001, "Lambda.ConvertToPolynomialWithoutObject()", "Failed to evaluate, converted as object"
        If VarType(buf_val) > vbArray Then Err.Raise 9001, "Lambda.ConvertToPolynomialWithoutObject()", "Unsupport array for evaluate"
        
        formula_str = Replace(formula_str, term_string, buf_val)
    Loop
    
    ConvertToPolynomialWithoutObject = formula_str
End Function


' example) "x.abc.def" -> 1
Private Function TryConvertMonomialToValue(eval_var, monominal_string As String, ByRef return_value) As Boolean
    Dim splited_by_dot, i, index_parenthes_open As Long
    
    monominal_string = Replace(monominal_string, " ", "")
    splited_by_dot = Split(monominal_string, ".")
    
    If UBound(splited_by_dot) = 0 Then
        If InStrOperator(CStr(monominal_string), 1) > 0 Then Err.Raise 9004, "Lambda.TryConvertMonomialToValue()", "Lambda string has error."
        If IsObject(eval_var) Then Set return_value = eval_var Else return_value = eval_var
    Else
        Set return_value = eval_var
        For i = 1 To UBound(splited_by_dot)
            If splited_by_dot(i) = "" Then Err.Raise 9004, "Lambda.TryConvertMonomialToValue()", "Lambda string has error."
            
            index_parenthes_open = InStr(splited_by_dot(i), "(")
            If index_parenthes_open > 0 Then
                Call SetCallByNameValue( _
                    called_object:=return_value, _
                    called_proc_name:=Left(splited_by_dot(i), index_parenthes_open - 1), _
                    return_value:=return_value, _
                    args:=Mid(splited_by_dot(i), index_parenthes_open + 1, InStr(splited_by_dot(i), ")") - index_parenthes_open - 1))
            Else
                Call SetCallByNameValue(return_value, CStr(splited_by_dot(i)), return_value)
            End If
        Next i
    End If
    
    If IsObject(return_value) Then TryConvertMonomialToValue = True Else TryConvertMonomialToValue = False
End Function

Private Function InStrOperator(string_has_operator, start_index) As Long
    Dim i As Long, str As String
    
    For i = start_index To Len(string_has_operator)
        str = Mid(string_has_operator, i)
        
        i = i + GetTagLength(str)
        
        If IsOperator(str) Then InStrOperator = i: Exit Function
    Next i
    InStrOperator = 0
End Function

Private Function IsOperator(operator_string) As Boolean
    IsOperator = True
    If Left(operator_string, 1) = "<" Then Exit Function
    If Left(operator_string, 2) = "<=" Then Exit Function
    If Left(operator_string, 1) = ">" Then Exit Function
    If Left(operator_string, 2) = ">=" Then Exit Function
    If Left(operator_string, 1) = "=" Then Exit Function
    If Left(operator_string, 1) = "+" Then Exit Function
    If Left(operator_string, 1) = "-" Then Exit Function
    If Left(operator_string, 1) = "*" Then Exit Function
    If Left(operator_string, 1) = "/" Then Exit Function
    If Left(operator_string, 1) = "^" Then Exit Function
    If Left(operator_string, 1) = "&" Then Exit Function
    IsOperator = False
End Function

Private Function SetCallByNameValue(called_object, called_proc_name As String, ByRef return_value, Optional args As Variant) As Boolean
    If IsMissing(args) Then
        On Error GoTo errGetWithoutArg
        Call SetObjectAsBuffer(CallByName(called_object, called_proc_name, VbGet), return_value)
        SetCallByNameValue = IsObject(return_value)
        Exit Function
errGetWithoutArg:
        If Err.Number = 450 Then
            Call SetObjectAsBuffer(CallByName(called_object, called_proc_name, VbMethod), return_value)
            SetCallByNameValue = IsObject(return_value)
        ElseIf Err.Number = 438 Then
            Err.Raise 9001, "Lambda.SetCallByNameValue()", "Not implemented property: " & TypeName(called_object) & "." & called_proc_name
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Else
        On Error GoTo errGetWitArg
        Call SetObjectAsBuffer(CallByName(called_object, called_proc_name, VbGet, args), return_value)
        SetCallByNameValue = IsObject(return_value)
        Exit Function
errGetWitArg:
        If Err.Number = 450 Then
            Call SetObjectAsBuffer(CallByName(called_object, called_proc_name, VbMethod, args), return_value)
            SetCallByNameValue = IsObject(return_value)
        ElseIf Err.Number = 438 Then
            Err.Raise 9001, "Lambda.SetCallByNameValue()", "Not implemented property: " & TypeName(called_object) & "." & called_proc_name
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If
    Exit Function
End Function

Private Function SetObjectAsBuffer(ByRef unverified_value, ByRef buffer_value)
    Dim isObj As Boolean
    isObj = IsObject(unverified_value)
    If isObj Then Set buffer_value = unverified_value Else buffer_value = unverified_value
End Function

' "1 + 1 + 1"  -> 3
Private Function RunEvaluate(non_parentheses_formula_str) As Variant
    Dim result_value
    result_value = SplitByOperator(non_parentheses_formula_str)
    RunEvaluate = result_value
End Function

' "1 + 2 + 3"  -> 6
Private Function SplitByOperator(string_has_operator)
    Dim i As Long, operator_strings, o, formular_info(0 To 2)
    
    ' note: element arrangement in array equal calaculated order
    operator_strings = Array(" Imp ", " Eqv ", " Xor ", " Or ", " And ", " Not ", _
                " Is ", " Like ", "<=", ">=", ">", "<", "<>", "=", _
                "&", "+", "-", " Mod ", "\", "*", "/", " ^ ")
    
On Error GoTo errSplitByOperator
    For Each o In operator_strings
        SplitByOperator = CalculateByOperator(string_has_operator, o)
        If IsEmpty(SplitByOperator) = False Then Exit Function
    Next
    
    If IsNumeric(string_has_operator) Then SplitByOperator = val(string_has_operator) Else SplitByOperator = string_has_operator
    Exit Function
    
errSplitByOperator:
   Err.Raise 9001, Err.Source & vbCrLf & " at Lambda.SplitByOperator(" & string_has_operator & ")", Err.Description & vbCrLf & "failed to split by operator"
    
End Function
 
' "1 + 2 + 3"  -> 6
Private Function CalculateByOperator(string_has_operator, o)
    Dim i As Long, vl, vr

On Error GoTo errCalculate
    For i = 1 To Len(string_has_operator)
        i = i + GetTagLength(Mid(string_has_operator, i))
    
        If o = Mid(string_has_operator, i, 1) Or o = Mid(string_has_operator, i, 2) Or o = Mid(string_has_operator, i, 3) Or o = Mid(string_has_operator, i, 4) Or o = Mid(string_has_operator, i, 5) Or o = Mid(string_has_operator, i, 6) Then
            vl = SplitByOperator(Left(string_has_operator, i - 1))
            vr = SplitByOperator(Mid(string_has_operator, i + Len(o)))
            CalculateByOperator = CalculateOperator(operator_string:=o, left_value:=vl, right_value:=vr)
            Exit Function
        End If
    Next i
    Exit Function
    
errCalculate:
    Err.Raise 9001, Err.Source & vbCrLf & " at Lambda.CalculateByOperator(" & string_has_operator & "," & o & ")", Err.Description & vbCrLf & "failed to calculate by operator: " & o
End Function

' 1 + 2  -> 3
Private Function CalculateOperator(operator_string, left_value, right_value)
    Dim calculated_value
    left_value = Trim(left_value)
    right_value = Trim(right_value)
    If IsNumeric(left_value) Then left_value = val(left_value)
    If IsNumeric(right_value) Then right_value = val(right_value)

On Error GoTo errCalc
    Select Case operator_string
        Case "<=": calculated_value = left_value <= right_value
        Case ">=": calculated_value = left_value >= right_value
        Case ">": calculated_value = left_value > right_value
        Case "<": calculated_value = left_value < right_value
        Case "<>": calculated_value = left_value <> right_value
        Case "=": calculated_value = left_value = right_value
        Case "&": calculated_value = left_value & right_value
        Case "+": calculated_value = left_value + right_value
        Case "-": calculated_value = left_value - right_value
        Case "*": calculated_value = left_value * right_value
        Case "/": calculated_value = left_value / right_value
        Case "\": calculated_value = left_value \ right_value
        Case " ^ ": calculated_value = left_value ^ right_value
        Case " Mod ": calculated_value = left_value Mod right_value
        Case " Imp ": calculated_value = left_value Imp right_value
        Case " Eqv ": calculated_value = left_value Eqv right_value
        Case " Xor ": calculated_value = left_value Xor right_value
        Case " Or ": calculated_value = left_value Or right_value
        Case " And ": calculated_value = left_value And right_value
        Case " Not ": calculated_value = Not left_value
        Case " Like ": calculated_value = left_value Like right_value
        Case " Is ": calculated_value = left_value Is right_value
    End Select
    
    CalculateOperator = calculated_value
    Exit Function
errCalc:
    Err.Raise 9001, "Lambda.CalculateOperator()", "formula string has error : " & left_value & " " & operator_string & " " & right_value
End Function

Private Function GetTagLength(str) As Long
    GetTagLength = 0
    If Left(str, Len(TAG_STRING_OPEN)) = TAG_STRING_OPEN Then _
        GetTagLength = InStr(str, TAG_STRING_CLOSE) + Len(TAG_STRING_CLOSE)
End Function



Private Function ConvertAsString(converted_string) As String
    If InStr(converted_string, TAG_STRING_OPEN) > 0 Then ConvertAsString = converted_string Else ConvertAsString = TAG_STRING_OPEN & converted_string & TAG_STRING_CLOSE
End Function
Private Function DeconvertAsString(deconverted_string)
    If InStr(deconverted_string, TAG_STRING_OPEN) > 0 Then DeconvertAsString = Replace(Replace(deconverted_string, TAG_STRING_OPEN, ""), TAG_STRING_CLOSE, "")
End Function
