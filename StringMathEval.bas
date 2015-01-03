Attribute VB_Name = "StringMathEval"
Option Explicit

Private Type Assortment
    WithinNum As Boolean
    OneChar As String
    PreviousChar As String
    Number As Double
    NumberCount As Integer
    PreviousNumIndex As Integer
    NextNumIndex As Integer
    ArrayIndex As Integer
    ArrayLevel As Integer
    ArrayLevelUbound As Integer
    DigitsBeforeDecimal As Integer
    DigitsAfterDecimal As Integer
    FoundDecimal As Boolean
End Type

Public Const Operators = "42, 43, 45, 47, 94"  '"*, +, -, /, ^"

Public Function EvalExpression(ThisExpr As String) As Variant
    Dim Result(3) As Variant
    Dim Count As Integer
    Dim count2 As Integer
    Dim count3 As Integer
    Dim Alpha() As Variant
    Dim Var As Assortment
    Dim MathErrorCodes(20) As String
    
    MathErrorCodes(0) = "No Error"
    MathErrorCodes(1) = "Malformed number. Too many decimals"
    MathErrorCodes(2) = "Multiple adjacent operators"
    MathErrorCodes(3) = "operator not valid next to paren."
    MathErrorCodes(4) = "Num. missing decimal part"
    MathErrorCodes(5) = "operator not valid at start"
    MathErrorCodes(6) = "Parenthesis pair with no operator"
    MathErrorCodes(7) = "Missing operator after number"
    MathErrorCodes(8) = "Empty parenthesis pair"
    MathErrorCodes(9) = "Missing num. before parenthesis"
    MathErrorCodes(10) = "Excess close parenthesis"
    MathErrorCodes(11) = "excess open parenthesis"
    MathErrorCodes(12) = "Division by zero"
    MathErrorCodes(13) = "Empty Expression"
    MathErrorCodes(14) = "Missing num. before operator"
    MathErrorCodes(15) = "Stray operator at end"
    MathErrorCodes(16) = "Error solving expression"
    MathErrorCodes(17) = "No alpha chars allowed"
    MathErrorCodes(18) = "Expected operator before number"
    
    Result(0) = 10
    Result(1) = 16
    Result(2) = MathErrorCodes(16)
    Result(3) = 0
    
    Var.NumberCount = 0
    
    If (ThisExpr = "") Then
        Result(0) = 10
        Result(1) = 13
        Result(2) = MathErrorCodes(13) '"Empty expression"
        EvalExpression = Result
        Exit Function
    End If
    
    ReDim Alpha(Len(ThisExpr) - 1, Var.ArrayLevelUbound)
    Var.PreviousChar = Chr(0)
    Var.ArrayIndex = -1
    
    'Populate the multi-level array
    For Count = 1 To Len(ThisExpr)
    
        Var.OneChar = Mid(ThisExpr, Count, 1)
        
        'Handle numbers
        '----------------------------
        If (Asc(Var.OneChar) > 47 And Asc(Var.OneChar) < 58 Or Var.OneChar = ".") Then
            
            'Count number within the expression
            If (Var.WithinNum = False) Then
                Var.WithinNum = True
                Var.Number = 0
                Var.DigitsAfterDecimal = 0
                Var.DigitsBeforeDecimal = 0
                Var.FoundDecimal = False
                Var.ArrayIndex = Var.ArrayIndex + 1
                Var.NumberCount = Var.NumberCount + 1
                Result(3) = Var.NumberCount
            End If
            
            'Handle decimal separator
            If (Var.OneChar = ".") Then
                If (Var.FoundDecimal = False) Then
                    Var.FoundDecimal = True
                Else
                    Result(0) = 10
                    Result(1) = 1
                    Result(2) = MathErrorCodes(1) '"Malformed number. Too many decimals"
                    EvalExpression = Result
                    Exit Function
                End If
            End If
            
            If (Var.PreviousChar = ")") Then
                    Result(0) = 10
                    Result(1) = 18
                    Result(2) = MathErrorCodes(18) '"Expected operator before number"
                    EvalExpression = Result
                    Exit Function
            End If
            
            'Construct the number
            If (Var.FoundDecimal = False) Then
                Var.DigitsBeforeDecimal = Var.DigitsBeforeDecimal + 1
                Var.Number = Var.Number * 10 + CInt(Var.OneChar)
            ElseIf (Var.OneChar <> ".") Then
                Var.DigitsAfterDecimal = Var.DigitsAfterDecimal + 1
                Var.Number = Var.Number + CInt(Var.OneChar) * (1 / 10 ^ (Var.DigitsAfterDecimal))
            End If
            
            If (Var.OneChar <> ".") Then Alpha(Var.ArrayIndex, Var.ArrayLevel) = Var.Number
        
        'Handle operators
        '----------------------------
        ElseIf (InStr(1, Operators, Asc(Var.OneChar))) Then
            Var.WithinNum = False
            
            
            If (Var.ArrayIndex >= 0) Then
            'FYI: we have not incremented the index yet
                If (Alpha(Var.ArrayIndex, Var.ArrayLevel) <> "") Then
                    'Check for double operator at the same level
                    If (InStr(1, Operators, Asc(Alpha(Var.ArrayIndex, Var.ArrayLevel)))) Then
                        Result(0) = 10
                        Result(1) = 2
                        Result(2) = MathErrorCodes(2) '"Multiple adjacent operators"
                        EvalExpression = Result
                        Exit Function
                    End If
                End If
                
                'check prenthesis neighbor condition
                If (Var.PreviousChar = "(" And (Var.OneChar <> "-" And Var.OneChar <> "+")) Then
                    Result(0) = 10
                    Result(1) = 3
                    Result(2) = "'" & Var.OneChar & "' " & MathErrorCodes(3) '"operator not valid next to paren."
                    EvalExpression = Result
                    Exit Function
                End If
                
                'check incomplete decimal
                If (Var.PreviousChar = ".") Then
                    Result(0) = 10
                    Result(1) = 4
                    Result(2) = MathErrorCodes(4) '"Num. missing decimal part"
                    EvalExpression = Result
                    Exit Function
                End If
            ElseIf (Var.ArrayIndex < 0) Then
                If (Var.OneChar <> "-" And Var.OneChar <> "+") Then
                    Result(0) = 10
                    Result(1) = 5
                    Result(2) = "'" & Var.OneChar & "' " & MathErrorCodes(5) '"operator not valid at start"
                    EvalExpression = Result
                    Exit Function
                End If
            End If
            
            'We are clear to increment the index
            Var.ArrayIndex = Var.ArrayIndex + 1
            
            Alpha(Var.ArrayIndex, Var.ArrayLevel) = Var.OneChar
        
        'Handle parenthesis
        '----------------------------
        ElseIf (Var.OneChar = "(") Then
        
            'Check for faulty conditions
            If (Var.PreviousChar = ")") Then
                Result(0) = 10
                Result(1) = 6
                Result(2) = MathErrorCodes(6) '"Parenthesis pair with no operator"
                EvalExpression = Result
                Exit Function
            ElseIf ((Asc(Var.PreviousChar) > 47 And Asc(Var.PreviousChar) < 58)) Then
                Result(0) = 10
                Result(1) = 7
                Result(2) = MathErrorCodes(7) '"Missing operator between paren. and num."
                EvalExpression = Result
                Exit Function
            ElseIf (Var.PreviousChar = ".") Then
                Result(0) = 10
                Result(1) = 4
                Result(2) = MathErrorCodes(4) ' "Num. missing decimal part"
                EvalExpression = Result
                Exit Function
            End If
        
            'We are clear to increment the level
            Var.ArrayLevel = Var.ArrayLevel + 1 'You do not need to increment the other index here
            If (Var.ArrayLevelUbound < Var.ArrayLevel) Then
                Var.ArrayLevelUbound = Var.ArrayLevel
                ReDim Preserve Alpha(Len(ThisExpr) - 1, Var.ArrayLevelUbound) 'Expand the array
            End If
        ElseIf (Var.OneChar = ")") Then
            'Check for faulty conditions
            If (Var.PreviousChar = "(") Then
                Result(0) = 10
                Result(1) = 8
                Result(2) = MathErrorCodes(8) '"Empty parenthesis pair"
                EvalExpression = Result
                Exit Function
            ElseIf (InStr(1, Operators, Asc(Var.PreviousChar))) Then
                Result(0) = 10
                Result(1) = 9
                Result(2) = MathErrorCodes(9) '"Missing num. before closed parenthesis"
                EvalExpression = Result
                Exit Function
            ElseIf (Var.PreviousChar = ".") Then
                Result(0) = 10
                Result(1) = 4
                Result(2) = MathErrorCodes(4) '"Num. missing decimal part"
                EvalExpression = Result
                Exit Function
            End If
            
            If (Var.ArrayLevel = 0) Then
                Result(0) = 10
                Result(1) = 10
                Result(2) = MathErrorCodes(10) '"Excess close parenthesis"
                EvalExpression = Result
                Exit Function
            End If
            
            'We are clear to decrement the level
            Var.ArrayLevel = Var.ArrayLevel - 1 'You do not need to increment the index here
        Else
            Result(0) = 10
            Result(1) = 17
            Result(2) = MathErrorCodes(17) '"No alpha chars allowed"
            EvalExpression = Result
            Exit Function
        End If
    
        Var.PreviousChar = Var.OneChar
    Next
    
    'The level needs to be back to zero (otherwise we have parenthesis problems)
    If (Var.ArrayLevel > 0) Then
        Result(0) = 10
        Result(1) = 11
        Result(2) = Var.ArrayLevel & " " & MathErrorCodes(11) '"excess open parenthesis"
        EvalExpression = Result
        Exit Function
    End If
    
    'Run the simple solver
    '-------------------------------------
    For Count = Var.ArrayLevelUbound To 0 Step -1 'Start solving from the highest level
        'Operator "^"
        For count2 = 0 To Var.ArrayIndex
            If (Alpha(count2, Count) = "^") Then
                Var.PreviousNumIndex = -1
                Var.NextNumIndex = -1
                'Find the index location of the previous number
                For count3 = count2 - 1 To 0 Step -1
                    If (Alpha(count3, Count) <> "") Then
                        Var.PreviousNumIndex = count3
                        Exit For
                    End If
                Next
                If (Var.PreviousNumIndex < 0) Then
                    Result(0) = 10
                    Result(1) = 14
                    Result(2) = MathErrorCodes(14) '"Missing num. before operator"
                    EvalExpression = Result
                    Exit Function
                End If
                'Find the index location of the previous number
                For count3 = count2 + 1 To Var.ArrayIndex
                    If (Alpha(count3, Count) <> "") Then
                        Var.NextNumIndex = count3
                        Exit For
                    End If
                Next
                If (Var.NextNumIndex < 0) Then
                    Result(0) = 10
                    Result(1) = 15
                    Result(2) = MathErrorCodes(15) '"Stray operator at end"
                    EvalExpression = Result
                    Exit Function
                End If
                Var.Number = Alpha(Var.PreviousNumIndex, Count) ^ Alpha(Var.NextNumIndex, Count)
                Alpha(Var.PreviousNumIndex, Count) = ""
                Alpha(count2, Count) = Var.Number
                Alpha(Var.NextNumIndex, Count) = ""
            End If
        Next
        
        'Operator "/"
        For count2 = 0 To Var.ArrayIndex
            If (Alpha(count2, Count) = "/") Then
                Var.PreviousNumIndex = -1
                Var.NextNumIndex = -1
                'Find the index location of the previous number
                For count3 = count2 - 1 To 0 Step -1
                    If (Alpha(count3, Count) <> "") Then
                        Var.PreviousNumIndex = count3
                        Exit For
                    End If
                Next
                If (Var.PreviousNumIndex < 0) Then
                    Result(0) = 10
                    Result(1) = 14
                    Result(2) = MathErrorCodes(14) '"Missing num. before operator"
                    EvalExpression = Result
                    Exit Function
                End If
                'Find the index location of the previous number
                For count3 = count2 + 1 To Var.ArrayIndex
                    If (Alpha(count3, Count) <> "") Then
                        Var.NextNumIndex = count3
                        Exit For
                    End If
                Next
                If (Var.NextNumIndex < 0) Then
                    Result(0) = 10
                    Result(1) = 15
                    Result(2) = MathErrorCodes(15) '"Stray operator at end"
                    EvalExpression = Result
                    Exit Function
                End If
                'Check for division by zero
                If (Alpha(Var.NextNumIndex, Count) = 0) Then
                    Result(0) = 10
                    Result(1) = 12
                    Result(2) = MathErrorCodes(12) '"Division by zero"
                    EvalExpression = Result
                    Exit Function
                End If
                Var.Number = Alpha(Var.PreviousNumIndex, Count) / Alpha(Var.NextNumIndex, Count)
                Alpha(Var.PreviousNumIndex, Count) = ""
                Alpha(count2, Count) = Var.Number
                Alpha(Var.NextNumIndex, Count) = ""
            End If
        Next
        
        'Operator "*"
        For count2 = 0 To Var.ArrayIndex
            If (Alpha(count2, Count) = "*") Then
                Var.PreviousNumIndex = -1
                Var.NextNumIndex = -1
                'Find the index location of the previous number
                For count3 = count2 - 1 To 0 Step -1
                    If (Alpha(count3, Count) <> "") Then
                        Var.PreviousNumIndex = count3
                        Exit For
                    End If
                Next
                If (Var.PreviousNumIndex < 0) Then
                    Result(0) = 10
                    Result(1) = 14
                    Result(2) = MathErrorCodes(14) '"Missing num. before operator"
                    EvalExpression = Result
                    Exit Function
                End If
                'Find the index location of the previous number
                For count3 = count2 + 1 To Var.ArrayIndex
                    If (Alpha(count3, Count) <> "") Then
                        Var.NextNumIndex = count3
                        Exit For
                    End If
                Next
                If (Var.NextNumIndex < 0) Then
                    Result(0) = 10
                    Result(1) = 15
                    Result(2) = MathErrorCodes(15) '"Stray operator at end"
                    EvalExpression = Result
                    Exit Function
                End If
                Var.Number = Alpha(Var.PreviousNumIndex, Count) * Alpha(Var.NextNumIndex, Count)
                Alpha(Var.PreviousNumIndex, Count) = ""
                Alpha(count2, Count) = Var.Number
                Alpha(Var.NextNumIndex, Count) = ""
            End If
        Next
        
        'Operator "-"
        For count2 = 0 To Var.ArrayIndex
            If (Alpha(count2, Count) = "-") Then
                Var.PreviousNumIndex = -1
                Var.NextNumIndex = -1
                'Find the index location of the previous number
                For count3 = count2 - 1 To 0 Step -1
                    If (Alpha(count3, Count) <> "") Then
                        Var.PreviousNumIndex = count3
                        Exit For
                    End If
                Next
                'Find the index location of the previous number
                For count3 = count2 + 1 To Var.ArrayIndex
                    If (Alpha(count3, Count) <> "") Then
                        Var.NextNumIndex = count3
                        Exit For
                    End If
                Next
                If (Var.NextNumIndex < 0) Then
                    Result(0) = 10
                    Result(1) = 15
                    Result(2) = MathErrorCodes(15) '"Stray operator at end"
                    EvalExpression = Result
                    Exit Function
                End If
                If (Var.PreviousNumIndex >= 0) Then
                    Var.Number = Alpha(Var.PreviousNumIndex, Count) - Alpha(Var.NextNumIndex, Count)
                    Alpha(Var.PreviousNumIndex, Count) = ""
                Else
                    Var.Number = 0 - Alpha(Var.NextNumIndex, Count)
                End If
                Alpha(count2, Count) = Var.Number
                Alpha(Var.NextNumIndex, Count) = ""
            End If
        Next
        
        'Operator "+"
        For count2 = 0 To Var.ArrayIndex
            If (Alpha(count2, Count) = "+") Then
                Var.PreviousNumIndex = -1
                Var.NextNumIndex = -1
                'Find the index location of the previous number
                For count3 = count2 - 1 To 0 Step -1
                    If (Alpha(count3, Count) <> "") Then
                        Var.PreviousNumIndex = count3
                        Exit For
                    End If
                Next
                'Find the index location of the previous number
                For count3 = count2 + 1 To Var.ArrayIndex
                    If (Alpha(count3, Count) <> "") Then
                        Var.NextNumIndex = count3
                        Exit For
                    End If
                Next
                If (Var.NextNumIndex < 0) Then
                    Result(0) = 10
                    Result(1) = 15
                    Result(2) = MathErrorCodes(15) '"Stray operator at end"
                    EvalExpression = Result
                    Exit Function
                End If
                If (Var.PreviousNumIndex >= 0) Then
                    Var.Number = Alpha(Var.PreviousNumIndex, Count) + Alpha(Var.NextNumIndex, Count)
                    Alpha(Var.PreviousNumIndex, Count) = ""
                Else
                    Var.Number = 0 + Alpha(Var.NextNumIndex, Count)
                End If
                Alpha(count2, Count) = Var.Number
                Alpha(Var.NextNumIndex, Count) = ""
            End If
        Next
        
        For count2 = 0 To Var.ArrayIndex
            'Copy the numbers of the solved level, down one level
            If (Alpha(count2, Count) <> "" And Count > 0) Then
                Alpha(count2, Count - 1) = Alpha(count2, Count)
                Alpha(count2, Count) = ""
            ElseIf (Alpha(count2, Count) <> "" And Count = 0) Then
                'If we made it down here, then all is good. Give the result
                Result(0) = Alpha(count2, 0)
                Result(1) = 0
                Result(2) = MathErrorCodes(0)
            End If
        Next
    Next
    
    EvalExpression = Result
End Function

'-----------------------------------------------------------------------
'OBSOLETE CODE section--------------------------------------------------
'-----------------------------------------------------------------------

'This function is obsolete in view of the "EvalExpression" function above, but it is kept here for sentimental reasons.
'It performs expression syntax checks *without* actually solving the expression (it only predicts errors)
'It was useful for strict keystroke-level guarding of textboxes.
'If the proposed ascii is unacceptable it returns ascii=0 (at that point you can choose to "eat the keystroke" before it goes to the textbox)
Public Function AsciiAfterExpressionCheck(KeyAscii As Integer, InThisString As String, AtSelStart As Integer, Optional ToSelLength As Integer = 0) As Variant
    Dim AllowedSymbols As String
    Dim Count As Integer
    Dim count2 As Integer
    Dim StrPortion As String
    Dim PredictedText As String
    Dim CheckNum As NumberInfo
    Dim ResultAscii(2) As Variant

    AllowedSymbols = "-46, 40, 41, 42, 43, 45, 46, 47, 94, 13, 8" ' del ( ) * + - . / ^ Enter Backspace
    'Operators = "42, 43, 45, 47, 94"  '"*, +, -, /, ^"
    
    ResultAscii(0) = KeyAscii
    ResultAscii(1) = ""
    ResultAscii(2) = ""
    
    'check for illegal characters
    If (InStr(1, AllowedSymbols, CStr(KeyAscii), vbTextCompare) = 0 And (KeyAscii < 48 Or KeyAscii > 57)) Then
        ResultAscii(0) = 0
        ResultAscii(1) = "No alpha chars."
    End If
    TrackExpSyntax.LastKeyStroke = KeyAscii

    'Fineshed check for illegal characters.
    'From here on all characters are legal (but are they in proper order ? )
    If (ResultAscii(0) <> 0) Then
    'Find the characters before and after the current insertion point
    If (AtSelStart = 0) Then
        TrackExpSyntax.AsciiBeforeInsPoint = 0
        If (Len(InThisString) = 0 Or (AtSelStart + ToSelLength) = Len(InThisString)) Then
            TrackExpSyntax.AsciiAfterInsPoint = 0
        ElseIf ((AtSelStart + ToSelLength) < Len(InThisString)) Then
            TrackExpSyntax.AsciiAfterInsPoint = Asc(Mid(InThisString, AtSelStart + ToSelLength + 1, 1))
        End If
    ElseIf (AtSelStart + ToSelLength) = Len(InThisString) Then
        TrackExpSyntax.AsciiAfterInsPoint = 0
        TrackExpSyntax.AsciiBeforeInsPoint = Asc(Mid(InThisString, AtSelStart, 1))
    Else
        TrackExpSyntax.AsciiAfterInsPoint = Asc(Mid(InThisString, AtSelStart + ToSelLength + 1, 1))
        TrackExpSyntax.AsciiBeforeInsPoint = Asc(Mid(InThisString, AtSelStart, 1))
    End If
    'ResultAscii(1) = TrackExpSyntax.AsciiBeforeInsPoint & ", " & TrackExpSyntax.AsciiAfterInsPoint
            
    'Check if requested keystroke is allowable
    Select Case KeyAscii
        '-------------------------
        Case Asc(".") '46
            'Check before the insertion point
            For Count = AtSelStart To 1 Step -1
                'if not a digit then exit (a dot is not a digit, but we counted for it in the line above)
                If (Not (Asc(Mid(InThisString, Count, 1)) > 47 And Asc(Mid(InThisString, Count, 1)) < 58) And Mid(InThisString, Count, 1) <> ".") Then Exit For
                'Not number is found, then check to see if it is a dot
                If (Mid(InThisString, Count, 1) = ".") Then CheckNum.DecimalCount = CheckNum.DecimalCount + 1
                CheckNum.DigitsBeforeDecim = CheckNum.DigitsBeforeDecim + 1
            Next
            'Check after the insertion point
            For Count = (AtSelStart + ToSelLength) + 1 To Len(InThisString)
                'if not a digit then exit (a dot is not a digit, but we counted for it in the line above)
                If (Not (Asc(Mid(InThisString, Count, 1)) > 47 And Asc(Mid(InThisString, Count, 1)) < 58) And Mid(InThisString, Count, 1) <> ".") Then Exit For
                'Not number is found, then check to see if it is a dot
                If (Mid(InThisString, Count, 1) = ".") Then CheckNum.DecimalCount = CheckNum.DecimalCount + 1
                CheckNum.DigitsAfterDecim = CheckNum.DigitsAfterDecim + 1
            Next
            If (CheckNum.DecimalCount <> 0) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "Too many decimal separators"
            ElseIf (CheckNum.DigitsBeforeDecim = 0) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "Missing leading digits"
            End If
        '-------------------------
        Case Asc("+"), Asc("-"), Asc("/"), Asc("*"), Asc("^") '43, 45, 47, 42, 94
            If (AtSelStart = 0 And KeyAscii <> Asc("-")) Then
                ResultAscii(1) = "'" & Chr(KeyAscii) & "' not expected at start"
                ResultAscii(0) = 0
            ElseIf ((InStr(1, Operators, TrackExpSyntax.AsciiBeforeInsPoint, vbTextCompare) And TrackExpSyntax.AsciiBeforeInsPoint <> 0) _
                    Or (TrackExpSyntax.AsciiAfterInsPoint <> 0 And InStr(1, Operators, TrackExpSyntax.AsciiAfterInsPoint, vbTextCompare))) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "Too many operators"
            ElseIf (KeyAscii <> Asc("-") And TrackExpSyntax.AsciiBeforeInsPoint = Asc("(")) Then
                ResultAscii(1) = "Number expected after paren."
                ResultAscii(0) = 0
            ElseIf (TrackExpSyntax.AsciiBeforeInsPoint = Asc(".") Or TrackExpSyntax.AsciiAfterInsPoint = Asc(".")) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "Incomplete decimal"
            End If
        '-------------------------
        Case Asc("(")
            If (TrackExpSyntax.AsciiBeforeInsPoint = Asc(".") Or TrackExpSyntax.AsciiAfterInsPoint = Asc(".")) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "Incomplete decimal"
            ElseIf (TrackExpSyntax.AsciiBeforeInsPoint = Asc(")")) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "'(' not allowed here"
            ElseIf ((TrackExpSyntax.AsciiBeforeInsPoint > 47 And TrackExpSyntax.AsciiBeforeInsPoint < 58)) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "Missing operator before '('"
            ElseIf ((Not (TrackExpSyntax.AsciiAfterInsPoint > 47 And TrackExpSyntax.AsciiAfterInsPoint < 58)) _
                    And TrackExpSyntax.AsciiAfterInsPoint <> Asc("-") And TrackExpSyntax.AsciiAfterInsPoint <> Asc("(") And TrackExpSyntax.AsciiAfterInsPoint <> 0) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "Missing num. after parenthesis"
            End If
        '-------------------------
        Case Asc(")")
            If (AtSelStart = 0) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "')' not expected at start"
            ElseIf (TrackExpSyntax.AsciiBeforeInsPoint = Asc(".") Or TrackExpSyntax.AsciiAfterInsPoint = Asc(".")) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "Incomplete decimal"
            ElseIf (TrackExpSyntax.AsciiBeforeInsPoint = Asc("(")) Then
                ResultAscii(0) = 0
                ResultAscii(1) = "No empty parenthesis allowed"
            ElseIf (Not (TrackExpSyntax.AsciiBeforeInsPoint > 47 And TrackExpSyntax.AsciiBeforeInsPoint < 58) _
                    And TrackExpSyntax.AsciiBeforeInsPoint <> Asc(")")) Then
                ResultAscii(1) = "Num. expected instead of '" & Chr(KeyAscii) & "'"
                ResultAscii(0) = 0
            ElseIf ((TrackExpSyntax.AsciiAfterInsPoint > 47 And TrackExpSyntax.AsciiAfterInsPoint < 58)) Then
                ResultAscii(1) = "Operator expected"
                ResultAscii(0) = 0
            End If
        Case 48 To 57 'This maybe too restrictive
            'If (TrackExpSyntax.AsciiBeforeInsPoint = Asc(")")) Then
            '    ResultAscii(0) = 0
            '    ResultAscii(1) = "Operator Expected before number"
            'ElseIf (TrackExpSyntax.AsciiBeforeInsPoint = Asc("(")) Then
            '    ResultAscii(0) = 0
            '    ResultAscii(1) = "Operator Expected before number"
            'End If
    End Select
    End If
       
    'find the resulting open and closed parenthesis
    PredictedText = PredictTextChange(InThisString, KeyAscii, AtSelStart, ToSelLength)  'to see how the string will be affected
    TrackExpSyntax.OpenParenCount = 0
    TrackExpSyntax.CloseParenCount = 0
    TrackExpSyntax.PoorParenPair = 0
    For Count = 1 To Len(PredictedText)
        If (Mid(PredictedText, Count, 1) = "(") Then TrackExpSyntax.OpenParenCount = TrackExpSyntax.OpenParenCount + 1
        If (Mid(PredictedText, Count, 1) = ")") Then
            TrackExpSyntax.CloseParenCount = TrackExpSyntax.CloseParenCount + 1
            If (TrackExpSyntax.OpenParenCount < TrackExpSyntax.CloseParenCount) Then
                TrackExpSyntax.PoorParenPair = 1
            End If
        End If
    Next
    If (TrackExpSyntax.OpenParenCount - TrackExpSyntax.CloseParenCount > 0) Then
        ResultAscii(2) = (TrackExpSyntax.OpenParenCount - TrackExpSyntax.CloseParenCount) & " excess open parenthesis"
    ElseIf (TrackExpSyntax.OpenParenCount - TrackExpSyntax.CloseParenCount < 0) Then
        ResultAscii(2) = Abs(TrackExpSyntax.OpenParenCount - TrackExpSyntax.CloseParenCount) & " excess close parenthesis"
    ElseIf (TrackExpSyntax.PoorParenPair <> 0) Then
        ResultAscii(2) = "Malformed parenthesis pair(s)"
    Else
        ResultAscii(2) = ""
    End If
    
    AsciiAfterExpressionCheck = ResultAscii
End Function

'predict text after keypress and predict the characters before and after the new insertion point.
'This proceedure is also obsolete (it was used in conjunction with the "AsciiAfterExpressionCheck" above.
Public Function PredictTextChange(ThisString As String, KeyAscii As Integer, AtSelStart As Integer, Optional ToSelLength As Integer = 0) As Variant
    Dim StrPortion As String
    ThisString = Chr(0) & Chr(0) & ThisString & Chr(0) & Chr(0)
    AtSelStart = AtSelStart + 2
    
    If (KeyAscii = 8) Then 'Backspace
        If (ToSelLength > 0) Then
            TrackExpSyntax.AsciiAfterInsPoint = Asc(Mid(ThisString, AtSelStart + ToSelLength + 1))
            TrackExpSyntax.AsciiBeforeInsPoint = Asc(Mid(ThisString, AtSelStart))
            StrPortion = Mid(ThisString, AtSelStart + 1 + ToSelLength, Len(ThisString))
            ThisString = Left(ThisString, AtSelStart) & StrPortion
        Else
            TrackExpSyntax.AsciiAfterInsPoint = Asc(Mid(ThisString, AtSelStart + 1))
            TrackExpSyntax.AsciiBeforeInsPoint = Asc(Mid(ThisString, AtSelStart - 1))
            StrPortion = Mid(ThisString, AtSelStart + 1, Len(ThisString))
            ThisString = Left(ThisString, AtSelStart - 1) & StrPortion
        End If
    ElseIf (TrackExpSyntax.LastKeyCode = 46) Then 'delete key
        If (ToSelLength > 0) Then
            TrackExpSyntax.AsciiAfterInsPoint = Asc(Mid(ThisString, AtSelStart + ToSelLength + 1))
            TrackExpSyntax.AsciiBeforeInsPoint = Asc(Mid(ThisString, AtSelStart))
            StrPortion = Mid(ThisString, AtSelStart + 1 + ToSelLength, Len(ThisString))
            ThisString = Left(ThisString, AtSelStart) & StrPortion
        Else
            TrackExpSyntax.AsciiAfterInsPoint = Asc(Mid(ThisString, AtSelStart + 2))
            TrackExpSyntax.AsciiBeforeInsPoint = Asc(Mid(ThisString, AtSelStart))
            StrPortion = Mid(ThisString, AtSelStart + 2, Len(ThisString))
            ThisString = Left(ThisString, AtSelStart) & StrPortion
        End If
    Else 'The other keys
        If (ToSelLength > 0) Then
            TrackExpSyntax.AsciiAfterInsPoint = Asc(Mid(ThisString, AtSelStart + ToSelLength + 1))
            TrackExpSyntax.AsciiBeforeInsPoint = KeyAscii
            StrPortion = Mid(ThisString, AtSelStart + 1 + ToSelLength, Len(ThisString))
            ThisString = Left(ThisString, AtSelStart) & Chr(KeyAscii) & StrPortion
        Else
            TrackExpSyntax.AsciiAfterInsPoint = Asc(Mid(ThisString, AtSelStart + 1, 1))
            TrackExpSyntax.AsciiBeforeInsPoint = KeyAscii
            StrPortion = Mid(ThisString, AtSelStart + 1, Len(ThisString))
            ThisString = Left(ThisString, AtSelStart) & Chr(KeyAscii) & StrPortion
        End If
    End If
    
    ThisString = Replace(ThisString, Chr(0), "")
    PredictTextChange = ThisString
    
End Function
