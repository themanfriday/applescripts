' Removes month, day, and year references (in numeric, spelled-out, and
' abbreviated formats) plus any remaining special characters from a string.
' Uses only native VBA - no VBScript.RegExp / external library reference
' required, so it works even without access to the Windows Script Host.
'
' Paste into a standard VBA module (e.g. in Excel) and call:
'   =RemoveDateReferences(A1)
' from a worksheet, or RemoveDateReferences(someString) from other VBA code.

Option Explicit

Public Function RemoveDateReferences(ByVal InputText As String) As String
    Dim normalizedInput As String
    Dim tokens() As String
    Dim outputParts As String
    Dim rawToken As String
    Dim cleanToken As String
    Dim keepToken As String
    Dim i As Long

    normalizedInput = Replace(InputText, vbTab, " ")
    normalizedInput = Replace(normalizedInput, vbCr, " ")
    normalizedInput = Replace(normalizedInput, vbLf, " ")

    tokens = Split(Trim(normalizedInput), " ")

    For i = LBound(tokens) To UBound(tokens)
        rawToken = tokens(i)
        If Len(Trim(rawToken)) > 0 Then
            cleanToken = CleanEdges(rawToken)

            If IsApostropheYearToken(rawToken) Then
                ' e.g. '24 - skip, it's a date reference
            ElseIf IsMonthName(cleanToken) Then
                ' e.g. January / Jan - skip
            ElseIf IsDayName(cleanToken) Then
                ' e.g. Monday / Mon - skip
            ElseIf IsFourDigitYear(cleanToken) Then
                ' e.g. 2024 - skip
            ElseIf IsOrdinalDayNumber(cleanToken) Then
                ' e.g. 1st, 22nd, 31st - skip
            ElseIf IsNumericDateToken(cleanToken) Then
                ' e.g. 01/02/2024, 2024-01-02 - skip
            Else
                keepToken = StripSpecialChars(rawToken)
                If Len(keepToken) > 0 Then
                    If Len(outputParts) > 0 Then
                        outputParts = outputParts & " " & keepToken
                    Else
                        outputParts = keepToken
                    End If
                End If
            End If
        End If
    Next i

    RemoveDateReferences = outputParts
End Function

' Strips non-alphanumeric characters from the leading and trailing edges
' of a token only, leaving internal separators (e.g. "/" in "01/02/2024")
' intact so date-pattern checks can still see them.
Private Function CleanEdges(ByVal s As String) As String
    Do While Len(s) > 0 And Not IsAlphaNumericChar(Left(s, 1))
        s = Mid(s, 2)
    Loop
    Do While Len(s) > 0 And Not IsAlphaNumericChar(Right(s, 1))
        s = Left(s, Len(s) - 1)
    Loop
    CleanEdges = s
End Function

' Removes every non-alphanumeric character anywhere in the token.
Private Function StripSpecialChars(ByVal s As String) As String
    Dim i As Long, ch As String, result As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If IsAlphaNumericChar(ch) Then
            result = result & ch
        End If
    Next i
    StripSpecialChars = result
End Function

Private Function IsAlphaNumericChar(ByVal ch As String) As Boolean
    IsAlphaNumericChar = (ch Like "[A-Za-z0-9]")
End Function

Private Function IsDigitsOnly(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then
        IsDigitsOnly = False
        Exit Function
    End If
    For i = 1 To Len(s)
        If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then
            IsDigitsOnly = False
            Exit Function
        End If
    Next i
    IsDigitsOnly = True
End Function

Private Function IsInArray(ByVal val As String, ByRef arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If val = arr(i) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Private Function IsMonthName(ByVal s As String) As Boolean
    Dim months As Variant
    months = Array("january", "jan", "february", "feb", "march", "mar", _
                    "april", "apr", "may", "june", "jun", "july", "jul", _
                    "august", "aug", "september", "sept", "sep", "october", _
                    "oct", "november", "nov", "december", "dec")
    IsMonthName = IsInArray(LCase(s), months)
End Function

Private Function IsDayName(ByVal s As String) As Boolean
    Dim days As Variant
    days = Array("monday", "mon", "tuesday", "tue", "tues", "wednesday", _
                 "wed", "thursday", "thu", "thur", "thurs", "friday", "fri", _
                 "saturday", "sat", "sunday", "sun")
    IsDayName = IsInArray(LCase(s), days)
End Function

Private Function IsFourDigitYear(ByVal s As String) As Boolean
    If Len(s) = 4 And IsDigitsOnly(s) Then
        Dim yr As Long
        yr = CLng(s)
        IsFourDigitYear = (yr >= 1900 And yr <= 2099)
    End If
End Function

' Detects a leading apostrophe followed by exactly two digits, e.g. '24
' (checked on the raw token, before edge-trimming, since the apostrophe
' itself is the signal).
Private Function IsApostropheYearToken(ByVal raw As String) As Boolean
    Dim s As String
    s = raw
    Do While Len(s) > 0 And Not IsAlphaNumericChar(Right(s, 1))
        s = Left(s, Len(s) - 1)
    Loop
    If Len(s) = 3 And Left(s, 1) = "'" Then
        IsApostropheYearToken = IsDigitsOnly(Mid(s, 2, 2))
    End If
End Function

' Detects ordinal day-of-month numbers: 1st, 2nd, 3rd, 4th ... 31st
Private Function IsOrdinalDayNumber(ByVal s As String) As Boolean
    Dim lower As String, suffix As String, numPart As String, n As Long
    lower = LCase(s)
    If Len(lower) >= 3 Then
        suffix = Right(lower, 2)
        If suffix = "st" Or suffix = "nd" Or suffix = "rd" Or suffix = "th" Then
            numPart = Left(lower, Len(lower) - 2)
            If IsDigitsOnly(numPart) And Len(numPart) <= 2 Then
                n = CInt(numPart)
                IsOrdinalDayNumber = (n >= 1 And n <= 31)
            End If
        End If
    End If
End Function

' Detects full or partial numeric dates using "/", "-", or "." separators,
' e.g. 01/02/2024, 2024-01-02, 12.31.24.
Private Function IsNumericDateToken(ByVal s As String) As Boolean
    Dim seps As Variant
    Dim sepChar As Variant
    Dim parts() As String
    Dim i As Long
    Dim allNumeric As Boolean

    seps = Array("/", "-", ".")

    For Each sepChar In seps
        If InStr(s, sepChar) > 0 Then
            parts = Split(s, sepChar)
            If UBound(parts) >= 1 And UBound(parts) <= 2 Then
                allNumeric = True
                For i = LBound(parts) To UBound(parts)
                    If Len(parts(i)) = 0 Or Len(parts(i)) > 4 Or Not IsDigitsOnly(parts(i)) Then
                        allNumeric = False
                        Exit For
                    End If
                Next i
                If allNumeric Then
                    IsNumericDateToken = True
                    Exit Function
                End If
            End If
        End If
    Next sepChar

    IsNumericDateToken = False
End Function
