' Removes every whole-word occurrence of any keyword in a supplied array
' from a string (case-insensitive, punctuation-insensitive), then collapses
' leftover whitespace. Uses only native VBA - no VBScript.RegExp / external
' library reference required.
'
' Standalone from RemoveDateReferences - call it separately, e.g.:
'
'   Dim kw() As Variant
'   kw = Array("draft", "confidential", "n/a")
'   result = RemoveKeywords(cellText, kw)
'
' or from a worksheet formula via a wrapper that hard-codes the array.

Option Explicit

Public Function RemoveKeywords(ByVal InputText As String, ByRef Keywords As Variant) As String
    Dim normalizedInput As String
    Dim tokens() As String
    Dim normKeywords() As String
    Dim outputParts As String
    Dim rawToken As String
    Dim normToken As String
    Dim keepToken As String
    Dim isKeyword As Boolean
    Dim i As Long, j As Long

    ReDim normKeywords(LBound(Keywords) To UBound(Keywords))
    For j = LBound(Keywords) To UBound(Keywords)
        normKeywords(j) = NormalizeToken(CStr(Keywords(j)))
    Next j

    normalizedInput = Replace(InputText, vbTab, " ")
    normalizedInput = Replace(normalizedInput, vbCr, " ")
    normalizedInput = Replace(normalizedInput, vbLf, " ")

    tokens = Split(Trim(normalizedInput), " ")

    For i = LBound(tokens) To UBound(tokens)
        rawToken = tokens(i)
        If Len(Trim(rawToken)) > 0 Then
            normToken = NormalizeToken(rawToken)
            isKeyword = False

            For j = LBound(normKeywords) To UBound(normKeywords)
                If Len(normKeywords(j)) > 0 And normToken = normKeywords(j) Then
                    isKeyword = True
                    Exit For
                End If
            Next j

            If Not isKeyword Then
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

    RemoveKeywords = outputParts
End Function

' Lowercases and strips special characters so tokens and keywords compare
' the same way regardless of punctuation or letter case, e.g. "N/A," and
' "n/a" both normalize to "na".
Private Function NormalizeToken(ByVal s As String) As String
    NormalizeToken = LCase(StripSpecialChars(s))
End Function

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
