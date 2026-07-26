' Removes every whole-word occurrence of any keyword in a supplied array
' from a string (case-insensitive), then collapses leftover whitespace.
' Standalone from RemoveDateReferences - call it separately, e.g.:
'
'   Dim kw() As Variant
'   kw = Array("draft", "confidential", "n/a")
'   result = RemoveKeywords(cellText, kw)
'
' or from a worksheet formula via a wrapper that hard-codes the array.

Option Explicit

Public Function RemoveKeywords(ByVal InputText As String, ByRef Keywords As Variant) As String
    Dim regEx As Object
    Dim result As String
    Dim i As Long

    result = InputText

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True

    For i = LBound(Keywords) To UBound(Keywords)
        If Len(Trim(Keywords(i))) > 0 Then
            regEx.Pattern = "\b" & EscapeRegexLiteral(CStr(Keywords(i))) & "\b"
            result = regEx.Replace(result, " ")
        End If
    Next i

    ' Collapse repeated whitespace and trim ends
    regEx.Pattern = "\s+"
    result = Trim(regEx.Replace(result, " "))

    RemoveKeywords = result
End Function

' Escapes regex special characters in a keyword so it is matched literally.
Private Function EscapeRegexLiteral(ByVal Text As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Pattern = "([.^$*+?()\[\]{}\\|\/])"
    EscapeRegexLiteral = regEx.Replace(Text, "\$1")
End Function
