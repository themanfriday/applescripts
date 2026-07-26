' Removes month, day, and year references (in numeric, spelled-out, and
' abbreviated formats) plus any remaining special characters from a string.
' Paste into a standard VBA module (e.g. in Excel) and call:
'   =RemoveDateReferences(A1)
' from a worksheet, or RemoveDateReferences(someString) from other VBA code.

Option Explicit

Public Function RemoveDateReferences(ByVal InputText As String) As String
    Dim regEx As Object
    Dim result As String
    result = InputText

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True

    ' 1. Full and abbreviated month names (with optional trailing period)
    regEx.Pattern = "\b(jan(uary)?|feb(ruary)?|mar(ch)?|apr(il)?|may|jun(e)?|jul(y)?|aug(ust)?|sep(t(ember)?)?|oct(ober)?|nov(ember)?|dec(ember)?)\.?\b"
    result = regEx.Replace(result, " ")

    ' 2. Full and abbreviated day-of-week names
    regEx.Pattern = "\b(mon(day)?|tue(s(day)?)?|wed(nesday)?|thu(rs(day)?)?|fri(day)?|sat(urday)?|sun(day)?)\.?\b"
    result = regEx.Replace(result, " ")

    ' 3. Full numeric dates, any order/separator: mm/dd/yyyy, dd-mm-yy, yyyy.mm.dd, etc.
    regEx.Pattern = "\b\d{1,4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,4}\b"
    result = regEx.Replace(result, " ")

    ' 4. Partial numeric dates: yyyy-mm or mm-yyyy
    regEx.Pattern = "\b\d{4}[\/\-\.]\d{1,2}\b|\b\d{1,2}[\/\-\.]\d{4}\b"
    result = regEx.Replace(result, " ")

    ' 5. Ordinal day-of-month numbers: 1st, 2nd, 3rd, 4th ... 31st
    regEx.Pattern = "\b\d{1,2}(st|nd|rd|th)\b"
    result = regEx.Replace(result, " ")

    ' 6. Stand-alone four-digit years (1900-2099)
    regEx.Pattern = "\b(19|20)\d{2}\b"
    result = regEx.Replace(result, " ")

    ' 7. Two-digit years introduced by an apostrophe, e.g. '24
    regEx.Pattern = "'\d{2}\b"
    result = regEx.Replace(result, " ")

    ' 8. Any remaining special characters (anything not a letter, digit, or space)
    regEx.Pattern = "[^A-Za-z0-9\s]"
    result = regEx.Replace(result, " ")

    ' 9. Collapse repeated whitespace and trim ends
    regEx.Pattern = "\s+"
    result = Trim(regEx.Replace(result, " "))

    RemoveDateReferences = result
End Function
