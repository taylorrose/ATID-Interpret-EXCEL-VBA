Attribute VB_Name = "Module1"
Function atidInterpret(ATIDString As String, vlRange As Range, primaryColumn As Integer, secondaryColumn As Integer)
Dim atidArray() As String
atidArray() = Split(ATIDString, ":")
ArrLength = UBound(atidArray)

Dim Vlook1 As String
Dim Vlook2 As String
Dim i As Integer

For i = 0 To ArrLength

On Error Resume Next
Err.Clear
    If Not Application.WorksheetFunction.VLookup(atidArray(i), vlRange, primaryColumn, 0) = 0 And Not Application.WorksheetFunction.VLookup(atidArray(i), vlRange, primaryColumn, 0) = "NULL" Then
        Vlook1 = Application.WorksheetFunction.VLookup(atidArray(i), vlRange, primaryColumn, 0)
    Else
         Vlook1 = Application.WorksheetFunction.VLookup(atidArray(i), vlRange, secondaryColumn, 0)
    End If
If Err.Number = 0 Then

    Vlook2 = Vlook2 & " " & i + 1 & ")" & Vlook1
Else
Vlook2 = Vlook2 & " " & i + 1 & ")" & "Unknown Source"
End If
Next i

atidInterpret = Vlook2


End Function
Function CountRegEx(ISIN As String, RegExpression As String)

    Dim RegEx As Object
    Set RegEx = CreateObject("vbscript.regexp")

    RegEx.Pattern = RegExpression
    RegEx.IgnoreCase = True
    RegEx.Global = True

    Dim Matches As Object
    Set Matches = RegEx.Execute(ISIN)

    CountRegEx = Matches.Count

End Function

Function ExtractRegEx(ByVal text As String, RegExpression As String) As String

Dim result As String
Dim allMatches As Object
Dim RE As Object
Set RE = CreateObject("vbscript.regexp")

RE.Pattern = RegExpression
RE.Global = True
RE.IgnoreCase = True
Set allMatches = RE.Execute(text)

If allMatches.Count <> 0 Then
    result = allMatches.Item(0).SubMatches.Item(0)
End If

ExtractRegEx = result

End Function



   
