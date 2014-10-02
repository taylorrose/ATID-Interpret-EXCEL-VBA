Attribute VB_Name = "atidInterpret"
'---------------------------------------------------------------------------------------------------------------------------------------------
'
'   ATID Interpret v1.0
'
'
'
'   Functions lists
'   ---------------
'
'       + Function atidInterpret(ATIDString As String, vlRange As Range, primaryColumn As Integer, secondaryColumn As Integer)
'           * Description : Use a lookup table to convert an ATID String, e.g., ****** to an ordered list of attribution names
'           * Specifications / limitations
'               - Multiline
'               - Not case sensitive
'           * Arguments
'               - ByVal txt As String : the text to search in
'               - ByVal matchPattern As String : the regular expression pattern
'               - ByVal replacePattern As String : the replacement pattern
'       
'       Revisions history
'       -----------------
'           - Taylor Rose        09/07/2014      v0.1        Creation
'
'---------------------------------------------------------------------------------------------------------------------------------
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


   
