
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
'           * Description : Use a lookup table to convert an ATID String, e.g., "28386:20220:28386:20203" to an ordered list, e.g.,  "1)Organic Search 2)Wireless 3)Organic Search 4)Unbranded"
'           * Specifications / limitations
'               -ATID in lookuptable must be formatted as string (text), not an integer (number)
'           * Arguments
'               ATIDString - The ATID colon deliminated string you'd like to interpet, e.g. , 28386:20220:20203 or b2
'               vlRange - The range of the lookup table you're using to interpret ATID, e.g., Sheet2!A:F 
'               primaryColumn - The number of the column you want to pass as your primary interpretation, e.g. , 6 
'               secondayColumn - The number of the column you want to pass as a fallback if the primary column you select is blank or "NULL" on certain rows, e.g. , 5
'
'
'       + Function ATID(ATIDstr As String, VRange As Range)
'           *Description : Similar to atidInterpret but a little simpler as doesn not produce ordered list
'           *Specification / limitations
'            - Does NOT require lookup table to be strings: Keep values as integers
'       
'       Revisions history
'       -----------------
'           - Taylor Rose        09/07/2014      v1.0       Creation
'           - Taylor Rose        10/09/2014      v1.01       Added Delim as parameter
'           - Taylor Rose        04/28/2015      v1.02      Added Simple Interprestion
'---------------------------------------------------------------------------------------------------------------------------------
Function atidInterpret(ATIDString As String, delim As String, vlRange As Range, primaryColumn As Integer, secondaryColumn As Integer)
Dim atidArray() As String
atidArray() = Split(ATIDString, delim)
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

Function ATID(ATIDstr As String, VRange As Range)

Dim Arr
Arr = Split(ATIDstr, ":")
ATID = Application.WorksheetFunction.VLookup(Val(Arr(0)), VRange, 2, 0)


For i = 1 To UBound(Arr)
    Dim res As Variant
    On Error Resume Next
    res = Application.WorksheetFunction.VLookup(Val(Arr(i)), VRange, 2, 0)
    If Err.Number = 0 Then
        ATID = ATID & " : " & res
    Else
        ATID = ATID & " : " & "Unknown"
    End If
Next i

End Function


   
