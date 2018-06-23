Option Compare Database
'Return excel column-value of specified integer; [1, 26, 27] = [A, Z, AA]
Function column(v As Integer) As String
    Dim val As Integer: val = v - 1
    If val >= 0 And val < 26 Then
        column = Chr(Asc("A") + val)
    ElseIf val > 26 Then
        column = column(val / 26) + column((val Mod 26) + 1)
    Else
        MsgBox "Invalid Column #" & v
    End If
End Function

'Return string encapsulated by quotes
Function quote(val As String) As String
    quote = Chr(34) & val & Chr(34)
End Function

'Return true if testDate falls between startDate and endDate (inclusively)
Function BetweenDates(startDate As String, endDate As String, testDate As String) As Boolean
    BetweenDates = CDate(testDate) >= CDate(startDate) And CDate(testDate) <= CDate(endDate)
End Function


