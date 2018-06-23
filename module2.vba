vOption Compare Database

'Convert specified table/query to dictionary structure
Function asDictionary(strSQL As String) As Scripting.Dictionary
    Dim map As Scripting.Dictionary: Set map = New Scripting.Dictionary
    Dim rs As Recordset: Set rs = CurrentDb.OpenRecordset(strSQL)
    For Each field In getFields(rs)
        map.Add CStr(field), getFieldValues(rs, field)
    Next
    rs.Close
    Set rs = Nothing
    Set asDictionary = map
End Function

'Return fieldname values of specified recordset
Private Function getFields(rs As Recordset) As Variant()
    Dim myArray() As Variant
    ReDim myArray(1 To rs.fields.Count)

    Dim i As Integer: i = 1
    For Each field In rs.fields
        myArray(i) = field.Name
        i = i + 1
    Next field
    getFields = myArray
End Function

'Return field values of specified field as variant array
Private Function getFieldValues(rs As Recordset, field) As Variant()
On Error Resume Next 'empty field values raise errors

    Dim myArray() As Variant
    ReDim myArray(1 To rs.RecordCount)

    Dim i As Integer: i = 0
    rs.MoveFirst
    Do Until rs.EOF
        rs.MoveNext
        myArray(i) = rs(field).value
        i = i + 1
    Loop

    getFieldValues = myArray
End Function
