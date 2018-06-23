Option Compare Database
Public Sub MailTest()
    Call sendEmail("mailtest@mail.com", "HELLO WORLD!")
End Sub

'Send email to specified email address
Public Sub sendEmail(emailAddress As String, msg As String, Optional subject As String = "No Subject", Optional copyName As String = "")
    Dim frmBody As String: frmBody = vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & AdditionalEmailMessage & vbCrLf
    Dim session As Object: Set session = CreateObject("notes.NotesSession")
    Dim username As String: username = session.username
    Dim mailDbName As String: mailDbName = Left$(username, 1) & Right$(username, (Len(username) - InStr(1, username, " "))) & ".nsf"
    Dim mailDb As Object: Set mailDb = session.getdatabase("", mailDbName)
    Dim fs As Object: Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Not mailDb.isopen Then
        mailDb.openmail
    End If
    Set mailDoc = mailDb.CREATEDOCUMENT
    
On Error GoTo NoSend
    With mailDoc
        .Form = "Memo"
        .sendto = emailAddress
        .copyto = IIf(IsNull(copyName), "Nobody", copyName)
        .subject = subject
        .Body = frmBody
        .SAVEMESSAGEONSEND = True
        .PostedDate = Now()
        .Send 0, emailAddress
    End With
    
NoSendExit:
    Set mailDb = Nothing
    Set mailDoc = Nothing
    Set session = Nothing
    Exit Sub
NoSend:
    MsgBox "Error sending mail. Check your Send Folder under Lotus Notes to see if the email was sent.", , "Problem Sending Email"
End Sub
