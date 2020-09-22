Attribute VB_Name = "mErrorLog"

Dim filepath As String

Function Log_Error(Source As String, Error As String)
  On Error GoTo oh_no
  
  If Right(App.Path, 1) <> "\" Then
    filepath = App.Path & "\"
  Else
    filepath = App.Path
  End If
  
  Open filepath & "ErrorLog.log" For Append As #1
  Print #1, " ' Description = " & Trim(Error)
  Print #1, " '      Source = " & Trim(Source)
  Print #1, " '        Date = " & Trim(Date)
  Print #1, " '        Time = " & Time
  Print #1, " ' ..........................."
  Print #1, " '     "
  Call MsgBox(Error, 16, "Error")
  Close #1
  Exit Function
  
oh_no:
Call MsgBox(Error, 16, "Error")
Exit Function
End Function

