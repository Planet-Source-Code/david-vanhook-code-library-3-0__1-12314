Attribute VB_Name = "mRecentFiles"
Dim filepath As String

Public Sub Ini_File()
  filepath = App.Path
  
  If Right(filepath, 1) <> "\" Then
    filepath = App.Path & "\"
  End If
  
End Sub

Sub UpdateFileMenu(FileName)
  WriteRecentFiles (FileName)
  GetRecentFiles
End Sub

Sub GetRecentFiles()
  Dim retval As Long
  Dim key As String
  Dim i As Integer
  Dim j As String
  Dim file As String
  Dim IniString As String
  IniString = String(255, 0)
  Dim NewString As String
  Ini_File
  file = filepath & "Library.ini"
  DoEvents

  For i = 0 To 9
    key = "Recent" & i
    retval = 0
    IniString = String(255, 0)
    retval = GetPrivateProfileString("Recent Documents", key, " ", IniString, Len(IniString), file)
    NewString = Left(IniString, retval)

    If retval And Left(IniString, 1) <> " " Then
      frmCodeLib.MenuToolBar.Bands("subFile").Tools("file" & i + 1).Visible = True
      frmCodeLib.MenuToolBar.Bands("subFile").Tools("removeRecent").Visible = True
      frmCodeLib.MenuToolBar.Bands("subFile").Tools("file" & i + 1).Description = NewString

      If Len(NewString) > 30 Then
        frmCodeLib.MenuToolBar.Bands("subFile").Tools("file" & i + 1).Caption = "&" & i & "..." & Left(NewString, 2) & "..\" & Right(NewString, 25)
      Else
        frmCodeLib.MenuToolBar.Bands("subFile").Tools("file" & i + 1).Caption = "&" & i & "..." & NewString
      End If
      
    Else
      frmCodeLib.MenuToolBar.Bands("subFile").Tools("file" & i + 1).Visible = False
    End If

    frmCodeLib.MenuToolBar.RecalcLayout
  Next

End Sub

Sub WriteRecentFiles(OpenFileName As String)
  Dim retval As String
  Dim key As String
  Dim i As Integer
  Dim j As String
  Dim IniString As String
  Dim file As String
  Dim HoldValue As Long
  Dim ReturnedString(0 To 9) As String
  IniString = String(255, 0)
  Ini_File
  file = filepath & "Library.ini"
  For i = 9 To -1 Step -1

    If i = 9 Then
    'Get the last entry
    retval = GetPrivateProfileString("Recent Documents", "Recent9", "", IniString, Len(IniString), file)
    'Return the value of the last entry
    ReturnedString(9) = Left(IniString, retval)

      'If the FileName to save is the last one in the list...
      If ReturnedString(9) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent8", "", IniString, Len(IniString), file)
        ReturnedString(8) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent9", ReturnedString(8), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent7", "", IniString, Len(IniString), file)
        ReturnedString(7) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent8", ReturnedString(7), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent6", "", IniString, Len(IniString), file)
        ReturnedString(6) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent7", ReturnedString(6), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent5", "", IniString, Len(IniString), file)
        ReturnedString(5) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent6", ReturnedString(5), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent4", "", IniString, Len(IniString), file)
        ReturnedString(4) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent5", ReturnedString(4), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent3", "", IniString, Len(IniString), file)
        ReturnedString(3) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent4", ReturnedString(3), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
        ReturnedString(2) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent3", ReturnedString(2), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
        ReturnedString(1) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If

    If i = 8 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent8", "", IniString, Len(IniString), file)
      'Return the value of the 9th entry
      ReturnedString(8) = Left(IniString, retval)

      If ReturnedString(8) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent7", "", IniString, Len(IniString), file)
        ReturnedString(7) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent8", ReturnedString(7), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent6", "", IniString, Len(IniString), file)
        ReturnedString(6) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent7", ReturnedString(6), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent5", "", IniString, Len(IniString), file)
        ReturnedString(5) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent6", ReturnedString(5), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent4", "", IniString, Len(IniString), file)
        ReturnedString(4) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent5", ReturnedString(4), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent3", "", IniString, Len(IniString), file)
        ReturnedString(3) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent4", ReturnedString(3), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
        ReturnedString(2) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent3", ReturnedString(2), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
        ReturnedString(1) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If

    If i = 7 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent7", "", IniString, Len(IniString), file)
      'Return the value of the 8th entry
      ReturnedString(7) = Left(IniString, retval)

      If ReturnedString(7) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent6", "", IniString, Len(IniString), file)
        ReturnedString(6) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent7", ReturnedString(6), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent5", "", IniString, Len(IniString), file)
        ReturnedString(5) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent6", ReturnedString(5), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent4", "", IniString, Len(IniString), file)
        ReturnedString(4) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent5", ReturnedString(4), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent3", "", IniString, Len(IniString), file)
        ReturnedString(3) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent4", ReturnedString(3), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
        ReturnedString(2) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent3", ReturnedString(2), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
        ReturnedString(1) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If

    If i = 6 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent6", "", IniString, Len(IniString), file)
      'Return the value of the 7th entry
      ReturnedString(6) = Left(IniString, retval)

      If ReturnedString(6) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent5", "", IniString, Len(IniString), file)
        ReturnedString(5) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent6", ReturnedString(5), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent4", "", IniString, Len(IniString), file)
        ReturnedString(4) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent5", ReturnedString(4), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent3", "", IniString, Len(IniString), file)
        ReturnedString(3) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent4", ReturnedString(3), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
        ReturnedString(2) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent3", ReturnedString(2), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
        ReturnedString(1) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If

    If i = 5 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent5", "", IniString, Len(IniString), file)
      'Return the value of the 6th entry
      ReturnedString(5) = Left(IniString, retval)

      If ReturnedString(5) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent4", "", IniString, Len(IniString), file)
        ReturnedString(4) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent5", ReturnedString(4), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent3", "", IniString, Len(IniString), file)
        ReturnedString(3) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent4", ReturnedString(3), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
        ReturnedString(2) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent3", ReturnedString(2), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
        ReturnedString(1) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If

    If i = 4 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent4", "", IniString, Len(IniString), file)
      'Return the value of the 5th entry
      ReturnedString(4) = Left(IniString, retval)

      If ReturnedString(4) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent3", "", IniString, Len(IniString), file)
        ReturnedString(3) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent4", ReturnedString(3), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
        ReturnedString(2) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent3", ReturnedString(2), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
        ReturnedString(1) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If

    If i = 3 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent3", "", IniString, Len(IniString), file)
      'Return the value of the 4th entry
      ReturnedString(3) = Left(IniString, retval)

      If ReturnedString(3) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
        ReturnedString(2) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent3", ReturnedString(2), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
        ReturnedString(1) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If

    If i = 2 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
      'Return the value of the 3rd entry
      ReturnedString(2) = Left(IniString, retval)

      If ReturnedString(2) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
        ReturnedString(1) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If

    If i = 1 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
      'Return the value of the 2nd entry
      ReturnedString(1) = Left(IniString, retval)

      If ReturnedString(1) = OpenFileName Then
        retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
        ReturnedString(0) = Left(IniString, retval)
        retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
        'Write to first
        retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
        Exit Sub
      End If

    End If


    If i = 0 Then
      retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
      'Return the value of the 1st entry
      ReturnedString(0) = Left(IniString, retval)

      If ReturnedString(0) = OpenFileName Then
        Exit Sub
      End If

    End If

  Next

  'The filename is not on the list.  Move other filenames down and
  'write the selected filename to the first entry in the file.
  retval = GetPrivateProfileString("Recent Documents", "Recent8", "", IniString, Len(IniString), file)
  ReturnedString(8) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent9", ReturnedString(8), file)
  retval = GetPrivateProfileString("Recent Documents", "Recent7", "", IniString, Len(IniString), file)
  ReturnedString(7) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent8", ReturnedString(7), file)
  retval = GetPrivateProfileString("Recent Documents", "Recent6", "", IniString, Len(IniString), file)
  ReturnedString(6) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent7", ReturnedString(6), file)
  retval = GetPrivateProfileString("Recent Documents", "Recent5", "", IniString, Len(IniString), file)
  ReturnedString(5) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent6", ReturnedString(5), file)
  retval = GetPrivateProfileString("Recent Documents", "Recent4", "", IniString, Len(IniString), file)
  ReturnedString(4) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent5", ReturnedString(4), file)
  retval = GetPrivateProfileString("Recent Documents", "Recent3", "", IniString, Len(IniString), file)
  ReturnedString(3) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent4", ReturnedString(3), file)
  retval = GetPrivateProfileString("Recent Documents", "Recent2", "", IniString, Len(IniString), file)
  ReturnedString(2) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent3", ReturnedString(2), file)
  retval = GetPrivateProfileString("Recent Documents", "Recent1", "", IniString, Len(IniString), file)
  ReturnedString(1) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent2", ReturnedString(1), file)
  retval = GetPrivateProfileString("Recent Documents", "Recent0", "", IniString, Len(IniString), file)
  ReturnedString(0) = Left(IniString, retval)
  retval = WritePrivateProfileString("Recent Documents", "Recent1", ReturnedString(0), file)
  'Write to first
  retval = WritePrivateProfileString("Recent Documents", "Recent0", OpenFileName, file)
End Sub


