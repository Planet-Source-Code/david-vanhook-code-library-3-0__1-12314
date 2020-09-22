VERSION 5.00
Begin VB.Form frmFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Find / Replace"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5475
      TabIndex        =   13
      Top             =   1790
      Width           =   5535
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   325
         Left            =   4320
         MouseIcon       =   "frmFind.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   240
         Picture         =   "frmFind.frx":0594
         ScaleHeight     =   420
         ScaleWidth      =   1950
         TabIndex        =   14
         Top             =   150
         Width           =   1950
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   0
      TabIndex        =   9
      Top             =   -40
      Width           =   5535
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   435
         Width           =   1935
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find (Next)"
         Height          =   315
         Left            =   2160
         MouseIcon       =   "frmFind.frx":0ABA
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtReplace 
         Height          =   285
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1125
         Width           =   1935
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         Height          =   315
         Left            =   2160
         MouseIcon       =   "frmFind.frx":0C0C
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace All"
         Height          =   315
         Left            =   2160
         MouseIcon       =   "frmFind.frx":0D5E
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "Undo"
         Height          =   315
         Left            =   2160
         MouseIcon       =   "frmFind.frx":0EB0
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Options"
         Height          =   1395
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   1575
         Begin VB.CheckBox Check2 
            Caption         =   "Whole word"
            Height          =   195
            Left            =   120
            MouseIcon       =   "frmFind.frx":1002
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Case sensitive"
            Height          =   195
            Left            =   120
            MouseIcon       =   "frmFind.frx":1154
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox checkTop 
            Caption         =   "Stay on top"
            Height          =   195
            Left            =   120
            MouseIcon       =   "frmFind.frx":12A6
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace with:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------
' frmFind
'-----------

Dim FoundPos As Long
Dim NewPos As Long
Dim FoundLine As Integer
Dim holding As String
Dim CaseSensitive As Boolean
Dim numwords As Long
Dim onTop As Boolean
Dim WholeWord As Boolean
Dim Y As Integer

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Only search for words that match exactly."
End Sub

Private Sub Check2_Click()
  
  If Check2.value = 1 Then
    WholeWord = True
  Else
    WholeWord = False
  End If
  
  txtFind.SetFocus
  NewPos = 0
End Sub

Private Sub Check1_Click()
  
  If Check1.value = 1 Then
    CaseSensitive = True
  Else
    CaseSensitive = False
  End If
  
  txtFind.SetFocus
  NewPos = 0
End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Find/replace only the whole word, if found."
End Sub

Private Sub checkTop_Click()
  
  If checkTop.value = 1 Then
    MakeWindowAlwaysTop frmFind.hWnd
    onTop = True
  Else
    MakeWindowNotTop frmFind.hWnd
    onTop = False
  End If
  
  txtFind.SetFocus
End Sub

Private Sub checkTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Stay on top of the main form.  Works until you bring up another form."
End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Find the first/next occurrence of this word in the current file."
End Sub

Private Sub cmdOK_Click()
  txtFind.SetFocus
  checkTop.value = 0
  MakeWindowNotTop frmFind.hWnd
  onTop = False
  Me.Hide
  On Error Resume Next
  frmCodeLib.rtb1(frmCodeLib.TabStrip1.SelectedItem.Index - 1).SetFocus
End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Return to the main form."
End Sub

Private Sub cmdReplace_Click()
  On Error GoTo ErrorLog_Trapper
  Y = frmCodeLib.TabStrip1.SelectedItem.Index - 1
  
  If frmCodeLib.rtb1(Y).Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  If txtReplace.Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  If txtFind.Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  Dim pString As Variant
  Dim KeyPos As Long
  Dim StrLen As Long
  Dim StartPos As Long
  Dim NewKeyWord As String
  StartPos = 1
  pString = frmCodeLib.rtb1(Y).Text
  StrLen = Len(pString)
  txtFind.SetFocus
  KeyPos = InStr(StartPos, pString, txtFind.Text)
  If KeyPos = 0 Then Exit Sub
  pString = Left(pString, KeyPos - 1) & txtReplace.Text & Mid(pString, KeyPos + Len(txtFind.Text), StrLen - KeyPos - Len(txtFind.Text) + 1)
  frmCodeLib.rtb1(Y).Text = pString
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("cmdReplace_Click-frmFind", Err.Description)
End Sub

Private Sub cmdReplace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Replace the first occurrence."
End Sub

Private Sub cmdReplaceAll_Click()
  On Error GoTo ErrorLog_Trapper
  txtFind.SetFocus
  
  If txtReplace.Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  If txtFind.Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  Y = frmCodeLib.TabStrip1.SelectedItem.Index - 1
  
  If frmCodeLib.rtb1(Y).Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  Me.Hide
  If MsgBox("Are you sure you want to replace all occurances?", 36, "Replace all?") = vbYes Then
    frmCodeLib.rtb1(Y).Text = replaceString(pString:=frmCodeLib.rtb1(Y).Text, Keyword:=txtFind.Text, NewKeyWord:=txtReplace.Text)
    Call MsgBox("Number of replacements: " & numwords, 32, "Done")
    numwords = 0
  End If
  
  Me.Show
  txtFind.SetFocus
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("cmdReplaceAll_Click-frmFind", Err.Description)
End Sub

Private Sub cmdReplaceAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Replace all occurrences."
End Sub

Private Sub cmdUndo_Click()
  On Error GoTo ErrorLog_Trapper
  
  If txtReplace.Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  If txtFind.Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  Y = frmCodeLib.TabStrip1.SelectedItem.Index - 1
  
  If frmCodeLib.rtb1(Y).Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  
  If InStr(frmCodeLib.rtb1(Y), txtReplace.Text) Then
    holding = txtReplace.Text
    txtReplace.Text = txtFind.Text
    txtFind.Text = holding
    cmdReplaceAll_Click
    holding = txtFind.Text
    txtFind.Text = txtReplace.Text
    txtReplace.Text = holding
  Else
    MakeWindowNotTop Parent.hWnd
    Call MsgBox("Search string not found.", 48, "Not found")
      
    If checkTop.value = 1 Then
      MakeWindowAlwaysTop frmFind.hWnd
    End If
      
    txtFind.SetFocus
  End If
  
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("cmdUndo_Click-frmFind", Err.Description)
End Sub

Private Sub cmdUndo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Returns the file to its previous state."
End Sub

Private Sub cmdFind_Click()
  On Error Resume Next
  Y = frmCodeLib.TabStrip1.SelectedItem.Index - 1
  
  If frmCodeLib.rtb1(Y).Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  If txtFind.Text = "" Then
    txtFind.SetFocus
    Exit Sub
  End If
  
  If NewPos = 0 Then
    frmCodeLib.rtb1(Y).SelStart = NewPos
  Else
    frmCodeLib.rtb1(Y).SelStart = NewPos ' - Len(txtFind.Text)
  End If
  
  If frmCodeLib.rtb1(Y).SelStart < 0 Then frmCodeLib.rtb1(Y).SelStart = 0
  
  
  '---------------------------
  'Whole word / Case sensitive
  '---------------------------
  If WholeWord = True And CaseSensitive = True Then
    FoundPos = frmCodeLib.rtb1(Y).Find(txtFind.Text, NewPos, , 6)
    
    If FoundPos = -1 Then
      txtFind.SetFocus
      MakeWindowNotTop Parent.hWnd
      Call MsgBox("Search string not found.", 48, "Not found")
      NewPos = 0
      
      If checkTop.value = 1 Then
        MakeWindowAlwaysTop frmFind.hWnd
      End If
      
      Exit Sub
    End If
      
    NewPos = FoundPos + Len(txtFind.Text)
  End If
  
  '---------------------------
  'Whole word
  '---------------------------
  If WholeWord = True And CaseSensitive = False Then
    FoundPos = frmCodeLib.rtb1(Y).Find(txtFind.Text, NewPos, , rtfWholeWord)
    
    If FoundPos = -1 Then
      txtFind.SetFocus
      MakeWindowNotTop Parent.hWnd
      Call MsgBox("Search string not found.", 48, "Not found")
      NewPos = 0
      
      If checkTop.value = 1 Then
        MakeWindowAlwaysTop frmFind.hWnd
      End If
      
      Exit Sub
    End If
      
    NewPos = FoundPos + Len(txtFind.Text)
  End If
  
  
  '---------------------------
  'Case sensitive
  '---------------------------
  If WholeWord = False And CaseSensitive = True Then
    FoundPos = frmCodeLib.rtb1(Y).Find(txtFind.Text, NewPos, , rtfMatchCase)
    
    If FoundPos = -1 Then
      txtFind.SetFocus
      MakeWindowNotTop Parent.hWnd
      Call MsgBox("Search string not found.", 48, "Not found")
      NewPos = 0
      
      If checkTop.value = 1 Then
        MakeWindowAlwaysTop frmFind.hWnd
      End If
      
      Exit Sub
    End If
      
    NewPos = FoundPos + Len(txtFind.Text)
  End If
  
  '---------------------------
  'Every match
  '---------------------------
  If WholeWord = False And CaseSensitive = False Then
    FoundPos = frmCodeLib.rtb1(Y).Find(txtFind.Text, NewPos)
    
    If FoundPos = -1 Then
      txtFind.SetFocus
      MakeWindowNotTop Parent.hWnd
      Call MsgBox("Search string not found.", 48, "Not found")
      NewPos = 0
      
      If checkTop.value = 1 Then
        MakeWindowAlwaysTop frmFind.hWnd
      End If
      
      Exit Sub
    End If
      
    NewPos = FoundPos + Len(txtFind.Text)
  End If
  
  txtFind.SetFocus
End Sub
'
'
'
'---------------------REPLACE FUNCTION------------------------'

Public Function replaceString(ByVal pString As Variant, _
ByVal Keyword As String, ByVal NewKeyWord As String) As String
  Dim StrLen       As Long
  Dim StrOut       As String
  Dim KeyPos       As Long
  Dim StrLenChange As Boolean
  Dim StartPos     As Long
  
  StartPos = 1
  StrLen = Len(pString)
  'pString is the text you are searching through
  'Keyword is the string you want to replace
  
  If CaseSensitive = True Then
  
    If Len(Keyword) <> Len(NewKeyWord) Then StrLenChange = True
      KeyPos = InStr(StartPos, pString, Keyword)
      
      While KeyPos > 0
        pString = Left(pString, KeyPos - 1) & NewKeyWord & Mid(pString, KeyPos + Len(Keyword), StrLen - KeyPos - Len(Keyword) + 1)
        If StrLenChange Then StrLen = Len(pString)
        StartPos = KeyPos + Len(NewKeyWord)
        KeyPos = InStr(StartPos, pString, Keyword)
        '-------------------------
        'Count the number of words
        '-------------------------
        numwords = numwords + 1
      Wend
  
    ' return the string
    replaceString = pString
    Exit Function
    
  Else
    'not case sensitive
    If Len(Keyword) <> Len(NewKeyWord) Then StrLenChange = True
    KeyPos = InStr(StartPos, UCase(pString), UCase(Keyword))
    
    While KeyPos > 0
      pString = Left(pString, KeyPos - 1) & NewKeyWord & Mid(pString, KeyPos + Len(Keyword), StrLen - KeyPos - Len(Keyword) + 1)
      If StrLenChange Then StrLen = Len(pString)
      StartPos = KeyPos + Len(NewKeyWord)
      KeyPos = InStr(StartPos, UCase(pString), UCase(Keyword))
      '-------------------------
      'Count the number of words
      '-------------------------
      numwords = numwords + 1
    Wend
    
    ' return the string
    replaceString = pString
    Exit Function
  
  End If
  
End Function

Private Sub Form_Activate()
  txtFind.SelStart = 0
  txtFind.SelLength = Len(txtFind.Text)
  txtReplace.SelStart = 0
  txtReplace.SelLength = Len(txtReplace.Text)
End Sub

Private Sub Form_Deactivate()
  On Error GoTo oh_no
  
  If checkTop.value = 0 Then
    Me.Hide
  End If
  
oh_no:
Exit Sub
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Find / Find Next / Replace"
End Sub

Private Sub optDown_Click()
  txtFind.SetFocus
End Sub

Private Sub optUp_Click()
  txtFind.SetFocus
End Sub
