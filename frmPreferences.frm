VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Options"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5085
   ControlBox      =   0   'False
   Icon            =   "frmPreferences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2010
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   4815
      Begin VB.ListBox listUsedFilters 
         Height          =   1425
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAllFiles 
         Caption         =   "Show All Files"
         Height          =   435
         Left            =   2640
         MouseIcon       =   "frmPreferences.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Use Default Filters"
         Height          =   435
         Left            =   2640
         MouseIcon       =   "frmPreferences.frx":0594
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   285
         Left            =   3600
         MouseIcon       =   "frmPreferences.frx":06E6
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   480
         Width           =   645
      End
      Begin VB.ListBox listFilters 
         Height          =   1425
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNewFilter 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   2640
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filters used:"
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add new filter:"
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available Filters:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2010
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   4815
      Begin VB.CommandButton cmdWindowsDefault 
         Caption         =   "..."
         Height          =   345
         Left            =   2280
         MouseIcon       =   "frmPreferences.frx":0838
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Select your own colors or use the Window's default color scheme for this program."
         Height          =   855
         Left            =   2520
         TabIndex        =   29
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   240
         MouseIcon       =   "frmPreferences.frx":098A
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   240
         MouseIcon       =   "frmPreferences.frx":0ADC
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Window's default"
         Height          =   315
         Left            =   2760
         MouseIcon       =   "frmPreferences.frx":0C2E
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Foreground Color"
         Height          =   315
         Left            =   720
         MouseIcon       =   "frmPreferences.frx":0D80
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Background Color"
         Height          =   315
         Left            =   720
         MouseIcon       =   "frmPreferences.frx":0ED2
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   480
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2010
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   4815
      Begin VB.TextBox txtJunk 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cmdSaveProject 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3360
         MouseIcon       =   "frmPreferences.frx":1024
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   950
         Width           =   375
      End
      Begin VB.TextBox txtProject 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdSaveRoot 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3360
         MouseIcon       =   "frmPreferences.frx":1176
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdSaveJunk 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3360
         MouseIcon       =   "frmPreferences.frx":12C8
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1525
         Width           =   375
      End
      Begin VB.TextBox txtRoot 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temp Folder"
         Height          =   195
         Left            =   375
         TabIndex        =   22
         Top             =   1575
         Width           =   885
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Root Folder"
         Height          =   195
         Left            =   100
         TabIndex        =   21
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Folder"
         Height          =   195
         Left            =   450
         TabIndex        =   20
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Browse"
         Height          =   195
         Left            =   4095
         TabIndex        =   19
         Top             =   975
         Width           =   525
      End
      Begin VB.Line Line1 
         X1              =   3960
         X2              =   3960
         Y1              =   480
         Y2              =   1680
      End
      Begin VB.Line Line2 
         X1              =   3720
         X2              =   4080
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line3 
         X1              =   3960
         X2              =   3720
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   3960
         X2              =   3720
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   360
      Picture         =   "frmPreferences.frx":141A
      ScaleHeight     =   420
      ScaleWidth      =   1950
      TabIndex        =   32
      Top             =   2880
      Width           =   1950
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   325
      Index           =   0
      Left            =   3960
      MouseIcon       =   "frmPreferences.frx":1940
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   2880
      Width           =   855
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   4471
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  File Filters  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Folder Options  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Colors  "
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "frmPreferences.frx":1A92
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4920
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   -120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5025
      TabIndex        =   33
      Top             =   2665
      Width           =   5090
   End
   Begin VB.Menu mnuRemove 
      Caption         =   "mnuremove"
      Visible         =   0   'False
      Begin VB.Menu mnuDelFilter 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------
' frmPreferences
'-----------

'----------------------------
Const Numtabs = 3
Dim DefaultPattern As String
Dim filepath As String
Dim file As String
Dim Temp1 As String * 75
Dim retval As Long
Dim bgcolor As String
Dim fgcolor As String
Dim DontClose As Boolean
Dim CurrentPos As Long
Dim holding As Long
Dim DontChange As Boolean
Dim GrandTotal As Integer

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Add the new filter to the list."
End Sub

Private Sub cmdOK_Click(Index As Integer)
  Me.Hide
End Sub

Private Sub cmdOk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Return to the main form."
  Label8.FontBold = False
  Label9.FontBold = False
  Label10.FontBold = False
  Label11.FontBold = False
  Label12.FontBold = False
  Label14.FontBold = False
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  frmFind.Hide
  Picture2.SetFocus
End Sub

Private Sub Form_Deactivate()
  On Error Resume Next
  Me.Hide
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorLog_Trapper

  For X = 0 To Numtabs - 1
  
    With Frame1(X)
      .BorderStyle = 1
      .Left = TabStrip1.ClientLeft
      .Top = TabStrip1.ClientTop - 5
      .Width = TabStrip1.ClientWidth
      .Height = TabStrip1.ClientHeight
      .Visible = False
    End With
  
  Next X

  TabStrip1.Tabs(1).Selected = True
  Frame1(TabStrip1.SelectedItem.Index - 1).Visible = True
  
  Get_Folders
  load_Filters
  txtNewFilter.Text = "*."
  txtNewFilter.SelStart = 3
  Label15.BackColor = frmCodeLib.rtb1(0).BackColor
  Label16.BackColor = frmCodeLib.rtb1(0).SelColor
  listUsedFilters.AddItem "Default"
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("Form_Load-frmPreferences", Err.Description)
End Sub

Private Sub Get_Folders()
  On Error GoTo ErrorLog_Trapper
  filepath = App.Path
  
  If Right(filepath, 1) <> "\" Then
    filepath = App.Path & "\"
  End If
  file = filepath & "Library.ini"
  
  'Projects
  Temp1 = ""
  retval = GetPrivateProfileString("Preferences", "ProjectFolder", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    txtProject.Text = Left(Temp1, retval)
  End If
  
  'Root
  Temp1 = ""
  retval = GetPrivateProfileString("Preferences", "RootDir", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    txtRoot.Text = Left(Temp1, retval)
  End If
  
  'On the fly folder [junk]
  Temp1 = ""
  retval = GetPrivateProfileString("Preferences", "FavFolder", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    txtJunk.Text = Left(Temp1, retval)
  End If
    
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("Get_Folders-frmPreferences", Err.Description)
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  DontChange = False
  Parent.StatusBar1.Panels(2).Text = "Define your preferences for this program."
  
  If Index = 1 Then
    Label8.FontBold = False
    Label9.FontBold = False
    Label10.FontBold = False
  End If
  
  If Index = 2 Then
    Label11.FontBold = False
    Label12.FontBold = False
    Label14.FontBold = False
  End If
  
End Sub

Private Sub Label11_Click()
  Label15_Click
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Specify the background color for the controls in this program."
End Sub

Private Sub Label12_Click()
  Label16_Click
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Specify the foreground color for the controls in this program."
End Sub

Private Sub Label14_Click()
  cmdWindowsDefault_Click
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Use Windows specified colors for the controls in this program."
End Sub

Private Sub Label15_Click()
  '-----------------
  'BACKGROUND COLOR
  '-----------------
  frmCodeLib.CommonDialog1.Flags = &H1
  frmCodeLib.CommonDialog1.Color = frmCodeLib.rtb1(0).BackColor
  frmCodeLib.CommonDialog1.ShowColor

  For X = 0 To 9
    frmCodeLib.rtb1(X).BackColor = frmCodeLib.CommonDialog1.Color
  Next
  
  frmCodeLib.Drive1.BackColor = frmCodeLib.CommonDialog1.Color
  frmCodeLib.Dir1.BackColor = frmCodeLib.CommonDialog1.Color
  frmCodeLib.File1.BackColor = frmCodeLib.CommonDialog1.Color
  frmCodeLib.txtFilename.BackColor = frmCodeLib.CommonDialog1.Color
  frmFind.txtFind.BackColor = frmCodeLib.CommonDialog1.Color
  frmFind.txtReplace.BackColor = frmCodeLib.CommonDialog1.Color
  frmFolders.Drive1.BackColor = frmCodeLib.CommonDialog1.Color
  frmFolders.Dir1.BackColor = frmCodeLib.CommonDialog1.Color
  frmFolders.File1.BackColor = frmCodeLib.CommonDialog1.Color
  listFilters.BackColor = frmCodeLib.CommonDialog1.Color
  listUsedFilters.BackColor = frmCodeLib.CommonDialog1.Color
  txtNewFilter.BackColor = frmCodeLib.CommonDialog1.Color
  txtProject.BackColor = frmCodeLib.CommonDialog1.Color
  txtRoot.BackColor = frmCodeLib.CommonDialog1.Color
  txtJunk.BackColor = frmCodeLib.CommonDialog1.Color
  frmProperties.txtFilename.BackColor = frmCodeLib.CommonDialog1.Color
  Ini_File
  bgcolor = frmCodeLib.rtb1(0).BackColor
  retval = WritePrivateProfileString("Preferences", "BackgroundColor", bgcolor, file)
  Label15.BackColor = bgcolor
End Sub

Private Sub Label16_Click()
  '-----------------
  'FOREGROUND COLOR
  '-----------------
  CurrentPos = frmCodeLib.rtb1(frmCodeLib.TabStrip1.SelectedItem.Index - 1).SelStart
  frmCodeLib.CommonDialog1.Flags = &H1
  frmCodeLib.CommonDialog1.Color = Label16.BackColor 'frmCodeLib.rtb1(0).SelColor
  frmCodeLib.CommonDialog1.ShowColor
  
  For X = 0 To 9
    frmCodeLib.rtb1(X).SelStart = 0
    frmCodeLib.rtb1(X).SelLength = Len(frmCodeLib.rtb1(X).Text)
    frmCodeLib.rtb1(X).SelColor = frmCodeLib.CommonDialog1.Color
    frmCodeLib.rtb1(X).SelLength = 0
    frmCodeLib.rtb1(X).SelStart = 0
  Next
  
  frmCodeLib.Drive1.ForeColor = frmCodeLib.CommonDialog1.Color
  frmCodeLib.Dir1.ForeColor = frmCodeLib.CommonDialog1.Color
  frmCodeLib.File1.ForeColor = frmCodeLib.CommonDialog1.Color
  frmCodeLib.txtFilename.ForeColor = frmCodeLib.CommonDialog1.Color
  frmFind.txtFind.ForeColor = frmCodeLib.CommonDialog1.Color
  frmFind.txtReplace.ForeColor = frmCodeLib.CommonDialog1.Color
  frmFolders.Drive1.ForeColor = frmCodeLib.CommonDialog1.Color
  frmFolders.Dir1.ForeColor = frmCodeLib.CommonDialog1.Color
  frmFolders.File1.ForeColor = frmCodeLib.CommonDialog1.Color
  listFilters.ForeColor = frmCodeLib.CommonDialog1.Color
  listUsedFilters.ForeColor = frmCodeLib.CommonDialog1.Color
  txtNewFilter.ForeColor = frmCodeLib.CommonDialog1.Color
  txtProject.ForeColor = frmCodeLib.CommonDialog1.Color
  txtRoot.ForeColor = frmCodeLib.CommonDialog1.Color
  txtJunk.ForeColor = frmCodeLib.CommonDialog1.Color
  frmProperties.txtFilename.ForeColor = frmCodeLib.CommonDialog1.Color
  Ini_File
  fgcolor = frmCodeLib.rtb1(0).SelColor
  retval = WritePrivateProfileString("Preferences", "ForegroundColor", fgcolor, file)
  Label16.BackColor = fgcolor
  frmCodeLib.rtb1(frmCodeLib.TabStrip1.SelectedItem.Index - 1).SelStart = CurrentPos
End Sub

Private Sub cmdWindowsDefault_Click()
  CurrentPos = frmCodeLib.rtb1(frmCodeLib.TabStrip1.SelectedItem.Index - 1).SelStart

  For X = 0 To 9
    frmCodeLib.rtb1(X).SelStart = 0
    frmCodeLib.rtb1(X).SelLength = Len(frmCodeLib.rtb1(X).Text)
    frmCodeLib.rtb1(X).BackColor = &H80000005
    frmCodeLib.rtb1(X).SelColor = &H80000012
    frmCodeLib.rtb1(X).SelLength = 0
  Next
  
  frmCodeLib.Drive1.BackColor = &H80000005
  frmCodeLib.Dir1.BackColor = &H80000005
  frmCodeLib.File1.BackColor = &H80000005
  frmCodeLib.txtFilename.BackColor = &H80000005
  frmFind.txtFind.BackColor = &H80000005
  frmFind.txtReplace.BackColor = &H80000005
  frmFolders.Drive1.BackColor = &H80000005
  frmFolders.Dir1.BackColor = &H80000005
  frmFolders.File1.BackColor = &H80000005
  listFilters.BackColor = &H80000005
  listUsedFilters.BackColor = &H80000005
  txtNewFilter.BackColor = &H80000005
  txtProject.BackColor = &H80000005
  txtRoot.BackColor = &H80000005
  txtJunk.BackColor = &H80000005
  frmProperties.txtFilename.ForeColor = &H80000005
  frmCodeLib.Drive1.ForeColor = &H80000012
  frmCodeLib.Dir1.ForeColor = &H80000012
  frmCodeLib.File1.ForeColor = &H80000012
  frmCodeLib.txtFilename.ForeColor = &H80000012
  frmFind.txtFind.ForeColor = &H80000012
  frmFind.txtReplace.ForeColor = &H80000012
  frmFolders.Drive1.ForeColor = &H80000012
  frmFolders.Dir1.ForeColor = &H80000012
  frmFolders.File1.ForeColor = &H80000012
  listFilters.ForeColor = &H80000012
  listUsedFilters.ForeColor = &H80000012
  txtNewFilter.ForeColor = &H80000012
  txtProject.ForeColor = &H80000012
  txtRoot.ForeColor = &H80000012
  txtJunk.ForeColor = &H80000012
  frmProperties.txtFilename.ForeColor = &H80000012
  Ini_File
  bgcolor = frmCodeLib.rtb1(0).BackColor
  fgcolor = frmCodeLib.rtb1(0).SelColor
  retval = 0
  retval = WritePrivateProfileString("Preferences", "BackgroundColor", bgcolor, file)
  retval = WritePrivateProfileString("Preferences", "ForegroundColor", fgcolor, file)
  Label15.BackColor = bgcolor
  Label16.BackColor = fgcolor
  frmCodeLib.rtb1(frmCodeLib.TabStrip1.SelectedItem.Index - 1).SelStart = CurrentPos
  TabStrip1_Click
End Sub

Private Sub listFilters_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Double-click to add to the filters used list.  Right-click to remove a filter."
End Sub

Private Sub listUsedFilters_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Filters that you are currently using.  Double-click to remove the filter."
End Sub

Private Sub mnuHide_Click()
  Parent.Hide
End Sub

Private Sub mnuShow_Click()
  Parent.Show
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label8.FontBold = False
  Label9.FontBold = False
  Label10.FontBold = False
  Label11.FontBold = False
  Label12.FontBold = False
  Label14.FontBold = False
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label8.FontBold = False
  Label9.FontBold = False
  Label10.FontBold = False
  Label11.FontBold = False
  Label12.FontBold = False
  Label14.FontBold = False
End Sub

Private Sub TabStrip1_Click()
  On Error Resume Next
  Dim i As Integer
  Dim Y As Integer
  Y = TabStrip1.SelectedItem.Index - 1
  
  For i = 0 To TabStrip1.Tabs.Count - 1
    
    If i = Y Then
      Frame1(i).Visible = True
      Frame1(i).Enabled = True
      Frame1(i).ZOrder 0
      Picture2.SetFocus
    Else
      Frame1(i).Visible = False
      Frame1(i).Enabled = False
    End If
    
  Next
  
End Sub
'
'
'
'---------------------FILTERS---------------------'
'-----------
' frmFilter
'-----------


Private Sub cmdAdd_Click()
  txtNewFilter.SetFocus
  
  If txtNewFilter.Text <> "*." And txtNewFilter.Text <> "" Then
    listFilters.AddItem txtNewFilter.Text
  End If
  
  txtNewFilter.Text = "*."
  txtNewFilter.SelStart = Len(txtNewFilter.Text) + 1
End Sub

Private Sub cmdAllFiles_Click()
  listUsedFilters.Clear
  frmCodeLib.File1.Pattern = "*.*"
  listUsedFilters.AddItem "*.*"
  listFilters.Clear
  load_Filters
  listFilters.SetFocus
  frmCodeLib.File1.Refresh
End Sub

Private Sub cmdAllFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Every file will be listed."
End Sub

Private Sub cmdDefault_Click()
  listUsedFilters.Clear
  listFilters.SetFocus
  listFilters.Clear
  load_Filters
  If listUsedFilters.List(0) = "Default" Then Exit Sub
  listUsedFilters.AddItem "Default"
  Dim Holdname As String
  Holdname = frmCodeLib.File1.FileName
  frmCodeLib.File1.Pattern = "*.txt;*.htm;*.frm;*.ctl;*.bas;*.ini;*.tmp;*.log;*.cls;*.bat;*.vbp"
  frmCodeLib.File1.Refresh
  
  For X = 0 To frmCodeLib.File1.ListCount - 1
    If frmCodeLib.File1.List(X) = Holdname Then
      frmCodeLib.File1.Selected(X) = True
    End If
  Next
  
End Sub

Private Sub cmdDefault_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Clear list of filters used and use default."
End Sub

Private Sub load_Filters()
  listFilters.AddItem "*.txt"
  listFilters.AddItem "*.htm"
  listFilters.AddItem "*.frm"
  listFilters.AddItem "*.bas"
  listFilters.AddItem "*.ctl"
  listFilters.AddItem "*.cls"
  listFilters.AddItem "*.ini"
  listFilters.AddItem "*.log"
  listFilters.AddItem "*.tmp"
  listFilters.AddItem "*.bat"
  listFilters.AddItem "*.vbp"
End Sub

Private Sub listFilters_dblClick()
  
  If listUsedFilters.List(0) = "Default" Then
    listUsedFilters.RemoveItem (0)
  End If
  
  If listUsedFilters.List(0) = "*.*" Then
    listUsedFilters.RemoveItem (0)
  End If
  
  listUsedFilters.AddItem listFilters.List(listFilters.ListIndex)
  listFilters.RemoveItem listFilters.ListIndex
  GetPattern
End Sub

Private Sub listFilters_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 2 Then
    Me.PopupMenu mnuRemove
  End If
  
End Sub

Private Sub listUsedFilters_dblClick()
  
  If listUsedFilters.Text <> "Default" And listUsedFilters.Text <> "*.*" Then
    listFilters.AddItem listUsedFilters.List(listUsedFilters.ListIndex)
    listUsedFilters.RemoveItem listUsedFilters.ListIndex
    GetPattern
  End If
  
End Sub

Private Sub mnuDelFilter_Click()
  If listFilters.ListIndex = -1 Then Exit Sub
  
  If MsgBox("Remove this filter: " & listFilters.List(listFilters.ListIndex) & "?", 36, "Remove") = vbYes Then
    listFilters.RemoveItem listFilters.ListIndex
  End If
  
End Sub

Private Sub GetPattern()
  
  For X = 0 To listUsedFilters.ListCount - 1
    
    If X = 0 Then
      DefaultPattern = listUsedFilters.List(X)
    Else
      DefaultPattern = DefaultPattern & ";" & listUsedFilters.List(X)
    End If
    
  Next
    
  frmCodeLib.File1.Pattern = DefaultPattern
End Sub

'Folder Options

Private Sub cmdSaveRoot_Click()
  On Error Resume Next
  txtRoot.SetFocus
  frmFolders.Dir1.Path = txtRoot.Text
  frmFolders.Caption = " Map [Root Folder]"
  frmFolders.Show vbModal
End Sub

Private Sub cmdSaveJunk_Click()
  On Error Resume Next
  txtJunk.SetFocus
  frmFolders.Dir1.Path = txtJunk.Text
  frmFolders.Caption = " Map [Temp Folder]"
  frmFolders.Show vbModal
End Sub

Private Sub cmdSaveProject_Click()
  On Error Resume Next
  txtProject.SetFocus
  frmFolders.Dir1.Path = txtProject.Text
  frmFolders.Caption = " Map [Main Folder]"
  frmFolders.Show vbModal
End Sub

Private Sub cmdSaveProject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If DontChange = True Then Exit Sub
  Label8.FontBold = True
  Label9.FontBold = False
  Label10.FontBold = False
  Parent.StatusBar1.Panels(2).Text = "Map the Main Folder button to one of your folders."
  DontChange = True
End Sub

Private Sub cmdSaveRoot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If DontChange = True Then Exit Sub
  Label9.FontBold = True
  Label8.FontBold = False
  Label10.FontBold = False
  Parent.StatusBar1.Panels(2).Text = "Map the Root Folder button to one of your folders."
  DontChange = True
End Sub

Private Sub cmdSaveJunk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If DontChange = True Then Exit Sub
  Label10.FontBold = True
  Label9.FontBold = False
  Label8.FontBold = False
  Parent.StatusBar1.Panels(2).Text = "Map the Temp Folder button to one of your folders."
  DontChange = True
End Sub


Private Sub txtNewFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Add your own user-defined file filters."
End Sub
