VERSION 5.00
Begin VB.Form frmFolders 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Map folder to button"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "frmFolders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   3120
      Left            =   0
      TabIndex        =   2
      Top             =   -40
      Width           =   4815
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   2640
         MouseIcon       =   "frmFolders.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Default         =   -1  'True
         Height          =   375
         Left            =   3960
         MouseIcon       =   "frmFolders.frx":0594
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2640
         Width           =   735
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2415
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   2640
         TabIndex        =   1
         Top             =   150
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------
' frmFolders
'-----------
Dim DontClose As Boolean

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Return to the main form."
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Map this folder to one of the buttons."
End Sub

Private Sub Dir1_Change()
  File1 = Dir1
End Sub

Private Sub Drive1_Change()
  On Error GoTo ErrorHandler
  Dir1.Path = Drive1
  Exit Sub
    
ErrorHandler:
Call MsgBox("That drive is not available.", 48, "Error")
End Sub

Private Sub cmdExit_Click()
  DontClose = True
  Dir1.SetFocus
  Me.Hide
End Sub
'
'
'
'----------------SAVE FREQUENTLY USED FOLDER------------------'

Private Sub cmdSave_Click()
  On Error GoTo ErrorLog_Trapper

  If frmFolders.Caption = " Map [Main Folder]" Then
    frmPreferences.txtProject.Text = Dir1.Path
    Ini_File
    Dir1.SetFocus
    Me.Hide
    DontClose = True
    Exit Sub
  End If
  
  If frmFolders.Caption = " Map [Root Folder]" Then
    frmPreferences.txtRoot.Text = Dir1.Path
    Ini_File
    Dir1.SetFocus
    Me.Hide
    DontClose = True
    Exit Sub
  End If
  
  If frmFolders.Caption = " Map [Temp Folder]" Then
    frmPreferences.txtJunk.Text = Dir1.Path
    Ini_File
    Dir1.SetFocus
    Me.Hide
    DontClose = True
    Exit Sub
  End If
  
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("cmdSave_Click-frmFolders", Err.Description)
End Sub

Public Sub Ini_File()
  On Error Resume Next
  Dim retval As Long
  retval = 0
  filepath = App.Path
  
  If Right(filepath, 1) <> "\" Then
    filepath = App.Path & "\"
  End If
  
  file = filepath & "Library.ini"
  Dim Proj As String
  Proj = frmPreferences.txtProject.Text
  Dim Root As String
  Root = frmPreferences.txtRoot.Text
  Dim Junk As String
  Junk = frmPreferences.txtJunk.Text
  
  If frmPreferences.txtProject.Text <> "" Then
    retval = WritePrivateProfileString("Preferences", "ProjectFolder", Proj, file)
  End If
  
  If frmPreferences.txtRoot.Text <> "" Then
    retval = WritePrivateProfileString("Preferences", "RootDir", Root, file)
  End If
  
  If frmPreferences.txtJunk.Text <> "" Then
    retval = WritePrivateProfileString("Preferences", "FavFolder", Junk, file)
  End If
  
End Sub

Private Sub Form_Deactivate()
  On Error Resume Next
  Me.Hide
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Map frequently used folders to the three buttons."
End Sub
