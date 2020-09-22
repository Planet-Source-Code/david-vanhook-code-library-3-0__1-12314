VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Map folder"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5355
      TabIndex        =   8
      Top             =   2280
      Width           =   5415
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Default         =   -1  'True
         Height          =   325
         Left            =   4200
         MouseIcon       =   "frmMap.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   325
         Index           =   1
         Left            =   3240
         MouseIcon       =   "frmMap.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   240
         Picture         =   "frmMap.frx":02A4
         ScaleHeight     =   420
         ScaleWidth      =   1950
         TabIndex        =   9
         Top             =   150
         Width           =   1950
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   -40
      Width           =   5415
      Begin VB.OptionButton Option1 
         Caption         =   "Temp Folder"
         Height          =   255
         Index           =   2
         Left            =   240
         MouseIcon       =   "frmMap.frx":07CA
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Main Folder"
         Height          =   255
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmMap.frx":091C
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Root Folder"
         Height          =   255
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmMap.frx":0A6E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblTemp 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Current folder:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1695
         Width           =   4650
      End
      Begin VB.Label lblMain 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Current folder:"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1095
         Width           =   4650
      End
      Begin VB.Label lblRoot 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Current folder:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   495
         Width           =   4650
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1680
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   840
      Width           =   135
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click(Index As Integer)
  Picture1.SetFocus
  frmPreferences.txtRoot.Text = lblRoot.Caption
  frmPreferences.txtProject.Text = lblMain.Caption
  frmPreferences.txtJunk.Text = lblTemp.Caption
  frmFolders.Ini_File
  Me.Hide
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Map the current folders."
End Sub

Private Sub Command2_Click()
  Picture1.SetFocus
  Me.Hide
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Return to the main form."
End Sub

Private Sub Form_Activate()
  Dim X As Integer
  lblRoot.Caption = frmPreferences.txtRoot.Text
  lblMain.Caption = frmPreferences.txtProject.Text
  lblTemp.Caption = frmPreferences.txtJunk.Text
  
  For X = 0 To 2
    Option1(X).value = False
  Next
  
  If lblRoot.Caption = frmCodeLib.Dir1.Path Then Option1(0).value = True
  If lblMain.Caption = frmCodeLib.Dir1.Path Then Option1(1).value = True
  If lblTemp.Caption = frmCodeLib.Dir1.Path Then Option1(2).value = True
  Me.Caption = " Map " & frmCodeLib.Dir1.Path & " to a folder."
End Sub

Private Sub Option1_Click(Index As Integer)
  On Error Resume Next
  
  Select Case Index
    Case Is = 0
      lblRoot.Caption = frmCodeLib.Dir1.Path
      lblMain.Caption = frmPreferences.txtProject.Text
      lblTemp.Caption = frmPreferences.txtJunk.Text
      lblRoot.FontBold = True
      lblMain.FontBold = False
      lblTemp.FontBold = False
    Case Is = 1
      lblMain.Caption = frmCodeLib.Dir1.Path
      lblRoot.Caption = frmPreferences.txtRoot.Text
      lblTemp.Caption = frmPreferences.txtJunk.Text
      lblMain.FontBold = True
      lblRoot.FontBold = False
      lblTemp.FontBold = False
    Case Is = 2
      lblTemp.Caption = frmCodeLib.Dir1.Path
      lblRoot.Caption = frmPreferences.txtRoot.Text
      lblMain.Caption = frmPreferences.txtProject.Text
      lblTemp.FontBold = True
      lblMain.FontBold = False
      lblRoot.FontBold = False
  End Select
  
  Picture1.SetFocus
End Sub
