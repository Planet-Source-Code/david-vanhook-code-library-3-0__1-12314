VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About..."
   ClientHeight    =   2325
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5415
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1604.757
   ScaleMode       =   0  'User
   ScaleWidth      =   5084.964
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   290
      Left            =   1080
      ScaleHeight     =   225
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   1080
      Width           =   3975
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code Library 3.0 is a handy utility used for viewing/saving commonly used pieces of text/code."
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   15
         Width           =   6675
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picIcon 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      MouseIcon       =   "frmAbout.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author: David VanHook"
      Height          =   195
      Left            =   1230
      TabIndex        =   6
      Top             =   761
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Be sure to check out my website:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   2355
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://home.earthlink.net/~zombiehead/"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      MouseIcon       =   "frmAbout.frx":09D6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2040
      Width           =   2880
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Library 3.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   4005
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  picIcon.SetFocus
  Unload Me
End Sub

Private Sub Command1_Click()
  picIcon.SetFocus
  Call MsgBox("I'm from Kalamazoo, Michigan.", 64)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblWeb.ForeColor = &H800000
End Sub

Private Sub lblWeb_Click()
  Set ie = New InternetExplorer
  ie.Visible = True
  ie.Navigate lblWeb.Caption
End Sub

Private Sub lblWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblWeb.ForeColor = &HFF0000
End Sub

Private Sub Timer1_Timer()
  lblDescription.Left = lblDescription.Left - 75
  
  If lblDescription.Left <= -6480 Then
    lblDescription.Left = 3500
  End If
  
End Sub
