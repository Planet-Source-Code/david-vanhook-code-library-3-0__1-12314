VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File Properties"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5355
      TabIndex        =   11
      Top             =   2280
      Width           =   5415
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   240
         Picture         =   "frmProperties.frx":0000
         ScaleHeight     =   420
         ScaleWidth      =   1950
         TabIndex        =   13
         Top             =   150
         Width           =   1950
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   325
         Index           =   1
         Left            =   4200
         MouseIcon       =   "frmProperties.frx":0526
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtFileName 
      Height          =   1215
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   75
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   -40
      Width           =   2535
      Begin VB.CheckBox Check1 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmProperties.frx":0678
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Read-Only"
         Height          =   255
         Left            =   1200
         MouseIcon       =   "frmProperties.frx":07CA
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Hidden"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmProperties.frx":091C
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "System"
         Height          =   255
         Left            =   1200
         MouseIcon       =   "frmProperties.frx":0A6E
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Archive"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmProperties.frx":0BC0
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   5415
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Modified:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   525
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File Size:"
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "File size:"
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   3840
      End
      Begin VB.Label Label3 
         Caption         =   "Modified:"
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   525
         Width           =   3840
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   4800
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   1920
      Width           =   135
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click() 'Normal
  On Error Resume Next
  Picture1.SetFocus
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Normal: The file has not been modified since it was created."
End Sub

Private Sub Check2_Click() 'Read-Only
  On Error Resume Next
  Picture1.SetFocus
End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Read-Only: The file cannot be changed or deleted."
End Sub

Private Sub Check3_Click() 'Hidden
  On Error Resume Next
  Picture1.SetFocus
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Hidden: The file is not supposed to be seen by the user."
End Sub

Private Sub Check4_Click() 'System
  On Error Resume Next
  Picture1.SetFocus
End Sub

Private Sub Check4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "System: The file is important to the operating system."
End Sub

Private Sub Check5_Click() 'Archive
  On Error Resume Next
  Picture1.SetFocus
End Sub

Private Sub Check5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Archive: The file has been modified since it was created."
End Sub

Private Sub cmdOK_Click(Index As Integer)
  On Error Resume Next
  Picture1.SetFocus
  Me.Hide
End Sub

Private Sub cmdOk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Return to the main form."
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Left(Parent.StatusBar1.Panels(2).Text, 4) <> "File" Then
    frmCodeLib.CountTheLines
  End If
  
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Left(Parent.StatusBar1.Panels(2).Text, 4) <> "File" Then
    frmCodeLib.CountTheLines
  End If
  
End Sub

Private Sub txtFileName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Left(Parent.StatusBar1.Panels(2).Text, 4) <> "File" Then
    frmCodeLib.CountTheLines
  End If
  
End Sub
