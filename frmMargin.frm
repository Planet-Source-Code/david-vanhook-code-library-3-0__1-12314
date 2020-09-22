VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMargin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   390
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.Slider Slider1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   529
      _Version        =   393216
      LargeChange     =   5000
      Min             =   1
      Max             =   50000
      SelStart        =   1
      TickFrequency   =   1000
      Value           =   1
      TextPosition    =   1
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4800
      ScaleHeight     =   735
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmMargin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  DoEvents
  Slider1.value = frmCodeLib.rtb1(0).RightMargin
  Me.Width = frmCodeLib.rtb1(0).Width - 90
  Slider1.Width = Me.Width - 150
  Me.Visible = True
  Me.Move frmCodeLib.rtb1(0).Left + 60 + Parent.Left + 60, frmCodeLib.rtb1(0).Top + 400 + Parent.Top, Me.Width, Me.Height
  On Error Resume Next
  Picture1.SetFocus
End Sub

Private Sub Form_Deactivate()
  Dim retval As Long
  Dim file As String
  Me.Visible = False
  
  If Right(App.Path, 1) <> "\" Then
    file = App.Path & "\" & "Library.ini"
  Else
    file = App.Path & "Library.ini"
  End If
  
  Dim value As String
  value = Slider1.value
  retval = WritePrivateProfileString("Preferences", "Right_Margin", value, file)
  Me.Hide
End Sub

Private Sub Form_Load()
  Slider1.value = frmCodeLib.rtb1(0).RightMargin
End Sub

Private Sub Slider1_Change()
  Slider1_Click
End Sub

Private Sub Slider1_Click()
  On Error Resume Next
  Slider1.ToolTipText = ""
  Picture1.SetFocus
  DoEvents
  
  For X = 0 To 9
    frmCodeLib.rtb1(X).RightMargin = Slider1.value
  Next
  
  Me.Refresh
End Sub

Private Sub Slider1_Scroll()
  Slider1_Click
End Sub
