VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ACTBAR.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCodeLib 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9180
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmCodeLib.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MousePointer    =   1  'Arrow
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9180
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6375
      Begin ActiveBarLibraryCtl.ActiveBar MenuToolBar 
         Left            =   0
         Tag             =   "Menu"
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Bands           =   "frmCodeLib.frx":0442
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "All files (*.*)|*.*|"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   8
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeLib.frx":7DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeLib.frx":829A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      BulletIndent    =   2
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":875E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Index           =   9
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      BulletIndent    =   2
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":8823
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3135
      Index           =   8
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      BulletIndent    =   2
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":88E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   415
      Left            =   0
      TabIndex        =   14
      Top             =   360
      Width           =   6375
      Begin ActiveBarLibraryCtl.ActiveBar MainToolBar 
         Left            =   0
         Tag             =   "ToolBar"
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Bands           =   "frmCodeLib.frx":89AD
      End
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawWidth       =   2
      Height          =   4095
      Left            =   6720
      ScaleHeight     =   4095
      ScaleWidth      =   2085
      TabIndex        =   10
      Top             =   960
      Width           =   2085
      Begin VB.PictureBox picSplitter 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         FillStyle       =   3  'Vertical Line
         Height          =   75
         Left            =   15
         ScaleHeight     =   32.66
         ScaleMode       =   0  'User
         ScaleWidth      =   21060
         TabIndex        =   17
         Top             =   2040
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtFilename 
         Height          =   315
         Left            =   0
         MousePointer    =   3  'I-Beam
         TabIndex        =   18
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Frame ButtonFrame 
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   15
         TabIndex        =   15
         Top             =   3600
         Width           =   2050
         Begin ActiveBarLibraryCtl.ActiveBar ButtonBar 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Bands           =   "frmCodeLib.frx":B16B
         End
      End
      Begin VB.FileListBox File1 
         Height          =   870
         Hidden          =   -1  'True
         Left            =   0
         Pattern         =   "*.txt*;*.htm*;*.frm;*.ctl;*.bas;*.ini;*.log;*.tmp;*.vbp"
         System          =   -1  'True
         TabIndex        =   13
         Top             =   2160
         Width           =   2055
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   0
         TabIndex        =   12
         Top             =   480
         Width           =   2040
      End
      Begin VB.Image imgSplitter 
         Height          =   75
         Left            =   0
         MousePointer    =   7  'Size N S
         Top             =   2040
         Width           =   2040
      End
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":B86D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Index           =   2
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":B932
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Index           =   3
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":B9F7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Index           =   4
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":BABC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Index           =   5
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":BB81
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Index           =   6
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":BC46
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3015
      Index           =   7
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MousePointer    =   3
      RightMargin     =   50000
      TextRTF         =   $"frmCodeLib.frx":BD0B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4575
      Left            =   0
      TabIndex        =   16
      Top             =   840
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   8070
      Placement       =   1
      Separators      =   -1  'True
      TabMinWidth     =   970
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code Library"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCodeLib.frx":BDD0
   End
   Begin ActiveBarLibraryCtl.ActiveBar PopUpToolBar 
      Left            =   120
      Tag             =   "Pop-up menu"
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmCodeLib.frx":BF32
   End
End
Attribute VB_Name = "frmCodeLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'####################################################################'
'## ##  -----------------------------------------------------   ## ##'
'## ##  Program name: Code Library                              ## ##'
'## ##       Started: November, 1999                            ## ##'
'## ##   Last Update: October, 2000                             ## ##'
'## ##        Author: David VanHook                             ## ##'
'## ##        E-Mail: zombiehead@earthlink.net                  ## ##'
'## ##       Website: http://home.earthlink.net/~zombiehead/    ## ##'
'## ##      Location: Kalamazoo, Michigan - USA                 ## ##'
'## ##  -----------------------------------------------------   ## ##'
'####################################################################'
'
'
'   This program was designed to hold pieces of code
'   from Visual Basic projects so they could be easily
'   referenced and copied into other projects.  I wanted
'   to create a program that could save code and view it
'   quickly to save time from having to write the same
'   code repeatedly.  I also used the ActiveBar control
'   because of how great the menus look compared to the
'   default gray ones that most programs use.
'
'
'
'---------------VARIABLE SECTION------------------'

Dim Numtabs As Integer         'Number of tabs on tabstrip (changes throughout program)
Dim X As Integer               'For/Next loop variable
Dim Y As Integer               'Variable to condense code
Dim Z As Integer               'Another looping variable
Dim lengthOfName As Long       'Length of filename (InStrRev) pulled from full filepath
Dim linecount As Long          'Number of lines in snippit used by line counter module
Dim filepath As String         'Full filepath of file to open
Dim answer As String           'MsgBox return variable
Dim TextToSave As String       'Contents of current RichTextBox to save
Dim DeSelect As Boolean        'Highlight filename in listbox
Dim DontMove As Boolean        'Avoid file1_click procedure when highlighting filename
Dim DontCount As Boolean       'Avoid the line counter procedure (font/color selection)
Dim ClearAll As Boolean        'Used to avoid certain procedures when clearing all
Dim CurrentPos As Long         'Current cursor position in text
Dim HowMany As Integer         'How many files are currently open (Window)
Dim HelpPath As String
'
'
'
'---------------INI FILE VARIABLES----------------'

Dim retval As Long              'Return variable for ini file procedure
Dim Temp1 As String * 75        'Variable saved in ini file
Dim bgcolor As String           'BackgroundColor
Dim fgcolor As String           'ForegroundColor
Dim fontname1 As String          'fontname1
Dim Fsize As String             'FontSize
Dim Fbold As String             'FontBold
Dim Fitalic As String           'FontItalic
Dim FullView As String          'FullView on/off
Dim ToolsOn As String           'ToolBarOn on/off
Dim toolTipsOn As String        'ToolTipsOn on/off
Dim StatusOn As String          'StatusBar on/off
Dim ProjectFolder As String     'ProjectFolder
Dim RootDir As String           'RootDir
Dim FavFolder As String         'FavFolder
Dim file As String              'Full path of .ini file
Dim key As String               'Subsection of .ini file
Dim getnumPages As Long         'Return variable for .ini file procedures
Dim mbMoving As Boolean         'If user is moving the splitter - dir/file listboxes
Const sglSplitLimit = 800       'Highest position the splitter can go(from top of picturebox)
Const cdlCCRGBInit = &H1        'Initial RGB color for commondialog1
'
'
'
'-------------------FORM LOAD---------------------'

Private Sub Form_Load()
  On Error GoTo ErrorLog_Trapper
  Numtabs = 10
  GetRecentFiles

  If frmCodeLib.MenuToolBar.Bands("subFile").Tools("file1").Caption = "  " Then
    frmCodeLib.MenuToolBar.Bands("subFile").Tools("file1").Caption = "Empty"
    frmCodeLib.MenuToolBar.Bands("subFile").Tools("file1").Enabled = False
  End If
  
  '-------------------
  'Setup ActiveBar
  '-------------------
  MainToolBar.Attach
  MainToolBar.AttachEx Frame1.hWnd
  MainToolBar.RecalcLayout
  MenuToolBar.Attach
  MenuToolBar.AttachEx Frame2.hWnd
  MenuToolBar.RecalcLayout
  ButtonBar.Attach
  ButtonBar.AttachEx ButtonFrame.hWnd
  ButtonBar.RecalcLayout
  '-------------------- INI FILE VALUES --------------------'
  '---------------
  'ForegroundColor
  '---------------
  Ini_File
  getnumPages = GetPrivateProfileString("Preferences", "BackgroundColor", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then 'It returned a value
    bgcolor = Left(Temp1, getnumPages)
    Drive1.BackColor = bgcolor
    Dir1.BackColor = bgcolor
    File1.BackColor = bgcolor
    txtFilename.BackColor = bgcolor
    frmFind.txtFind.BackColor = bgcolor
    frmFind.txtReplace.BackColor = bgcolor
    frmFolders.Drive1.BackColor = bgcolor
    frmFolders.Dir1.BackColor = bgcolor
    frmFolders.File1.BackColor = bgcolor
    frmPreferences.Label15.BackColor = bgcolor
    frmPreferences.listFilters.BackColor = bgcolor
    frmPreferences.listUsedFilters.BackColor = bgcolor
    frmPreferences.txtNewFilter.BackColor = bgcolor
    frmPreferences.txtProject.BackColor = bgcolor
    frmPreferences.txtRoot.BackColor = bgcolor
    frmPreferences.txtJunk.BackColor = bgcolor
    frmProperties.txtFilename.BackColor = bgcolor
  Else
    bgcolor = "-2147483643"
  End If
  
  '---------------
  'ForegroundColor
  '---------------
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "ForegroundColor", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    fgcolor = Left(Temp1, getnumPages)
    Drive1.ForeColor = fgcolor
    Dir1.ForeColor = fgcolor
    File1.ForeColor = fgcolor
    txtFilename.ForeColor = fgcolor
    frmFind.txtFind.ForeColor = fgcolor
    frmFind.txtReplace.ForeColor = fgcolor
    frmFolders.Drive1.ForeColor = fgcolor
    frmFolders.Dir1.ForeColor = fgcolor
    frmFolders.File1.ForeColor = fgcolor
    frmPreferences.listFilters.ForeColor = fgcolor
    frmPreferences.listUsedFilters.ForeColor = fgcolor
    frmPreferences.txtNewFilter.ForeColor = fgcolor
    frmPreferences.txtProject.ForeColor = fgcolor
    frmPreferences.txtRoot.ForeColor = fgcolor
    frmPreferences.txtJunk.ForeColor = fgcolor
    frmPreferences.Label16.BackColor = fgcolor
    frmProperties.txtFilename.ForeColor = fgcolor
  Else
    fgcolor = "0"
  End If
  
'--------------FONT NAME
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "FontName", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    fontname1 = Left(Temp1, getnumPages)
  Else
    fontname1 = "Courier New"
  End If
  
'--------------FONT SIZE
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "FontSize", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    Fsize = Left(Temp1, getnumPages)
  Else
    Fsize = "10"
  End If
  
'--------------FONT BOLD
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "FontBold", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    Fbold = Left(Temp1, getnumPages)
  Else
    Fbold = "False"
  End If
  
'--------------FONT ITALIC
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "FontItalic", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    Fitalic = Left(Temp1, getnumPages)
  Else
    Fitalic = "False"
  End If

'--------------TOOLBAR ON/OFF
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "ToolBarOn", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    ToolsOn = Left(Temp1, getnumPages)
    
    If ToolsOn = "True" Then
      TabStrip1.Top = Frame1.Top + Frame1.Height
      Frame1.Visible = True
    Else
      TabStrip1.Top = Frame2.Top + Frame2.Height
      Frame1.Visible = False
    End If
    
  Else
    ToolsOn = "True"
    Frame1.Visible = True
  End If
  
  Form_Resize
  
'--------------STATUSBAR ON/OFF
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "StatusBar", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    StatusOn = Left(Temp1, getnumPages)
    
    If StatusOn = "True" Then
      Parent.StatusBar1.Visible = True
    Else
      Parent.StatusBar1.Visible = False
    End If
    
  Else
    StatusOn = "True"
  End If
  
'--------------CONTROLS ON/OFF
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "FullView", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    FullView = Left(Temp1, getnumPages)
    
    If FullView = "True" Then
      Picture6.Tag = "NotVisible"
      Picture6.Visible = False
    Else
      Picture6.Tag = "Visible"
      Picture6.Visible = True
    End If
    
  Else
    FullView = "False"
    Picture6.Tag = "Visible"
    Picture6.Visible = True
  End If

'--------------TOOLTIPS ON/OFF
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "ToolTipsOn", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  
  If X = 0 Then
    toolTipsOn = Left(Temp1, getnumPages)
  End If
  
  MainToolBar.DisplayToolTips = toolTipsOn
  ButtonBar.DisplayToolTips = toolTipsOn
  
'--------------RIGHT-MARGIN
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "Right_Margin", file, Temp1, Len(Temp1), file)
  Dim RMargin As String
  RMargin = Left(Temp1, getnumPages)
  
  If Val(RMargin) <> 0 Then
    
    For X = 0 To 9
      rtb1(X).RightMargin = RMargin
    Next
  
  End If
  
  Parent.Height = 7200
  Parent.Width = 9550
  Frame1.Top = Frame2.Height
  Frame1.Width = Parent.Width
  Frame2.Width = Parent.Width
  
'--------------BACKGROUND COLOR
  For X = 0 To 9
    rtb1(X).BackColor = bgcolor
    rtb1(X).SelStart = 0
    rtb1(X).SelLength = Len(rtb1(X))
    DontCount = True
    rtb1(X).SelColor = fgcolor
    rtb1(X).SelStart = 0
    rtb1(X).SelLength = 0
    rtb1(X).Font = fontname1
    rtb1(X).Font.Size = Fsize
    rtb1(X).Font.Bold = Fbold
    rtb1(X).Font.Italic = Fitalic
  Next
  
'--------------ROOT FOLDER
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "RootDir", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  On Error GoTo errFolder
  
  If X = 0 Then
    RootDir = Left(Temp1, getnumPages)
    Dir1.Path = RootDir
    Drive1.Drive = RootDir
    File1.Path = RootDir
  End If
  
'--------------PROJECT FOLDER
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "ProjectFolder", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  On Error GoTo errFolder
  
  If X = 0 Then
    ProjectFolder = Left(Temp1, getnumPages)
  End If
  
'--------------FAVORITE FOLDER
  Temp1 = ""
  getnumPages = GetPrivateProfileString("Preferences", "FavFolder", file, Temp1, Len(Temp1), file)
  X = InStr(Temp1, ".ini")
  On Error GoTo errFolder
  
  If X = 0 Then
    FavFolder = Left(Temp1, getnumPages)
  End If
  
  Dir1_Change
  rtb1(0).ZOrder 0
  rtb1(0).Enabled = True
errFolder:
Exit Sub

ErrorLog_Trapper:
Call Log_Error("Form_Load-frmCodeLib", Err.Description)
End Sub
'
'
'
'------------------FORM RESIZE--------------------'

Private Sub Form_Resize()
  On Error Resume Next
  DoEvents
  TabStrip1.Width = Me.Width - 130
  If Parent.WindowState = vbMinimized Then Exit Sub
  
  If ToolsOn = "False" Then
    TabStrip1.Top = Frame2.Top + Frame2.Height
    TabStrip1.Height = Me.ScaleHeight - 395
    Picture6.Top = TabStrip1.Top + 50
    Picture6.Height = Me.Height - 895
  End If
  
  If ToolsOn = "True" Then
    Frame1.Visible = False
    TabStrip1.Top = Frame1.Top + Frame1.Height
    TabStrip1.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - 50
    Picture6.Top = TabStrip1.Top + 50
    Picture6.Height = Me.Height - 1300
    Frame1.Visible = True
  End If
  
  Picture6.Left = Me.ScaleWidth - Picture6.Width - 90
  Drive1.Top = 0.01 * Picture6.Height
  Dir1.Top = Drive1.Top + 375
  Dir1.Height = imgSplitter.Top - Dir1.Top
  If Dir1.Height < 540 Then Dir1.Height = 540
  imgSplitter.Top = Dir1.Top + Dir1.Height
  File1.Top = imgSplitter.Top + imgSplitter.Height
  ButtonFrame.Top = Picture6.Height - ButtonFrame.Height
  txtFilename.Top = ButtonFrame.Top - txtFilename.Height - 50
  File1.Height = txtFilename.Top - imgSplitter.Top - imgSplitter.Height
  
  If File1.Height <= 675 Then
    File1.Top = txtFilename.Top - 700
    imgSplitter.Top = File1.Top - imgSplitter.Height
    File1.Height = 675
    Dir1.Top = Drive1.Top + 375
    Dir1.Height = File1.Top - imgSplitter.Height - Dir1.Top
    imgSplitter.Top = Dir1.Top + Dir1.Height
    File1.Top = imgSplitter.Top + imgSplitter.Height
  End If
  
  '---------------------
  'Set up Richtextboxes
  '---------------------
  If Picture6.Tag = "Visible" Then
    
    For X = 0 To 9
      rtb1(X).Visible = False
      rtb1(X).Top = TabStrip1.Top + 30
      rtb1(X).Height = TabStrip1.Height - 325
      rtb1(X).Left = TabStrip1.Left + 25
      rtb1(X).Width = Me.Width - Picture6.Width - 285
      rtb1(X).Visible = True
    Next
    
  Else
    
    For X = 0 To 9
      rtb1(X).Visible = False
      rtb1(X).Top = TabStrip1.Top + 30
      rtb1(X).Height = TabStrip1.Height - 325
      rtb1(X).Left = TabStrip1.Left + 25
      rtb1(X).Width = Me.Width - 190
      rtb1(X).Visible = True
    Next
  
  End If
  
  Frame2.Width = Me.Width - 135
  Frame1.Width = Me.Width - 135
  Me.Refresh
End Sub
'
'
'
'------------------LINE COUNTER-------------------'

Function CountTheLines()
  On Error Resume Next
  Parent.StatusBar1.Panels(2).Text = "Counting the lines..."
  Y = TabStrip1.SelectedItem.Index - 1
  
  If rtb1(Y).Text = "" Then
    Parent.StatusBar1.Panels(2).Text = "No file currently open."
    Exit Function
  End If
  
  Dim FileSize As String
  Dim NumLines As String
  Dim Numpages As String
  
  '---------------------
  'File information
  '---------------------
  If TabStrip1.SelectedItem.Caption = "Untitled" Then
    FileSize = "Unknown"
  End If
  
  If Len(FileSize) >= 11 Then
    FileSize = Format(FileLen(rtb1(Y).FileName) / 1024, "#,###.##") & "KB"
  End If
  
  NumLines = Format$(linecount, "#,###") & " lines"
  Numpages = Format(linecount / 52, "#,###.00") & " pages"
  Parent.StatusBar1.Panels(2).Text = "File info: "
  
  If rtb1(Y).FileName <> "" And TabStrip1.SelectedItem.Caption <> "Untitled" And Parent.Caption <> " Code Library 3.0 - Untitled" Then
    FileSize = Format(FileLen(rtb1(Y).FileName), "#,###") & " bytes"
    Parent.StatusBar1.Panels(2).Text = Parent.StatusBar1.Panels(2).Text & "  " & rtb1(Y).FileName
    Parent.StatusBar1.Panels(2).Text = Parent.StatusBar1.Panels(2).Text & "   " & "Size: " & FileSize
    Parent.StatusBar1.Panels(2).Text = Parent.StatusBar1.Panels(2).Text & "   " & "Lines: " & Format$(linecount, "##,###")
  Else
    Parent.StatusBar1.Panels(2).Text = Parent.StatusBar1.Panels(2).Text & "  " & "Untitled" & "   Lines: " & Format$(linecount, "##,###")
  End If
  
  If linecount >= 52 Then
    Parent.StatusBar1.Panels(2).Text = Parent.StatusBar1.Panels(2).Text & "   " & "Pages: " & Format(linecount / 52, "#,###.00")
  Else
    Parent.StatusBar1.Panels(2).Text = Parent.StatusBar1.Panels(2).Text & "   " & "Pages:  1"
  End If
  
End Function

Function Get_Help()
  
  If Right(App.Path, 1) = "\" Then
    HelpPath = App.Path
  Else
    HelpPath = App.Path & "\"
  End If
  
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If TypeOf Screen.ActiveControl Is TextBox Then
    Exit Sub
  Else
    rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  End If
  
  If KeyCode = vbKeyF3 Then
    Hide_frmFind
    frmAbout.Show vbModal
    Exit Sub
  End If
  
  If KeyCode = vbKeyF4 Then
    Get_Properties
    Exit Sub
  End If
  
  If KeyCode = vbKeyF5 Then
    frmPreferences.TabStrip1.Tabs(1).Selected = True
    frmPreferences.Show vbModal
    Exit Sub
  End If
  
  If KeyCode <> vbKeyF1 Then
    Call MenuToolBar.OnKeyDown(KeyCode, Shift)
  End If
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  
  If TypeOf Screen.ActiveControl Is TextBox Then
    Exit Sub
  Else
    rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  End If
      
  If KeyCode = vbKeyF1 Then
    Get_Help
    retval = ShellExecute(hWnd, "Open", "CODELIB.HLP", "", HelpPath, 1)
    
    If retval = 2 Then
      Call MsgBox("The Help file was not found.", 16, " Error")
    End If
    
    Exit Sub
  End If

  Call MenuToolBar.OnKeyUp(KeyCode, Shift)
End Sub

Sub FilesSearch(DrivePath As String, Ext As String)
  Dim XDir() As String
  Dim TmpDir As String
  Dim FFound As String
  Dim DirCount As Integer
  Dim X As Integer
  'Initialises Variables
  DirCount = 0
  ReDim XDir(0) As String
  XDir(DirCount) = ""
    
  If Right(DrivePath, 1) <> "\" Then
    DrivePath = DrivePath & "\"
  End If

  DoEvents
  TmpDir = Dir(DrivePath, vbDirectory)

  Do While TmpDir <> ""
    DoEvents

    If TmpDir <> "." And TmpDir <> ".." Then

      If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then
        XDir(DirCount) = DrivePath & TmpDir & "\"
        DirCount = DirCount + 1
        ReDim Preserve XDir(DirCount) As String
      End If

    End If
            
    TmpDir = Dir
  Loop

  'Searches for the files given by extension Ext
  FFound = Dir(DrivePath & Ext)

  Do Until FFound = ""
    DoEvents
    Kill DrivePath & FFound
    FFound = Dir
  Loop

  'Recursive searches through all sub directories

  For X = 0 To (UBound(XDir) - 1)
    DoEvents
    FilesSearch XDir(X), Ext
    RmDir (XDir(X))
  Next X

End Sub

Function Get_Properties()
  Dim retval As Integer
  Get_Path
  If File1.FileName = "" Then Exit Function
  retval = GetAttr(filepath & File1.FileName)
  frmProperties.Check1.value = 0
  frmProperties.Check2.value = 0
  frmProperties.Check3.value = 0
  frmProperties.Check4.value = 0
  frmProperties.Check5.value = 0
  If retval = 0 Then frmProperties.Check1.value = 1  'Normal
  If retval = 1 Then frmProperties.Check2.value = 1  'Read-only
  If retval = 2 Then frmProperties.Check3.value = 1  'Hidden
  If retval = 4 Then frmProperties.Check4.value = 1  'System file
  If retval = 32 Then frmProperties.Check5.value = 1 'Archive
  
  If retval = 3 Then
    frmProperties.Check2.value = 1
    frmProperties.Check3.value = 1
  End If
  
  If retval = 5 Then
    frmProperties.Check2.value = 1
    frmProperties.Check4.value = 1
  End If
        
  If retval = 6 Then
    frmProperties.Check3.value = 1
    frmProperties.Check4.value = 1
  End If
        
  If retval = 7 Then
    frmProperties.Check2.value = 1
    frmProperties.Check3.value = 1
    frmProperties.Check4.value = 1
  End If
        
  If retval = 33 Then
    frmProperties.Check2.value = 1
    frmProperties.Check5.value = 1
  End If
  
  If retval = 35 Then
    frmProperties.Check2.value = 1
    frmProperties.Check3.value = 1
    frmProperties.Check5.value = 1
  End If
  
  If retval = 36 Then
    frmProperties.Check4.value = 1
    frmProperties.Check5.value = 1
  End If
  
  If retval = 37 Then
    frmProperties.Check2.value = 1
    frmProperties.Check4.value = 1
    frmProperties.Check5.value = 1
  End If
  
  If retval = 38 Then
    frmProperties.Check3.value = 1
    frmProperties.Check4.value = 1
    frmProperties.Check5.value = 1
  End If
  
  If retval = 39 Then
    frmProperties.Check2.value = 1
    frmProperties.Check3.value = 1
    frmProperties.Check4.value = 1
    frmProperties.Check5.value = 1
  End If
  
  If FileLen(filepath & File1.FileName) >= 999 Then
    frmProperties.txtFilename.Text = filepath & File1.FileName
    frmProperties.Label2.Caption = Format(FileLen(filepath & File1.FileName) / 1024, "#,###.##") & "KB"
    frmProperties.Label3.Caption = Format(FileDateTime(filepath & File1.FileName), "dddd, mmm d yyyy, hh:mm AMPM")
  Else
    frmProperties.txtFilename.Text = filepath & File1.FileName
    frmProperties.Label2.Caption = FileLen(filepath & File1.FileName) & " bytes."
    frmProperties.Label3.Caption = Format(FileDateTime(filepath & File1.FileName), "dddd, mmm d yyyy, hh:mm AMPM")
  End If
  
  Hide_frmFind
  frmProperties.Show vbModal
End Function
'
'
'
'------------------PATH MODULES-------------------'

Public Sub Ini_File()
  filepath = App.Path
  
  If Right(filepath, 1) <> "\" Then
    filepath = App.Path & "\"
  End If
  
  file = filepath & "Library.ini"
End Sub

Public Sub Get_Path()
  
  If Right(Dir1.Path, 1) <> "\" Then
    filepath = Dir1.Path & "\"
  Else
    filepath = Dir1.Path
  End If
    
End Sub
'
'
'
'-----------------FOLDER BUTTONS------------------'

Private Sub cmdRoot_Click()
  If RootDir = "" Then Exit Sub
  Drive1.Drive = Left(RootDir, 1)
  Dir1.Path = RootDir
End Sub

Private Sub cmdProjects_Click()
  If ProjectFolder = "" Then Exit Sub
  Drive1.Drive = Left(ProjectFolder, 1)
  Dir1.Path = ProjectFolder
End Sub

Private Sub cmdFolder_Click()
  If FavFolder = "" Then Exit Sub
  Drive1.Drive = Left(FavFolder, 1)
  Dir1.Path = FavFolder
End Sub
'
'
'
'--------------CREATE NEW DIRECTORY---------------'

Private Sub mnuNewFolder_Click()
  Get_Path
  answer = InputBox("Enter a name for the new folder.", "Create a new folder")
    
  If answer <> "" Then
    MkDir (filepath & answer)
  End If
    
  Dir1.Refresh
End Sub
'
'
'
'-----------------REMOVE DIRECTORY----------------'

Private Sub mnuDelFolder_Click()
  On Error GoTo oh_no
  '------------------------
  'Don't want to delete
  'these folders!
  '------------------------
  If UCase(Dir1.Path) <> "C:\" And UCase(Dir1.Path) <> "C:\WINDOWS" Then
        
    If MsgBox("Are you sure you want to remove " & vbCr & Dir1.Path & " ?" & vbCr & "All files and subfolders will be deleted!", 36 + vbDefaultButton2, " Remove directory") = 6 Then
      filepath = Dir1.Path & "\"
      DoEvents
      FilesSearch filepath, "*.*"
      Dir1.Path = Dir1.List(Dir1.ListIndex - 1)
      RmDir (filepath)
      Dir1.Refresh
      File1.Refresh
    End If
        
  End If
  
  Exit Sub
  
oh_no:
Exit Sub
End Sub
'
'
'
'---------------OPEN & VIEW FILE------------------'

Private Sub File1_Click()
  On Error GoTo ErrorLog_Trapper
  If DontMove = True Then Exit Sub
  If DeSelect = True Then Exit Sub
  Dim WhichOne As Integer
  WhichOne = TabStrip1.SelectedItem.Index - 1
  DoEvents
  Get_Path
  MenuToolBar.Bands("subFile").Tools("file1").Enabled = True
  rtb1(WhichOne).ZOrder 0
  rtb1(WhichOne).Visible = True
  DoEvents
  rtb1(WhichOne).LoadFile filepath & File1.FileName, 1
  DoEvents
  lengthOfName = InStrRev(rtb1(WhichOne).FileName, "\", Len(rtb1(WhichOne).FileName))
  txtFilename.Text = Right(rtb1(WhichOne).FileName, Len(rtb1(WhichOne).FileName) - lengthOfName)
  TabStrip1.Tabs(WhichOne + 1).Caption = txtFilename.Text
    
    For NumChecked = 0 To 9
      MenuToolBar.Bands("subWindow").Tools("window" & NumChecked).Checked = False
    Next
    
  MenuToolBar.Bands("subWindow").Tools("window" & Y).Checked = True
  TabStrip1.SelectedItem.Image = 1
  DoEvents
  CountTheLines
  TabStrip1.SelectedItem.Caption = txtFilename.Text
  UpdateFileMenu filepath & File1.FileName
  Parent.Caption = " Code Library 3.0 - " & filepath & File1.FileName
  Picture6.ZOrder 0
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("File1_Click-frmCodeLib", Err.Description)
End Sub
'
'
'
'------------------DRIVE CHANGE-------------------'

Private Sub Drive1_Change()
  On Error GoTo oh_no
  Dir1.Path = Drive1.Drive
  
  If Picture6.Visible = True Then
    Dir1.SetFocus
  Else
    rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  End If
  
  Exit Sub
    
oh_no:
Call MsgBox("That drive is not available.", 48, " Error")
Exit Sub

End Sub
'
'
'
'----------------DIRECTORY CHANGE-----------------'

Private Sub Dir1_Change()
  On Error Resume Next
  File1.Path = Dir1.Path
  Drive1.Drive = Dir1.Path
  
  If Dir1.Path = RootDir Then
    MainToolBar.Bands("Toolbar").Tools("root").Checked = True
    MainToolBar.Bands("Toolbar").Tools("main").Checked = False
    MainToolBar.Bands("Toolbar").Tools("temp").Checked = False
  Else
    MainToolBar.Bands("Toolbar").Tools("root").Checked = False
  End If
  
  If Dir1.Path = ProjectFolder Then
    MainToolBar.Bands("Toolbar").Tools("root").Checked = False
    MainToolBar.Bands("Toolbar").Tools("main").Checked = True
    MainToolBar.Bands("Toolbar").Tools("temp").Checked = False
  Else
    MainToolBar.Bands("Toolbar").Tools("main").Checked = False
  End If
  
  If Dir1.Path = FavFolder Then
    MainToolBar.Bands("Toolbar").Tools("root").Checked = False
    MainToolBar.Bands("Toolbar").Tools("main").Checked = False
    MainToolBar.Bands("Toolbar").Tools("temp").Checked = True
  Else
    MainToolBar.Bands("Toolbar").Tools("temp").Checked = False
  End If
  
End Sub
'
'
'
'-------------SAVE SNIPPIT TO FILE----------------'

Private Sub mnuSave_Click()
  cmdSave_Click
End Sub

Private Sub cmdSave_Click()
  On Error GoTo ErrorLog_Trapper
  
  If rtb1(TabStrip1.SelectedItem.Index - 1).Text = "" Then
    Dir1.SetFocus
    Exit Sub
  End If
  
  '--------------------------
  'No filename selected
  '--------------------------
  If txtFilename.Text = "" Then
    Call MsgBox("No filename selected.", 48, " Error")
    txtFilename.SetFocus
    Exit Sub
  End If
    
  '--------------------------
  'Save text to file
  '--------------------------
  Get_Path
  TextToSave = rtb1(TabStrip1.SelectedItem.Index - 1).Text
  
  If MsgBox("Save this as " & txtFilename & "?", 36, " Save file") = vbYes Then
    Dim retval As Integer
    Dim returnFile As String
    returnFile = Dir(filepath & txtFilename.Text)
    'If the file already exists
    If returnFile <> "" Then
      retval = GetAttr(filepath & txtFilename.Text)
      
      If retval <> 32 And retval <> 0 Then
        
        If MsgBox("That file is read-only, hidden or a system file." & vbCr & "Do you still want to modify the file?", 36, "Secured file") = vbYes Then
          SetAttr filepath & txtFilename.Text, vbNormal
        Else
          Exit Sub
        End If
       
      End If
      
      rtb1(TabStrip1.SelectedItem.Index - 1).SaveFile filepath & txtFilename.Text, 1
    End If
    
    rtb1(TabStrip1.SelectedItem.Index - 1).SaveFile filepath & txtFilename.Text, 1
  End If
    
  File1.Refresh
    
  For X = 0 To File1.ListCount - 1
        
    If txtFilename.Text = File1.List(X) Then
      DontMove = True
      File1.ListIndex = X
      DontMove = False
      Exit For
    End If
            
  Next
    
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("cmdSave_Click-frmCodeLib", Err.Description)
End Sub
'
'
'
'-------------------CLEAR ALL---------------------'

Private Sub cmdClear_Click()
  On Error GoTo ErrorLog_Trapper
  If MsgBox("Close all open files?", 36, " Are you sure?") = vbNo Then Exit Sub
  PopUpToolBar.Bands("Toolbar").Tools("recent").Text = ""
  
  For X = 0 To 9
    rtb1(X).Text = ""
    rtb1(X).SelText = ""
    rtb1(X).FileName = ""
    TabStrip1.Tabs(X + 1).Caption = ""
    TabStrip1.Tabs(X + 1).Image = 2
  Next
  
  rtb1(0).Enabled = True
  rtb1(0).ZOrder 0
  rtb1(0).Visible = True
  txtFilename.Text = ""
  ClearAll = True
  TabStrip1.Tabs(1).Caption = "Code Library"
  TabStrip1.Tabs(1).Selected = True
  '------------------------------------------------------------
  'Deselect is so the file gets highlighted in the filelistbox
  'but the file1_click procedure is skipped to avoid problems
  '------------------------------------------------------------
  For X = 0 To File1.ListCount - 1
    DeSelect = True
    File1.Selected(X) = False
  Next
  
  DeSelect = False
  ClearAll = False
  rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  Parent.Caption = " Code Library 3.0"
  Exit Sub

ErrorLog_Trapper:
ClearAll = False
Call Log_Error("cmdClear_Click-frmCodeLib", Err.Description)
End Sub

Private Sub mnuClearAll_Click()
  cmdClear_Click
  Parent.StatusBar1.Panels(2).Text = "Close all open files."
End Sub
'
'
'
'------------------DELETE FILE--------------------'

Private Sub mnuDel_Click()
  On Error GoTo ErrorLog_Trapper
  Dim retString As String
  answer = MsgBox("Are you sure you want to delete " & File1.FileName & "?", 36 + vbDefaultButton2, "Delete file?")
    
  If answer = vbYes Then
    Get_Path
    'What a great VB word, Kill
    retString = Dir(filepath & File1.FileName)
    
    If retString = "" Then
      Call MsgBox("File not found.", 16, " Error")
      Exit Sub
    End If
    
    Kill (filepath & File1.FileName)
    File1.Refresh
    TabStrip1_Click
    txtFilename.Text = ""
  End If
  
  Dir1.SetFocus
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("mnuDel_Click-frmCodeLib", Err.Description)
End Sub
'
'
'
'-----------------EXIT PROGRAM--------------------'

Private Sub mnuExitParent_Click()
  mnuExit_Click
End Sub

Private Sub mnuExit_Click()
  Dim frm As Form
  
  For Each frm In Forms
    Unload frm
    Set frm = Nothing
  Next frm
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ErrorLog_Trapper
  Ini_File
  Dim WinState As String
  WinState = Parent.WindowState
  retval = WritePrivateProfileString("Preferences", "WindowState", WinState, file)
  
  Dim frm As Form
  
  For Each frm In Forms
    Unload frm
    Set frm = Nothing
  Next frm
  
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("Form_QueryUnload-frmCodeLib", Err.Description)
End Sub
'
'
'
'---------OPEN FILE WITH DEFAULT PROGRAM----------'

Private Sub mnuopen_Click()
  On Error Resume Next
  
  If File1.FileName = "" Then
    answer = MsgBox("There is no file selected.  Open Windows Explorer?", 36, "Search...")
  Else
    answer = MsgBox("Open " & File1.FileName & " with default program?", 36, "Open file...")
  End If
  
  If answer = 6 Then
    Get_Path
    Call ShellExecute(hWnd, "Open", File1.FileName, "", filepath, 1)
  End If
  
End Sub
'
'
'
'-----------------RENAME FOLDER-------------------'

Private Sub mnuRename_Click()
  On Error GoTo ErrorLog_Trapper
  If UCase(Dir1.Path) = "C:\" Or Dir1.Path = "C:\WINDOWS" Then Exit Sub
  answer = InputBox("Enter the new name for " & Dir1.Path & ".", "Rename folder")
  Get_Path
  
  If answer <> "" Then
    Name filepath As Dir1.List(Dir1.ListIndex - 1) & "\" & answer
    Me.Refresh
    Dir1.Path = Dir1.List(Dir1.ListIndex - 1)
  End If
  
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("mnuRename_Click-frmCodeLib", Err.Description)
End Sub
'
'
'
'--------------POP-UP MENU ON FORM----------------'

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 2 Then
    PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
  End If
  
End Sub

Private Sub TabStrip1_GotFocus()
  rtb1(TabStrip1.SelectedItem.Index - 1).Enabled = True
  rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
End Sub

Private Sub TabStrip1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
  End If
  
End Sub

Private Sub txtFilename_KeyUp(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 13 Then
    cmdSave_Click
    rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  End If
  
End Sub

Private Sub txtFileName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Type in a name for the file and click the Save button below."
End Sub
'
'
'
'------------------POP-UP MENU--------------------'

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorLog_Trapper
  
  If Button = 2 Then
    PopUpToolBar.Bands("FileOptions").TrackPopup -1, -1
  End If
    
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("File1_MouseDown-frmCodeLib", Err.Description)
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  If Button = 2 Then
    PopUpToolBar.Bands("Folder Options").TrackPopup -1, -1
  End If
    
End Sub
'
'
'
'-----------------STATUS BAR TEXT-----------------'

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Click on a file to view the contents. - Right-click for file options.  Number of files: " & File1.ListCount
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 2 Then
    PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
  End If
  
End Sub

Private Sub rtb1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 2 Then
    PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
  End If
  
End Sub

Private Sub rtb1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  If rtb1(Index).Text <> "" And TabStrip1.SelectedItem.Caption <> "Code Library" And TabStrip1.SelectedItem.Caption <> "" Then
    
    If Left(Parent.StatusBar1.Panels(2).Text, 4) <> "File" Then
      CountTheLines
    End If
    
  Else
  
    If rtb1(Index).Text <> "" And TabStrip1.SelectedItem.Caption <> "Code Library" Then
      Parent.StatusBar1.Panels(2).Text = "To save this text, type in a filename below and click Save."
    Else
      Parent.StatusBar1.Panels(2).Text = "No file currently open."
    End If
    
  End If
  
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Parent.StatusBar1.Panels(2).Text = "Double-click on a directory to open it. - Right-click for folder options."
End Sub

Private Sub Hide_frmFind()
  
  If frmFind.checkTop.value = 1 Then
    frmFind.checkTop.value = 0
    MakeWindowNotTop frmFind.hWnd
    frmFind.Hide
  End If
  
End Sub
'
'
'
'-------------------FONT OPTIONS------------------'

Private Sub mnuFontOptions_Click()
  On Error GoTo oh_no
  Parent.StatusBar1.Panels(2).Text = "Specify font type, size and other font options."
  Dim numUsed As Integer
  numUsed = TabStrip1.SelectedItem.Index - 1
  Hide_frmFind
  '--------------------
  'Change selected text
  '--------------------
  If rtb1(numUsed).SelText <> "" Then
    CommonDialog1.Flags = &H3
    CommonDialog1.fontname = rtb1(numUsed).SelFontName
    CommonDialog1.FontSize = rtb1(numUsed).SelFontSize
    CommonDialog1.FontBold = rtb1(numUsed).SelBold
    CommonDialog1.FontItalic = rtb1(numUsed).SelItalic
    CommonDialog1.ShowFont
    rtb1(numUsed).SelFontName = CommonDialog1.fontname
    rtb1(numUsed).SelBold = CommonDialog1.FontBold
    rtb1(numUsed).SelItalic = CommonDialog1.FontItalic
    rtb1(numUsed).SelFontSize = CommonDialog1.FontSize
  Else
  '--------------------
  'Change all the text
  '--------------------
    CurrentPos = rtb1(TabStrip1.SelectedItem.Index - 1).SelStart
    CommonDialog1.Flags = &H3
    CommonDialog1.fontname = rtb1(numUsed).Font.Name
    CommonDialog1.FontSize = rtb1(numUsed).Font.Size
    CommonDialog1.FontBold = rtb1(numUsed).Font.Bold
    CommonDialog1.FontItalic = rtb1(numUsed).Font.Italic
    CommonDialog1.ShowFont
    
    For numUsed = 0 To 9
      rtb1(numUsed).Font.Name = CommonDialog1.fontname
      rtb1(numUsed).Font.Bold = CommonDialog1.FontBold
      rtb1(numUsed).Font.Italic = CommonDialog1.FontItalic
      rtb1(numUsed).Font.Size = CommonDialog1.FontSize
    Next
    
    Ini_File
    Temp1 = ""
    getnumPages = GetPrivateProfileString("Preferences", "ForegroundColor", file, Temp1, Len(Temp1), file)
    X = InStr(Temp1, ".ini")
    
    If X = 0 Then
      fgcolor = Left(Temp1, getnumPages)
    End If
    
    For numUsed = 0 To 9
      DontCount = True
      rtb1(numUsed).SelStart = 0
      rtb1(numUsed).SelLength = Len(rtb1(numUsed))
      rtb1(numUsed).SelColor = fgcolor
      rtb1(numUsed).SelStart = 0
      rtb1(numUsed).SelLength = 0
    Next
    
    DontCount = False
    Ini_File
    fontname1 = rtb1(0).Font.Name
    Fsize = rtb1(0).Font.Size
    Fbold = rtb1(0).Font.Bold
    Fitalic = rtb1(0).Font.Italic
    retval = WritePrivateProfileString("Preferences", "FontName", fontname1, file)
    retval = WritePrivateProfileString("Preferences", "FontSize", Fsize, file)
    retval = WritePrivateProfileString("Preferences", "FontBold", Fbold, file)
    retval = WritePrivateProfileString("Preferences", "FontItalic", Fitalic, file)
  End If
    
  rtb1(TabStrip1.SelectedItem.Index - 1).SelStart = CurrentPos
  rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  Exit Sub
  
oh_no:
Exit Sub
End Sub
'
'
'
'------------------PRINT OPTIONS------------------'

Private Sub mnuPrint_Click()
  On Error GoTo oh_no
  If rtb1(TabStrip1.SelectedItem.Index - 1).Text = "" Then Exit Sub
  CommonDialog1.ShowPrinter
  Printer.Font = fontname1
  Printer.FontBold = Fbold
  Printer.FontItalic = Fitalic
  Printer.FontSize = Fsize
  DoEvents
  Printer.Print rtb1(TabStrip1.SelectedItem.Index - 1).Text
  Printer.EndDoc
  Exit Sub

oh_no:
  
  If Err.Number = 32755 Then
    Exit Sub
  Else
    Call MsgBox("There was a printer error.", 48, "Error")
    Exit Sub
  End If
  
End Sub

Private Sub mnuPrintAll_Click()
  mnuPrint_Click
End Sub

Private Sub mnuPrintSelected_Click()
  On Error GoTo oh_no
  If rtb1(TabStrip1.SelectedItem.Index - 1).SelText = "" Then Exit Sub
  answer = MsgBox("Are you sure you want to print the selected text?", 36, "Print selected text...")
  
  If answer = 6 Then
    Printer.Font = fontname1
    Printer.FontBold = Fbold
    Printer.FontItalic = Fitalic
    Printer.FontSize = Fsize
    DoEvents
    Printer.Print rtb1(TabStrip1.SelectedItem.Index - 1).SelText
    Printer.EndDoc
  End If
  
  CommonDialog1.CancelError = False
  Exit Sub
  
oh_no:
  
  If Err.Number = 32755 Then
    Exit Sub
  Else
    Call MsgBox("There was a printer error.", 48, "Error")
    Exit Sub
  End If
  
End Sub

Private Sub mnuPrintsetup_Click()
  On Error GoTo oh_no
  CommonDialog1.CancelError = False
  CommonDialog1.ShowPrinter
  Exit Sub
  
oh_no:
Exit Sub
End Sub
'
'
'
'-----------------EDIT MENU ITEMS-----------------'

Private Sub mnuCopy_Click()
  Parent.StatusBar1.Panels(2).Text = "Copy the selected text to the clipboard."
  Clipboard.Clear
  Clipboard.SetText rtb1(TabStrip1.SelectedItem.Index - 1).SelText
End Sub

Private Sub mnuCut_Click()
  Parent.StatusBar1.Panels(2).Text = "Cut the selected text."
  Clipboard.Clear
  Clipboard.SetText rtb1(TabStrip1.SelectedItem.Index - 1).SelText
  rtb1(TabStrip1.SelectedItem.Index - 1).SelText = ""
End Sub

Private Sub mnuPaste_Click()
  Parent.StatusBar1.Panels(2).Text = "Paste from the clipboard."
  
  If Clipboard.GetText <> "" Then
    rtb1(TabStrip1.SelectedItem.Index - 1).SelText = Clipboard.GetText
    
    If TabStrip1.SelectedItem.Caption = "Code Library" Or TabStrip1.SelectedItem.Caption = "" Then
      TabStrip1.SelectedItem.Caption = "Untitled"
      TabStrip1.SelectedItem.Image = 1
      File1.Refresh
    End If
    
  Else
    Parent.StatusBar1.Panels(2).Text = "The clipboard is empty."
  End If
  
  TabStrip1_Click
End Sub

Private Sub mnuSelectAll_Click()
  Parent.StatusBar1.Panels(2).Text = "Select all text in the current file."
  Y = TabStrip1.SelectedItem.Index - 1
  rtb1(Y).SelStart = 0
  rtb1(Y).SelLength = Len(rtb1(Y))
End Sub
'
'
'
'-------------SHOW NEXT AVAILABLE TAB-------------'

Private Sub cmdNew_Click()
  mnuFileNew_Click
End Sub

Private Sub mnuFileNew_Click()
  On Error GoTo ErrorLog_Trapper
  Dim NewBox As Integer
  Dim AllFull As Boolean
  AllFull = False
  
  For X = 0 To 9
    
    If rtb1(X).Text = "" Then
      TabStrip1.Tabs(X + 1).Selected = True
      rtb1(X).ZOrder 0
      Exit Sub
    Else
      AllFull = True
    End If
    
  Next
  
  If AllFull = True Then Exit Sub
  NewBox = TabStrip1.SelectedItem.Index - 1
  If rtb1(NewBox).Text = "" Then Exit Sub
  
  For X = 0 To 9
    
    If rtb1(X).Text = "" Then
      TabStrip1.Tabs(X + 1).Selected = True
      Exit Sub
    End If
    
  Next
  
  DoEvents
  With rtb1(Numtabs - 1)
    .Text = ""
    .SelStart = 0
    .SelLength = Len(rtb1(NewBox))
    .SelColor = rtb1(0).SelColor
    .Font = rtb1(0).Font
    .Font.Size = rtb1(0).Font.Size
    .SelLength = 0
    .SelStart = 0
    .Visible = True
    .ZOrder 0
  End With
  
  TabStrip1.Tabs(Numtabs).Selected = True
  Form_Resize
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("mnuFileNew_Click-frmCodeLib", Err.Description)
Exit Sub
End Sub

Private Sub mnuClose_Click()
  mnuFileClose_Click
End Sub

Private Sub mnuFileClose_Click()
  rtb1(TabStrip1.SelectedItem.Index - 1).Text = ""
  TabStrip1.SelectedItem.Caption = ""
  
  If TabStrip1.SelectedItem.Index = 1 Then
    TabStrip1.Tabs(1).Caption = "Code Library"
  End If
  
  TabStrip1.SelectedItem.Image = 2
  TabStrip1_Click
  Form_Resize
End Sub
'
'
'
'-------------------RENAME FILE-------------------'

Private Sub mnuRenameFile_Click()
  On Error GoTo ErrorLog_Trapper
  Get_Path
  filepath = filepath & File1.FileName
  answer = InputBox("Enter the new name for " & vbCr & filepath & ".", "Rename file")
  If answer = "" Then Exit Sub
  
  If InStr(answer, ".") = 0 Then
    Call MsgBox("No extension given.  Can't rename file.", 16)
    Exit Sub
  End If
  
  If answer <> "" Then
    Dim NewAnswer As String
    NewAnswer = Left(filepath, Len(filepath) - Len(File1.FileName)) & answer
    Name filepath As NewAnswer
    File1.Refresh
  End If
  
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("mnuRename_Click-frmCodeLib", Err.Description)
End Sub

Private Sub mnuFind_Click()
  Parent.StatusBar1.Panels(2).Text = "Use the Find/Find Next/Replace utility."
  frmFind.Show
End Sub

Private Sub mnuSetPaths_Click()
  Hide_frmFind
  frmPreferences.Show vbModal
End Sub
'
'
'
'----------------STATUSBAR ON/OFF-----------------'

Private Sub mnuStatusOn_Click()
  
  If StatusOn = "True" Then
    StatusOn = "False"
    Parent.StatusBar1.Visible = False
  Else
    StatusOn = "True"
    Parent.StatusBar1.Visible = True
  End If
  
  rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  Ini_File
  retval = WritePrivateProfileString("Preferences", "StatusBar", StatusOn, file)
End Sub
'
'
'
'-----------------FULLVIEW ON/OFF-----------------'

Private Sub mnuHideDir_Click()
  On Error Resume Next
  rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  
  If Picture6.Tag = "Visible" Then
    Picture6.Visible = False
    Picture6.Tag = "NotVisible"
    FullView = "True"
    Parent.StatusBar1.Panels(2).Text = "Controls are currently off."
  Else
    Picture6.Tag = "Visible"
    Picture6.Visible = True
    FullView = "False"
    Parent.StatusBar1.Panels(2).Text = "Controls are currently on."
  End If
  
  Form_Resize
  Ini_File
  retval = WritePrivateProfileString("Preferences", "FullView", FullView, file)
End Sub
'
'
'
'-------------------TIPS ON/OFF-------------------'

Private Sub mnuTipsOn_Click()
  
  If toolTipsOn = "True" Then
    toolTipsOn = "False"
    MainToolBar.DisplayToolTips = False
    ButtonBar.DisplayToolTips = False
    Dir1.ToolTipText = ""
    File1.ToolTipText = ""
  Else
    toolTipsOn = "True"
    MainToolBar.DisplayToolTips = True
    ButtonBar.DisplayToolTips = True
    Dir1.ToolTipText = "Right-click for folder options"
    File1.ToolTipText = "Right-click for file options"
  End If
   
  If toolTipsOn = "True" Then
    Parent.StatusBar1.Panels(2).Text = "Tooltips are currently on."
  Else
    Parent.StatusBar1.Panels(2).Text = "Tooltips are currently off."
  End If
    
  Ini_File
  retval = WritePrivateProfileString("Preferences", "ToolTipsOn", toolTipsOn, file)
End Sub
'
'
'
'-----------------TOOLBAR ON/OFF------------------'

Private Sub mnuToolsOn_Click()
  rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  Me.Refresh
  DoEvents
  
  If ToolsOn = "True" Then
    ToolsOn = "False"
    Frame1.Visible = False
  Else
    ToolsOn = "True"
    Frame1.Visible = True
  End If

  Dim VisibleNow As Boolean
  
  If Picture6.Visible = True Then
    Picture6.Visible = False
    VisibleNow = True
  End If
  
  Form_Resize
  
  If Picture6.Visible = False And VisibleNow = True Then
    Picture6.Visible = True
  End If
  
  Ini_File
  retval = WritePrivateProfileString("Preferences", "ToolBarOn", ToolsOn, file)
End Sub
'
'
'
'------------RESIZE DIR/FILE LISTBOXES------------'

Private Sub imgSplitter_DblClick()
  picSplitter.Visible = False
  picSplitter.Top = 0.45 * Picture6.Height
  imgSplitter.Top = picSplitter.Top
  mbMoving = False
  Form_Resize
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  picSplitter.Visible = True
  picSplitter.Top = imgSplitter.Top
  mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim sglPos As Single
  
  If mbMoving Then
    sglPos = Y + imgSplitter.Top
    picSplitter.Top = sglPos
    imgSplitter.Top = picSplitter.Top
  End If
  
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If File1.Height < 480 Then
    File1.Top = txtFilename.Top - 500
    imgSplitter.Top = File1.Top - imgSplitter.Height
    File1.Height = 480
    Dir1.Top = Drive1.Top + 375
    Dir1.Height = File1.Top - imgSplitter.Height - Dir1.Top
  End If
  
  If Dir1.Height < 540 Then
    Dir1.Height = 540
    picSplitter.Top = Dir1.Top + Dir1.Height
  End If
  
  If picSplitter.Top > 0.8 * Picture6.Height Then
    picSplitter.Top = 0.8 * Picture6.Height
  End If
  
  If picSplitter.Top < 0.15 * Picture6.Height Then
    picSplitter.Top = 0.15 * Picture6.Height
  End If
  
  picSplitter.Visible = False
  imgSplitter.Top = picSplitter.Top
  mbMoving = False
  rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  Form_Resize
End Sub
'
'
'
'---------------RICHTEXTBOX CHANGE----------------'

Private Sub rtb1_Change(Index As Integer)
  On Error GoTo ErrorLog_Trapper
  
  If DontCount = True Then
    DontCount = False
    Exit Sub
  End If
  
  linecount = SendMessageLong(rtb1(Index).hWnd, EM_GETLINECOUNT, 0&, 0&)
  Exit Sub
  
ErrorLog_Trapper:
Call Log_Error("rtb1_Change-frmCodeLib", Err.Description)
End Sub
'
'
'
'-----------------TABSTRIP CLICK------------------'

Private Sub TabStrip1_Click()
  On Error GoTo ErrorLog_Trapper
  If ClearAll = True Then Exit Sub
  Y = TabStrip1.SelectedItem.Index - 1
  
  For i = 0 To TabStrip1.Tabs.Count - 1
    
    If i = Y Then
      rtb1(i).Enabled = True
      rtb1(i).SetFocus
      rtb1(Y).Visible = True
      rtb1(Y).ZOrder 0
    Else
      rtb1(i).Enabled = False
    End If
    
  Next
    
  For NumChecked = 0 To 9
    MenuToolBar.Bands("subWindow").Tools("window" & NumChecked).Checked = False
  Next
  
  MenuToolBar.Bands("subWindow").Tools("window" & Y).Checked = True
  '-----------------------------------
  'If the Richtextbox has a file open
  '-----------------------------------
  If rtb1(Y).FileName <> "" And rtb1(Y).Text <> "" And TabStrip1.SelectedItem.Caption <> "Untitled" Then
    lengthOfName = InStrRev(rtb1(Y).FileName, "\", Len(rtb1(Y).FileName))
    txtFilename.Text = Right(rtb1(Y).FileName, Len(rtb1(Y).FileName) - lengthOfName)
    TabStrip1.Tabs(Y + 1).Caption = txtFilename.Text
    Y = TabStrip1.SelectedItem.Index - 1
    On Error Resume Next
    Drive1.Drive = Left(rtb1(Y).FileName, 2)
    Dir1.Path = Left(rtb1(Y).FileName, lengthOfName)
    
    For X = 0 To File1.ListCount - 1
      
      If txtFilename.Text = File1.List(X) Then
        DontMove = True
        File1.ListIndex = X
        DontMove = False
        TabStrip1.SelectedItem.Image = 1
        Exit For
      End If
          
    Next
  
  End If
       
  linecount = SendMessageLong(rtb1(Y).hWnd, EM_GETLINECOUNT, 0&, 0&)
  
  If TabStrip1.Tabs(Y + 1).Caption = "" Or TabStrip1.Tabs(Y + 1).Caption = "Code Library" Then
    Parent.Caption = " Code Library 3.0"
    
    For NumWindows = 0 To 9
      MenuToolBar.Bands("subWindow").Tools("window" & NumWindows).Checked = False
    Next

    File1.Refresh
    txtFilename.Text = ""
    CountTheLines
    Exit Sub
  End If
      
  If File1.FileName <> "" And TabStrip1.SelectedItem.Caption <> "Untitled" Then
    Get_Path
    Parent.Caption = " Code Library 3.0 - " & filepath & File1.FileName
  Else
    File1.Refresh
    txtFilename.Text = ""
    Parent.Caption = " Code Library 3.0 - Untitled"
  End If
  
  CountTheLines
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("TabStrip1_Click-frmCodeLib", Err.Description)
End Sub
'
'
'
'---------------ACTIVEBAR SECTION-----------------'From here to the bottom is Activebar stuff...

Private Sub PopUpToolBar_BandOpen(ByVal Band As ActiveBarLibraryCtl.Band)
  Tools_Enabled
End Sub

'-------------
'MAIN TOOLBAR
'-------------
Private Sub MainToolBar_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  On Error GoTo ErrorLog_Trapper
  
  Select Case Tool.Name
    Case Is = "options"
      Hide_frmFind
      frmPreferences.Show vbModal
    Case Is = "undo"
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
      SendKeys "^z", True
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
    Case Is = "Help"
      Get_Help
      Call ShellExecute(hWnd, "Open", "CODELIB.HLP", "", HelpPath, 1)
    Case Is = "root"
      cmdRoot_Click
    Case Is = "main"
      cmdProjects_Click
    Case Is = "temp"
      cmdFolder_Click
    Case Is = "save"
      If rtb1(TabStrip1.SelectedItem.Index - 1).Text = "" Then Exit Sub
      If rtb1(TabStrip1.SelectedItem.Index - 1).FileName = "" Then Exit Sub
      If File1.FileName = "" Then Exit Sub
      cmdSave_Click
    Case Is = "Find"
      mnuFind_Click
    Case Is = "selectAll"
        mnuSelectAll_Click
    Case Is = "copy"
      mnuCopy_Click
    Case Is = "paste"
      mnuPaste_Click
    Case Is = "clear"
      rtb1(TabStrip1.SelectedItem.Index - 1).Text = ""
      TabStrip1.SelectedItem.Caption = ""
      TabStrip1.SelectedItem.Image = 2
      txtFilename.Text = ""
      Parent.Caption = " Code Library 3.0"
      File1.Refresh
      
      If TabStrip1.SelectedItem.Index = 1 Then
        TabStrip1.Tabs(1).Caption = "Code Library"
        TabStrip1.Tabs(1).Image = 2
      End If
      
    Case Is = "exit"
      mnuExit_Click
    Case Is = "print"
      mnuPrint_Click
    Case Is = "cut"
      mnuCut_Click
  End Select
  
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("MainToolBar_Click-frmCodeLib", Err.Description)
End Sub

'-------------
'MENU TOOLBAR
'-------------
Private Sub MenuToolBar_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  On Error GoTo ErrorLog_Trapper
  
  Select Case Tool.Name
    Case Is = "clearclip"
      
      If MsgBox("Clear the Clipboard?  Are you sure?", 36, " Clear Clipboard") = 6 Then
        Clipboard.Clear
      End If
      
    Case Is = "saveas"
      On Error Resume Next
      CommonDialog1.FileName = txtFilename.Text
      CommonDialog1.ShowSave
      
      If Err.Number <> 32755 Then
        rtb1(TabStrip1.SelectedItem.Index - 1).SaveFile CommonDialog1.FileName, 1
        File1.Refresh
      End If
      
    Case Is = "properties"
      Hide_frmFind
      Get_Properties
      Exit Sub
    Case Is = "open"
      On Error Resume Next
      CommonDialog1.ShowOpen
      
      If Err.Number <> 32755 Then
        rtb1(TabStrip1.SelectedItem.Index - 1).LoadFile CommonDialog1.FileName
        TabStrip1_Click
        File1_Click
      End If
      
      Exit Sub
    Case Is = "about"
      Hide_frmFind
      frmAbout.Show vbModal
    Case Is = "delete"
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
      SendKeys "{Del}", True
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
    Case Is = "undo"
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
      SendKeys "^z", True
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
    Case Is = "margin"
      frmMargin.Show
      Exit Sub
    Case Is = "filters"
      frmPreferences.TabStrip1.Tabs(1).Selected = True
      frmPreferences.Show vbModal
      Exit Sub
    Case Is = "folderOptions"
      frmPreferences.TabStrip1.Tabs(2).Selected = True
      frmPreferences.Show vbModal
      Exit Sub
    Case Is = "setColors"
      frmPreferences.TabStrip1.Tabs(3).Selected = True
      frmPreferences.Show vbModal
      Exit Sub
    Case Is = "window0"
      TabStrip1.Tabs(1).Selected = True
    Case Is = "window1"
      TabStrip1.Tabs(2).Selected = True
    Case Is = "window2"
      TabStrip1.Tabs(3).Selected = True
    Case Is = "window3"
      TabStrip1.Tabs(4).Selected = True
    Case Is = "window4"
      TabStrip1.Tabs(5).Selected = True
    Case Is = "window5"
      TabStrip1.Tabs(6).Selected = True
    Case Is = "window6"
      TabStrip1.Tabs(7).Selected = True
    Case Is = "window7"
      TabStrip1.Tabs(8).Selected = True
    Case Is = "window8"
      TabStrip1.Tabs(9).Selected = True
    Case Is = "window9"
      TabStrip1.Tabs(10).Selected = True
    Case Is = "exit"
      mnuExit_Click
      Exit Sub
    Case Is = "New"
      mnuFileNew_Click
    Case Is = "Close"
      mnuFileClose_Click
    Case Is = "save"
      If rtb1(TabStrip1.SelectedItem.Index - 1).Text = "" Then Exit Sub
      If rtb1(TabStrip1.SelectedItem.Index - 1).FileName = "" Then Exit Sub
      NoQuestionSave = True
      cmdSave_Click
    Case Is = "Find"
      mnuFind_Click
      Exit Sub
    Case Is = "selectAll"
        mnuSelectAll_Click
    Case Is = "copy"
      mnuCopy_Click
    Case Is = "paste"
      mnuPaste_Click
    Case Is = "clear"
      rtb1(TabStrip1.SelectedItem.Index - 1).Text = ""
      TabStrip1.SelectedItem.Caption = ""
      TabStrip1.SelectedItem.Image = 2
      txtFilename.Text = ""
      Parent.Caption = " Code Library 3.0"
      File1.Refresh
      
      If TabStrip1.SelectedItem.Index = 1 Then
        TabStrip1.SelectedItem.Caption = " Code Library"
      End If
      
    Case Is = "clearall"
      mnuClearAll_Click
    Case Is = "exit"
      mnuExit_Click
      Exit Sub
    Case Is = "print"
      mnuPrint_Click
    Case Is = "cut"
      mnuCut_Click
    Case Is = "properties"
      Hide_frmFind
      Get_Properties
      Exit Sub
    Case Is = "fonts"
      mnuFontOptions_Click
    Case Is = "rename"
      mnuRenameFile_Click
    Case Is = "readme"
      Get_Help
      Call ShellExecute(hWnd, "Open", "CODELIB.HLP", "", HelpPath, 1)
    Case Is = "controls"
      mnuHideDir_Click
    Case Is = "tooltips"
      mnuTipsOn_Click
    Case Is = "toolbars"
      mnuToolsOn_Click
      
      If Left(ToolsOn, 4) = "True" Then
        Parent.StatusBar1.Panels(2).Text = "The toolbar is currently on."
      Else
        Parent.StatusBar1.Panels(2).Text = "The toolbar is currently off."
      End If
      
    Case Is = "status"
      mnuStatusOn_Click
    Case Is = "print"
      mnuPrint_Click
    Case Is = "printall"
      mnuPrintAll_Click
    Case Is = "printselected"
      mnuPrintSelected_Click
    Case Is = "printsetup"
      mnuPrintsetup_Click
    Case Is = "minimize"
      Parent.WindowState = vbMinimized
    Case Is = "maximize"
      
      If Parent.WindowState = vbMaximized Then
        Tool.Caption = "Maximize"
        Parent.WindowState = vbNormal
      Else
        Tool.Caption = "Restore"
        Parent.WindowState = vbMaximized
      End If
      
    Case Is = "file1"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 0
    Case Is = "file2"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 1
    Case Is = "file3"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 2
    Case Is = "file4"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 3
    Case Is = "file5"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 4
    Case Is = "file6"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 5
    Case Is = "file7"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 6
    Case Is = "file8"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 7
    Case Is = "file9"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 8
    Case Is = "file10"
      PopUpToolBar.Bands("Toolbar").Tools("recent").CBListIndex = 9
    Case Is = "removeRecent"
      
      For i = 1 To 10
        MenuToolBar.Bands("subFile").Tools("file1").Caption = "Empty"
      Next

      retval = WritePrivateProfileString("Recent Documents", "Recent0", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent1", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent2", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent3", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent4", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent5", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent6", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent7", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent8", "", file)
      retval = WritePrivateProfileString("Recent Documents", "Recent9", "", file)
      GetRecentFiles
  End Select
  
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("MenuToolBar_Click-frmCodeLib", Err.Description)
End Sub

Private Sub PopUpToolBar_MouseEnter(ByVal Tool As ActiveBarLibraryCtl.Tool)
  
  Select Case Tool.Name
    Case Is = "clearclip"
      Parent.StatusBar1.Panels(2).Text = "Clear Windows Clipboard to free up some memory."
    Case Is = "saveas"
      Parent.StatusBar1.Panels(2).Text = "Save the current file with a different filename or in a different location."
    Case Is = "delete"
      Parent.StatusBar1.Panels(2).Text = "Delete the currenly selected text."
    Case Is = "filters"
      Parent.StatusBar1.Panels(2).Text = "Specify which File Filters to use right now."
    Case Is = "folderOptions"
      Parent.StatusBar1.Panels(2).Text = "Assign a folder to one of the Folder buttons on the toolbar."
    Case Is = "setColors"
      Parent.StatusBar1.Panels(2).Text = "Select the colors you would like for the background and foreground of this program."
    Case Is = "properties"
      Parent.StatusBar1.Panels(2).Text = "View the properties of this file."
    Case Is = "undo"
      Parent.StatusBar1.Panels(2).Text = "Undo the last edit operation."
    Case Is = "margin"
      Parent.StatusBar1.Panels(2).Text = "Set the right margin for easier viewing."
    Case Is = "map"
      Parent.StatusBar1.Panels(2).Text = "Map " & Dir1.Path & " to a button for quick access."
    Case Is = "openFolder"
      Parent.StatusBar1.Panels(2).Text = "Open this folder with Windows Explorer."
    Case Is = "open"
      Parent.StatusBar1.Panels(2).Text = "Open this file with the default program."
    Case Is = "File"
      Parent.StatusBar1.Panels(2).Text = "A list of common file commands."
    Case Is = "edit"
      Parent.StatusBar1.Panels(2).Text = "A list of common editing commands."
    Case Is = "options"
      Parent.StatusBar1.Panels(2).Text = "A list of options that are available for this program."
    Case Is = "fonts"
      Parent.StatusBar1.Panels(2).Text = "Choose the font style you would like to use."
    Case Is = "rename"
      Parent.StatusBar1.Panels(2).Text = "Rename the current file."
    Case Is = "about"
      Parent.StatusBar1.Panels(2).Text = "Some info about this program."
    Case Is = "readme"
      Parent.StatusBar1.Panels(2).Text = "A brief explanation of this program's features."
    Case Is = "controls"
    
      If Left(FullView, 4) = "True" Then
        Parent.StatusBar1.Panels(2).Text = "Controls are currently off."
      Else
        Parent.StatusBar1.Panels(2).Text = "Controls are currently on."
      End If
      
    Case Is = "status"
      
      If Left(StatusOn, 4) = "True" Then
        Parent.StatusBar1.Panels(2).Text = "Statusbar is currently on."
      End If
      
    Case Is = "toolbars"
      
      If Left(ToolsOn, 4) = "True" Then
        Parent.StatusBar1.Panels(2).Text = "The toolbar is currently on."
      Else
        Parent.StatusBar1.Panels(2).Text = "The toolbar is currently off."
      End If
      
    Case Is = "help"
      Parent.StatusBar1.Panels(2).Text = "View some information about this program."
    Case Is = "minimize"
      Parent.StatusBar1.Panels(2).Text = "Minimize this program to Windows taskbar."
    Case Is = "maximize"
      
      If MenuToolBar.Bands("subWindow").Tools("maximize").Caption = "Restore" Then
        Parent.StatusBar1.Panels(2).Text = "Restore this program to the normal size."
      Else
        Parent.StatusBar1.Panels(2).Text = "Maximize this program."
      End If
      
    Case Is = "New"
      Parent.StatusBar1.Panels(2).Text = "Creates a new tab so you can view multiple files."
    Case Is = "Close"
      Parent.StatusBar1.Panels(2).Text = "Close the currently selected file."
    Case Is = "save"
      Parent.StatusBar1.Panels(2).Text = "Save the currently selected file to disk."
    Case Is = "print"
      Parent.StatusBar1.Panels(2).Text = "Print options and print setup."
    Case Is = "printall"
      Parent.StatusBar1.Panels(2).Text = "Print the currently selected file."
    Case Is = "printselected"
      Parent.StatusBar1.Panels(2).Text = "Print only the currently highlighted text."
    Case Is = "printsetup"
      Parent.StatusBar1.Panels(2).Text = "Setup your printer, page layout and more."
    Case Is = "Find"
      Parent.StatusBar1.Panels(2).Text = "Find / Find Next / Replace"
    Case Is = "selectAll"
      Parent.StatusBar1.Panels(2).Text = "Select all text in the current file."
    Case Is = "cut"
      Parent.StatusBar1.Panels(2).Text = "Cut the selected text."
    Case Is = "copy"
      Parent.StatusBar1.Panels(2).Text = "Copy the selected text to the clipboard."
    Case Is = "paste"
      Parent.StatusBar1.Panels(2).Text = "Paste from the clipboard."
    Case Is = "clear"
      Parent.StatusBar1.Panels(2).Text = "Clear the contents of the selected text box."
    Case Is = "clearall"
      Parent.StatusBar1.Panels(2).Text = "Close all open files."
    Case Is = "recent"
      Parent.StatusBar1.Panels(2).Text = "A list of your 10 most recently opened files."
    Case Is = "exit"
      Parent.StatusBar1.Panels(2).Text = "Exit this program.  Have a nice day!"
    Case Is = "tooltips"
      
      If toolTipsOn = "True" Then
        Parent.StatusBar1.Panels(2).Text = "Tooltips are currently on."
      Else
        Parent.StatusBar1.Panels(2).Text = "Tooltips are currently off."
      End If
      
    Case Is = "renameFolder"
      Parent.StatusBar1.Panels(2).Text = "Rename the selected folder."
    Case Is = "deleteFolder"
      Parent.StatusBar1.Panels(2).Text = "Delete the selected folder."
    Case Is = "newFolder"
      Parent.StatusBar1.Panels(2).Text = "Create a new folder."
    Case Is = "deleteFile"
      Parent.StatusBar1.Panels(2).Text = "Delete the current file."
    Case Is = "Filters"
      Parent.StatusBar1.Panels(2).Text = "Specify the file filters for this program."
  End Select
  
End Sub

Private Sub MenuToolBar_MouseEnter(ByVal Tool As ActiveBarLibraryCtl.Tool)
  
  Select Case Tool.Name
    Case Is = "clearclip"
      Parent.StatusBar1.Panels(2).Text = "Clear Windows Clipboard to free up some memory."
    Case Is = "open"
      Parent.StatusBar1.Panels(2).Text = "Search for a file to open."
    Case Is = "saveas"
      Parent.StatusBar1.Panels(2).Text = "Save the current file with a different filename or in a different location."
    Case Is = "properties"
      Parent.StatusBar1.Panels(2).Text = "View the properties of this file."
    Case Is = "about"
      Parent.StatusBar1.Panels(2).Text = "Some info about this program."
    Case Is = "delete"
      Parent.StatusBar1.Panels(2).Text = "Delete the currenly selected text."
    Case Is = "undo"
      Parent.StatusBar1.Panels(2).Text = "Undo the last edit operation."
    Case Is = "margin"
      Parent.StatusBar1.Panels(2).Text = "Set the right margin for easier viewing."
    Case Is = "filters"
      Parent.StatusBar1.Panels(2).Text = "Specify which File Filters to use right now."
    Case Is = "folderOptions"
      Parent.StatusBar1.Panels(2).Text = "Assign a folder to one of the Folder buttons on the toolbar."
    Case Is = "setColors"
      Parent.StatusBar1.Panels(2).Text = "Select the colors you would like for the background and foreground of this program."
    Case Is = "window0"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window0").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window0").Caption) - 5)
    Case Is = "window1"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window1").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window1").Caption) - 5)
    Case Is = "window2"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window2").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window2").Caption) - 5)
    Case Is = "window3"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window3").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window3").Caption) - 5)
    Case Is = "window4"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window4").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window4").Caption) - 5)
    Case Is = "window5"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window5").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window5").Caption) - 5)
    Case Is = "window6"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window6").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window6").Caption) - 5)
    Case Is = "window7"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window7").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window7").Caption) - 5)
    Case Is = "window8"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window8").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window8").Caption) - 5)
    Case Is = "window9"
      Parent.StatusBar1.Panels(2).Text = "Switch to file: " & Right(MenuToolBar.Bands("subWindow").Tools("window9").Caption, Len(MenuToolBar.Bands("subWindow").Tools("window9").Caption) - 5)
    Case Is = "removeRecent"
      Parent.StatusBar1.Panels(2).Text = "Remove all entries from the recent file list."
    Case Is = "Font"
      Parent.StatusBar1.Panels(2).Text = "Choose the type of font this program will use."
    Case Is = "File"
      Parent.StatusBar1.Panels(2).Text = "Create a new tab, close a file or save the current file."
    Case Is = "edit"
      Parent.StatusBar1.Panels(2).Text = "A list of common editing commands."
    Case Is = "options"
      Parent.StatusBar1.Panels(2).Text = "A list of options that are available for this program."
    Case Is = "fonts"
      Parent.StatusBar1.Panels(2).Text = "Choose the type of font this program will use."
    Case Is = "rename"
      Parent.StatusBar1.Panels(2).Text = "Rename a file."
    Case Is = "about"
      Parent.StatusBar1.Panels(2).Text = "View some information about the author."
    Case Is = "readme"
      Parent.StatusBar1.Panels(2).Text = "Read the 'User's Manual' for this program."
    Case Is = "controls"
    
      If FullView = "True" Then
        Parent.StatusBar1.Panels(2).Text = "Controls are currently off."
      Else
        Parent.StatusBar1.Panels(2).Text = "Controls are currently on."
      End If
      
    Case Is = "status"
      
      If StatusOn = "True" Then
        Parent.StatusBar1.Panels(2).Text = "Statusbar is currently on."
      End If
      
    Case Is = "toolbars"
      
      If ToolsOn = "True" Then
        Parent.StatusBar1.Panels(2).Text = "The toolbar is currently on."
      Else
        Parent.StatusBar1.Panels(2).Text = "The toolbar is currently off."
      End If
      
    Case Is = "help"
      Parent.StatusBar1.Panels(2).Text = "View some information about this program."
    Case Is = "minimize"
      Parent.StatusBar1.Panels(2).Text = "Minimize this program to Windows taskbar."
    Case Is = "maximize"
      
      If MenuToolBar.Bands("subWindow").Tools("maximize").Caption = "Restore" Then
        Parent.StatusBar1.Panels(2).Text = "Restore this program to the normal size."
      Else
        Parent.StatusBar1.Panels(2).Text = "Maximize this program."
      End If
      
    Case Is = "preferences"
      Parent.StatusBar1.Panels(2).Text = "Configure this program to your liking."
    Case Is = "New"
      Parent.StatusBar1.Panels(2).Text = "Creates a new tab so you can view multiple files."
    Case Is = "Close"
      Parent.StatusBar1.Panels(2).Text = "Close the currently selected file."
    Case Is = "save"
      Parent.StatusBar1.Panels(2).Text = "Save the currently selected file to disk."
    Case Is = "print"
      Parent.StatusBar1.Panels(2).Text = "Print options and print setup."
    Case Is = "printall"
      Parent.StatusBar1.Panels(2).Text = "Print the currently selected file."
    Case Is = "printselected"
      Parent.StatusBar1.Panels(2).Text = "Print only the currently highlighted text."
    Case Is = "printsetup"
      Parent.StatusBar1.Panels(2).Text = "Setup your printer, page layout and more."
    Case Is = "Find"
      Parent.StatusBar1.Panels(2).Text = "Find / Find Next / Replace"
    Case Is = "selectAll"
      Parent.StatusBar1.Panels(2).Text = "Select all text in the current file."
    Case Is = "cut"
      Parent.StatusBar1.Panels(2).Text = "Cut the selected text."
    Case Is = "copy"
      Parent.StatusBar1.Panels(2).Text = "Copy the selected text to the clipboard."
    Case Is = "paste"
      Parent.StatusBar1.Panels(2).Text = "Paste from the clipboard."
    Case Is = "clear"
      Parent.StatusBar1.Panels(2).Text = "Clear the contents of the selected text box."
    Case Is = "clearall"
      Parent.StatusBar1.Panels(2).Text = "Close all open files."
    Case Is = "recent"
      Parent.StatusBar1.Panels(2).Text = "A list of your 10 most recently opened files."
    Case Is = "exit"
      Parent.StatusBar1.Panels(2).Text = "Exit the program.  Have a nice day!"
    Case Is = "tooltips"
      
      If toolTipsOn = True Then
        Parent.StatusBar1.Panels(2).Text = "Tooltips are currently on."
      Else
        Parent.StatusBar1.Panels(2).Text = "Tooltips are currently off."
      End If
      
    Case Is = "window"
      Parent.StatusBar1.Panels(2).Text = "Minimize, maximize or restore the window."
    Case Is = "file1"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 0: " & MenuToolBar.Bands("subFile").Tools("file1").Description
    Case Is = "file2"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 1: " & MenuToolBar.Bands("subFile").Tools("file2").Description
    Case Is = "file3"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 2: " & MenuToolBar.Bands("subFile").Tools("file3").Description
    Case Is = "file4"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 3: " & MenuToolBar.Bands("subFile").Tools("file4").Description
    Case Is = "file5"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 4: " & MenuToolBar.Bands("subFile").Tools("file5").Description
    Case Is = "file6"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 5: " & MenuToolBar.Bands("subFile").Tools("file6").Description
    Case Is = "file7"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 6: " & MenuToolBar.Bands("subFile").Tools("file7").Description
    Case Is = "file8"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 7: " & MenuToolBar.Bands("subFile").Tools("file8").Description
    Case Is = "file9"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 8: " & MenuToolBar.Bands("subFile").Tools("file9").Description
    Case Is = "file10"
      Parent.StatusBar1.Panels(2).Text = "Open RecentFile 9: " & MenuToolBar.Bands("subFile").Tools("file10").Description
  
  End Select
  
End Sub

Private Sub PopUpToolBar_MenuItemEnter(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Me.MousePointer = 0
End Sub

Private Sub PopUpToolBar_MenuItemExit(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Me.MousePointer = 99
End Sub

Private Sub MainToolBar_MouseEnter(ByVal Tool As ActiveBarLibraryCtl.Tool)
  ProjectFolder = frmPreferences.txtProject.Text
  RootDir = frmPreferences.txtRoot.Text
  FavFolder = frmPreferences.txtJunk.Text
  
  Select Case Tool.Name
    Case Is = "undo"
      Parent.StatusBar1.Panels(2).Text = "Undo the last edit operation."
    Case Is = "options"
      Parent.StatusBar1.Panels(2).Text = "Choose file filters, assign folders to the folder buttons and set background / foreground color."
    Case Is = "Help"
      Parent.StatusBar1.Panels(2).Text = "Read the 'User's Manual' for this program."
    Case Is = "root"
      Tool.ToolTipText = RootDir
      Parent.StatusBar1.Panels(2).Text = RootDir
    Case Is = "main"
      Tool.ToolTipText = ProjectFolder
      Parent.StatusBar1.Panels(2).Text = ProjectFolder
    Case Is = "temp"
      Tool.ToolTipText = FavFolder
      Parent.StatusBar1.Panels(2).Text = FavFolder
    Case Is = "Close"
      Parent.StatusBar1.Panels(2).Text = "Close the currently selected file."
    Case Is = "print"
      Parent.StatusBar1.Panels(2).Text = "Print the current document."
    Case Is = "save"
      Parent.StatusBar1.Panels(2).Text = "Save the currently selected file to disk."
    Case Is = "Find"
      Parent.StatusBar1.Panels(2).Text = "Use the Find/Find Next/Replace utility."
    Case Is = "selectAll"
      Parent.StatusBar1.Panels(2).Text = "Select all text in the current file."
    Case Is = "cut"
      Parent.StatusBar1.Panels(2).Text = "Cut the selected text."
    Case Is = "copy"
      Parent.StatusBar1.Panels(2).Text = "Copy the selected text to the clipboard."
    Case Is = "paste"
      
      If Clipboard.GetText <> "" Then
        Parent.StatusBar1.Panels(2).Text = "Paste from the clipboard."
      Else
        Parent.StatusBar1.Panels(2).Text = "The clipboard is empty."
      End If
      
    Case Is = "clear"
      Parent.StatusBar1.Panels(2).Text = "Clear the contents of the selected text box."
    Case Is = "exit"
      Parent.StatusBar1.Panels(2).Text = "Exit this program.  Have a nice day!"
  End Select
  
End Sub
'
'
'
'------------DISABLE/ENABLE MENU ITEMS------------'

Private Sub MenuToolBar_BandOpen(ByVal Band As ActiveBarLibraryCtl.Band)
  DoEvents
  Tools_Enabled
  
  If Band.Name <> "subFile" And _
     Band.Name <> "subWindow" And _
     Band.Name <> "Edit" Then Exit Sub
  
  If Band.Name = "subFile" Then
    GetRecentFiles
    
    If frmCodeLib.MenuToolBar.Bands("subFile").Tools("file1").Caption = "Empty" Then
      MenuToolBar.Bands("subFile").Tools("removeRecent").Visible = False
    Else
      MenuToolBar.Bands("subFile").Tools("removeRecent").Visible = True
    End If
    
    Exit Sub
  End If
  
'----------SUBWINDOW SECTION----------------------------|

  If Band.Name = "subWindow" Then
    Dim NumCaption As Integer
    
    For NumCaption = 1 To 10
      
      If TabStrip1.Tabs(NumCaption).Caption <> "" Then
        MenuToolBar.Bands("subWindow").Tools("window" & NumCaption - 1).Visible = True
        HowMany = HowMany + 1
        MenuToolBar.Bands("subWindow").Tools("window" & NumCaption - 1).Caption = "&" & HowMany - 1 & "..." & TabStrip1.Tabs(NumCaption).Caption
      Else
        MenuToolBar.Bands("subWindow").Tools("window" & NumCaption - 1).Visible = False
      End If
      
    Next
    
      
    For NumWindows = 2 To 9
      
      If TabStrip1.Tabs(NumWindows).Caption <> "" Then
        numOpen = numOpen + 1
      End If
      
    Next
    
    If TabStrip1.SelectedItem.Caption = "Code Library" Then
      MenuToolBar.Bands("subWindow").Tools("window0").Checked = True
    End If
    
  End If
  
'----------END SUBWINDOW SECTION----------------------------|
  
  HowMany = 0
  
End Sub

Private Sub Tools_Enabled()
  
  If rtb1(TabStrip1.SelectedItem.Index - 1).FileName <> "" _
    And TabStrip1.SelectedItem.Caption <> "" _
    And rtb1(TabStrip1.SelectedItem.Index - 1).Text <> "" Then
    PopUpToolBar.Bands("subFile").Tools("deleteFile").Enabled = True
    PopUpToolBar.Bands("subFile").Tools("rename").Enabled = True
    MenuToolBar.Bands("Edit").Tools("Find").Enabled = True
    
    If TabStrip1.SelectedItem.Index > 0 Then
      PopUpToolBar.Bands("subFile").Tools("Close").Enabled = True
      MenuToolBar.Bands("subFile").Tools("Close").Enabled = True
    End If
    
 Else
    PopUpToolBar.Bands("subFile").Tools("deleteFile").Enabled = False
    PopUpToolBar.Bands("subFile").Tools("rename").Enabled = False
    PopUpToolBar.Bands("subFile").Tools("Close").Enabled = False
    MenuToolBar.Bands("subFile").Tools("Close").Enabled = False
    MenuToolBar.Bands("Edit").Tools("Find").Enabled = False
  End If
  
  If rtb1(TabStrip1.SelectedItem.Index - 1).SelText = "" Then
    PopUpToolBar.Bands("Edit").Tools("cut").Enabled = False
    PopUpToolBar.Bands("subPrint").Tools("printselected").Enabled = False
    MenuToolBar.Bands("subPrint").Tools("printselected").Enabled = False
    MenuToolBar.Bands("Edit").Tools("cut").Enabled = False
    MenuToolBar.Bands("Edit").Tools("copy").Enabled = False
    MenuToolBar.Bands("Edit").Tools("delete").Enabled = False
    PopUpToolBar.Bands("Edit").Tools("cut").Enabled = False
    PopUpToolBar.Bands("Edit").Tools("copy").Enabled = False
    PopUpToolBar.Bands("Edit").Tools("delete").Enabled = False
  Else
    PopUpToolBar.Bands("Edit").Tools("cut").Enabled = True
    PopUpToolBar.Bands("subPrint").Tools("printselected").Enabled = True
    MenuToolBar.Bands("subPrint").Tools("printselected").Enabled = True
    MenuToolBar.Bands("Edit").Tools("cut").Enabled = True
    MenuToolBar.Bands("Edit").Tools("copy").Enabled = True
    MenuToolBar.Bands("Edit").Tools("delete").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("cut").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("copy").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("delete").Enabled = True
  End If
  
  If Parent.WindowState = vbMaximized Then
    PopUpToolBar.Bands("Toolbar").Tools("maximize").Caption = "Restore"
    MenuToolBar.Bands("subWindow").Tools("maximize").Caption = "Restore"
    PopUpToolBar.RecalcLayout
    MenuToolBar.RecalcLayout
  Else
    PopUpToolBar.Bands("Toolbar").Tools("maximize").Caption = "Maximize"
    MenuToolBar.Bands("subWindow").Tools("maximize").Caption = "Maximize"
    PopUpToolBar.RecalcLayout
    MenuToolBar.RecalcLayout
  End If
      
  If Clipboard.GetText = "" Then 'Clipboard is empty
    PopUpToolBar.Bands("Edit").Tools("paste").Enabled = False
    PopUpToolBar.Bands("Edit").Tools("clearclip").Enabled = False
    MenuToolBar.Bands("Edit").Tools("paste").Enabled = False
    MenuToolBar.Bands("Edit").Tools("clearclip").Enabled = False
  Else
    PopUpToolBar.Bands("Edit").Tools("paste").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("clearclip").Enabled = True
    MenuToolBar.Bands("Edit").Tools("paste").Enabled = True
    MenuToolBar.Bands("Edit").Tools("clearall").Enabled = True
    MenuToolBar.Bands("Edit").Tools("clearclip").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("clearall").Enabled = True
  End If
  
  If rtb1(TabStrip1.SelectedItem.Index - 1).Text = "" Then
    PopUpToolBar.Bands("subFile").Tools("New").Enabled = False
    PopUpToolBar.Bands("subFile").Tools("save").Enabled = False
    PopUpToolBar.Bands("subFile").Tools("saveas").Enabled = False
    PopUpToolBar.Bands("subFile").Tools("deleteFile").Enabled = False
    PopUpToolBar.Bands("subFile").Tools("rename").Enabled = False
    PopUpToolBar.Bands("subPrint").Tools("printall").Enabled = False
    PopUpToolBar.Bands("subPrint").Tools("printselected").Enabled = False
    PopUpToolBar.Bands("Edit").Tools("selectAll").Enabled = False
    PopUpToolBar.Bands("Edit").Tools("clear").Enabled = False
    PopUpToolBar.Bands("Edit").Tools("Find").Enabled = False
    MenuToolBar.Bands("Edit").Tools("selectAll").Enabled = False
    MenuToolBar.Bands("Edit").Tools("clear").Enabled = False
    MenuToolBar.Bands("subFile").Tools("New").Enabled = False
    MenuToolBar.Bands("subFile").Tools("Close").Enabled = False
    MenuToolBar.Bands("subFile").Tools("save").Enabled = False
    MenuToolBar.Bands("subFile").Tools("saveas").Enabled = False
    MenuToolBar.Bands("subPrint").Tools("printselected").Enabled = False
    MenuToolBar.Bands("subPrint").Tools("printall").Enabled = False
    MenuToolBar.Bands("Edit").Tools("Find").Enabled = False
  Else
    PopUpToolBar.Bands("subFile").Tools("New").Enabled = True
    PopUpToolBar.Bands("subFile").Tools("save").Enabled = True
    PopUpToolBar.Bands("subFile").Tools("saveas").Enabled = True
    PopUpToolBar.Bands("subPrint").Tools("printall").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("selectAll").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("clear").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("Find").Enabled = True
    MenuToolBar.Bands("Edit").Tools("selectAll").Enabled = True
    MenuToolBar.Bands("Edit").Tools("clear").Enabled = True
    MenuToolBar.Bands("subFile").Tools("New").Enabled = True
    MenuToolBar.Bands("subFile").Tools("save").Enabled = True
    MenuToolBar.Bands("subFile").Tools("saveas").Enabled = True
    MenuToolBar.Bands("subPrint").Tools("printall").Enabled = True
    MenuToolBar.Bands("Edit").Tools("Find").Enabled = True
  End If
  
  For Z = 0 To File1.ListCount - 1
    
    If File1.Selected(Z) = True Then
      PopUpToolBar.Bands("subFile").Tools("deleteFile").Enabled = True
      PopUpToolBar.Bands("FileOptions").Tools("properties").Enabled = True
      PopUpToolBar.Bands("subFile").Tools("properties").Enabled = True
      MenuToolBar.Bands("subFile").Tools("properties").Enabled = True
      Exit For
    Else
      PopUpToolBar.Bands("subFile").Tools("deleteFile").Enabled = False
      PopUpToolBar.Bands("FileOptions").Tools("properties").Enabled = False
      PopUpToolBar.Bands("subFile").Tools("properties").Enabled = False
      MenuToolBar.Bands("subFile").Tools("properties").Enabled = False
    End If
    
  Next
  
  If File1.FileName = "" Then
    PopUpToolBar.Bands("subFile").Tools("rename").Enabled = False
  Else
    PopUpToolBar.Bands("subFile").Tools("rename").Enabled = True
  End If
  
  For Z = 0 To TabStrip1.Tabs.Count - 1
    
    If rtb1(Z).Text = "" Then
      Dim LeaveEnabled As Boolean
      LeaveEnabled = False
    Else
      LeaveEnabled = True
      Exit For
    End If
    
  Next
  
  If LeaveEnabled = True Then
    MenuToolBar.Bands("Edit").Tools("clearall").Enabled = True
    PopUpToolBar.Bands("Edit").Tools("clearall").Enabled = True
  Else
    MenuToolBar.Bands("Edit").Tools("clearall").Enabled = False
    PopUpToolBar.Bands("Edit").Tools("clearall").Enabled = False
  End If

End Sub

Private Sub ButtonFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 2 Then
    PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
  End If
  
End Sub

Private Sub ButtonBar_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  
  Select Case Tool.Name
    Case Is = "new"
      cmdNew_Click
    Case Is = "clear"
      cmdClear_Click
    Case Is = "save"
      cmdSave_Click
  End Select
  
End Sub

Private Sub ButtonBar_MouseEnter(ByVal Tool As ActiveBarLibraryCtl.Tool)
  
  Select Case Tool.Name
    Case Is = "new"
      Parent.StatusBar1.Panels(2).Text = "Create a new tab to view other files."
    Case Is = "clear"
      Parent.StatusBar1.Panels(2).Text = "Close all open files."
    Case Is = "save"
      Parent.StatusBar1.Panels(2).Text = "Save the current text with the filename above."
  End Select
  
End Sub

Private Sub ButtonBar_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  Cancel = True
  PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
End Sub

'-------------
'POP-UP TOOLBAR
'-------------
Private Sub PopUpToolBar_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  On Error GoTo ErrorLog_Trapper
  rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  
  Select Case Tool.Name
    Case Is = "clearclip"
      Clipboard.Clear
    Case Is = "saveas"
      On Error Resume Next
      CommonDialog1.FileName = txtFilename.Text
      CommonDialog1.ShowSave
      
      If Err.Number <> 32755 Then
        rtb1(TabStrip1.SelectedItem.Index - 1).SaveFile CommonDialog1.FileName, 1
        File1.Refresh
      End If
      
    Case Is = "filters"
      frmPreferences.TabStrip1.Tabs(1).Selected = True
      frmPreferences.Show vbModal
      Exit Sub
    Case Is = "folderOptions"
      frmPreferences.TabStrip1.Tabs(2).Selected = True
      frmPreferences.Show vbModal
      Exit Sub
    Case Is = "setColors"
      frmPreferences.TabStrip1.Tabs(3).Selected = True
      frmPreferences.Show vbModal
      Exit Sub
    Case Is = "about"
      Hide_frmFind
      frmAbout.Show vbModal
    Case Is = "delete"
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
      SendKeys "{Del}", True
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
    Case Is = "undo"
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
      SendKeys "^z", True
      rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
    Case Is = "properties"
      Hide_frmFind
      Get_Properties
      Exit Sub
    Case Is = "margin"
      frmMargin.Show
    Case Is = "map"
      frmMap.Tag = Dir1.Path
      frmMap.Show vbModal
    Case Is = "openFolder"
      Call ShellExecute(hWnd, "Open", Dir1.Path, "", Dir1.Path, 1)
    Case Is = "open"
      mnuopen_Click
    Case Is = "New"
      mnuFileNew_Click
    Case Is = "save"
      If rtb1(TabStrip1.SelectedItem.Index - 1).Text = "" Then Exit Sub
      cmdSave_Click
    Case Is = "Find"
      mnuFind_Click
    Case Is = "selectAll"
        mnuSelectAll_Click
    Case Is = "copy"
      mnuCopy_Click
    Case Is = "paste"
      mnuPaste_Click
    Case Is = "clear"
      rtb1(TabStrip1.SelectedItem.Index - 1).Text = ""
      TabStrip1.SelectedItem.Caption = ""
      TabStrip1.SelectedItem.Image = 2
      txtFilename.Text = ""
      Parent.Caption = " Code Library 3.0"
      File1.Refresh
      
      If TabStrip1.SelectedItem.Index = 1 Then
        TabStrip1.SelectedItem.Caption = " Code Library"
      End If
      
    Case Is = "clearall"
      mnuClearAll_Click
    Case Is = "Close"
      mnuClose_Click
    Case Is = "exit"
      mnuExit_Click
    Case Is = "print"
      mnuPrint_Click
    Case Is = "cut"
      mnuCut_Click
    Case Is = "properties"
      Hide_frmFind
      Get_Properties
      Exit Sub
    Case Is = "fonts"
      mnuFontOptions_Click
    Case Is = "Filters"
      Hide_frmFind
      frmPreferences.TabStrip1.Tabs(1).Selected = True
      frmPreferences.Show vbModal
      Exit Sub
    Case Is = "rename"
      mnuRenameFile_Click
    Case Is = "readme"
      Get_Help
      Call ShellExecute(hWnd, "Open", "CODELIB.HLP", "", HelpPath, 1)
    Case Is = "controls"
      mnuHideDir_Click
    Case Is = "tooltips"
      mnuTipsOn_Click
    Case Is = "toolbars"
      mnuToolsOn_Click
    Case Is = "status"
      mnuStatusOn_Click
    Case Is = "print"
      mnuPrint_Click
    Case Is = "printall"
      mnuPrintAll_Click
    Case Is = "printselected"
      mnuPrintSelected_Click
    Case Is = "printsetup"
      mnuPrintsetup_Click
    Case Is = "deleteFile"
      mnuDel_Click
    Case Is = "newFolder"
      mnuNewFolder_Click
    Case Is = "renameFolder"
      mnuRename_Click
    Case Is = "deleteFolder"
      mnuDelFolder_Click
    Case Is = "minimize"
      Parent.WindowState = vbMinimized
    Case Is = "maximize"
      
      If Parent.WindowState = vbMaximized Then
        Tool.Caption = "Maximize"
        MenuToolBar.Bands("subWindow").Tools("maximize").Caption = "Maximize"
        MenuToolBar.RecalcLayout
        Parent.WindowState = vbNormal
      Else
        Tool.Caption = "Restore"
        MenuToolBar.Bands("subWindow").Tools("maximize").Caption = "Restore"
        MenuToolBar.RecalcLayout
        Parent.WindowState = vbMaximized
      End If
      
  End Select
  
  Exit Sub

ErrorLog_Trapper:
Call Log_Error("PopUpToolBar_Click-frmCodeLib", Err.Description)
End Sub

Private Sub PopUpToolBar_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  'By default, the ActiveBar control allows a
  'user to customize the toolbar.  I don't
  'see any use in someone messing around with
  'this program's toolbar, so I disabled it.
  Cancel = True
End Sub

Private Sub MainToolBar_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  Cancel = True
  PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
End Sub

Private Sub MenuToolBar_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  Cancel = True
  PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
End Sub

Private Sub PopUpToolBar_ComboDrop(ByVal Tool As ActiveBarLibraryCtl.Tool)
  On Error Resume Next
  Tool.CBList.Clear
  
  'Add recent files to comboBox
  For X = 1 To 10
    lengthOfName = InStrRev(MenuToolBar.Bands("subFile").Tools("file" & X).Caption, "\", Len(MenuToolBar.Bands("subFile").Tools("file" & X).Caption))
    filetoopen = Right(MenuToolBar.Bands("subFile").Tools("file" & X).Caption, Len(MenuToolBar.Bands("subFile").Tools("file" & X).Caption) - lengthOfName)
    If filetoopen = "Empty" Then Exit Sub
    If Left(filetoopen, 6) = "Recent" Then Exit Sub
    
    If filetoopen <> Tool.CBList.Item(X - 2) Then
      Tool.CBList.AddItem filetoopen
    End If
  
  Next
  
End Sub

Private Sub PopUpToolBar_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  On Error Resume Next
  filetoopen = MenuToolBar.Bands("subFile").Tools("file" & Tool.CBListIndex + 1).Description
  
  If filetoopen = "" Then
    rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
    Exit Sub
  End If
  
  Dim returnFile As String
  returnFile = Dir(filetoopen)
  
  If returnFile <> "" Then
    rtb1(TabStrip1.SelectedItem.Index - 1).LoadFile filetoopen, 1
    TabStrip1.Tabs(TabStrip1.SelectedItem.Index).Caption = returnFile
    UpdateFileMenu filetoopen
    TabStrip1_Click
    rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  Else
    Call MsgBox("That file no longer exists" & vbCr & "or is not accessible at this time.", 48, "File Access Error")
    rtb1(TabStrip1.SelectedItem.Index - 1).SetFocus
  End If
  
  TabStrip1_Click
End Sub
