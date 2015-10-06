VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "VBIDEUtils"
   ClientHeight    =   6060
   ClientLeft      =   255
   ClientTop       =   465
   ClientWidth     =   8805
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0FF&
   Icon            =   "About.frx":0000
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "About.frx":27A2
   ScaleHeight     =   6060
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgRegisterCmd 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7200
      MouseIcon       =   "About.frx":808E
      MousePointer    =   99  'Custom
      Picture         =   "About.frx":8398
      ScaleHeight     =   315
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   3000
      Width           =   1110
   End
   Begin MSComctlLib.TreeView treeAbout 
      Height          =   2175
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3836
      _Version        =   393217
      Indentation     =   88
      LabelEdit       =   1
      Style           =   5
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   8480
      ScaleHeight     =   120
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   10
      Width           =   300
      Begin VB.Image imgClosebutton 
         Height          =   135
         Left            =   0
         Picture         =   "About.frx":963A
         Top             =   1
         Width           =   300
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   10
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   5
      Top             =   10
      Width           =   330
      Begin VB.Image imgMoveButton 
         Height          =   330
         Left            =   0
         MouseIcon       =   "About.frx":9898
         MousePointer    =   99  'Custom
         Picture         =   "About.frx":9BA2
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.PictureBox imgEmailCmd 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   6960
      MouseIcon       =   "About.frx":A1BC
      MousePointer    =   99  'Custom
      Picture         =   "About.frx":A4C6
      ScaleHeight     =   315
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   2640
      Width           =   1110
   End
   Begin VB.PictureBox imgWebCmd 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   6600
      MouseIcon       =   "About.frx":B768
      MousePointer    =   99  'Custom
      Picture         =   "About.frx":BA72
      ScaleHeight     =   315
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   2040
      Width           =   1110
   End
   Begin VB.PictureBox imgBackCmd 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7440
      MouseIcon       =   "About.frx":CD14
      MousePointer    =   99  'Custom
      Picture         =   "About.frx":D01E
      ScaleHeight     =   315
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   3360
      Width           =   1110
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "About.frx":E2C0
            Key             =   "CLOSE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "About.frx":E5DA
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "About.frx":E8F4
            Key             =   "MAIL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "About.frx":EC0E
            Key             =   "WEBLINK"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "About.frx":EF28
            Key             =   "PERSON"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "It is me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1999-2015"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   10
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label lbVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version ?.?µ323"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lbApp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VBIDEUtils"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   3795
      TabIndex        =   8
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label lblWebSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to : (Double click to see web site,  e-mail or register)µ309"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Image imgMoveFocus 
      Height          =   330
      Left            =   5400
      Picture         =   "About.frx":F242
      Top             =   6240
      Width           =   330
   End
   Begin VB.Image imgMoveNoFocus 
      Height          =   330
      Left            =   5160
      Picture         =   "About.frx":F85C
      Top             =   6240
      Width           =   330
   End
   Begin VB.Image imgMoveClicked 
      Height          =   330
      Left            =   4920
      Picture         =   "About.frx":FE76
      Top             =   6240
      Width           =   330
   End
   Begin VB.Image imgCloseNoFocus 
      Height          =   135
      Left            =   5760
      Picture         =   "About.frx":10490
      Top             =   6240
      Width           =   300
   End
   Begin VB.Image imgCloseClicked 
      Height          =   135
      Left            =   6480
      Picture         =   "About.frx":106EE
      Top             =   6240
      Width           =   300
   End
   Begin VB.Image imgCloseFocus 
      Height          =   135
      Left            =   6120
      Picture         =   "About.frx":1094C
      Top             =   6240
      Width           =   300
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:50
' * Module Name      : frmAbout
' * Module Filename  : About.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Public bAbout           As Boolean

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Type POINTAPI
   x                    As Long
   y                    As Long
End Type

Private Type RECT
   remLaunchPath        As String * 100
End Type

Private Const RGNAND = 1
Private Const RGNCOPY = 5
Private Const RGNDIFF = 4
Private Const RGNOR = 2
Private Const RGNXOR = 3

Dim mouseIsDown         As Boolean
Dim cx                  As Single
Dim cy                  As Single
Dim strLaunchPath       As String

Private Function CreateFormRegion() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : CreateFormRegion
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim ResultRegion     As Long
   Dim HolderRegion     As Long
   Dim ObjectRegion     As Long
   Dim nRet             As Long
   Dim PolyPoints()     As POINTAPI

   ResultRegion = CreateRectRgn(0, 0, 0, 0)
   HolderRegion = CreateRectRgn(0, 0, 0, 0)

   'This procedure was generated by VB Shaped Form Creator.  This copy has
   'NOT been registered for commercial use.  It may only be used for non-
   'profit making programs.  If you intend to sell your program, I think
   'it's only fair you pay for mine.  Commercial registration costs $20,
   'and can be performed online.  See "Registration" item on the help menu
   'for details.

   'Latest versions of VB Shaped Form Creator can be found at my website at
   'http://www.comports.com/AlexV/VBSFC.html or you can visit my main site
   'with many other free programs and utilities at http://www.comports.com/AlexV

   'Lines starting with '! are required for reading the form shape using the
   'Import Form command in VB Shaped Form Creator, but are not necessary for
   'Visual Basic to display the form correctly.

   '!Shaped Form Region Definition
   '!1,1,0,0,0,0,0,1
   '!:1,1,134,1,136,3,136,4,138,6,138,7,140,9,144,18,146,23,147,26,148,29,149,32,150,37,151,42,151,47,152,48,152,51,153,52,326,52,327,51,327,27,353,27,353,51,354,52,532,52,533,51,533,40,557,40,557,427,38,427,22,411,22,410,19,407,19,406,17,404,17,403,15,401,15,400,13,398,9,389,7,384,6,381,5,378,4,375,3,370,2,365,1,354,End
   ReDim PolyPoints(0 To 46)
   PolyPoints(0).x = 1
   PolyPoints(0).y = 1
   PolyPoints(1).x = 134
   PolyPoints(1).y = 1
   PolyPoints(2).x = 136
   PolyPoints(2).y = 3
   PolyPoints(3).x = 136
   PolyPoints(3).y = 4
   PolyPoints(4).x = 138
   PolyPoints(4).y = 6
   PolyPoints(5).x = 138
   PolyPoints(5).y = 7
   PolyPoints(6).x = 140
   PolyPoints(6).y = 9
   PolyPoints(7).x = 144
   PolyPoints(7).y = 18
   PolyPoints(8).x = 146
   PolyPoints(8).y = 23
   PolyPoints(9).x = 147
   PolyPoints(9).y = 26
   PolyPoints(10).x = 148
   PolyPoints(10).y = 29
   PolyPoints(11).x = 149
   PolyPoints(11).y = 32
   PolyPoints(12).x = 150
   PolyPoints(12).y = 37
   PolyPoints(13).x = 151
   PolyPoints(13).y = 42
   PolyPoints(14).x = 151
   PolyPoints(14).y = 47
   PolyPoints(15).x = 152
   PolyPoints(15).y = 48
   PolyPoints(16).x = 152
   PolyPoints(16).y = 51
   PolyPoints(17).x = 153
   PolyPoints(17).y = 52
   PolyPoints(18).x = 326
   PolyPoints(18).y = 52
   PolyPoints(19).x = 327
   PolyPoints(19).y = 51
   PolyPoints(20).x = 327
   PolyPoints(20).y = 27
   PolyPoints(21).x = 353
   PolyPoints(21).y = 27
   PolyPoints(22).x = 353
   PolyPoints(22).y = 51
   PolyPoints(23).x = 354
   PolyPoints(23).y = 52
   PolyPoints(24).x = 532
   PolyPoints(24).y = 52
   PolyPoints(25).x = 533
   PolyPoints(25).y = 51
   PolyPoints(26).x = 533
   PolyPoints(26).y = 40
   PolyPoints(27).x = 557
   PolyPoints(27).y = 40
   PolyPoints(28).x = 557
   PolyPoints(28).y = 427
   PolyPoints(29).x = 38
   PolyPoints(29).y = 427
   PolyPoints(30).x = 22
   PolyPoints(30).y = 411
   PolyPoints(31).x = 22
   PolyPoints(31).y = 410
   PolyPoints(32).x = 19
   PolyPoints(32).y = 407
   PolyPoints(33).x = 19
   PolyPoints(33).y = 406
   PolyPoints(34).x = 17
   PolyPoints(34).y = 404
   PolyPoints(35).x = 17
   PolyPoints(35).y = 403
   PolyPoints(36).x = 15
   PolyPoints(36).y = 401
   PolyPoints(37).x = 15
   PolyPoints(37).y = 400
   PolyPoints(38).x = 13
   PolyPoints(38).y = 398
   PolyPoints(39).x = 9
   PolyPoints(39).y = 389
   PolyPoints(40).x = 7
   PolyPoints(40).y = 384
   PolyPoints(41).x = 6
   PolyPoints(41).y = 381
   PolyPoints(42).x = 5
   PolyPoints(42).y = 378
   PolyPoints(43).x = 4
   PolyPoints(43).y = 375
   PolyPoints(44).x = 3
   PolyPoints(44).y = 370
   PolyPoints(45).x = 2
   PolyPoints(45).y = 365
   PolyPoints(46).x = 1
   PolyPoints(46).y = 354
   ObjectRegion = CreatePolygonRgn(PolyPoints(0), 47, 1)
   nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGNCOPY)
   DeleteObject ObjectRegion
   CreateFormRegion = ResultRegion

End Function

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Call InternationalizeForm(Me)

   If gsExpired <> "" Then
      frmAbout.bAbout = True
      frmAbout.Label1(2).AutoSize = True
      frmAbout.Label1(2).Caption = gsExpired
      frmAbout.Label1(2).ForeColor = &HC000&
      frmAbout.Label1(2).FontSize = 17
   End If

   If bAbout Then
      InitAbout
   Else
      InitStatistics
   End If

   Call InitTooltips

   ' *** Init ButtonPictures
   imgClosebutton = imgCloseNoFocus

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : Form_MouseDown
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Next two lines enable window drag from anywhere on form.  Uncomment them
   'to allow full window dragging.

   'ReleaseCapture
   'SendMessage Me.hWnd, &HA1, 2, 0&

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : Form_MouseMove
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   imgClosebutton = imgCloseNoFocus
   imgMoveButton = imgMoveNoFocus

End Sub

Private Sub picBlack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : picBlack_MouseMove
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   imgClosebutton = imgCloseNoFocus
   imgMoveButton = imgMoveNoFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error Resume Next
   Set clsTooltips = Nothing

End Sub

Private Sub imgBackCmd_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgBackCmd_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :We leave
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Unload Me

End Sub

Private Sub imgClosebutton_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgClosebutton_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   imgBackCmd_Click

End Sub

Private Sub imgClosebutton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgClosebutton_MouseDown
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   imgClosebutton = imgCloseClicked

End Sub

Private Sub imgClosebutton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgClosebutton_MouseMove
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   imgClosebutton = imgCloseFocus

End Sub

Private Sub imgClosebutton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgClosebutton_MouseUp
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   imgClosebutton = imgCloseNoFocus

End Sub

Private Sub imgEmailCmd_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgEmailCmd_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Start the email client

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ShellExecute 0&, "", "mailto:removed", "", "", vbNormalFocus

   If (err.number <> 0) Then
      MsgBox "Sorry, I failed to open the email client to send a email to: removed due to an error." & vbCrLf & vbCrLf & "[" & err.Description & "]", vbExclamation
   End If

End Sub

Private Sub imgMoveButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgMoveButton_MouseDown
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   imgMoveButton.Picture = imgMoveClicked.Picture
   mouseIsDown = True
   cx = x
   cy = y

End Sub

Private Sub imgMoveButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgMoveButton_MouseMove
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If imgMoveButton <> imgMoveClicked Then
      imgMoveButton.Picture = imgMoveFocus.Picture
   End If
   If mouseIsDown Then
      Me.Move Me.left + (x - cx), Me.top + (y - cy)
   End If

End Sub

Private Sub imgMoveButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgMoveButton_MouseUp
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   imgMoveButton.Picture = imgMoveNoFocus
   mouseIsDown = False

End Sub

Private Sub Form_Activate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : Form_Activate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Me.ZOrder 0
   AlwaysOnTop Me, True
   AlwaysOnTop Me, False

   DoEvents

   TextEffect Me, "", 12, 12, , 10, 0, RGB(&H80, 0, 0)

End Sub

Private Sub imgRegisterCmd_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgRegisterCmd_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Start the web browser

   On Error Resume Next

End Sub

Private Sub imgWebCmd_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : imgWebCmd_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Start the web browser

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ShellExecute 0&, "", "http://www.ppreview.net", "", "", vbNormalFocus

   If (err.number <> 0) Then
      MsgBox "Sorry, I failed to open the web site: http://www.ppreview.net due to an error." & vbCrLf & vbCrLf & "[" & err.Description & "]", vbExclamation
   End If

End Sub

Private Sub treeAbout_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : treeAbout_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cHourglass       As class_Hourglass

   On Error Resume Next

   If bAbout = False Then Exit Sub

   Select Case treeAbout.SelectedItem.Image
      Case "WEBLINK":
         ' *** Start the web browser

         On Error Resume Next

         Set cHourglass = New class_Hourglass

         ShellExecute 0&, "", treeAbout.SelectedItem.Text, "", "", vbNormalFocus

         If (err.number <> 0) Then
            MsgBox "Sorry, I failed to open the web site: " & treeAbout.SelectedItem.Text & " due to an error." & vbCrLf & vbCrLf & "[" & err.Description & "]", vbExclamation
         End If

      Case "MAIL":
         ' *** Start the email client

         On Error Resume Next

         Set cHourglass = New class_Hourglass

         ShellExecute 0&, "", "mailto:" & treeAbout.SelectedItem.Text, "", "", vbNormalFocus

         If (err.number <> 0) Then
            MsgBox "Sorry, I failed to open the email client to send a email to: removed due to an error." & vbCrLf & vbCrLf & "[" & err.Description & "]", vbExclamation
         End If

   End Select

End Sub

Private Sub InitAbout()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : InitAbout
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sKey             As String
   Dim sParent          As String

   ' *** Init the about form

   On Error Resume Next

   ' *** Show eeded controls
   If CInt(Rnd(50) * 100) Mod 2 Then
      imgBackCmd.left = 6720
      imgRegisterCmd.left = 6960
      imgEmailCmd.left = 7200
      imgWebCmd.left = 7440
   Else
      imgBackCmd.left = 7440
      imgRegisterCmd.left = 7200
      imgEmailCmd.left = 6960
      imgWebCmd.left = 6720
   End If

   lblWebSite.Visible = True
   imgWebCmd.Visible = True
   imgEmailCmd.Visible = True
   imgRegisterCmd.Visible = True
   imgBackCmd.Visible = True

   If Not gbRegistered Then
      lbApp.Caption = "UNREGISTERED " & gsREG_APP & " UNREGISTERED"
   Else
      lbApp.Caption = gsREG_APP & " registered to " & gsRegistered
   End If
   lbVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

   ' *** Add all links in the treeview
   treeAbout.Nodes.Clear

   ' *** FreeVBCode
   sParent = "FreeVBCode"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sParent

   sKey = "http://www.FreeVBCode.com/"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "feedback@FreeVBCode.com"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** SourceCode4Free
   sParent = "SourceCode4Free"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sParent

   sKey = "http://www.SourceCode4Free.com/"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "roger.johansson@snabbmat.se"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** VBWeb
   sParent = "VBWeb"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sParent

   sKey = "http://www.VBWeb.co.uk/"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "james@vbweb.co.uk"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** CodeHound VB Search
   sParent = "CodeHound, search in all VB sites"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sParent

   sKey = "http://www.codehound.com/"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "info@codehound.com"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** Bruno Paris
   sParent = "Beta tester : Bruno Paris"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sParent

   sKey = "Bruno Paris"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "ameba@zg.tel.hr"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** Visual Basic Accelerator
   sParent = "Visual Basic Accelerator"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sKey = "http://vbaccelerator.com/"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "Steve McMahon"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "steve@dogma.demon.co.uk"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** VBnet, The Visual Basic Developers Resource Centre
   sParent = "VBnet, The Visual Basic Developers Resource Centre"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sKey = "http://www.mvps.org/vbnet/"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "Randy Birch"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "randy_birch@msn.com"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** Karl E. Peterson's One-Stop Source Shop
   sParent = "Karl E. Peterson's One-Stop Source Shop"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sKey = "http://www.mvps.org/vb/"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "Karl E. Peterson"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "karl@mvps.org"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** vbhelp.NET - Your Online Visual Basic Source
   sParent = "vbhelp.NET - Your Online Visual Basic Source"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sKey = "http://www.vbhelp.net/"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "Georgia Wall"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "wallg@vbhelp.net"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** VB Helper
   sParent = "VB Helper"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sKey = "http://www.vb-helper.com"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "Rod Stephens"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "RodStephens@vb-helper.comformat"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** Visual Basic Web Magazine
   sParent = "Visual Basic Web Magazine"
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sKey = "http://www.bowmansoft.com/vbwm"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "Richard Bowman"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "vbwm@bowmansoft.com"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** Select the first one
   treeAbout.SelectedItem = treeAbout.Nodes(1)
   treeAbout.SelectedItem.EnsureVisible
   treeAbout.SelectedItem.Expanded = True

End Sub

Private Sub InitStatistics()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:50
   ' * Module Name      : frmAbout
   ' * Module Filename  : About.frm
   ' * Procedure Name   : InitStatistics
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sKey             As String
   Dim sParent          As String

   Dim sSQL             As String
   Dim record           As Recordset
   Dim record2          As Recordset
   Dim nI               As Long

   ' *** Init for statistics

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Hide uneeded controls
   lblWebSite.Visible = False
   imgWebCmd.Visible = False
   imgEmailCmd.Visible = False
   imgRegisterCmd.Visible = False

   If Not gbRegistered Then
      lbApp.Caption = "UNREGISTERED " & gsREG_APP & " UNREGISTERED"
   Else
      lbApp.Caption = gsREG_APP & " registered to " & gsRegistered
   End If
   lbVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

   ' *** General
   sParent = Translation("General statisticsµ174")
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sKey = "VBDiamond Source Code and Freeware"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "http://http://www.ppreview.net"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "WEBLINK"
   treeAbout.Nodes(sKey).ExpandedImage = "WEBLINK"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "removed"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   sKey = "removed"
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "MAIL"
   treeAbout.Nodes(sKey).ExpandedImage = "MAIL"
   treeAbout.Nodes(sKey).Tag = sKey

   If gbRegistered Then
      sKey = "Registered Version"
   Else
      sKey = "Shareware Version"
   End If
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "PERSON"
   treeAbout.Nodes(sKey).ExpandedImage = "PERSON"
   treeAbout.Nodes(sKey).Tag = sKey

   ' *** Categories
   sParent = Translation("Categoriesµ82")
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sSQL = "Select * From Categories "
   Set record = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
   If record.RecordCount > 0 Then
      record.MoveLast
      record.MoveFirst
   End If

   sKey = Translation("Number of Categories in the database : µ224") & " " & record.RecordCount & " " & Translation("itemµ195") & IIf(record.RecordCount > 1, "s", "")
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sKey

   For nI = 1 To record.RecordCount
      sSQL = "Select * From Items "
      sSQL = sSQL & "Where Category = " & ReadRecordSet(record, "ID")
      Set record2 = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
      If record2.RecordCount > 0 Then record2.MoveLast

      sKey = ReadRecordSet(record, "Category") & " : " & record2.RecordCount & " " & Translation("itemµ195") & IIf(record2.RecordCount > 1, "s", "")
      treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
      treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
      treeAbout.Nodes(sKey).Tag = sKey

      record2.Close
      Set record2 = Nothing

      record.MoveNext
   Next

   record.Close
   Set record = Nothing

   ' *** Items
   sParent = Translation("Itemsµ196")
   treeAbout.Nodes.Add , , sParent, sParent, "CLOSE"
   treeAbout.Nodes(sParent).ExpandedImage = "OPEN"
   treeAbout.Nodes(sParent).Tag = sParent

   sSQL = "Select * From Items "
   Set record = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
   If record.RecordCount > 0 Then record.MoveLast

   sKey = Translation("Number of Items in the database : µ229") & " " & record.RecordCount & " " & Translation("itemµ195") & IIf(record.RecordCount > 1, "s", "")
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sKey

   record.Close
   Set record = Nothing

   sSQL = "Select * From Code "
   Set record = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
   If record.RecordCount > 0 Then record.MoveLast

   sKey = Translation("Number of code in the database : µ225") & " " & record.RecordCount & " " & Translation("itemµ195") & IIf(record.RecordCount > 1, "s", "")
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sKey

   record.Close
   Set record = Nothing

   sSQL = "Select * From Samples "
   Set record = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
   If record.RecordCount > 0 Then record.MoveLast

   sKey = Translation("Number of samples in the database : µ230") & " " & record.RecordCount & " " & Translation("itemµ195") & IIf(record.RecordCount > 1, "s", "")
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sKey

   record.Close
   Set record = Nothing

   sSQL = "Select * From Articles "
   Set record = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
   If record.RecordCount > 0 Then record.MoveLast

   sKey = Translation("Number of articles in the database : µ223") & " " & record.RecordCount & " " & Translation("itemµ195") & IIf(record.RecordCount > 1, "s", "")
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sKey

   record.Close
   Set record = Nothing

   sSQL = "Select * From Files "
   Set record = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
   If record.RecordCount > 0 Then record.MoveLast

   sKey = Translation("Number of files in the database : µ226") & " " & record.RecordCount & " " & Translation("itemµ195") & IIf(record.RecordCount > 1, "s", "")
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sKey

   record.Close
   Set record = Nothing

   sSQL = "Select * From HTML Where Cache = True "
   Set record = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
   If record.RecordCount > 0 Then record.MoveLast

   sKey = Translation("Number of HTML pages stored in the database : µ228") & " " & record.RecordCount & " " & Translation("itemµ195") & IIf(record.RecordCount > 1, "s", "")
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sKey

   record.Close
   Set record = Nothing

   sSQL = "Select * From HTML Where Cache = False "
   Set record = DB.OpenRecordset(sSQL, DAO.dbOpenDynaset)
   If record.RecordCount > 0 Then record.MoveLast

   sKey = Translation("Number of HTML links in the database : µ227") & " " & record.RecordCount & " " & Translation("itemµ195") & IIf(record.RecordCount > 1, "s", "")
   treeAbout.Nodes.Add sParent, tvwChild, sKey, sKey, "CLOSE"
   treeAbout.Nodes(sKey).ExpandedImage = "OPEN"
   treeAbout.Nodes(sKey).Tag = sKey

   record.Close
   Set record = Nothing

End Sub

Private Sub InitTooltips(Optional Flags As TooltipFlagConstants)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/11/1999
   ' * Time             : 11:15
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : InitTooltips
   ' * Parameters       :
   ' *                    Optional Flags As TooltipFlagConstants
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer

   On Error Resume Next

   With clsTooltips
      .Create Me, ttfBalloon
      .MaxTipWidth = 70
      .Icon = itInfoIcon
      .Title = "VBIDEUtils"
      For nI = 0 To Me.Controls.Count
         If Trim$(Me.Controls(nI).ToolTipText) <> "" Then
            If err = 0 Then
               .AddTool Me.Controls(nI), tfTransparent, Me.Controls(nI).ToolTipText
               Me.Controls(nI).ToolTipText = ""
            Else
               err.Clear
            End If
         End If
      Next
   End With

End Sub
