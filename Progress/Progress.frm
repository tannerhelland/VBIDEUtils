VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Indenting Progressµ193"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   750
   ClientWidth     =   7020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser HTMLLoad 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
      ExtentX         =   1508
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   1140
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox imgCancel 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2400
      MouseIcon       =   "Progress.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Progress.frx":030A
      ScaleHeight     =   315
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1110
   End
   Begin MSComctlLib.ProgressBar ProgressHTML 
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image imgBanner 
      Height          =   900
      Left            =   0
      Picture         =   "Progress.frx":15AC
      Top             =   120
      Width           =   7020
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   2080
      X2              =   2080
      Y1              =   900
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2200
      X2              =   2200
      Y1              =   900
      Y2              =   2895
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   200
      Picture         =   "Progress.frx":86B8
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caption of progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   4485
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:51
' * Module Name      : frmProgress
' * Module Filename  : Progress.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Public bCancel          As Boolean

Private Sub Form_Activate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmProgress
   ' * Module Filename  : Progress.frm
   ' * Procedure Name   : Form_Activate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Me.Progress = 0

   gbCancelProgress = False
   If bCancel Then
      imgCancel.Visible = True
   Else
      imgCancel.Visible = False
   End If

End Sub

Property Let MessageText(sText As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmProgress
   ' * Module Filename  : Progress.frm
   ' * Procedure Name   : MessageText
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next
   lblMessage.Caption = sText
   DoEvents

End Property

Property Let Progress(nProgress As Double)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmProgress
   ' * Module Filename  : Progress.frm
   ' * Procedure Name   : Progress
   ' * Parameters       :
   ' *                    nProgress As Double
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   ProgressBar.Value = CLng(nProgress)
   Me.ZOrder 0
   DoEvents

End Property

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmProgress
   ' * Module Filename  : Progress.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   'Call YingYang(Me)

   imgBanner.left = (Me.Width - imgBanner.Width) / 2

   Me.MessageText = ""
   Me.Progress = 0
   bCancel = False

End Sub

Public Property Get Maximum() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmProgress
   ' * Module Filename  : Progress.frm
   ' * Procedure Name   : Maximum
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Maximum = ProgressBar.Max

End Property

Public Property Let Maximum(ByVal nNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmProgress
   ' * Module Filename  : Progress.frm
   ' * Procedure Name   : Maximum
   ' * Parameters       :
   ' *                    ByVal nNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   ProgressBar.Max = nNewValue
   ProgressBar.Value = 0
   ProgressBar.Min = 0

End Property

Private Sub imgCancel_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmProgress
   ' * Module Filename  : Progress.frm
   ' * Procedure Name   : imgCancel_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   gbCancelProgress = True
   Unload Me

End Sub

Private Sub HTMLLoad_ProgressChange(ByVal ProgressS As Long, ByVal ProgressMax As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmProgress
   ' * Module Filename  : Progress.frm
   ' * Procedure Name   : HTMLLoad_ProgressChange
   ' * Parameters       :
   ' *                    ByVal ProgressS As Long
   ' *                    ByVal ProgressMax As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   ProgressHTML.Max = ProgressMax
   If ProgressS > 0 Then
      ProgressHTML.Value = ProgressS
   Else
      ProgressHTML.Value = ProgressMax
   End If

End Sub
