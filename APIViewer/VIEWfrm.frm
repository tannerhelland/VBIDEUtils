VERSION 5.00
Begin VB.Form VIEWfrm 
   Caption         =   "PREVIEW"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VIEWfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   6720
      Begin VB.TextBox APItext 
         Height          =   1485
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   210
         Width           =   6450
      End
   End
   Begin VB.CommandButton VIEWpri 
      Caption         =   "Pr&ivate"
      Height          =   315
      Left            =   105
      TabIndex        =   2
      ToolTipText     =   "Ctrl + I"
      Top             =   1980
      Width           =   1365
   End
   Begin VB.CommandButton VIEWpub 
      Caption         =   "P&ublic"
      Height          =   315
      Left            =   1885
      TabIndex        =   3
      ToolTipText     =   "Ctrl + U"
      Top             =   1980
      Width           =   1365
   End
   Begin VB.CommandButton VIEWcop 
      Caption         =   "C&opy"
      Height          =   315
      Left            =   3665
      TabIndex        =   4
      ToolTipText     =   "Ctrl + C"
      Top             =   1980
      Width           =   1365
   End
   Begin VB.CommandButton VIEWok 
      Caption         =   "O&k"
      Height          =   315
      Left            =   5445
      TabIndex        =   5
      Top             =   1980
      Width           =   1365
   End
End
Attribute VB_Name = "VIEWfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim ll, lt As Long, retval1 As Boolean
   lt = (Screen.Height \ 2 - Me.Height \ 2) \ Screen.TwipsPerPixelY
   ll = (Screen.Width \ 2 - Me.Width \ 2) \ Screen.TwipsPerPixelX
   retval1 = SetWindowPos(Me.hwnd, HWND_TOPMOST, ll, lt, 0&, 0&, SWP_NOSIZE)
End Sub
Private Sub VIEWcop_Click()
   Clipboard.Clear
   Clipboard.SetText Trim(VIEWfrm.APItext.Text), vbCFText
   Unload Me
End Sub
Private Sub VIEWok_Click()
   Unload Me
End Sub
Private Sub VIEWpri_Click()
   MsgBox VIEWfrm.APItext.Text
   APItext.Text = "Private " & Mid(VIEWfrm.APItext.Text, 8, Len(APItext.Text))
   MsgBox VIEWfrm.APItext.Text
End Sub
Private Sub VIEWpub_Click()
   MsgBox VIEWfrm.APItext.Text
   APItext.Text = "Public" & Mid(VIEWfrm.APItext.Text, 8, Len(APItext.Text))
   MsgBox VIEWfrm.APItext.Text
End Sub
