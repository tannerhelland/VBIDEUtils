VERSION 5.00
Begin VB.Form frmCharPicker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Char Picker"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   675
      Left            =   5040
      Picture         =   "CharPicker.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Bye bye"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCopyClipboard 
      Caption         =   "&Copy to Clipboard"
      Height          =   675
      Left            =   3480
      Picture         =   "CharPicker.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Copy the code to the clipboard"
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Canvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   1
      Top             =   840
      Width           =   7755
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtAcii 
      Height          =   315
      Left            =   1860
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   3
      Top             =   480
      Width           =   675
   End
   Begin VB.TextBox txtChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   2580
      Locked          =   -1  'True
      MaxLength       =   1
      OLEDragMode     =   1  'Automatic
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Ascii Code :"
      Height          =   315
      Left            =   900
      TabIndex        =   5
      Top             =   540
      Width           =   1035
   End
End
Attribute VB_Name = "frmCharPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 06/08/2001
' * Time             : 23:08
' * Module Name      : frmCharTable
' * Module Filename  : CharTable.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit
Dim xpos                As Long
Dim ypos                As Long

Private Sub Canvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:08
   ' * Module Name      : frmCharTable
   ' * Module Filename  : CharTable.frm
   ' * Procedure Name   : Canvas_MouseDown
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    X As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim z                As Long
   If Button = 1 Then
      xpos = Int(x / 16)
      ypos = Int(y / 16)
      If ypos > 7 Then ypos = 7
      If xpos > 31 Then xpos = 31
      If ypos < 0 Then ypos = 0
      If xpos < 0 Then xpos = 0

      z = xpos + 32 * ypos
      txtChar.Text = Chr(z)
      txtAcii.Text = z
      Canvas.Cls
      Canvas.FillColor = RGB(100, 100, 100)
      Canvas.FillStyle = 0
      Canvas.DrawMode = 7
      Canvas.Line (xpos * 16, ypos * 16)-(xpos * 16 + 16, ypos * 16 + 16), RGB(100, 100, 100), B
      Canvas.DrawMode = 13
      Canvas.Refresh
   End If
End Sub

Private Sub Canvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:08
   ' * Module Name      : frmCharTable
   ' * Module Filename  : CharTable.frm
   ' * Procedure Name   : Canvas_MouseMove
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    X As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Canvas_MouseDown Button, Shift, x, y
End Sub

Private Sub cboFont_Change()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:08
   ' * Module Name      : frmCharTable
   ' * Module Filename  : CharTable.frm
   ' * Procedure Name   : cboFont_Change
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim x                As Long
   Dim y                As Long
   Dim z                As Long
   Set Canvas.Picture = Nothing
   Canvas.Cls
   Canvas.FontName = cboFont.Text
   Canvas.FontBold = False
   Canvas.FontItalic = False
   Canvas.FontSize = 10
   Canvas.ForeColor = RGB(190, 190, 190)
   For y = 1 To 8
      Canvas.Line (0, y * 16)-(Canvas.ScaleWidth, y * 16)
   Next
   For x = 1 To 32
      Canvas.Line (x * 16, 0)-(x * 16, Canvas.ScaleHeight)
   Next
   Canvas.ForeColor = 0
   For y = 0 To 7
      For x = 0 To 31
         Canvas.CurrentX = x * 16 + 2
         Canvas.CurrentY = y * 16 + 2
         Canvas.Print Chr(z)
         z = z + 1
      Next
   Next
   txtChar.FontName = Canvas.FontName
   txtChar.FontSize = 25
   txtChar.FontItalic = False
   txtChar.FontBold = False
   Canvas.Refresh
   Set Canvas.Picture = Canvas.Image
End Sub

Private Sub cboFont_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:08
   ' * Module Name      : frmCharTable
   ' * Module Filename  : CharTable.frm
   ' * Procedure Name   : cboFont_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   cboFont_Change
End Sub

Private Sub cmdCopyClipboard_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:08
   ' * Module Name      : frmCharTable
   ' * Module Filename  : CharTable.frm
   ' * Procedure Name   : cmdCopyClipboard_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim sMsgbox          As String

   If txtAcii.Text <> vbNullString Then
      sTmp = "Chr$(" & txtAcii.Text & ")"

      ' *** Copy to the clipboard
      Clipboard.Clear
      Clipboard.SetText sTmp, vbCFText

      ' *** Copy done
      sMsgbox = "The character " & Chr$(13) & "has been copied to the clipboard." & Chr$(13)
      sMsgbox = sMsgbox & "You can paste it in your code."

      Call MsgBoxTop(Me.hWnd, sMsgbox, vbOKOnly + vbInformation, "Char Picker")

   End If

   ' *** Exit
   Call cmdExit_Click

End Sub

Private Sub cmdExit_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:08
   ' * Module Name      : frmCharTable
   ' * Module Filename  : CharTable.frm
   ' * Procedure Name   : cmdExit_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call UnloadEffect(Me)
   Unload Me

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:08
   ' * Module Name      : frmCharTable
   ' * Module Filename  : CharTable.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim i                As Long
   cboFont.Clear
   For i = 0 To Screen.FontCount - 1
      cboFont.AddItem Screen.Fonts(i)
   Next
   cboFont.ListIndex = 0
End Sub
