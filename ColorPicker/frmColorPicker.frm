VERSION 5.00
Begin VB.Form frmColorPicker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Picker"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   795
      Left            =   3720
      Picture         =   "frmColorPicker.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Bye bye"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopyClipboard 
      Caption         =   "&Copy to Clipboard"
      Height          =   795
      Left            =   2520
      Picture         =   "frmColorPicker.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Copy the code to the clipboard"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.PictureBox fake 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   4080
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
      Begin VB.PictureBox picSlider 
         AutoRedraw      =   -1  'True
         Height          =   3900
         Left            =   120
         Picture         =   "frmColorPicker.frx":0454
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox SliderArrows 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   0
         Picture         =   "frmColorPicker.frx":1A54
         ScaleHeight     =   165
         ScaleWidth      =   525
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3780
         Width           =   525
      End
   End
   Begin VB.PictureBox Canvas 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   0
      MouseIcon       =   "frmColorPicker.frx":2E08
      MousePointer    =   99  'Custom
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   3900
      Begin VB.PictureBox FastCanvas 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   60
         ScaleHeight     =   32
         ScaleMode       =   0  'User
         ScaleWidth      =   16
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1155
      Left            =   1320
      ScaleHeight     =   1095
      ScaleWidth      =   1035
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox tR 
      Height          =   315
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4200
      Width           =   555
   End
   Begin VB.TextBox tG 
      Height          =   315
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4560
      Width           =   555
   End
   Begin VB.TextBox tB 
      Height          =   315
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4920
      Width           =   555
   End
   Begin VB.TextBox tRGB 
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Red :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   4260
      Width           =   650
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Green :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   4620
      Width           =   650
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Blue :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   4980
      Width           =   650
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 06/08/2001
' * Time             : 23:26
' * Module Name      : frmColorPicker
' * Module Filename  : frmColorPicker.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type LOGFONT
   lfHeight             As Long
   lfWidth              As Long
   lfEscapement         As Long
   lfOrientation        As Long
   lfWeight             As Long
   lfItalic             As Byte
   lfUnderline          As Byte
   lfStrikeOut          As Byte
   lfCharSet            As Byte
   lfOutPrecision       As Byte
   lfClipPrecision      As Byte
   lfQuality            As Byte
   lfPitchAndFamily     As Byte
   lfFaceName           As String * 32
End Type

Private Type BITMAPINFOHEADER '40 bytes
   biSize               As Long
   biWidth              As Long
   biHeight             As Long
   biPlanes             As Integer
   biBitCount           As Integer
   biCompression        As Long
   biSizeImage          As Long
   biXPelsPerMeter      As Long
   biYPelsPerMeter      As Long
   biClrUsed            As Long
   biClrImportant       As Long
End Type

Private Type BITMAPINFO
   bmiHeader            As BITMAPINFOHEADER
   bmiColors            As Long
End Type

Private Type DWORD
   low                  As Integer
   high                 As Integer
End Type

Private BBuffer(255, 255) As Long
Private FastBuffer(31, 31) As Long
Private bminfo          As BITMAPINFO
Private fbminfo         As BITMAPINFO
Private xpos            As Single
Private ypos            As Single
Private the_color       As Long

Private Type RGB
   r                    As Byte
   g                    As Byte
   b                    As Byte
   a                    As Byte
End Type

Private Type color
   Value                As Long
End Type

Private Sub Generate(lngColor As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : Generate
   ' * Parameters       :
   ' *                    lngColor As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim col              As color
   Dim rgbcol           As RGB

   col.Value = lngColor
   LSet rgbcol = col

   Dim r                As Double
   Dim g                As Double
   Dim b                As Double

   Dim x                As Double
   Dim y                As Double
   Dim dark             As Double
   Dim white            As Double
   Dim color            As Double

   b = rgbcol.b
   g = rgbcol.g
   r = rgbcol.r

   For y = 0 To 255
      For x = 0 To 255
         dark = (255 - y) / 255
         white = (255 - x)
         color = x / 255
         BBuffer(x, 255 - y) = RGB((white + b * color) * dark, (white + g * color) * dark, (white + r * color) * dark)
      Next
   Next
   SetDIBits Canvas.hdc, Canvas.Picture.Handle, 0, 256, BBuffer(0, 0), bminfo, 0
   'StretchDIBits Canvas.hdc, 0, 0, 255, 255, 0, 0, 127, 127, BBuffer(0, 0), bminfo, 0, vbSrcCopy
   Canvas.Cls
   Canvas.Refresh

End Sub

Private Sub FastGenerate(lngColor As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : FastGenerate
   ' * Parameters       :
   ' *                    lngColor As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim col              As color
   Dim rgbcol           As RGB

   col.Value = lngColor
   LSet rgbcol = col

   Dim r                As Double
   Dim g                As Double
   Dim b                As Double

   Dim x                As Double
   Dim y                As Double
   Dim dark             As Double
   Dim white            As Double
   Dim color            As Double

   b = rgbcol.b
   g = rgbcol.g
   r = rgbcol.r

   For y = 0 To 255 Step 8
      For x = 0 To 255 Step 8
         dark = (255 - y) / 255
         white = (255 - x)
         color = x / 255
         FastBuffer(x / 8, 31 - y / 8) = RGB((white + b * color) * dark, (white + g * color) * dark, (white + r * color) * dark)
      Next
   Next
   SetDIBits FastCanvas.hdc, FastCanvas.Picture.Handle, 0, 32, FastBuffer(0, 0), fbminfo, 0
   FastCanvas.Cls
   StretchBlt Canvas.hdc, 0, 0, 256, 256, FastCanvas.hdc, 0, 0, 32, 32, vbSrcCopy
   Canvas.Refresh
End Sub

Private Sub Canvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
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
   Canvas_MouseMove Button, Shift, x, y
End Sub

Private Sub Canvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
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
   If x > 255 Then x = 255
   If x < 0 Then x = 0
   If y > 255 Then y = 255
   If y < 0 Then y = 0
   Dim col              As color
   Dim rgba             As RGB

   If Button = 1 Then
      xpos = x
      ypos = y
      col.Value = BBuffer(x, 255 - y)
      LSet rgba = col
      picColor.BackColor = RGB(rgba.b, rgba.g, rgba.r)
      tR = rgba.b
      tG = rgba.g
      tB = rgba.r
      tRGB = "&H" & Hex(RGB(rgba.b, rgba.g, rgba.r))
      picColor.Refresh
      Canvas.Cls
      DrawCursor
   End If
End Sub

Private Sub cmdCancel_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : cmdCancel_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Me.Hide
End Sub

Private Sub cmdCopyClipboard_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : cmdCopyClipboard_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim sMsgbox          As String

   If tRGB.Text <> vbNullString Then
      sTmp = tRGB.Text

      ' *** Copy to the clipboard
      Clipboard.Clear
      Clipboard.SetText sTmp, vbCFText

      ' *** Copy done
      sMsgbox = "The color " & Chr$(13) & "has been copied to the clipboard." & Chr$(13)
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
   ' * Time             : 23:28
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
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
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   With bminfo.bmiHeader
      .biSize = Len(bminfo.bmiHeader)
      .biWidth = 256
      .biHeight = 256
      .biPlanes = 1
      .biBitCount = 32
   End With

   With fbminfo.bmiHeader
      .biSize = Len(bminfo.bmiHeader)
      .biWidth = 32
      .biHeight = 32
      .biPlanes = 1
      .biBitCount = 32
   End With
   Set Canvas.Picture = Canvas.Image
   Set FastCanvas.Picture = FastCanvas.Image
   Generate vbMagenta
End Sub

Private Sub picSlider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : picSlider_MouseDown
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
   picSlider_MouseMove Button, Shift, x, y
End Sub

Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : picSlider_MouseMove
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
   If y < 0 Then y = 0
   If y > 255 Then y = 0
   Dim col              As color
   Dim rgbcol           As RGB
   If Button = 1 Then
      col.Value = picSlider.Point(1, y) 'returns bgr not rgb :/
      the_color = col.Value
      LSet rgbcol = col
      FastGenerate col.Value
      'Canvas_MouseDown 1, 0, (xPos + 0), (yPos + 0)
      SliderArrows.top = y - 4
      fake.Refresh
      DrawCursor
   End If
End Sub

Private Sub picSlider_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : picSlider_MouseUp
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
   Generate the_color
   DrawCursor
End Sub

Private Sub DrawCursor()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:26
   ' * Module Name      : frmColorPicker
   ' * Module Filename  : frmColorPicker.frm
   ' * Procedure Name   : DrawCursor
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Canvas.Circle (xpos, ypos), 5, vbWhite
   Canvas.Refresh
End Sub
