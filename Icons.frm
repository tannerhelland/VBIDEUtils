VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmIcons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List of icons for toolbar"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pictVBDoc 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3120
      Picture         =   "Icons.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   51
      Top             =   2880
      Width           =   540
   End
   Begin VB.PictureBox ProjectAnalyzer 
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   120
      Picture         =   "Icons.frx":0C42
      ScaleHeight     =   300
      ScaleWidth      =   330
      TabIndex        =   50
      Top             =   720
      Width           =   390
   End
   Begin VB.PictureBox MouseZoom 
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   2760
      Picture         =   "Icons.frx":15EC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   49
      Top             =   600
      Width           =   315
   End
   Begin VB.PictureBox ListOLEServers 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3360
      Picture         =   "Icons.frx":19A2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   48
      Top             =   600
      Width           =   540
   End
   Begin VB.PictureBox ColorPicker 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3360
      Picture         =   "Icons.frx":4144
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   47
      Top             =   2520
      Width           =   540
   End
   Begin VB.PictureBox CharPicker 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3000
      Picture         =   "Icons.frx":4586
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   46
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox IndentProject 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1560
      Picture         =   "Icons.frx":4B10
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   45
      Top             =   600
      Width           =   300
   End
   Begin VB.PictureBox IndentModule 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1200
      Picture         =   "Icons.frx":4E9A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   44
      Top             =   600
      Width           =   300
   End
   Begin VB.PictureBox IndentProcedure 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   840
      Picture         =   "Icons.frx":5224
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   43
      Top             =   600
      Width           =   300
   End
   Begin VB.PictureBox CloseAllWindows 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3000
      Picture         =   "Icons.frx":55AE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   42
      Top             =   3600
      Width           =   300
   End
   Begin VB.PictureBox PropertyBuilder 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2640
      Picture         =   "Icons.frx":5B38
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   41
      Top             =   3600
      Width           =   300
   End
   Begin VB.PictureBox ExportHTML 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2280
      Picture         =   "Icons.frx":60C2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   40
      Top             =   3600
      Width           =   300
   End
   Begin VB.PictureBox GenerateGUID 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1920
      Picture         =   "Icons.frx":644C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   39
      Top             =   3600
      Width           =   300
   End
   Begin VB.PictureBox GenerateDLLBaseAdress 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1320
      Picture         =   "Icons.frx":6596
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   38
      Top             =   3360
      Width           =   540
   End
   Begin VB.PictureBox ADOConnectionCreator 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   600
      Picture         =   "Icons.frx":8D38
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   37
      Top             =   3360
      Width           =   540
   End
   Begin VB.PictureBox FindAndReplace 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   120
      Picture         =   "Icons.frx":C542
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   36
      Top             =   3600
      Width           =   300
   End
   Begin VB.PictureBox ProjProcHeader 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1920
      Picture         =   "Icons.frx":C68C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   35
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox Accelerator 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3960
      Picture         =   "Icons.frx":CC16
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   2520
      Width           =   540
   End
   Begin VB.PictureBox RemoveLineNumbering 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2640
      Picture         =   "Icons.frx":D058
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   32
      Top             =   3240
      Width           =   300
   End
   Begin VB.PictureBox AddLineNumbering 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2280
      Picture         =   "Icons.frx":D1A2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   31
      Top             =   3240
      Width           =   300
   End
   Begin VB.PictureBox VBDiamond 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2640
      Picture         =   "Icons.frx":D2EC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   30
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox ProjectExplorer 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1920
      Picture         =   "Icons.frx":D876
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox Dependencies 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2280
      Picture         =   "Icons.frx":DE00
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox Properties 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1560
      Picture         =   "Icons.frx":E38A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   27
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox StringExtractor 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1560
      Picture         =   "Icons.frx":E914
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox EnhancedErrorMod 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   840
      Picture         =   "Icons.frx":EA5E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox EnhancedErrorProc 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   840
      Picture         =   "Icons.frx":EDA0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox ErrorModule 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   120
      Picture         =   "Icons.frx":F0E2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox ErrorProcedure 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   120
      Picture         =   "Icons.frx":F46C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox ModProcHeader 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1560
      Picture         =   "Icons.frx":F7F6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox Swap 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2640
      Picture         =   "Icons.frx":FD80
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox Alphabetize 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1560
      Picture         =   "Icons.frx":1010A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox ActiveXExplorer 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   840
      Picture         =   "Icons.frx":10494
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   1920
      Width           =   540
   End
   Begin VB.PictureBox DBCreator 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "Icons.frx":108D6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   1920
      Width           =   540
   End
   Begin VB.PictureBox IconExplorer 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2760
      Picture         =   "Icons.frx":13078
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   1920
      Width           =   540
   End
   Begin VB.PictureBox Register 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2280
      Picture         =   "Icons.frx":134BA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   1320
      Width           =   540
   End
   Begin VB.PictureBox ClassSpy 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1560
      Picture         =   "Icons.frx":13A3C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   1320
      Width           =   540
   End
   Begin VB.PictureBox CodeDatabase 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   840
      Picture         =   "Icons.frx":13D46
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   1320
      Width           =   540
   End
   Begin VB.PictureBox Toolbar 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   120
      Picture         =   "Icons.frx":14188
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   300
   End
   Begin VB.PictureBox MsgBox 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   4080
      Picture         =   "Icons.frx":14712
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   300
   End
   Begin VB.PictureBox KeyCode 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3360
      Picture         =   "Icons.frx":14A9C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   1920
      Width           =   540
   End
   Begin VB.PictureBox APIError 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3720
      Picture         =   "Icons.frx":14DA6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   300
   End
   Begin VB.PictureBox ClearDebug 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3360
      Picture         =   "Icons.frx":15330
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   300
   End
   Begin VB.PictureBox TabOrder 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3000
      Picture         =   "Icons.frx":158BA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   300
   End
   Begin VB.PictureBox VBIDEUtils 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "Icons.frx":15E44
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Inline 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3120
      Picture         =   "Icons.frx":1614E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox Pending 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2640
      Picture         =   "Icons.frx":166D8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   300
   End
   Begin VB.PictureBox RemoveAll 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2760
      Picture         =   "Icons.frx":16C62
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox ProcHeader 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1200
      Picture         =   "Icons.frx":171EC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox ModHeader 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   840
      Picture         =   "Icons.frx":17776
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox AllComment 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2400
      Picture         =   "Icons.frx":17D00
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   240
      Width           =   300
   End
   Begin RichTextLib.RichTextBox rtfColorize 
      Height          =   735
      Left            =   3840
      TabIndex        =   34
      Top             =   3360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"Icons.frx":1828A
   End
End
Attribute VB_Name = "frmIcons"
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
' * Module Name      : frmIcons
' * Module Filename  : Icons.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

' *** Make transparent Bitmap
Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Sub MakeTransparentBitmap(OutDstDC As Long, DstDC As Long, SrcDC As Long, SrcRect As RECT, DstX As Integer, DstY As Integer, TransColor As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:28
   ' * Module Name      : frmIcons
   ' * Module Filename  : Icons.frm
   ' * Procedure Name   : MakeTransparentBitmap
   ' * Parameters       :
   ' *                    OutDstDC As Long
   ' *                    DstDC As Long
   ' *                    SrcDC As Long
   ' *                    SrcRect As RECT
   ' *                    DstX As Integer
   ' *                    DstY As Integer
   ' *                    TransColor As Long
   ' **********************************************************************
   ' * Comments         :
   ' * DstDC- Device context into which image must be
   ' * drawn transparently
   ' *
   ' * OutDstDC- Device context into image is actually drawn,
   ' * even though it is made transparent in terms of DstDC
   ' *
   ' * Src- Device context of source to be made transparent
   ' * in color TransColor
   ' *
   ' * SrcRect- Rectangular region within SrcDC to be made
   ' * transparent in terms of DstDC, and drawn to OutDstDC
   ' *
   ' * DstX, DstY - Coordinates in OutDstDC (and DstDC)
   ' * where the transparent bitmap must go. In most
   ' * cases, OutDstDC and DstDC will be the same
   ' *
   ' *
   ' **********************************************************************

   Dim nRet             As Long
   Dim w                As Integer
   Dim h                As Integer
   Dim MonoMaskDC       As Long
   Dim hMonoMask        As Long
   Dim MonoInvDC        As Long
   Dim hMonoInv         As Long
   Dim ResultDstDC      As Long
   Dim hResultDst       As Long
   Dim ResultSrcDC      As Long
   Dim hResultSrc       As Long
   Dim hPrevMask        As Long
   Dim hPrevInv         As Long
   Dim hPrevSrc         As Long
   Dim hPrevDst         As Long

   w = SrcRect.right - SrcRect.left + 1
   h = SrcRect.bottom - SrcRect.top + 1

   'create monochrome mask and inverse masks
   MonoMaskDC = CreateCompatibleDC(DstDC)
   MonoInvDC = CreateCompatibleDC(DstDC)
   hMonoMask = CreateBitmap(w, h, 1, 1, ByVal 0&)
   hMonoInv = CreateBitmap(w, h, 1, 1, ByVal 0&)
   hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
   hPrevInv = SelectObject(MonoInvDC, hMonoInv)

   'create keeper DCs and bitmaps
   ResultDstDC = CreateCompatibleDC(DstDC)
   ResultSrcDC = CreateCompatibleDC(DstDC)
   hResultDst = CreateCompatibleBitmap(DstDC, w, h)
   hResultSrc = CreateCompatibleBitmap(DstDC, w, h)
   hPrevDst = SelectObject(ResultDstDC, hResultDst)
   hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)

   'copy src to monochrome mask
   Dim OldBC            As Long
   OldBC = SetBkColor(SrcDC, TransColor)
   nRet = BitBlt(MonoMaskDC, 0, 0, w, h, SrcDC, SrcRect.left, SrcRect.top, vbSrcCopy)
   TransColor = SetBkColor(SrcDC, OldBC)

   'create inverse of mask
   nRet = BitBlt(MonoInvDC, 0, 0, w, h, MonoMaskDC, 0, 0, vbNotSrcCopy)

   'get background
   nRet = BitBlt(ResultDstDC, 0, 0, w, h, DstDC, DstX, DstY, vbSrcCopy)

   'AND with Monochrome mask
   nRet = BitBlt(ResultDstDC, 0, 0, w, h, MonoMaskDC, 0, 0, vbSrcAnd)

   'get overlapper
   nRet = BitBlt(ResultSrcDC, 0, 0, w, h, SrcDC, SrcRect.left, SrcRect.top, vbSrcCopy)

   'AND with inverse monochrome mask
   nRet = BitBlt(ResultSrcDC, 0, 0, w, h, MonoInvDC, 0, 0, vbSrcAnd)

   'XOR these two
   nRet = BitBlt(ResultDstDC, 0, 0, w, h, ResultSrcDC, 0, 0, vbSrcInvert)

   'output results
   nRet = BitBlt(OutDstDC, DstX, DstY, w, h, ResultDstDC, 0, 0, vbSrcCopy)

   'clean up
   hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
   DeleteObject hMonoMask

   hMonoInv = SelectObject(MonoInvDC, hPrevInv)
   DeleteObject hMonoInv

   hResultDst = SelectObject(ResultDstDC, hPrevDst)
   DeleteObject hResultDst

   hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
   DeleteObject hResultSrc

   DeleteDC MonoMaskDC
   DeleteDC MonoInvDC
   DeleteDC ResultDstDC
   DeleteDC ResultSrcDC

End Sub

Private Sub SetTransparent(pict As PictureBox)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmIcons
   ' * Module Filename  : Icons.frm
   ' * Procedure Name   : SetTransparent
   ' * Parameters       :
   ' *                    pict As PictureBox
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Set the image as transparent

   Dim r                As RECT

   With r
      .left = 0
      .top = 0
      .right = pict.ScaleWidth
      .bottom = pict.ScaleHeight
   End With

   'Call MakeTransparentBitmap(Pending.hdc, Pending.hdc, pict.hdc, R, 20, 20, &H8000000F)
   Call MakeTransparentBitmap(Pending.hdc, pict.hdc, pict.hdc, r, 0, 0, &HFF&)

End Sub

Private Sub Form_Load()

   'Call SetTransparent(Pending2)

End Sub
