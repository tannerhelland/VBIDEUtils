VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 09/11/1999
' * Time             : 11:44
' * Module Name      : class_Logo
' * Module Filename  : Logo.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

' ===========================================================================
' Name:     class_Logo
' Author:   Steve McMahon
' Date:     24 January 1999
'
' Shows how to draw a gradient and overlay an icon
' from an ImageList, then draw vertical text onto it.
'
' ---------------------------------------------------------------------------
' Visit vbAccelerator - advanced, free VB source code.
'     http://vbaccelerator.com
' ===========================================================================

Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Const LF_FACESIZE = 32
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
   lfFaceName(LF_FACESIZE) As Byte
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
   ByVal hIml As Long, _
   ByVal i As Long, _
   ByVal hdcDst As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal fStyle As Long _
   ) As Long
Private Const ILD_TRANSPARENT = 1

Private m_picThis       As PictureBox
Private m_sCaption      As String
Private m_bRGBStart(1 To 3) As Integer
Private m_oStartColor   As OLE_COLOR
Private m_bRGBEnd(1 To 3) As Integer
Private m_oEndColor     As OLE_COLOR
Private m_hIml          As Long
Private m_lIconIndex    As Long

Public Property Let hImageList(ByVal hIml As Long)
   m_hIml = hIml
End Property

Public Property Let IconIndex(ByVal lIndex As Long)
   m_lIconIndex = lIndex
End Property

Public Property Let Caption(ByVal sCaption As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : Caption
   ' * Parameters       :
   ' *                    ByVal sCaption As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_sCaption = sCaption
End Property

Public Property Get Caption() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : Caption
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Caption = m_sCaption
End Property

Public Property Let DrawingObject(ByRef picThis As PictureBox)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : DrawingObject
   ' * Parameters       :
   ' *                    ByRef picThis As PictureBox
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set m_picThis = picThis
End Property

Public Property Get StartColor() As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : StartColor
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   StartColor = m_oStartColor
End Property

Public Property Let StartColor(ByVal oColor As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : StartColor
   ' * Parameters       :
   ' *                    ByVal oColor As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lColor           As Long
   If (m_oStartColor <> oColor) Then
      m_oStartColor = oColor
      OleTranslateColor oColor, 0, lColor
      m_bRGBStart(1) = lColor And &HFF&
      m_bRGBStart(2) = ((lColor And &HFF00&) \ &H100)
      m_bRGBStart(3) = ((lColor And &HFF0000) \ &H10000)
      If Not (m_picThis Is Nothing) Then
         Draw
      End If
   End If

End Property

Public Property Get EndColor() As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : EndColor
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   EndColor = m_oEndColor
End Property

Public Property Let EndColor(ByVal oColor As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : EndColor
   ' * Parameters       :
   ' *                    ByVal oColor As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lColor           As Long
   If (m_oEndColor <> oColor) Then
      m_oEndColor = oColor
      OleTranslateColor oColor, 0, lColor
      m_bRGBEnd(1) = lColor And &HFF&
      m_bRGBEnd(2) = ((lColor And &HFF00&) \ &H100)
      m_bRGBEnd(3) = ((lColor And &HFF0000) \ &H10000)
      If Not (m_picThis Is Nothing) Then
         Draw
      End If
   End If
End Property

Public Sub Draw()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : Draw
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lHeight          As Long
   Dim lYStep           As Long
   Dim lY               As Long
   Dim bRGB(1 To 3)     As Integer
   Dim tLF              As LOGFONT
   Dim hFnt             As Long
   Dim hFntOld          As Long
   Dim lR               As Long
   Dim rct              As RECT
   Dim hBr              As Long
   Dim hdc              As Long
   Dim dR(1 To 3)       As Double
   On Error GoTo DrawError

   hdc = m_picThis.hdc
   lHeight = m_picThis.Height \ Screen.TwipsPerPixelY
   rct.right = m_picThis.Width \ Screen.TwipsPerPixelY
   ' Set a graduation of 255 pixels:
   lYStep = lHeight \ 255
   If (lYStep = 0) Then
      lYStep = 1
   End If
   rct.bottom = lHeight

   bRGB(1) = m_bRGBStart(1)
   bRGB(2) = m_bRGBStart(2)
   bRGB(3) = m_bRGBStart(3)
   dR(1) = m_bRGBEnd(1) - m_bRGBStart(1)
   dR(2) = m_bRGBEnd(2) - m_bRGBStart(2)
   dR(3) = m_bRGBEnd(3) - m_bRGBStart(3)

   For lY = lHeight To 0 Step -lYStep
      ' Draw bar:
      rct.top = rct.bottom - lYStep
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hdc, rct, hBr
      DeleteObject hBr
      rct.bottom = rct.top
      ' Adjust colour:
      bRGB(1) = m_bRGBStart(1) + dR(1) * (lHeight - lY) / lHeight
      bRGB(2) = m_bRGBStart(2) + dR(2) * (lHeight - lY) / lHeight
      bRGB(3) = m_bRGBStart(3) + dR(3) * (lHeight - lY) / lHeight
      'Debug.Print bRGB(1), (lHeight - lY) / lHeight
   Next lY

   If (m_hIml <> 0) And (m_lIconIndex > 0) Then
      lHeight = lHeight - 16
      ImageList_Draw m_hIml, m_lIconIndex, hdc, (rct.right - rct.left - 16) \ 2, lHeight - 2, ILD_TRANSPARENT
      lHeight = lHeight - 4
   Else
      lHeight = lHeight - 2
   End If

   pOLEFontToLogFont m_picThis.Font, hdc, tLF
   tLF.lfEscapement = 900
   hFnt = CreateFontIndirect(tLF)
   If (hFnt <> 0) Then
      hFntOld = SelectObject(hdc, hFnt)
      lR = TextOut(hdc, 0, lHeight, m_sCaption, Len(m_sCaption))
      SelectObject hdc, hFntOld
      DeleteObject hFnt
   End If

   m_picThis.Refresh
   Exit Sub

DrawError:
End Sub

Private Sub pOLEFontToLogFont(fntThis As StdFont, hdc As Long, tLF As LOGFONT)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : pOLEFontToLogFont
   ' * Parameters       :
   ' *                    fntThis As StdFont
   ' *                    hdc As Long
   ' *                    tLF As LOGFONT
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sFont            As String
   Dim iChar            As Integer

   ' Convert an OLE StdFont to a LOGFONT structure:
   With tLF
      sFont = fntThis.Name
      ' There is a quicker way involving StrConv and CopyMemory, but
      ' this is simpler!:
      For iChar = 1 To Len(sFont)
         .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
      Next iChar
      ' Based on the Win32SDK documentation:
      .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
      .lfItalic = fntThis.Italic
      If (fntThis.Bold) Then
         .lfWeight = FW_BOLD
      Else
         .lfWeight = FW_NORMAL
      End If
      .lfUnderline = fntThis.Underline
      .lfStrikeOut = fntThis.Strikethrough

   End With

End Sub

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:44
   ' * Module Name      : class_Logo
   ' * Module Filename  : Logo.cls
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   StartColor = &H0
   EndColor = vbButtonFace
End Sub

