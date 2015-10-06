Attribute VB_Name = "TextEffect_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:53
' * Module Name      : TextEffect_Module
' * Module Filename  : TextEffect.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long

Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const COLOR_BTNFACE = 15

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Private Const DT_DISPFILE = 6            '  Display-file
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_LEFT = &H0
Private Const DT_METAFILE = 5            '  Metafile, VDM
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_PLOTTER = 0             '  Vector plotter
Private Const DT_RASCAMERA = 3           '  Raster camera
Private Const DT_RASDISPLAY = 1          '  Raster display
Private Const DT_RASPRINTER = 2          '  Raster printer
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Public Sub TextEffect(obj As Object, ByVal sText As String, ByVal lX As Long, ByVal lY As Long, Optional ByVal bLoop As Boolean = False, Optional ByVal lStartSpacing As Long = 128, Optional ByVal lEndSpacing As Long = -1, Optional ByVal oColor As OLE_COLOR = vbWindowText)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : TextEffect_Module
   ' * Module Filename  : TextEffect.bas
   ' * Procedure Name   : TextEffect
   ' * Parameters       :
   ' *                    obj As Object
   ' *                    ByVal sText As String
   ' *                    ByVal lX As Long
   ' *                    ByVal lY As Long
   ' *                    Optional ByVal bLoop As Boolean = False
   ' *                    Optional ByVal lStartSpacing As Long = 128
   ' *                    Optional ByVal lEndSpacing As Long = -1
   ' *                    Optional ByVal oColor As OLE_COLOR = vbWindowText
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Kerning describes the spacing between characters when a font is written out.
   ' *** By default, fonts have a preset default kerning, but this very easy to modify
   ' *** under the Win32 API.

   ' *** The following (rather unusally named?) API function is all you need:

   ' *** Private Declare Function SetTextCharacterExtra Lib "gdi32" () (ByVal hdc As Long, ByVal nCharExtra As Long) As Long

   ' *** By setting nCharExtra to a negative value, you bring the characters closer together,
   ' *** and by setting to a positive values the characters space out.
   ' *** It works with VB's print methods too.

   Dim lHDC             As Long
   Dim i                As Long
   Dim x                As Long
   Dim lLen             As Long
   Dim hBrush           As Long
   Static tR            As RECT
   Dim iDir             As Long
   Dim lTime            As Long
   Dim lIter            As Long
   Dim bSlowDown        As Boolean
   Dim lColor           As Long
   Dim bDoIt            As Boolean

   lHDC = obj.hdc
   iDir = -1
   i = lStartSpacing
   tR.left = lX: tR.top = lY: tR.right = lX: tR.bottom = lY
   OleTranslateColor oColor, 0, lColor

   hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
   lLen = Len(sText)

   SetTextColor lHDC, lColor
   bDoIt = True

   Do While bDoIt
      lTime = timeGetTime
      If (i < -3) And Not (bLoop) And Not (bSlowDown) Then
         bSlowDown = True
         iDir = 1
         lIter = (i + 4)
      End If
      If (i > 128) Then iDir = -1
      If Not (bLoop) And iDir = 1 Then
         If (i = lEndSpacing) Then
            ' Stop
            bDoIt = False
         Else
            lIter = lIter - 1
            If (lIter <= 0) Then
               i = i + iDir
               lIter = (i + 4)
            End If
         End If
      Else
         i = i + iDir
      End If

      FillRect lHDC, tR, hBrush
      x = 32 - (i * lLen)
      SetTextCharacterExtra lHDC, i
      DrawText lHDC, sText, lLen, tR, DT_CALCRECT
      tR.right = tR.right + 4
      If (tR.right > obj.ScaleWidth \ Screen.TwipsPerPixelX) Then tR.right = obj.ScaleWidth \ Screen.TwipsPerPixelX
      DrawText lHDC, sText, lLen, tR, DT_LEFT
      obj.Refresh

      Do
         DoEvents
         If obj.Visible = False Then Exit Sub
      Loop While (timeGetTime - lTime) < 20

   Loop
   DeleteObject hBrush

End Sub
