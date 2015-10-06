VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNCMetrics"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SPI_GETICONMETRICS = 45
Private Const SPI_GETICONTITLELOGFONT = 31
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

' Normal log font structure:
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

' Here but thanks to a bit of hacking (testing cbSize
' parameter until it worked - see XXXTest1!) is the required
' definition for NONCLIENTMETRICS structure - differs
' to the one in WIN32API.TXT with VB.
' For some bizarre reason, maybe to do with byte
' alignment, the LOGFONT structure we must apply
' to NONCLIENTMETRICS seems to require an LF_FACESIZE
' 4 bytes smaller than normal:
Private Type NMLOGFONT
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
   lfFaceName(LF_FACESIZE - 4) As Byte
End Type
Private Type NONCLIENTMETRICS
   cbSize               As Long
   iBorderWidth         As Long
   iScrollWidth         As Long
   iScrollHeight        As Long
   iCaptionWidth        As Long
   iCaptionHeight       As Long
   lfCaptionFont        As NMLOGFONT
   iSMCaptionWidth      As Long
   iSMCaptionHeight     As Long
   lfSMCaptionFont      As NMLOGFONT
   iMenuWidth           As Long
   iMenuHeight          As Long
   lfMenuFont           As NMLOGFONT
   lfStatusFont         As NMLOGFONT
   lfMessageFont        As NMLOGFONT
End Type
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_SETNONCLIENTMETRICS = 42
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
   ByVal uAction As Long, _
   ByVal uParam As Long, _
   lpvParam As Any, _
   ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Enum CNCMetricsFontWeightConstants
   FW_DONTCARE = 0
   FW_THIN = 100
   FW_EXTRALIGHT = 200
   FW_ULTRALIGHT = 200
   FW_LIGHT = 300
   FW_NORMAL = 400
   FW_REGULAR = 400
   FW_MEDIUM = 500
   FW_SEMIBOLD = 600
   FW_DEMIBOLD = 600
   FW_BOLD = 700
   FW_EXTRABOLD = 800
   FW_ULTRABOLD = 800
   FW_HEAVY = 900
   FW_BLACK = 900
End Enum

Private m_tNCM          As NONCLIENTMETRICS
Private m_tLF           As LOGFONT
Private m_hFont(1 To 6) As Long

Public Enum CNCMetricsFontTypes
   IconFont = 1
   CaptionFont = 2
   SMCaptionFont = 3
   MenuFOnt = 4
   StatusFont = 5
   MessageFont = 6
End Enum

Public Function GetMetrics() As Boolean
   Dim lR               As Long

   ClearUp
   ' Get Non-client metrics:
   m_tNCM.cbSize = 340 'Len(m_tNCM) - why doesn't this go?
   lR = SystemParametersInfo( _
      SPI_GETNONCLIENTMETRICS, _
      0, _
      m_tNCM, _
      0)
   If (lR <> 0) Then
      ' Get icon font:
      lR = SystemParametersInfo( _
         SPI_GETICONTITLELOGFONT, _
         0, _
         m_tLF, _
         0)
      GetMetrics = (lR <> 0)
   End If
End Function
Property Get FontHandle( _
   ByVal eFontNum As CNCMetricsFontTypes _
   ) As Long
   Dim tLF              As LOGFONT
   Select Case eFontNum
      Case StatusFont
         CopyMemory tLF, m_tNCM.lfStatusFont, LenB(m_tNCM.lfStatusFont)
      Case SMCaptionFont
         CopyMemory tLF, m_tNCM.lfSMCaptionFont, LenB(m_tNCM.lfSMCaptionFont)
      Case MessageFont
         CopyMemory tLF, m_tNCM.lfMessageFont, LenB(m_tNCM.lfMessageFont)
      Case MenuFOnt
         CopyMemory tLF, m_tNCM.lfMenuFont, LenB(m_tNCM.lfMenuFont)
      Case IconFont
         CopyMemory tLF, m_tLF, LenB(m_tLF)
      Case CaptionFont
         CopyMemory tLF, m_tNCM.lfCaptionFont, LenB(m_tNCM.lfCaptionFont)
   End Select

   FontHandle = plGetFont(eFontNum, tLF)

End Property
Property Get BoldenedFontHandle( _
   ByVal eFontNum As CNCMetricsFontTypes _
   ) As Long
   Dim tLF              As LOGFONT
   Select Case eFontNum
      Case StatusFont
         CopyMemory tLF, m_tNCM.lfStatusFont, LenB(m_tNCM.lfStatusFont)
      Case SMCaptionFont
         CopyMemory tLF, m_tNCM.lfSMCaptionFont, LenB(m_tNCM.lfSMCaptionFont)
      Case MessageFont
         CopyMemory tLF, m_tNCM.lfMessageFont, LenB(m_tNCM.lfMessageFont)
      Case MenuFOnt
         CopyMemory tLF, m_tNCM.lfMenuFont, LenB(m_tNCM.lfMenuFont)
      Case IconFont
         CopyMemory tLF, m_tLF, LenB(m_tLF)
      Case CaptionFont
         CopyMemory tLF, m_tNCM.lfCaptionFont, LenB(m_tNCM.lfCaptionFont)
   End Select
   If (tLF.lfWeight < FW_BOLD) Then
      tLF.lfWeight = FW_BOLD
   Else
      tLF.lfWeight = FW_BLACK
   End If
   ' Not cleared up by this class....
   BoldenedFontHandle = CreateFontIndirect(tLF)
End Property
Property Get SmallerFontHandle( _
   ByVal eFontNum As CNCMetricsFontTypes _
   ) As Long
   Dim tLF              As LOGFONT
   Select Case eFontNum
      Case StatusFont
         CopyMemory tLF, m_tNCM.lfStatusFont, LenB(m_tNCM.lfStatusFont)
      Case SMCaptionFont
         CopyMemory tLF, m_tNCM.lfSMCaptionFont, LenB(m_tNCM.lfSMCaptionFont)
      Case MessageFont
         CopyMemory tLF, m_tNCM.lfMessageFont, LenB(m_tNCM.lfMessageFont)
      Case MenuFOnt
         CopyMemory tLF, m_tNCM.lfMenuFont, LenB(m_tNCM.lfMenuFont)
      Case IconFont
         CopyMemory tLF, m_tLF, LenB(m_tLF)
      Case CaptionFont
         CopyMemory tLF, m_tNCM.lfCaptionFont, LenB(m_tNCM.lfCaptionFont)
   End Select
   tLF.lfHeight = tLF.lfHeight * 5 / 6
   ' Not cleared up by this class....
   SmallerFontHandle = CreateFontIndirect(tLF)
End Property
Property Get CaptionHeight() As Long
   CaptionHeight = m_tNCM.iCaptionHeight
End Property
Property Get CaptionWIdth() As Long
   CaptionWIdth = m_tNCM.iCaptionWidth
End Property
Property Get MenuHeight() As Long
   MenuHeight = m_tNCM.iMenuHeight
End Property
Property Get MenuWidth() As Long
   MenuWidth = m_tNCM.iMenuWidth
End Property
Property Get ScrollHeight() As Long
   ScrollHeight = m_tNCM.iScrollHeight
End Property
Property Get ScrollWidth() As Long
   ScrollWidth = m_tNCM.iScrollWidth
End Property
Property Get SMCaptionHeight() As Long
   SMCaptionHeight = m_tNCM.iSMCaptionHeight
End Property
Property Get SMCaptionWIdth() As Long
   SMCaptionWIdth = m_tNCM.iSMCaptionWidth
End Property
Property Get BorderWidth() As Long
   BorderWidth = m_tNCM.iBorderWidth
End Property
Public Function ClearUp()
   Dim l                As Long
   For l = 1 To 6
      If (m_hFont(l) <> 0) Then
         DeleteObject m_hFont(l)
      End If
      m_hFont(l) = 0
   Next l
End Function
Private Function plGetFont( _
   ByVal eFontNum As CNCMetricsFontTypes, _
   ByRef tLF As LOGFONT _
   )
   If (m_hFont(eFontNum) = 0) Then
      m_hFont(eFontNum) = CreateFontIndirect(tLF)
   End If
   plGetFont = m_hFont(eFontNum)
End Function
Private Function XXXTest1()
   Dim lSize            As Long
   Dim lR               As Long
   For lSize = 0 To 1024
      m_tNCM.cbSize = lSize
      lR = SystemParametersInfo( _
         SPI_GETNONCLIENTMETRICS, _
         0, _
         m_tNCM, _
         0)
      If (lR <> 0) Then
         Debug.Print lSize
      End If
   Next lSize
End Function

Private Sub Class_Terminate()
   ClearUp
End Sub

