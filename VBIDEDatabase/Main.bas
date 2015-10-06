Attribute VB_Name = "Main_Module"

' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:53
' * Module Name      : Main_Module
' * Module Filename  : Main.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private Const LIC_CLSID = "C444195A-7C4A-40D9-090D-AC1A2A35E29A"    'License ID
Private Const LIC_KEY = "056-9182736"                               'License Key

Global gsGUID           As String
Global gsExpired        As String
Global gnSynchronized   As Long

'Global Const gbRegistered = False
Global gbRegistered     As Boolean
Global gsRegistered     As String
Global gsLicenseKey     As String

Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' *** Returns the handle of the active window.
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

' Used to transfer side logo onto the owner-draw menu:
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lParam As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

' *** Preserve the refresh of a window
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, lp As Any) As Long

Public Const WM_USER = &H400
Public Const WM_QUIT = &H12
Public Const EM_FORMATRANGE As Long = WM_USER + 57

Global gsDBPath         As String
Global gsDatabase       As String
Global gsDBName         As String

Global gsOldDBPath      As String
Global gsOldDatabase    As String
Global gsOldDBName      As String

Global DB               As Database
Global gcolImages       As New Collection
Global gcolExtension    As New Collection
Global gbAutoColorizePaste As Boolean
Global gbAutoColorizeLoad As Boolean
Global gbHTML           As Boolean
Global gbAutoRefresh    As Boolean
Global gbTrial          As Boolean

Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type

Private Type CharRange
   cpMin                As Long
   cpMax                As Long
End Type

Private Type FormatRange
   hdc                  As Long
   hdcTarget            As Long
   rc                   As RECT
   rcPage               As RECT
   chrg                 As CharRange
End Type

Public Const OFS_MAXPATHNAME = 260
Public Const OF_READWRITE = &H2

Type OFSTRUCT
   cBytes               As Byte
   fFixedDisk           As Byte
   nErrCode             As Integer
   Reserved1            As Integer
   Reserved2            As Integer
   szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type FILETIME
   dwLowDateTime        As Long
   dwHighDateTime       As Long
End Type

Public Type SYSTEMTIME
   wYear                As Integer
   wMonth               As Integer
   wDayOfWeek           As Integer
   wDay                 As Integer
   wHour                As Integer
   wMinute              As Integer
   wSecond              As Integer
   wMilliseconds        As Long
End Type

Public Declare Sub GetLocalTime Lib "kernel32" _
   (lpSystemTime As SYSTEMTIME)

Public Declare Function GetFileTime Lib "kernel32" _
   (ByVal hFile As Long, lpCreationTime As FILETIME, _
   lpLastAccessTime As FILETIME, _
   lpLastWriteTime As FILETIME) As Long

Public Declare Function SetFileTime Lib "kernel32" _
   (ByVal hFile As Long, _
   lpCreationTime As FILETIME, _
   lpLastAccessTime As FILETIME, _
   lpLastWriteTime As FILETIME) As Long

Public Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, _
   lpSystemTime As SYSTEMTIME) As Long

Public Declare Function SystemTimeToFileTime Lib "kernel32" _
   (lpSystemTime As SYSTEMTIME, _
   lpFileTime As FILETIME) As Long

Public Declare Function OpenFile Lib "kernel32" _
   (ByVal lpFileName As String, _
   lpReOpenBuff As OFSTRUCT, _
   ByVal wStyle As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" _
   (ByVal hFile As Long) As Long

Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Type LVBKIMAGE
   ulFlags              As Long
   hbm                  As Long
   pszImage             As String
   cchImageMax          As Long
   xOffsetPercent       As Long
   yOffsetPercent       As Long
End Type

Public Const LVBKIF_SOURCE_URL = &H2
Public Const LVBKIF_STYLE_TILE = &H10
Public Const LVM_FIRST  As Long = &H1000
Public Const LVM_SETBKIMAGEA = (LVM_FIRST + 68)
Public Const LVM_SETBKIMAGE = LVM_SETBKIMAGEA
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const CLR_NONE = &HFFFFFFFF

Global gbDBLoaded       As Boolean
Global gbCancelProgress As Boolean
Global gbCompressed     As Boolean
Global gnColorize       As Integer ' 0 = VB  1 = Javascript  2 = Java
Global gbSaveImages     As Boolean
Global gbToolbarCaption As Boolean
Global gnCategoryChoosen As Long
Global gsCategoryChoosen As String

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&

Global Const gsCommentChar = "'"

Public Sub RemoveCancelMenuItem(frm As Form)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/02/2001
   ' * Time             : 12:39
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : RemoveCancelMenuItem
   ' * Parameters       :
   ' *                    frm As Form
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim hSysMenu         As Long
   'get the system menu for this form
   hSysMenu = GetSystemMenu(frm.hWnd, 0)
   'remove the close item
   Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
   'remove the separator that was over the close item
   Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

Function FileExist(sFile As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : FileExist
   ' * Parameters       :
   ' *                    sFile As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'It does not handle wildcard characters
   Dim lSize            As Long

   On Error Resume Next
   'Preset length to -1 because files can be zero bytes in length
   lSize = -1
   'Get the length of the file
   lSize = FileLen(sFile)
   If lSize > -1 Then
      FileExist = True
   Else
      FileExist = False
   End If

End Function

Function GetFilePath(sPathIn As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 22/10/1999
   ' * Time             : 16:53
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetFilePath
   ' * Parameters       :
   ' *                    sPathIn As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer

   For nI = Len(sPathIn) To 1 Step -1
      If InStr(":\", Mid$(sPathIn, nI, 1)) Then Exit For
   Next
   GetFilePath = left$(sPathIn, nI)

End Function

Function GetFileName(sFileIn As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 22/10/1999
   ' * Time             : 16:53
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetFileName
   ' * Parameters       :
   ' *                    sFileIn As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** This function will return just the file name from a
   ' *** string containing a path and file name.

   Dim nI               As Integer

   For nI = Len(sFileIn) To 1 Step -1
      If InStr("\", Mid$(sFileIn, nI, 1)) Then Exit For
   Next
   GetFileName = Mid$(sFileIn, nI + 1, Len(sFileIn) - nI)

End Function

Function GetFileNameNoExtension(sFileIn As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 22/10/1999
   ' * Time             : 16:53
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetFileNameNoExtension
   ' * Parameters       :
   ' *                    sFileIn As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** This function will return just the file name
   ' *** with no extension from a string containing
   ' *** a path and file name.

   Dim nI               As Integer
   Dim sTmp             As String

   For nI = Len(sFileIn) To 1 Step -1
      If InStr("\", Mid$(sFileIn, nI, 1)) Then Exit For
   Next
   sTmp = Mid$(sFileIn, nI + 1, Len(sFileIn) - nI)

   For nI = Len(sTmp) To 1 Step -1
      ' *** Find the last ocurrence of "." in the string
      If InStr(".", Mid$(sTmp, nI, 1)) Then Exit For
   Next

   GetFileNameNoExtension = left$(sTmp, nI - 1)

End Function

Function GetFileExt(sFileName As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 22/10/1999
   ' * Time             : 16:53
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetFileExt
   ' * Parameters       :
   ' *                    sFileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** This function will return the file extension from a
   ' *** string containing a path and file name.

   Dim nI               As Integer

   For nI = Len(sFileName) To 1 Step -1
      ' *** Find the last ocurrence of "." in the string
      If InStr(".", Mid$(sFileName, nI, 1)) Then Exit For
   Next
   GetFileExt = right$(sFileName, Len(sFileName) - nI)

End Function

Public Sub SetRedraw(frm As Object, bRedraw As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/07/98
   ' * Time             : 10:55
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : SetRedraw
   ' * Parameters       :
   ' *                    frm As Object
   ' *                    bRedraw As Boolean
   ' **********************************************************************
   ' * Comments         : Preserve the refresh of a window
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   If frm.hWnd = 0 Then Exit Sub

   If (bRedraw = True) Then
      LockWindowUpdate 0
      frm.Refresh
   Else
      LockWindowUpdate frm.hWnd
   End If
   DoEvents

End Sub

Public Sub PrintRTF(rtf As RichTextBox, nnLeftMarginWidth As Long, nnTopMarginHeight As Long, nnRightMarginWidth As Long, nnBottomMarginHeight As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:43
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : PrintRTF
   ' * Parameters       :
   ' *                    rtf As RichTextBox
   ' *                    nnLeftMarginWidth As Long
   ' *                    nnTopMarginHeight As Long
   ' *                    nnRightMarginWidth As Long
   ' *                    nnBottomMarginHeight As Long
   ' **********************************************************************
   ' * Comments         : Print RTF
   ' *
   ' *
   ' **********************************************************************

   Dim nLeftOffset      As Long
   Dim nTopOffset       As Long
   Dim nLeftMargin      As Long
   Dim nTopMargin       As Long
   Dim nRightMargin     As Long
   Dim nBottomMargin    As Long
   Dim fr               As FormatRange
   Dim rcDrawTo         As RECT
   Dim rcPage           As RECT
   Dim nTextLength      As Long
   Dim nNextCharPos     As Long
   Dim nRet             As Long

   Printer.Print Space$(1)
   Printer.ScaleMode = vbTwips
   nLeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
   nTopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)
   nLeftMargin = nnLeftMarginWidth - nLeftOffset
   nTopMargin = nnTopMarginHeight - nTopOffset
   nRightMargin = (Printer.Width - nnRightMarginWidth) - nLeftOffset
   nBottomMargin = (Printer.Height - nnBottomMarginHeight) - nTopOffset
   rcPage.left = 0
   rcPage.top = 0
   rcPage.right = Printer.ScaleWidth
   rcPage.bottom = Printer.ScaleHeight
   rcDrawTo.left = nLeftMargin
   rcDrawTo.top = nTopMargin
   rcDrawTo.right = nRightMargin
   rcDrawTo.bottom = nBottomMargin
   fr.hdc = Printer.hdc
   fr.hdcTarget = Printer.hdc
   fr.rc = rcDrawTo
   fr.rcPage = rcPage
   fr.chrg.cpMin = 0
   fr.chrg.cpMax = -1
   nTextLength = Len(rtf.Text)
   Do
      fr.hdc = Printer.hdc
      fr.hdcTarget = Printer.hdc
      nNextCharPos = SendMessage(rtf.hWnd, EM_FORMATRANGE, True, fr)
      If nNextCharPos >= nTextLength Then Exit Do
      fr.chrg.cpMin = nNextCharPos
      Printer.NewPage
      Printer.Print Space$(1)
   Loop
   Printer.EndDoc
   nRet = SendMessage(rtf.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))

End Sub

' #HOut# ********************
' #HOut# Programmer Name  : removed
' #HOut# Date             : 23/02/2000
' #HOut# Time             : 13:01
' #HOut# Comment          :
' #HOut# Comment          :
' #HOut# Comment          :
' #HOut# ********************
' #Out# Public Sub AlwaysOnTopHWND(nHwnd As Long, bOnTop As Boolean)
' #Out#    ' #VBIDEUtils#************************************************************
' #Out#    ' * Programmer Name  : removed
' #Out#    ' * Web Site         : http://www.ppreview.net
' #Out#    ' * E-Mail           : removed
' #Out#    ' * Date             : 09/09/1999
' #Out#    ' * Time             : 17:09
' #Out#    ' * Module Name      : Main_Module
' #Out#    ' * Module Filename  : Main.bas
' #Out#    ' * Procedure Name   : AlwaysOnTopHWND
' #Out#    ' * Parameters       :
' #Out#    ' *                    nHwnd As Long
' #Out#    ' *                    bOnTop As Boolean
' #Out#    ' **********************************************************************
' #Out#    ' * Comments         : Set a form as alway on top
' #Out#    ' *
' #Out#    ' * Pass any non-zero value to Place on top
' #Out#    ' * Pass zero to remove top-mostness
' #Out#    ' *
' #Out#    ' **********************************************************************
' #Out#
' #Out#    Const SWP_NOMOVE = 2
' #Out#    Const SWP_NOSIZE = 1
' #Out#    Const Flags = SWP_NOMOVE Or SWP_NOSIZE
' #Out#    Const HWND_TOPMOST = -1
' #Out#    Const HWND_NOTOPMOST = -2
' #Out#
' #Out#    If bOnTop Then
' #Out#       bOnTop = SetWindowPos(nHwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
' #Out#    Else
' #Out#       bOnTop = SetWindowPos(nHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
' #Out#    End If
' #Out#
' #Out# End Sub
' #Out#
' #HOut# ********************

Public Sub AlwaysOnTopHWND(nHwnd As Long, bOnTop As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/09/1999
   ' * Time             : 17:09
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : AlwaysOnTopHWND
   ' * Parameters       :
   ' *                    nHwnd As Long
   ' *                    bOnTop As Boolean
   ' **********************************************************************
   ' * Comments         : Set a form as alway on top
   ' *
   ' * Pass any non-zero value to Place on top
   ' * Pass zero to remove top-mostness
   ' *
   ' **********************************************************************

   Const SWP_NOMOVE = 2
   Const SWP_NOSIZE = 1
   Const Flags = SWP_NOMOVE Or SWP_NOSIZE
   Const HWND_TOPMOST = -1
   Const HWND_NOTOPMOST = -2

   If bOnTop Then
      bOnTop = SetWindowPos(nHwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
   Else
      bOnTop = SetWindowPos(nHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
   End If

End Sub

Public Sub AlwaysOnTop(frmID As Form, OnTop As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/09/1999
   ' * Time             : 17:09
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : AlwaysOnTop
   ' * Parameters       :
   ' *                    frmID As Form
   ' *                    OnTop As Integer
   ' **********************************************************************
   ' * Comments         : Set a form as alway on top
   ' *
   ' * Pass any non-zero value to Place on top
   ' * Pass zero to remove top-mostness
   ' *
   ' **********************************************************************

   Const SWP_NOMOVE = 2
   Const SWP_NOSIZE = 1
   Const Flags = SWP_NOMOVE Or SWP_NOSIZE
   Const HWND_TOPMOST = -1
   Const HWND_NOTOPMOST = -2

   If OnTop Then
      OnTop = SetWindowPos(frmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
   Else
      OnTop = SetWindowPos(frmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
   End If

End Sub

Public Function Split(ByVal sIn As String, sOut() As String, Optional sDelim As String = " ") As Variant
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/11/98
   ' * Time             : 12:06
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : Split
   ' * Parameters       :
   ' *                    ByVal sIn As String
   ' *                    sOut() As String
   ' *                    Optional sDelim As String
   ' **********************************************************************
   ' * Comments         : Split a string into a variant array
   ' *
   ' *
   ' **********************************************************************

   Dim nC               As Long
   Dim nI               As Long

   Dim nPos             As Long
   Dim noldPos          As Long

   If sDelim = "" Then
      Split = sIn
   End If

   nC = CountTokens(sIn, sDelim)

   ReDim sOut(nC + 1)

   nPos = 0
   noldPos = 1
   For nI = 1 To nC
      If nI = 93 Then
         Debug.Print
      End If
      nPos = InStr(nPos + 1, sIn, sDelim)

      If nPos > 0 Then
         sOut(nI) = Mid$(sIn, noldPos, nPos - noldPos)
         noldPos = nPos + Len(sDelim)
      End If
   Next
   sOut(nC + 1) = Mid$(sIn, noldPos, Len(sIn) - noldPos + 1)
   Split = sOut

End Function

Public Function Join(Source() As String, Optional sDelim As String = " ") As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/11/98
   ' * Time             : 12:06
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : Join
   ' * Parameters       :
   ' *                    Source() As String
   ' *                    Optional sDelim As String = " "
   ' **********************************************************************
   ' * Comments         : Join arrays elements
   ' *
   ' *
   ' **********************************************************************

   Dim sOut             As String
   Dim nI               As Long

   Dim sArray()         As String
   Dim nArray           As Long

   Dim nLow             As Long
   Dim nHigh            As Long

On Error GoTo errh:
   nLow = LBound(Source)
   nHigh = UBound(Source)

   nArray = 0

   ReDim sArray((nHigh \ 100) + 1)

   For nI = nLow To nHigh - 1
      If nI Mod 50 = 0 Then frmProgress.Progress = nHigh + nI
      If nI Mod 1000 = 0 Then nArray = nArray + 1
      sArray(nArray) = sArray(nArray) & Source(nI)
   Next
   sArray(nArray) = sArray(nArray) & Source(nHigh)
   sOut = ""
   For nI = 1 To nArray
      sOut = sOut & sArray(nI)
   Next
   Join = sOut

   Exit Function
errh:
   err.Raise err.number

End Function

Public Function CountTokens(sStr As String, sItem As String) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 22/10/1999
   ' * Time             : 16:53
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : CountTokens
   ' * Parameters       :
   ' *                    sStr As String
   ' *                    sItem As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Count the number of items in sStr

   Dim nPos             As Long
   Dim nCount           As Long

   nPos = InStr(sStr, sItem)
   nCount = 0

   Do While nPos > 0
      nCount = nCount + 1

      nPos = InStr(nPos + 1, sStr, sItem)
   Loop

   CountTokens = nCount

End Function

Public Function GetToken(s As String, token As String, ByVal Nth As Integer) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/10/98
   ' * Time             : 18:11
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetToken
   ' * Parameters       :
   ' *                    s As String
   ' *                    token As String
   ' *                    ByVal Nth As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *  This function returns the Nth token in a string
   ' *    Ex.  GetToken("This is a test.", " ", 2) = "is"
   ' *
   ' **********************************************************************

   Dim i                As Integer
   Dim P                As Integer
   Dim r                As Integer

   If Nth < 1 Then
      GetToken = ""
      Exit Function
   End If

   r = 0

   For i = 1 To Nth
      P = r
      r = InStr(P + 1, s, token)
      If r = 0 Then
         If i = Nth Then
            GetToken = Mid$(s, P + 1, Len(s) - P)
         Else
            GetToken = ""
         End If
         Exit Function
      End If
   Next

   GetToken = Mid$(s, P + 1, r - P - 1)

End Function

Public Function FindWindowLike(ByVal nHWNDStart As Long, sSearchWindowText As String, sSearchClassName As String) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/02/2000
   ' * Time             : 11:56
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : FindWindowLike
   ' * Parameters       :
   ' *                    ByVal nHWNDStart As Long
   ' *                    sSearchWindowText As String
   ' *                    sSearchClassName As String
   ' **********************************************************************
   ' * Comments         :
   ' *   Find Applications Matching a Specific Class or Window Title
   ' *
   ' **********************************************************************

   Dim nHwnd            As Long
   Dim sWindowText      As String
   Dim sClassName       As String
   Dim nRet             As Long

   Static nLevel        As Integer

   If nLevel = 0 Then If nHWNDStart = 0 Then nHWNDStart = GetDesktopWindow()

   ' *** Increase recursion counter
   nLevel = nLevel + 1

   ' *** Get first child window
   nHwnd = GetWindow(nHWNDStart, GW_CHILD)

   Do Until nHwnd = 0
      ' *** Search children by recursion
      nRet = FindWindowLike(nHwnd, sSearchWindowText, sSearchClassName)

      ' *** Get the window text and class name
      sWindowText = Space$(255)
      nRet = GetWindowText(nHwnd, sWindowText, 255)
      sWindowText = left$(sWindowText, nRet)

      sClassName = Space$(255)
      nRet = GetClassName(nHwnd, sClassName, 255)
      sClassName = left$(sClassName, nRet)

      ' *** Check if window found matches the search parameters
      If (sWindowText Like sSearchWindowText) And (sClassName Like sSearchClassName) Then
         FindWindowLike = nHwnd
         Exit Do
      End If

      ' *** Get next child window
      nHwnd = GetWindow(nHwnd, GW_HWNDNEXT)

   Loop

   ' *** Reduce the recursion counter
   nLevel = nLevel - 1

End Function

Public Function GetTrial() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 22/10/1999
   ' * Time             : 16:53
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetTrial
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Expire           As Date

   Dim sDate            As String
   Dim sTmp             As String

   Dim nFile            As Integer

   GetTrial = True
   gsExpired = ""
   If gbRegistered Then Exit Function

   On Error GoTo ERROR_GETFile
   nFile = FreeFile
   Open GetWindowsDirectory() & "Win.res" For Input As #nFile
   Input #nFile, gsGUID
   Close #nFile
END_GET_IT:

   On Error GoTo ERROR_handler

   Dim clsRegistry      As New class_Registry

   With clsRegistry
      .ClassKey = HKEY_CLASSES_ROOT
      .SectionKey = "CLSID\" & gsGUID
      .ValueKey = "A"
      .ValueType = REG_SZ
      gsRegistered = .Value
   End With

   With clsRegistry
      .ClassKey = HKEY_CLASSES_ROOT
      .SectionKey = "CLSID\" & gsGUID
      .ValueKey = "B"
      .ValueType = REG_SZ
      gsLicenseKey = .Value
   End With

   gbRegistered = True
   GetTrial = True
   Exit Function
   'If KeyGen(gsRegistered, "36BRI4TQ1WJ83ET6XYZ", 3) = gsLicenseKey Then
   '   gbRegistered = True
   '   GetTrial = True
   '   Exit Function
   'Else
   '   gbRegistered = False
   'End If

   With clsRegistry
      .ClassKey = HKEY_CLASSES_ROOT
      .SectionKey = "CLSID\" & gsGUID
      .ValueKey = ""
      .ValueType = REG_SZ
      sTmp = .Value
   End With
   sDate = Mid$(sTmp, 7, 2) & "/" & Mid$(sTmp, 1, 2) & "/" & Mid$(sTmp, 3, 4)

   If Abs(DateDiff("d", sDate, Date)) > 30 Then
      ' *** Expired
      gsExpired = Translation("It is time to register, the trial version has expiredµ194")
      frmAbout.Show vbModal
      GetTrial = False
   End If

   Exit Function

ERROR_GETFile:
   If err = 53 Then
      On Error Resume Next
      gsGUID = CreateGUID()
      Open GetWindowsDirectory() & "Win.res" For Binary Access Write As #nFile
      Put #nFile, , gsGUID
      Close #nFile
      Call SetOtherFileDate(GetWindowsDirectory() & "Win.res")

      With clsRegistry
         .ClassKey = HKEY_CLASSES_ROOT
         .SectionKey = "CLSID\" & gsGUID
         .ValueKey = ""
         .ValueType = REG_SZ
         .Value = Format$(Date, "MMYYYYDD")
      End With
   End If
   Resume END_GET_IT

ERROR_handler:
   GetTrial = False
   Exit Function

End Function

Public Sub SetOtherFileDate(sFileName As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 22/10/1999
   ' * Time             : 16:53
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : SetOtherFileDate
   ' * Parameters       :
   ' *                    sFileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim r                As Long
   Dim hFile            As Long
   Dim OFS              As OFSTRUCT
   Dim SYS_TIME         As SYSTEMTIME
   Dim FT_CREATE        As FILETIME
   Dim FT_ACCESS        As FILETIME
   Dim FT_WRITE         As FILETIME
   Dim NEW_TIME         As FILETIME

   Randomize

   hFile = OpenFile(sFileName, OFS, OF_READWRITE)

   If hFile Then
      GetLocalTime SYS_TIME
      SYS_TIME.wDay = Int(25 * Rnd + 1)
      SYS_TIME.wMonth = Int(12 * Rnd + 1)
      SYS_TIME.wYear = SYS_TIME.wYear - 1
      Call GetFileTime(hFile, FT_CREATE, FT_ACCESS, FT_WRITE)
      Call SystemTimeToFileTime(SYS_TIME, NEW_TIME)
      Call SetFileTime(hFile, NEW_TIME, NEW_TIME, NEW_TIME)
   End If

   Call CloseHandle(hFile)

End Sub

Public Sub SetListViewBkImage(frm As Form, lv As ListView)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/09/1999
   ' * Time             : 15:19
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : SetListViewBkImage
   ' * Parameters       :
   ' *                    frm As Form
   ' *                    lv As ListView
   ' **********************************************************************
   ' * Comments         : Tiling an image onto a list view
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_SetListViewBkImage

   Dim tLBI             As LVBKIMAGE
   Dim sTemp            As String

   sTemp = GetTempFileName()
   SavePicture frm.pictBackground, sTemp
   tLBI.pszImage = sTemp & Chr$(0)
   tLBI.cchImageMax = Len(sTemp) + 1
   tLBI.ulFlags = LVBKIF_SOURCE_URL Or LVBKIF_STYLE_TILE
   SendMessage lv.hWnd, LVM_SETBKIMAGE, 0, tLBI
   SendMessageByLong lv.hWnd, LVM_SETTEXTBKCOLOR, 0, CLR_NONE

EXIT_SetListViewBkImage:
   On Error Resume Next
   Kill sTemp
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_SetListViewBkImage:
   Resume EXIT_SetListViewBkImage

End Sub

Function oldRTF2HTML(sRTF As String, Optional sOptions As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Brady
   ' * Web Site         : http://www2.bitstream.net/~bradyh/downloads/rtf2htmlrm.html
   ' * E-Mail           : bradyh@bitstream.net
   ' * Date             : 15/09/1999
   ' * Time             : 15:02
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : RTF2HTML
   ' * Parameters       :
   ' *                    sRTF As String
   ' *                    Optional sOptions As String
   ' **********************************************************************
   ' * Comments         : revised by removed (removed)
   ' *
   ' *
   ' **********************************************************************

   'Version 2.11

   'The current version of this function is available at
   'http://www2.bitstream.net/~bradyh/downloads/rtf2html.zip

   'More information can be found at
   'http://www2.bitstream.net/~bradyh/downloads/rtf2htmlrm.html

   'Converts Rich Text encoded text to HTML format
   'if you find some text that this function doesn't
   'convert properly please email the text to
   'bradyh@bitstream.net

   'Options:
   '+H              add an HTML header and footer
   '+G              add a generator Metatag
   '+T="MyTitle"    add a title (only works if +H is used)
   Dim sHTML            As String
   Dim l                As Long
   Dim lTmp             As Long
   Dim lTmp2            As Long
   Dim lRTFLen          As Long
   Dim lBOS             As Long              'beginning of section
   Dim lEOS             As Long              'end of section
   Dim sTmp             As String
   Dim sTmp2            As String
   Dim sEOS             As String            'string to be added to end of section
   Dim sBOS             As String            'string to be added to beginning of section
   Dim sEOP             As String            'string to be added to end of paragraph
   Dim sBOL             As String            'string to be added to the begining of each new line
   Dim sEOL             As String            'string to be added to the end of each new line
   Dim sEOLL            As String            'string to be added to the end of previous line
   Dim sCurFont         As String            'current font code eg: "f3"
   Dim sCurFontSize     As String            'current font size eg: "fs20"
   Dim sCurColor        As String            'current font color eg: "cf2"
   Dim sFontFace        As String            'Font face for current font
   Dim sFontColor       As String            'Font color for current font
   Dim lFontSize        As Integer           'Font size for current font

   Const gHellFrozenOver = False             'always false
   Dim gSkip            As Boolean           'skip to next word/command
   Dim sCodes           As String              'codes for ascii to HTML char conversion
   Dim sCurLine         As String            'temp storage for text for current line before being added to sHTML
   Dim sColorTable()    As String            'table of colors
   Dim lColors          As Long              '# of colors
   Dim sFontTable()     As String            'table of fonts
   Dim lFonts           As Long              '# of fonts
   Dim sFontCodes       As String            'list of font code modifiers
   Dim bSeekinbText     As Boolean           'True if we have to hit text before inserting a </FONT>
   Dim bText            As Boolean           'true if there is text (as opposed to a control code) in sTmp
   Dim sAlign           As String            '"center" or "right"
   Dim gAlign           As Boolean           'if current text is aligned
   Dim sGen             As String            'Temp store for Generator Meta Tag if requested
   Dim sTitle           As String            'Temp store for Title if requested

   'setup HTML codes
   sCodes = "&nbsp;  {00}&copy;  {a9}&acute; {b4}&laquo; {ab}&raquo; {bb}&iexcl; {a1}&iquest;{bf}&Agrave;{c0}&agrave;{e0}&Aacute;{c1}"
   sCodes = sCodes & "&aacute;{e1}&Acirc; {c2}&acirc; {e2}&Atilde;{c3}&atilde;{e3}&Auml;  {c4}&auml;  {e4}&Aring; {c5}&aring; {e5}&AElig; {c6}"
   sCodes = sCodes & "&aelig; {e6}&Ccedil;{c7}&ccedil;{e7}&ETH;   {d0}&eth;   {f0}&Egrave;{c8}&egrave;{e8}&Eacute;{c9}&eacute;{e9}&Ecirc; {ca}"
   sCodes = sCodes & "&ecirc; {ea}&Euml;  {cb}&euml;  {eb}&Igrave;{cc}&igrave;{ec}&Iacute;{cd}&iacute;{ed}&Icirc; {ce}&icirc; {ee}&Iuml;  {cf}"
   sCodes = sCodes & "&iuml;  {ef}&Ntilde;{d1}&ntilde;{f1}&Ograve;{d2}&ograve;{f2}&Oacute;{d3}&oacute;{f3}&Ocirc; {d4}&ocirc; {f4}&Otilde;{d5}"
   sCodes = sCodes & "&otilde;{f5}&Ouml;  {d6}&ouml;  {f6}&Oslash;{d8}&oslash;{f8}&Ugrave;{d9}&ugrave;{f9}&Uacute;{da}&uacute;{fa}&Ucirc; {db}"
   sCodes = sCodes & "&ucirc; {fb}&Uuml;  {dc}&uuml;  {fc}&Yacute;{dd}&yacute;{fd}&yuml;  {ff}&THORN; {de}&thorn; {fe}&szlig; {df}&sect;  {a7}"
   sCodes = sCodes & "&para;  {b6}&micro; {b5}&brvbar;{a6}&plusmn;{b1}&middot;{b7}&uml;   {a8}&cedil; {b8}&ordf;  {aa}&ordm;  {ba}&not;   {ac}"
   sCodes = sCodes & "&shy;   {ad}&macr;  {af}&deg;   {b0}&sup1;  {b9}&sup2;  {b2}&sup3;  {b3}&frac14;{bc}&frac12;{bd}&frac34;{be}&times; {d7}"
   sCodes = sCodes & "&divide;{f7}&cent;  {a2}&pound; {a3}&curren;{a4}&yen;   {a5}...     {85}"

   'setup color table
   lColors = 0
   ReDim sColorTable(0)
   lBOS = InStr(sRTF, "\colortbl")
   If lBOS <> 0 Then
      lEOS = InStr(lBOS, sRTF, ";}")
      If lEOS <> 0 Then
         lBOS = InStr(lBOS, sRTF, "\red")
         While ((lBOS <= lEOS) And (lBOS <> 0))
            ReDim Preserve sColorTable(lColors)
            sTmp = Trim$(Hex(Mid$(sRTF, lBOS + 4, 1) & IIf(IsNumeric(Mid$(sRTF, lBOS + 5, 1)), Mid$(sRTF, lBOS + 5, 1), "") & IIf(IsNumeric(Mid$(sRTF, lBOS + 6, 1)), Mid$(sRTF, lBOS + 6, 1), "")))
            If Len(sTmp) = 1 Then sTmp = "0" & sTmp
            sColorTable(lColors) = sColorTable(lColors) & sTmp
            lBOS = InStr(lBOS, sRTF, "\green")
            sTmp = Trim$(Hex(Mid$(sRTF, lBOS + 6, 1) & IIf(IsNumeric(Mid$(sRTF, lBOS + 7, 1)), Mid$(sRTF, lBOS + 7, 1), "") & IIf(IsNumeric(Mid$(sRTF, lBOS + 8, 1)), Mid$(sRTF, lBOS + 8, 1), "")))
            If Len(sTmp) = 1 Then sTmp = "0" & sTmp
            sColorTable(lColors) = sColorTable(lColors) & sTmp
            lBOS = InStr(lBOS, sRTF, "\blue")
            sTmp = Trim$(Hex(Mid$(sRTF, lBOS + 5, 1) & IIf(IsNumeric(Mid$(sRTF, lBOS + 6, 1)), Mid$(sRTF, lBOS + 6, 1), "") & IIf(IsNumeric(Mid$(sRTF, lBOS + 7, 1)), Mid$(sRTF, lBOS + 7, 1), "")))
            If Len(sTmp) = 1 Then sTmp = "0" & sTmp
            sColorTable(lColors) = sColorTable(lColors) & sTmp
            lBOS = InStr(lBOS, sRTF, "\red")
            lColors = lColors + 1
         Wend
      End If
   End If

   'setup font table
   lFonts = 0
   ReDim sFontTable(0)
   lBOS = InStr(sRTF, "\fonttbl")
   If lBOS <> 0 Then
      lEOS = InStr(lBOS, sRTF, ";}}")
      If lEOS <> 0 Then
         lBOS = InStr(lBOS, sRTF, "\f0")
         While ((lBOS <= lEOS) And (lBOS <> 0))
            ReDim Preserve sFontTable(lFonts)
            While ((Mid$(sRTF, lBOS, 1) <> " ") And (lBOS <= lEOS))
               lBOS = lBOS + 1
            Wend
            lBOS = lBOS + 1
            sTmp = Mid$(sRTF, lBOS, InStr(lBOS, sRTF, ";") - lBOS)
            sFontTable(lFonts) = sFontTable(lFonts) & sTmp
            lBOS = InStr(lBOS, sRTF, "\f" & (lFonts + 1))
            lFonts = lFonts + 1
         Wend
      End If
   End If

   sHTML = "<PRE>"
   lRTFLen = Len(sRTF)
   'seek first line with text on it
   lBOS = InStr(sRTF, vbCrLf & "\deflang")
   If lBOS = 0 Then
      GoTo finally
   Else
      lBOS = lBOS + 2
   End If
   lEOS = InStr(lBOS, sRTF, vbCrLf & "\par")
   If lEOS = 0 Then GoTo finally

   While Not gHellFrozenOver
      sTmp = Mid$(sRTF, lBOS, lEOS - lBOS)
      l = lBOS
      While l <= lEOS
         sTmp = Mid$(sRTF, l, 1)
         Select Case sTmp
            Case "{"
               l = l + 1
            Case "}"
               sCurLine = sCurLine & sEOS
               sEOS = ""
               l = l + 1
            Case "\"    'special code
               l = l + 1
               sTmp = Mid$(sRTF, l, 1)
               Select Case sTmp
                  Case "b"
                     If ((Mid$(sRTF, l + 1, 1) = " ") Or (Mid$(sRTF, l + 1, 1) = "\")) Then
                        'b = bold
                        sCurLine = sCurLine & "<B>"
                        sEOS = "</B>" & sEOS
                        If (Mid$(sRTF, l + 1, 1) = " ") Then l = l + 1
                     ElseIf (Mid$(sRTF, l, 7) = "bullet ") Then
                        sTmp = "•"     'bullet
                        l = l + 6
                        bText = True
                     Else
                        gSkip = True
                     End If
                  Case "c"
                     If ((Mid$(sRTF, l, 2) = "cf") And (IsNumeric(Mid$(sRTF, l + 2, 1)))) Then
                        'cf = color font
                        lTmp = Val(Mid$(sRTF, l + 2, 5))
                        If lTmp <= UBound(sColorTable) Then
                           sCurColor = "cf" & lTmp
                           sFontColor = "#" & sColorTable(lTmp)
                           bSeekinbText = True
                        End If
                        'move "cursor" position to next rtf code
                        lTmp = l
                        While ((Mid$(sRTF, lTmp, 1) <> " ") And (Mid$(sRTF, lTmp, 1) <> "\"))
                           lTmp = lTmp + 1
                        Wend
                        If (Mid$(sRTF, lTmp, 1) = " ") Then
                           l = lTmp
                        Else
                           l = lTmp - 1
                        End If
                     Else
                        gSkip = True
                     End If
                  Case "e"
                     If (Mid$(sRTF, l, 7) = "emdash ") Then
                        sTmp = "—"
                        l = l + 6
                        bText = True
                     Else
                        gSkip = True
                     End If
                  Case "f"
                     If IsNumeric(Mid$(sRTF, l + 1, 1)) Then
                        'f# = font
                        'first get font number
                        lTmp = l + 2
                        sTmp2 = Mid$(sRTF, l + 1, 1)
                        While IsNumeric(Mid$(sRTF, lTmp, 1))
                           sTmp2 = sTmp2 & Mid$(sRTF, lTmp2, 1)
                           lTmp = lTmp + 1
                        Wend
                        lTmp = Val(sTmp2)
                        sCurFont = "f" & lTmp
                        If ((lTmp <= UBound(sFontTable)) And (sFontTable(lTmp) <> sFontTable(0))) Then
                           'insert codes if lTmp is a valid font # AND the font is not the default font
                           sFontFace = sFontTable(lTmp)
                           bSeekinbText = True
                        End If
                        'move "cursor" position to next rtf code
                        lTmp = l
                        While ((Mid$(sRTF, lTmp, 1) <> " ") And (Mid$(sRTF, lTmp, 1) <> "\"))
                           lTmp = lTmp + 1
                        Wend
                        If (Mid$(sRTF, lTmp, 1) = " ") Then
                           l = lTmp
                        Else
                           l = lTmp - 1
                        End If
                     ElseIf ((Mid$(sRTF, l + 1, 1) = "s") And (IsNumeric(Mid$(sRTF, l + 2, 1)))) Then
                        'fs# = font size
                        'first get font size
                        lTmp = l + 3
                        sTmp2 = Mid$(sRTF, l + 2, 1)
                        While IsNumeric(Mid$(sRTF, lTmp, 1))
                           sTmp2 = sTmp2 & Mid$(sRTF, lTmp, 1)
                           lTmp = lTmp + 1
                        Wend
                        lTmp = Val(sTmp2)
                        sCurFontSize = "fs" & lTmp
                        lFontSize = Int((lTmp / 5) - 2)
                        If lFontSize = 2 Then
                           sCurFontSize = ""
                           lFontSize = 0
                        Else
                           bSeekinbText = True
                           If lFontSize > 8 Then lFontSize = 8
                           If lFontSize < 1 Then lFontSize = 1
                        End If
                        'move "cursor" position to next rtf code
                        lTmp = l
                        While ((Mid$(sRTF, lTmp, 1) <> " ") And (Mid$(sRTF, lTmp, 1) <> "\"))
                           lTmp = lTmp + 1
                        Wend
                        If (Mid$(sRTF, lTmp, 1) = " ") Then
                           l = lTmp
                        Else
                           l = lTmp - 1
                        End If
                     Else
                        gSkip = True
                     End If
                  Case "i"
                     If ((Mid$(sRTF, l + 1, 1) = " ") Or (Mid$(sRTF, l + 1, 1) = "\")) Then
                        sCurLine = sCurLine & "<I>"
                        sEOS = "</I>" & sEOS
                        If (Mid$(sRTF, l + 1, 1) = " ") Then l = l + 1
                     Else
                        gSkip = True
                     End If
                  Case "l"
                     If (Mid$(sRTF, l, 10) = "ldblquote ") Then
                        'left doublequote
                        sTmp = "“"
                        l = l + 9
                        bText = True
                     ElseIf (Mid$(sRTF, l, 7) = "lquote ") Then
                        'left quote
                        sTmp = "‘"
                        l = l + 6
                        bText = True
                     Else
                        gSkip = True
                     End If
                  Case "p"
                     If ((Mid$(sRTF, l, 6) = "plain\") Or (Mid$(sRTF, l, 6) = "plain ")) Then
                        If (Len(sFontColor & sFontFace) > 0) Then
                           If Not bSeekinbText Then sCurLine = sCurLine & "</FONT>"
                           sFontColor = ""
                           sFontFace = ""
                        End If
                        If gAlign Then
                           sCurLine = sCurLine & "</TD></TR></TABLE><BR>"
                           gAlign = False
                        End If
                        sCurLine = sCurLine & sEOS
                        sEOS = ""
                        If Mid$(sRTF, l + 5, 1) = "\" Then l = l + 4 Else l = l + 5    'catch next \ but skip a space
                     ElseIf (Mid$(sRTF, l, 9) = "pnlvlblt\") Then
                        'bulleted list
                        sEOS = ""
                        sBOS = "<UL>"
                        sBOL = "<LI>"
                        sEOL = "</LI>"
                        sEOP = "</UL>"
                        l = l + 7    'catch next \
                     ElseIf (Mid$(sRTF, l, 7) = "pntext\") Then
                        l = InStr(l, sRTF, "}")   'skip to end of braces
                     ElseIf (Mid$(sRTF, l, 6) = "pntxtb") Then
                        l = InStr(l, sRTF, "}")   'skip to end of braces
                     ElseIf (Mid$(sRTF, l, 10) = "pard\plain") Then
                        sCurLine = sCurLine & sEOS & sEOP
                        sEOS = ""
                        sEOP = ""
                        sBOL = ""
                        sEOL = "<BR>"
                        l = l + 3    'catch next \
                     Else
                        gSkip = True
                     End If
                  Case "q"
                     If ((Mid$(sRTF, l, 3) = "qc\") Or (Mid$(sRTF, l, 3) = "qc ")) Then
                        'qc = centered
                        sAlign = "center"
                        'move "cursor" position to next rtf code
                        If (Mid$(sRTF, l + 2, 1) = " ") Then l = l + 2
                        l = l + 1
                     ElseIf ((Mid$(sRTF, l, 3) = "qr\") Or (Mid$(sRTF, l, 3) = "qr ")) Then
                        'qr = right justified
                        sAlign = "right"
                        'move "cursor" position to next rtf code
                        If (Mid$(sRTF, l + 2, 1) = " ") Then l = l + 2
                        l = l + 1
                     Else
                        gSkip = True
                     End If
                  Case "r"
                     If (Mid$(sRTF, l, 7) = "rquote ") Then
                        'reverse quote
                        sTmp = "’"
                        l = l + 6
                        bText = True
                     ElseIf (Mid$(sRTF, l, 10) = "rdblquote ") Then
                        'reverse doublequote
                        sTmp = "”"
                        l = l + 9
                        bText = True
                     Else
                        gSkip = True
                     End If
                  Case "s"
                     'strikethrough
                     If ((Mid$(sRTF, l, 7) = "strike\") Or (Mid$(sRTF, l, 7) = "strike ")) Then
                        sCurLine = sCurLine & "<STRIKE>"
                        sEOS = "</STRIKE>" & sEOS
                        l = l + 6
                     Else
                        gSkip = True
                     End If
                  Case "t"
                     If (Mid$(sRTF, l, 4) = "tab ") Then
                        sTmp = "&#9;"   'tab
                        l = l + 2
                        bText = True
                     Else
                        gSkip = True
                     End If
                  Case "u"
                     'underline
                     If ((Mid$(sRTF, l, 3) = "ul ") Or (Mid$(sRTF, l, 3) = "ul\")) Then
                        sCurLine = sCurLine & "<U>"
                        sEOS = "</U>" & sEOS
                        l = l + 1
                     Else
                        gSkip = True
                     End If
                  Case "'"
                     'special characters
                     sTmp2 = "{" & Mid$(sRTF, l + 1, 2) & "}"
                     lTmp = InStr(sCodes, sTmp2)
                     If lTmp = 0 Then
                        sTmp = Chr$("&H" & Mid$(sTmp2, 2, 2))
                     Else
                        sTmp = Trim$(Mid$(sCodes, lTmp - 8, 8))
                     End If
                     l = l + 1
                     bText = True
                  Case "~"
                     sTmp = " "
                     bText = True
                  Case "{", "}", "\"
                     bText = True
                  Case vbLf, vbCr, vbCrLf    'always use vbCrLf
                     sCurLine = sCurLine & vbCrLf
                  Case Else
                     gSkip = True
               End Select
               If gSkip = True Then
                  'skip everything up until the next space or "\" or "}"
                  While InStr(" \}", Mid$(sRTF, l, 1)) = 0
                     l = l + 1
                  Wend
                  gSkip = False
                  If (Mid$(sRTF, l, 1) = "\") Then l = l - 1
               End If
               l = l + 1
            Case vbLf, vbCr, vbCrLf
               l = l + 1
            Case "<"
               sCurLine = sCurLine & "&lt;"
               sTmp = ""
               bText = True
               'l = l + 1
            Case ">"
               sCurLine = sCurLine & "&gt;"
               sTmp = ""
               bText = True
               'l = l + 1
            Case "&"
               sCurLine = sCurLine & "&amp;"
               sTmp = ""
               bText = True
               'l = l + 3
            Case """"
               sCurLine = sCurLine & "&quot;"
               sTmp = ""
               bText = True
               '                l = l + 1
            Case Else
               bText = True
         End Select
         If bText Then
            If ((Len(sFontColor & sFontFace) > 0) And bSeekinbText) Then
               If Len(sAlign) > 0 Then
                  gAlign = True
                  If sAlign = "center" Then
                     sCurLine = sCurLine & "<TABLE ALIGN=""left"" CELLSPACING=0 CELLPADDING=0 WIDTH=""100%""><TR ALIGN=""center""><TD>"
                  ElseIf sAlign = "right" Then
                     sCurLine = sCurLine & "<TABLE ALIGN=""left"" CELLSPACING=0 CELLPADDING=0 WIDTH=""100%""><TR ALIGN=""right""><TD>"
                  End If
                  sAlign = ""
               End If
               If Len(sFontFace) > 0 Then
                  sFontCodes = sFontCodes & " FACE=""" & sFontFace & """"
               End If
               If Len(sFontColor) > 0 Then
                  sFontCodes = sFontCodes & " COLOR=""" & sFontColor & """"
               End If
               If Len(sCurFontSize) > 0 Then
                  sFontCodes = sFontCodes & " SIZE = """ & lFontSize & """"
               End If
               sCurLine = sCurLine & "<FONT" & sFontCodes & ">"
               sFontCodes = ""
            End If
            sCurLine = sCurLine & sTmp
            l = l + 1
            bSeekinbText = False
            bText = False
         End If
      Wend

      lBOS = lEOS + 2
      lEOS = InStr(lEOS + 1, sRTF, vbCrLf & "\par")
      sHTML = sHTML & sEOLL & sBOS & sBOL & sCurLine
      sEOLL = sEOL
      If Len(sEOL) = 0 Then sEOL = "<BR>"

      If lEOS = 0 Then GoTo finally
      sBOS = ""
      sCurLine = ""
   Wend

finally:
   sHTML = sHTML & sEOS

   'clear up any hanging fonts
   If (Len(sFontColor & sFontFace) > 0) Then sHTML = sHTML & "</FONT>" & "</PRE>"

   'Add Generator Metatag if requested
   If InStr(sOptions, "+G") <> 0 Then
      sGen = "<META NAME=""GENERATOR"" CONTENT=""RTF2HTML by VBCodeDatabase"">"
   Else
      sGen = ""
   End If

   'Add Title if requested
   If InStr(sOptions, "+T") <> 0 Then
      lTmp = InStr(sOptions, "+T") + 4
      lTmp2 = InStr(lTmp + 1, sOptions, """")
      On Error Resume Next
      sTitle = Mid$(sOptions, lTmp, lTmp2 - lTmp)
   Else
      sTitle = ""
   End If

   'add header and footer if requested
   If InStr(sOptions, "+H") <> 0 Then sHTML = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">" & vbCrLf _
      & "<HTML>" & vbCrLf _
      & "<HEAD>" & vbCrLf _
      & "<TITLE>" & sTitle & "</TITLE>" & vbCrLf _
      & sGen & vbCrLf _
      & "</HEAD>" & vbCrLf _
      & "<BODY>" & vbCrLf _
      & sHTML _
      & "</BODY>" & vbCrLf _
      & "</HTML>"
   oldRTF2HTML = sHTML

End Function

Public Sub SelectAll()
   ' #VBIDEUtils#************************************************************
   ' * Module Filename  :
   ' * Module Name      : frmSample
   ' * Procedure Name   : SelectText
   ' * Parameters       :
   ' * Date             : 14/12/98
   ' * E-Mail           : acidbuzz@videotron.ca
   ' * Programmer Name  : Michel Gratton
   ' **********************************************************************
   ' * Comments         : This sub is use to highlight text in a control.
   ' **********************************************************************

   On Error Resume Next

   Screen.ActiveControl.SelStart = 0
   Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)

End Sub

Public Function PadR(sTmp As String, lNbr As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/12/1999
   ' * Time             : 11:54
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : PadR
   ' * Parameters       :
   ' *                    sTmp As String
   ' *                    lNbr As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   PadR = left$(RTrim$(sTmp) & Space$(lNbr), lNbr)

End Function

Public Function PadL(sTmp As String, lNbr As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/12/1999
   ' * Time             : 11:55
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : PadL
   ' * Parameters       :
   ' *                    sTmp As String
   ' *                    lNbr As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   PadL = right$(Space$(lNbr) & RTrim$(sTmp), lNbr)

End Function

Public Sub GetAllSettings(sAppName As String, sSection As String, ByRef sValues() As String)

   Dim nCount           As Long

   Dim clsRegistry      As New class_Registry
   With clsRegistry
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & sAppName & "\" & sSection
      Call .EnumerateValues(sValues, nCount)
   End With

End Sub

Public Function GetSetting(sAppName As String, sSection As String, sKey As String, sDefault As String) As String

   Dim sTmp             As String

   Dim clsRegistry      As New class_Registry
   With clsRegistry
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & sAppName & "\" & sSection
      .ValueKey = sKey
      .ValueType = REG_SZ
      sTmp = .Value
   End With
   If sTmp = "" Then
      SaveSetting sAppName, sSection, sKey, sDefault
      sTmp = sDefault
   End If
   GetSetting = sTmp

End Function

Public Sub SaveSetting(sAppName As String, sSection As String, sKey As String, vValue As Variant)

   Dim clsRegistry      As New class_Registry
   With clsRegistry
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & sAppName & "\" & sSection
      .ValueKey = sKey
      .ValueType = REG_SZ
      .Value = vValue
   End With

End Sub

Public Sub DeleteSection(sAppName As String, sSection As String)

   Dim clsRegistry      As New class_Registry
   With clsRegistry
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & sAppName & "\" & sSection
      .DeleteKey
   End With

End Sub

Public Sub DeleteKeyValue(sAppName As String, sSection As String, sKey As String)

   Dim clsRegistry      As New class_Registry
   With clsRegistry
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\" & sAppName & "\" & sSection
      .ValueKey = sKey
      .DeleteValue
   End With

End Sub

Public Function ReplaceAccent(sString As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 01/06/2001
   ' * Time             : 20:07
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : ReplaceAccent
   ' * Parameters       :
   ' *                    sString As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim i                As Integer

   sString = Replace(sString, Chr$(10), vbCrLf)

   ' Replace A & a
   For i = 192 To 197
      sString = Replace(sString, Chr$(i), "A")
   Next
   For i = 224 To 229
      sString = Replace(sString, Chr$(i), "a")
   Next
   ' Replace C & c
   sString = Replace(sString, Chr$(199), "C")
   sString = Replace(sString, Chr$(231), "c")
   ' Replace E & e
   For i = 200 To 203
      sString = Replace(sString, Chr$(i), "E")
   Next
   For i = 232 To 235
      sString = Replace(sString, Chr$(i), "e")
   Next
   ' Replace I & i
   For i = 204 To 207
      sString = Replace(sString, Chr$(i), "I")
   Next
   For i = 236 To 239
      sString = Replace(sString, Chr$(i), "i")
   Next
   ' Replace D
   sString = Replace(sString, Chr$(208), "D")
   ' Replace N & n
   sString = Replace(sString, Chr$(209), "N")
   sString = Replace(sString, Chr$(241), "n")
   ' Replace O & o
   For i = 210 To 214
      sString = Replace(sString, Chr$(i), "O")
   Next
   For i = 242 To 246
      sString = Replace(sString, Chr$(i), "o")
   Next
   ' Replace U & u
   For i = 217 To 220
      sString = Replace(sString, Chr$(i), "U")
   Next
   For i = 249 To 252
      sString = Replace(sString, Chr$(i), "u")
   Next
   ' Replace Y & y
   sString = Replace(sString, Chr$(221), "Y")
   sString = Replace(sString, Chr$(253), "y")
   sString = Replace(sString, Chr$(255), "y")
   sString = Replace(sString, "æ", "ae")
   sString = Replace(sString, "´", "'")
   sString = Replace(sString, "’", "'")

   For i = 123 To 254
      sString = Replace(sString, Chr(i), " ")
   Next

   ReplaceAccent = sString

End Function

Public Function GetWrap(sInput As String, nLen As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 01/13/2001
   ' * Time             : 19:52
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetWrap
   ' * Parameters       :
   ' *                    sInput As String
   ' *                    nLen As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetWrap

   Dim nI               As Long
   Dim sTmp             As String
   Dim sChar            As String

   If Len(sInput) <= nLen + 15 Then
      sTmp = sInput
      sInput = ""
   Else
      For nI = 1 To Len(sInput)
         sChar = Mid$(sInput, nI, 1)
         If (sChar = " ") Then
            If nI > nLen Then
               sTmp = Mid$(sInput, 1, nI)
               sInput = Trim$(Mid$(sInput, nI))
               Exit For
            End If
         End If
      Next
      If sTmp = "" Then
         sTmp = sInput
         sInput = ""
      End If

   End If

   GetWrap = sTmp

EXIT_GetWrap:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_GetWrap:
   Resume EXIT_GetWrap

End Function

Public Function GetStringBetweenTags(ByVal sSearchIn As String, ByVal sFrom As String, ByVal sUntil As String, Optional nPosAfter As Long, Optional ByVal nStartAtPos As Long = 0) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 01/15/2001
   ' * Time             : 13:31
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetStringBetweenTags
   ' * Parameters       :
   ' *                    ByVal sSearchIn As String
   ' *                    ByVal sFrom As String
   ' *                    ByVal sUntil As String
   ' *                    Optional ByVal nStartAtPos As Long = 0
   ' **********************************************************************
   ' * Comments         :
   ' * This function gets in a string and two keywords
   ' * and returns the string between the keywords
   ' *
   ' **********************************************************************

   Dim s1               As Long
   Dim s2               As Long
   Dim s                As Long
   Dim l                As Long
   Dim sFound           As String

   On Error GoTo ERROR_GetStringBetweenTags

   s1 = InStr(nStartAtPos + 1, sSearchIn, sFrom, vbTextCompare)
   s2 = InStr(s1 + 1, sSearchIn, sUntil, vbTextCompare)

   If s1 = 0 Or s2 = 0 Or IsNull(s1) Or IsNull(s2) Then
      sFound = ""
   Else
      s = s1 + Len(sFrom)
      l = s2 - s
      sFound = Mid$(sSearchIn, s, l)
   End If

   GetStringBetweenTags = sFound
   If s + l > 0 Then
      nPosAfter = (s + l) - 1
   End If

   Exit Function

ERROR_GetStringBetweenTags:
   GetStringBetweenTags = ""

End Function

Public Function StringStripper(sInput As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 01/18/2001
   ' * Time             : 12:21
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : StringStripper
   ' * Parameters       :
   ' *                    sInput As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sCopy            As String

   Dim sResult          As String
   Dim nPos             As Integer
   Dim nI               As Integer

   sCopy = sInput

   For nI = 0 To 255
      If nI = 32 Then nI = 33         ' <space>
      If nI = 48 Then nI = 58         ' 0 - 9
      If nI = 65 Then nI = 91         ' A - Z
      If nI = 97 Then nI = 123        ' a - z

      Do While InStr(1, sCopy, Chr$(nI)) > 0
         nPos = InStr(1, sCopy, Chr$(nI))

         If nPos > 0 Then
            sResult = Mid$(sCopy, 1, nPos - 1)
            sCopy = sResult & Mid$(sCopy, nPos + 1)
         End If
      Loop

   Next nI

   StringStripper = sCopy

End Function

Public Function IsLoaded(sFormName As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           :
   ' * Date             : 12/10/99
   ' * Time             : 12:06
   ' * Procedure Name   : IsLoaded
   ' * Parameters       :
   ' *                    sFormName As String
   ' **********************************************************************
   ' * Comments         : his function returns true if a form is loaded
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer

   sFormName = UCase$(sFormName)
   IsLoaded = False
   For nI = 0 To (Forms.Count - 1)
      If UCase$(Forms(nI).Name) = sFormName Then
         IsLoaded = True
         Exit For
      End If
   Next

End Function

Public Function GetDateFormatString(cSeparator As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/15/2001
   ' * Time             : 20:24
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : GetDateFormatString
   ' * Parameters       :
   ' *                    cSeparator As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sLocaleDate      As String
   Dim sWorkDate        As String
   Dim sFormat          As String
   Dim sTemp            As String
   Dim iLoc             As Integer
   Dim iLeft            As Integer
   Dim iRight           As Integer
   Dim iMid             As Integer
   Dim iDD              As Integer
   Dim iMM              As Integer
   Dim iYY              As Integer
   Dim iCC              As Integer
   Dim iYYYY            As Integer
   Dim sYYYY            As String

   Const sKnownDate = "11/07/71"

   '
   ' Convert a known date to the regional format
   '
   sLocaleDate = Format$(sKnownDate, "Short Date")
   sWorkDate = sLocaleDate

   '
   ' Determine the separator character
   '
   For iLoc = 1 To 5
      cSeparator = Mid$(sLocaleDate, iLoc, 1)
      If Not IsNumeric(cSeparator) Then Exit For
   Next

   '
   ' Parse the date into its components
   '
   iLoc = InStr(1, sWorkDate, cSeparator)
   iLeft = Val(left$(sLocaleDate, iLoc - 1))
   sWorkDate = Mid$(sWorkDate, iLoc + 1)
   iLoc = InStr(1, sWorkDate, cSeparator)

   '
   'Correct regional date format 4/98
   '
   sTemp = Mid$(sWorkDate, 1, iLoc - 1)
   If IsNumeric(sTemp) Then
      iMid = Val(Mid$(sWorkDate, 1, iLoc - 1))
   Else
      iMid = Month(sLocaleDate)
   End If

   sWorkDate = Mid$(sWorkDate, iLoc + 1)
   iRight = Val(Mid$(sWorkDate, 1))

   '
   ' Locale aware functions
   '
   iDD = Day(sLocaleDate)
   iMM = Month(sLocaleDate)
   iYYYY = Year(sLocaleDate)
   sYYYY = CStr(iYYYY)
   iCC = Val(left$(sYYYY, 2))
   iYY = Val(right(sYYYY, 2))

   '
   ' Is the left component the day, month or year??
   '
   Select Case iLeft
      Case iDD
         sFormat = "dd/"
      Case iMM
         sFormat = "mm/"
      Case iYY, iYYYY
         sFormat = "yyyy/"
   End Select

   '
   ' How about the middle?
   '
   Select Case iMid
      Case iDD
         sFormat = sFormat & "dd/"
      Case iMM
         sFormat = sFormat & "mm/"
      Case iYY, iYYYY
         sFormat = sFormat & "yyyy/"
   End Select

   '
   ' And the right component is:
   '
   Select Case iRight
      Case iDD
         sFormat = sFormat & "dd"
      Case iMM
         sFormat = sFormat & "mm"
      Case iYY, iYYYY
         sFormat = sFormat & "yyyy"
   End Select
   GetDateFormatString = sFormat

End Function

Public Function URLEncode(StringToEncode As String, Optional UsePlusRatherThanHexForSpace As Boolean = False) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/20/2001
   ' * Time             : 21:02
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : URLEncode
   ' * Parameters       :
   ' *                    StringToEncode As String
   ' *                    Optional UsePlusRatherThanHexForSpace As Boolean = False
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim TempAns          As String
   Dim CurChr           As Integer
   CurChr = 1
   Do Until CurChr - 1 = Len(StringToEncode)
      Select Case Asc(Mid$(StringToEncode, CurChr, 1))
         Case 48 To 57, 65 To 90, 97 To 122
            TempAns = TempAns & Mid$(StringToEncode, CurChr, 1)
         Case 32
            If UsePlusRatherThanHexForSpace = True Then
               TempAns = TempAns & "+"
            Else
               TempAns = TempAns & "%" & Hex(32)
            End If
         Case Else
            TempAns = TempAns & "%" & Hex(Asc(Mid$(StringToEncode, CurChr, 1)))
      End Select

      CurChr = CurChr + 1
   Loop

   URLEncode = TempAns
End Function

Public Function UrlDecode(StringToDecode As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/20/2001
   ' * Time             : 21:02
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : URLDecode
   ' * Parameters       :
   ' *                    StringToDecode As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim TempAns          As String
   Dim CurChr           As Integer

   CurChr = 1

   Do Until CurChr - 1 = Len(StringToDecode)
      Select Case Mid$(StringToDecode, CurChr, 1)
         Case "+"
            TempAns = TempAns & " "
         Case "%"
            TempAns = TempAns & Chr$(Val("&h" & _
               Mid$(StringToDecode, CurChr + 1, 2)))
            CurChr = CurChr + 2
         Case Else
            TempAns = TempAns & Mid$(StringToDecode, CurChr, 1)
      End Select

      CurChr = CurChr + 1
   Loop

   UrlDecode = TempAns
End Function

'Public Sub CheckLicense()
'   ' #VBIDEUtils#************************************************************
'   ' * Programmer Name  : removed
'   ' * Web Site         : http://www.ppreview.net
'   ' * E-Mail           : removed
'   ' * Date             : 03/22/2001
'   ' * Time             : 10:04
'   ' * Module Name      : Main_Module
'   ' * Module Filename  : Main.bas
'   ' * Procedure Name   : CheckLicense
'   ' * Parameters       :
'   ' **********************************************************************
'   ' * Comments         :
'   ' *
'   ' *
'   ' **********************************************************************
'
'   Exit Sub
'
'   Dim oLicense    As New class_License
'   Dim SName       As String
'
'   With oLicense
'      .licCLSID = LIC_CLSID
'      .licKEY = LIC_KEY
'
'      .Register
'
'      'check the registery for a valid license key
'      Call .CheckRegistration
'   End With
'
'   Set oLicense = Nothing
'
'End Sub

Public Function InDesignMode() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/25/2001
   ' * Time             : 15:01
   ' * Module Name      : Main_Module
   ' * Module Filename  : Main.bas
   ' * Procedure Name   : InDebugMode
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next
   Debug.Assert 0 / 0
   InDesignMode = err.number <> 0

End Function

