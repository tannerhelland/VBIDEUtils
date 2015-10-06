Attribute VB_Name = "MsgBox_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 09/14/2000
' * Time             : 15:44
' * Module Name      : MsgBox_Module
' * Module Filename  : MsgBox.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private Const NV_CLOSEMSGBOX = &H5000&
Private Const NV_MOVEMSGBOX = &H5001&
Private Const NV_TOPMSGBOX = &H5002&
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1

Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, _
   ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
   ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
   ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, _
   ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
   ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private mTitle          As String
Private mX              As Long
Private mY              As Long
Private mPause          As Long
Private mHandle         As Long

Public Function MsgBoxTop(ByVal hWnd As Long, ByVal sPrompt As String, Optional nButtons As Long = vbInformation + vbOKOnly, Optional sTitle As String = "") As Integer
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/08/2001
   ' * Time             : 23:42
   ' * Module Name      : MsgBox_Module
   ' * Module Filename  : MsgBox.bas
   ' * Procedure Name   : MsgBoxTop
   ' * Parameters       :
   ' *                    ByVal hwnd As Long
   ' *                    ByVal sPrompt As String
   ' *                    Optional nButtons As Long = vbInformation + vbOKOnly
   ' *                    Optional sTitle As String = ""
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   mTitle = sTitle
   SetTimer hWnd, NV_TOPMSGBOX, 0&, AddressOf NewTimerProc
   MsgBoxTop = MessageBox(hWnd, sPrompt, sTitle, nButtons)

End Function

Public Function MsgBoxMove(ByVal hWnd As Long, ByVal sPrompt As String, ByVal sTitle As String, ByVal nButtons As Long, ByVal inX As Long, ByVal inY As Long) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/14/2000
   ' * Time             : 15:44
   ' * Module Name      : MsgBox_Module
   ' * Module Filename  : MsgBox.bas
   ' * Procedure Name   : MsgBoxMove
   ' * Parameters       :
   ' *                    ByVal hwnd As Long
   ' *                    ByVal sPrompt As String
   ' *                    ByVal sTitle As String
   ' *                    ByVal nButtons As Long
   ' *                    ByVal inX As Long
   ' *                    ByVal inY As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   mTitle = sTitle: mX = inX:  mY = inY
   SetTimer hWnd, NV_MOVEMSGBOX, 0&, AddressOf NewTimerProc
   MsgBoxMove = MessageBox(hWnd, sPrompt, sTitle, nButtons)
End Function

Public Function MsgBoxPause(ByVal hWnd As Long, ByVal sPrompt As String, ByVal sTitle As String, ByVal nButtons As Long, ByVal inPause As Integer) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/14/2000
   ' * Time             : 15:44
   ' * Module Name      : MsgBox_Module
   ' * Module Filename  : MsgBox.bas
   ' * Procedure Name   : MsgBoxPause
   ' * Parameters       :
   ' *                    ByVal hwnd As Long
   ' *                    ByVal sPrompt As String
   ' *                    ByVal sTitle As String
   ' *                    ByVal nButtons As Long
   ' *                    ByVal inPause As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   mTitle = sTitle: mPause = inPause * 1000
   SetTimer hWnd, NV_CLOSEMSGBOX, mPause, AddressOf NewTimerProc
   MsgBoxPause = MessageBox(hWnd, sPrompt, sTitle, nButtons)
End Function

Public Function NewTimerProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/14/2000
   ' * Time             : 15:44
   ' * Module Name      : MsgBox_Module
   ' * Module Filename  : MsgBox.bas
   ' * Procedure Name   : NewTimerProc
   ' * Parameters       :
   ' *                    ByVal hwnd As Long
   ' *                    ByVal Msg As Long
   ' *                    ByVal wparam As Long
   ' *                    ByVal lparam As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim w                As Single
   Dim h                As Single
   Dim mBox             As RECT

   KillTimer hWnd, wParam
   Select Case wParam
      Case NV_CLOSEMSGBOX:
         ' A system class is a window class registered by the system which cannot
         ' be destroyed by a processed, e.g. #32768 (a menu), #32769 (desktop
         ' window), #32770 (dialog box), #32771 (task switch window).
         mHandle = FindWindow("#32770", mTitle)
         If mHandle <> 0 Then
            SetForegroundWindow mHandle
            SendKeys "{enter}"
         End If

      Case NV_TOPMSGBOX:
         mHandle = FindWindow("#32770", mTitle)
         If mHandle <> 0 Then
            w = Screen.Width / Screen.TwipsPerPixelX
            h = Screen.Height / Screen.TwipsPerPixelY
            GetWindowRect mHandle, mBox
            mX = (w - (mBox.right - mBox.left) - 1) / 2
            mY = (h - (mBox.bottom - mBox.top) - 1) / 2
            ' SWP_NOSIZE is to use current size, ignoring 3rd & 4th parameters.
            SetWindowPos mHandle, HWND_TOPMOST, mX, mY, 0, 0, SWP_NOSIZE
         End If

      Case NV_MOVEMSGBOX:
         mHandle = FindWindow("#32770", mTitle)
         If mHandle <> 0 Then
            w = Screen.Width / Screen.TwipsPerPixelX
            h = Screen.Height / Screen.TwipsPerPixelY
            GetWindowRect mHandle, mBox
            If mX > (w - (mBox.right - mBox.left) - 1) Then mX = (w - (mBox.right - mBox.left) - 1)
            If mY > (h - (mBox.bottom - mBox.top) - 1) Then mY = (h - (mBox.bottom - mBox.top) - 1)
            If mX < 1 Then mX = 1: If mY < 1 Then mY = 1
            ' SWP_NOSIZE is to use current size, ignoring 3rd & 4th parameters.
            SetWindowPos mHandle, HWND_TOPMOST, mX, mY, 0, 0, SWP_NOSIZE
         End If
   End Select
End Function

