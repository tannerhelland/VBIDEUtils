Attribute VB_Name = "Unload_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 16/02/2000
' * Time             : 14:19
' * Module Name      : Unload
' * Module Filename  :
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************

Option Explicit

Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Const IDANI_CAPTION = &H3

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Sub UnloadEffect(frm As Form)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 16/02/2000
   ' * Time             : 14:19
   ' * Module Name      : Unload
   ' * Module Filename  :
   ' * Procedure Name   : UnloadEffect
   ' * Parameters       :
   ' * *******************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * *******************************************************************

   Dim FromRect         As RECT
   Dim ToRect           As RECT

   GetWindowRect frm.hWnd, FromRect
   GetWindowRect frm.hWnd, ToRect
   ToRect.top = FromRect.bottom
   ToRect.bottom = FromRect.bottom
   ToRect.left = FromRect.right
   ToRect.right = FromRect.right

   DrawAnimatedRects frm.hWnd, IDANI_CAPTION, FromRect, ToRect
End Sub
