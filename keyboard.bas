Attribute VB_Name = "Keyboard_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 20/03/2000
' * Time             : 15:57
' * Module Name      : Keyboard_Module
' * Module Filename  : keyboard.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************
Option Explicit

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadID As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const WH_KEYBOARD = 2
Private Const KBH_MASK = &H20000000
Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11

Private hHook           As Long

Public Sub SetKeyHook()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 20/03/2000
   ' * Time             : 15:58
   ' * Module Name      : Keyboard_Module
   ' * Module Filename  : keyboard.bas
   ' * Procedure Name   : SetKeyHook
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Exit Sub

   On Error Resume Next
   hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, 0&, App.ThreadID)

End Sub

Public Sub UnsetKeyHook()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 20/03/2000
   ' * Time             : 15:57
   ' * Module Name      : Keyboard_Module
   ' * Module Filename  : keyboard.bas
   ' * Procedure Name   : UnsetKeyHook
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Exit Sub

   On Error Resume Next
   If Not hHook = 0 Then Call UnhookWindowsHookEx(hHook)

End Sub

Private Function KeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 20/03/2000
   ' * Time             : 15:56
   ' * Module Name      : Keyboard_Module
   ' * Module Filename  : keyboard.bas
   ' * Procedure Name   : KeyboardProc
   ' * Parameters       :
   ' *                    ByVal nCode As Long
   ' *                    ByVal wParam As Long
   ' *                    ByVal lParam As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Exit Function

   If (nCode > 0) And (lParam And KBH_MASK) = 0 Then
      Debug.Print "Alt " & time
      If GetKeyState(VK_CONTROL) <= 1 Then
         If GetKeyState(VK_SHIFT) Then
            If wParam = Asc("H") Then
               Call InsertProcedureHeader

               KeyboardProc = 1
               Exit Function
            End If
            If wParam = Asc("M") Then
               Call InsertModuleHeader

               KeyboardProc = 1
               Exit Function
            End If
            If wParam = Asc("W") Then
               ' *** Search the web
               Call AlwaysOnTop(frmFindWeb, True)
               frmFindWeb.Show

               KeyboardProc = 1
               Exit Function
            End If

            If wParam = Asc("P") Then
               ' *** Project explorer

               KeyboardProc = 1
               Exit Function
            End If
         End If

         If wParam = Asc("U") Then
            Call GetPending

            KeyboardProc = 1
            Exit Function
         End If
         If wParam = Asc("K") Then
            Call ClearDebug
            nCode = 0
            lParam = 0
            wParam = 0
            KeyboardProc = 1
            Exit Function
         End If
         If wParam = Asc("W") Then
            Call SwapEgual

            KeyboardProc = 1
            Exit Function
         End If
         If wParam = Asc("M") Then
            frmMessageBox.Show

            KeyboardProc = 1
            Exit Function
         End If

      End If
   End If

   KeyboardProc = CallNextHookEx(hHook, nCode, wParam, lParam)

End Function
