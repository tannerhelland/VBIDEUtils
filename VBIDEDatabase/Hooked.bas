Attribute VB_Name = "Hooked_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:52
' * Module Name      : Hooked_Module
' * Module Filename  : Hooked.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

' ==========================================================================
' Module:   mDialogHook
' Filename: mDialogHook.bas
' Author:   Steve McMahon
' Date:     24 May 1998
'
' Provides functions which can be called via AddressOf for common
' dialog hook support.
' ==========================================================================
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private m_cHookedDialog As Long

Property Let HookedDialog(ByRef cThis As class_GCommonDialog)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : HookedDialog
   ' * Parameters       :
   ' *                    ByRef cThis As class_GCommonDialog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Set cHookedDialog = cThis
   m_cHookedDialog = ObjPtr(cThis)
End Property

Property Get HookedDialog() As class_GCommonDialog
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : HookedDialog
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim oT               As class_GCommonDialog
   If (m_cHookedDialog <> 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory oT, m_cHookedDialog, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set HookedDialog = oT
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory oT, 0&, 4
   End If
End Property

Public Sub ClearHookedDialog()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : ClearHookedDialog
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_cHookedDialog = 0
End Sub

Public Function DialogHookFunction(ByVal hDlg As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : DialogHookFunction
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' *                    ByVal Msg As Long
   ' *                    ByVal wParam As Long
   ' *                    ByVal lParam As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cD               As class_GCommonDialog
   Set cD = HookedDialog
   If Not (cD Is Nothing) Then
      DialogHookFunction = cD.DialogHook(hDlg, Msg, wParam, lParam)
   End If
End Function

Public Function PrintHookProc(ByVal hDlg As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : PrintHookProc
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' *                    ByVal Msg As Long
   ' *                    ByVal wParam As Long
   ' *                    ByVal lParam As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cD               As class_GCommonDialog
   Set cD = HookedDialog
   If Not (cD Is Nothing) Then
      PrintHookProc = cD.DialogHook(hDlg, Msg, wParam, lParam)
   End If
End Function

Public Function PrintSetupHookProc(ByVal hDlg As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : PrintSetupHookProc
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' *                    ByVal Msg As Long
   ' *                    ByVal wParam As Long
   ' *                    ByVal lParam As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cD               As class_GCommonDialog
   Set cD = HookedDialog
   If Not (cD Is Nothing) Then
      PrintSetupHookProc = cD.DialogHook(hDlg, Msg, wParam, lParam)
   End If
End Function

Public Function PageSetupHook(ByVal hDlg As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : PageSetupHook
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' *                    ByVal Msg As Long
   ' *                    ByVal wParam As Long
   ' *                    ByVal lParam As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cD               As class_GCommonDialog
   Set cD = HookedDialog
   If Not (cD Is Nothing) Then
      PageSetupHook = cD.DialogHook(hDlg, Msg, wParam, lParam)
   End If
End Function

Public Function CCHookProc(ByVal hDlg As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : CCHookProc
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' *                    ByVal Msg As Long
   ' *                    ByVal wParam As Long
   ' *                    ByVal lParam As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cD               As class_GCommonDialog
   Set cD = HookedDialog
   If Not (cD Is Nothing) Then
      CCHookProc = cD.DialogHook(hDlg, Msg, wParam, lParam)
   End If

End Function

Public Function CFHookProc(ByVal hDlg As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Hooked_Module
   ' * Module Filename  : Hooked.bas
   ' * Procedure Name   : CFHookProc
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' *                    ByVal Msg As Long
   ' *                    ByVal wParam As Long
   ' *                    ByVal lParam As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cD               As class_GCommonDialog
   Set cD = HookedDialog
   If Not (cD Is Nothing) Then
      CFHookProc = cD.DialogHook(hDlg, Msg, wParam, lParam)
   End If
End Function
