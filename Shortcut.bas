Attribute VB_Name = "Shortcut_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 23/08/99
' * Time             : 13:33
' * Module Name      : Shortcut_Module
' * Module Filename  : Shortcut.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'
'Public Sub ShortCutTreatment()
'   ' #VBIDEUtils#************************************************************
'   ' * Programmer Name  : removed
'   ' * Web Site         : http://www.ppreview.net
'   ' * E-Mail           : removed
'   ' * Date             : 23/08/99
'   ' * Time             : 13:33
'   ' * Module Name      : Shortcut_Module
'   ' * Module Filename  : Shortcut.bas
'   ' * Procedure Name   : ShortCutTreatment
'   ' * Parameters       :
'   ' **********************************************************************
'   ' * Comments         :
'   ' *
'   ' *
'   ' **********************************************************************
'
'   Exit Sub
'
'   If (GetAsyncKeyState(vbKeyControl)) Then
'      If (GetAsyncKeyState(vbKeyShift)) Then
'         If GetAsyncKeyState(Asc("H")) Then Call InsertProcedureHeader
'         If GetAsyncKeyState(Asc("M")) Then Call InsertModuleHeader
'         If GetAsyncKeyState(Asc("W")) Then
'            ' *** Search the web
'            Call AlwaysOnTop(frmFindWeb, True)
'            frmFindWeb.Show
'         End If
'
'         If GetAsyncKeyState(Asc("P")) Then
'            ' *** Project explorer
'         End If
'      End If
'
'      If GetAsyncKeyState(Asc("U")) Then Call GetPending
'      If GetAsyncKeyState(Asc("K")) Then Call ClearDebug
'      If GetAsyncKeyState(Asc("W")) Then Call SwapEgual
'      If GetAsyncKeyState(Asc("M")) Then frmMessageBox.Show
'
'   End If
'
'End Sub
