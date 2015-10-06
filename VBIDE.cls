VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_VBIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 07/04/2001
' * Time             : 13:46
' * Module Name      : class_VBIDE
' * Module Filename  : VBIDE.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Public Sub HookVBIDE()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 07/04/2001
   ' * Time             : 13:46
   ' * Module Name      : class_VBIDE
   ' * Module Filename  : VBIDE.cls
   ' * Procedure Name   : Go
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'load our macros
   On Error GoTo err

   'subclass the vbide

   MDIClientHWND = FindWindowEx(VBInstance.MainWindow.hWnd, 0, "MDIClient", vbNullString)
   oProc = GetWindowLong(MDIClientHWND, GWL_WNDPROC)
   SetWindowLong MDIClientHWND, GWL_WNDPROC, AddressOf WindowProc
   oMainProc = GetWindowLong(VBInstance.MainWindow.hWnd, GWL_WNDPROC)
   SetWindowLong VBInstance.MainWindow.hWnd, GWL_WNDPROC, AddressOf MainWindowProc

   'set our bg pic propps
   SubclassIDE_Module.Width = frmSplash.ScaleWidth
   SubclassIDE_Module.Height = frmSplash.ScaleHeight
   SubclassIDE_Module.dc = GetDC(MDIClientHWND)
   Exit Sub
err:
   'hmm.. got an error?
   err.Clear
   MsgBox "error " & err.Description

End Sub

Public Sub unHookVBIDE()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 07/04/2001
   ' * Time             : 13:46
   ' * Module Name      : class_VBIDE
   ' * Module Filename  : VBIDE.cls
   ' * Procedure Name   : Go
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo err

   ' *** unsubclass the vbide
   Call UnhookWindowsHookEx(VBInstance.MainWindow.hWnd)

   Exit Sub
err:
   'hmm.. got an error?
   err.Clear
   MsgBox "error " & err.Description

End Sub
