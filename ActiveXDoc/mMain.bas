Attribute VB_Name = "mMain"
Option Explicit

Public Const WM_QUIT = &H12
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub Main()
Dim sCmd As String
    sCmd = Command
    frmDocHelp.CommandLine = sCmd
    frmDocHelp.Show
End Sub

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
   InDesignMode = Err.Number <> 0

End Function

