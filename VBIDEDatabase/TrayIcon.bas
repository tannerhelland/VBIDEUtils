Attribute VB_Name = "TrayIcon_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:53
' * Module Name      : TrayIcon_Module
' * Module Filename  : TrayIcon.bas
' **********************************************************************
' * Comments         : Module containing all the code to manage
' * the application in the TrayIcon Taskbar
' *
' **********************************************************************

Option Explicit

' Public declarations required for frmTrayIcon
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
