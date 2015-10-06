VERSION 5.00
Begin VB.Form frmTrayIcon 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1620
   ClientLeft      =   4155
   ClientTop       =   4590
   ClientWidth     =   2700
   Icon            =   "TrayIcon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1620
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:51
' * Module Name      : frmTrayIcon
' * Module Filename  : TrayIcon.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private Type NOTIFYICONDATA
   cbSize               As Long
   hWnd                 As Long
   uId                  As Long
   uFlags               As Long
   ucallbackMessage     As Long
   hIcon                As Long
   szTip                As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private TrayIcon        As NOTIFYICONDATA

Private IconIsOn        As Boolean

Public Owner            As Form
Public ShowTray         As Boolean
Public sToolTips        As String

Public Sub SetCallback(ob As Object)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmTrayIcon
   ' * Module Filename  : TrayIcon.frm
   ' * Procedure Name   : SetCallback
   ' * Parameters       :
   ' *                    ob As Object
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Set's which object to call "TrayIconCallback(msg As Long)" in.

   Set Owner = ob

End Sub

Public Sub Update(IconOn As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmTrayIcon
   ' * Module Filename  : TrayIcon.frm
   ' * Procedure Name   : Update
   ' * Parameters       :
   ' *                    IconOn As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Turns the Trayicon on and off.
   ' *** Note: the Icon, and ToolTip are taken from the
   ' ***   form's icon and caption respectively.
   If IconOn Then
      TrayIcon.cbSize = Len(TrayIcon)
      TrayIcon.hWnd = Me.hWnd
      TrayIcon.uId = 1&
      TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      If sToolTips <> "" Then
         TrayIcon.uFlags = TrayIcon.uFlags Or NIF_TIP
      End If
      TrayIcon.ucallbackMessage = WM_MOUSEMOVE
      TrayIcon.hIcon = Me.Icon
      TrayIcon.szTip = sToolTips & Chr$(0)

      If Not IconIsOn Then
         Shell_NotifyIcon NIM_ADD, TrayIcon
      Else
         Shell_NotifyIcon NIM_MODIFY, TrayIcon
      End If
      IconIsOn = True
   Else
      If IconIsOn Then
         TrayIcon.cbSize = Len(TrayIcon)
         TrayIcon.hWnd = Me.hWnd
         TrayIcon.uId = 1&
         Shell_NotifyIcon NIM_DELETE, TrayIcon
      End If
      IconIsOn = False
   End If

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmTrayIcon
   ' * Module Filename  : TrayIcon.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   IconIsOn = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmTrayIcon
   ' * Module Filename  : TrayIcon.frm
   ' * Procedure Name   : Form_MouseMove
   ' * Parameters       :
   ' *                    Button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Static rec As Boolean, Msg As Long

   Msg = x / Screen.TwipsPerPixelX
   If rec = False Then
      rec = True

      ' Ignore an error if the call to
      ' Owner.TrayIconCallback fails
      On Error Resume Next
      If Msg >= WM_LBUTTONDOWN And Msg <= WM_RBUTTONDBLCLK Then
         Call Owner.TrayIconCallback(Msg)
      End If

      rec = False
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmTrayIcon
   ' * Module Filename  : TrayIcon.frm
   ' * Procedure Name   : Form_Unload
   ' * Parameters       :
   ' *                    Cancel As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Me.Update False

End Sub
