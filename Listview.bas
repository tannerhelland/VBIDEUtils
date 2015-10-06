Attribute VB_Name = "Listview_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 15/10/99
' * Time             : 12:21
' * Module Name      : Listview_Module
' * Module Filename  : Listview.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_CHECKBOXES = &H4
Private Const LVS_EX_HEADERDRAGDROP = &H10
Private Const LVS_EX_TRACKSELECT = &H8

Public Sub SetListviewGridLines(lv As ListView, bGrid As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 12:22
   ' * Module Name      : Listview_Module
   ' * Module Filename  : Listview.bas
   ' * Procedure Name   : SetListviewGridLines
   ' * Parameters       :
   ' *                    lv As ListView
   ' *                    bGrid As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call SendMessageLong(lv.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, bGrid)

End Sub

Public Sub SetListviewFullSelectLine(lv As ListView, bSelect As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 12:22
   ' * Module Name      : Listview_Module
   ' * Module Filename  : Listview.bas
   ' * Procedure Name   : SetListviewFullSelectLine
   ' * Parameters       :
   ' *                    lv As ListView
   ' *                    bSelect As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call SendMessageLong(lv.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT, bSelect)

End Sub
