Attribute VB_Name = "TaskBar_Module"
' #VBIDEUtils#************************************************************
' *Programmer Name  : VBNet
' *Web Site         : http://www.mvps.org/vbnet/code/comctl/tbarie5prob.htm
' *E-Mail           : removed
' *Date             : 07/09/1999
' *Time             : 16:37
' *Module Name      : TaskBar_Module
' *Module Filename  : Taskbar.bas
' **********************************************************************
' *Comments         : Changing a VB Toolbar to a Rebar-Style Toolbar
' * Fixing the IE5/MSComCtl 5 Toolbar Problem
' *
' **********************************************************************

Option Explicit

Private Const WM_USER   As Long = &H400
Private Const TB_SETSTYLE As Long = WM_USER + 56
Private Const TB_GETSTYLE As Long = WM_USER + 57
Private Const TBSTYLE_WRAPABLE As Long = &H200  'buttons to wrap when form resized
Private Const TBSTYLE_FLAT As Long = &H800      'flat IE3+ style toolbar
Private Const TBSTYLE_LIST As Long = &H1000     'places captions beside buttons

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
   (ByVal hWnd1 As Long, _
   ByVal hWnd2 As Long, _
   ByVal lpsz1 As String, _
   ByVal lpsz2 As String) As Long

' *** Taskbar
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREPOSITION = &H200
Private Const SWP_NOSIZE = &H1

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' *** Set the Show In Taskbar property at run time
'
' *** VB provides a ShowInTask bar property for forms which allows you to choose whether
' *** a form is shown in the Alt-Tab sequence and the shell's task bar.
' *** However there are two limitations to this:
' ***  The property can't be set at run-time.
' ***  it doesn 't seem to work at all for modal forms under NT4.0
'
' *** Often it is rather handy to put a modal form in the task bar - say for example a
' *** login dialog which shows before your main form - otherwise the user can 'loose'
' *** this window behind other ones.
'
' *** The rules the taskbar uses to decide whether a button should be shown for a window
' *** aren 't very well documented. Here is how it is done:
' *** When you create a window, the taskbar examines the window’s extended style to see
' *** if either the WS_EX_APPWINDOW (&H40000) or WS_EX_TOOLWINDOW (defined as &H80) style
' *** is turned on. If WS_EX_APPWINDOW is turned on, the taskbar shows a button
' *** for the window, and if WS_EX_ TOOLWINDOW is turned on, the taskbar does not show
' *** a button for the window. A window should never have both of these extended styles.
' *** If the window doesn't have either of these styles, the taskbar decides to create
' *** a button if the window is unowned and does not create a button if the window is owned.
' *** Incidentally, VB forms seem to have neither of these extended styles.
'
' *** Here is some code which allows you to set the WS_EX_APPWINDOW style at run time:

Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_TOOLWINDOW = &H80&
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE = (-20)

Public Function MakeFlatToolbar(TBar As ToolBar, fHorizontal As Boolean, fWrapable As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : VBNet
   ' * Web Site         : http://www.mvps.org/vbnet/code/comctl/tbarie5prob.htm
   ' * E-Mail           : removed
   ' * Date             : 07/09/1999
   ' * Time             : 16:37
   ' * Module Name      : TaskBar_Module
   ' * Module Filename  : taskbar.bas
   ' * Procedure Name   : MakeFlatToolbar
   ' * Parameters       :
   ' *                    TBar As Toolbar
   ' *                    fHorizontal As Boolean
   ' *                    fWrapable As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim hTBar            As Long
   Dim Style            As Long

   'to assure that the toolbar has correctly calculated
   'the button widths based on the assigned captions,
   'force a refresh
   TBar.Refresh

   'get the handle of the toolbar
   hTBar = FindWindowEx(TBar.hWnd, 0&, "ToolbarWindow32", "")

   'retrieve the current toolbar style
   Style = SendMessageLong(hTBar, TB_GETSTYLE, 0&, 0&)

   'Set the new style flags.
   'To assure that the button caption will set the correct
   'button width, the TBSTYLE_WRAPABLE style must be
   'specified before making flat.
   Style = Style Or TBSTYLE_FLAT Or TBSTYLE_WRAPABLE

   'if a horizontal layout was specified, add that style
   If fHorizontal Then Style = Style Or TBSTYLE_LIST

   'apply the new style to the toolbar and refresh
   Call SendMessageLong(hTBar, TB_SETSTYLE, 0&, Style)
   TBar.Refresh

   'now that the toolbar is flat, if the wrapable
   'style is not desired, it can be removed.
   'A refresh is not required.
   If fWrapable = False Then
      Style = Style Xor TBSTYLE_WRAPABLE
      Call SendMessageLong(hTBar, TB_SETSTYLE, 0&, Style)
   End If

End Function

Public Sub MakeHorizontalToolbar(TBar As ToolBar, fHorizontal As Boolean, fWrapable As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : VBNet
   ' * Web Site         : http://www.mvps.org/vbnet/code/comctl/tbarie5prob.htm
   ' * E-Mail           : removed
   ' * Date             : 07/09/1999
   ' * Time             : 16:37
   ' * Module Name      : TaskBar_Module
   ' * Module Filename  : taskbar.bas
   ' * Procedure Name   : MakeHorizontalToolbar
   ' * Parameters       :
   ' *                    TBar As Toolbar
   ' *                    fHorizontal As Boolean
   ' *                    fWrapable As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim hTBar            As Long
   Dim Style            As Long

   'get the handle of the toolbar
   hTBar = FindWindowEx(TBar.hWnd, 0&, "ToolbarWindow32", "")

   'retrieve the current toolbar style
   Style = SendMessageLong(hTBar, TB_GETSTYLE, 0&, 0&)

   'Set the new style flags
   If fHorizontal Then
      Style = Style Or TBSTYLE_LIST
   End If

   If fWrapable Then
      Style = Style Or TBSTYLE_WRAPABLE
   End If

   'apply the new style to the toolbar and refresh
   Call SendMessageLong(hTBar, TB_SETSTYLE, 0&, Style)
   TBar.Refresh

End Sub

Public Sub SetToolbarCaption(TBar As ToolBar, TBIndex As Long, newCaption As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : VBNet
   ' * Web Site         : http://www.mvps.org/vbnet/code/comctl/tbarie5prob.htm
   ' * E-Mail           : removed
   ' * Date             : 07/09/1999
   ' * Time             : 16:37
   ' * Module Name      : TaskBar_Module
   ' * Module Filename  : taskbar.bas
   ' * Procedure Name   : SetToolbarCaption
   ' * Parameters       :
   ' *                    TBar As Toolbar
   ' *                    TBIndex As Integer
   ' *                    newCaption As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim hTBar            As Long
   Dim inStyle          As Long
   Dim tempStyle        As Long

   'get the handle to the toolbar and its current style
   hTBar = FindWindowEx(TBar.hWnd, 0&, "ToolbarWindow32", "")
   inStyle = SendMessageLong(hTBar, TB_GETSTYLE, 0&, 0&)

   'if the toolbar has had the flat style applied
   If inStyle And TBSTYLE_FLAT Then

      'This will all happen too quickly to
      'reshow the old raised style

      'temporarily remove the flat style
      tempStyle = inStyle Xor TBSTYLE_FLAT
      Call SendMessageLong(hTBar, TB_SETSTYLE, 0&, tempStyle)

      'change and refresh the caption. A refresh
      'is required here (before resetting to flat)
      'for the toolbar to recalculate the correct
      'sizes of the toolbar buttons based on the
      'longest item.
      TBar.Buttons(TBIndex).Caption = newCaption
      TBar.Refresh

      'restore the previous style, and refresh once more
      Call SendMessageLong(hTBar, TB_SETSTYLE, 0&, inStyle)
      TBar.Refresh

   Else
      'its not flat, so just change the text
      TBar.Buttons(TBIndex).Caption = newCaption

   End If

End Sub

