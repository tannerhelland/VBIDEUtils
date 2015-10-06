VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Tooltips"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : Thomas Kabir
' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
' * E-Mail           : contact@vbfrood.be
' * Date             : 16/02/2000
' * Time             : 12:58
' * Module Name      : cTooltips
' * Module Filename  : cTooltips.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

#Const WIN32_IE = &H400

Private hWndTT          As Long
Private m_Enabled       As Boolean
Private m_Icon          As InfoTitleConstants
Private m_Title         As String

Public Enum DelayTimeConstants
   dtAutoPop = &H2
   dtInitial = &H3
   dtReshow = &H1
End Enum

Public Enum SetDelayTimeConstants
   sdtAutoPop = &H2
   sdtInitial = &H3
   sdtReshow = &H1
   sdtAutomatic = &H0
End Enum

Public Enum InfoTitleConstants
   itNoIcon = 0
   itInfoIcon = 1
   itWarningIcon = 2
   itErrorIcon = 3
End Enum

Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type

Private Type TOOLINFO
   cbSize               As Long
   uFlags               As Long
   hWnd                 As Long
   uId                  As Long
   WinRect              As RECT
   hinst                As Long
   lpszText             As String
   #If WIN32_IE >= &H300 Then
      lParam               As Long
   #End If
End Type

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Const WS_EX_TOPMOST = &H8&
Private Const TOOLTIPS_CLASS = "tooltips_class32"
Private Const WS_POPUP = &H80000000
Private Const CW_USEDEFAULT = &H80000000

Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As INITCOMMONCONTROLEXSTRUCT) As Long
Private Const ICC_WIN95_CLASSES = &HFF
Private Const ICC_BAR_CLASSES = &H4

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
#If Not UNICODE Then
   Private Const TTM_ADDTOOL = WM_USER + 4
   Private Const TTM_DELTOOL = WM_USER + 5
   Private Const TTM_ENUMTOOLS = WM_USER + 14
   Private Const TTM_HITTEST = WM_USER + 10
   Private Const TTM_NEWTOOLRECT = WM_USER + 6
   Private Const TTM_SETTITLE = WM_USER + 32
   Private Const TTM_UPDATETIPTEXT = WM_USER + 12
#Else
   Private Const TTM_ADDTOOL = WM_USER + 50
   Private Const TTM_DELTOOL = WM_USER + 51
   Private Const TTM_ENUMTOOLS = WM_USER + 58
   Private Const TTM_HITTEST = WM_USER + 55
   Private Const TTM_NEWTOOLRECT = WM_USER + 52
   Private Const TTM_SETTITLE = WM_USER + 33
   Private Const TTM_UPDATETIPTEXT = WM_USER + 57
#End If
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_ADJUSTRECT = WM_USER + 31
Private Const TTM_GETTOOLCOUNT = WM_USER + 13
Private Const TTM_GETBUBBLESIZE = WM_USER + 30
Private Const TTM_GETDELAYTIME = WM_USER + 21
Private Const TTM_GETMARGIN = WM_USER + 27
Private Const TTM_GETMAXTIPWIDTH = WM_USER + 25
Private Const TTM_GETTIPBKCOLOR = WM_USER + 22
Private Const TTM_GETTIPTEXTCOLOR = WM_USER + 23
Private Const TTM_POP = WM_USER + 28
Private Const TTM_SETDELAYTIME = WM_USER + 3
Private Const TTM_SETMARGIN = WM_USER + 26
Private Const TTM_SETMAXTIPWIDTH = WM_USER + 24
Private Const TTM_SETTIPBKCOLOR = WM_USER + 19
Private Const TTM_SETTIPTEXTCOLOR = WM_USER + 20
Private Const TTM_TRACKACTIVATE = WM_USER + 17
Private Const TTM_TRACKPOSITION = WM_USER + 18

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const TTF_IDISHWND = &H1
Private Const TTF_SUBCLASS = &H10

Private Type Size
   cx                   As Long
   cy                   As Long
End Type
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpString As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Enum TooltipFlagConstants
   ttfAlwaysTip = &H1
   ttfNoPrefix = &H2
   ttfNoAnimate = &H10
   ttfNoFade = &H20
   ttfBalloon = &H40
End Enum

Public Enum ToolFlagConstants
   tfCenterTip = &H2
   tfRtlReading = &H4
   tfTrack = &H20
   tfAbsolute = &H80
   tfTransparent = &H100
End Enum

Private Type INITCOMMONCONTROLEXSTRUCT
   dwSize               As Long
   dwICC                As Long
End Type

Private Type POINTAPI
   x                    As Long
   y                    As Long
End Type

Private Type TTHITTESTINFO
   hWnd                 As Long
   pt                   As POINTAPI
   ti                   As TOOLINFO
End Type

Private Type OLECOLOR
   RedOrSys             As Byte
   Green                As Byte
   Blue                 As Byte
   Type                 As Byte
End Type

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Sub AddTool(Ctl As Control, Flags As ToolFlagConstants, Optional Text As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : AddTool
   ' * Parameters       :
   ' *                    Ctl As Control
   ' *                    Flags As ToolFlagConstants
   ' *                    Optional Text As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Info             As TOOLINFO

   On Error Resume Next

   If hWndTT = 0 Then Exit Sub

   Info.cbSize = Len(Info)
   Info.uFlags = Flags
   If Not (Flags And tfTrack) Then
      Info.uFlags = Info.uFlags Or TTF_SUBCLASS
   End If
   Info.uFlags = Info.uFlags Or TTF_IDISHWND
   Info.hWnd = Ctl.Container.hWnd
   Info.hinst = App.hInstance
   Info.uId = Ctl.hWnd
   If Len(Text) > 0 Then
      Info.lpszText = Text
   End If

   SendMessage hWndTT, TTM_ADDTOOL, 0, Info
End Sub

Public Sub Create(frm As Form, Flags As TooltipFlagConstants)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : Create
   ' * Parameters       :
   ' *                    Frm As Form
   ' *                    Flags As TooltipFlagConstants
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim InitCtrls        As INITCOMMONCONTROLEXSTRUCT

   InitCtrls.dwSize = Len(InitCtrls)
   InitCtrls.dwICC = ICC_WIN95_CLASSES Or ICC_BAR_CLASSES

   Class_Terminate
   m_Enabled = True
   m_Icon = itNoIcon
   m_Title = ""

   InitCommonControlsEx InitCtrls

   hWndTT = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, "", _
      WS_POPUP Or Flags, CW_USEDEFAULT, CW_USEDEFAULT, _
      CW_USEDEFAULT, CW_USEDEFAULT, frm.hWnd, 0, App.hInstance, _
      ByVal 0&)

   SetWindowPos hWndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or _
      SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

Public Sub CreateHwnd(nHwnd As Long, Flags As TooltipFlagConstants)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : CreateHnwd
   ' * Parameters       :
   ' *                    nHwnd As Long
   ' *                    Flags As TooltipFlagConstants
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim InitCtrls        As INITCOMMONCONTROLEXSTRUCT

   InitCtrls.dwSize = Len(InitCtrls)
   InitCtrls.dwICC = ICC_WIN95_CLASSES Or ICC_BAR_CLASSES

   Class_Terminate
   m_Enabled = True
   m_Icon = itNoIcon
   m_Title = ""

   InitCommonControlsEx InitCtrls

   hWndTT = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, "", _
      WS_POPUP Or Flags, CW_USEDEFAULT, CW_USEDEFAULT, _
      CW_USEDEFAULT, CW_USEDEFAULT, nHwnd, 0, App.hInstance, _
      ByVal 0&)

   SetWindowPos hWndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or _
      SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

Public Sub AddToolHwnd(Ctl As Control, nHwndParent As Long, Flags As ToolFlagConstants, Optional Text As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : AddToolHwnd
   ' * Parameters       :
   ' *                    Ctl As Control
   ' *                    nHwndParent As Long
   ' *                    Flags As ToolFlagConstants
   ' *                    Optional Text As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Info             As TOOLINFO

   If hWndTT = 0 Then Exit Sub

   Info.cbSize = Len(Info)
   Info.uFlags = Flags
   If Not (Flags And tfTrack) Then
      Info.uFlags = Info.uFlags Or TTF_SUBCLASS
   End If
   Info.uFlags = Info.uFlags Or TTF_IDISHWND
   Info.hWnd = nHwndParent
   Info.hinst = App.hInstance
   Info.uId = Ctl.hWnd
   If Len(Text) > 0 Then
      Info.lpszText = Text
   End If

   SendMessage hWndTT, TTM_ADDTOOL, 0, Info
End Sub

Public Sub DeleteTool(Ctl As Control)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : DeleteTool
   ' * Parameters       :
   ' *                    Ctl As Control
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Info             As TOOLINFO

   If hWndTT = 0 Then Exit Sub

   Info = GetToolInfo(Ctl.hWnd)

   If Info.cbSize <> 0 Then
      SendMessage hWndTT, TTM_DELTOOL, 0, Info
   End If
End Sub

Public Property Get Enabled() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : Enabled
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   Enabled = m_Enabled
End Property

Public Property Let Enabled(NewState As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : Enabled
   ' * Parameters       :
   ' *                    NewState As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_Enabled = NewState

   If hWndTT = 0 Then Exit Property

   SendMessage hWndTT, TTM_ACTIVATE, m_Enabled, ByVal 0&
End Property

Public Function GetDelayTime(TimeMode As DelayTimeConstants) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : GetDelayTime
   ' * Parameters       :
   ' *                    TimeMode As DelayTimeConstants
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Function

   GetDelayTime = SendMessage(hWndTT, TTM_GETDELAYTIME, TimeMode, _
      0&)
End Function

Public Property Get Icon() As InfoTitleConstants
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : Icon
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   Icon = m_Icon
End Property

Public Function MakeLong(Int1 As Integer, Int2 As Integer) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MakeLong
   ' * Parameters       :
   ' *                    Int1 As Integer
   ' *                    Int2 As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Ints(1 To 2)     As Integer

   CopyMemory Ints(1), Int1, Len(Int1)
   CopyMemory Ints(2), Int2, Len(Int2)

   CopyMemory MakeLong, Ints(1), Len(MakeLong)
End Function

Public Sub SetTracking(Ctl As Control, Tracking As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : SetTracking
   ' * Parameters       :
   ' *                    Ctl As Control
   ' *                    Tracking As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Info             As TOOLINFO

   If hWndTT = 0 Then Exit Sub

   Info = GetToolInfo(Ctl.hWnd)
   If Info.cbSize <> 0 Then
      SendMessage hWndTT, TTM_TRACKACTIVATE, Tracking, Info
   End If
End Sub

Public Sub SetTrackPosition(Ctl As Control, x As Integer, y As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : SetTrackPosition
   ' * Parameters       :
   ' *                    Ctl As Control
   ' *                    X As Integer
   ' *                    Y As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Info             As TOOLINFO

   If hWndTT = 0 Then Exit Sub

   Info = GetToolInfo(Ctl.hWnd)
   If Info.cbSize <> 0 Then
      SendMessage hWndTT, TTM_TRACKPOSITION, 0, _
         ByVal MakeLong(x, y)
   End If
End Sub

Public Property Get Title() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : Title
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   Title = m_Title
End Property

Public Property Let Icon(NewIcon As InfoTitleConstants)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : Icon
   ' * Parameters       :
   ' *                    NewIcon As InfoTitleConstants
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   m_Icon = NewIcon

   SendMessage hWndTT, TTM_SETTITLE, m_Icon, ByVal Title
End Property

Public Property Let Title(newTitle As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : Title
   ' * Parameters       :
   ' *                    NewTitle As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   m_Title = newTitle

   SendMessage hWndTT, TTM_SETTITLE, Icon, ByVal m_Title
End Property

Public Sub SetDelayTime(TimeMode As SetDelayTimeConstants, time As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : SetDelayTime
   ' * Parameters       :
   ' *                    TimeMode As SetDelayTimeConstants
   ' *                    Time As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Sub

   SendMessage hWndTT, TTM_SETDELAYTIME, TimeMode, ByVal time
End Sub

Public Sub HideTips()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : HideTips
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Sub

   SendMessage hWndTT, TTM_POP, 0, ByVal 0&
End Sub

Public Function HitTest(Ctl As Control, x As Long, y As Long) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : HitTest
   ' * Parameters       :
   ' *                    Ctl As Control
   ' *                    X As Long
   ' *                    Y As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Info             As TTHITTESTINFO

   If hWndTT = 0 Then Exit Function

   With Info
      .hWnd = Ctl.hWnd
      .pt.x = x
      .pt.y = y
      .ti.cbSize = Len(.ti)
   End With

   HitTest = (SendMessage(Ctl.hWnd, TTM_HITTEST, 0, Info) <> 0)
End Function

Public Property Get MarginLeft() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MarginLeft
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Property

   SendMessage hWndTT, TTM_GETMARGIN, 0, rct

   MarginLeft = rct.left
End Property

Public Property Let MarginLeft(NewMargin As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MarginLeft
   ' * Parameters       :
   ' *                    NewMargin As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Property

   With rct
      .left = NewMargin
      .top = MarginTop
      .right = MarginRight
      .bottom = MarginBottom
   End With

   SendMessage hWndTT, TTM_SETMARGIN, 0, rct
End Property

Public Property Let MarginTop(NewMargin As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MarginTop
   ' * Parameters       :
   ' *                    NewMargin As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Property

   With rct
      .left = MarginLeft
      .top = NewMargin
      .right = MarginRight
      .bottom = MarginBottom
   End With

   SendMessage hWndTT, TTM_SETMARGIN, 0, rct
End Property

Public Property Let MarginRight(NewMargin As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MarginRight
   ' * Parameters       :
   ' *                    NewMargin As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Property

   With rct
      .left = MarginLeft
      .top = MarginTop
      .right = NewMargin
      .bottom = MarginBottom
   End With

   SendMessage hWndTT, TTM_SETMARGIN, 0, rct
End Property

Public Property Let MarginBottom(NewMargin As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MarginBottom
   ' * Parameters       :
   ' *                    NewMargin As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Property

   With rct
      .left = MarginLeft
      .top = MarginTop
      .right = MarginRight
      .bottom = NewMargin
   End With

   SendMessage hWndTT, TTM_SETMARGIN, 0, rct
End Property

Public Property Get MarginTop() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MarginTop
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Property

   SendMessage hWndTT, TTM_GETMARGIN, 0, rct

   MarginTop = rct.top
End Property

Public Property Get MarginRight() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MarginRight
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Property

   SendMessage hWndTT, TTM_GETMARGIN, 0, rct

   MarginRight = rct.right
End Property

Public Property Get MarginBottom() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MarginBottom
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Property

   SendMessage hWndTT, TTM_GETMARGIN, 0, rct

   MarginBottom = rct.bottom
End Property

Public Function GetShellVersion() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : GetShellVersion
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

End Function

Private Function GetToolInfo(ToolWnd As Long) As TOOLINFO
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : GetToolInfo
   ' * Parameters       :
   ' *                    ToolWnd As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim n                As Long, Info As TOOLINFO, Found As Boolean

   If hWndTT = 0 Then Exit Function

   Info.cbSize = Len(Info)

   For n = 0 To ToolCount - 1
      If SendMessage(hWndTT, TTM_ENUMTOOLS, n, Info) <> 0 Then
         If Info.uId = ToolWnd Then
            Found = True
            Exit For
         End If
      End If
   Next n

   If Found Then GetToolInfo = Info
End Function

Public Function GetTooltipWidth(Ctl As Control) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : GetTooltipWidth
   ' * Parameters       :
   ' *                    Ctl As Control
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim TipSize          As Long, Info As TOOLINFO

   If hWndTT = 0 Then Exit Function

   Info = GetToolInfo(Ctl.hWnd)

   If Info.cbSize <> 0 Then
      TipSize = SendMessage(hWndTT, TTM_GETBUBBLESIZE, 0, _
         Info)
   End If

   GetTooltipWidth = LoWord(TipSize)
End Function

Public Function GetTooltipHeight(Ctl As Control) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : GetTooltipHeight
   ' * Parameters       :
   ' *                    Ctl As Control
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim TipSize          As Long, Info As TOOLINFO

   If hWndTT = 0 Then Exit Function

   Info = GetToolInfo(Ctl.hWnd)

   If Info.cbSize <> 0 Then
      TipSize = SendMessage(hWndTT, TTM_GETBUBBLESIZE, 0, _
         Info)
   End If

   GetTooltipHeight = HiWord(TipSize)
End Function

Private Function LoWord(DWORD As Long) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : LoWord
   ' * Parameters       :
   ' *                    DWord As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Ints(1 To 2)     As Integer

   CopyMemory Ints(1), DWORD, Len(DWORD)

   LoWord = Ints(1)
End Function

Private Function HiWord(DWORD As Long) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : HiWord
   ' * Parameters       :
   ' *                    DWord As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Ints(1 To 2)     As Integer

   CopyMemory Ints(1), DWORD, Len(DWORD)

   HiWord = Ints(2)
End Function

Public Property Get MaxTipWidth() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MaxTipWidth
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   MaxTipWidth = SendMessage(hWndTT, TTM_GETMAXTIPWIDTH, 0, ByVal 0&)
End Property

Public Property Let MaxTipWidth(NewWidth As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : MaxTipWidth
   ' * Parameters       :
   ' *                    NewWidth As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   SendMessage hWndTT, TTM_SETMAXTIPWIDTH, 0, ByVal NewWidth
End Property

Public Sub SetToolRect(Ctl As Control, left As Long, top As Long, right As Long, bottom As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : SetToolRect
   ' * Parameters       :
   ' *                    Ctl As Control
   ' *                    Left As Long
   ' *                    Top As Long
   ' *                    Right As Long
   ' *                    Bottom As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Info             As TOOLINFO

   If hWndTT = 0 Then Exit Sub

   Info = GetToolInfo(Ctl.hWnd)
   With Info.WinRect
      .left = left
      .top = top
      .right = right
      .bottom = bottom
   End With

   If Info.cbSize <> 0 Then
      SendMessage hWndTT, TTM_NEWTOOLRECT, 0, Info
   End If
End Sub

Public Property Get ToolCount() As Long
   If hWndTT = 0 Then Exit Property

   ToolCount = SendMessage(hWndTT, TTM_GETTOOLCOUNT, 0, ByVal 0&)
End Property

Public Property Get BackColor() As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : BackColor
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   BackColor = SendMessage(hWndTT, TTM_GETTIPBKCOLOR, 0, ByVal 0&)
End Property

Public Property Let BackColor(NewColor As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : BackColor
   ' * Parameters       :
   ' *                    NewColor As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   SendMessage hWndTT, TTM_SETTIPBKCOLOR, WinColor(NewColor), ByVal 0&
End Property

Public Sub SetTipText(Ctl As Control, NewText As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : SetTipText
   ' * Parameters       :
   ' *                    Ctl As Control
   ' *                    NewText As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Info             As TOOLINFO

   If hWndTT = 0 Then Exit Sub

   Info = GetToolInfo(Ctl.hWnd)
   If Info.cbSize <> 0 Then
      Info.lpszText = NewText
      Info.cbSize = Len(Info)
      SendMessage hWndTT, TTM_UPDATETIPTEXT, 0, Info
   End If
End Sub

Public Property Let ForeColor(NewColor As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : ForeColor
   ' * Parameters       :
   ' *                    NewColor As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   SendMessage hWndTT, TTM_SETTIPTEXTCOLOR, WinColor(NewColor), ByVal 0&
End Property

Private Function WinColor(VBColor As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : WinColor
   ' * Parameters       :
   ' *                    VBColor As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim SysClr           As OLECOLOR

   CopyMemory SysClr, VBColor, Len(SysClr)

   If SysClr.Type = &H80 Then 'It is a system color
      'SysClr.RedOrSys is the index of the system color
      WinColor = GetSysColor(SysClr.RedOrSys)
   Else
      WinColor = VBColor
   End If
End Function

Public Property Get ForeColor() As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : ForeColor
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT = 0 Then Exit Property

   ForeColor = SendMessage(hWndTT, TTM_GETTIPTEXTCOLOR, 0, ByVal 0&)
End Property

Public Function ToolTipWidthFromText(hdc As Long, Text As String) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : ToolTipWidthFromText
   ' * Parameters       :
   ' *                    hDC As Long
   ' *                    Text As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT, Siz As Size

   If hWndTT = 0 Then Exit Function

   GetTextExtentPoint32 hdc, Text, LenB(Text), Siz

   rct.right = Siz.cx
   rct.bottom = Siz.cy
   SendMessage hWndTT, TTM_ADJUSTRECT, True, rct

   ToolTipWidthFromText = (rct.right - rct.left)
End Function

Public Function ToolTipHeightFromText(hdc As Long, Text As String) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : ToolTipHeightFromText
   ' * Parameters       :
   ' *                    hDC As Long
   ' *                    Text As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT
   Dim Siz              As Size

   If hWndTT = 0 Then Exit Function

   GetTextExtentPoint32 hdc, Text, LenB(Text), Siz

   rct.right = Siz.cx
   rct.bottom = Siz.cy
   SendMessage hWndTT, TTM_ADJUSTRECT, True, rct

   ToolTipHeightFromText = (rct.bottom - rct.top)
End Function

Public Function TextHeightFromToolTipHeight(hdc As Long, TextHeight As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : TextHeightFromToolTipHeight
   ' * Parameters       :
   ' *                    hDC As Long
   ' *                    TextHeight As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Function

   rct.bottom = TextHeight
   SendMessage hWndTT, TTM_ADJUSTRECT, False, rct

   TextHeightFromToolTipHeight = (rct.bottom - rct.top)
End Function

Public Function TextWidthFromToolTipWidth(hdc As Long, TextWidth As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : TextWidthFromToolTipWidth
   ' * Parameters       :
   ' *                    hDC As Long
   ' *                    TextWidth As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim rct              As RECT

   If hWndTT = 0 Then Exit Function

   rct.right = TextWidth
   SendMessage hWndTT, TTM_ADJUSTRECT, False, rct

   TextWidthFromToolTipWidth = (rct.right - rct.left)
End Function

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Thomas Kabir
   ' * Web Site         : http://home.t-online.de/home/vbfrood/english/about.htm
   ' * E-Mail           : contact@vbfrood.be
   ' * Date             : 16/02/2000
   ' * Time             : 12:58
   ' * Module Name      : cTooltips
   ' * Module Filename  : cTooltips.cls
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If hWndTT <> 0 Then DestroyWindow hWndTT
End Sub

