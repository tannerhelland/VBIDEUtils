VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_HWndToolTip As Long

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    rct As RECT
    hInst As Long
    lpszText As String 'Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type ToolTipText
    hdr As NMHDR
    lpszText As Long
    szText As String * 80
    hInst As Long
    uFlags As Long
End Type

'Misc Constants
Private Const TOOLTIPS_CLASS = "tooltips_class32"

'Windows Messages
Private Const WM_CANCELMODE = &H1F

'Resource String Indexes
Private Const giINVALID_PIC_TYPE As Integer = 10

'Get Windows Long Constants
Private Const GWL_USERDATA = (-21)
Private Const GWL_WNDPROC = (-4)

'Messages to relay to ToolTip
Private Const WM_USER = &H400
Private Const WM_NOTIFY = &H4E
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

'ToolTip style
Private Const TTF_IDISHWND = &H1

'Tool Tip messages
Private Const TTM_ACTIVATE = (WM_USER + 1)
#If UNICODE Then
    Private Const TTM_ADDTOOLW = (WM_USER + 50)
    Private Const TTM_ADDTOOL = TTM_ADDTOOLW
#Else
    Private Const TTM_ADDTOOLA = (WM_USER + 4)
    Private Const TTM_ADDTOOL = TTM_ADDTOOLA
#End If

'ToolTip Notification
Private Const TTN_FIRST = (H_MAX - 520&)
#If UNICODE Then
    Private Const TTN_NEEDTEXTW = (TTN_FIRST - 10&)
    Private Const TTN_NEEDTEXT = TTN_NEEDTEXTW
#Else
    Private Const TTN_NEEDTEXTA = (TTN_FIRST - 0&)
    Private Const TTN_NEEDTEXT = TTN_NEEDTEXTA
#End If

'Misc ToolTip
Private Const LPSTR_TEXTCALLBACK As Long = -1

 'VB Errors
Private Const giOBJECT_VARIABLE_NOT_SET As Integer = 91
Private Const giINVALID_PICTURE As Integer = 481
Private Const giDLL_FUNCTION_NOT_FOUND As Integer = 453

'Windows Errors
Private Const ERROR_CALL_NOT_IMPLEMENTED As Long = 120

Public Property Get hWnd() As Long
    hWnd = m_HWndToolTip
End Property

 
Public Sub AddTool(hWnd As Long)
 'Add the hWnd of the Control being subclassed
Dim ti As TOOLINFO

    With ti
        .cbSize = Len(ti)
        .uId = hWnd
        .hWnd = hWnd
        .hInst = App.hInstance
        .uFlags = TTF_IDISHWND
        .lpszText = LPSTR_TEXTCALLBACK
    End With

    SendMessage m_HWndToolTip, TTM_ADDTOOL, 0, ti
    SendMessage m_HWndToolTip, TTM_ACTIVATE, 1, ByVal hWnd

End Sub
 
Public Property Get ToolTipText() As String
    ToolTipText = ToolTipText
End Property
Public Property Let ToolTipText(ByVal sToolTip As String)
    ToolTipText = sToolTip
    
    If Len(sToolTip) <> 0 Then
        ' Doesn't do anything...
        'mbToolTipNotInExtender = True
    End If
   
End Property
 

 
Public Sub Create()
    InitCommonControls
    m_HWndToolTip = CreateWindowEX(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, 0, _
              CW_USEDEFAULT, CW_USEDEFAULT, _
              CW_USEDEFAULT, CW_USEDEFAULT, _
              0, 0, _
              App.hInstance, _
              ByVal 0)
    SendMessage m_HWndToolTip, TTM_ACTIVATE, 1, ByVal 0
End Sub


