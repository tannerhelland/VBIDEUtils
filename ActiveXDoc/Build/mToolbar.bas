Attribute VB_Name = "mToolbar"
Option Explicit

Public m_cTT As New cToolTip

Private m_HWndToolTip As Long
Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uId As Long
    rct As RECT
    hinst As Long
    lpszText As Long
End Type

'Tooltips
Private Type msg
    hwnd As Long
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

Public Type NMTTDISPINFO
    hdr As NMHDR
    lpszText As Long
    szText(0 To 79) As Byte
    hinst As Long
    uFlags As Long
'#if (_WIN32_IE >= 0x0300)
    lParam As Long
'#End If
End Type
Private Const TOOLTIPS_CLASS As String = "tooltips_class32"
 
Private Const CW_USEDEFAULT  As Long = &H80000000
Private Const glSUNKEN_OFFSET = 1
Private Const GDI_ERROR = &HFFFFFFFF

'Windows Messages
Private Const WM_CANCELMODE = &H1F

'Resource String Indexes
Private Const giINVALID_PIC_TYPE As Integer = 10
 
'Image type
Private Const DST_ICON = &H3&
Private Const DST_BITMAP = &H4&

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
Private Const TTM_RELAYEVENT = (WM_USER + 7)

'ToolTip Notification
Private Const TTN_FIRST = (H_MAX - 520&)
#If UNICODE Then
Private Const TTN_NEEDTEXTW = (TTN_FIRST - 10&)
Private Const TTN_NEEDTEXT = TTN_NEEDTEXTW
#Else
Private Const TTN_NEEDTEXTA = (TTN_FIRST - 0&)
Private Const TTN_NEEDTEXT = TTN_NEEDTEXTA
#End If

Private Const LPSTR_TEXTCALLBACK As Long = -1

Private Const TBN_FIRST = -700&
Private Const TBN_DROPDOWN = (TBN_FIRST - 10)


Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    bReserved1 As Byte
    bReserved2 As Byte
    dwData As Long
    iString As Long
End Type
 
Type NMTOOLBAR
    hdr As NMHDR
    iItem As Long
    TBBUTTON As TBBUTTON
    cchText As Long
    pszText As String
End Type

