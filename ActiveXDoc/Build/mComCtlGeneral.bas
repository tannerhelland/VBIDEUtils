Attribute VB_Name = "mComCtlGeneral"
Option Explicit

Public Declare Sub InitCommonControls Lib "Comctl32.dll" ()
Public Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Const ICC_BAR_CLASSES = &H4
Public Const ICC_COOL_CLASSES = &H400

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Public Type ToolTipText
    hdr As NMHDR
    lpszText As Long
    szText As String * 80
    hInst As Long
    uFlags As Long
End Type

Public Const TTM_RELAYEVENT = (WM_USER + 7)
'Tool Tip messages
Public Const TTM_ACTIVATE = (WM_USER + 1)
#If UNICODE Then
    Public Const TTM_ADDTOOLW = (WM_USER + 50)
    Public Const TTM_ADDTOOL = TTM_ADDTOOLW
#Else
    Public Const TTM_ADDTOOLA = (WM_USER + 4)
    Public Const TTM_ADDTOOL = TTM_ADDTOOLA
#End If

'ToolTip Notification
Public Const TTN_FIRST = (H_MAX - 520&)
#If UNICODE Then
    Public Const TTN_NEEDTEXTW = (TTN_FIRST - 10&)
    Public Const TTN_NEEDTEXT = TTN_NEEDTEXTW
#Else
    Public Const TTN_NEEDTEXTA = (TTN_FIRST - 0&)
    Public Const TTN_NEEDTEXT = TTN_NEEDTEXTA
#End If

'//Common Control Constants
Public Const CCS_TOP = &H1
Public Const CCS_NOMOVEY = &H2
Public Const CCS_BOTTOM = &H3
Public Const CCS_NORESIZE = &H4
Public Const CCS_NOPARENTALIGN = &H8
Public Const CCS_NODIVIDER = &H40
Public Const CCS_VERT = &H80
Public Const CCS_LEFT = (CCS_VERT Or CCS_TOP)
Public Const CCS_RIGHT = (CCS_VERT Or CCS_BOTTOM)

