VERSION 5.00
Begin VB.UserControl TabControl 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ControlContainer=   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   2685
   ToolboxBitmap   =   "cTabCtrl.ctx":0000
End
Attribute VB_Name = "TabControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Declares and types:
' ======================================================================
' Windows general:
Private Const WM_USER = &H400
Private Const WM_NOTIFY = &H4E
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Const SW_HIDE = 0
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_BORDER = &H800000
Private Const WM_SETFONT = &H30
' Font
Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Const BITSPIXEL = 12
    Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
    Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Common controls general:
Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
Private Const TCM_FIRST = &H1300                   '// Tab control messages
Private Const CCM_FIRST = &H2000                   '// Common control shared messages
Private Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
Private Const H_MAX As Long = &HFFFF + 1
Private Const TCN_FIRST = H_MAX - 550                  '// tab control
Private Const NM_FIRST = H_MAX
Private Const NM_RCLICK = (NM_FIRST - 5)               '// uses NMCLICK struct


' //====== TAB CONTROL ==========================================================

' #ifndef NOTABCONTROL

' #ifdef _WIN32

Private Const WC_TABCONTROLA = "SysTabControl32"
'private const WC_TABCONTROLW          L"SysTabControl32"
' #ifdef UNICODE
'private const  WC_TABCONTROL          WC_TABCONTROLW
' #Else
Private Const WC_TABCONTROL = WC_TABCONTROLA
' #End If

' #Else
'private const WC_TABCONTROL           "SysTabControl"
' #End If

' // begin_r_commctrl

' #if (_WIN32_IE >= =&H0300)
Private Const TCS_SCROLLOPPOSITE = &H1          ' // assumes multiline tab
Private Const TCS_BOTTOM = &H2
Private Const TCS_RIGHT = &H2
Private Const TCS_MULTISELECT = &H4            ' // allow multi-select in button mode
' #End If
' #if (_WIN32_IE >= =&H0400)
Private Const TCS_FLATBUTTONS = &H8
' #End If
Private Const TCS_FORCEICONLEFT = &H10
Private Const TCS_FORCELABELLEFT = &H20
' #if (_WIN32_IE >= =&H0300)
Private Const TCS_HOTTRACK = &H40
Private Const TCS_VERTICAL = &H80
' #End If
Private Const TCS_TABS = &H0
Private Const TCS_BUTTONS = &H100
Private Const TCS_SINGLELINE = &H0
Private Const TCS_MULTILINE = &H200
Private Const TCS_RIGHTJUSTIFY = &H0
Private Const TCS_FIXEDWIDTH = &H400
Private Const TCS_RAGGEDRIGHT = &H800
Private Const TCS_FOCUSONBUTTONDOWN = &H1000
Private Const TCS_OWNERDRAWFIXED = &H2000
Private Const TCS_TOOLTIPS = &H4000
Private Const TCS_FOCUSNEVER = &H8000

' #if (_WIN32_IE >= =&H0400)
' // EX styles for use with TCM_SETEXTENDEDSTYLE
Private Const TCS_EX_FLATSEPARATORS = &H1
Private Const TCS_EX_REGISTERDROP = &H2
' #End If

' // end_r_commctrl


Private Const TCM_GETIMAGELIST = (TCM_FIRST + 2)
'private const TabCtrl_GetImageList(hwnd) \
'    (HIMAGELIST)SNDMSG((hwnd), TCM_GETIMAGELIST, 0, 0L)


Private Const TCM_SETIMAGELIST = (TCM_FIRST + 3)
    'private const TabCtrl_SetImageList(hwnd, himl) \
    '    (HIMAGELIST)SNDMSG((hwnd), TCM_SETIMAGELIST, 0, (LPARAM)(UINT)(HIMAGELIST)(himl))


Private Const TCM_GETITEMCOUNT = (TCM_FIRST + 4)
    'private const TabCtrl_GetItemCount(hwnd) \
    '    (int)SNDMSG((hwnd), TCM_GETITEMCOUNT, 0, 0L)

Private Const TCIF_TEXT = &H1
Private Const TCIF_IMAGE = &H2
Private Const TCIF_RTLREADING = &H4
Private Const TCIF_PARAM = &H8
' #if (_WIN32_IE >= =&H0300)
Private Const TCIF_STATE = &H10


Private Const TCIS_BUTTONPRESSED = &H1
' #End If
' #if (_WIN32_IE >= =&H0400)
Private Const TCIS_HIGHLIGHTED = &H2
' #End If

' #if (_WIN32_IE >= =&H0300)
'Private Const TC_ITEMHEADERA = TCITEMHEADERA
'private const TC_ITEMHEADERW         TCITEMHEADERW
' #Else
'private const tagTCITEMHEADERA       _TC_ITEMHEADERA
'private const    TCITEMHEADERA        TC_ITEMHEADERA
'private const tagTCITEMHEADERW       _TC_ITEMHEADERW
'private const    TCITEMHEADERW        TC_ITEMHEADERW
' #End If
'private const TC_ITEMHEADER          TCITEMHEADER

Private Type TCITEMHEADER
    mask As Long
    lpReserved1 As Long
    lpReserved2 As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
End Type
Private Type TCITEMHEADER_NOTEXT
    mask As Long
    lpReserved1 As Long
    lpReserved2 As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
End Type

'typedef struct tagTCITEMHEADERW
'{
'    UINT mask;
'    UINT lpReserved1;
'    UINT lpReserved2;
'    LPWSTR pszText;
'    int cchTextMax;
'    int iImage;
'} TCITEMHEADERW, FAR *LPTCITEMHEADERW;

' #ifdef UNICODE
'private const  TCITEMHEADER          TCITEMHEADERW
'private const  LPTCITEMHEADER        LPTCITEMHEADERW
'' #Else
'private const  TCITEMHEADER          TCITEMHEADERA
'private const  LPTCITEMHEADER        LPTCITEMHEADERA
' #End If


' #if (_WIN32_IE >= =&H0300)
'private const TC_ITEMA                TCITEMA
'private const TC_ITEMW                TCITEMW
' #Else
'private const tagTCITEMA              _TC_ITEMA
'private const    TCITEMA               TC_ITEMA
'private const tagTCITEMW              _TC_ITEMW
'private const    TCITEMW               TC_ITEMW
' #End If
'private const TC_ITEM                 TCITEM

Private Type TCITEM
    mask As Long
' #if (_WIN32_IE >= =&H0300)
    dwState As Long
    dwStateMask As Long
' #Else
'    UINT lpReserved1;
'    UINT lpReserved2;
' #End If
    pszText As String
    cchTextMax As Long
    iImage As Long

    lParam As Long
End Type
Private Type TCITEM_NOTEXT
    mask As Long
' #if (_WIN32_IE >= =&H0300)
    dwState As Long
    dwStateMask As Long
' #Else
'    UINT lpReserved1;
'    UINT lpReserved2;
' #End If
    pszText As Long
    cchTextMax As Long
    iImage As Long

    lParam As Long
End Type


'typedef struct tagTCITEMW
'{
'    UINT mask;
' #if (_WIN32_IE >= =&H0300)
'    DWORD dwState;
'    DWORD dwStateMask;
' #Else
'    UINT lpReserved1;
'    UINT lpReserved2;
' #End If
'    LPWSTR pszText;
'    int cchTextMax;
'    int iImage;

'    LPARAM lParam;
'} TCITEMW, FAR *LPTCITEMW;
'
' #ifdef UNICODE
'private const  TCITEM                 TCITEMW
'private const  LPTCITEM               LPTCITEMW
' #Else
'private const  TCITEM                 TCITEMA
'private const  LPTCITEM               LPTCITEMA
' #End If


Private Const TCM_GETITEMA = (TCM_FIRST + 5)
Private Const TCM_GETITEMW = (TCM_FIRST + 60)

' #ifdef UNICODE
'private const TCM_GETITEM             TCM_GETITEMW
' #Else
Private Const TCM_GETITEM = TCM_GETITEMA
' #End If

'private const TabCtrl_GetItem(hwnd, iItem, pitem) \
'    (BOOL)SNDMSG((hwnd), TCM_GETITEM, (WPARAM)(int)iItem, (LPARAM)=(TC_ITEM FAR*)(pitem))


Private Const TCM_SETITEMA = (TCM_FIRST + 6)
Private Const TCM_SETITEMW = (TCM_FIRST + 61)

' #ifdef UNICODE
'private const TCM_SETITEM             TCM_SETITEMW
' #Else
Private Const TCM_SETITEM = TCM_SETITEMA
' #End If

'private const TabCtrl_SetItem(hwnd, iItem, pitem) \
'    (BOOL)SNDMSG((hwnd), TCM_SETITEM, (WPARAM)(int)iItem, (LPARAM)=(TC_ITEM FAR*)(pitem))


Private Const TCM_INSERTITEMA = (TCM_FIRST + 7)
Private Const TCM_INSERTITEMW = (TCM_FIRST + 62)

' #ifdef UNICODE
'private const TCM_INSERTITEM          TCM_INSERTITEMW
' #Else
Private Const TCM_INSERTITEM = TCM_INSERTITEMA
' #End If

'private const TabCtrl_InsertItem(hwnd, iItem, pitem)   \
'    (int)SNDMSG((hwnd), TCM_INSERTITEM, (WPARAM)(int)iItem, (LPARAM)(const TC_ITEM FAR*)(pitem))


Private Const TCM_DELETEITEM = (TCM_FIRST + 8)
'private const TabCtrl_DeleteItem(hwnd, i) \
'    (BOOL)SNDMSG((hwnd), TCM_DELETEITEM, (WPARAM)(int)(i), 0L)


Private Const TCM_DELETEALLITEMS = (TCM_FIRST + 9)
'private const TabCtrl_DeleteAllItems(hwnd) \
'    (BOOL)SNDMSG((hwnd), TCM_DELETEALLITEMS, 0, 0L)


Private Const TCM_GETITEMRECT = (TCM_FIRST + 10)
'private const TabCtrl_GetItemRect(hwnd, i, prc) \
'    (BOOL)SNDMSG((hwnd), TCM_GETITEMRECT, (WPARAM)(int)(i), (LPARAM)(RECT FAR*)(prc))


Private Const TCM_GETCURSEL = (TCM_FIRST + 11)
'private const TabCtrl_GetCurSel(hwnd) \
'    (int)SNDMSG((hwnd), TCM_GETCURSEL, 0, 0)


Private Const TCM_SETCURSEL = (TCM_FIRST + 12)
'private const TabCtrl_SetCurSel(hwnd, i) \
'    (int)SNDMSG((hwnd), TCM_SETCURSEL, (WPARAM)i, 0)


Private Const TCHT_NOWHERE = &H1
Private Const TCHT_ONITEMICON = &H2
Private Const TCHT_ONITEMLABEL = &H4
Private Const TCHT_ONITEM = (TCHT_ONITEMICON Or TCHT_ONITEMLABEL)

' #if (_WIN32_IE >= =&H0300)
'private const LPTC_HITTESTINFO        LPTCHITTESTINFO
'private const TC_HITTESTINFO          TCHITTESTINFO
' #Else
'private const tagTCHITTESTINFO        _TC_HITTESTINFO
'private const    TCHITTESTINFO         TC_HITTESTINFO
'private const  LPTCHITTESTINFO       LPTC_HITTESTINFO
' #End If

Private Type TCHITTESTINFO
    pt As POINTAPI
    flags As Long
End Type

Private Const TCM_HITTEST = (TCM_FIRST + 13)
'private const TabCtrl_HitTest(hwndTC, pinfo) \
'    (int)SNDMSG((hwndTC), TCM_HITTEST, 0, (LPARAM)=(TC_HITTESTINFO FAR*)(pinfo))


Private Const TCM_SETITEMEXTRA = (TCM_FIRST + 14)
'private const TabCtrl_SetItemExtra(hwndTC, cb) \
'    (BOOL)SNDMSG((hwndTC), TCM_SETITEMEXTRA, (WPARAM)(cb), 0L)


Private Const TCM_ADJUSTRECT = (TCM_FIRST + 40)
'private const TabCtrl_AdjustRect(hwnd, bLarger, prc) \
'    (int)SNDMSG(hwnd, TCM_ADJUSTRECT, (WPARAM)(BOOL)bLarger, (LPARAM)(RECT FAR *)prc)


Private Const TCM_SETITEMSIZE = (TCM_FIRST + 41)
'private const TabCtrl_SetItemSize(hwnd, x, y) \
'    (DWORD)SNDMSG((hwnd), TCM_SETITEMSIZE, 0, MAKELPARAM(x,y))


Private Const TCM_REMOVEIMAGE = (TCM_FIRST + 42)
'private const TabCtrl_RemoveImage(hwnd, i) \
'        (void)SNDMSG((hwnd), TCM_REMOVEIMAGE, i, 0L)


Private Const TCM_SETPADDING = (TCM_FIRST + 43)
'private const TabCtrl_SetPadding(hwnd,  cx, cy) \
'        (void)SNDMSG((hwnd), TCM_SETPADDING, 0, MAKELPARAM(cx, cy))


Private Const TCM_GETROWCOUNT = (TCM_FIRST + 44)
'private const TabCtrl_GetRowCount(hwnd) \
'        (int)SNDMSG((hwnd), TCM_GETROWCOUNT, 0, 0L)


Private Const TCM_GETTOOLTIPS = (TCM_FIRST + 45)
'private const TabCtrl_GetToolTips(hwnd) \
'        (HWND)SNDMSG((hwnd), TCM_GETTOOLTIPS, 0, 0L)


Private Const TCM_SETTOOLTIPS = (TCM_FIRST + 46)
'private const TabCtrl_SetToolTips(hwnd, hwndTT) \
'        (void)SNDMSG((hwnd), TCM_SETTOOLTIPS, (WPARAM)hwndTT, 0L)


Private Const TCM_GETCURFOCUS = (TCM_FIRST + 47)
'private const TabCtrl_GetCurFocus(hwnd) \
'    (int)SNDMSG((hwnd), TCM_GETCURFOCUS, 0, 0)

Private Const TCM_SETCURFOCUS = (TCM_FIRST + 48)
'private const TabCtrl_SetCurFocus(hwnd, i) \
'    SNDMSG((hwnd),TCM_SETCURFOCUS, i, 0)

' #if (_WIN32_IE >= =&H0300)
Private Const TCM_SETMINTABWIDTH = (TCM_FIRST + 49)
'private const TabCtrl_SetMinTabWidth(hwnd, x) \
'        (int)SNDMSG((hwnd), TCM_SETMINTABWIDTH, 0, x)


Private Const TCM_DESELECTALL = (TCM_FIRST + 50)
'private const TabCtrl_DeselectAll(hwnd, fExcludeFocus)\
'        (void)SNDMSG((hwnd), TCM_DESELECTALL, fExcludeFocus, 0)
' #End If

' #if (_WIN32_IE >= =&H0400)

Private Const TCM_HIGHLIGHTITEM = (TCM_FIRST + 51)
'private const TabCtrl_HighlightItem(hwnd, i, fHighlight) \
'    (BOOL)SNDMSG((hwnd), TCM_HIGHLIGHTITEM, (WPARAM)i, (LPARAM)MAKELONG (fHighlight, 0))

Private Const TCM_SETEXTENDEDSTYLE = (TCM_FIRST + 52)    ' // optional wParam == mask
'private const TabCtrl_SetExtendedStyle(hwnd, dw)\
'        (DWORD)SNDMSG((hwnd), TCM_SETEXTENDEDSTYLE, 0, dw)

Private Const TCM_GETEXTENDEDSTYLE = (TCM_FIRST + 53)
'private const TabCtrl_GetExtendedStyle(hwnd)\
'        (DWORD)SNDMSG((hwnd), TCM_GETEXTENDEDSTYLE, 0, 0)

Private Const TCM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
'private const TabCtrl_SetUnicodeFormat(hwnd, fUnicode)  \
'    (BOOL)SNDMSG((hwnd), TCM_SETUNICODEFORMAT, (WPARAM)(fUnicode), 0)

Private Const TCM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
'private const TabCtrl_GetUnicodeFormat(hwnd)  \
'    (BOOL)SNDMSG((hwnd), TCM_GETUNICODEFORMAT, 0, 0)

' #End If     ' // _WIN32_IE >= =&H0400

Private Const TCN_KEYDOWN = (TCN_FIRST - 0)

' #if (_WIN32_IE >= =&H0300)
'private const TC_KEYDOWN              NMTCKEYDOWN
' #Else
'private const tagTCKEYDOWN            _TC_KEYDOWN
'private const  NMTCKEYDOWN             TC_KEYDOWN
' #End If

Private Type TCKEYDOWN
    hdr As NMHDR
    wVKey As Long
    flags As Long
End Type

Private Const TCN_SELCHANGE = (TCN_FIRST - 1)
Private Const TCN_SELCHANGING = (TCN_FIRST - 2)
' #if (_WIN32_IE >= =&H0400)
Private Const TCN_GETOBJECT = (TCN_FIRST - 3)
' #End If     ' // _WIN32_IE >= =&H0400

' #End If     ' // NOTABCONTROL


' ======================================================================
' Interface:
' ======================================================================

' ======================================================================
' Private Implementation:
' ======================================================================
Implements ISubclass
Private m_emr As EMsgResponse
Private m_bSubClassing As Boolean

Private m_hWnd As Long
Private m_hIml As Long
Private m_sKey() As String
Private m_tULF As LOGFONT
Private m_hFnt As Long

Public Event BeforeClick(ByVal lTab As Long, ByRef bCancel As Boolean)
Public Event TabClick(ByVal lTab As Long)
Public Event TabRightClick()


Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Set Font(sFont As StdFont)
Dim hFnt As Long
    If Not (UserControl.Font Is sFont) Then
        Set UserControl.Font = sFont
        ' Store a log font structure for this font:
        pOLEFontToLogFont sFont, UserControl.hdc, m_tULF
        ' Store old font handle:
        hFnt = m_hFnt
        ' Create a new version of the font:
        m_hFnt = CreateFontIndirect(m_tULF)
        ' Ensure the edit portion has the correct font:
        If (m_hWnd <> 0) Then
            SendMessage m_hWnd, WM_SETFONT, m_hFnt, 1
        End If
        ' Delete previous version, if we had one:
        If (hFnt <> 0) Then
            DeleteObject hFnt
        End If
        PropertyChanged "Font"
    End If
End Property
Private Sub pOLEFontToLogFont(fntThis As StdFont, hdc As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer

    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name
        ' There is a quicker way involving StrConv and CopyMemory, but
        ' this is simpler!:
        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
        Next iChar
        ' Based on the Win32SDK documentation:
        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
    End With

End Sub

Public Property Let ImageList( _
        ByRef vImageList As Variant _
    )
    If VarType(vImageList) = vbObject Then
        ' Assume VB ImageList control.  Note that unless
        ' some call has been made to an object within a
        ' VB ImageList the image list itself is not
        ' created.  Therefore hImageList returns error. So
        ' ensure that the ImageList has been initialised by
        ' drawing into nowhere:
        On Error Resume Next
        ' Get the image list initialised..
        vImageList.ListImages(1).Draw 0, 0, 0, 1
        m_hIml = vImageList.hImageList
        If (Err.Number <> 0) Then
            m_hIml = 0
            Debug.Print "vImageList Parameter is invalid."
            On Error GoTo 0
        Else
            On Error GoTo 0
            pSetImageList
        End If
    ElseIf VarType(vImageList) = vbLong Then
        m_hIml = vImageList
        pSetImageList
    Else
        Debug.Print "vImageList Parameter is invalid."
    End If
End Property
Private Sub pSetImageList()
    SendMessageLong m_hWnd, TCM_SETIMAGELIST, 0, m_hIml
End Sub

Public Sub AddTab( _
        ByVal sText As String, _
        Optional ByVal iIconIndex As Long = -1, _
        Optional ByVal lIndex As Long = -1, _
        Optional ByVal sKey As String _
    )
Dim tTCI As TCITEM
Dim lTabCount As Long
Dim lKey As Long

    ' Set up the tab to add:
    lTabCount = TabCount
    With tTCI
        .lParam = lTabCount
        .mask = TCIF_TEXT Or TCIF_IMAGE
        .iImage = iIconIndex
        .cchTextMax = Len(sText)
        .pszText = sText
    End With
    ReDim Preserve m_sKey(0 To lTabCount) As String
    If (lIndex = -1) Then
        lIndex = lTabCount
    Else
        lIndex = lIndex - 1
    End If
    
    ' Add the tab:
    If (SendMessage(m_hWnd, TCM_INSERTITEM, lIndex, tTCI) = 0) Then
        Debug.Print "Failed to insert tab"
    Else
        ' Add the key:
        For lKey = lIndex + 1 To lTabCount
            m_sKey(lKey) = m_sKey(lKey - 1)
        Next lKey
        m_sKey(lIndex) = sKey
    End If

End Sub
Public Property Get SelectedTab() As Long
Dim lTab As Long
    SelectedTab = SendMessageLong(m_hWnd, TCM_GETCURSEL, 0, 0) + 1
End Property
Public Property Get ClientLeft() As Long
Dim rc As RECT
    pGetClientRect rc
    ClientLeft = rc.Left * Screen.TwipsPerPixelX
End Property
Public Property Get ClientTop() As Long
Dim rc As RECT
    pGetClientRect rc
    ClientTop = rc.Top * Screen.TwipsPerPixelY
End Property
Public Property Get ClientWidth() As Long
Dim rc As RECT
    pGetClientRect rc
    ClientWidth = (rc.Right - rc.Left) * Screen.TwipsPerPixelX
End Property
Public Property Get ClientHeight() As Long
Dim rc As RECT
    pGetClientRect rc
    ClientHeight = (rc.Bottom - rc.Top) * Screen.TwipsPerPixelY
End Property
Private Sub pGetClientRect(rc As RECT)
Dim tP As POINTAPI
    ' Get window rect of the user control:
    GetWindowRect UserControl.hwnd, rc
    tP.x = rc.Left
    tP.y = rc.Top
    ' Adjust to coordinates of user control's container:
    ScreenToClient GetParent(UserControl.hwnd), tP
    rc.Right = rc.Right + (tP.x - rc.Left)
    rc.Bottom = rc.Bottom + (tP.y - rc.Top)
    rc.Left = tP.x
    rc.Top = tP.y
    ' Calculate the useable area of the tab:
    SendMessage m_hWnd, TCM_ADJUSTRECT, 0, rc
End Sub
Public Property Get TabText(ByVal vKey As Variant) As String
Dim lIndex As Long
    
End Property

Public Property Get TabIndex(ByVal vKey As Variant) As Long
Dim lS As Long
Dim lKey As Long
    TabIndex = -1
    If IsNumeric(vKey) Then
        lKey = CLng(vKey)
    Else
        For lS = 0 To TabCount - 1
            If (m_sKey(lS) = vKey) Then
                TabIndex = lKey
                Exit For
            End If
        Next lS
    End If
    
    If (lKey > 0) And (lKey <= TabCount) Then
        TabIndex = lKey
    Else
        Debug.Print "TabIndex '" & vKey & "' does not exist"
    End If

End Property
Public Property Get TabCount() As Long
    TabCount = SendMessageLong(m_hWnd, TCM_GETITEMCOUNT, 0, 0)
End Property

Public Property Get hwnd() As Long
    hwnd = m_hWnd
End Property
Private Sub pInitialise()
Dim dwStyle As Long
        
    ' Ensure we don't already have Tab control:
    pTerminate
    
    ' Ensure common controls:
    InitCommonControls
    
    ' Create the control:
    dwStyle = WS_VISIBLE Or WS_CHILD Or WS_CLIPSIBLINGS
    
    
    m_hWnd = CreateWindowEX( _
        0, WC_TABCONTROL, "", _
        dwStyle, _
        0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, _
        UserControl.hwnd, 0, _
        App.hInstance, 0)
        
    Debug.Assert m_hWnd <> 0
    If (m_hWnd <> 0) Then
        If (UserControl.Ambient.UserMode) Then
            ' Attach messages to the control:
            pAttachMessages
        Else
            AddTab "Tab Control"
        End If
    End If
    
End Sub
Private Sub pAttachMessages()
    m_emr = emrPreprocess
    AttachMessage Me, UserControl.hwnd, WM_NOTIFY
    m_bSubClassing = True
End Sub
Private Sub pDetachMessages()
    If (m_bSubClassing) Then
        DetachMessage Me, UserControl.hwnd, WM_NOTIFY
        m_bSubClassing = False
    End If
End Sub
Private Sub pTerminate()
    If (m_hWnd <> 0) Then
        ' Stop subclassing:
        pDetachMessages
        ' Destroy the window:
        ShowWindow m_hWnd, SW_HIDE
        SetParent m_hWnd, 0
        Debug.Print DestroyWindow(m_hWnd)
        ' store that we haven't a window:
        m_hWnd = 0
    End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As VBIDEUtils0.EMsgResponse)
    RHS = m_emr
End Property

Private Property Get ISubclass_MsgResponse() As VBIDEUtils0.EMsgResponse
    ISubclass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tNM As NMHDR
Dim lTab As Long
Dim bCancel As Boolean
    
    If (iMsg = WM_NOTIFY) Then
        CopyMemory tNM, ByVal lParam, Len(tNM)
        If (tNM.hwndFrom = m_hWnd) Then
            Select Case tNM.code
            Case TCN_SELCHANGING
                lTab = SelectedTab
                RaiseEvent BeforeClick(lTab, bCancel)
                If (bCancel) Then
                    ISubclass_WindowProc = 1
                End If
            Case TCN_SELCHANGE
                lTab = SelectedTab
                RaiseEvent TabClick(lTab)
            Case NM_RCLICK
                RaiseEvent TabRightClick
            End Select
        End If
    End If
End Function

Private Sub UserControl_Initialize()
    Debug.Print "cTabCtrl:Initialize"
End Sub

Private Sub UserControl_InitProperties()
    pInitialise
    Set Font = UserControl.Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    pInitialise
    
    Dim sFnt As New StdFont
    sFnt.Name = "MS Sans Serif"
    sFnt.Size = 8
    Set Font = PropBag.ReadProperty("Font", sFnt)
    
End Sub

Private Sub UserControl_Resize()
    If (m_hWnd <> 0) Then
        MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1
    End If
End Sub

Private Sub UserControl_Terminate()
    pTerminate
    Debug.Print "cUpDown:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    pTerminate
    
    Dim sFnt As New StdFont
    sFnt.Name = "MS Sans Serif"
    sFnt.Size = 8
    PropBag.WriteProperty "Font", Font, sFnt
End Sub

