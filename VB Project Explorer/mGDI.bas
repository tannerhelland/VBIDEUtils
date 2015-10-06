Attribute VB_Name = "mAPIAndCallbacks"
Option Explicit
' ======================================================================================
' Name:     mGDI
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     22 December 1998
'
' Copyright © 1998-1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Various GDI declares and helper functions for the vbAcceleratorGrid
' control.
'
' FREE SOURCE CODE - ENJOY!
' ======================================================================================
#Const DEBUGMODE = 0

Type BITMAP '14 bytes
   bmType               As Long
   bmWidth              As Long
   bmHeight             As Long
   bmWidthBytes         As Long
   bmPlanes             As Integer
   bmBitsPixel          As Integer
   bmBits               As Long
End Type

Public Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type
Public Type LOGBRUSH
   lbStyle              As Long
   lbColor              As Long
   lbHatch              As Long
End Type
Public Type POINTAPI
   x                    As Long
   y                    As Long
End Type

Public Enum ImageTypes
   IMAGE_BITMAP = 0
   IMAGE_ICON = 1
   IMAGE_CURSOR = 2
End Enum

Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function DrawFrameControl Lib "user32" (ByVal lHDC As Long, tR As RECT, ByVal eFlag As DFCFlags, ByVal eStyle As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Const DT_NOFULLWIDTHCHARBREAK = &H80000
Public Const DT_HIDEPREFIX = &H100000
Public Const DT_PREFIXONLY = &H200000
Public Const HC_ACTION = 0

Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Private Const LF_FACESIZE = 32
Public Type LOGFONT
   lfHeight             As Long ' The font size (see below)
   lfWidth              As Long ' Normally you don't set this, just let Windows create the Default
   lfEscapement         As Long ' The angle, in 0.1 degrees, of the font
   lfOrientation        As Long ' Leave as default
   lfWeight             As Long ' Bold, Extra Bold, Normal etc
   lfItalic             As Byte ' As it says
   lfUnderline          As Byte ' As it says
   lfStrikeOut          As Byte ' As it says
   lfCharSet            As Byte ' As it says
   lfOutPrecision       As Byte ' Leave for default
   lfClipPrecision      As Byte ' Leave for default
   lfQuality            As Byte ' Leave for default
   lfPitchAndFamily     As Byte ' Leave for default
   lfFaceName(LF_FACESIZE) As Byte ' The font name converted to a byte array
End Type

Private Const BITSPIXEL = 12         '  Number of bits per pixel

' Owner draw information:
Public Const ODS_CHECKED = &H8
Public Const ODS_DISABLED = &H4
Public Const ODS_FOCUS = &H10
Public Const ODS_GRAYED = &H2
Public Const ODS_SELECTED = &H1
Public Const ODT_BUTTON = 4
Public Const ODT_COMBOBOX = 3
Public Const ODT_LISTBOX = 2
Public Const ODT_MENU = 1

Public Const MIIM_STATE = &H1&
Public Const MIIM_ID = &H2&
Public Const MIIM_SUBMENU = &H4&
Public Const MIIM_CHECKMARKS = &H8&
Public Const MIIM_TYPE = &H10&
Public Const MIIM_DATA = &H20&

' /* State type */
Public Const DSS_NORMAL = &H0&
Public Const DSS_UNION = &H10& ' Dither
Public Const DSS_DISABLED = &H20&
Public Const DSS_MONO = &H80& ' Draw in colour of brush specified in hBrush
Public Const DSS_RIGHT = &H8000&

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const BF_DIAGONAL = &H10

Public Enum DFCFlags
   DFC_CAPTION = 1
   DFC_MENU = 2
   DFC_SCROLL = 3
   DFC_BUTTON = 4
   'Win98/2000 only
   DFC_POPUPMENU = 5
End Enum
Public Enum ECGTextAlignFlags
   DT_TOP = &H0&
   DT_LEFT = &H0&
   DT_CENTER = &H1&
   DT_RIGHT = &H2&
   DT_VCENTER = &H4&
   DT_BOTTOM = &H8&
   DT_WORDBREAK = &H10&
   DT_SINGLELINE = &H20&
   DT_EXPANDTABS = &H40&
   DT_TABSTOP = &H80&
   DT_NOCLIP = &H100&
   DT_EXTERNALLEADING = &H200&
   DT_CALCRECT = &H400&
   DT_NOPREFIX = &H800&
   DT_INTERNAL = &H1000&
   '#if(WINVER >= =&H0400)
   DT_EDITCONTROL = &H2000&
   DT_PATH_ELLIPSIS = &H4000&
   DT_END_ELLIPSIS = &H8000&
   DT_MODIFYSTRING = &H10000
   DT_RTLREADING = &H20000
   DT_WORD_ELLIPSIS = &H40000
End Enum

Public Const DFCS_INACTIVE = &H100
Public Const DFCS_PUSHED = &H200
Public Const DFCS_CHECKED = &H400

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadID As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Const WH_MSGFILTER As Long = (-1)
Public Const WH_KEYBOARD As Long = 2
Public Const MSGF_MENU = 2
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8
Public Const COLORONCOLOR = 3

Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const FF_DONTCARE = 0
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_CHARSET = 1
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Const CLR_INVALID = -1
' Corrected Draw State function declarations:
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
   ByVal hBrush As Long, ByVal lpDrawStateProc As Long, _
   ByVal lParam As Long, ByVal wParam As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal cx As Long, ByVal cy As Long, _
   ByVal fuFlags As Long) As Long

' Missing Draw State constants declarations:
'/* Image type */
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

Public Enum DFCCaptionTypeFlags
   ' Caption types:
   DFCS_CAPTIONCLOSE = &H0&
   DFCS_CAPTIONMIN = &H1&
   DFCS_CAPTIONMAX = &H2&
   DFCS_CAPTIONRESTORE = &H3&
   DFCS_CAPTIONHELP = &H4&
End Enum
Public Enum DFCMenuTypeFlags
   ' Menu types:
   DFCS_MENUARROW = &H0&
   DFCS_MENUCHECK = &H1&
   DFCS_MENUBULLET = &H2&
   DFCS_MENUARROWRIGHT = &H4&
End Enum
Public Enum DFCScrollTypeFlags
   ' Scroll types:
   DFCS_SCROLLUP = &H0&
   DFCS_SCROLLDOWN = &H1&
   DFCS_SCROLLLEFT = &H2&
   DFCS_SCROLLRIGHT = &H3&
   DFCS_SCROLLCOMBOBOX = &H5&
   DFCS_SCROLLSIZEGRIP = &H8&
   DFCS_SCROLLSIZEGRIPRIGHT = &H10&
End Enum
Public Enum SystemMetricsIndexConstants
   SM_CMETRICS = 44&
   SM_CMOUSEBUTTONS = 43&
   SM_CXBORDER = 5&
   SM_CXCURSOR = 13&
   SM_CXDLGFRAME = 7&
   SM_CXDOUBLECLK = 36&
   SM_CXFIXEDFRAME = SM_CXDLGFRAME
   SM_CXFRAME = 32&
   SM_CXFULLSCREEN = 16&
   SM_CXHSCROLL = 21&
   SM_CXHTHUMB = 10&
   SM_CXICON = 11&
   SM_CXICONSPACING = 38&
   SM_CXMIN = 28&
   SM_CXMINTRACK = 34&
   SM_CXSCREEN = 0&
   SM_CXSIZE = 30&
   SM_CXSIZEFRAME = SM_CXFRAME
   SM_CXSMSIZE = 30
   SM_CXVSCROLL = 2&
   SM_CYBORDER = 6&
   SM_CYCAPTION = 4&
   SM_CYCURSOR = 14&
   SM_CYDLGFRAME = 8&
   SM_CYDOUBLECLK = 37&
   SM_CYFIXEDFRAME = SM_CYDLGFRAME
   SM_CYFRAME = 33&
   SM_CYFULLSCREEN = 17&
   SM_CYHSCROLL = 3&
   SM_CYICON = 12&
   SM_CYICONSPACING = 39&
   SM_CYKANJIWINDOW = 18&
   SM_CYMENU = 15&
   SM_CYMIN = 29&
   SM_CYMINTRACK = 35&
   SM_CYSCREEN = 1&
   SM_CYSIZE = 31&
   SM_CYSIZEFRAME = SM_CYFRAME
   SM_CYSMSIZE = 31
   SM_CYVSCROLL = 20&
   SM_CYVTHUMB = 9&
   SM_DBCSENABLED = 42&
   SM_DEBUG = 22&
   SM_MENUDROPALIGNMENT = 40&
   SM_MOUSEPRESENT = 19&
   SM_PENWINDOWS = 41&
   SM_SWAPBUTTON = 23&
End Enum
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long

Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Public Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long

Public Declare Function AppendMenuBylong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function AppendMenuByString Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function InsertMenuByString Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function HiliteMenuItem Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long

Public Declare Function MenuItemFromPoint Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
Public Declare Function TrackPopupMenuByLong Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Type TPMPARAMS
   cbSize               As Long
   rcExclude            As RECT
End Type
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long

' =======================================================================
' General API Declares:
' =======================================================================
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Const GW_CHILD = 5
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA As Long = 48&
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' Create a new icon based on an image list icon:
Private Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal diIgnore As Long) As Long
' Draw an item in an ImageList:
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, _
   ByVal i As Long, ByVal hdcDst As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal fStyle As Long) As Long
' Draw an item in an ImageList with more control over positioning
' and colour:
Private Declare Function ImageList_DrawEx Lib "COMCTL32.DLL" (ByVal hIml As Long, _
   ByVal i As Long, ByVal hdcDst As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal dx As Long, ByVal dy As Long, _
   ByVal rgbBk As Long, ByVal rgbFg As Long, _
   ByVal fStyle As Long) As Long
' Built in ImageList drawing methods:
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_OVERLAYMASK = 3840
' Use default rgb colour:
Public Const CLR_NONE = -1

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function LoadImageByNum Lib "user32" Alias "LoadImageA" (ByVal hinst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Type MEASUREITEMSTRUCT
   CtlType              As Long
   CtlID                As Long
   itemID               As Long
   itemWidth            As Long
   itemHeight           As Long
   ItemData             As Long
End Type

Public Type DRAWITEMSTRUCT
   CtlType              As Long
   CtlID                As Long
   itemID               As Long
   itemAction           As Long
   itemState            As Long
   hwndItem             As Long
   hdc                  As Long
   rcItem               As RECT
   ItemData             As Long
End Type

Public Type MENUITEMINFO
   cbSize               As Long
   fMask                As Long
   fType                As Long
   fState               As Long
   wID                  As Long
   hSubMenu             As Long
   hbmpChecked          As Long
   hbmpUnchecked        As Long
   dwItemData           As Long
   dwTypeData           As String
   cch                  As Long
End Type

Public Type MENUITEMTEMPLATE
   mtOption             As Integer
   mtID                 As Integer
   mtString             As Byte
End Type

Public Type MENUITEMTEMPLATEHEADER
   versionNumber        As Integer
   offset               As Integer
End Type
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&

Public Const TPM_NONOTIFY = &H80&           '/* Don't send any notification msgs */
Public Const TPM_RETURNCMD = &H100
Public Const TPM_HORIZONTAL = &H0          '/* Horz alignment matters more */
Public Const TPM_VERTICAL = &H40           '/* Vert alignment matters more */

Public Type tMenuItem
   sHelptext            As String
   sInputCaption        As String
   sCaption             As String
   sAccelerator         As String
   sShortCutDisplay     As String
   iShortCutShiftMask   As Integer
   iShortCutShiftKey    As Integer
   lID                  As Long
   lActualID            As Long       ' The ID gets modified if we add a sub-menu to the hMenu of the popup
   lItemData            As Long
   lIndex               As Long
   lParentId            As Long
   lIconIndex           As Long
   bChecked             As Boolean
   bRadioCheck          As Boolean
   bEnabled             As Boolean
   hMenu                As Long
   lHeight              As Long
   lWidth               As Long
   bCreated             As Boolean
   bIsAVBMenu           As Boolean
   lShortCutStartPos    As Long
   bMarkTODestroy       As Boolean
   sKey                 As String
   lParentIndex         As Long
   bTitle               As Boolean
   bDefault             As Boolean
   bOwnerDraw           As Boolean
   bMenuBarBreak        As Boolean
   bMenuBreak           As Boolean
End Type
' Menu flag constants:
' Menu flag constants:
Public Const MF_APPEND = &H100&
Public Const MF_BITMAP = &H4&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_CALLBACKS = &H8000000
Public Const MF_CHANGE = &H80&
Public Const MF_CHECKED = &H8&
Public Const MF_CONV = &H40000000
Public Const MF_DELETE = &H200&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_END = &H80
Public Const MF_ERRORS = &H10000000
Public Const MF_GRAYED = &H1&
Public Const MF_HELP = &H4000&
Public Const MF_HILITE = &H80&
Public Const MF_HSZ_INFO = &H1000000
Public Const MF_INSERT = &H0&
Public Const MF_LINKS = &H20000000
Public Const MF_MASK = &HFF000000
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_POSTMSGS = &H4000000
Public Const MF_REMOVE = &H1000&
Public Const MF_SENDMSGS = &H2000000
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_SYSMENU = &H2000&
Public Const MF_UNCHECKED = &H0&
Public Const MF_UNHILITE = &H0&
Public Const MF_USECHECKBITMAPS = &H200&
Public Const MF_DEFAULT = &H1000&
Public Const MFT_RADIOCHECK = &H200&
Public Const MFT_RIGHTORDER = &H2000&
' New versions of the names...
Public Const MFS_GRAYED = &H3&
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = MF_CHECKED
Public Const MFS_HILITE = MF_HILITE
Public Const MFS_ENABLED = MF_ENABLED
Public Const MFS_UNCHECKED = MF_UNCHECKED
Public Const MFS_UNHILITE = MF_UNHILITE
Public Const MFS_DEFAULT = MF_DEFAULT

Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long

' Work DC
Private m_hdcMono       As Long
Private m_hbmpMono      As Long
Private m_hBmpOld       As Long

' Keyboard hook (for accelerators):
Private m_hKeyHook      As Long
Private m_lKeyHookPtr() As Long
Private m_iKeyHookCount As Long
Private m_iCurrentMessage As Long
Private m_iProcOld      As Long


Public Sub ImageListDrawIcon( _
   ByVal hdc As Long, _
   ByVal hIml As Long, _
   ByVal iIconIndex As Long, _
   ByVal lX As Long, _
   ByVal lY As Long, _
   Optional ByVal bSelected As Boolean = False, _
   Optional ByVal bBlend25 As Boolean = False _
   )
   Dim lFlags           As Long
   Dim lR               As Long

   lFlags = ILD_TRANSPARENT
   If (bSelected) Then
      lFlags = lFlags Or ILD_SELECTED
   End If
   If (bBlend25) Then
      lFlags = lFlags Or ILD_BLEND25
   End If
   lR = ImageList_Draw(hIml, iIconIndex, hdc, lX, lY, lFlags)
   If (lR = 0) Then
      Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
   End If
End Sub

Public Sub ImageListDrawIconDisabled(ByVal hdc As Long, _
   ByVal hIml As Long, ByVal iIconIndex As Long, _
   ByVal lX As Long, ByVal lY As Long, _
   ByVal lSize As Long)
   Dim lR               As Long
   Dim hIcon            As Long

   hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
   lR = DrawState(hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
   DestroyIcon hIcon

End Sub

Public Property Get LighterColour(ByVal oColor As OLE_COLOR) As Long
   Dim lC               As Long
   Dim h                As Single, s As Single, l As Single
   Dim lR               As Long, lG As Long, lB As Long
   Static s_lColLast As Long
   Static s_lLightColLast As Long

   lC = TranslateColor(oColor)
   If (lC <> s_lColLast) Then
      s_lColLast = lC
      RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
      If (l > 0.99) Then
         l = l * 0.8
      Else
         l = l * 1.25
         If (l > 1) Then
            l = 1
         End If
      End If
      HLSToRGB h, s, l, lR, lG, lB
      s_lLightColLast = RGB(lR, lG, lB)
   End If
   LighterColour = s_lLightColLast
End Property

Public Property Get NoPalette(Optional ByVal bForce As Boolean = False) As Boolean
   Static bOnce As Boolean
   Static bNoPalette As Boolean
   Dim lHDC             As Long
   Dim lBits            As Long
   If (bForce) Then
      bOnce = False
   End If
   If Not (bOnce) Then
      lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      If (lHDC <> 0) Then
         lBits = GetDeviceCaps(lHDC, BITSPIXEL)
         If (lBits <> 0) Then
            bOnce = True
         End If
         bNoPalette = (lBits > 8)
         DeleteDC lHDC
      End If
   End If
   NoPalette = bNoPalette
End Property

Private Sub RGBToHLS(ByVal r As Long, ByVal g As Long, ByVal b As Long, _
   h As Single, s As Single, l As Single)
   Dim Max              As Single
   Dim Min              As Single
   Dim delta            As Single
   Dim rR               As Single, rG As Single, rB As Single

   rR = r / 255: rG = g / 255: rB = b / 255

   '{Given: rgb each in [0,1].
   ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
   Max = Maximum(rR, rG, rB)
   Min = Minimum(rR, rG, rB)
   l = (Max + Min) / 2 '{This is the lightness}
   '{Next calculate saturation}
   If Max = Min Then
      'begin {Acrhomatic case}
      s = 0
      h = 0
      'end {Acrhomatic case}
   Else
      'begin {Chromatic case}
      '{First calculate the saturation.}
      If l <= 0.5 Then
         s = (Max - Min) / (Max + Min)
      Else
         s = (Max - Min) / (2 - Max - Min)
      End If
      '{Next calculate the hue.}
      delta = Max - Min
      If rR = Max Then
         h = (rG - rB) / delta '{Resulting color is between yellow and magenta}
      ElseIf rG = Max Then
         h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
      ElseIf rB = Max Then
         h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
      End If
      'end {Chromatic Case}
   End If
End Sub

Private Sub HLSToRGB(ByVal h As Single, ByVal s As Single, ByVal l As Single, _
   r As Long, g As Long, b As Long)
   Dim rR               As Single, rG As Single, rB As Single
   Dim Min              As Single, Max As Single

   If s = 0 Then
      ' Achromatic case:
      rR = l: rG = l: rB = l
   Else
      ' Chromatic case:
      ' delta = Max-Min
      If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = l * (1 - s)
      Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = l - s * (1 - l)
      End If
      ' Get the Max value:
      Max = 2 * l - Min

      ' Now depending on sector we can evaluate the h,l,s:
      If (h < 1) Then
         rR = Max
         If (h < 0) Then
            rG = Min
            rB = rG - h * (Max - Min)
         Else
            rB = Min
            rG = h * (Max - Min) + rB
         End If
      ElseIf (h < 3) Then
         rG = Max
         If (h < 2) Then
            rB = Min
            rR = rB - (h - 2) * (Max - Min)
         Else
            rR = Min
            rB = (h - 2) * (Max - Min) + rR
         End If
      Else
         rB = Max
         If (h < 4) Then
            rR = Min
            rG = rR - (h - 4) * (Max - Min)
         Else
            rG = Min
            rR = (h - 4) * (Max - Min) + rG
         End If

      End If

   End If
   r = rR * 255: g = rG * 255: b = rB * 255
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function

Public Sub ClearUpWorkDC()
   If m_hBmpOld <> 0 Then
      SelectObject m_hdcMono, m_hBmpOld
      m_hBmpOld = 0
   End If
   If m_hbmpMono <> 0 Then
      DeleteObject m_hbmpMono
      m_hbmpMono = 0
   End If
   If m_hdcMono <> 0 Then
      DeleteDC m_hdcMono
      m_hdcMono = 0
   End If
End Sub

Public Sub DrawMaskedFrameControl( _
   ByVal hdcDest As Long, _
   ByRef trWhere As RECT, _
   ByVal kind As DFCFlags, _
   ByVal Style As Long _
   )
   Dim hbrMenu          As Long
   Dim saveBkMode       As Long, saveBkColor As Long, saveBrush As Long
   Dim tRWhereOnTmp     As RECT

   Static s_lLastRight As Long, s_lLastBottom As Long

   With tRWhereOnTmp
      .right = trWhere.right - trWhere.left
      .bottom = trWhere.bottom - trWhere.top
      If .right > s_lLastRight Or .bottom > s_lLastBottom Or (m_hdcMono = 0) Or (m_hbmpMono = 0) Or (m_hBmpOld = 0) Then
         ClearUpWorkDC
         ' Create memory device context for our temporary mask
         m_hdcMono = CreateCompatibleDC(0)
         If m_hdcMono <> 0 Then
            ' Create monochrome bitmap and select it into DC
            m_hbmpMono = CreateCompatibleBitmap(m_hdcMono, .right, .bottom)
            If m_hbmpMono <> 0 Then
               m_hBmpOld = SelectObject(m_hdcMono, m_hbmpMono)
               SetBkColor m_hdcMono, &HFFFFFF
            End If
         End If
         If m_hBmpOld = 0 Then
            ' Failed...
            ClearUpWorkDC
         End If
      End If
      s_lLastRight = .right
      s_lLastBottom = .bottom
   End With

   DrawFrameControl m_hdcMono, tRWhereOnTmp, kind, Style
   ' We have black where tick & white elsewhere
   SetBkColor hdcDest, &HFFFFFF
   BitBlt hdcDest, trWhere.left, trWhere.top, trWhere.right, trWhere.bottom, m_hdcMono, 0, 0, vbSrcAnd

   ' Clean up everything.
   If saveBrush <> 0 Then
      SelectObject hdcDest, saveBrush
   End If
   If hbrMenu <> 0 Then
      DeleteObject hbrMenu
   End If
   If saveBkMode <> 0 Then
      SetBkMode hdcDest, saveBkMode
   End If
   If saveBkColor <> 0 Then
      SetBkColor hdcDest, saveBkColor
   End If

End Sub

Public Sub DrawGradient( _
   ByVal hdc As Long, _
   ByRef rct As RECT, _
   ByVal lEndColour As Long, _
   ByVal lStartColour As Long, _
   ByVal bVertical As Boolean _
   )
   Dim lStep            As Long
   Dim lPos             As Long, lSize As Long
   Dim bRGB(1 To 3)     As Integer
   Dim bRGBStart(1 To 3) As Integer
   Dim dR(1 To 3)       As Double
   Dim dPos             As Double, d As Double
   Dim hBr              As Long
   Dim tR               As RECT

   LSet tR = rct
   If bVertical Then
      lSize = (tR.bottom - tR.top)
   Else
      lSize = (tR.right - tR.left)
   End If
   lStep = lSize \ 255
   If (lStep < 3) Then
      lStep = 3
   End If

   bRGB(1) = lStartColour And &HFF&
   bRGB(2) = (lStartColour And &HFF00&) \ &H100&
   bRGB(3) = (lStartColour And &HFF0000) \ &H10000
   bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
   dR(1) = (lEndColour And &HFF&) - bRGB(1)
   dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
   dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)

   For lPos = lSize To 0 Step -lStep
      ' Draw bar:
      If bVertical Then
         tR.top = tR.bottom - lStep
      Else
         tR.left = tR.right - lStep
      End If
      If tR.top < rct.top Then
         tR.top = rct.top
      End If
      If tR.left < rct.left Then
         tR.left = rct.left
      End If

      'Debug.Print tR.Right, tR.left, (bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1))
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hdc, tR, hBr
      DeleteObject hBr

      ' Adjust colour:
      dPos = ((lSize - lPos) / lSize)
      If bVertical Then
         tR.bottom = tR.top
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      Else
         tR.right = tR.left
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      End If

   Next lPos

End Sub

Private Property Get PopupMenuFromPtr(ByVal lPtr As Long) As cPopupMenu
   Dim oTemp            As Object
   If lPtr <> 0 Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory oTemp, lPtr, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set PopupMenuFromPtr = oTemp
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory oTemp, 0&, 4
      ' OK, hit the End button if you must--you'll probably still crash,
      ' but it will be because of the subclass, not the uncounted reference
   End If
End Property

Private Function KeyboardFilter(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim bKeyUp           As Boolean
   Dim bAlt             As Boolean, bCtrl As Boolean, bShift As Boolean
   Dim bFKey            As Boolean, bEscape As Boolean, bDelete As Boolean
   Dim wMask            As KeyCodeConstants
   Dim i                As Long

   On Error GoTo ErrorHandler

   If nCode = HC_ACTION And m_iKeyHookCount > 0 Then
      ' Key up or down:
      bKeyUp = ((lParam And &H80000000) = &H80000000)
      If Not bKeyUp Then
         bShift = (GetAsyncKeyState(vbKeyShift) <> 0)
         bAlt = ((lParam And &H20000000) = &H20000000)
         bCtrl = (GetAsyncKeyState(vbKeyControl) <> 0)
         bFKey = ((wParam >= vbKeyF1) And (wParam <= vbKeyF12))
         bEscape = (wParam = vbKeyEscape)
         bDelete = (wParam = vbKeyDelete)
         If bAlt Or bCtrl Or bFKey Or bEscape Or bDelete Then
            wMask = Abs(bShift * vbShiftMask) Or Abs(bCtrl * vbCtrlMask) Or Abs(bAlt * vbAltMask)
            For i = m_iKeyHookCount To 1 Step -1
               If m_lKeyHookPtr(i) <> 0 Then
                  ' Alt- or Ctrl- key combination pressed:
                  ' #HOut# ********************
                  ' #HOut# Programmer Name  : removed
                  ' #HOut# Date             : 06/18/2001
                  ' #HOut# Time             : 10:57
                  ' #HOut# Comment          :
                  ' #HOut# Comment          :
                  ' #HOut# Comment          :
                  ' #HOut# ********************
                  ' #Out#                   Set cT = PopupMenuFromPtr(m_lKeyHookPtr(i))
                  ' #Out#                   If Not cT Is Nothing Then
                  ' #Out#                      If cT.AcceleratorPress(wParam, wMask) Then
                  ' #Out#                         KeyboardFilter = 1
                  ' #Out#                         Exit Function
                  ' #Out#                      End If
                  ' #Out#                   End If
                  ' #HOut# ********************
               End If
            Next i
         End If
      End If
   End If
   KeyboardFilter = CallNextHookEx(m_hKeyHook, nCode, wParam, lParam)

   Exit Function

ErrorHandler:
   Debug.Print "Keyboard Hook Error!"
   Exit Function

End Function

Public Sub AttachKeyboardHook(cThis As cPopupMenu)
   Dim lpFn             As Long
   Dim lPtr             As Long
   Dim i                As Long

   If m_iKeyHookCount = 0 Then
      lpFn = HookAddress(AddressOf KeyboardFilter)
      m_hKeyHook = SetWindowsHookEx(WH_KEYBOARD, lpFn, 0&, GetCurrentThreadId())
      Debug.Assert (m_hKeyHook <> 0)
   End If
   lPtr = ObjPtr(cThis)
   For i = 1 To m_iKeyHookCount
      If lPtr = m_lKeyHookPtr(i) Then
         ' we already have it:
         Debug.Assert False
         Exit Sub
      End If
   Next i
   ReDim Preserve m_lKeyHookPtr(1 To m_iKeyHookCount + 1) As Long
   m_iKeyHookCount = m_iKeyHookCount + 1
   m_lKeyHookPtr(m_iKeyHookCount) = lPtr

End Sub

Public Sub DetachKeyboardHook(cThis As cPopupMenu)
   Dim i                As Long
   Dim lPtr             As Long
   Dim iThis            As Long

   lPtr = ObjPtr(cThis)
   For i = 1 To m_iKeyHookCount
      If m_lKeyHookPtr(i) = lPtr Then
         iThis = i
         Exit For
      End If
   Next i
   If iThis <> 0 Then
      If m_iKeyHookCount > 1 Then
         For i = iThis To m_iKeyHookCount - 1
            m_lKeyHookPtr(i) = m_lKeyHookPtr(i + 1)
         Next i
      End If
      m_iKeyHookCount = m_iKeyHookCount - 1
      If m_iKeyHookCount >= 1 Then
         ReDim Preserve m_lKeyHookPtr(1 To m_iKeyHookCount) As Long
      Else
         Erase m_lKeyHookPtr
      End If
   Else
      ' Trying to detach a toolbar which was never attached...
      ' This will happen at design time
   End If

   If m_iKeyHookCount <= 0 Then
      If (m_hKeyHook <> 0) Then
         UnhookWindowsHookEx m_hKeyHook
         m_hKeyHook = 0
      End If
   End If

End Sub

Public Function HookAddress(ByVal lPtr As Long) As Long
   HookAddress = lPtr
End Function

Public Sub TileArea( _
   ByVal hdc As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal Width As Long, _
   ByVal Height As Long, _
   ByVal lSrcDC As Long, _
   ByVal lBitmapW As Long, _
   ByVal lBitmapH As Long, _
   ByVal lSrcOffsetX As Long, _
   ByVal lSrcOffsetY As Long _
   )
   Dim lSrcX            As Long
   Dim lSrcY            As Long
   Dim lSrcStartX       As Long
   Dim lSrcStartY       As Long
   Dim lSrcStartWidth   As Long
   Dim lSrcStartHeight  As Long
   Dim lDstX            As Long
   Dim lDstY            As Long
   Dim lDstWidth        As Long
   Dim lDstHeight       As Long

   lSrcStartX = ((x + lSrcOffsetX) Mod lBitmapW)
   lSrcStartY = ((y + lSrcOffsetY) Mod lBitmapH)
   lSrcStartWidth = (lBitmapW - lSrcStartX)
   lSrcStartHeight = (lBitmapH - lSrcStartY)
   lSrcX = lSrcStartX
   lSrcY = lSrcStartY

   lDstY = y
   lDstHeight = lSrcStartHeight

   Do While lDstY < (y + Height)
      If (lDstY + lDstHeight) > (y + Height) Then
         lDstHeight = y + Height - lDstY
      End If
      lDstWidth = lSrcStartWidth
      lDstX = x
      lSrcX = lSrcStartX
      Do While lDstX < (x + Width)
         If (lDstX + lDstWidth) > (x + Width) Then
            lDstWidth = x + Width - lDstX
            If (lDstWidth = 0) Then
               lDstWidth = 4
            End If
         End If
         'If (lDstWidth > Width) Then lDstWidth = Width
         'If (lDstHeight > Height) Then lDstHeight = Height
         BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDC, lSrcX, lSrcY, vbSrcCopy
         lDstX = lDstX + lDstWidth
         lSrcX = 0
         lDstWidth = lBitmapW
      Loop
      lDstY = lDstY + lDstHeight
      lSrcY = 0
      lDstHeight = lBitmapH
   Loop
End Sub

Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateColor) Then
      TranslateColor = CLR_INVALID
   End If
End Function

