VERSION 5.00
Begin VB.UserControl VBRichEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "VBRichEdit.ctx":0000
End
Attribute VB_Name = "VBRichEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' cRICHEDIT
' Steve McMahon 1998
' 14 June 1998
'
' An ultra-light RichEdit control all in VB for display only purposes
' ======================================================================

' ======================================================================
' Declares and types:
' ======================================================================
' Windows general:

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const SW_HIDE = 0
Private Const WS_CHILD = &H40000000
Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000
Private Const WS_VISIBLE = &H10000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_BORDER = &H800000
Private Const WS_EX_CLIENTEDGE = &H200
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_DLGFRAME = &H400000
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_COPY = &H301

Public Enum ERECControlVersion
    eRICHED32
    eRICHED20
End Enum

Private m_hWnd As Long
Private m_bSubClassing As Boolean
Private m_hLib As Long
Private m_eVersion As ERECControlVersion

Public Enum ERECFileTypes
    SF_TEXT = &H1
    SF_RTF = &H2
End Enum

Implements ISubclass
Private m_emr As EMsgResponse

Public Sub SelectAll()
Dim tc As CHARRANGE
    tc.cpMax = -1
    tc.cpMin = 0
    SendMessage m_hWnd, EM_EXSETSEL, 0, tc
End Sub
Public Sub SelectNone()
Dim tc As CHARRANGE
    tc.cpMax = 0
    tc.cpMin = 0
    SendMessage m_hWnd, EM_EXSETSEL, 0, tc
End Sub
Public Property Get CharacterCount() As Long
    CharacterCount = SendMessageLong(m_hWnd, WM_GETTEXTLENGTH, 0, 0)
End Property

Public Sub Copy()
    SendMessageLong m_hWnd, WM_COPY, 0, 0
End Sub

Public Property Let UseVersion(ByVal eVersion As ERECControlVersion)
    If (UserControl.Ambient.UserMode) Then
        ' can't set at run time in this implementation.
    Else
        m_eVersion = eVersion
    End If
End Property
Public Property Get UseVersion() As ERECControlVersion
    UseVersion = m_eVersion
End Property

Public Property Let Contents(ByVal eType As ERECFileTypes, ByRef sContents As String)
Dim tStream As EDITSTREAM

    tStream.dwCookie = m_hWnd
    tStream.pfnCallback = plAddressOf(AddressOf LoadCallBack)
    tStream.dwError = 0
    StreamText = sContents
    ' The text will be streamed in though the LoadCallback function:
    SendMessage m_hWnd, EM_STREAMIN, eType, tStream
    
End Property
Public Property Get Contents(ByVal eType As ERECFileTypes) As String
Dim tStream As EDITSTREAM
    
    tStream.dwCookie = m_hWnd
    tStream.pfnCallback = plAddressOf(AddressOf SaveCallBack)
    tStream.dwError = 0
    ' The text will be streamed out though the SaveCallback function:
    ClearStreamText
    SendMessage m_hWnd, EM_STREAMOUT, eType, tStream
    Contents = StreamText()
    
End Property
Public Sub PrintDoc(ByVal sDocTitle As String)
Dim fr As FORMATRANGE
Dim di As DOCINFO
Dim lTextOut As Long, lTextAmt As Long
Dim pd As PRINTDLG
Dim b() As Byte
Dim hJob As Long
Dim lR As Long
Dim lWidthTwips As Long, lHeightTwips As Long

'// Initialize the PRINTDLG structure.
pd.lStructSize = Len(pd)
pd.hWndOwner = m_hWnd
pd.hDevMode = 0
pd.hDevNames = 0
pd.nFromPage = 0
pd.nToPage = 0
pd.nMinPage = 0
pd.nMaxPage = 0
pd.nCopies = 0
pd.hInstance = App.hInstance
pd.flags = PD_RETURNDC Or PD_NOPAGENUMS Or PD_NOSELECTION Or PD_PRINTSETUP
pd.lpfnSetupHook = 0
pd.lpSetupTemplateName = 0
pd.lpfnPrintHook = 0
pd.lpPrintTemplateName = 0

'// Get the printer DC.
If (PRINTDLG(pd) <> 0) Then

    '// Be sure that the printer DC is in text mode.
    SetMapMode pd.hdc, MM_TEXT

    '// Fill out the FORMATRANGE structure for the RTF output.
    fr.hdc = pd.hdc '; // HDC
    fr.hdcTarget = fr.hdc
    fr.chrg.cpMin = 0 '; // print
    fr.chrg.cpMax = -1 '; // entire contents
    fr.rcPage.Top = 0: fr.rc.Left = 0: fr.rcPage.Left = 0: fr.rc.Top = 0
    ' This is the number of pixels across (INCORRECT IN 'PROGRAMMING WIN95 INTERFACE'):
    lWidthTwips = GetDeviceCaps(pd.hdc, HORZRES)
    ' Now divide by the number of pixels per logical inch and multiply by 1440 (twips/inch):
    lWidthTwips = (lWidthTwips * 1440) / GetDeviceCaps(pd.hdc, LOGPIXELSX)
    fr.rcPage.Right = lWidthTwips
    fr.rc.Right = fr.rcPage.Right
    ' This is the number of pixels down (INCORRECT IN 'PROGRAMMING WIN95 INTERFACE'):
    lHeightTwips = GetDeviceCaps(pd.hdc, VERTRES) * 4
    ' Now divide by the number of pixels per logical inch and multiply by 1440 (twips/inch):
    lHeightTwips = (lHeightTwips * 1440) / GetDeviceCaps(pd.hdc, LOGPIXELSY)
    fr.rcPage.Bottom = lHeightTwips
    fr.rc.Bottom = fr.rcPage.Bottom

    '// Fill out the DOCINFO structure.
    di.cbSize = Len(di)
    If (sDocTitle = "") Then
        sDocTitle = "RTF Document"
    End If
    ReDim b(0 To Len(sDocTitle) - 1) As Byte
    b = StrConv(sDocTitle, vbFromUnicode)
    di.lpszDocName = VarPtr(b(0))
    di.lpszOutput = 0

    hJob = StartDoc(pd.hdc, di)
    If (hJob <> 0) Then
        StartPage pd.hdc
    
        lTextOut = 0
        lTextAmt = SendMessage(m_hWnd, WM_GETTEXTLENGTH, 0, 0)
    
        Do While (lTextOut < lTextAmt)
            lTextOut = SendMessage(m_hWnd, EM_FORMATRANGE, True, fr)
            If (lTextOut < lTextAmt) Then
                EndPage pd.hdc
                StartPage pd.hdc
                fr.chrg.cpMin = lTextOut
                fr.chrg.cpMax = -1
            End If
        Loop
    
        '// Reset the formatting of the rich edit control.
        SendMessageLong m_hWnd, EM_FORMATRANGE, True, 0
    
        '// Finish the document.
        EndPage pd.hdc
        EndDoc pd.hdc
    Else
        Debug.Print "Failed to start print job"
    End If
    
    '// Delete the printer DC.
    DeleteDC pd.hdc
End If

End Sub

Private Function plAddressOf(ByVal lAddr As Long) As Long
    ' Why do we have to write nonsense like this?
    plAddressOf = lAddr
End Function
Public Property Get hwnd() As Long
    hwnd = m_hWnd
End Property

Private Sub pInitialise()
Dim dwStyle As Long
Dim lS As Long
Dim hP As Long
Dim sLib As String
Dim sClass As String

    pTerminate

    If (UserControl.Ambient.UserMode) Then
        If (m_eVersion = eRICHED20) Then
            sLib = "RICHED20.DLL"
            sClass = RICHEDIT_CLASSA
        Else
            sLib = "RICHED32.DLL"
            sClass = RICHEDIT_CLASS10A
        End If
        m_hLib = LoadLibrary(sLib)
        If m_hLib <> 0 Then
            dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS
            dwStyle = dwStyle Or WS_HSCROLL Or WS_VSCROLL
            dwStyle = dwStyle Or ES_MULTILINE Or ES_SAVESEL Or ES_SUNKEN
            dwStyle = dwStyle Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
        
        '// Create the rich edit control.
            m_hWnd = CreateWindowEX( _
                WS_EX_CLIENTEDGE, _
                sClass, _
                "", _
                dwStyle, _
                0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, _
                UserControl.hwnd, _
                0, _
                App.hInstance, _
                0)
            If (m_hWnd <> 0) Then
                
                SetParent m_hWnd, UserControl.hwnd
                MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1
                
                dwStyle = GetWindowLong(m_hWnd, GWL_STYLE)
                dwStyle = dwStyle And Not (WS_POPUP Or WS_DLGFRAME Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
                SetWindowLong m_hWnd, GWL_STYLE, dwStyle
                SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOMOVE
                
                dwStyle = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
                dwStyle = dwStyle And Not WS_EX_CLIENTEDGE
                SetWindowLong UserControl.hwnd, GWL_EXSTYLE, dwStyle
                SetWindowPos UserControl.hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOMOVE
                
                EnableWindow m_hWnd, 1
                
                'SetFocusAPI UserControl.hwnd
                pAttachMessages
            End If
        End If
    End If
End Sub
Private Function pTerminate()
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
    If (m_hLib <> 0) Then
        FreeLibrary m_hLib
        m_hLib = 0
    End If
End Function
Private Sub pAttachMessages()
    m_emr = emrPreprocess
    'AttachMessage Me, UserControl.hwnd, WM_NOTIFY
    m_bSubClassing = True
End Sub
Private Sub pDetachMessages()
    If (m_bSubClassing) Then
        'DetachMessage Me, UserControl.hwnd, WM_NOTIFY
        m_bSubClassing = False
    End If
End Sub


Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    RHS = m_emr
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' debug.print
End Function

Private Sub UserControl_Initialize()
    Debug.Print "RichEditControl:Initialise"
    m_eVersion = eRICHED32
End Sub

Private Sub UserControl_InitProperties()
    pInitialise
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UseVersion = PropBag.ReadProperty("Version", eRICHED32)
    pInitialise
End Sub

Private Sub UserControl_Resize()
    If (m_hWnd <> 0) Then
        MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1
    End If
End Sub

Private Sub UserControl_Terminate()
    Debug.Print "RichEditControl:Terminate"
    pTerminate
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Version", UseVersion, eRICHED32
End Sub
