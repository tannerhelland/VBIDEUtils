VERSION 5.00
Begin VB.UserControl cReBar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   525
   ScaleWidth      =   4905
   Begin VB.Label lblRebar 
      Caption         =   "'Rebar Control'"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "cReBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ==============================================================================
' Declares, constants and types required for toolbar:
' ==============================================================================
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
'WINDOW STYLES
Private Const WS_BORDER = &H800000
Private Const WS_CHILD = &H40000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_VISIBLE = &H10000000
'EXTENDED WINDOW STYLES
Private Const WS_EX_TOOLWINDOW = &H80
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Integer) As Long
Private Const GW_CHILD = 5
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_WINDOWTEXT = 8
Private Const COLORONCOLOR = 3
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

' General common controls:
Private Type CommonControlsEx
   dwSize As Long '// size of this structure
   dwICC As Long  '// flags indicating which classes to be initialized
End Type
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As CommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200& '// comboex
Private Const ICC_BAR_CLASSES = &H4
Private Const ICC_COOL_CLASSES = &H400
Private Const ICC_WIN95_CLASSES = &HFF&

Private Type nmhdr
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type REBARINFO
    cbSize As Integer
    fMask As Integer
    himl As Long
End Type

Private Type REBARBANDINFO
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As String
    cch As Long
    iImage As Long
    hWndChild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long
    wId As Long
End Type

Private Type REBARBANDINFO_NOTEXT
    cbSize As Long
    fMask As Long
    fStyle As Long
    clrFore As Long
    clrBack As Long
    lpText As Long
    cch As Long
    iImage As Integer 'Image
    hWndChild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long 'hBitmap
    wId As Long
End Type

Private Const CCS_TOP = &H1
Private Const CCS_NOMOVEY = &H2
Private Const CCS_BOTTOM = &H3
Private Const CCS_NORESIZE = &H4
Private Const CCS_NOPARENTALIGN = &H8
Private Const CCS_ADJUSTABLE = &H20
Private Const CCS_NODIVIDER = &H40
Private Const CCS_VERT = &H80
Private Const CCS_LEFT = (CCS_VERT Or CCS_TOP)
Private Const CCS_RIGHT = (CCS_VERT Or CCS_BOTTOM)
Private Const CCS_NOMOVEX = (CCS_VERT Or CCS_NOMOVEY)


Private Const REBARCLASSNAME = "ReBarWindow32"
Private Const RBN_FIRST = 0 - 831
Private Const RBN_LAST = 0 - 859
Private Const RBIM_IMAGELIST = &H1

'Rebar Styles
Private Const RBS_AUTOSIZE = &H2000
Private Const RBS_VERTICALGRIPPER = &H4000 '  // this always has the vertical gripper (default for horizontal mode)
Private Const RBS_TOOLTIPS = &H100
Private Const RBS_VARHEIGHT = &H200
Private Const RBS_BANDBORDERS = &H400
Private Const RBS_FIXEDORDER = &H800

Private Const RBBS_VARIABLEHEIGHT = &H40
Private Const RBBS_GRIPPERALWAYS = &H80      ' always show the gripper
Private Const RBBS_BREAK = &H1               ' break to new line
Private Const RBBS_FIXEDSIZE = &H2           ' band can't be sized
Private Const RBBS_CHILDEDGE = &H4           ' edge around top & bottom of child window
Private Const RBBS_HIDDEN = &H8              ' don't show
Private Const RBBS_NOVERT = &H10             ' don't show when vertical
Private Const RBBS_FIXEDBMP = &H20           ' bitmap doesn't move during band resize

Private Const RBBIM_STYLE = &H1
Private Const RBBIM_COLORS = &H2
Private Const RBBIM_TEXT = &H4
Private Const RBBIM_IMAGE = &H8
Private Const RBBIM_CHILD = &H10
Private Const RBBIM_CHILDSIZE = &H20
Private Const RBBIM_SIZE = &H40
Private Const RBBIM_BACKGROUND = &H80
Private Const RBBIM_ID = &H100
Private Const WM_USER = &H400
Private Const RB_BEGINDRAG = (WM_USER + 24)
Private Const RB_ENDDRAG = (WM_USER + 25)
Private Const RB_DRAGMOVE = (WM_USER + 26)
Private Const RB_HITTEST = (WM_USER + 8)
Private Const RB_INSERTBANDA = (WM_USER + 1)
Private Const RB_DELETEBAND = (WM_USER + 2)
Private Const RB_GETBARINFO = (WM_USER + 3)
Private Const RB_SETBARINFO = (WM_USER + 4)
Private Const RB_GETBANDINFO = (WM_USER + 5)
Private Const RB_SETBANDINFOA = (WM_USER + 6)
Private Const RB_SETPARENT = (WM_USER + 7)
Private Const RB_INSERTBANDW = (WM_USER + 10)
Private Const RB_SETBANDINFOW = (WM_USER + 11)
Private Const RB_GETBANDCOUNT = (WM_USER + 12)
Private Const RB_GETROWCOUNT = (WM_USER + 13)
Private Const RB_GETROWHEIGHT = (WM_USER + 14)
Private Const RB_SETBKCOLOR = (WM_USER + 19)
Private Const RB_GETBKCOLOR = (WM_USER + 20)
Private Const RB_SETTEXTCOLOR = (WM_USER + 21)
Private Const RB_GETTEXTCOLOR = (WM_USER + 22)
Private Const RBHT_NOWHERE = &H1
Private Const RBHT_CAPTION = &H2
Private Const RBHT_CLIENT = &H3
Private Const RBHT_GRABBER = &H4

Private Const RB_INSERTBAND = RB_INSERTBANDA
Private Const RB_SETBANDINFO = RB_SETBANDINFOA
Private Const RBN_HEIGHTCHANGE = (RBN_FIRST - 0)

Private Const WM_NOTIFY = &H4E&

Implements ISubclass
Private m_emr As EMsgResponse

Private m_hWnd As Long
'Private m_HwndControl As Long
Private m_HwndParent As Long
Private m_bSubClassing As Boolean
Private m_pic As StdPicture
Private m_bPictureLoaded As Boolean
Private m_sPicture As String
Private m_bInTerminate As Boolean

Private m_iWndItemCount As Integer
Private m_hWndItem() As Long
Private m_hWNdItemParent() As Long

Private m_wID() As Long
Private m_vData() As Variant
Private m_lIDCount As Long

Public Event HeightChanged(lNewHeight As Long)

Private Sub pCreateSubClass()
    If Not (m_bSubClassing) Then
        If (m_HwndParent = 0) Then
            If (UserControl.Ambient.UserMode) Then
                m_HwndParent = UserControl.Parent.hWnd
            End If
        End If
        If (m_HwndParent > 0) Then
            'Debug.Print "Subclassing window: " & m_HwndParent
            AttachMessage Me, m_HwndParent, WM_NOTIFY
            m_bSubClassing = True
        End If
    End If
End Sub

Private Sub pDestroySubClass()
    If (m_bSubClassing) Then
        DetachMessage Me, m_HwndParent, WM_NOTIFY
        m_HwndParent = 0
        m_bSubClassing = False
    End If
End Sub


Private Function ISubclass_WindowProc(ByVal hWnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
Dim lHeight As Long
   'Debug.Print "Rebar:Got WindowProc"
   If Not (m_bInTerminate) Then
      'If (wParam = 1000) Then
         Dim tNMH As nmhdr
         CopyMemory tNMH, ByVal lParam, Len(tNMH)
         If (tNMH.code = -831) Then
            RaiseEvent HeightChanged(lHeight)
            m_emr = emrPostProcess
         End If
      'End If
   End If

End Function
' Interface properties
Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = m_emr
End Property
Private Property Let ISubclass_MsgResponse(ByVal emrA As EMsgResponse)
    m_emr = emrA
End Property
Property Get BandChildMinHeight(ByVal lBand As Long) As Long
Dim cy As Long
    If (lBand >= 0) And (lBand < BandCount) Then
        If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_CHILDSIZE, cyMinChild:=cy)) Then
            BandChildMinHeight = cy
        End If
    Else
        BandChildMinHeight = -1
    End If
End Property
Property Let BandChildMinHeight(ByVal lBand As Long, lHeight As Long)
    If (lBand >= 0) And (lBand < BandCount) Then
        Dim tRbbi As REBARBANDINFO_NOTEXT
        Dim lR As Long
        tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD
        tRbbi.cbSize = Len(tRbbi)
        lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
        If (lR <> 0) Then
            If (tRbbi.hWndChild <> 0) Then
                tRbbi.fMask = RBBIM_CHILDSIZE
                tRbbi.cyMinChild = lHeight
                lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
            End If
        End If
    Else
    End If
End Property
Property Get BandChildMinWidth(ByVal lBand As Long) As Long
Dim cx As Long
    If (lBand >= 0) And (lBand < BandCount) Then
        If (pbGetBandInfo(m_hWnd, lBand, fMask:=RBBIM_CHILDSIZE, cxMinChild:=cx)) Then
            BandChildMinWidth = cx
        End If
    Else
        BandChildMinWidth = -1
    End If

End Property
Property Let BandChildMinWidth(ByVal lBand As Long, lWidth As Long)
    If (lBand >= 0) And (lBand < BandCount) Then
        Dim tRbbi As REBARBANDINFO_NOTEXT
        Dim lR As Long
        tRbbi.fMask = RBBIM_CHILDSIZE Or RBBIM_CHILD
        tRbbi.cbSize = Len(tRbbi)
        lR = SendMessage(m_hWnd, RB_GETBANDINFO, lBand, tRbbi)
        If (lR <> 0) Then
            If (tRbbi.hWndChild <> 0) Then
                tRbbi.fMask = RBBIM_CHILDSIZE
                tRbbi.cxMinChild = lWidth
                lR = SendMessage(m_hWnd, RB_SETBANDINFOA, lBand, tRbbi)
            End If
        End If
    Else
    End If
End Property
Property Get BandCount() As Long
    BandCount = SendMessage(m_hWnd, RB_GETBANDCOUNT, 0&, ByVal 0&)
End Property

Private Function pbGetBandInfo( _
        ByVal lHwnd As Long, _
        ByVal lBand As Long, _
        Optional ByRef fMask As Long, _
        Optional ByRef fStyle As Long, _
        Optional ByRef clrFore As Long, _
        Optional ByRef clrBack As Long, _
        Optional ByRef cch As Long, _
        Optional ByRef iImage As Integer, _
        Optional ByRef hWndChild As Long, _
        Optional ByRef cxMinChild As Long, _
        Optional ByRef cyMinChild As Long, _
        Optional ByRef cx As Long, _
        Optional ByRef hbmpBack As Long, _
        Optional ByRef wId As Long _
    ) As Boolean
Dim lParam As Long
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim lR As Long

    tRbbi.cbSize = LenB(tRbbi)
    tRbbi.fMask = fMask
    lR = SendMessage(lHwnd, RB_GETBANDINFO, lBand, tRbbi)
    If (lR <> 0) Then
        With tRbbi
            fMask = .fMask
            fStyle = .fStyle
            clrFore = .clrFore
            clrBack = .clrBack
            cch = .cch
            iImage = .iImage
            hWndChild = .hWndChild
            cxMinChild = .cxMinChild
            cyMinChild = .cyMinChild
            cx = .cx
            hbmpBack = .hbmBack
            wId = .wId
        End With
        pbGetBandInfo = True
    End If
End Function

Public Property Get BackgroundBitmap() As String
    BackgroundBitmap = m_sPicture
End Property
Public Property Let BackgroundBitmap( _
        ByVal sBitmap As String _
    )
Dim sError As String
On Error GoTo BackgroundBitmapError
    
    ClearPicture
    'If (pbLoadTransPic(m_pic, sBitmap, sError)) Then
    '    m_sPicture = sBitmap
    '    m_bPictureLoaded = True
    'Else
        Debug.Print "Failed to load bitmap:" & sError
    'End If
    
    Exit Property
    
BackgroundBitmapError:
    sError = Err.Description
    Debug.Print "Background bitmap error!" & sError
    Exit Property
End Property

Public Function AddBandByHwnd( _
        ByVal hWnd As Long, _
        Optional ByVal sBandText As String = "", _
        Optional ByVal bBreakLine As Boolean = True, _
        Optional ByVal bFixedSize As Boolean = False, _
        Optional ByVal vData As Variant _
    ) As Long
Dim hBMP As Long
Dim lX As Long
Dim lBand As Long
Dim hWndP As Long
Dim wId As Long
    
    If (m_hWnd = 0) Then
        
    End If
    
    If (m_hWnd <> 0) Then
        If (m_bPictureLoaded) Then
            hBMP = m_pic.Handle
        Else
            hBMP = 0
        End If
        
        hWndP = GetParent(hWnd)
        If (hWndP <> 0) Then
            pAddWnds hWnd, hWndP
        End If
        wId = plAddId(vData)
        If (Not (pbRBAddBandByhWnd(m_hWnd, wId, hWnd, sBandText, hBMP, bBreakLine, bFixedSize, lBand))) Then
            Debug.Print "Failed to add Band"
            pRemoveID wId
        Else
            AddBandByHwnd = wId
            If Not (m_bSubClassing) Then
                ' Start subclassing:
                'Debug.Print "Start subclassing"
                pCreateSubClass
            End If
            RebarSize
        End If
    End If
End Function
Private Function pbRBAddBandByhWnd( _
        ByVal hWndRebar As Long, _
        ByVal wId As Long, _
        Optional ByVal hWndChild As Long = 0, _
        Optional ByVal sBandText As String = "", _
        Optional ByVal hBMP As Long = 0, _
        Optional ByVal bBreakLine As Boolean = True, _
        Optional ByVal bFixedSize As Boolean = False, _
        Optional ByRef ltRBand As Long _
    ) As Boolean

If hWndRebar = 0 Then
    MsgBox "No hWndRebar!"
    Exit Function
End If

Dim sClassName As String
Dim hWndReal As Long
Dim tRBand As REBARBANDINFO
Dim tRBandNT As REBARBANDINFO_NOTEXT
Dim bNoText As Boolean
Dim rct As RECT
Dim fMask As Long
Dim fStyle As Long

   hWndReal = hWndChild
   
   If Not (hWndChild = 0) Then
      'Check to see if it's a toolbar (so we can
      'make if flat)
      fMask = RBBIM_CHILD Or RBBIM_CHILDSIZE
      sClassName = Space$(255)
      GetClassName hWndChild, sClassName, 255
      'see if it's a real Windows toolbar
      If InStr(UCase$(sClassName), "TOOLBARWINDOW32") Then
          SetWindowLong hWndChild, GWL_STYLE, 1442875725
      End If
      'Could be a VB Toolbar -- make it flat anyway.
      If InStr(UCase$(sClassName), "TOOLBARWNDCLASS") Then
          hWndReal = GetWindow(hWndChild, GW_CHILD)
          SetWindowLong hWndReal, GWL_STYLE, 1442875725
      End If
   End If
   
   GetWindowRect hWndReal, rct
   rct.Bottom = rct.Bottom + 2
   
   If hBMP <> 0 Then
       fMask = fMask Or RBBIM_BACKGROUND
   End If
   fMask = fMask Or RBBIM_STYLE Or RBBIM_ID Or RBBIM_COLORS Or RBBIM_SIZE
   If sBandText <> "" Then
      fMask = fMask Or RBBIM_TEXT
      tRBand.lpText = sBandText
      tRBand.cch = Len(sBandText)
   Else
      bNoText = True
   End If
   fStyle = RBBS_CHILDEDGE Or RBBS_FIXEDBMP
   If bBreakLine = True Then
       fStyle = fStyle Or RBBS_BREAK
   End If
   If bFixedSize = True Then
       fStyle = fStyle Or RBBS_FIXEDSIZE
   Else
       fStyle = fStyle And Not RBBS_FIXEDSIZE
   End If
   
   If (bNoText) Then
      With tRBandNT
         .fMask = fMask
         .fStyle = fStyle
         'Only set if there's a child window
         If hWndReal <> 0 Then
            .hWndChild = hWndReal
            .cxMinChild = rct.Right - rct.Left
            .cyMinChild = rct.Bottom - rct.Top
         End If
         'Set the rest OK
         .wId = wId
         .clrBack = GetSysColor(COLOR_BTNFACE)
         .clrFore = GetSysColor(COLOR_BTNTEXT)
         .cx = 200
         .hbmBack = hBMP
         'The length of the type
         .cbSize = LenB(tRBand)
      End With
      pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBandNT) <> 0)
   Else
      With tRBand
         .fMask = fMask
         .fStyle = fStyle
         'Only set if there's a child window
         If hWndReal <> 0 Then
            .hWndChild = hWndReal
            .cxMinChild = rct.Right - rct.Left
            .cyMinChild = rct.Bottom - rct.Top
         End If
         'Set the rest OK
         .wId = wId
         .clrBack = GetSysColor(COLOR_BTNFACE)
         .clrFore = GetSysColor(COLOR_BTNTEXT)
         .cx = 200
         .hbmBack = hBMP
         'The length of the type
         .cbSize = LenB(tRBand)
      End With
      pbRBAddBandByhWnd = (SendMessage(hWndRebar, RB_INSERTBAND, -1, tRBand) <> 0)
   End If
   
   ltRBand = BandCount
   MoveWindow hWndRebar, 0, 0, 200, 10, True

End Function

Private Sub pRemoveID( _
        ByVal wId As Long _
    )
Dim lItem As Long
Dim lAt As Long
    
    For lItem = 1 To m_lIDCount
        If (m_wID(lItem) = wId) Then
            lAt = lItem
            Exit For
        End If
    Next lItem
    
    If (lAt > 0) Then
        If (m_lIDCount > 1) Then
            For lItem = lAt + 1 To m_lIDCount - 1
                m_wID(lItem) = m_wID(lItem + 1)
                m_vData(lItem) = m_vData(lItem + 1)
            Next lItem
            m_lIDCount = m_lIDCount - 1
            ReDim Preserve m_wID(1 To m_lIDCount) As Long
            ReDim Preserve m_vData(1 To m_lIDCount) As Variant
        Else
            m_lIDCount = 0
            Erase m_wID
            Erase m_vData
        End If
    End If
    
End Sub
Property Get BandIndexForId( _
        ByVal wId As Long _
    ) As Long
Dim lItem As Long
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim lIndex As Long
Dim lR As Long

    lIndex = -1
    tRbbi.cbSize = Len(tRbbi)
    tRbbi.fMask = RBBIM_ID
    For lItem = 0 To BandCount - 1
        lR = SendMessage(m_hWnd, RB_GETBANDINFO, lItem, tRbbi)
        If (lR <> 0) Then
            If (wId = tRbbi.wId) Then
                lIndex = lItem
                Exit For
            End If
        End If
    Next lItem
    BandIndexForId = lIndex
End Property
Property Get BandIndexForData( _
        ByVal vData As Variant _
    ) As Long
Dim lItem As Long
Dim lAt As Long
    lAt = -1
    For lItem = 1 To m_lIDCount
        If (m_vData(lItem) = vData) Then
            lAt = lItem
            Exit For
        End If
    Next lItem
    If (lAt > 0) Then
        lAt = BandIndexForId(m_wID(lAt))
    End If
    BandIndexForData = lAt
End Property
Private Function plAddId( _
        ByVal vData As Variant _
    ) As Long
    m_lIDCount = m_lIDCount + 1
    ReDim Preserve m_wID(1 To m_lIDCount) As Long
    ReDim Preserve m_vData(1 To m_lIDCount) As Variant
    m_wID(m_lIDCount) = m_lIDCount
    m_vData(m_lIDCount) = vData
    plAddId = m_lIDCount
End Function
Private Sub pAddWnds( _
        ByVal hWndItem As Long, _
        ByVal hWndParent As Long _
    )
    m_iWndItemCount = m_iWndItemCount + 1
    ReDim Preserve m_hWndItem(1 To m_iWndItemCount) As Long
    ReDim Preserve m_hWNdItemParent(1 To m_iWndItemCount) As Long
    m_hWndItem(m_iWndItemCount) = hWndItem
    m_hWNdItemParent(m_iWndItemCount) = hWndParent
End Sub
Private Sub pResetParent( _
        ByVal hWndItem As Long _
    )
Dim iItem As Integer
Dim iThisItem As Integer
    For iItem = 1 To m_iWndItemCount
        If (m_hWndItem(iItem) = hWndItem) Then
            iThisItem = iItem
            Exit For
        End If
    Next iItem
    If (iThisItem > 0) Then
        SetParent hWndItem, m_hWNdItemParent(iThisItem)
        If (m_iWndItemCount > 1) Then
            For iItem = iThisItem To m_iWndItemCount - 1
                m_hWndItem(iItem) = m_hWndItem(iItem + 1)
                m_hWNdItemParent(iItem) = m_hWNdItemParent(iItem + 1)
            Next iItem
            ReDim Preserve m_hWndItem(1 To m_iWndItemCount - 1) As Long
            ReDim Preserve m_hWNdItemParent(1 To m_iWndItemCount - 1) As Long
        Else
            If (m_iWndItemCount = 1) Then
                Erase m_hWndItem
                Erase m_hWNdItemParent
            End If
        End If
        m_iWndItemCount = m_iWndItemCount - 1
        If (m_iWndItemCount < 0) Then
            m_iWndItemCount = 0
        End If
    Else
        Debug.Print "Failed to reset parent.."
        ShowWindow hWndItem, SW_HIDE
        SetParent hWndItem, 0
    End If
End Sub
Private Sub RebarSize()
Dim lWidth As Long
Dim lHeight As Long
Dim rc As RECT
    If (m_hWnd <> 0) Then
        lWidth = UserControl.Parent.Width \ Screen.TwipsPerPixelX
        lHeight = RebarHeight
        If (lWidth > 0) And (lHeight > 0) Then
            MoveWindow m_hWnd, 0, 0, lWidth, lHeight, 1
        End If
         rc.Right = lWidth: rc.Bottom = lHeight
        InvalidateRect UserControl.Parent.hWnd, rc, True
         UpdateWindow m_hWnd
    End If
End Sub
Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
Property Get RebarHwnd() As Long
    RebarHwnd = m_hWnd
End Property
Public Property Get RebarHeight() As Long
Dim tc As RECT
    If (m_hWnd <> 0) Then
      GetWindowRect m_hWnd, tc
      RebarHeight = (tc.Bottom - tc.Top)
    End If
End Property
Private Property Get RebarWidth() As Long
Dim tc As RECT
   If (m_hWnd <> 0) Then
      GetWindowRect m_hWnd, tc
      RebarWidth = (tc.Right - tc.Left)
   End If
End Property
Public Sub Resize(frm As Object)
On Error Resume Next
Dim rc As RECT
   GetWindowRect m_hWnd, rc
   rc.Right = frm.Width / Screen.TwipsPerPixelX - 8
   MoveWindow m_hWnd, 0, 0, rc.Right, (rc.Bottom - rc.Top), True
   UpdateWindow m_hWnd
   On Error GoTo 0
End Sub
Private Function pbLoadCommCtls() As Boolean
Dim ctEx As CommonControlsEx

    ctEx.dwSize = Len(ctEx)
    ctEx.dwICC = ICC_COOL_CLASSES Or _
        ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES
    
    pbLoadCommCtls = (InitCommonControlsEx(ctEx) <> 0)

End Function
Private Function plBuildCoolBar( _
        ByVal hWndParent As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        Optional ByVal bVertical As Boolean = False _
    ) As Long
Dim hwndCoolBar As Long
Dim lResult As Long
Dim cStyle As Long

    cStyle = WS_CHILD Or WS_BORDER Or _
        WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or _
        WS_VISIBLE Or RBS_VARHEIGHT Or _
        RBS_BANDBORDERS Or RBS_TOOLTIPS _
        Or CCS_NODIVIDER Or CCS_NOPARENTALIGN
    If bVertical Then
        cStyle = (cStyle Or CCS_VERT)
    End If
    
    plBuildCoolBar = CreateWindowEX(WS_EX_TOOLWINDOW, _
                            REBARCLASSNAME, "", _
                            cStyle, 0, 0, lWidth, lHeight, _
                            hWndParent, ICC_COOL_CLASSES, App.hInstance, ByVal 0&)

End Function

Private Function pbCreateRebar() As Boolean
Dim lWidth As Long
Dim lHeight As Long
Dim hWndParent As Long
        
    ' Try to load the Common Controls support for the
    ' rebar control:
    If (UserControl.Ambient.UserMode) Then
      If (pbLoadCommCtls()) Then
         'Debug.Print "Loaded Coolbar support"
         ' If we have done this, then build a rebar:
         lWidth = UserControl.Parent.Width \ Screen.TwipsPerPixelX
         lHeight = UserControl.Height \ Screen.TwipsPerPixelY
         hWndParent = UserControl.Parent.hWnd
         m_hWnd = plBuildCoolBar(hWndParent, lWidth, lHeight)
         If (m_hWnd <> 0) Then
            ' Debug.Print "Created Rebar Window"
            pbCreateRebar = True
         End If
      End If
    End If
    
End Function
Private Sub pDestroyRebar()
    If (m_hWnd <> 0) Then
        ' Debug.Print "Destroying rebar window"
        RemoveAllRebarBands
        SetParent m_hWnd, 0
        DestroyWindow m_hWnd
        m_hWnd = 0
        pDestroySubClass
    End If
End Sub

Public Sub RemoveAllRebarBands()
Dim lBands As Long
Dim lBand As Long
    If (m_hWnd <> 0) Then
        pDestroySubClass
        lBands = BandCount
        For lBand = 0 To lBands - 1
            RemoveBand 0
        Next lBand
    End If
End Sub
Public Sub RemoveBand( _
        ByVal lBand As Long _
    )
Dim lHwnd As Long
    If (m_hWnd <> 0) Then
        ' If a valid band:
        If (lBand >= 0) And (lBand < BandCount) Then
            lHwnd = plGetHwndOfBandChild(m_hWnd, lBand)
            If (lHwnd <> 0) Then
                pResetParent lHwnd
            End If
            SendMessageLong m_hWnd, RB_DELETEBAND, lBand, 0&
            If (BandCount = 0) Then
                pDestroySubClass
            End If
        End If
    End If
End Sub
Private Function plGetHwndOfBandChild( _
        ByVal lHwnd As Long, _
        ByVal lBand As Long _
    ) As Long
Dim lParam As Long
Dim tRbbi As REBARBANDINFO_NOTEXT
Dim lR As Long

    tRbbi.cbSize = Len(tRbbi)
    tRbbi.fMask = RBBIM_CHILD
    lR = SendMessage(lHwnd, RB_GETBANDINFO, lBand, tRbbi)
    If (lR <> 0) Then
        plGetHwndOfBandChild = tRbbi.hWndChild
    End If
End Function


Private Sub ClearPicture()
    If (m_bPictureLoaded) Then
        m_sPicture = ""
        m_bPictureLoaded = False
    End If
End Sub

Private Sub UserControl_Initialize()
    Debug.Print "cRebar:Initialise"
    Set m_pic = New StdPicture
    'm_HwndControl = hWnd
End Sub

Private Sub UserControl_InitProperties()
    ' Set up the rebar:
    If (pbCreateRebar()) Then
        ' Ini init properties we must be in design mode.
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    ' Read in properties here:
    ' ...
    
    ' Set up the rebar:
    If (pbCreateRebar()) Then
        ' If we're not in design mode, then go
        ' for a sub class:
        If (UserControl.Ambient.UserMode) Then
            m_HwndParent = UserControl.Parent.hWnd
        End If
        ShowWindow UserControl.hWnd, SW_HIDE
    End If
    
End Sub

Private Sub UserControl_Resize()
   If (UserControl.Ambient.UserMode) Then
      UserControl.Width = 0
      UserControl.Height = 0
   End If
End Sub

Private Sub UserControl_Terminate()
    'pbSubClass False
    'ClearPicture
    'Set m_pic = LoadPicture("")
    'Set m_pic = Nothing
    m_bInTerminate = True
    pDestroyRebar
    Debug.Print "cRebar:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' Write properties here:
    ' ...
    
End Sub


