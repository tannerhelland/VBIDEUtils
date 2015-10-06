VERSION 5.00
Begin VB.UserControl cToolbar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   3855
   Begin VB.Label lblInfo 
      Caption         =   "'Toolbar control'"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4275
   End
End
Attribute VB_Name = "cToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' Declares, constants and types required for toolbar:
' ==============================================================================
' Types:
Private Type ButtonInfo
   idNum As Integer
   BitMapNum As Integer
   PosNum As Integer
   TipsText As String
   TextIndexNum As Integer
   ButnText As String
   Large As Integer
   xWidth As Integer
   xHeight As Integer
   sKey As String
End Type

Private Type BCommand
   Command As Integer
   TipNum As Integer
End Type

Private Type TBADDBITMAP
   hinst As Long
   nID As Long
End Type

Private Type BUTTONDATA
   iBitmap As Integer
   idCommand As Integer
   fsState As Long
   fsStyle As Long
   lpszButtonText As String
   lpszTooltip As String
End Type

' Toolbar window class name:
Private Const TOOLBARCLASSNAME = "ToolbarWindow32"

' Toolbar and button styles:
Private Const TBSTYLE_BUTTON = &H0
Private Const TBSTYLE_SEP = &H1
Private Const TBSTYLE_CHECK = &H2
Private Const TBSTYLE_GROUP = &H4
Private Const TBSTYLE_CHECKGROUP = (TBSTYLE_GROUP Or TBSTYLE_CHECK)
Private Const TBSTYLE_DROPDOWN = &H8
Private Const TBSTATE_ENABLED = &H4
Private Const TBSTYLE_TOOLTIPS = &H100
Private Const TBSTYLE_WRAPABLE = &H200
Private Const TBSTYLE_ALTDRAG = &H400

Private Const TBSTYLE_FLAT = &H800
Private Const TBSTYLE_LIST = &H1000
Private Const TBSTYLE_TRANSPARENT = &H8000

' Toolbar button states:
Private Enum ectbButtonStates
   TBSTATE_CHECKED = &H1
   TBSTATE_Pressed = &H2
   TBSTATE_WRAP = &H20
   TBSTATE_ELLIPSES = &H40
   TBSTATE_INDETERMINATE = &H10
   TBSTATE_HIDDEN = &H8
End Enum

' Toolbar notification messages:
Private Const TBN_LAST = &H720
Private Const TBN_FIRST = -700&
Private Const TBN_CLOSEUP = (TBN_FIRST - 11)
Private Const TBN_DROPDOWN = (TBN_FIRST - 10)

' Toolbar messages:
Private Const TB_ADDBUTTONS = (WM_USER + 20)
Private Const TB_INSERTBUTTON = (WM_USER + 21)
Private Const TB_GETBUTTON = (WM_USER + 23)
Private Const TB_BUTTONCOUNT = (WM_USER + 24)
Private Const TB_COMMANDTOINDEX = (WM_USER + 25)
Private Const TB_BUTTONSTRUCTSIZE = (WM_USER + 30)
Private Const TB_SETMAXTEXTROWS = (WM_USER + 60)
Private Const TB_SETBITMAPSIZE = (WM_USER + 32)
Private Const TB_SETBUTTONWIDTH = (WM_USER + 59)
Private Const TB_SETBUTTONSIZE = (WM_USER + 31)
Private Const TB_AUTOSIZE = (WM_USER + 33)

Private Const TB_DELETEBUTTON = (WM_USER + 22)
Private Const TB_ENABLEBUTTON = (WM_USER + 1)
Private Const TB_CHECKBUTTON = (WM_USER + 2)
Private Const TB_PRESSBUTTON = (WM_USER + 3)
Private Const TB_HIDEBUTTON = (WM_USER + 4)
Private Const TB_INDETERMINATE = (WM_USER + 5)

Private Const TB_ISBUTTONENABLED = (WM_USER + 9)
Private Const TB_ISBUTTONCHECKED = (WM_USER + 10)
Private Const TB_ISBUTTONPRESSED = (WM_USER + 11)
Private Const TB_ISBUTTONHIDDEN = (WM_USER + 12)
Private Const TB_ISBUTTONINDETERMINATE = (WM_USER + 13)

Private Const TB_GETBUTTONTEXTA = (WM_USER + 45)
Private Const TB_ADDSTRING = (WM_USER + 28)
Private Const TB_SETSTATE = (WM_USER + 17)
Private Const TB_GETSTATE = (WM_USER + 18)
Private Const TB_ADDBITMAP = (WM_USER + 19)
Private Const TB_SETPARENT = (WM_USER + 37)
Private Const TB_GETITEMRECT = (WM_USER + 29)
' Extended style:
Private Const TB_SETEXTENDEDSTYLE = (WM_USER + 84)    ' // For TBSTYLE_EX_*
Private Const TBSTYLE_EX_DRAWDDARROWS = &H1
Private Const TB_GETEXTENDEDSTYLE = (WM_USER + 85)    ' // For TBSTYLE_EX_*

Private Declare Function CreateToolbarEx Lib "COMCTL32" (ByVal hwnd As Long, ByVal ws As Long, ByVal wId As Long, ByVal nBitmaps As Long, ByVal hBMInst As Long, ByVal wBMID As Long, ByRef lpButtons As TBBUTTON, ByVal iNumButtons As Long, ByVal dxButton As Long, ByVal dyButton As Long, ByVal dxBitmap As Long, ByVal dyBitmap As Long, ByVal uStructSize As Long) As Long

' ==============================================================================
' INTERFACE
' ==============================================================================
' Enumerations:
Public Enum ECTBToolButtonSyle
   CTBNormal = TBSTYLE_BUTTON
   CTBSeparator = TBSTYLE_SEP
   CTBCheck = TBSTYLE_CHECK
   CTBCheckGroup = TBSTYLE_CHECKGROUP
   CTBDropDown = TBSTYLE_DROPDOWN
End Enum
Public Enum ECTBToolbarStyle
   CTBFlat = TBSTYLE_FLAT
   CTBList = TBSTYLE_LIST
   CTBTransparent = -1 ' special - here we remove Toolbar from owner window
End Enum
Public Enum ECTBImageSourceTypes
   CTBResourceBitmap
   CTBLoadFromFile
   CTBExternalImageList
   CTBInternalImageLists
   CTBPicture
End Enum
' Events:
Public Event ButtonClick(ByVal lButton As Long)
Public Event DropDownPress(ByVal lButton As Long)

' ==============================================================================
' INTERNAL INFORMATION
' ==============================================================================
' Subclassing
Implements ISubclass
Private m_emr As EMsgResponse
Private m_bInSubClass As Boolean

' Hwnd of tool bar itself:
Private m_hToolBarWnd As Long
' Where the button images are coming from
Private m_eImageSourceType As ECTBImageSourceTypes
Private m_pic As Long
Private m_sFileName As String
Private m_lResourceId As Long
Private m_hIml As Long

Private CaptionLen As Integer
Private m_iButtonWidth As Integer
Private m_iButtonHeight As Integer
Private Buttons32 As Integer
Private m_tBInfo() As ButtonInfo
Private m_lR As Long

Public Property Get ButtonToolTip(ByVal vButton As Variant) As String
   Dim iB As Long
   iB = ButtonIndex(vButton)
   If (iB > -1) Then
      ButtonToolTip = m_tBInfo(iB).TipsText
   End If
End Property
Public Property Let ButtonToolTip(ByVal vButton As Variant, ByVal sToolTip As String)
   Dim iB As Long
   iB = ButtonIndex(vButton)
   If (iB > -1) Then
      m_tBInfo(iB).TipsText = sToolTip
   End If
End Property
Private Function pbGetIndexForID(ByVal iBtnId As Long) As Long
   Dim iB As Long
   pbGetIndexForID = -1
   For iB = 0 To UBound(m_tBInfo)
      If (m_tBInfo(iB).idNum = iBtnId) Then
         pbGetIndexForID = iB
         Exit For
      End If
   Next iB
End Function

Public Property Get ButtonIndex(ByVal vButton As Variant) As Integer
   Dim iB As Integer
   Dim iIndex As Integer
   iIndex = -1
   If (IsNumeric(vButton)) Then
      iIndex = CInt(vButton)
   Else
      For iB = 0 To UBound(m_tBInfo)
         If (m_tBInfo(iB).sKey = vButton) Then
            iIndex = iB
            Exit For
         End If
      Next iB
   End If
   If (iIndex > -1) And (iIndex <= UBound(m_tBInfo)) Then
      ButtonIndex = iIndex
   Else
      ' error
      Debug.Print "Button index failed"
      ButtonIndex = -1
   End If

End Property

Public Property Get ButtonEnabled(ByVal vButton As Variant) As Boolean
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      ButtonEnabled = pbGetState(iID, TBSTATE_ENABLED)
   End If
End Property
Public Property Let ButtonEnabled(ByVal vButton As Variant, ByVal bState As Boolean)
   Dim iButton As Long
   Dim iID As Long
   Dim lEnable As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      pbSetState iID, TBSTATE_ENABLED, bState
   End If
End Property
Public Property Get ButtonVisible(ByVal vButton As Variant) As Boolean
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      ButtonVisible = Not (pbGetState(iID, TBSTATE_HIDDEN))
   End If
End Property
Public Property Let ButtonVisible(ByVal vButton As Variant, ByVal bState As Boolean)
   Dim iButton As Long
   Dim iID As Long
   Dim lEnable As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      pbSetState iID, TBSTATE_HIDDEN, Not (bState)
   End If
End Property

Public Property Get ButtonChecked(ByVal vButton As Variant) As Boolean
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      ButtonChecked = pbGetState(iID, TBSTATE_CHECKED)
   End If
End Property
Public Property Let ButtonChecked(ByVal vButton As Variant, ByVal bState As Boolean)
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      pbSetState iID, TBSTATE_CHECKED, bState
   End If
End Property
Public Property Get ButtonPressed(ByVal vButton As Variant) As Boolean
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      ButtonPressed = pbGetState(iID, TBSTATE_Pressed)
   End If
End Property
Public Property Let ButtonPressed(ByVal vButton As Variant, ByVal bState As Boolean)
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      pbSetState iID, TBSTATE_Pressed, bState
   End If
End Property
Public Property Get ButtonTextWrap(ByVal vButton As Variant) As Boolean
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      ButtonTextWrap = pbGetState(iID, TBSTATE_WRAP)
   End If
End Property
Public Property Let ButtonTextWrap(ByVal vButton As Variant, ByVal bState As Boolean)
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      pbSetState iID, TBSTATE_WRAP, bState
   End If
End Property
Public Property Get ButtonTextEllipses(ByVal vButton As Variant) As Boolean
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      ButtonTextEllipses = pbGetState(iID, TBSTATE_ELLIPSES)
   End If
End Property
Public Property Let ButtonTextEllipses(ByVal vButton As Variant, ByVal bState As Boolean)
   Dim iButton As Long
   Dim iID As Long
   iButton = ButtonIndex(vButton)
   If (iButton <> -1) Then
      iID = m_tBInfo(iButton).idNum
      pbSetState iID, TBSTATE_ELLIPSES, bState
   End If
End Property
Private Function pbGetState(ByVal iIDBtn As Long, ByVal fStateFlag As ectbButtonStates) As Boolean
   Dim fState As Long
   fState = SendMessageLong(m_hToolBarWnd, TB_GETSTATE, iIDBtn, 0)
   pbGetState = ((fState And fStateFlag) = fStateFlag)
End Function
Private Function pbSetState(ByVal iIDBtn As Long, ByVal fStateFlag As ectbButtonStates, ByVal bState As Boolean)
   Dim fState As Long
   fState = SendMessageLong(m_hToolBarWnd, TB_GETSTATE, iIDBtn, 0)
   If (bState) Then
      fState = fState Or fStateFlag
   Else
      fState = fState And Not fStateFlag
   End If
   If (SendMessageLong(m_hToolBarWnd, TB_SETSTATE, iIDBtn, fState) = 0) Then
      Debug.Print "Button state failed"
   Else
      pbSetState = True
   End If
End Function

Public Property Get hwnd() As Long
   hwnd = m_hToolBarWnd
End Property

Public Sub DestroyToolBar()
   On Error Resume Next
   'We need to clean up our windows
   pSubClass False
   If (m_hToolBarWnd <> 0) Then
      ShowWindow m_hToolBarWnd, SW_HIDE
      DestroyWindow (m_hToolBarWnd)
      m_hToolBarWnd = 0
   End If
End Sub

Public Sub CreateToolbar(Optional ButtonSize As Integer = 16, Optional StyleList As Boolean, Optional WithText As Boolean, Optional Wrappable As Boolean, Optional PicSize As Integer)
   On Error Resume Next
   Dim Button As TBBUTTON
   Dim lParam As Long
   Dim ListButtons As Boolean
   Dim dwStyle As Long

   DestroyToolBar

   dwStyle = WS_CHILD Or WS_VISIBLE Or TB_AUTOSIZE Or _
      CCS_NODIVIDER Or TBSTYLE_TOOLTIPS Or WS_CLIPCHILDREN Or _
      CCS_NOPARENTALIGN Or CCS_NORESIZE Or TBSTYLE_FLAT
   If (StyleList) Then
      dwStyle = dwStyle Or TBSTYLE_LIST
   End If
   If (Wrappable) Then
      dwStyle = dwStyle Or TBSTYLE_WRAPABLE
   End If

   m_hToolBarWnd = CreateWindowEX(0, "ToolbarWindow32", "", _
      dwStyle, _
      0, 0, 0, 0, UserControl.Parent.hwnd, 0&, App.hInstance, 0&)

   SendMessageLong m_hToolBarWnd, TB_SETPARENT, UserControl.Parent.hwnd, 0

   m_lR = SendMessageLong(m_hToolBarWnd, TB_BUTTONSTRUCTSIZE, LenB(Button), 0)

   lParam = ButtonSize + (ButtonSize * 65536)
   m_lR = SendMessageLong(m_hToolBarWnd, TB_SETBITMAPSIZE, 0, lParam)

   AddBitmapIfRequired

   'tell the toolbar the size of the buttons
   If ButtonSize = 16 And WithText = False Then
      lParam = ButtonSize + ((ButtonSize - 1) * &H10000)
      m_lR = SendMessageLong(m_hToolBarWnd, TB_SETBUTTONSIZE, 0, lParam)
   ElseIf ButtonSize = 16 And WithText = True Then 'Else
      lParam = ButtonSize + ((ButtonSize - 2) * &H10000)
      m_lR = SendMessageLong(m_hToolBarWnd, TB_SETBUTTONSIZE, 0, lParam)
   ElseIf ButtonSize = 32 And WithText = False Then 'Else
      lParam = ButtonSize + (ButtonSize * &H10000)
      m_lR = SendMessageLong(m_hToolBarWnd, TB_SETBUTTONSIZE, 0, lParam)
   ElseIf ButtonSize = 32 And WithText = True Then
      lParam = 50 + (50 * &H10000)
      m_lR = SendMessageLong(m_hToolBarWnd, TB_SETBUTTONSIZE, 0, lParam)
   End If

   pSubClass True, UserControl.Parent.hwnd
   m_cTT.Create
   m_cTT.AddTool m_hToolBarWnd

End Sub
Public Property Let ButtonImageSource( _
   ByVal eType As ECTBImageSourceTypes _
   )
   m_eImageSourceType = eType
End Property
Public Property Let ButtonImageResourceID(ByVal lResourceId As Long)
   m_lResourceId = lResourceId
End Property
Public Property Let ButtonImageFile(ByVal sFIle As String)
   m_sFileName = sFIle
End Property
Public Property Let ButtonImageList(ByRef imlThis As Object)

End Property
Public Property Let ButtonImagePicture(ByVal hBMP As Long)
   m_pic = hBMP
End Property

Private Sub AddBitmapIfRequired()
   Dim tbab As TBADDBITMAP

   Select Case m_eImageSourceType
      Case CTBPicture
         tbab.hinst = 0
         tbab.nID = m_pic
         ' Add the bitmap containing button images to the toolbar.
         m_lR = SendMessage(m_hToolBarWnd, TB_ADDBITMAP, 54, tbab)
      Case CTBLoadFromFile
         tbab.hinst = 0
         tbab.nID = LoadImage(0, m_sFileName, IMAGE_BITMAP, 0, 0, _
            LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
         m_lR = SendMessage(m_hToolBarWnd, TB_ADDBITMAP, 54, tbab)
      Case CTBResourceBitmap
         tbab.hinst = 0
         tbab.nID = LoadImageLong(App.hInstance, m_lResourceId, IMAGE_BITMAP, 0, 0, _
            LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT)
         m_lR = SendMessage(m_hToolBarWnd, TB_ADDBITMAP, 54, tbab)
      Case CTBInternalImageLists
      Case CTBExternalImageList
   End Select
End Sub

Public Sub AddButton( _
   ByVal id As Integer, _
   Optional ByVal zTip As String = "", _
   Optional ByVal BitPic As Integer = -1, _
   Optional ByVal PosNumber As Integer = -1, _
   Optional ByVal xLarge As Integer = 0, _
   Optional ByVal ButtonText As String, _
   Optional ByVal ButtonStyle As ECTBToolButtonSyle, _
   Optional ByVal Key As String = "" _
   )
   Dim sBuffer As String
   Dim NewStyle As Long
   Dim Button As TBBUTTON
   Dim lParam As Long
   Static s_iNum As Integer

   On Error Resume Next

   If id = 0 Then
      SendMessageLong m_hToolBarWnd, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS
   End If

   If id > UBound(m_tBInfo) Then
      ReDim Preserve m_tBInfo(id)
   End If
   If PosNumber = -1 Then PosNumber = id

   m_tBInfo(id).idNum = id
   m_tBInfo(id).BitMapNum = BitPic
   m_tBInfo(id).TipsText = zTip
   m_tBInfo(id).PosNum = PosNumber
   m_tBInfo(id).Large = xLarge
   m_tBInfo(id).sKey = Key

   If Len(ButtonText) > 0 Then
      sBuffer = String$(50, 0)
      sBuffer = Trim(ButtonText)
      m_lR = SendMessageStr(m_hToolBarWnd, TB_ADDSTRING, 0, sBuffer)
      m_tBInfo(id).ButnText = sBuffer
   End If

   Button.iBitmap = BitPic
   Button.idCommand = id
   Button.fsState = TBSTATE_ENABLED
   Button.fsStyle = ButtonStyle
   Button.dwData = 0
   Button.iString = m_lR

   m_lR = SendMessage(m_hToolBarWnd, TB_ADDBUTTONS, 1, Button)

   ' Size window:
   pResizeToolbar
End Sub
Private Sub pResizeToolbar()
   Dim tR As RECT
   Dim rc As RECT
   Dim lCount As Long
   Dim Button As TBBUTTON

   ' Get number of buttons:
   lCount = SendMessageLong(m_hToolBarWnd, TB_BUTTONCOUNT, 0, 0)
   If (lCount > 0) Then
      ' Get rectangle for last button:
      SendMessage m_hToolBarWnd, TB_GETITEMRECT, m_tBInfo(lCount - 1).idNum, rc
      ' Get rectangle for toolbar:
      GetWindowRect m_hToolBarWnd, tR
      ' Make window correct size:
      MoveWindow m_hToolBarWnd, tR.Left, tR.Top, rc.Right, rc.Bottom - 2, 1

   End If
End Sub
Public Sub ButtonSize(xWidth As Integer, xHeight As Integer)
   m_iButtonWidth = xWidth
   m_iButtonHeight = xHeight
End Sub
Public Sub GetDropDownPosition( _
   ByVal id As Integer, _
   ByRef x As Long, _
   ByRef y As Long _
   )
   Dim rc As RECT
   Dim tP As POINTAPI

   SendMessage m_hToolBarWnd, TB_GETITEMRECT, id - 1, rc
   tP.x = rc.Left
   tP.y = rc.Bottom
   MapWindowPoints m_hToolBarWnd, UserControl.Parent.hwnd, tP, 2
   x = tP.x * Screen.TwipsPerPixelX
   y = tP.y * Screen.TwipsPerPixelY

   'Me.DropDownPress 0, CInt(wParam), tp.x, tp.y, rc.Top, rc.Left, rc.Bottom, rc.Right
   'tpm.cbSize = len(tpm)sizeof(TPMPARAMS);
   'tpm.rcExclude.top    = rc.top;
   'tpm.rcExclude.left   = rc.left;
   'tpm.rcExclude.bottom = rc.bottom;
   'tpm.rcExclude.right  = rc.right;

End Sub

Public Sub Resize(frm As Object)
   On Error Resume Next
   'We need to Hide the toolbars so they will repaint
   'correctly.  Not required with IE 4.0 updated files
   Call ShowWindow(m_hToolBarWnd, SW_HIDE)

   'Resize toolbars
   Call MoveWindow(m_hToolBarWnd, 0, 0, frm.Width / Screen.TwipsPerPixelX - 5, frm.Height / Screen.TwipsPerPixelY - 5, True)

   'Update the Windows
   Call UpdateWindow(m_hToolBarWnd)

   'Show the Toolbars
   Call ShowWindow(m_hToolBarWnd, SW_SHOW)
End Sub
Public Sub PressButton(id As Integer)
   On Error Resume Next

   'Initialize button structure
   Dim Button As TBBUTTON

   'Check the button
   Call SendMessage(m_hToolBarWnd, TB_CHECKBUTTON, id, Button)
   'Update the Window
   Call UpdateWindow(m_hToolBarWnd)

End Sub

Private Sub pInitialise()
   If Not (UserControl.Ambient.UserMode) Then
      ' We are in design mode:

   Else
      ' We are in run
      Dim iccex As tagInitCommonControlsEx

      With iccex
         .lngSize = LenB(iccex)
         .lngICC = ICC_BAR_CLASSES
      End With

      'We need to make this call to make sure the common controls are loaded
      InitCommonControlsEx iccex

      m_hToolBarWnd = 0

   End If
End Sub
Private Sub pSubClass(ByVal bState As Boolean, Optional ByVal lHwnd As Long = 0)
   Static s_lhWndSave As Long

   m_emr = emrPreprocess
   If (m_bInSubClass <> bState) Then
      If (bState) Then
         Debug.Print "Subclassing:Start"
         Debug.Assert (lHwnd <> 0)
         If (s_lhWndSave <> 0) Then
            pSubClass False
         End If
         s_lhWndSave = lHwnd
         pAttMsg lHwnd, WM_COMMAND
         pAttMsg lHwnd, WM_MOUSEMOVE
         pAttMsg lHwnd, WM_LBUTTONDOWN
         pAttMsg lHwnd, WM_LBUTTONUP
         pAttMsg lHwnd, WM_RBUTTONDOWN
         pAttMsg lHwnd, WM_RBUTTONUP
         pAttMsg lHwnd, WM_MBUTTONDOWN
         pAttMsg lHwnd, WM_MBUTTONUP
         pAttMsg lHwnd, WM_NOTIFY
         s_lhWndSave = lHwnd
         m_bInSubClass = True
      Else
         Debug.Print "Subclassing:End"
         pDelMsg s_lhWndSave, WM_COMMAND
         pDelMsg s_lhWndSave, WM_MOUSEMOVE
         pDelMsg s_lhWndSave, WM_LBUTTONDOWN
         pDelMsg s_lhWndSave, WM_LBUTTONUP
         pDelMsg s_lhWndSave, WM_RBUTTONDOWN
         pDelMsg s_lhWndSave, WM_RBUTTONUP
         pDelMsg s_lhWndSave, WM_MBUTTONDOWN
         pDelMsg s_lhWndSave, WM_MBUTTONUP
         pDelMsg s_lhWndSave, WM_NOTIFY
         s_lhWndSave = 0
         m_bInSubClass = False
      End If
   End If
End Sub
Private Sub pTerminate()
   DestroyToolBar
End Sub
Private Sub pAttMsg(ByVal lHwnd As Long, ByVal lMsg As Long)
   AttachMessage Me, lHwnd, lMsg
End Sub
Private Sub pDelMsg(ByVal lHwnd As Long, ByVal lMsg As Long)
   DetachMessage Me, lHwnd, lMsg
End Sub

Public Function RaiseButtonClick(ByVal iIDButton As Long)
   ' Required as part of the WM_COMMAND handler:
   RaiseEvent ButtonClick(iIDButton)
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As VBIDEUtils0.EMsgResponse)
   m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As VBIDEUtils0.EMsgResponse
   ISubclass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim msgStruct As msg
   Dim hdr As NMHDR
   Dim ttt As NMTTDISPINFO
   Dim MinMax As MINMAXINFO
   Dim pt32 As POINTAPI
   Dim ptx As Long
   Dim pty As Long
   Dim hWndOver As Long
   Dim sToolTipBuffer As String
   Dim b() As Byte
   Dim iB As Long
   Dim lPtr As Long
   Dim cTbar As cToolbar

   On Error Resume Next

   ' Process messages here:
   m_emr = emrPreprocess
   Select Case iMsg
      Case WM_COMMAND
         'Debug.Print "WM_COMMAND", hWnd, iMsg, wParam, lParam, UserControl.Name
         If (lParam = m_hToolBarWnd) Then
            iB = SendMessage(m_hToolBarWnd, TB_COMMANDTOINDEX, wParam, 0)
            RaiseEvent ButtonClick(iB + 1)
            ISubclass_WindowProc = 0
         End If

      Case WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP
         'Debug.Print "MouseEvent"
         ' What is this for?
         With msgStruct
            .lParam = lParam
            .wParam = wParam
            .message = iMsg
            .hwnd = hwnd
         End With

         'Pass the structure
         SendMessage m_cTT.hwnd, TTM_RELAYEVENT, 0, 0

      Case WM_NOTIFY
         CopyMemory hdr, ByVal lParam, Len(hdr)
         'Debug.Print "WM_NOTIFY", Hex$(hdr.code)

         If hdr.code = TTN_NEEDTEXT Then
            'Debug.Print "TTN_NEEDTEXT"
            Dim idNum As Integer
            idNum = hdr.idfrom
            On Error Resume Next

            iB = pbGetIndexForID(idNum)
            If (iB > -1) Then
               sToolTipBuffer = ButtonToolTip(iB) 'StrConv(ButtonToolTip(iB), vbFromUnicode)
               If Err.Number = 0 Then
                  Debug.Print "Show tool tip", ButtonToolTip(iB)
                  If (Len(sToolTipBuffer) > 0) Then
                     ReDim b(0 To Len(sToolTipBuffer) - 1)
                     b = StrConv(sToolTipBuffer, vbFromUnicode)

                     CopyMemory ttt, ByVal lParam, Len(ttt)
                     'Show the tooltips
                     CopyMemory ByVal ttt.lpszText, b(0), Len(sToolTipBuffer)
                     CopyMemory ttt.szText(0), b(0), Len(sToolTipBuffer)
                     CopyMemory ByVal lParam, ttt, Len(ttt)
                  End If
               Else
                  Err.Clear
               End If
            End If

         ElseIf hdr.code = TBN_DROPDOWN Then
            'Debug.Print "TBN_DROPDOWN", m_hToolBarWnd, hWndOver

            Dim nmTB As NMTOOLBAR
            Dim rc As RECT
            Dim tP As POINTAPI

            'Call GetCursorPos(pt32)               ' Get cursor position
            '     ptx = pt32.x
            '     pty = pt32.y
            'hWndOver = WindowFromPointXY(ptx, pty)    ' Get window cursor is over
            'If (hWndOver = m_hToolBarWnd) Then
            If (hdr.hwndFrom = m_hToolBarWnd) Then
               CopyMemory nmTB, ByVal lParam, Len(nmTB)
               iB = SendMessage(m_hToolBarWnd, TB_COMMANDTOINDEX, nmTB.iItem, 0)
               RaiseEvent DropDownPress(iB + 1)
            End If
         End If
   End Select

End Function
Private Property Get ObjectFromPtr(ByVal lPtr As Long) As cToolbar
   Dim oThis As Object
   Dim oThisT As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oThisT, lPtr, 4
   ' Assign to legal reference
   Set oThis = oThisT
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oThisT, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
   Set ObjectFromPtr = oThis
End Property

Private Sub UserControl_Initialize()
   Debug.Print "cToolbar:Initialize"
End Sub

Private Sub UserControl_InitProperties()
   ' Initialise the control
   pInitialise
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   ' Read properties:

   ' Initialise the control
   pInitialise
End Sub

Private Sub UserControl_Terminate()
   ' Clear up as required:
   pTerminate
   Debug.Print "cToolbar:Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   ' Write properties:
End Sub

