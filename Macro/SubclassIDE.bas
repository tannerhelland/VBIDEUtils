Attribute VB_Name = "SubclassIDE_Module"
Option Explicit

'vb ide mdi window dc
Public dc               As Long
'hwnd to the vbide mdi cliend
Public MDIClientHWND    As Long

'wdh and hgt of our bg pic
Public Width            As Long
Public Height           As Long

'holders of our ide addin objects
Public CurrentCode      As CodeModule

Public oProc            As Long
Public oMainProc        As Long

'tmp holder for our "subbed" propp
Private Subbed          As Long

'--
'holders of our macros "from -> to" data
Public macroFrom()      As String
Public macroTo()        As String
Public macroFast        As String
'--

Public Const GWL_WNDPROC = (-4)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Type Msg
   hWnd                 As Long
   message              As Long
   wParam               As Long
   lParam               As Long
   time                 As Long
   pt                   As POINTAPI
End Type

Public Const WM_PAINT = &HF
Public Const PM_REMOVE = &H1
Public Const WM_QUIT = &H12
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_KEYDOWN = &H100
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_KEYUP = &H101
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYLAST = &H108
Public Const WM_ACTIVATE = &H6
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_LBUTTONDBLCLK = &H203

'---------------------------------------------------------
'wndProc for the vb ide mdi window..
'---------------------------------------------------------
Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   Static OldWnd        As Long
   Dim ActiveHWND       As Long
   Dim state            As Long

   On Error Resume Next
   If uMsg = WM_PAINT Then
      If Not gbRegistered Then
         WindowProc = CallWindowProc(oProc, hw, uMsg, wParam, lParam)
         BitBlt dc, 20, 0, Width, Height, frmSplash.hdc, 0, 0, vbSrcCopy

         Exit Function
      End If
   End If

   ActiveHWND = FindWindowEx(hw, 0, "VbaWindow", VBInstance.ActiveWindow.Caption)
   If VBInstance.ActiveWindow Is VBInstance.ActiveCodePane.Window Then
      'were in a codewindow
      Select Case uMsg
         Case 553
            If ActiveHWND > 0 Then
               Subbed = GetProp(ActiveHWND, "SUBBED")
               If Subbed = 0 Then
                  'tell our app that this window should
                  'be subclassed
                  'store the wndproc in our "subbed" propp
                  Subbed = GetWindowLong(ActiveHWND, GWL_WNDPROC)
                  SetProp ActiveHWND, "SUBBED", Subbed
                  'now we have the old winproc in "SUBBED"
                  SetWindowLong ActiveHWND, GWL_WNDPROC, AddressOf CodeWindowProc
                  'this makes many windows assosiated to the same winproc.
                  'but since every window has its own "SUBBED"(containing that windows oldwinproc) property
                  'its not a problem
                  '
               End If
            End If
      End Select
      state = 1
   Else
      'we are on a designer
      state = 2
   End If

   If ActiveHWND <> OldWnd Then
      If state = 1 Then
         'youre in a codewindow
      Else
         'youre in a designer
      End If
   End If

   OldWnd = ActiveHWND
   WindowProc = CallWindowProc(oProc, hw, uMsg, wParam, lParam)

End Function

'-----------------------------------------
'wndProc for the codewindow!!!!!!!!!!!!!!
'-----------------------------------------
Public Function CodeWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   Dim pt               As POINTAPI
   Static OldWnd As Long
   Static Dirty As Boolean
   Static OldLineNr As Long                                'old line nr
   Dim LineNr           As Long                                      'active linenr
   Static OldLineStr As String
   Dim LineStr          As String

   Dim ThisOproc        As Long
   Dim line             As String

   Dim i                As Long

   GetCaretPos pt
   'oke what proc is assosated with this hwnd?
   'get our "subbed" prop from this window..
   ThisOproc = GetProp(hw, "SUBBED")
   'call the proc stored in our subbed propp
   CodeWindowProc = CallWindowProc(ThisOproc, hw, uMsg, wParam, lParam)
   Select Case uMsg
      Case 8                                                  'deactivate
         DetachWindow hw
         'detatch this window
         'we only want to subclass the active window
   End Select

End Function

'------------------------------
'the vb ide main window proc
'------------------------------
Public Function MainWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   MainWindowProc = CallWindowProc(oMainProc, hw, uMsg, wParam, lParam)
End Function

Private Sub CloseWindow(ByVal hWnd As Long, ByVal lParam As Long)
   'RESET WNDPROCS FOR A WINDOW THAT SHOULD BE CLOSED!
   DetachWindow hWnd
End Sub

Private Sub DetachWindow(ByVal hWnd As Long)
   Dim ThisOproc        As Long
   'oke what proc is assosated with this hwnd?
   ThisOproc = GetProp(hWnd, "SUBBED")
   'the winodw is not subclassed
   If ThisOproc = 0 Then Exit Sub

   'set our "subbed" propp to "0" to let our app know that this window is NOT
   'subclassed

   SetProp hWnd, "SUBBED", 0
   'reset the oldProc for this window
   SetWindowLong hWnd, GWL_WNDPROC, ThisOproc
End Sub
