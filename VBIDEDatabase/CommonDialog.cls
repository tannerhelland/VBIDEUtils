VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:54
' * Module Name      : class_CommonDialog
' * Module Filename  : CommonDialog.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

' ==========================================================================
' Class:    class_CommonDialog
' Filename: class_CommonDialog.cls
' Author:   Steve McMahon
' Date:     24 May 1998
'
' A wrapper around class_GCommonDialog to make it look more
' like the standard common dialog control.
' ==========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48&

' Properties to emulate the CommonDialog control:
Private m_bCancelError  As Boolean
Private m_sFilter       As String
Private m_lFilterIndex  As Long
Private m_sFileName     As String
Private m_oColor        As OLE_COLOR
Private m_lCopies       As Long
Private m_lFlags        As Long
Private m_sDialogTitle  As String
Private m_sDefaultExt   As String
Private m_font          As New StdFont
Private m_oFontColor    As OLE_COLOR
Private m_lFromPage     As Long
Private m_lhWnd         As Long
Private m_eHelpCommand  As EShowHelpCommands
Private m_sHelpContext  As String
Private m_sHelpFile     As String
Private m_sHelpKey      As String
Private m_sInitDir      As String
Private m_lMax          As Long
Private m_lMaxFileSize  As Long
Private m_lMin          As Long
Private m_objObject     As Object
Private m_iPrinterDefault As Integer
Private m_lToPage       As Long
Private m_sFileTitle    As String
Private m_hDC           As Long
Private m_bHookDialog   As Boolean
Private mCommonDialog   As New class_GCommonDialog

'API function inside ShowHelp method
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
' WinHelp Commands:
Public Enum EShowHelpCommands
   HELP_COMMAND = &H102&
   HELP_CONTENTS = &H3&
   HELP_CONTEXT = &H1          '  Display topic in ulTopic
   HELP_CONTEXTPOPUP = &H8&
   HELP_FORCEFILE = &H9&
   HELP_HELPONHELP = &H4       '  Display help on using help
   HELP_INDEX = &H3            '  Display index
   HELP_KEY = &H101            '  Display topic for keyword in offabData
   HELP_MULTIKEY = &H201&
   HELP_PARTIALKEY = &H105&
   HELP_QUIT = &H2             '  Terminate help
   HELP_SETCONTENTS = &H5&
   HELP_SETINDEX = &H5         '  Set current Index for multi index help
   HELP_SETWINPOS = &H203&

   HELP_FINDER = &HB           ' Win95 version of HELP_CONTENTS
End Enum
Private Type HELPWININFO
   wStructSize          As Long
   x                    As Long
   y                    As Long
   dx                   As Long
   dy                   As Long
   wMax                 As Long
   rgchMember           As String * 2
End Type
Public Enum EShowHelpWindowPos
   SW_HIDE = 0
   SW_MAXIMIZE = 3
   SW_MINIMIZE = 6
   SW_SHOW = 5
   SW_SHOWMAXIMIZED = 3
   SW_SHOWMINIMIZED = 2
   SW_SHOWMINNOACTIVE = 7
   SW_SHOWNA = 8
   SW_SHOWNOACTIVATE = 4
   SW_SHOWNORMAL = 1
End Enum
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type
Private Type POINTAPI
   x                    As Long
   y                    As Long
End Type
Private m_bFileDialog   As Boolean

Public Event InitDialog(ByVal hDlg As Long)
Public Event FileChange(ByVal hDlg As Long)
Public Event FolderChange(ByVal hDlg As Long)
Public Event DialogOK(ByRef bCancel As Boolean)
Public Event TypeChange(ByVal hDlg As Long)
Public Event DialogClose()

Public Sub ParseMultiFileName( _
   ByRef sDir As String, _
   ByRef sFiles() As String, _
   ByRef iFileCount As Long _
   )
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ParseMultiFileName
   ' * Parameters       :
   ' *                    ByRef sDir As String
   ' *                    ByRef sFiles() As String
   ' *                    ByRef iFileCount As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim iPos             As Long
   Dim iNextPos         As Long
   Dim sAllFiles        As String
   Dim i                As Long

   iPos = InStr(m_sFileName, vbNullChar & vbNullChar)
   If iPos <> 0 Then
      ' multi names
      sAllFiles = left$(m_sFileName, iPos - 1)
      iPos = 1
      iNextPos = InStr(sAllFiles, vbNullChar)
      Do While iNextPos <> 0
         If (sDir = "") Then
            sDir = Mid$(sAllFiles, iPos, iNextPos - iPos)
         Else
            iFileCount = iFileCount + 1
            ReDim Preserve sFiles(1 To iFileCount) As String
            sFiles(iFileCount) = Mid$(sAllFiles, iPos, iNextPos - iPos)
         End If
         iPos = iNextPos + 1
         iNextPos = InStr(iPos, sAllFiles, vbNullChar)
      Loop
      iFileCount = iFileCount + 1
      ReDim Preserve sFiles(1 To iFileCount) As String
      sFiles(iFileCount) = Mid$(sAllFiles, iPos)
   Else
      ' single file
      iFileCount = 1
      ReDim sFiles(1 To 1) As String
      For i = Len(m_sFileName) To 1 Step -1
         If Mid$(m_sFileName, i, 1) = "\" Then
            If (i > 1) Then
               sDir = left$(m_sFileName, i - 1)
               sFiles(1) = Mid$(m_sFileName, i + 1)
            Else
               sDir = ""
               sFiles(1) = m_sFileName
            End If
            Exit Sub
         End If
      Next i
      sDir = ""
      sFiles(1) = m_sFileName
   End If
End Sub

Public Sub DialogClose()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : DialogClose
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RaiseEvent DialogClose
End Sub

Public Function TypeChange(ByVal hDlg As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : TypeChange
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RaiseEvent TypeChange(hDlg)
End Function

Public Function InitDialog(ByVal hDlg As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : InitDialog
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RaiseEvent InitDialog(hDlg)
End Function

Public Function FileChange(ByVal hDlg As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FileChange
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RaiseEvent FileChange(hDlg)
End Function

Public Function FolderChange(ByVal hDlg As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FolderChange
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RaiseEvent FolderChange(hDlg)
End Function

Public Function ConfirmOK() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ConfirmOK
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim bCancel          As Boolean
   bCancel = False
   RaiseEvent DialogOK(bCancel)
   If (bCancel) Then
      ConfirmOK = False
   Else
      ConfirmOK = True
   End If
End Function

Public Sub CentreDialog(ByVal hDlg As Long, ByRef oCentreTo As Object)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : CentreDialog
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' *                    ByRef oCentreTo As Object
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lHwnd            As Long
   Dim tWR              As RECT
   Dim tDR              As RECT
   Dim lHwndCentreTo    As Long
   Dim lL               As Long
   Dim lT               As Long
   Dim lR               As Long

   ' If we're showing a file dialog, then the rectangle is the
   ' parent of the dialog itself:
   If (m_bFileDialog) Then
      lHwnd = GetParent(hDlg)
   Else
      lHwnd = hDlg
   End If
   GetWindowRect lHwnd, tDR
   'Debug.Print tDR.Right - tDR.Left, tDR.Bottom - tDR.Top

   On Error Resume Next
   lHwndCentreTo = oCentreTo.hWnd
   If (err.number = 0) Then
      GetWindowRect lHwndCentreTo, tWR
   Else
      ' Assume the screen object:
      lR = SystemParametersInfo(SPI_GETWORKAREA, 0, tWR, 0)
      If (lR = 0) Then
         ' Call failed - just use standard screen:
         tWR.left = 0
         tWR.top = 0
         tWR.right = Screen.Width \ Screen.TwipsPerPixelX
         tWR.bottom = Screen.Height \ Screen.TwipsPerPixelY
      End If
   End If
   On Error GoTo 0
   If (tWR.right > 0) And (tWR.bottom > 0) Then
      lL = tWR.left + (((tWR.right - tWR.left) - (tDR.right - tDR.left)) \ 2)
      lT = tWR.top + (((tWR.bottom - tWR.top) - (tDR.bottom - tDR.top)) \ 2)
      'Debug.Print tDR.Right - tDR.Left, tDR.Bottom - tDR.Top
      MoveWindow lHwnd, lL, lT, (tDR.right - tDR.left), (tDR.bottom - tDR.top), 1
   End If
End Sub

Public Sub GetDialogSize( _
   ByVal hDlg As Long, _
   ByRef lL As Long, _
   ByRef lT As Long, _
   ByRef lW As Long, _
   ByRef lH As Long _
   )
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : GetDialogSize
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' *                    ByRef lL As Long
   ' *                    ByRef lT As Long
   ' *                    ByRef lW As Long
   ' *                    ByRef lH As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lHwnd            As Long
   Dim tDR              As RECT
   If (m_bFileDialog) Then
      lHwnd = GetParent(hDlg)
   Else
      lHwnd = hDlg
   End If
   GetWindowRect lHwnd, tDR
   lL = tDR.left
   lT = tDR.top
   lW = tDR.right - tDR.left
   lH = tDR.bottom - tDR.top
End Sub

Public Sub SetHelpPosition(x As Long, y As Long, dx As Long, dy As Long, eWindowType As EShowHelpWindowPos)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : SetHelpPosition
   ' * Parameters       :
   ' *                    x As Long
   ' *                    Y As Long
   ' *                    dx As Long
   ' *                    dy As Long
   ' *                    eWindowType As EShowHelpWindowPos
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim tW               As HELPWININFO
   Dim lPtr             As Long

   With tW
      .x = x
      .y = y
      .dx = dx
      .dy = dy
      .wMax = eWindowType
      .wStructSize = Len(tW)
   End With
   Dim cM               As New class_Memory
   cM.AllocateMemory tW.wStructSize
   lPtr = cM.Pointer
   CopyMemory ByVal lPtr, tW, tW.wStructSize
   WinHelp m_lhWnd, m_sHelpFile, HELP_SETWINPOS, lPtr
   cM.FreeMemory

End Sub

Public Property Get CancelError() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : CancelError
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   CancelError = m_bCancelError
End Property

Public Property Let CancelError(ByVal bCancelError As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : CancelError
   ' * Parameters       :
   ' *                    ByVal bCancelError As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_bCancelError = bCancelError
End Property

Public Property Get FileName() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Filename
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's FileName property
   FileName = m_sFileName
End Property

Public Property Let FileName(ByVal sFileName As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Filename
   ' * Parameters       :
   ' *                    ByVal sFileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's FileName property
   m_sFileName = sFileName
End Property

Public Property Get Filter() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Filter
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's Filter property
   Filter = m_sFilter
End Property

Public Property Let Filter(ByVal sFilter As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Filter
   ' * Parameters       :
   ' *                    ByVal sFilter As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's Filter property
   m_sFilter = sFilter
End Property

Public Property Get FilterIndex() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FilterIndex
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's FilterIndex property
   FilterIndex = m_lFilterIndex
End Property

Public Property Let FilterIndex(ByVal lFilterIndex As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FilterIndex
   ' * Parameters       :
   ' *                    ByVal lFilterIndex As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's FilterIndex property
   m_lFilterIndex = lFilterIndex
End Property

Public Property Get color() As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Color
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's Color property
   color = m_oColor
End Property

Public Property Let color(ByVal oColor As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Color
   ' * Parameters       :
   ' *                    ByVal oColor As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's Color property
   m_oColor = oColor
End Property

Public Property Get Copies() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Copies
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's Copies property
   Copies = m_lCopies
End Property

Public Property Let Copies(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Copies
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's Copies property
   m_lCopies = vNewValue
End Property

Public Property Get DefaultExt() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : DefaultExt
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's DefaultExt property
   DefaultExt = m_sDefaultExt
End Property

Public Property Let DefaultExt(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : DefaultExt
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's DefaultExt property
   m_sDefaultExt = vNewValue
End Property

Public Property Get DialogTitle() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : DialogTitle
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's FileName property
   DialogTitle = m_sDialogTitle
End Property

Public Property Let DialogTitle(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : DialogTitle
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's DialogTitle property
   m_sDialogTitle = vNewValue
End Property

Public Property Get Flags() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : flags
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's Flags property
   Flags = m_lFlags
End Property

Public Property Let Flags(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : flags
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's Flags property
   m_lFlags = vNewValue
End Property

Public Property Get FontBold() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontBold
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's FontBold property
   FontBold = m_font.Bold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontBold
   ' * Parameters       :
   ' *                    ByVal vNewValue As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's FontBold property
   m_font.Bold = vNewValue
End Property

Public Property Get FontItalic() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontItalic
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's FontItalic property
   FontItalic = m_font.Italic
End Property

Public Property Let FontItalic(ByVal vNewValue As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontItalic
   ' * Parameters       :
   ' *                    ByVal vNewValue As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's FontItalic property
   m_font.Italic = vNewValue
End Property

Public Property Get FontName() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontName
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's Fontname property
   FontName = m_font.Name
End Property

Public Property Let FontName(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontName
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's FontName property
   m_font.Name = vNewValue
End Property

Public Property Get FontSize() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontSize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's FontSize property
   FontSize = m_font.Size
End Property

Public Property Let FontSize(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontSize
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's FontSize property
   m_font.Size = vNewValue
End Property

Public Property Get FontStrikethru() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontStrikethru
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's FontStrikethru property
   FontStrikethru = m_font.Strikethrough
End Property

Public Property Let FontStrikethru(ByVal vNewValue As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontStrikethru
   ' * Parameters       :
   ' *                    ByVal vNewValue As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's - property
   m_font.Strikethrough = vNewValue
End Property

Public Property Get FontUnderline() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontUnderline
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's FontUnderline property
   FontUnderline = m_font.Underline
End Property

Public Property Let FontUnderline(ByVal vNewValue As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontUnderline
   ' * Parameters       :
   ' *                    ByVal vNewValue As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's FontUnderline property
   m_font.Underline = vNewValue
End Property

Public Property Get Font() As StdFont
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Font
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set Font = m_font
End Property

Public Property Let Font(sFont As StdFont)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Font
   ' * Parameters       :
   ' *                    sFont As StdFont
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set m_font = sFont
End Property

Public Property Get FontColor() As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontColor
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   FontColor = m_oFontColor
End Property

Public Property Let FontColor(oColor As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FontColor
   ' * Parameters       :
   ' *                    oColor As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_oFontColor = oColor
End Property

Public Property Get FromPage() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FromPage
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's FromPAge property
   FromPage = m_lFromPage
End Property

Public Property Let FromPage(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FromPage
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's FromPage property
   m_lFromPage = vNewValue
End Property

Public Property Get hWnd() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : hwnd
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's hWnd property
   hWnd = m_lhWnd
End Property

Public Property Let hWnd(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : hwnd
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's hWnd property
   m_lhWnd = vNewValue
End Property

Public Property Get HelpCommand() As EShowHelpCommands
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HelpCommand
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's HelpCommand property
   HelpCommand = m_eHelpCommand
End Property

Public Property Let HelpCommand(ByVal vNewValue As EShowHelpCommands)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HelpCommand
   ' * Parameters       :
   ' *                    ByVal vNewValue As EShowHelpCommands
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's HelpCommand property
   m_eHelpCommand = vNewValue
End Property

Public Property Get HelpContext() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HelpContext
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's HelpContext property
   HelpContext = m_sHelpContext
End Property

Public Property Let HelpContext(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HelpContext
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's HelpContext property
   m_sHelpContext = vNewValue
End Property

Public Property Get HelpFile() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HelpFile
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's HelpFile property
   HelpFile = m_sHelpFile
End Property

Public Property Let HelpFile(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HelpFile
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's HelpFile property
   m_sHelpFile = vNewValue
End Property

Public Property Get HelpKey() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HelpKey
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's HelpKey property
   HelpKey = m_sHelpKey
End Property

Public Property Let HelpKey(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HelpKey
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's HelpKey property
   m_sHelpKey = vNewValue
End Property

Public Property Get InitDir() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : InitDir
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's InitDir property
   InitDir = m_sInitDir
End Property

Public Property Let InitDir(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : InitDir
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's InitDir property
   m_sInitDir = vNewValue
End Property

Public Property Get Max() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Max
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's Max property
   Max = m_lMax
End Property

Public Property Let Max(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Max
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's - property
   m_lMax = vNewValue
End Property

Public Property Get MaxFileSize() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : MaxFileSize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's MaxFileSize property
   MaxFileSize = m_lMaxFileSize
End Property

Public Property Let MaxFileSize(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : MaxFileSize
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's MaxFileSize property
   m_lMaxFileSize = vNewValue
End Property

Public Property Get Min() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Min
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's Min property
   Min = m_lMin
End Property

Public Property Let Min(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Min
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's Min property
   m_lMin = vNewValue
End Property

Public Property Get Object() As Object
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Object
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's Object property
   Object = m_objObject
End Property

Public Property Let Object(ByVal vNewValue As Object)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Object
   ' * Parameters       :
   ' *                    ByVal vNewValue As Object
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's Object property
   Set m_objObject = vNewValue
End Property

Public Property Get PrinterDefault() As Integer
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : PrinterDefault
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's PrinterDefault property
   PrinterDefault = m_iPrinterDefault
End Property

Public Property Let PrinterDefault(ByVal vNewValue As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : PrinterDefault
   ' * Parameters       :
   ' *                    ByVal vNewValue As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's PrinterDefault property
   m_iPrinterDefault = vNewValue
End Property

Public Property Get ToPage() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ToPage
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Return object's ToPage property
   ToPage = m_lToPage
End Property

Public Property Let ToPage(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ToPage
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Assign object's ToPage property
   m_lToPage = vNewValue
End Property

Public Property Get FileTitle() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FileTitle
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'return object's FileTitle property
   FileTitle = m_sFileTitle
End Property

Public Property Let FileTitle(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FileTitle
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'assign object's FileTitle property
   m_sFileTitle = vNewValue
End Property

Property Get CustomColor(ByVal i As Integer) As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : CustomColor
   ' * Parameters       :
   ' *                    ByVal I As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   CustomColor = mCommonDialog.CustomColor(i)
End Property

Property Let CustomColor(ByVal i As Integer, oValue As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : CustomColor
   ' * Parameters       :
   ' *                    ByVal I As Integer
   ' *                    oValue As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   mCommonDialog.CustomColor(i) = oValue
End Property

Public Sub ShowOpen()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ShowOpen
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim bFileMustExist   As Boolean
   Dim bMultiSelect     As Boolean
   Dim bReadOnly        As Boolean
   Dim bHideReadOnly    As Boolean

   m_bFileDialog = True
   bFileMustExist = FlagSet(m_lFlags, OFN_FILEMUSTEXIST)
   bMultiSelect = FlagSet(m_lFlags, OFN_ALLOWMULTISELECT)
   bReadOnly = FlagSet(m_lFlags, OFN_READONLY)
   bHideReadOnly = FlagSet(m_lFlags, OFN_HIDEREADONLY)

   If (m_lFilterIndex = 0) Then m_lFilterIndex = 1

   If Not (mCommonDialog.VBGetOpenFileName( _
      m_sFileName, _
      m_sFileTitle, _
      bFileMustExist, bMultiSelect, bReadOnly, bHideReadOnly, _
      m_sFilter, m_lFilterIndex, _
      m_sInitDir, _
      m_sDialogTitle, _
      m_sDefaultExt, _
      m_lhWnd, _
      m_lFlags, _
      m_bHookDialog, _
      Me)) Then
      pCommonDialogError
   End If
End Sub

Public Sub ShowSave()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ShowSave
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim bOverWritePrompt As Boolean

   m_bFileDialog = True
   bOverWritePrompt = FlagSet(m_lFlags, OFN_OVERWRITEPROMPT)
   If Not (mCommonDialog.VBGetSaveFileName( _
      m_sFileName, _
      m_sFileTitle, _
      bOverWritePrompt, _
      m_sFilter, m_lFilterIndex, _
      m_sInitDir, _
      m_sDialogTitle, _
      m_sDefaultExt, _
      m_lhWnd, _
      m_lFlags, _
      m_bHookDialog, _
      Me)) Then
      pCommonDialogError
   End If
End Sub

Public Sub ShowColor()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ShowColor
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim bAnyColor        As Boolean
   Dim bFullOpen        As Boolean
   Dim bDisableFullOpen As Boolean
   Dim lColor           As Long

   m_bFileDialog = False
   lColor = TranslateColor(m_oColor)
   bAnyColor = FlagSet(m_lFlags, CC_ANYCOLOR)
   bFullOpen = FlagSet(m_lFlags, CC_FULLOPEN)
   bDisableFullOpen = FlagSet(m_lFlags, CC_PREVENTFULLOPEN)
   If Not (mCommonDialog.VBChooseColor( _
      lColor, _
      bAnyColor, bFullOpen, bDisableFullOpen, _
      m_lhWnd, m_lFlags, _
      m_bHookDialog, _
      Me)) Then
      pCommonDialogError
   Else
      m_oColor = lColor
   End If
End Sub

Public Sub ShowFont()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ShowFont
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_bFileDialog = False
   If Not (mCommonDialog.VBChooseFont( _
      m_font, _
      -1, _
      m_lhWnd, _
      m_oFontColor, _
      m_lMin, _
      m_lMax, _
      m_lFlags, _
      m_bHookDialog, _
      Me)) Then
      pCommonDialogError
   End If
End Sub

Public Sub ShowPrinter()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ShowPrinter
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim bDisablePageNumbers As Boolean
   Dim bShowPrintToFile As Boolean
   Dim bPrintToFile     As Boolean
   Dim bDisablePrintToFile As Boolean
   Dim bCollate         As Boolean
   Dim bPreventWarning  As Boolean
   Dim bDisableSelection As Boolean
   Dim ePR              As EPrintRange
   Dim iCopies          As Integer

   m_bFileDialog = False
   iCopies = m_lCopies
   bDisablePageNumbers = Not (FlagSet(m_lFlags, PD_PAGENUMS))
   bDisableSelection = FlagSet(m_lFlags, PD_NOSELECTION)
   bShowPrintToFile = Not (FlagSet(m_lFlags, PD_HIDEPRINTTOFILE))
   bDisablePrintToFile = FlagSet(m_lFlags, PD_DISABLEPRINTTOFILE)
   bPrintToFile = FlagSet(m_lFlags, PD_PRINTTOFILE)
   bCollate = FlagSet(m_lFlags, PD_COLLATE)
   bPreventWarning = FlagSet(m_lFlags, PD_NOWARNING)

   If (mCommonDialog.VBPrintDlg( _
      m_hDC, _
      ePR, _
      bDisablePageNumbers, _
      m_lFromPage, _
      m_lToPage, _
      bDisableSelection, _
      iCopies, _
      bShowPrintToFile, _
      bDisablePrintToFile, _
      bPrintToFile, _
      bCollate, _
      bPreventWarning, _
      m_lhWnd, _
      m_objObject, _
      m_lFlags, _
      m_bHookDialog, _
      Me)) Then
      ' Success
      m_lCopies = iCopies
   End If
End Sub

Public Sub ShowHelp()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : ShowHelp
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'run winhelp.exe with the specified help file
   Dim sHelpFileBuff    As String
   Dim lData            As Long
   Dim lR               As Long

   On Error GoTo ShowHelpError

   m_bFileDialog = False
   '***    prepare the buffers and parameters for the API function
   'sHelpFile is a null terminated string
   sHelpFileBuff = m_sHelpFile & Chr$(0)

   'sData is dependent on lHelpCommand
   Select Case m_eHelpCommand
      Case HELP_CONTEXT, HELP_CONTEXTPOPUP, HELP_SETCONTENTS, HELP_SETINDEX
         ' lData should be an unsigned long integer pointing to the context identifier

      Case HELP_COMMAND, HELP_PARTIALKEY
         ' lData is an address of a string

      Case HELP_CONTENTS, HELP_FORCEFILE, HELP_HELPONHELP, HELP_INDEX, HELP_QUIT
         ' lData is not required
         lData = 0
      Case HELP_SETWINPOS, HELP_MULTIKEY
         ' lData should point to a structure.  Can't call here - use separate functions to perform
         err.Raise eeBaseCommonDialog + &HFF, App.EXEName, "Invalid WinHelp Command Passed to ShowHelp function."
   End Select

   '***    call the API function
   lR = WinHelp(m_lhWnd, m_sHelpFile, m_eHelpCommand, lData)    ' - Store to APIReturn property
   If (lR <> 0) Then
      ' Success
   Else
      err.Raise LastApiError, App.EXEName & ".class_CommonDialog", APIError(LastApiError)
   End If

   Exit Sub

ShowHelpError:
   err.Raise err.number, App.EXEName & ".class_CommonDialog"
   Exit Sub

End Sub

Public Property Get HookDialog() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HookDialog
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   HookDialog = m_bHookDialog
End Property

Public Property Let HookDialog(ByVal bHook As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : HookDialog
   ' * Parameters       :
   ' *                    ByVal bHook As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_bHookDialog = bHook
End Property

Private Function APIError(ByVal e As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : APIError
   ' * Parameters       :
   ' *                    ByVal e As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim s                As String
   Dim c                As Long
   s = String$(256, 0)

   c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
      FORMAT_MESSAGE_IGNORE_INSERTS, _
      0, e, 0&, s, Len(s), ByVal 0)
   If c Then APIError = left$(s, c)
End Function

Private Function LastApiError() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : LastApiError
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   LastApiError = APIError(err.LastDllError)
End Function

Private Function FlagSet(ByVal lWord As Long, ByVal lFlagValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : FlagSet
   ' * Parameters       :
   ' *                    ByVal lWord As Long
   ' *                    ByVal lFlagValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   FlagSet = ((lWord And lFlagValue) = lFlagValue)
End Function

Private Function TranslateColor(ByVal clr As OLE_COLOR, _
   Optional hPal As Long = 0) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : TranslateColor
   ' * Parameters       :
   ' *                    ByVal clr As OLE_COLOR
   ' *                    Optional hPal As Long = 0
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If OleTranslateColor(clr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If
End Function

Private Sub pCommonDialogError()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : pCommonDialogError
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' We have an error:
   If (mCommonDialog.APIReturn = 0) Then
      ' Cancelled:
      If (m_bCancelError) Then
         ' Note if your code stops here, that is because your error
         ' options in VB are "Break in Class Module".  Change your
         ' error option to "Break on Unhandled Errors" to see how
         ' this works at runtime.
         err.Raise 20001, App.EXEName & ".class_CommonDialog", "User selected cancel."
      End If
   Else
      err.Raise eeBaseCommonDialog Or mCommonDialog.ExtendedError, App.EXEName & ".class_CommonDialog"
   End If
End Sub

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_lFilterIndex = 1
End Sub

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:54
   ' * Module Name      : class_CommonDialog
   ' * Module Filename  : CommonDialog.cls
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set mCommonDialog = Nothing
End Sub

