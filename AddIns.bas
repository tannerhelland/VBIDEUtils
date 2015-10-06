Attribute VB_Name = "AddIn_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:51
' * Module Name      : AddIn_Module
' * Module Filename  : AddIns.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

'Define some public constants for the Registry keys
Public Const gsREG_SECTION As String = "Copyright © 2001 removed"
Public Const gsREG_APP  As String = "VBIDEUtils"

Global Const guidAdvFndRpl$ = "{54F45E01-1BB7-4D10-9BBC-EC7638D8D554}"
Global Const guidMouseZoom$ = "{374A295B-F4C7-407D-B72F-AE23EB36A05B}"
Global Const guidProjectExplorer$ = "{F79F051D-A76B-4841-A8D2-C687F1E2205E}"

Public Const AdvFndRplCaption = "Advanced Find/Replace"
Public Const MouseZoomCaption = "Mouse Zoom"
Public Const ProjectExplorerCaption = "Project Explorer"

Global VBInstance       As VBIDE.VBE
Global gAddInInst       As VBIDE.AddIn

Public mCommonDialog    As New class_GCommonDialog

Public gwinAdvFndRpl    As VBIDE.Window   'used to make sure we only run one instance
Public gdocAdvFndRpl    As docAdvFndRpl         'user doc object

Public gwinMouseZoom    As VBIDE.Window   'used to make sure we only run one instance
Public gdocMouseZoom    As docMouseZoom         'user doc object

Public gwinProjectExplorer As VBIDE.Window   'used to make sure we only run one instance
Public gdocProjectExplorer As docProjectExplorer         'user doc object

Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Function GenerateDLLBaseAdress() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/07/2001
   ' * Time             : 19:23
   ' * Module Name      : AddIn_Module
   ' * Module Filename  : AddIns.bas
   ' * Procedure Name   : GenerateDLLBaseAdress
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Dim nNumber          As Long
   Dim sNumber          As String

   Randomize

   nNumber = Rnd * 100

   sNumber = "&H" & Hex$(&H11000000 + nNumber * &H10000)

   Clipboard.Clear
   Clipboard.SetText sNumber

   GenerateDLLBaseAdress = sNumber

   MsgBox "The Base Address " & sNumber & vbCrLf & "has been generated and copied to the clipboard", vbInformation + vbOKOnly, "Generates Base Address"

End Function

Public Function GenerateGUID() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/07/2001
   ' * Time             : 19:23
   ' * Module Name      : AddIn_Module
   ' * Module Filename  : AddIns.bas
   ' * Procedure Name   : GenerateGUID
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Dim sGUID            As String

   sGUID = CreateGUID()

   Clipboard.Clear
   Clipboard.SetText sGUID

   GenerateGUID = sGUID

   MsgBox "The GUID " & sGUID & vbCrLf & "has been generated and copied to the clipboard", vbInformation + vbOKOnly, "Generates GUID"

End Function

'====================================================================
'this sub should be executed from the Immediate window
'in order to get this app added to the VBADDIN.INI file
'you must change the name in the 2nd argument to reflect
'the correct name of your project
'====================================================================
Sub AddToINI()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : AddIn_Module
   ' * Module Filename  : AddIns.bas
   ' * Procedure Name   : AddToINI
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim ErrCode          As Long
   ErrCode = WritePrivateProfileString("Add-Ins32", "VBIDEUtils.clsConnect", "0", "vbaddin.ini")
End Sub

Public Sub ShowMouseZoom()
   gwinMouseZoom.Visible = True
End Sub

Public Sub ShowProjectExplorer()
   gwinProjectExplorer.Visible = True
End Sub

Public Sub ShowFindAndReplace()
   gwinAdvFndRpl.Visible = True
   gdocAdvFndRpl.Initialize
   gdocAdvFndRpl.SelectedText = GetSelectedText
End Sub

Private Function GetSelectedText() As String
   Dim llStartCol       As Long
   Dim llStartLine      As Long
   Dim llEndCol         As Long
   Dim llEndLine        As Long
   Dim llTopLine        As Long
   Dim lsCodeLine       As String
   Dim cpActivePane     As VBIDE.CodePane

   On Error GoTo GetSelectedText_Err

   Set cpActivePane = VBInstance.ActiveCodePane
   cpActivePane.GetSelection llStartLine, llStartCol, llEndLine, llEndCol
   lsCodeLine = cpActivePane.CodeModule.Lines(llStartLine, 1)
   If llEndLine > llStartLine Then
      GetSelectedText = Mid$(lsCodeLine, llStartCol, Len(lsCodeLine) - llStartCol + 1)
   Else
      GetSelectedText = Mid$(lsCodeLine, llStartCol, llEndCol - llStartCol)
   End If

   'If window is linked (docked), then reposition code line within code pane
   'just in case the window overlays the selected text.
   If gwinAdvFndRpl.LinkedWindowFrame.Caption <> AdvFndRplCaption Then
      llTopLine = llStartLine - (cpActivePane.CountOfVisibleLines / 2)
      If llTopLine < 0 Then
         llTopLine = cpActivePane.CountOfVisibleLines / 2
      End If
      cpActivePane.TopLine = llTopLine
   End If

GetSelectedText_Err:

End Function
