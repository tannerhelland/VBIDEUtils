Attribute VB_Name = "Divers_Module"
' #VBIDEUtils#************************************************************
' * Author           : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 10/21/2002
' * Project Name     : VBIDEUtils
' * Module Name      : Divers_Module
' * Module Filename  : Divers.bas
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' * Example          :
' *
' * Screenshot       :
' *
' * See Also         :
' *
' * History          :
' *
' *
' **********************************************************************

Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd&, ByVal Msg&, ByVal wp&, ByVal lp&)
Private Declare Sub SetFocus Lib "user32" (ByVal hWnd&)
Private Declare Function GetParent Lib "user32" (ByVal hWnd&) As Long
Const WM_SYSKEYDOWN = &H104
Const WM_SYSKEYUP = &H105
Const WM_SYSCHAR = &H106
Const VK_F = 70                        ' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
Dim hwndMenu            As Long             ' needed to pass the menu keystrokes to VB

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

Global gsTemplate       As String
Global gsDevelopper     As String
Global gsWebSite        As String
Global gsEmail          As String
Global gsTelephone      As String
Global gsInline         As String
Global gsComment        As String
Global gsUserComment    As String
Global gsCommentString  As String
Global gsErrorHandler   As String

Global Const gsDefaultDevelopper = "removed"
Global Const gsDefaultWebSite = "http://www.ppreview.net"
Global Const gsDefaultEmail = "removed"

Global gcolFind         As Collection

Public Function CountLine(sText As String) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:49
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : CountLine
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         : Count the number of lines
   ' *
   ' *
   ' **********************************************************************

   Dim nPos             As Long
   Dim nI               As Long

   If Len(sText) = 0 Then
      CountLine = 0
      Exit Function
   End If

   nI = 1
   nPos = 1

   Do While (nPos > 0)
      ' *** Get the position of the CR
      nPos = InStr(nPos, sText, vbCrLf)

      If (nPos = 0) Then Exit Do

      nPos = nPos + 2
      nI = nI + 1

   Loop

   CountLine = nI

End Function

Public Function GetLine(sText As String, ByVal nLine As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:50
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : GetLine
   ' * Parameters       :
   ' *                    sText As String
   ' *                    ByVal nLine As Long
   ' **********************************************************************
   ' * Comments         : Get a specific line a complete set of line
   ' *
   ' *
   ' **********************************************************************

   Dim nPos             As Long
   Dim nStart           As String
   Dim nI               As Long
   Dim sTmp             As String

   nStart = 1

   nPos = 1

   ' *** get the right line
   For nI = 1 To nLine
      ' *** Get the position of the CR
      nPos = InStr(nPos, sText, vbCrLf)

      ' *** Only one line
      If nPos = 0 Then
         If nI = 1 Then
            sTmp = RTrim(sText)
         Else
            sTmp = RTrim(Mid$(sText, nStart))
         End If

         GetLine = sTmp
         Exit Function

      ElseIf nI = nLine Then
         sTmp = RTrim(Mid$(sText, nStart, nPos - nStart))

         GetLine = sTmp
         Exit Function
      End If

      nPos = nPos + 2
      nStart = nPos
   Next

End Function

Public Function GetNextLine(sText As String, nPreviousStart As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:50
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : GetNextLine
   ' * Parameters       :
   ' *                    sText As String
   ' *                    nPreviousStart As Long
   ' **********************************************************************
   ' * Comments         : Get a specific line a complete set of line
   ' *
   ' *
   ' **********************************************************************

   Dim nPos             As Long
   Dim nStart           As String
   Dim sTmp             As String

   nStart = nPreviousStart

   nPos = nPreviousStart

   ' *** get the next line
   ' *** Get the position of the CR
   nPos = InStr(nPos, sText, vbCrLf)

   ' *** Only one line
   If nPos = 0 Then
      sTmp = RTrim(Mid$(sText, nStart))

      GetNextLine = sTmp
      Exit Function

   Else
      sTmp = RTrim(Mid$(sText, nStart, nPos - nStart))

      nPreviousStart = nPos + 2

      GetNextLine = sTmp
      Exit Function
   End If

End Function

Public Sub Add_DefaultTemplate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:39
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : Add_DefaultTemplate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Add a default template
   ' *
   ' *
   ' **********************************************************************

   Dim sTemplate        As String

   sTemplate = "Template_Default"

   SaveSetting gsREG_APP, sTemplate, 1, twaCommentProgrammerName
   SaveSetting gsREG_APP, sTemplate, 2, twaCommentWebSite
   SaveSetting gsREG_APP, sTemplate, 3, twaCommentEMail
   SaveSetting gsREG_APP, sTemplate, 4, twaCommentDate
   SaveSetting gsREG_APP, sTemplate, 5, twaCommentTime
   SaveSetting gsREG_APP, sTemplate, 6, twaCommentModuleName
   SaveSetting gsREG_APP, sTemplate, 7, twaCommentModuleFileName
   SaveSetting gsREG_APP, sTemplate, 8, twaCommentProcedureName
   SaveSetting gsREG_APP, sTemplate, 9, twaCommentPurpose
   SaveSetting gsREG_APP, sTemplate, 10, twaCommentProcedureParameters
   SaveSetting gsREG_APP, sTemplate, 11, twaCommentPrefered
   SaveSetting gsREG_APP, sTemplate, 12, twaCommentPurpose
   SaveSetting gsREG_APP, sTemplate, 13, twaCommentSample
   SaveSetting gsREG_APP, sTemplate, 14, twaCommentSeeAlso
   SaveSetting gsREG_APP, sTemplate, 15, twaCommentHistory

   SaveSetting gsREG_APP, "Template", sTemplate, sTemplate

   gsTemplate = sTemplate

End Sub

Public Function RemoveAmpersand(sInput As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:39
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : RemoveAmpersand
   ' * Parameters       :
   ' *                    sInput As String
   ' **********************************************************************
   ' * Comments         : Remove all the & of a string
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim nI               As Integer

   sTmp = ""
   For nI = 1 To Len(sInput)
      If (Mid$(sInput, nI, 1) <> "&") Then sTmp = sTmp & Mid$(sInput, nI, 1)
   Next

   RemoveAmpersand = sTmp

End Function

Function InRunMode(VBInst As VBIDE.VBE) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : InRunMode
   ' * Parameters       :
   ' *                    VBInst As VBIDE.VBE
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)

End Function

Sub HandleKeyDown(ud As Object, KeyCode As Integer, Shift As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : HandleKeyDown
   ' * Parameters       :
   ' *                    ud As Object
   ' *                    KeyCode As Integer
   ' *                    Shift As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If Shift <> 4 Then Exit Sub
   If KeyCode < 65 Or KeyCode > 90 Then Exit Sub
   If VBInstance.DisplayModel = vbext_dm_SDI Then Exit Sub

   If hwndMenu = 0 Then hwndMenu = FindHwndMenu(ud.hWnd)
   PostMessage hwndMenu, WM_SYSKEYDOWN, KeyCode, &H20000000
   KeyCode = 0
   SetFocus hwndMenu

End Sub

Function FindHwndMenu(ByVal hWnd As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : FindHwndMenu
   ' * Parameters       :
   ' *                    ByVal hwnd As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim h                As Long

Loop2:
   h = GetParent(hWnd)
   If h = 0 Then FindHwndMenu = hWnd: Exit Function
   hWnd = h
   GoTo Loop2

End Function

Public Sub CloseUnusedWindows()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 9/02/99
   ' * Time             : 12:08
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : CloseUnusedWindows
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Close all unused windows
   ' *
   ' *
   ' **********************************************************************

   Dim pWindow          As Window

   For Each pWindow In VBInstance.Windows
      If Not pWindow Is VBInstance.ActiveWindow Then
         If pWindow.Type = 0 Or pWindow.Type = 1 Then pWindow.Close
      End If
   Next

End Sub

Public Function GetALine(sFile As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/12/1999
   ' * Time             : 12:31
   ' * Module Name      : Divers_Module
   ' * Module Filename  : Divers.bas
   ' * Procedure Name   : GetALine
   ' * Parameters       :
   ' *                    sFile As String
   ' **********************************************************************
   ' * Comments         :
   ' * Return the next line
   ' *
   ' **********************************************************************

   Dim sReturn          As String
   Dim nPos             As Integer

   sReturn = ""

   nPos = InStr(sFile, vbCrLf)

   If nPos = 0 Then
      ' *** Not found

      sReturn = ""
      sFile = ""

   Else
      ' *** Found
      sReturn = left$(sFile, nPos - 1)
      sFile = Mid$(sFile, nPos + 2)

   End If

   GetALine = sReturn

End Function

Public Sub InputNumeric(nKeyAscii As Integer, ctrName As Control, nDecimal As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/02/2000
   ' * Time             : 09:59
   ' * Module Name      : Lib_Module
   ' * Module Filename  : Lib.bas
   ' * Procedure Name   : InputNumeric
   ' * Parameters       :
   ' *                    nKeyAscii As Integer
   ' *                    ctrName As Control
   ' *                    nDecimal As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** N'accepte que les chiffres ***

   If nDecimal = True Then
      ' *** Test si point dÚcimal ***
      If nKeyAscii = Asc(".") Then
         If InStr(Trim$(ctrName.Text), ".") > 0 Then
            If InStr(Trim$(ctrName.SelText), ".") < 1 Then
               nKeyAscii = 0
               Beep
            End If
            Exit Sub
         End If
         Exit Sub
      End If
   End If
   If nKeyAscii = vbKeyBack Or (nKeyAscii > 47 And nKeyAscii < 58) Then
      ' *** Zone numÚrique ***
   Else
      nKeyAscii = 0
      Beep
   End If

End Sub
