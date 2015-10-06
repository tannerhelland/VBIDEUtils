Attribute VB_Name = "Swap_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 13/10/99
' * Time             : 13:29
' * Module Name      : Swap_Module
' * Module Filename  : Swap.bas
' **********************************************************************
' * Comments         : Swap Left part and Right par of a =
' *
' *
' **********************************************************************

Option Explicit

Public Sub SwapEgual()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : Swap_Module
   ' * Module Filename  : Swap.bas
   ' * Procedure Name   : SwapEgual
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Swap Left part and Right par of a =
   ' *
   ' *
   ' **********************************************************************

   Dim nStartLine       As Long
   Dim nStartColumn     As Long
   Dim nEndline         As Long
   Dim nEndColumn       As Long

   Dim sLine            As String
   Dim sCode            As String

   Dim nLine            As Integer
   Dim sToken1          As String
   Dim sToken2          As String

   Dim sNew             As String

   Dim nI               As Integer

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim nPos             As Integer

   On Error Resume Next

   ' *** Get the active project
   Set prjProject = VBInstance.ActiveVBProject

   ' *** If we couldn't get it, quit
   If prjProject Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current project", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the project")
      Exit Sub
   End If

   ' *** Try to find the active code pane
   Set cpCodePane = VBInstance.ActiveCodePane

   ' *** If we couldn't get it, quit
   If cpCodePane Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current module", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the module")
      Exit Sub
   End If

   cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn
   If nEndColumn > 1 Then nEndline = nEndline + 1
   sCode = cpCodePane.CodeModule.Lines(nStartLine, IIf(nEndline - nStartLine = 0, 1, nEndline - nStartLine))

   If (sCode = "") Then Exit Sub

   ' *** Swap each line
   sNew = ""
   nLine = nStartLine
   For nI = 1 To CountLine(sCode)
      sLine = GetLine(sCode, nI)
      nPos = InStr(sLine, "=")
      If nPos > 0 Then
         sToken1 = RTrim(left(sLine, nPos - 1))
         sToken2 = right$(sLine, Len(sLine) - nPos)
         sNew = Space(Len(sToken1) - Len(LTrim(sToken1))) & Trim$(sToken2) & " = " & Trim$(sToken1)

         ' *** Replace that line
         cpCodePane.CodeModule.ReplaceLine nLine, sNew
      End If

      nLine = nLine + 1
   Next

   cpCodePane.SetSelection nStartLine, nStartColumn, nStartLine, nStartColumn

End Sub
