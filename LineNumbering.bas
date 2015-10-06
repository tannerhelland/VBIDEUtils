Attribute VB_Name = "LineNumbering_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 10/02/2000
' * Time             : 15:11
' * Module Name      : LineNumbering_Module
' * Module Filename  : LineNumbering.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Public Sub AddLineNumbering(bRemove As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/02/2000
   ' * Time             : 15:11
   ' * Module Name      : LineNumbering_Module
   ' * Module Filename  : LineNumbering.bas
   ' * Procedure Name   : AddLineNumbering
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim modCode          As VBIDE.CodeModule
   Dim nLine            As Long
   Dim nStart           As Long
   Dim nEnd             As Long
   Dim sProcName        As String
   Dim nProcType        As vbext_ProcKind
   Dim sLine            As String
   Dim sTemp            As String
   Dim bSkipNextLine    As Boolean
   Dim bOldLine         As Boolean
   Dim sTmp             As String

   Dim nCount           As Integer
   Dim nI               As Integer

   Dim nLineNum         As Long

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set modCode = VBInstance.ActiveCodePane.CodeModule

   frmProgress.MessageText = "Numbering all procedures"
   frmProgress.Maximum = 0
   frmProgress.Show

   nCount = 0
   For nI = 1 To modCode.members.Count
      If (modCode.members(nI).Type = 5) Or (modCode.members(nI).CodeLocation <= modCode.CountOfDeclarationLines) Then
         ' *** Declaration
      Else
         nCount = nCount + 1
      End If
   Next

   frmProgress.Maximum = nCount
   frmProgress.Show
   DoEvents
   nCount = 1

   ' *** First procedure is after all module-level declarations
   nLine = modCode.CountOfDeclarationLines + 1
   If nLine < modCode.CountOfLines Then
      ' *** Get name of first procedure
      sProcName = modCode.ProcOfLine(nLine, nProcType)

      Do While sProcName > ""
         frmProgress.Progress = nCount
         DoEvents
         nCount = nCount + 1

         nStart = modCode.ProcBodyLine(sProcName, nProcType)
         nEnd = modCode.ProcStartLine(sProcName, nProcType) _
            + modCode.ProcCountLines(sProcName, nProcType) - 2
         nLineNum = 1
         bSkipNextLine = True ' *** Skip procedure declaration line, but still process incase it has a line-contination
         For nLine = nStart To nEnd Step 1
            sLine = modCode.Lines(nLine, 1)
            sTemp = Trim$(sLine)
            ' *** Skip blank lines, comments, and compilation constants
            If Not bSkipNextLine Then
               If LineNumberEnd(sLine) = 0 Then
                  ' *** No Line numbering
                  sTmp = Trim$(LCase$(sLine))
                  If (sTmp = "") Or _
                     (sTmp = "end sub") Or _
                     (sTmp = "end function") Or _
                     (sTmp = "end property") Or _
                     (left(sTmp, 1) = "#") Or _
                     (left(sTmp, 1) = "'") Or _
                     (left(sTmp, 3) = "rem") Or _
                     (left(sTmp, 3) = "dim") Then
                     ' *** Do nothing
                  Else
                     If bRemove = False Then
                        ' *** Add line numbering
                        'modCode.ReplaceLine nLine, (nLine - nStart) & " " & sLine
                        modCode.ReplaceLine nLine, nLineNum & " " & sLine
                        nLineNum = nLineNum + 1
                     End If
                  End If
               Else
                  ' *** Else already has line number Remove it
                  bOldLine = False
                  Do While Len(sLine) > 0 And IsNumeric(left(sLine, 1))
                     sLine = Mid$(sLine, 2)
                     bOldLine = True
                  Loop
                  On Error Resume Next
                  If bOldLine Then sLine = Mid$(sLine, 2)

                  sTmp = Trim$(LCase$(sLine))
                  If (sTmp = "") Or _
                     (sTmp = "end sub") Or _
                     (sTmp = "end function") Or _
                     (sTmp = "end property") Or _
                     (left(sTmp, 1) = "#") Or _
                     (left(sTmp, 1) = "'") Or _
                     (left(sTmp, 3) = "rem") Or _
                     (left(sTmp, 3) = "dim") Then
                     ' *** Do nothing
                  Else
                     If bRemove = False Then
                        ' *** Add line numbering
                        modCode.ReplaceLine nLine, (nLine - nStart) & " " & sLine
                     Else
                        ' *** Remove line numbering
                        modCode.ReplaceLine nLine, sLine
                     End If
                  End If
               End If
            End If
            bSkipNextLine = SkipNextLine(sLine)
         Next

         ' *** Find next procedure
         nLine = modCode.ProcStartLine(sProcName, nProcType) + modCode.ProcCountLines(sProcName, nProcType)
         If nLine > modCode.CountOfLines Then Exit Do '-------------------------\/

         sProcName = modCode.ProcOfLine(nLine, nProcType)
      Loop
   End If

   Unload frmProgress
   Set frmProgress = Nothing

End Sub

Private Function SkipNextLine(sLine As String) As Boolean 'returns true if should not put a Line Number on next line of code 'because either Line Continuation or Select Case
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/02/2000
   ' * Time             : 15:11
   ' * Module Name      : LineNumbering_Module
   ' * Module Filename  : LineNumbering.bas
   ' * Procedure Name   : SkipNextLine
   ' * Parameters       :
   ' *                    sLine As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If InStr(sLine, " _") Then
      SkipNextLine = True
   ElseIf left$(LCase$(Trim$(sLine)), 6) = "select" Then
      SkipNextLine = True
   End If
End Function

Private Function LineNumberEnd(sLine As String) As Integer 'if sLine starts with a line number, returns the number of digits plus 1 'LineNumberEnd("12 If x = 0 Then") = 3
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/02/2000
   ' * Time             : 15:11
   ' * Module Name      : LineNumbering_Module
   ' * Module Filename  : LineNumbering.bas
   ' * Procedure Name   : LineNumberEnd
   ' * Parameters       :
   ' *                    sLine As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Useful for removing line numbers
   Dim nPos             As Long
   Dim sWord            As String
   nPos = InStr(sLine, " ")
   If nPos > 0 Then
      sWord = Trim$(left$(sLine, nPos - 1))
      If IsNumeric(sWord) Then
         LineNumberEnd = nPos
      End If
   Else
      sWord = Trim$(sLine)
      If IsNumeric(sWord) Then
         ' *** Line consists of line number only
         LineNumberEnd = Len(sLine)
      End If
   End If
End Function

