Attribute VB_Name = "Comment_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 18/09/1998
' * Time             : 12:27
' * Module Name      : Comment_Module
' * Module Filename  : Comment.bas
' **********************************************************************
' * Comments         : Add Headers to modules, procedures
' *                    Un/Comment blocks of code
' *
' *
' **********************************************************************

Option Explicit

Global Const twaCommentProgrammerName = 1
Global Const twaCommentWebSite = 2
Global Const twaCommentEMail = 3
Global Const twaCommentTel = 4
Global Const twaCommentDate = 5
Global Const twaCommentTime = 6
Global Const twaCommentProjectName = 7
Global Const twaCommentModuleName = 8
Global Const twaCommentModuleFileName = 9
Global Const twaCommentProcedureName = 10
Global Const twaCommentProcedureParameters = 11
Global Const twaCommentPrefered = 12
Global Const twaCommentPurpose = 13
Global Const twaCommentSample = 14
Global Const twaCommentScreenshot = 15
Global Const twaCommentSeeAlso = 16
Global Const twaCommentHistory = 17

Public Sub InsertText(sText As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:05
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : InsertText
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         : Insert Text at the current position
   ' *
   ' *
   ' **********************************************************************

   Dim nStartLine       As Long
   Dim nStartColumn     As Long
   Dim nEndline         As Long
   Dim nEndColumn       As Long

   Dim sLine            As String

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

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
   sLine = cpCodePane.CodeModule.Lines(nStartLine, 1) & sText
   cpCodePane.CodeModule.ReplaceLine nStartLine, sLine
   cpCodePane.SetSelection nStartLine, nStartColumn + Len(sText), nEndline, nEndColumn + Len(sText)

End Sub

Public Sub InsertProcedureHeader()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 21:31
   ' * Project Name     : VBIDEUtils
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : InsertProcedureHeader
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Insert a procedure Header
   ' *
   ' *
   ' **********************************************************************

   Dim nStartLine       As Long
   Dim nStartColumn     As Long
   Dim nEndline         As Long
   Dim nEndColumn       As Long

   Dim sLine            As String

   Dim sProcName        As String
   Dim sDeclaration     As String
   Dim nLine            As Long
   Dim nProcStart       As Integer

   Dim nI               As Integer
   Dim nJ               As Integer
   Dim nK               As Integer

   Dim nItem            As Integer

   Dim sChar            As String * 1
   Dim sTmp             As String

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim bUseTabs         As Boolean
   Dim iIndentSpaces    As Integer
   Dim sIndent          As String

   Dim nParenthese      As Integer

   Dim bAnotherGo       As Boolean

   Dim nMaxLines        As Long

   Dim nLineInsertStart As Long

   On Error Resume Next

   bAnotherGo = False

   ' *** Get the active project
   Set prjProject = VBInstance.ActiveVBProject

   ' *** If we couldn't get it, quit
   If prjProject Is Nothing Then Exit Sub

   ' *** Try to find the active code pane
   Set cpCodePane = VBInstance.ActiveCodePane

   ' *** If we couldn't get it, quit
   If cpCodePane Is Nothing Then Exit Sub

   bUseTabs = (GetSetting(gsREG_APP, "Indent", "UseTabs", "N") = "Y")
   iIndentSpaces = Val(GetSetting(gsREG_APP, "Indent", "IndentSpaces", "3"))
   If bUseTabs Then
      sIndent = String(1, miTAB)
   Else
      sIndent = String(iIndentSpaces, " ")
   End If

   ' *** Get the active line
   cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn

   ' *** Get the name of the procedure
   sProcName = cpCodePane.CodeModule.ProcOfLine(nStartLine, vbext_pk_Proc)

   If sProcName = "" Then Exit Sub

Next_Pass:
   nProcStart = 0
   ' *** Get the line of the declaration procedure
   nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Proc)
   If (nProcStart = 0) Then
      nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Get)
      'If (cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Get) > nStartLine) Or (nStartLine > cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Get) + cpCodePane.CodeModule.ProcCountLines(sProcName, vbext_pk_Get)) Then
      If bAnotherGo = False Then
         nProcStart = 0
      End If
   End If
   If (nProcStart = 0) Then nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Let)
   If (nProcStart = 0) Then nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Set)
   If (nProcStart = 0) Then nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Get)
   nLine = nProcStart

   If bAnotherGo = False Then
      If (cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Let) > 0) Then
         ' *** A property
         If (cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Get) > 0) Then bAnotherGo = True
      End If
   Else
      bAnotherGo = False
   End If

   If nLine = 0 Then Exit Sub

   ' *** Check if not on more than 1 line
   sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
   Do While right$(sDeclaration, 1) = "_"
      nLine = nLine + 1
      sDeclaration = sDeclaration & Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
   Loop

   ' *** Get the first line of comment
   sLine = cpCodePane.CodeModule.Lines(nLine + 1, 1)

   ' *** Check if comment already placed
   If (left$(Trim$(sLine), Len("' #VBIDEUtils#")) = "' #VBIDEUtils#") Then
      ' *** Already done, modify the parameters

      If gsUserComment <> "" Then Exit Sub

      ' *** Get the number of lines in the procedure
      nMaxLines = cpCodePane.CodeModule.CountOfLines

      nLine = nLine + 1
      sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))

      Do While (nLine < nMaxLines) And (sLine <> "")
         Select Case UCase$(left$(sLine, 22))
            Case "' " & gsCommentString & " MODULE NAME      :":
               sLine = sIndent & "' " & gsCommentString & " Module Name      : " & cpCodePane.CodeModule.Parent.Name
               cpCodePane.CodeModule.ReplaceLine nLine, sLine

            Case "' " & gsCommentString & " MODULE FILENAME  :":
               sLine = sIndent & "' " & gsCommentString & " Module Filename  : " & GetFileName(cpCodePane.CodeModule.Parent.FileNames(1))
               cpCodePane.CodeModule.ReplaceLine nLine, sLine

            Case "' " & gsCommentString & " PROCEDURE NAME   :":
               sLine = sIndent & "' " & gsCommentString & " Procedure Name   : " & sProcName
               cpCodePane.CodeModule.ReplaceLine nLine, sLine

            Case "' " & gsCommentString & "                   ": ' *** Old parameter do delete
               cpCodePane.CodeModule.DeleteLines nLine, 1

            Case "' " & gsCommentString & " PARAMETERS       :":
               cpCodePane.CodeModule.DeleteLines nLine, 1

               sLine = sIndent & "' " & gsCommentString & " Parameters       : "
               cpCodePane.CodeModule.InsertLines nLine, sLine

               ' *** Delete all the parameters line
               nLine = nLine + 1
               sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
               Do While left$(sLine, 22) = "' " & gsCommentString & "                   "
                  cpCodePane.CodeModule.DeleteLines nLine, 1
                  sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
               Loop
               nLine = nLine - 1

               sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nProcStart, 1))
               nK = 1
               Do While right$(sDeclaration, 1) = "_"
                  sDeclaration = left$(sDeclaration, Len(sDeclaration) - 1) & Trim$(cpCodePane.CodeModule.Lines(nProcStart + nK, 1))
                  nK = nK + 1
               Loop

               sTmp = ""

               ' *** Skip the declaration of procedure
               Do While Len(sDeclaration) > 0
                  If left$(sDeclaration, 1) <> "(" Then
                     sDeclaration = right$(sDeclaration, Len(sDeclaration) - 1)
                  Else
                     sDeclaration = right$(sDeclaration, Len(sDeclaration) - 1)
                     Exit Do
                  End If
               Loop

               nJ = 1
               nParenthese = 0
               Do While Len(sDeclaration) >= nJ
                  sChar = Mid$(sDeclaration, nJ, 1)
                  If (sChar = ",") Or ((sChar = ")") And (nParenthese = 0)) Or ((sChar = "_") And (nJ = Len(sDeclaration))) Then
                     ' *** Go to next variable

                     If (Trim$(sTmp) = "") Then
                     Else
                        sLine = sIndent & "' " & gsCommentString & "                    " & Trim$(sTmp)
                        nLine = nLine + 1
                        cpCodePane.CodeModule.InsertLines nLine, sLine

                        sTmp = ""

                        ' *** End of variables
                        If (sChar = ")") Then
                           If (nParenthese = 0) Then
                              Exit Do
                           Else
                              nParenthese = nParenthese - 1
                           End If
                        End If

                        If (sChar = "_") And (nJ = Len(sDeclaration)) Then
                           nLine = nLine + 1
                           sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nProcStart, 1))
                           nJ = 1
                        End If
                     End If
                  Else
                     sTmp = sTmp & sChar
                     If (sChar = "(") Then nParenthese = nParenthese + 1

                     If (sChar = ")") Then
                        If (nParenthese = 0) Then
                        Else
                           nParenthese = nParenthese - 1
                        End If
                     End If

                  End If
                  nJ = nJ + 1
               Loop
         End Select
         nLine = nLine + 1
         sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
      Loop

      If bAnotherGo = True Then GoTo Next_Pass

      Exit Sub
   End If

   If gsUserComment <> "" Then
      nLine = nLine + 1
      cpCodePane.CodeModule.InsertLines nLine, gsUserComment
      Exit Sub
   End If

   ' *** Add a line
   If sLine <> "" Then
      'cpCodePane.CodeModule.InsertLines nLine + 1, ""
      sLine = ""
   End If
   nLine = nLine + 1
   nLineInsertStart = nLine
   sLine = sLine & sIndent & "' #VBIDEUtils#" & String(60, gsCommentString) & vbCrLf
   nLine = nLine + 1

   ' *** Go through all the parameters
   nI = 1
   nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))
   Do While nItem <> 0
      Select Case nItem
         Case twaCommentProgrammerName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Author           : " & gsDevelopper & vbCrLf
            nLine = nLine + 1

         Case twaCommentWebSite:
            sLine = sLine & sIndent & "' " & gsCommentString & " Web Site         : " & gsWebSite & vbCrLf
            nLine = nLine + 1

         Case twaCommentEMail:
            sLine = sLine & sIndent & "' " & gsCommentString & " E-Mail           : " & gsEmail & vbCrLf
            nLine = nLine + 1

         Case twaCommentTel:
            sLine = sLine & sIndent & "' " & gsCommentString & " Telephone        : " & gsTelephone & vbCrLf
            nLine = nLine + 1

         Case twaCommentDate:
            sLine = sLine & sIndent & "' " & gsCommentString & " Date             : " & Format(Now, "mm/dd/yyyy") & vbCrLf
            nLine = nLine + 1

         Case twaCommentTime:
            sLine = sLine & sIndent & "' " & gsCommentString & " Time             : " & Format(Now, "Short Time") & vbCrLf
            nLine = nLine + 1

         Case twaCommentProjectName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Project Name     : " & prjProject.Name & vbCrLf
            If err = 0 Then
               nLine = nLine + 1
            End If

         Case twaCommentModuleName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Module Name      : " & cpCodePane.CodeModule.Parent.Name & vbCrLf
            nLine = nLine + 1

         Case twaCommentModuleFileName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Module Filename  : " & GetFileName(cpCodePane.CodeModule.Parent.FileNames(1)) & vbCrLf
            nLine = nLine + 1

         Case twaCommentProcedureName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Procedure Name   : " & sProcName & vbCrLf
            nLine = nLine + 1

         Case twaCommentPurpose:
            sLine = sLine & sIndent & "' " & gsCommentString & " Purpose          : " & vbCrLf
            nLine = nLine + 1

         Case twaCommentProcedureParameters:
            sLine = sLine & sIndent & "' " & gsCommentString & " Parameters       : " & vbCrLf

            sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nProcStart, 1))
            nK = 1
            Do While right$(sDeclaration, 1) = "_"
               sDeclaration = left$(sDeclaration, Len(sDeclaration) - 1) & Trim$(cpCodePane.CodeModule.Lines(nProcStart + nK, 1))
               nK = nK + 1
            Loop

            sTmp = ""

            ' *** Skip the declaration of procedure
            Do While Len(sDeclaration) > 0
               If left$(sDeclaration, 1) <> "(" Then
                  sDeclaration = right$(sDeclaration, Len(sDeclaration) - 1)
               Else
                  sDeclaration = right$(sDeclaration, Len(sDeclaration) - 1)
                  Exit Do
               End If
            Loop

            nJ = 1
            nParenthese = 0
            Do While Len(sDeclaration) >= nJ
               sChar = Mid$(sDeclaration, nJ, 1)
               If (sChar = ",") Or ((sChar = ")") And (nParenthese = 0)) Or ((sChar = "_") And (nJ = Len(sDeclaration))) Then
                  ' *** Go to next variable

                  If (Trim$(sTmp) = "") Then
                  Else
                     sLine = sLine & sIndent & "' " & gsCommentString & "                    " & Trim$(sTmp) & vbCrLf
                     nLine = nLine + 1

                     sTmp = ""

                     ' *** End of variables
                     If (sChar = ")") Then
                        If (nParenthese = 0) Then
                           Exit Do
                        Else
                           nParenthese = nParenthese - 1
                        End If
                     End If

                     If (sChar = "_") And (nJ = Len(sDeclaration)) Then
                        nLine = nLine + 1
                        sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nProcStart, 1))
                        nJ = 1
                     End If
                  End If
               Else
                  sTmp = sTmp & sChar
                  If (sChar = "(") Then nParenthese = nParenthese + 1

                  If (sChar = ")") Then
                     If (nParenthese = 0) Then
                     Else
                        nParenthese = nParenthese - 1
                     End If
                  End If

               End If
               nJ = nJ + 1
            Loop
            nLine = nLine + 1

      End Select

      ' *** Get next parameter
      nI = nI + 1
      nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))
   Loop

   ' *** Add a line
   If gbRegistered Then
      sLine = sLine & sIndent & "' " & String(70, gsCommentString) & vbCrLf
   Else
      sLine = sLine & sIndent & "' * *************" & String(54, gsCommentString) & vbCrLf
   End If

   ' *** Comments
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & " Comments         : " & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** Add remaining
   nI = 1
   nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))
   Do While nItem <> 0
      Select Case nItem

         Case twaCommentSample:
            sLine = sLine & sIndent & "' " & gsCommentString & " Example          : " & vbCrLf
            nLine = nLine + 1
            sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf
            nLine = nLine + 1

         Case twaCommentScreenshot:
            sLine = sLine & sIndent & "' " & gsCommentString & " Screenshot       : " & vbCrLf
            nLine = nLine + 1
            sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf
            nLine = nLine + 1

         Case twaCommentSeeAlso:
            sLine = sLine & sIndent & "' " & gsCommentString & " See Also         : " & vbCrLf
            nLine = nLine + 1
            sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf
            nLine = nLine + 1

         Case twaCommentHistory:
            sLine = sLine & sIndent & "' " & gsCommentString & " History          : " & vbCrLf
            nLine = nLine + 1
            sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf
            nLine = nLine + 1

      End Select

      ' *** Get next parameter
      nI = nI + 1
      nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))

   Loop

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   If gbRegistered Then
      sLine = sLine & sIndent & "' " & String(70, gsCommentString) & vbCrLf
   Else
      sLine = sLine & sIndent & "' * *************" & String(54, gsCommentString) & vbCrLf
   End If

   If right$(sLine, 2) = vbCrLf Then sLine = left$(sLine, Len(sLine) - 2)

   cpCodePane.CodeModule.InsertLines nLineInsertStart, sLine

   If bAnotherGo = True Then GoTo Next_Pass

End Sub

Public Sub RemoveProcedureHeader()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 21:31
   ' * Project Name     : VBIDEUtils
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : RemoveProcedureHeader
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Remove a procedure Header
   ' *
   ' *
   ' **********************************************************************

   Dim nStartLine       As Long
   Dim nStartColumn     As Long
   Dim nEndline         As Long
   Dim nEndColumn       As Long

   Dim sLine            As String

   Dim sProcName        As String
   Dim sDeclaration     As String
   Dim nLine            As Long
   Dim nProcStart       As Integer

   Dim nI               As Integer
   Dim nJ               As Integer
   Dim nK               As Integer

   Dim nItem            As Integer

   Dim sChar            As String * 1
   Dim sTmp             As String

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim bUseTabs         As Boolean
   Dim iIndentSpaces    As Integer
   Dim sIndent          As String

   Dim nParenthese      As Integer

   Dim bAnotherGo       As Boolean

   Dim nMaxLines        As Long

   Dim nLineInsertStart As Long

   On Error Resume Next

   bAnotherGo = False

   ' *** Get the active project
   Set prjProject = VBInstance.ActiveVBProject

   ' *** If we couldn't get it, quit
   If prjProject Is Nothing Then Exit Sub

   ' *** Try to find the active code pane
   Set cpCodePane = VBInstance.ActiveCodePane

   ' *** If we couldn't get it, quit
   If cpCodePane Is Nothing Then Exit Sub

   bUseTabs = (GetSetting(gsREG_APP, "Indent", "UseTabs", "N") = "Y")
   iIndentSpaces = Val(GetSetting(gsREG_APP, "Indent", "IndentSpaces", "3"))
   If bUseTabs Then
      sIndent = String(1, miTAB)
   Else
      sIndent = String(iIndentSpaces, " ")
   End If

   ' *** Get the active line
   cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn

   ' *** Get the name of the procedure
   sProcName = cpCodePane.CodeModule.ProcOfLine(nStartLine, vbext_pk_Proc)

   If sProcName = "" Then Exit Sub

Next_Pass:
   nProcStart = 0
   ' *** Get the line of the declaration procedure
   nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Proc)
   If (nProcStart = 0) Then
      nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Get)
      'If (cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Get) > nStartLine) Or (nStartLine > cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Get) + cpCodePane.CodeModule.ProcCountLines(sProcName, vbext_pk_Get)) Then
      If bAnotherGo = False Then
         nProcStart = 0
      End If
   End If
   If (nProcStart = 0) Then nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Let)
   If (nProcStart = 0) Then nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Set)
   If (nProcStart = 0) Then nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, vbext_pk_Get)
   nLine = nProcStart

   If bAnotherGo = False Then
      If (cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Let) > 0) Then
         ' *** A property
         If (cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Get) > 0) Then bAnotherGo = True
      End If
   Else
      bAnotherGo = False
   End If

   If nLine = 0 Then Exit Sub

   ' *** Check if not on more than 1 line
   sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
   Do While right$(sDeclaration, 1) = "_"
      nLine = nLine + 1
      sDeclaration = sDeclaration & Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
      nProcStart = nLine
   Loop

   ' *** Get the first line of comment
   sLine = cpCodePane.CodeModule.Lines(nLine + 1, 1)

   ' *** Check if comment already placed
   If (left$(Trim$(sLine), Len("' #VBIDEUtils#")) = "' #VBIDEUtils#") Then
      ' *** Already done, Remove it

      If gsUserComment <> "" Then Exit Sub

      ' *** Get the number of lines in the procedure
      nMaxLines = cpCodePane.CodeModule.CountOfLines

      nLine = nLine + 2
      sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))

      Do While left$(sLine, 3) = "' *"
         nLine = nLine + 1
         sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
      Loop

      cpCodePane.CodeModule.DeleteLines nProcStart + 1, nLine - nProcStart - 1
      bAnotherGo = False
      nLine = nProcStart
   End If

   If gsUserComment <> "" Then
      nLine = nLine + 1
      cpCodePane.CodeModule.InsertLines nLine, gsUserComment
      Exit Sub
   End If

   ' *** Add a line
   If sLine <> "" Then
      'cpCodePane.CodeModule.InsertLines nLine + 1, ""
      sLine = ""
   End If
   nLine = nLine + 1
   nLineInsertStart = nLine
   sLine = sLine & sIndent & "' #VBIDEUtils#" & String(60, gsCommentString) & vbCrLf
   nLine = nLine + 1

   ' *** Go through all the parameters
   nI = 1
   nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))
   Do While nItem <> 0
      Select Case nItem
         Case twaCommentProgrammerName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Author           : " & gsDevelopper & vbCrLf
            nLine = nLine + 1

         Case twaCommentWebSite:
            sLine = sLine & sIndent & "' " & gsCommentString & " Web Site         : " & gsWebSite & vbCrLf
            nLine = nLine + 1

         Case twaCommentEMail:
            sLine = sLine & sIndent & "' " & gsCommentString & " E-Mail           : " & gsEmail & vbCrLf
            nLine = nLine + 1

         Case twaCommentTel:
            sLine = sLine & sIndent & "' " & gsCommentString & " Telephone        : " & gsTelephone & vbCrLf
            nLine = nLine + 1

         Case twaCommentDate:
            sLine = sLine & sIndent & "' " & gsCommentString & " Date             : " & Format(Now, "mm/dd/yyyy") & vbCrLf
            nLine = nLine + 1

         Case twaCommentTime:
            sLine = sLine & sIndent & "' " & gsCommentString & " Time             : " & Format(Now, "Short Time") & vbCrLf
            nLine = nLine + 1

         Case twaCommentProjectName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Project Name     : " & prjProject.Name & vbCrLf
            If err = 0 Then
               nLine = nLine + 1
            End If

         Case twaCommentModuleName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Module Name      : " & cpCodePane.CodeModule.Parent.Name & vbCrLf
            nLine = nLine + 1

         Case twaCommentModuleFileName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Module Filename  : " & GetFileName(cpCodePane.CodeModule.Parent.FileNames(1)) & vbCrLf
            nLine = nLine + 1

         Case twaCommentProcedureName:
            sLine = sLine & sIndent & "' " & gsCommentString & " Procedure Name   : " & sProcName & vbCrLf
            nLine = nLine + 1

         Case twaCommentPurpose:
            sLine = sLine & sIndent & "' " & gsCommentString & " Purpose          : " & vbCrLf
            nLine = nLine + 1

         Case twaCommentProcedureParameters:
            sLine = sLine & sIndent & "' " & gsCommentString & " Parameters       : " & vbCrLf

            sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nProcStart, 1))
            nK = 1
            Do While right$(sDeclaration, 1) = "_"
               sDeclaration = left$(sDeclaration, Len(sDeclaration) - 1) & Trim$(cpCodePane.CodeModule.Lines(nProcStart + nK, 1))
               nK = nK + 1
            Loop

            sTmp = ""

            ' *** Skip the declaration of procedure
            Do While Len(sDeclaration) > 0
               If left$(sDeclaration, 1) <> "(" Then
                  sDeclaration = right$(sDeclaration, Len(sDeclaration) - 1)
               Else
                  sDeclaration = right$(sDeclaration, Len(sDeclaration) - 1)
                  Exit Do
               End If
            Loop

            nJ = 1
            nParenthese = 0
            Do While Len(sDeclaration) >= nJ
               sChar = Mid$(sDeclaration, nJ, 1)
               If (sChar = ",") Or ((sChar = ")") And (nParenthese = 0)) Or ((sChar = "_") And (nJ = Len(sDeclaration))) Then
                  ' *** Go to next variable

                  If (Trim$(sTmp) = "") Then
                  Else
                     sLine = sLine & sIndent & "' " & gsCommentString & "                    " & Trim$(sTmp) & vbCrLf
                     nLine = nLine + 1

                     sTmp = ""

                     ' *** End of variables
                     If (sChar = ")") Then
                        If (nParenthese = 0) Then
                           Exit Do
                        Else
                           nParenthese = nParenthese - 1
                        End If
                     End If

                     If (sChar = "_") And (nJ = Len(sDeclaration)) Then
                        nLine = nLine + 1
                        sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nProcStart, 1))
                        nJ = 1
                     End If
                  End If
               Else
                  sTmp = sTmp & sChar
                  If (sChar = "(") Then nParenthese = nParenthese + 1

                  If (sChar = ")") Then
                     If (nParenthese = 0) Then
                     Else
                        nParenthese = nParenthese - 1
                     End If
                  End If

               End If
               nJ = nJ + 1
            Loop
            nLine = nLine + 1

      End Select

      ' *** Get next parameter
      nI = nI + 1
      nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))
   Loop

   ' *** Add a line
   If gbRegistered Then
      sLine = sLine & sIndent & "' " & String(70, gsCommentString) & vbCrLf
   Else
      sLine = sLine & sIndent & "' * *************" & String(54, gsCommentString) & vbCrLf
   End If

   ' *** Comments
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & " Comments         : " & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** Example
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & " Example          : " & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** See Also
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & " See Also         : " & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   sLine = sLine & sIndent & "' " & gsCommentString & vbCrLf

   ' *** Add a line
   nLine = nLine + 1
   If gbRegistered Then
      sLine = sLine & sIndent & "' " & String(70, gsCommentString) & vbCrLf
   Else
      sLine = sLine & sIndent & "' * *************" & String(54, gsCommentString) & vbCrLf
   End If

   If right$(sLine, 2) = vbCrLf Then sLine = left$(sLine, Len(sLine) - 2)

   cpCodePane.CodeModule.InsertLines nLineInsertStart, sLine

   If bAnotherGo = True Then GoTo Next_Pass

End Sub

Public Sub InsertModuleProcedureHeader()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/10/98
   ' * Time             : 17:26
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : InsertModuleProcedureHeader
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Locates the active module, add headers in all procedures
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim nI               As Integer

   Dim sName            As String

   Dim nLine            As Long
   Dim nStartLine       As Long

   Dim bNext            As Boolean

   Dim nMaxLine         As Long

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set prjProject = VBInstance.ActiveVBProject

   If prjProject Is Nothing Then Exit Sub

   Set cpCodePane = VBInstance.ActiveCodePane

   If cpCodePane Is Nothing Then Exit Sub

   If Not IsThereCode(cpCodePane.CodeModule) Then
      MsgBox "The current module does not contain any code to Error."
      Exit Sub
   End If

   Load frmProgress
   frmProgress.Show
   DoEvents
   frmProgress.MessageText = "Adding headers"
   frmProgress.Maximum = cpCodePane.CodeModule.members.Count
   DoEvents
   frmProgress.ZOrder
   nStartLine = cpCodePane.CodeModule.CountOfDeclarationLines
   nMaxLine = cpCodePane.CodeModule.CountOfLines
   For nI = 1 To cpCodePane.CodeModule.members.Count
      frmProgress.Progress = nI

      If (cpCodePane.CodeModule.members(nI).Type = 5) Or (cpCodePane.CodeModule.members(nI).CodeLocation <= cpCodePane.CodeModule.CountOfDeclarationLines) Then GoTo NEXT_ONE

      sName = cpCodePane.CodeModule.members(nI).Name

      frmProgress.MessageText = "Header on " & sName

      nLine = 0
      bNext = False
      ' *** Get the line of the declaration procedure
      nLine = cpCodePane.CodeModule.ProcBodyLine(sName, vbext_pk_Proc)
      If (nLine = 0) Then nLine = cpCodePane.CodeModule.ProcBodyLine(sName, vbext_pk_Get)
      If (nLine = 0) Then nLine = cpCodePane.CodeModule.ProcBodyLine(sName, vbext_pk_Let)
      If (nLine = 0) Then nLine = cpCodePane.CodeModule.ProcBodyLine(sName, vbext_pk_Set)

      cpCodePane.SetSelection nLine, 1, nLine, 1
      DoEvents
      Call InsertProcedureHeader

NEXT_ONE:
   Next
   nStartLine = cpCodePane.CodeModule.CountOfLines - 1
   cpCodePane.SetSelection nStartLine, 1, nStartLine, 1
   DoEvents
   Call InsertModuleHeader

   Unload frmProgress
   Set frmProgress = Nothing

End Sub

Public Sub RemoveModuleProcedureHeader()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/10/98
   ' * Time             : 17:26
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : RemoveModuleProcedureHeader
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Locates the active module, add headers in all procedures
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim nI               As Integer

   Dim sName            As String

   Dim nLine            As Long
   Dim nStartLine       As Long

   Dim bNext            As Boolean

   Dim nMaxLine         As Long

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set prjProject = VBInstance.ActiveVBProject

   If prjProject Is Nothing Then Exit Sub

   Set cpCodePane = VBInstance.ActiveCodePane

   If cpCodePane Is Nothing Then Exit Sub

   If Not IsThereCode(cpCodePane.CodeModule) Then
      MsgBox "The current module does not contain any code to Error."
      Exit Sub
   End If

   Load frmProgress
   frmProgress.Show
   DoEvents
   frmProgress.MessageText = "Adding headers"
   frmProgress.Maximum = cpCodePane.CodeModule.members.Count
   DoEvents
   frmProgress.ZOrder
   nStartLine = cpCodePane.CodeModule.CountOfDeclarationLines
   nMaxLine = cpCodePane.CodeModule.CountOfLines
   For nI = 1 To cpCodePane.CodeModule.members.Count
      frmProgress.Progress = nI

      If (cpCodePane.CodeModule.members(nI).Type = 5) Or (cpCodePane.CodeModule.members(nI).CodeLocation <= cpCodePane.CodeModule.CountOfDeclarationLines) Then GoTo NEXT_ONE

      sName = cpCodePane.CodeModule.members(nI).Name

      frmProgress.MessageText = "Replace Header on " & sName

      nLine = 0
      bNext = False
      ' *** Get the line of the declaration procedure
      nLine = cpCodePane.CodeModule.ProcBodyLine(sName, vbext_pk_Proc)
      If (nLine = 0) Then nLine = cpCodePane.CodeModule.ProcBodyLine(sName, vbext_pk_Get)
      If (nLine = 0) Then nLine = cpCodePane.CodeModule.ProcBodyLine(sName, vbext_pk_Let)
      If (nLine = 0) Then nLine = cpCodePane.CodeModule.ProcBodyLine(sName, vbext_pk_Set)

      cpCodePane.SetSelection nLine, 1, nLine, 1
      DoEvents
      Call RemoveProcedureHeader

NEXT_ONE:
   Next
   nStartLine = cpCodePane.CodeModule.CountOfLines - 1
   cpCodePane.SetSelection nStartLine, 1, nStartLine, 1
   DoEvents
   Call InsertModuleHeader

   Unload frmProgress
   Set frmProgress = Nothing

End Sub

Public Sub BlockOutCode()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : BlockOutCode
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Block out a complete set of code
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

   Dim nI               As Integer

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim bUseTabs         As Boolean
   Dim iIndentSpaces    As Integer
   Dim sIndent          As String

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
   sCode = cpCodePane.CodeModule.Lines(nStartLine, IIf(nEndline - nStartLine = 0, 1, nEndline - nStartLine))

   If (sCode = "") Then Exit Sub

   bUseTabs = (GetSetting(gsREG_APP, "Indent", "UseTabs", "N") = "Y")
   iIndentSpaces = Val(GetSetting(gsREG_APP, "Indent", "IndentSpaces", "3"))
   If bUseTabs Then
      sIndent = String(1, miTAB)
   Else
      sIndent = String(iIndentSpaces, " ")
   End If

   ' *** Add lines
   nLine = nStartLine
   sLine = sIndent & "' #HOut# " & String(20, gsCommentString)
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   sLine = sIndent & "' #HOut# Programmer Name  : " & gsDevelopper
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   sLine = sIndent & "' #HOut# Date             : " & Format(Now, "mm/dd/yyyy")
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   sLine = sIndent & "' #HOut# Time             : " & Format(Now, "Short Time")
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   sLine = sIndent & "' #HOut# Comment          : "
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   sLine = sIndent & "' #HOut# Comment          : "
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   sLine = sIndent & "' #HOut# Comment          : "
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   sLine = sIndent & "' #HOut# " & String(20, gsCommentString)
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   ' *** Comment each line
   For nI = 1 To CountLine(sCode)
      sLine = sIndent & "' #Out# " & GetLine(sCode, nI)
      cpCodePane.CodeModule.ReplaceLine nLine, sLine
      nLine = nLine + 1
   Next
   sLine = sIndent & "' #HOut# " & String(20, gsCommentString)
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   cpCodePane.SetSelection nLine, nStartColumn, nLine, nStartColumn

End Sub

Public Sub UnBlockOutCode()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 18/09/1998
   ' * Time             : 12:28
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : UnBlockOutCode
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : UnBlock out a complete set of code
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

   Dim nI               As Integer

   Dim nPos             As Integer

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim bUseTabs         As Boolean
   Dim iIndentSpaces    As Integer
   Dim sIndent          As String

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
   sCode = cpCodePane.CodeModule.Lines(nStartLine, nEndline - nStartLine + 1)

   If (sCode = "") Then Exit Sub

   bUseTabs = (GetSetting(gsREG_APP, "Indent", "UseTabs", "N") = "Y")
   iIndentSpaces = Val(GetSetting(gsREG_APP, "Indent", "IndentSpaces", "3"))
   If bUseTabs Then
      sIndent = String(1, miTAB)
   Else
      sIndent = String(iIndentSpaces, " ")
   End If

   nLine = nStartLine

   ' *** UnComment each line
   For nI = CountLine(sCode) To 1 Step -1
      sLine = GetLine(sCode, nI)
      If left$(Trim$(sLine), Len("' #Out#")) = "' #Out#" Then
         If Trim$(sLine) = "' #Out#" Then
            cpCodePane.CodeModule.ReplaceLine nLine + nI - 1, ""
         Else
            nPos = InStr(sLine, "' #Out#")
            cpCodePane.CodeModule.ReplaceLine nLine + nI - 1, right$(sLine, Len(sLine) - Len("' #Out#") - nPos)
         End If

      ElseIf left$(Trim$(sLine), Len("' #HOut#")) = "' #HOut#" Then
         cpCodePane.CodeModule.DeleteLines nLine + nI - 1, 1
      End If
   Next

   cpCodePane.SetSelection nLine, nStartColumn, nLine, nStartColumn

End Sub

Public Sub InsertModuleHeader()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 18/09/1998
   ' * Time             : 12:28
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : InsertModuleHeader
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Insert a module header
   ' *
   ' *
   ' **********************************************************************

   Dim sLine            As String
   Dim nLine            As Long

   Dim nI               As Integer
   Dim nItem            As Integer

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim sIndent          As String

   Dim nMaxLines        As Long

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

   ' *** Get the first line of module
   nLine = 0

   ' *** Get the first line of comment
   sLine = cpCodePane.CodeModule.Lines(nLine + 1, 1)

   ' *** Check if comment already placed
   If (left$(Trim$(sLine), Len("' #VBIDEUtils#")) = "' #VBIDEUtils#") Then
      ' *** Already done, modify the parameters

      ' *** Get the number of lines in the procedure
      nMaxLines = cpCodePane.CodeModule.CountOfLines

      nLine = nLine + 1
      sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))

      Do While (nLine < nMaxLines) And (sLine <> "")
         Select Case UCase$(left$(sLine, 22))
            Case "' " & gsCommentString & " MODULE NAME      :":
               sLine = sIndent & "' " & gsCommentString & " Module Name      : " & cpCodePane.CodeModule.Parent.Name
               cpCodePane.CodeModule.ReplaceLine nLine, sLine

            Case "' " & gsCommentString & " MODULE FILENAME  :":
               sLine = sIndent & "' " & gsCommentString & " Module Filename  : " & GetFileName(cpCodePane.CodeModule.Parent.FileNames(1))
               cpCodePane.CodeModule.ReplaceLine nLine, sLine

            Case "' " & gsCommentString & " PROJECT NAME     :":
               sLine = sIndent & "' " & gsCommentString & " Project Name     : " & prjProject.Name
               cpCodePane.CodeModule.ReplaceLine nLine, sLine
         End Select
         nLine = nLine + 1
         sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))
      Loop

      Exit Sub
   End If

   ' *** Add a line
   nLine = nLine + 1
   sLine = sIndent & "' #VBIDEUtils#" & String(60, gsCommentString)
   cpCodePane.CodeModule.InsertLines nLine, sLine
   nLine = nLine + 1

   ' *** Go through all the parameters
   nI = 1
   nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))
   Do While nItem <> 0
      Select Case nItem
         Case twaCommentProgrammerName:
            sLine = sIndent & "' " & gsCommentString & " Author           : " & gsDevelopper
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

         Case twaCommentWebSite:
            sLine = sIndent & "' " & gsCommentString & " Web Site         : " & gsWebSite
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

         Case twaCommentEMail:
            sLine = sIndent & "' " & gsCommentString & " E-Mail           : " & gsEmail
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

         Case twaCommentTel:
            sLine = sIndent & "' " & gsCommentString & " Telephone        : " & gsTelephone
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

         Case twaCommentDate:
            sLine = sIndent & "' " & gsCommentString & " Date             : " & Format(Now, "mm/dd/yyyy")
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

         Case twaCommentTime:
            sLine = sIndent & "' " & gsCommentString & " Time             : " & Format(Now, "Short Time")
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

         Case twaCommentProjectName:
            sLine = sIndent & "' " & gsCommentString & " Project Name     : " & prjProject.Name
            If err = 0 Then
               cpCodePane.CodeModule.InsertLines nLine, sLine
               nLine = nLine + 1
            End If

         Case twaCommentModuleName:
            sLine = sIndent & "' " & gsCommentString & " Module Name      : " & cpCodePane.CodeModule.Parent.Name
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

         Case twaCommentModuleFileName:
            sLine = sIndent & "' " & gsCommentString & " Module Filename  : " & GetFileName(cpCodePane.CodeModule.Parent.FileNames(1))
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

         Case twaCommentPurpose:
            sLine = sIndent & "' " & gsCommentString & " Purpose          : "
            cpCodePane.CodeModule.InsertLines nLine, sLine
            nLine = nLine + 1

      End Select

      ' *** Get next parameter
      nI = nI + 1
      nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))
   Loop

   ' *** Add a line
   If gbRegistered Then
      sLine = sIndent & "' " & String(70, gsCommentString)
   Else
      sLine = sIndent & "' * *************" & String(54, gsCommentString)
   End If
   cpCodePane.CodeModule.InsertLines nLine, sLine

   ' *** Comments
   nLine = nLine + 1
   sLine = sIndent & "' " & gsCommentString & " Comments         : "
   cpCodePane.CodeModule.InsertLines nLine, sLine

   ' *** Add a line
   nLine = nLine + 1
   sLine = sIndent & "' " & gsCommentString
   cpCodePane.CodeModule.InsertLines nLine, sLine

   ' *** Add a line
   nLine = nLine + 1
   sLine = sIndent & "' " & gsCommentString
   cpCodePane.CodeModule.InsertLines nLine, sLine

   ' *** Add remaining
   nI = 1
   nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))
   Do While nItem <> 0
      Select Case nItem
         Case twaCommentSample:
            sLine = sIndent & "' " & gsCommentString & " Example          : "
            nLine = nLine + 1
            cpCodePane.CodeModule.InsertLines nLine, sLine
            sLine = sIndent & "' " & gsCommentString
            nLine = nLine + 1
            cpCodePane.CodeModule.InsertLines nLine, sLine

         Case twaCommentScreenshot:
            sLine = sIndent & "' " & gsCommentString & " Screenshot       : "
            nLine = nLine + 1
            cpCodePane.CodeModule.InsertLines nLine, sLine
            sLine = sIndent & "' " & gsCommentString
            nLine = nLine + 1
            cpCodePane.CodeModule.InsertLines nLine, sLine

         Case twaCommentSeeAlso:
            sLine = sIndent & "' " & gsCommentString & " See Also         : "
            nLine = nLine + 1
            cpCodePane.CodeModule.InsertLines nLine, sLine
            sLine = sIndent & "' " & gsCommentString
            nLine = nLine + 1
            cpCodePane.CodeModule.InsertLines nLine, sLine

         Case twaCommentHistory:
            sLine = sIndent & "' " & gsCommentString & " History          : "
            nLine = nLine + 1
            cpCodePane.CodeModule.InsertLines nLine, sLine
            sLine = sIndent & "' " & gsCommentString
            nLine = nLine + 1
            cpCodePane.CodeModule.InsertLines nLine, sLine

      End Select

      ' *** Get next parameter
      nI = nI + 1
      nItem = CInt(GetSetting(gsREG_APP, gsTemplate, CStr(nI), "0"))

   Loop

   ' *** Add a line
   nLine = nLine + 1
   sLine = sIndent & "' " & gsCommentString
   cpCodePane.CodeModule.InsertLines nLine, sLine

   ' *** Add a line
   nLine = nLine + 1
   If gbRegistered Then
      sLine = sIndent & "' " & String(70, gsCommentString)
   Else
      sLine = sIndent & "' * *************" & String(54, gsCommentString)
   End If
   cpCodePane.CodeModule.InsertLines nLine, sLine

   ' *** Blank line
   nLine = nLine + 1
   sLine = ""
   cpCodePane.CodeModule.InsertLines nLine, sLine

End Sub

Public Sub InsertProjectHeader()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/10/98
   ' * Time             : 17:26
   ' * Module Name      : Comment_Module
   ' * Module Filename  : Comment.bas
   ' * Procedure Name   : InsertProjectHeader
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Locates the active project, add header in all procedures
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject

   Dim cmp              As VBComponent

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Get the active project
   Set prjProject = VBInstance.ActiveVBProject

   ' *** If we couldn't get it, quit
   If prjProject Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current project", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the project")
      Exit Sub
   End If

   For Each cmp In prjProject.VBComponents
      If (cmp.Type <> vbext_ct_RelatedDocument) And (cmp.Type <> vbext_ct_ResFile) Then
         If IsThereCode(cmp.CodeModule) Then
            cmp.CodeModule.CodePane.Show
            DoEvents

            ' *** Process this module
            Call InsertModuleHeader
            Call InsertModuleProcedureHeader
            cmp.CodeModule.CodePane.Window.Close
         End If
      End If
   Next

End Sub
