Attribute VB_Name = "EnhancedErrorTreatment_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 26/04/99
' * Time             : 13:47
' * Module Name      : EnhancedErrorTreatment_Module
' * Module Filename  : EnhancedError.bas
' **********************************************************************
' * Comments         : Enhanced Error treatment
' *
' *
' **********************************************************************

Option Explicit

Private isSeparator(0 To 255) As Boolean
Private byteArrayString() As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal bytes As Long)

Public Sub InsertEnhancedProcedureError()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 21:31
   ' * Project Name     : VBIDEUtils
   ' * Module Name      : EnhancedErrorTreatment_Module
   ' * Module Filename  : EnhancedError.bas
   ' * Procedure Name   : InsertEnhancedProcedureError
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Insert an enhanced procedure Error handler
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

   Dim nMaxLines        As Long

   Dim nI               As Integer

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim bUseTabs         As Boolean
   Dim iIndentSpaces    As Integer
   Dim sIndent          As String

   Dim nProcCount       As Integer
   Dim nProcType        As Integer
   Dim sFileName        As String
   Dim sButton          As String

   Dim sEndSub          As String

   Dim sFirstWord       As String
   Dim sSecondWord      As String
   Dim sLastWord        As String

   Dim sAllParameters   As String
   Dim nK               As Long

   On Error Resume Next

   Const sSeparators = vbTab & " ,.:;!?""()=-><+&#" & vbCrLf

   For nI = 1 To Len(sSeparators)
      isSeparator(Asc(Mid$(sSeparators, nI, 1))) = True
   Next

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

   bUseTabs = (GetSetting(gsREG_APP, "Indent", "UseTabs", "N") = "Y")
   iIndentSpaces = Val(GetSetting(gsREG_APP, "Indent", "IndentSpaces", "3"))
   If bUseTabs Then
      sIndent = String(1, miTAB)
   Else
      sIndent = String(iIndentSpaces, " ")
   End If

   '*** Get button
   sButton = GetSetting(gsREG_APP, "Setting", "Button", "vbAbortRetryIgnore") & " + " & _
      GetSetting(gsREG_APP, "Setting", "Icon", "vbCritical") & " + " & _
      GetSetting(gsREG_APP, "Setting", "DefaultButton", "vbDefaultButton1")

   ' *** Get the active line
   cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn

   ' *** Get the name of the procedure
   sProcName = cpCodePane.CodeModule.ProcOfLine(nStartLine, vbext_pk_Proc)

   If sProcName = "" Then Exit Sub

   sFileName = cpCodePane.CodeModule.Parent.Name

   ' *** Get the line of the declaration procedure
   nProcType = vbext_pk_Proc
   nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, nProcType)
   If (nProcStart = 0) Then
      nProcType = vbext_pk_Get
      nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, nProcType)
      If (cpCodePane.CodeModule.ProcStartLine(sProcName, nProcType) > nStartLine) Or (nStartLine > cpCodePane.CodeModule.ProcStartLine(sProcName, nProcType) + cpCodePane.CodeModule.ProcCountLines(sProcName, nProcType)) Then
         nProcStart = 0
      End If
   End If
   If (nProcStart = 0) Then
      nProcType = vbext_pk_Let
      nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, nProcType)
   End If
   If (nProcStart = 0) Then
      nProcType = vbext_pk_Set
      nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, nProcType)
   End If
   nLine = nProcStart

   ' *** Check if not on more than 1 line
   sDeclaration = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))

   nK = 1
   Do While right(sDeclaration, 1) = "_"
      sDeclaration = left$(sDeclaration, Len(sDeclaration) - 1) & Trim$(cpCodePane.CodeModule.Lines(nProcStart + nK, 1))
      nK = nK + 1
   Loop

   ' *** Get the type of function
   Call GetWords(sDeclaration, sFirstWord, sSecondWord, sLastWord)

   If (sFirstWord = "sub") Or _
      ((sFirstWord = "static") And (sSecondWord = "sub")) Or _
      ((sFirstWord = "private") And (sSecondWord = "sub")) Or _
      ((sFirstWord = "friend") And (sSecondWord = "sub")) Or _
      ((sFirstWord = "public") And (sSecondWord = "sub")) Then
      sEndSub = "Sub"
   ElseIf (sFirstWord = "function") Or _
      ((sFirstWord = "static") And (sSecondWord = "function")) Or _
         ((sFirstWord = "private") And (sSecondWord = "function")) Or _
         ((sFirstWord = "friend") And (sSecondWord = "function")) Or _
         ((sFirstWord = "public") And (sSecondWord = "function")) Then
      sEndSub = "Function"
   ElseIf (sFirstWord = "property") Or _
      ((sFirstWord = "static") And (sSecondWord = "property")) Or _
         ((sFirstWord = "private") And (sSecondWord = "property")) Or _
         ((sFirstWord = "friend") And (sSecondWord = "property")) Or _
         ((sFirstWord = "public") And (sSecondWord = "property")) Then
      sEndSub = "Property"
   Else
      sEndSub = "Sub"
   End If

   Dim nJ               As Long
   Dim sTmp             As String
   Dim nParenthese      As Long
   Dim sChar            As String

   sTmp = ""

   ' *** Skip the declaration of procedure
   Do While Len(sDeclaration) > 0
      If left(sDeclaration, 1) <> "(" Then
         sDeclaration = right$(sDeclaration, Len(sDeclaration) - 1)
      Else
         sDeclaration = right$(sDeclaration, Len(sDeclaration) - 1)
         Exit Do
      End If
   Loop

   nJ = 1
   nParenthese = 0

   nLine = nLine + nK - 1

   sAllParameters = ""

   Do While Len(sDeclaration) >= nJ
      sChar = Mid$(sDeclaration, nJ, 1)
      If (sChar = ",") Or ((sChar = ")") And (nParenthese = 0)) Or ((sChar = "_") And (nJ = Len(sDeclaration))) Then
         ' *** Go to next variable

         If (Trim$(sTmp) = "") Then
         Else
            sAllParameters = sAllParameters & Trim$(sTmp)
            sTmp = Trim$(Replace(sTmp, "ByVal ", ""))
            sTmp = Trim$(Replace(sTmp, "ByRef ", ""))
            sTmp = Trim$(Replace(sTmp, "Optional ", ""))
            sTmp = Trim$(Replace(sTmp, "ParamArray ", ""))
            If InStr(sTmp, "As String") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As String", ""))
               sAllParameters = sAllParameters & " = """""" & " & sTmp & " & """""" ; "

            ElseIf InStr(sTmp, "As Byte") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As Byte", ""))
               sAllParameters = sAllParameters & " = "" & " & sTmp & " & "" ; "

            ElseIf InStr(sTmp, "As Integer") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As Integer", ""))
               sAllParameters = sAllParameters & " = "" & " & sTmp & " & "" ; "

            ElseIf InStr(sTmp, "As Long") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As Long", ""))
               sAllParameters = sAllParameters & " = "" & " & sTmp & " & "" ; "

            ElseIf InStr(sTmp, "As Single") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As Single", ""))
               sAllParameters = sAllParameters & " = "" & " & sTmp & " & "" ; "

            ElseIf InStr(sTmp, "As Double") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As Double", ""))
               sAllParameters = sAllParameters & " = "" & " & sTmp & " & "" ; "

            ElseIf InStr(sTmp, "As Decimal") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As Decimal", ""))
               sAllParameters = sAllParameters & " = "" & " & sTmp & " & "" ; "

            ElseIf InStr(sTmp, "As Date") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As Date", ""))
               sAllParameters = sAllParameters & " = "" & " & sTmp & " & "" ; "

            ElseIf InStr(sTmp, "As Boolean") > 0 Then
               sTmp = Trim$(Replace(sTmp, "As Boolean", ""))
               sAllParameters = sAllParameters & " = "" & CStr(" & sTmp & ") & "" ; "

            End If

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

   ' *** Get the first line of comment
   sLine = LCase$(Trim$(cpCodePane.CodeModule.Lines(nLine + 1, 1)))

   ' *** Go to the end of procedure comment
   Do While (sLine = "") Or (left(sLine, 1) = "'") Or (left(sLine, 3) = "rem")
      ' *** Check if an error handler is alread placed
      If (LCase$(left(Trim$(sLine), Len("' #VBIDEUtilsERROR#"))) = LCase$("' #VBIDEUtilsERROR#")) Then
         ' *** Already done, exit
         Exit Do
      End If
      nLine = nLine + 1
      sLine = LCase$(Trim$(cpCodePane.CodeModule.Lines(nLine + 1, 1)))
   Loop

   ' *** Get the number of lines in the procedure
   nMaxLines = cpCodePane.CodeModule.CountOfLines

   ' *** Check if an error handler is already placed
   If (LCase$(left(Trim$(sLine), Len("' #VBIDEUtilsERROR#"))) = LCase$("' #VBIDEUtilsERROR#")) Then
      ' *** Already done, modify it
      Do While (nLine < nMaxLines)
         sLine = Trim$(cpCodePane.CodeModule.Lines(nLine, 1))

         If LCase$(left$(sLine, 53)) = LCase$("Call clsAutoErrorHandler.Error_HandlerStartProcedure(") Then
            sLine = sIndent & "Call clsAutoErrorHandler.Error_HandlerStartProcedure(""" & sFileName & """, """ & sProcName & """, """ & sAllParameters & """)"
            cpCodePane.CodeModule.ReplaceLine nLine, sLine
            Exit Do
         End If

         nLine = nLine + 1
      Loop

   Else
      ' *** Add a new error handler

      If LCase$(Trim$(cpCodePane.CodeModule.Lines(nLine, 1))) <> "" Then nLine = nLine + 1

      ' *** Blank line
      sLine = ""
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      '----------------------------------------------------------------
      '---------------------Study HEADER section ----------
      '----------------------------------------------------------------
      ' *** Add a line  -
      sLine = sIndent & "' #VBIDEUtilsERROR#   HEADER Begin "
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add the automatic Error Handler
      sLine = sIndent & "Dim clsAutoErrorHandler As class_AutomateErrorHandler"
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1
      sLine = sIndent & "Set clsAutoErrorHandler = New class_AutomateErrorHandler"
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add ErrorHandlerEvent line
      sLine = sIndent & "Call clsAutoErrorHandler.Error_HandlerStartProcedure(""" & sFileName & """, """ & sProcName & """, """ & sAllParameters & """)"
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add the error handler
      sLine = sIndent & "On Error Goto ERROR_" & sProcName
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add a comment line
      sLine = sIndent & "' #VBIDEUtilsERROR#   HEADER End  | " & Now
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Blank line if needed
      sLine = LCase$(Trim$(cpCodePane.CodeModule.Lines(nLine, 1)))
      If sLine <> "" Then
         sLine = ""
         cpCodePane.CodeModule.InsertLines nLine, sLine
         nLine = nLine + 1
      End If

      ' *** Get start line for procedure
      nProcStart = cpCodePane.CodeModule.ProcBodyLine(sProcName, nProcType)

      ' *** Get the number of lines in the procedure
      nProcCount = cpCodePane.CodeModule.ProcCountLines(sProcName, nProcType)

      ' *** Get the active line
      cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn
      nLine = nProcStart

      ' *** Go to the end of the function
      Do While (nLine < nProcStart + nProcCount) And (sLine <> "end " & LCase$(sEndSub))
         nLine = nLine + 1
         sLine = LCase$(Trim$(cpCodePane.CodeModule.Lines(nLine + 1, 1)))
      Loop

      If LCase$(Trim$(cpCodePane.CodeModule.Lines(nLine, 1))) <> "" Then nLine = nLine + 1

      '----------------------------------------------------------------
      '---------------------Study FOOTER section ----------
      '----------------------------------------------------------------

      ' *** Blank line
      sLine = ""
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add a comment line  -
      sLine = sIndent & "' #VBIDEUtilsERROR#   FOOTER  Begin "
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add a stop line
      sLine = sIndent & "Goto END_" & sProcName
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Blank line
      sLine = ""
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add the error handler
      sLine = "ERROR_" & sProcName & ":"
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add a line
      sLine = sIndent & "Debug.Assert Err = 0"
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      sLine = sIndent & "Call clsAutoErrorHandler.Error_TreatErrorHandler(Error_GetOptionalConfiguration())"
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Blank line
      sLine = ""
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add a line
      sLine = sIndent & "END_" & sProcName & ":"
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Add a line
      sLine = sIndent & "' #VBIDEUtilsERROR#   FOOTER  End | " & Now
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** Blank line
      sLine = ""
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      ' *** On Error Resume Next
      sLine = sIndent & "On Error Resume Next"
      cpCodePane.CodeModule.InsertLines nLine, sLine
      nLine = nLine + 1

      If LCase$(Trim$(cpCodePane.CodeModule.Lines(nLine, 1))) <> "" Then
         ' *** Blank line
         sLine = ""
         cpCodePane.CodeModule.InsertLines nLine, sLine
         nLine = nLine + 1
      End If
   End If

End Sub

Private Sub GetWords(ByVal sLine As String, sFirstWord As String, sSecondWord As String, sLastWord As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 24/11/98
   ' * Time             : 17:09
   ' * Module Name      : EnhancedErrorTreatment_Module
   ' * Module Filename  : EnhancedError.bas
   ' * Procedure Name   : GetWords
   ' * Parameters       :
   ' *                    ByVal sLine As String
   ' *                    sFirstWord As String
   ' *                    sSecondWord As String
   ' *                    sLastWord As String
   ' **********************************************************************
   ' * Comments         : Get all the needed words
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Long
   Dim nJ               As Long

   sFirstWord = ""
   sSecondWord = ""
   sLastWord = ""

   sLine = Trim$(LCase$(sLine)) ' this line is not optim.

   ' *** Remove things in strings
   nI = InStr(1, sLine, """")
   Do Until nI = 0
      nJ = InStr(nI + 1, sLine, """")
      If nJ = 0 Then nJ = nI + 1
      sLine = left$(sLine, nI) & Mid$(sLine, nJ)
      nI = InStr(nI + 2, sLine, """")
   Loop

   ' *** Remove trailing comments from the line
   nI = InStr(1, sLine, "'")
   If nI > 0 Then sLine = left$(sLine, nI - 1)

   nI = InStr(sLine, Chr(39)) - 1
   If nI > 0 Then sLine = left$(sLine, nI)

   If left(sLine, 1) = Chr(39) Then Exit Sub
   If Len(sLine) = 0 Then Exit Sub

   Dim nSize            As Integer

   sLine = Trim$(sLine)

   nSize = Len(sLine)
   If nSize = 0 Then Exit Sub
   ReDim byteArrayString(1 To nSize)

   ' *** Copy string to byte array
   CopyMemory byteArrayString(1), ByVal sLine, nSize

   Dim nStartWord       As Integer
   Dim nSavPosition     As Integer

   ' *** Get FirstWord
   nStartWord = 1
   For nI = 1 To nSize
      If isSeparator(byteArrayString(nI)) Then
         sFirstWord = Space$(nI - nStartWord)
         CopyMemory ByVal sFirstWord, byteArrayString(nStartWord), nI - nStartWord
         nSavPosition = nI + 1
         Exit For
      End If
   Next

   If nSavPosition = 0 Then
      sFirstWord = Space$(nI - nStartWord)
      CopyMemory ByVal sFirstWord, byteArrayString(nStartWord), nI - nStartWord
      Exit Sub
   End If

   ' *** Get SecondWord
   nStartWord = nSavPosition
   For nI = nSavPosition To nSize
      If isSeparator(byteArrayString(nI)) Then
         sSecondWord = Space$(nI - nStartWord)
         CopyMemory ByVal sSecondWord, byteArrayString(nStartWord), nI - nStartWord
         nSavPosition = nI + 1
         Exit For
      End If
   Next

   If nSavPosition = nStartWord Then
      sSecondWord = Space$(nI - nStartWord)
      If (nI - nStartWord) > 0 Then
         CopyMemory ByVal sSecondWord, byteArrayString(nStartWord), nI - nStartWord
      End If
      sLastWord = sSecondWord
      Exit Sub
   End If

   ' *** Get LastWord
   For nI = nSize To nSavPosition Step -1
      If isSeparator(byteArrayString(nI)) Then
         sLastWord = Space$(nSize - nI)
         If (nSize - nI) > 0 Then
            CopyMemory ByVal sLastWord, byteArrayString(nI + 1), nSize - nI
         End If
         Exit For
      End If
   Next

   If (nI = nSavPosition - 1) Then
      If isSeparator(byteArrayString(nI)) Then
         sLastWord = Space$(nSize - nI)
         If (nSize - nI) > 0 Then
            CopyMemory ByVal sLastWord, byteArrayString(nI + 1), nSize - nI
         End If
      End If
   End If

End Sub

Public Sub InsertEnhancedModuleError()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/10/98
   ' * Time             : 17:26
   ' * Module Name      : EnhancedErrorTreatment_Module
   ' * Module Filename  : EnhancedError.bas
   ' * Procedure Name   : InsertEnhancedModuleError
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Locates the active module, add error handler in all procedures
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim nI               As Integer

   Dim nLine            As Long

   Dim nStartLine       As Long

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

   ' *** Check if the module contains any code
   If Not IsThereCode(cpCodePane.CodeModule) Then
      MsgBox "The current module does not contain any code to Error."
      Exit Sub
   End If

   ' *** Show the status bar user form.
   ' *** The activate of the userform runs the Adding Error handler
   ' *** routine, so it can update the status bar form as it progresses.
   Load frmProgress
   frmProgress.Show
   DoEvents
   frmProgress.MessageText = "Adding error handler"
   frmProgress.Maximum = cpCodePane.CodeModule.members.Count
   DoEvents
   frmProgress.ZOrder
   nStartLine = cpCodePane.CodeModule.CountOfDeclarationLines
   nMaxLines = cpCodePane.CodeModule.CountOfLines

   For nI = 1 To cpCodePane.CodeModule.members.Count
      frmProgress.Progress = nI

      If (cpCodePane.CodeModule.members(nI).Type = 5) Or (cpCodePane.CodeModule.members(nI).CodeLocation < cpCodePane.CodeModule.CountOfDeclarationLines) Then GoTo NEXT_ONE

      nLine = nStartLine
      Do While (nLine < nMaxLines) And (cpCodePane.CodeModule.ProcOfLine(nLine, vbext_pk_Proc) <> cpCodePane.CodeModule.members(nI).Name)
         nLine = nLine + 1
      Loop

      If nLine <> nMaxLines Then
         nStartLine = cpCodePane.CodeModule.members(nI).CodeLocation + cpCodePane.CodeModule.ProcCountLines(cpCodePane.CodeModule.members(nI).Name, vbext_pk_Proc)
         frmProgress.MessageText = "Error handler on " & cpCodePane.CodeModule.members(nI).Name

         cpCodePane.SetSelection nLine, 1, nLine, 1
         DoEvents
         Call InsertEnhancedProcedureError
      End If
NEXT_ONE:
   Next

   Unload frmProgress
   Set frmProgress = Nothing

End Sub

Public Sub InsertEnhancedProjectError()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/10/98
   ' * Time             : 17:26
   ' * Module Name      : EnhancedErrorTreatment_Module
   ' * Module Filename  : EnhancedError.bas
   ' * Procedure Name   : InsertEnhancedProjectError
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Locates the active project, add error handler in all modules
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim cmp              As VBComponent

   Dim nI               As Integer

   Dim nLine            As Long
   Dim nCountLine       As Long

   On Error Resume Next

   ' *** Get the active project
   Set prjProject = VBInstance.ActiveVBProject

   ' *** If we couldn't get it, quit
   If prjProject Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current project", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the project")
      Exit Sub
   End If

   ' *** Show the status bar user form.
   ' *** The activate of the userform runs the Adding Error handler
   ' *** routine, so it can update the status bar form as it progresses.
   Load frmProgress
   frmProgress.Show
   frmProgress.ZOrder

   For Each cmp In prjProject.VBComponents
      If (cmp.Type <> vbext_ct_RelatedDocument) And (cmp.Type <> vbext_ct_ResFile) Then
         If IsThereCode(cmp.CodeModule) Then
            ' *** Process this module

            Set cpCodePane = cmp.CodeModule.CodePane '

            nLine = 1
            nCountLine = cmp.CodeModule.CountOfLines    'cpCodePane.CodeModule.CountOfLines '

            For nI = 1 To cpCodePane.CodeModule.members.Count
               frmProgress.MessageText = "Adding Error handler " & cpCodePane.CodeModule.members(nI).Name
               Debug.Print "InsertProjectError -> ", cpCodePane.CodeModule.members(nI).Name
               DoEvents '
               Do While (nLine < nCountLine) And (cpCodePane.CodeModule.ProcOfLine(nLine, vbext_pk_Proc) <> cpCodePane.CodeModule.members(nI).Name)
                  nLine = nLine + 1
               Loop

               cpCodePane.SetSelection nLine, 1, nLine, 1

               DoEvents
               Call InsertEnhancedProcedureError
            Next
         End If
      End If
   Next

   Dim sMessage         As String
   sMessage = "You need add a reference to the Enhanced Error Handler"
   MsgBox sMessage, vbInformation

   Unload frmProgress
   Set frmProgress = Nothing

End Sub

