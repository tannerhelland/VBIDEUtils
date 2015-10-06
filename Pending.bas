Attribute VB_Name = "Pending_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:53
' * Module Name      : Pending_Module
' * Module Filename  : Pending.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private isSeparator(0 To 255) As Boolean
Private byteArrayString() As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal bytes As Long)

Public Sub GetPending()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:49
   ' * Module Name      : Pending_Module
   ' * Module Filename  : Pending.bas
   ' * Procedure Name   : GetPending
   ' * Parameters       :
   ' *
   ' **********************************************************************
   ' * Comments         : Show the pending code
   ' *
   ' *
   ' **********************************************************************
   Dim nStartLine       As Long
   Dim nStartColumn     As Long
   Dim nEndline         As Long
   Dim nEndColumn       As Long
   Dim nMaxLines        As Long

   Dim sLine            As String

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim sFirstWord       As String
   Dim sSecondWord      As String
   Dim sLastWord        As String

   Dim sFirstSearch     As String
   Dim sSecondSearch    As String
   Dim sLastSearch      As String
   Dim bDown            As Boolean

   Dim nLevel           As Integer

   Dim nCol             As Integer

   Dim bErrorHandler    As Boolean

   On Error Resume Next

   Const sSeparators = vbTab & " ,.:;!?""()=-><+&#" & vbCrLf
   Dim nI               As Integer

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

   cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn
   sLine = cpCodePane.CodeModule.Lines(nStartLine, 1)
   nMaxLines = cpCodePane.CodeModule.CountOfLines

   ' *** Get the needed words in the line
   Call GetWords(sLine, sFirstWord, sSecondWord, sLastWord)

   ' *** Analyse
   sFirstSearch = ""
   sSecondSearch = ""
   sLastSearch = ""

   ' *** Is this is the label of an error handler
   If Mid$(sLine, Len(sFirstWord) + 1, 1) = ":" Then
      bErrorHandler = True
      sFirstSearch = "on"
      sSecondSearch = "error"
      sLastSearch = sFirstWord
   Else
      bErrorHandler = False
   End If

   Select Case sFirstWord
      ' *** Search Down
      Case "on":
         If sSecondWord = "error" Then
            sFirstSearch = sLastWord
         End If
         bDown = True

      Case "goto":
         sFirstSearch = sLastWord
         bDown = True

      Case "if":
         sFirstSearch = "end"
         sSecondSearch = "if"
         bDown = True

      Case "#if":
         sFirstSearch = "#end"
         sSecondSearch = "if"
         bDown = True

      Case "select":
         Select Case sSecondWord
            Case "case":
               sFirstSearch = "end"
               sSecondSearch = "select"
               bDown = True
         End Select

      Case "with":
         sFirstSearch = "end"
         sSecondSearch = "with"
         bDown = True

      Case "for":
         sFirstSearch = "next"
         bDown = True

      Case "do":
         sFirstSearch = "loop"
         bDown = True

      Case "while":
         sFirstSearch = "wend"
         bDown = True

      Case "sub":
         sFirstSearch = "end"
         sSecondSearch = "sub"
         bDown = True

      Case "function":
         sFirstSearch = "end"
         sSecondSearch = "function"
         bDown = True

      Case "property":
         sFirstSearch = "end"
         sSecondSearch = "property"
         bDown = True

      Case "enum":
         sFirstSearch = "end"
         sSecondSearch = "enum"
         bDown = True

      Case "type":
         sFirstSearch = "end"
         sSecondSearch = "type"
         bDown = True

      Case "public", "private", "friend", "static":
         Select Case sSecondWord
            Case "sub":
               sFirstSearch = "end"
               sSecondSearch = "sub"
               bDown = True

            Case "function":
               sFirstSearch = "end"
               sSecondSearch = "function"
               bDown = True

            Case "property":
               sFirstSearch = "end"
               sSecondSearch = "property"
               bDown = True

            Case "enum":
               sFirstSearch = "end"
               sSecondSearch = "enum"
               bDown = True

            Case "type":
               sFirstSearch = "end"
               sSecondSearch = "type"
               bDown = True

         End Select

         ' *** Search up
      Case "#end"
         sFirstSearch = "#if"
         bDown = False

      Case "next":
         sFirstSearch = "for"
         bDown = False

      Case "loop":
         sFirstSearch = "do"
         bDown = False

      Case "wend":
         sFirstSearch = "while"
         bDown = False

      Case "end":
         Select Case sSecondWord
            Case "if"
               sFirstSearch = "if"
               bDown = False

            Case "select":
               sFirstSearch = "select"
               sSecondSearch = "case"
               bDown = False

            Case "enum":
               sFirstSearch = "enum|public|private|"
               sSecondSearch = "enum"
               bDown = False

            Case "type":
               sFirstSearch = "type|public|private|"
               sSecondSearch = "type"
               bDown = False

            Case "with":
               sFirstSearch = "with"
               bDown = False

            Case "sub":
               sFirstSearch = "sub|private|public|friend|static|"
               sSecondSearch = "sub"
               bDown = False

            Case "function":
               sFirstSearch = "function|private|public|friend|static|"
               sSecondSearch = "function"
               bDown = False

            Case "property":
               sFirstSearch = "property|private|public|friend|static|"
               sSecondSearch = "property"
               bDown = False

         End Select

   End Select

   If sFirstSearch = "" Then Exit Sub

   ' *** Begin the search
   nLevel = 0
   If bDown = True Then
      nStartLine = nStartLine + 1 ' *** Continues on next line
   Else
      nStartLine = nStartLine - 1 ' *** Continues on previous line
   End If

   ' *** Read through all the lines
   Do While (nStartLine >= 0) And (nStartLine <= nMaxLines)
      ' *** Do Some verifications
      If nStartLine < 0 Then Exit Do
      If (nLevel < 0) And (bDown) Then
         ' *** Error
         Call MsgBox("Pending line not found", vbCritical + vbOKOnly, "Pending problem")
         Exit Do
      End If

      sLine = cpCodePane.CodeModule.Lines(nStartLine, 1)

      ' *** Get the needed words in the line
      Call GetWords(sLine, sFirstWord, sSecondWord, sLastWord)

      Select Case sFirstWord
         ' *** Increase level

         Case "if":
            If (sLastWord = "then") Then nLevel = nLevel + 1

         Case "#if":
            nLevel = nLevel + 1

         Case "select":
            Select Case sSecondWord
               Case "case":
                  nLevel = nLevel + 1
            End Select

         Case "enum":
            nLevel = nLevel + 1

         Case "with":
            nLevel = nLevel + 1

         Case "for":
            nLevel = nLevel + 1

         Case "do":
            nLevel = nLevel + 1

         Case "while":
            nLevel = nLevel + 1

         Case "sub":
            nLevel = nLevel + 1

         Case "function":
            nLevel = nLevel + 1

         Case "property":
            nLevel = nLevel + 1

         Case "enum":
            nLevel = nLevel + 1

         Case "type":
            nLevel = nLevel + 1

         Case "public", "private", "friend", "static":
            Select Case sSecondWord
               Case "sub":
                  nLevel = nLevel + 1

               Case "function":
                  nLevel = nLevel + 1

               Case "property":
                  nLevel = nLevel + 1

               Case "enum":
                  nLevel = nLevel + 1

               Case "type":
                  nLevel = nLevel + 1

            End Select

            ' *** Decrease level
         Case "end":
            If sSecondWord <> "" Then nLevel = nLevel - 1

         Case "#end":
            nLevel = nLevel - 1

         Case "next":
            nLevel = nLevel - 1

         Case "loop":
            nLevel = nLevel - 1

         Case "wend":
            nLevel = nLevel - 1

      End Select

      ' *** Search in multiple value
      If bDown = False Then
         If InStr(sFirstSearch, sFirstWord & "|") > 0 Then
            sFirstWord = sFirstSearch
         End If
      End If

      If (sFirstWord = sFirstSearch) Then
         If bDown Then
            If ((sSecondSearch = "") And (nLevel = -1)) Or ((sSecondWord = sSecondSearch) And (nLevel = -1)) Or ((sSecondSearch = "") And (sSecondWord = sSecondSearch) And nLevel = 0) Then
               nCol = 1
               Do While (Mid$(sLine, nCol, 1) = " ") And (nCol < Len(sLine) - 1)
                  nCol = nCol + 1
               Loop
               cpCodePane.SetSelection nStartLine, nCol, nStartLine, nCol
               Exit Do
            End If
         Else
            If (((sSecondSearch = "") Or ((sSecondSearch <> "") And (sSecondWord = sSecondSearch))) And (nLevel = 1)) _
               Or ((sSecondWord = sSecondSearch) And (sLastWord = sLastSearch)) Then
               nCol = 1
               Do While (Mid$(sLine, nCol, 1) = " ") And (nCol < Len(sLine) - 1)
                  nCol = nCol + 1
               Loop
               cpCodePane.SetSelection nStartLine, nCol, nStartLine, nCol
               Exit Do
            End If
         End If
      End If

      If bDown = True Then
         nStartLine = nStartLine + 1 ' *** Continues on next line
      Else
         nStartLine = nStartLine - 1 ' *** Continues on previous line
      End If

   Loop

End Sub

Private Sub GetWords(ByVal sLine As String, sFirstWord As String, sSecondWord As String, sLastWord As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 24/11/98
   ' * Time             : 17:09
   ' * Module Name      : Pending_Module
   ' * Module Filename  : Pending.bas
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

