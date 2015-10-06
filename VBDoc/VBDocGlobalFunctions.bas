Attribute VB_Name = "GlobalFunctions"
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : GlobalFunctions
' * Module Filename  : VBDocGlobalFunctions.bas
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Purpose: This module contains all public functions and sub used by all class<BR>
'   in this project.
'Example: This is a test for
'   an example.<BR><BR>
'Code:      dim a as string
'Code:      a = 0
'Code:      b =1
'
'Another example
'Code:'In the Example Code you can write comments!!!
'Code:      Public Function Prova() as String
'Code:          Dim h as long 'comment
'Code:      End Function
'
'
'End of examples.
Option Explicit

'Purpose: Determine if the parameter strTemp is an Object or a Value during
'   the creation of the html file.<BR>
'Remarks: It check the gTypeValues Collection that contains the standard value type of
'   Visual Basic and the Enums and the UDTs defined in the project.
'Paramter: strType the type of the Object or value
Public Function isAnObject(ByVal strType As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : isAnObject
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal strType As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim temp             As Variant
   strType = UCase$(strType)
   isAnObject = True
   For Each temp In gTypeValues
      If strType = CStr(temp) Then isAnObject = False
   Next
End Function

'Purpose: Return tru if the member is static
'Parameter: Definition the definition elaborated that have or not the static keyword
Public Function IsStaticMemb(Definition As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : IsStaticMemb
   ' * Purpose          :
   ' * Parameters       :
   ' *                    Definition As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   IsStaticMemb = False
   If FirstLeftPart(Definition, "Static", False, True) Then IsStaticMemb = True
End Function

'Purpose: Read the file and storing the return the result text
Public Function ReadTextFile(FileName As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : ReadTextFile
   ' * Purpose          :
   ' * Parameters       :
   ' *                    FileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ReadTextFile

   Dim intFile          As Integer
   Dim strTemp          As String

   intFile = FreeFile
   ReadTextFile = ""
   Open FileName For Input As #intFile
   Do While Not EOF(intFile)
      Input #intFile, strTemp
      ReadTextFile = ReadTextFile & strTemp & vbCrLf
   Loop
   Close #intFile

EXIT_ReadTextFile:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_ReadTextFile:
   Select Case MsgBox("Error " & err.number & ": " & err.Description & vbCrLf & "in ReadTextFile", vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_ReadTextFile
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select

   Resume EXIT_ReadTextFile

End Function

'Purpose: Return the scope of the memb in string format
Public Function ScopeOfMemb(Definition As String, _
   Optional DefaultValue As vbext_Scope = vbext_Public) As vbext_Scope
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : ScopeOfMemb
   ' * Purpose          :
   ' * Parameters       :
   ' *                    Definition As String
   ' *                    Optional DefaultValue As vbext_Scope = vbext_Public
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error GoTo err_ScopeOfMemb
   ScopeOfMemb = DefaultValue
   If FirstLeftPart(Definition, "Public", False, True) Then
      ScopeOfMemb = vbext_Public
   ElseIf FirstLeftPart(Definition, "Private", False, True) Then
      ScopeOfMemb = vbext_Private
   ElseIf FirstLeftPart(Definition, "Friend", False, True) Then
      ScopeOfMemb = vbext_Friend
   End If
   Exit Function
err_ScopeOfMemb:
   ScopeOfMemb = DefaultValue
End Function

'Purpose: Write the HTML File using the Scripting.FileSystemObject
'Parameter FileName     The name of the file to be create
'Parameter strFile      The HTML text
Public Sub WriteTextFile(FileName As String, strFile As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : WriteTextFile
   ' * Purpose          :
   ' * Parameters       :
   ' *                    FileName As String
   ' *                    strFile As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_WriteTextFile

   Dim intFile          As Integer

   intFile = FreeFile
   Open FileName For Output As #intFile
   Print #intFile, strFile
   Close #intFile

EXIT_WriteTextFile:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_WriteTextFile:
   Select Case MsgBox("Error " & err.number & ": " & err.Description & vbCrLf & "in WriteTextFile", vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_WriteTextFile
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select

   Resume EXIT_WriteTextFile

End Sub

'Purpose: Get the current line of code and check the nexts lines of comment without
'   tags.<BR> It Returns the last line of comment.
'Parameter: IsExample if Is an Example test for Code Comments.
Public Function NextComments(VBcode As CodeModule, _
   ByRef CurrLine As Integer, _
   Optional IsExample As Boolean = False) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : NextComments
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    ByRef CurrLine As Integer
   ' *                    Optional IsExample As Boolean = False
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim strComment       As String
   NextComments = ""
   strComment = VBcode.Lines(CurrLine + 1, 1)

   Do While FirstLeftPart(strComment, "'", True, True)
      If Not IsKeyTag(strComment) Then
         If IsExample And (FirstLeftPart(CStr(strComment), gBLOCK_CODE, False, True) Or FirstLeftPart(CStr(strComment), gBLOCK_TEXT, False, True)) Then
            If FirstLeftPart(CStr(strComment), gBLOCK_CODE, False, True) Then
               strComment = Replace(strComment, " ", "&nbsp;")
               LeftPart strComment, gBLOCK_CODE, False, True
               If left$(strComment, 1) = ":" Then strComment = Mid$(strComment, 2)
               strComment = Replace(strComment, " ", "&nbsp;")
               NextComments = NextComments & "<FONT face=""Courier New"">" & HTMLCodeLine(strComment) & "</FONT>"
            Else
               LeftPart strComment, gBLOCK_TEXT, False, True
               If left$(strComment, 1) = ":" Then strComment = Mid$(strComment, 2)
               strComment = Replace(Trim$(strComment), " ", "&nbsp;")
               strComment = Replace(strComment, " ", "&nbsp;")
               NextComments = NextComments & strComment & "<BR>"
            End If
         Else
            If left$(strComment, 1) = "*" Then
               strComment = Trim$(Mid$(strComment, 2))
            End If
            If strComment <> "*********************************************************************" Then
               NextComments = NextComments & " " & strComment & vbCrLf
            End If
         End If
         CurrLine = CurrLine + 1
         strComment = VBcode.Lines(CurrLine + 1, 1)
      Else
         Exit Do
      End If
   Loop
   strComment = strComment

End Function

'Purpose: Get the current line of code and check the nexts lines of comment without
'   tags.<BR> It Returns the last line of comment.
'Parameter: IsExample if Is an Example test for Code Comments.
Public Function NextCommentsParameters(VBcode As CodeModule, _
   ByRef CurrLine As Integer, _
   Optional IsExample As Boolean = False) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : NextCommentsParameters
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    ByRef CurrLine As Integer
   ' *                    Optional IsExample As Boolean = False
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim strComment       As String
   NextCommentsParameters = ""
   strComment = VBcode.Lines(CurrLine + 1, 1)

   Do While FirstLeftPart(strComment, "'", True, True)
      If (left$(Trim$(strComment), 10) = "**********") Then Exit Do

      If (Not IsKeyTag(strComment)) And ((left$(Trim$(strComment), 10) <> "*         ") Or (Len(NextCommentsParameters) = 0)) Then
         If IsExample And (FirstLeftPart(CStr(strComment), gBLOCK_CODE, False, True) Or FirstLeftPart(CStr(strComment), gBLOCK_CODE, False, True)) Then
            If FirstLeftPart(CStr(strComment), gBLOCK_CODE, False, True) Then
               strComment = Replace(strComment, " ", "&nbsp;")
               LeftPart strComment, gBLOCK_CODE, False, True
               If left$(strComment, 1) = ":" Then strComment = Mid$(strComment, 2)
               strComment = Replace(strComment, " ", "&nbsp;")
               NextCommentsParameters = NextCommentsParameters & "<FONT face=""Courier New"">" & HTMLCodeLine(strComment) & "</FONT>"
            Else
               strComment = Replace(strComment, " ", "&nbsp;")
               LeftPart strComment, gBLOCK_TEXT, False, True
               If left$(strComment, 1) = ":" Then strComment = Mid$(strComment, 2)
               strComment = Replace(strComment, " ", "&nbsp;")
               NextCommentsParameters = NextCommentsParameters & strComment & "<BR>"
            End If
         Else
            If left$(strComment, 1) = "*" Then
               strComment = Trim$(Mid$(strComment, 2))
            End If
            If strComment <> "*********************************************************************" Then
               NextCommentsParameters = NextCommentsParameters & " " & strComment & vbCrLf
            End If
         End If
         CurrLine = CurrLine + 1
         strComment = VBcode.Lines(CurrLine + 1, 1)
      Else
         Exit Do
      End If
   Loop
   strComment = strComment

End Function

'Purpose: Return true if the first left part of the parameter
'   strComment is a recognized tag
Private Function IsKeyTag(ByVal strComment As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : IsKeyTag
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal strComment As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   IsKeyTag = False
   If FirstLeftPart(strComment, gBLOCK_AUTHOR) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_PURPOSE) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_DATE_CREATION) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_DATE_LAST_MOD) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_EXAMPLE) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_SEEALSO) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_SCREEENSHOT) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_PARAMETER) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_PROJECT) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_REMARKS) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_VERSION) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_NO_COMMENT) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_WEBSITE) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_EMAIL) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_TIME) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_TEL) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_PROCEDURE_NAME) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_MODULE_NAME) Then
      IsKeyTag = True
   ElseIf FirstLeftPart(strComment, gBLOCK_MODULE_FILE) Then
      IsKeyTag = True
   End If

End Function

'Purpose:Remove double blank from a string
Public Function RemoveDoubleBlank(myStr As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : RemoveDoubleBlank
   ' * Purpose          :
   ' * Parameters       :
   ' *                    myStr As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   myStr = Trim$(myStr)
   Do While InStr(1, myStr, "  ")
      myStr = Replace(myStr, "  ", " ")
   Loop
   RemoveDoubleBlank = myStr
End Function

Public Function ReplaceCRLF(sText As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : ReplaceCRLF
   ' * Purpose          :
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'Purpose: Replace all CRLF with <BR> and remove the last <BR>

   Dim sTmp             As String

   ' *** Remove the first vbrlf
   If left$(sText, 2) = vbCrLf Then sText = Mid$(sText, 3)

   sTmp = Replace(sText, vbCr, vbNullString)
   sTmp = Trim$(Replace(sTmp, vbLf, vbNullString))
   If Len(sTmp) > 0 Then
      ReplaceCRLF = Replace(sText, vbCrLf, "<BR>")
   Else
      ReplaceCRLF = vbNullString
   End If

End Function

'Purpose: Returns true if the left part of myStr is equal to myChars
'Remarks: This function is used for parsing the declaration and the comment<BR>
'   The myStr parameter is passed byRef and then it is truncated if the left part
'   of myStr is myChars and DeleteMyChars is True.
'Parameter: DeleteMyChars Remove the myChars string from myStr
Public Function FirstLeftPart(ByRef myStr As String, _
   myChars As String, _
   Optional MatchCase As Boolean = False, _
   Optional DeleteMyChars As Boolean = False) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : FirstLeftPart
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByRef myStr As String
   ' *                    myChars As String
   ' *                    Optional MatchCase As Boolean = False
   ' *                    Optional DeleteMyChars As Boolean = False
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim tempMystr        As String
   Dim nI               As Integer

   tempMystr = myStr

   tempMystr = Trim$(tempMystr)
   If Not MatchCase Then
      tempMystr = UCase$(tempMystr)
      myChars = UCase$(myChars)
   End If
   For nI = 1 To 5
      FirstLeftPart = (InStr(nI, tempMystr, myChars) = nI)
      If FirstLeftPart Then Exit For
   Next
   If FirstLeftPart And DeleteMyChars Then LeftPart myStr, myChars, MatchCase, True

   If left$(myStr, 1) = ":" Then
      myStr = Trim$(Mid$(myStr, 2))
   End If

End Function

'Purpose: This function returns the first part of myStr at the left of
'   the myChar character.<BR> The myStr variable is truncated of the left part.
'Remarks: The parameter DeleteChar indicate if the myChar character will be deleted.<BR>
'   If myChar is not encountered return all the myStr parameter.
'Example:
'Code:
'Code:       myVar = "lngTemp = 6"
'Code:       myResult = LeftPart(myvar,"=",TRUE,FALSE,FALSE)
'Code:
'       myReult is "lngTemp"
Public Function LeftPart(ByRef myStr As String, _
   myChar As String, _
   Optional MatchCase As Boolean = False, _
   Optional DeleteChar As Boolean = True, _
   Optional NotInQuotes As Boolean = True) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : LeftPart
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByRef myStr As String
   ' *                    myChar As String
   ' *                    Optional MatchCase As Boolean = False
   ' *                    Optional DeleteChar As Boolean = True
   ' *                    Optional NotInQuotes As Boolean = True
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim i                As Integer
   Dim inext            As Integer
   Dim InQuotes         As Boolean
   Dim tempMystr        As String

   tempMystr = UCase$(myStr)
   If Not MatchCase Then myChar = UCase$(myChar)
   If Not NotInQuotes Then
      i = InStr(1, tempMystr, myChar)
      If i > 0 Then
         LeftPart = Trim$(left$(myStr, i - 1))
         myStr = Trim$(right$(myStr, Len(myStr) - i + 1 - IIf(DeleteChar, Len(myChar), 0)))
      Else
         LeftPart = myStr
         myStr = ""
      End If
   Else
      inext = 0
      InQuotes = True
      Do While InQuotes
         InQuotes = False
         inext = InStr(inext + 1, tempMystr, myChar)
         If inext = 0 Then Exit Do
         For i = 1 To inext
            If Mid$(tempMystr, i, 1) = """" Then InQuotes = Not InQuotes
         Next
      Loop
      If inext > 0 Then
         LeftPart = Trim$(left$(myStr, inext - 1))
         myStr = Trim$(right$(myStr, Len(myStr) - inext + 1 - IIf(DeleteChar, Len(myChar), 0)))
      Else
         LeftPart = Trim$(myStr)
         myStr = ""
      End If
   End If
End Function

'Purpose: Return the String 'version' of the scope of a member<BR>
'   It's used during the creaton of HTML files.
Public Function ScopeToString(vScope As vbext_Scope) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : ScopeToString
   ' * Purpose          :
   ' * Parameters       :
   ' *                    vScope As vbext_Scope
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Select Case vScope
      Case vbext_Public
         ScopeToString = "Public"
      Case vbext_Private
         ScopeToString = "Private"
      Case vbext_Friend
         ScopeToString = "Friend"
   End Select
End Function

'Purpose: This function used by the Events,Enum,UDTs and Declaration Classes
'   retrieve the first line of comment before the definition
'Remarks: It's important a blank line between a declaration af a member
'   and the next initial comment for another member
'Parameter: VBcode      The Code Module for the member
'Parameter: intLine     The Line number of the definition
Public Function GetFirstComment(VBcode As CodeModule, intLine As Integer) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : GetFirstComment
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    intLine As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error GoTo GetFirstComment_Error
   GetFirstComment = intLine - 1
   Do While FirstLeftPart(VBcode.Lines(GetFirstComment - 1, 1), "'", True, False)
      GetFirstComment = GetFirstComment - 1
   Loop
   Exit Function
GetFirstComment_Error:

End Function

'Purpose: This complex function fix a bug of the VB IDE.<BR>
'   When there are some lines with the underscore at final of string
'   the next CodeLocation property are incorrect.<BR>
'   The if I want know the correct codeline i must count the underline at the
'   end of lines before ma declaration.
Private Function PrevUnderScore(VBModule As CodeModule, _
   CodeLocation As Integer) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : PrevUnderScore
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBModule As CodeModule
   ' *                    CodeLocation As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim i                As Integer
   Dim intTemp          As Integer
   Dim strTemp          As String
   Dim intRowToScroll   As Integer
   Dim intTempUnderScore As Integer

   PrevUnderScore = 0
   intTemp = 0
   intRowToScroll = CodeLocation
   Do
      intTempUnderScore = 0
      For i = intTemp + 1 To intTemp + intRowToScroll
         strTemp = Trim$(VBModule.Lines(i, 1))
         If right$(strTemp, 1) = "_" Then
            intTempUnderScore = intTempUnderScore + 1
         End If
      Next
      PrevUnderScore = PrevUnderScore + intTempUnderScore
      intTemp = intTemp + intRowToScroll
      intRowToScroll = intTempUnderScore
   Loop While intTempUnderScore > 0
End Function

'Purpose: It's used only by cEnum, then maybe it will move in this class
Public Function cLngP(var As Variant) As Long
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : cLngP
   ' * Purpose          :
   ' * Parameters       :
   ' *                    var As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error Resume Next
   cLngP = 0
   cLngP = CLng(var)
End Function

'Purpose: This function Parse a single line of code and returns it colored
Public Function HTMLCodeLine(line) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : HTMLCodeLine
   ' * Purpose          :
   ' * Parameters       :
   ' *                    Line
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim strTemp          As String
   Dim strComment       As String
   Dim strDelimiter     As String
   Dim strLine          As String

   strTemp = ""
   HTMLCodeLine = ""

   'read all string and replace spaces with &nbsp;
   strComment = line
   strComment = Replace(strComment, " ", "&nbsp;")
   'remove comments
   strLine = LeftPart(strComment, "'", True, False, True)

   strTemp = ""
   Do While Len(strLine) > 0

      strDelimiter = GetDelimiter(strLine)

      strTemp = LeftPart(strLine, LCase$(strDelimiter), False, True, True)
      strTemp = Replace(strTemp, "<", "&lt;")
      strTemp = Replace(strTemp, ">", "&gt;")
      If isKeyWord(strTemp) Then
         HTMLCodeLine = HTMLCodeLine & "<FONT color=Blue>" & strTemp & "</FONT>" & strDelimiter
      Else
         HTMLCodeLine = HTMLCodeLine & strTemp & strDelimiter
      End If
   Loop

   HTMLCodeLine = HTMLCodeLine & IIf(Len(strComment) > 0, "<font color=Green>" & strComment & "</font>", "") & "<br>"
End Function

'Purpose: Get the next valid delimiter in the source code line
Private Function GetDelimiter(line As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : GetDelimiter
   ' * Purpose          :
   ' * Parameters       :
   ' *                    Line As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim intCloseDel      As Integer
   Dim intTemp          As Integer

   intCloseDel = 256
   GetDelimiter = "&nbsp;"
   intTemp = InStr(1, line, "&nbsp;")
   If intTemp > 0 And intTemp < intCloseDel Then
      intCloseDel = intTemp
      GetDelimiter = "&nbsp;"
   End If
   intTemp = InStr(1, line, ",")
   If intTemp > 0 And intTemp < intCloseDel Then
      intCloseDel = intTemp
      GetDelimiter = ","
   End If
   intTemp = InStr(1, line, ")")
   If intTemp > 0 And intTemp < intCloseDel Then
      intCloseDel = intTemp
      GetDelimiter = ")"
   End If
   intTemp = InStr(1, line, "(")
   If intTemp > 0 And intTemp < intCloseDel Then
      intCloseDel = intTemp
      GetDelimiter = "("
   End If

End Function

'Purpose: If I have forget some keyword please add it in this list
Private Function isKeyWord(strWord As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : GlobalFunctions
   ' * Module Filename  : VBDocGlobalFunctions.bas
   ' * Procedure Name   : isKeyWord
   ' * Purpose          :
   ' * Parameters       :
   ' *                    strWord As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   isKeyWord = False
   Dim a                As Date
   Select Case UCase$(strWord)
      Case "AND", "ANY", "AS", "BOOLEAN", "BYREF", "BYTE", "BYVAL", _
         "CASE", "CONST", "CURRENCY", "DATE", "DECLARE", "DIM", "DO", _
         "DOUBLE", "EACH", "ELSE", "ELESEIF", "END", "ENUM", _
         "ERROR", "EVENT", "EXIT", "EXPLICIT", "FALSE", "FOR", "FRIEND", _
         "FUNCTION", "GET", "GOSUB", "GOTO", "IF", "IMPLEMENTS", _
         "IN", "INTEGER", "IS", "LET", "LIB", "LONG", "LOOP", "NEXT", _
         "NEW", "NOT", "OBJECT", "ON", "OPTION", "OPTIONAL", "OR", "PARAMARRAY", _
         "PRIVATE", "PROPERTY", "PUBLIC", "REDIM", "RESUME", "RETURN", "SELECT", _
         "SET", "STEP", "STOP", "STRING", "SUB", "THEN", "TIME", "TO", "TRUE", _
         "TYPE", "UNTIL", "VARIANT", "WEND", "WHILE", "WITH"
         '...... add other
         isKeyWord = True
   End Select
End Function

Public Function RemoveCRLF(sText As String) As String

   Dim sTmp             As String

   sTmp = Replace(sText, vbCr, "")
   sTmp = Trim$(Replace(sTmp, vbLf, " "))

   If right$(sTmp, 1) <> "." Then sTmp = sTmp & "."

   If Len(sTmp) < 3 Then sTmp = ""

   RemoveCRLF = Trim$(sTmp)

End Function

