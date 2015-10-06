Attribute VB_Name = "ExtractString_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 13/10/99
' * Time             : 16:19
' * Module Name      : ExtractString_Module
' * Module Filename  : ExtractString.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Public colExceptLine    As Collection
Public colExceptAssign  As Collection
Public colExceptString  As Collection

Public gbExceptLine     As Boolean
Public gbExceptAssign   As Boolean
Public gbExceptString   As Boolean
Public gbExceptAllUpperCase As Boolean
Public gbGetMinimumSize As Boolean
Public gnGetMinimumSize As Integer

Public colExtracted     As Collection
Public colRefused       As Collection

Public Sub ExtractString()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:47
   ' * Module Name      : ExtractString_Module
   ' * Module Filename  : ExtractString.bas
   ' * Procedure Name   : ExtractString
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Extract all strings from a project
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ExtractString

   Dim sBuffer          As String
   Dim nI               As Long
   Dim nJ               As Long
   Dim nK               As Long

   Dim nPos             As Long
   Dim nPos1            As Long
   Dim nPreviousPos     As Long

   Dim nSelLen          As Long

   Dim nStart           As Long

   Dim nNextStart       As Long

   Dim nComponent       As Integer

   Dim sLine            As String
   Dim sTmp             As String

   Dim bExceptString    As Boolean

   Dim cmp              As VBComponent

   Dim nFile            As Integer

   Dim nPosChar         As Integer

   Dim nLines           As Long

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set frmProgress = Nothing
   Set colExtracted = Nothing
   Set colRefused = Nothing

   ' *** If we couldn't get it, quit
   If VBInstance.ActiveVBProject Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current project", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the project")
      Exit Sub
   End If

   Set colExtracted = New Collection
   Set colRefused = New Collection

   frmProgress.MessageText = "Extracting strings from the project"
   Load frmProgress
   frmProgress.Show
   frmProgress.ZOrder
   DoEvents

   frmProgress.Maximum = VBInstance.ActiveVBProject.VBComponents.Count
   nComponent = 1

   ' *** Get all components in the project
   For Each cmp In VBInstance.ActiveVBProject.VBComponents
      frmProgress.Progress = nComponent

      ' *** Get the code of the component
      On Error Resume Next
      nFile = FreeFile
      Open cmp.FileNames(1) For Binary Access Read As #nFile
      sBuffer = Input(LOF(nFile), nFile)
      Close #nFile
      On Error GoTo ERROR_ExtractString

      nNextStart = 1
      nLines = CountLine(sBuffer)
      For nI = 1 To nLines
         sLine = GetNextLine(sBuffer, nNextStart)
         sLine = Trim$(sLine)

         nPos = InStr(sLine, Chr$(34))
         If nPos > 0 Then
            ' *** Inside a string
            nPreviousPos = 1
            bExceptString = False
Get_String:
            ' *** Check if we need to except all the line
            ' *** Check if we are in a comment
            nPos1 = InStr(sLine, "'")
            If (nPos1 > 0) And (nPos1 < nPos) Then GoTo Next_Line

            ' *** Check if we need to except the line
            If gbExceptLine Then
               For nK = 1 To colExceptLine.Count
                  If InStr(sLine, colExceptLine(nK)) > 0 Then GoTo Next_Line
               Next
            End If

            ' *** Check if we need to except the string due to something before
            If gbExceptAssign Then
               For nK = 1 To colExceptAssign.Count
                  nPos1 = InStr(sLine, colExceptAssign(nK))
                  If (nPos1 > 0) And (nPos1 < nPos) Then
                     bExceptString = True
                     nPreviousPos = nPos
                     Exit For
                  End If
               Next
            End If

            ' *** We can continue
            nSelLen = 1
            nStart = nPos + 1
            For nJ = nStart To Len(sLine)
               If (Mid$(sLine, nJ, 1) = Chr(34)) Then
                  If Mid$(sLine, nJ + 1, 1) = Chr(34) Then
                     nSelLen = nSelLen + 1
                  Else
                     nSelLen = nSelLen - 1
                     Exit For
                  End If
               Else
                  nSelLen = nSelLen + 1
               End If
            Next
            On Error Resume Next

            sTmp = Mid$(sLine, nStart, nSelLen)
            If (sTmp <> "") Then
               ' *** Check for string containing exceptions
               If gbExceptString Then
                  If bExceptString = False Then
                     For nK = 1 To colExceptString.Count
                        nPos1 = InStr(UCase$(sTmp), colExceptString(nK))
                        If (nPos1 > 0) Then
                           bExceptString = True
                           Exit For
                        End If
                     Next
                  End If
               End If

               ' *** Check if all char are in uppercase
               If (bExceptString = False) And (gbExceptAllUpperCase) Then
                  bExceptString = (UCase$(sTmp) = sTmp)
               End If

               ' *** Check the minimum size
               If gbGetMinimumSize And (Len(sTmp) < gnGetMinimumSize) Then bExceptString = True

               If (bExceptString = False) Then
                  nPosChar = InStr(sTmp, "µ")
                  If nPosChar > 0 Then sTmp = left$(sTmp, nPosChar - 1)

                  colExtracted.Add sTmp, sTmp
               Else
                  colRefused.Add sTmp, sTmp
               End If
            End If
            err = 0

            ' *** Continue the line for possible second string
            If nJ < Len(sLine) Then
               bExceptString = False
               nPos = InStr(nJ + 1, sLine, Chr$(34))
               If nPos > 0 Then GoTo Get_String
            End If

            On Error GoTo ERROR_ExtractString

         End If
Next_Line:

      Next
      nComponent = nComponent + 1
   Next
   frmProgress.Progress = VBInstance.ActiveVBProject.VBComponents.Count

EXIT_ExtractString:
   Unload frmProgress

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ExtractString:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in ExtractString", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_ExtractString
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_ExtractString

End Sub

Public Sub InitExtractString()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 11:53
   ' * Module Name      : ExtractString_Module
   ' * Module Filename  : ExtractString.bas
   ' * Procedure Name   : InitExtractString
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Initialise the string extraction
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_InitExtractString

   Dim sLine            As String
   Dim nFile            As Integer
   Dim sBuffer          As String

   Dim nI               As Long

   ' *** Get all the line exception
   Set colExceptLine = Nothing
   Set colExceptLine = New Collection

   On Error Resume Next
   nFile = FreeFile
   Open App.Path & "\Extract1.txt" For Binary Access Read As #nFile
   sBuffer = Input(LOF(nFile), nFile)
   Close #nFile
   On Error GoTo ERROR_InitExtractString

   For nI = 1 To CountLine(sBuffer)
      sLine = GetLine(sBuffer, nI)
      On Error Resume Next
      If Trim$(sLine) <> "" Then colExceptLine.Add sLine, sLine
      On Error GoTo ERROR_InitExtractString
   Next

   ' *** Get all the assignment string exception
   Set colExceptAssign = Nothing
   Set colExceptAssign = New Collection

   On Error Resume Next
   nFile = FreeFile
   Open App.Path & "\Extract2.txt" For Binary Access Read As #nFile
   sBuffer = Input(LOF(nFile), nFile)
   Close #nFile

   For nI = 1 To CountLine(sBuffer)
      sLine = GetLine(sBuffer, nI)
      On Error Resume Next
      If Trim$(sLine) <> "" Then colExceptAssign.Add sLine, sLine
      On Error GoTo ERROR_InitExtractString
   Next

   ' *** Get all the string exception
   Set colExceptString = Nothing
   Set colExceptString = New Collection

   On Error Resume Next
   nFile = FreeFile
   Open App.Path & "\Extract3.txt" For Binary Access Read As #nFile
   sBuffer = Input(LOF(nFile), nFile)
   Close #nFile

   For nI = 1 To CountLine(sBuffer)
      sLine = GetLine(sBuffer, nI)
      On Error Resume Next
      If Trim$(sLine) <> "" Then colExceptString.Add sLine, sLine
      On Error GoTo ERROR_InitExtractString
   Next

   gbExceptLine = (GetSetting(gsREG_APP, "ExtractString", "ExceptLine", "Y") = "Y")
   gbExceptAssign = (GetSetting(gsREG_APP, "ExtractString", "ExceptAssign", "Y") = "Y")
   gbExceptString = (GetSetting(gsREG_APP, "ExtractString", "ExceptString", "Y") = "Y")

   gbExceptAllUpperCase = (GetSetting(gsREG_APP, "ExtractString", "ExceptAllUpperCase", "Y") = "Y")

   gbGetMinimumSize = (GetSetting(gsREG_APP, "ExtractString", "IsGetMinimumSize", "Y") = "Y")
   gnGetMinimumSize = GetSetting(gsREG_APP, "ExtractString", "NumGetMinimumSize", "2")

EXIT_InitExtractString:
   DoEvents

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_InitExtractString:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in InitExtractString", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_InitExtractString
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_InitExtractString

End Sub

Public Sub InternationalizeProject(sDB As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/10/98
   ' * Time             : 14:47
   ' * Module Name      : InternationalizeProject_Module
   ' * Module Filename  : InternationalizeProject.bas
   ' * Procedure Name   : InternationalizeProject
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Internationalize all the project
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_InternationalizeProject

   Dim sBuffer          As String
   Dim nI               As Long
   Dim nJ               As Long
   Dim nK               As Long

   Dim nPos             As Long
   Dim nPos1            As Long
   Dim nPreviousPos     As Long

   Dim nSelLen          As Long

   Dim nStart           As Long

   Dim nNextStart       As Long

   Dim nComponent       As Integer

   Dim sLine            As String
   Dim sTmp             As String
   Dim sTmp2            As String

   Dim bExceptString    As Boolean

   Dim cmp              As VBComponent

   Dim nFile            As Integer

   Dim bHeader          As Boolean

   Dim DBTranslate      As Database
   Dim sSQL             As String
   Dim record           As Recordset

   Dim sNewFile         As String
   Dim sNewLine         As String

   Dim sName            As String

   Dim nPosChar         As Integer

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set frmProgress = Nothing
   Set colExtracted = Nothing

   ' *** If we couldn't get it, quit
   If VBInstance.ActiveVBProject Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current project", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the project")
      Exit Sub
   End If

   Set DBTranslate = DAO.OpenDatabase(sDB)

   frmProgress.MessageText = "Internationalizing the current project"
   Load frmProgress
   frmProgress.Show
   frmProgress.ZOrder
   DoEvents

   frmExtractString.lstResult.AddItem "Internationalizing " & VBInstance.ActiveVBProject.Name
   frmExtractString.lstResult.AddItem ""
   frmExtractString.lstResult.AddItem VBInstance.ActiveVBProject.VBComponents.Count & " components to internationalize"
   frmExtractString.lstResult.AddItem ""

   frmProgress.Maximum = VBInstance.ActiveVBProject.VBComponents.Count
   nComponent = 1

   ' *** Get all components in the project
   For Each cmp In VBInstance.ActiveVBProject.VBComponents
      sName = cmp.Name
      frmExtractString.lstResult.AddItem "Working on " & sName
      frmProgress.Progress = nComponent

      ' *** Get the code of the component
      On Error Resume Next
      nFile = FreeFile
      Open cmp.FileNames(1) For Binary Access Read As #nFile
      sBuffer = Input(LOF(nFile), nFile)
      Close #nFile
      On Error GoTo ERROR_InternationalizeProject

      bHeader = True
      sNewFile = ""

      nNextStart = 1
      For nI = 1 To CountLine(sBuffer)
         sLine = GetNextLine(sBuffer, nNextStart)
         sNewLine = sLine

         nPos = InStr(sLine, Chr$(34))
         If nPos > 0 Then
            ' *** Check if we are in the file header
            If InStr(sLine, "Attribute VB_Name = ") > 0 Then bHeader = False

            ' *** Inside a string
            nPreviousPos = 1
            bExceptString = False

Get_String:
            ' *** Check if we need to except all the line
            ' *** Check if we are in a comment
            nPos1 = InStr(sLine, "'")
            If (nPos1 > 0) And (nPos1 < nPos) Then GoTo Next_Line

            ' *** Check if we need to except the line
            If gbExceptLine Then
               For nK = 1 To colExceptLine.Count
                  If InStr(sLine, colExceptLine(nK)) > 0 Then GoTo Next_Line
               Next
            End If

            ' *** Check if we need to except the string due to something before
            If gbExceptAssign Then
               For nK = 1 To colExceptAssign.Count
                  nPos1 = InStr(sLine, colExceptAssign(nK))
                  If (nPos1 > 0) And (nPos1 < nPos) Then
                     bExceptString = True
                     nPreviousPos = nPos
                     Exit For
                  End If
               Next
            End If

            ' *** We can continue
            nSelLen = 1
            nStart = nPos + 1
            For nJ = nStart To Len(sLine)
               If (Mid$(sLine, nJ, 1) = Chr(34)) Then
                  If Mid$(sLine, nJ + 1, 1) = Chr(34) Then
                     nSelLen = nSelLen + 1
                  Else
                     nSelLen = nSelLen - 1
                     Exit For
                  End If
               Else
                  nSelLen = nSelLen + 1
               End If
            Next
            On Error Resume Next

            sTmp = Mid$(sLine, nStart, nSelLen)
            If (sTmp <> "") Then
               ' *** Check for string containing exceptions
               If gbExceptString Then
                  If bExceptString = False Then
                     For nK = 1 To colExceptString.Count
                        nPos1 = InStr(UCase$(sTmp), colExceptString(nK))
                        If (nPos1 > 0) Then
                           bExceptString = True
                           Exit For
                        End If
                     Next
                  End If
               End If

               ' *** Check if all char are in uppercase
               If (bExceptString = False) And (gbExceptAllUpperCase) Then
                  bExceptString = (UCase$(sTmp) = sTmp)
               End If

               ' *** Check the minimum size
               If gbGetMinimumSize And (Len(sTmp) < gnGetMinimumSize) Then bExceptString = True

               If (bExceptString = False) Then
                  ' *** Internationalize this string
                  ' *** Verify if already internaionalized
                  sTmp2 = sTmp
                  nPosChar = InStr(sTmp, "µ")
                  If nPosChar > 0 Then sTmp = left$(sTmp, nPosChar - 1)

                  sSQL = "Select * From Translation "
                  sSQL = sSQL & "Where Original = '" & Replace(sTmp, "'", "''") & "' "
                  Set record = DBTranslate.OpenRecordset(sSQL, DAO.dbOpenDynaset)
                  If record.RecordCount > 0 Then
                     If nPosChar = 0 Then
                        ' *** Not yet translated
                        If bHeader Then
                           sNewLine = Replace(sNewLine, """" & sTmp & """", """" & record("Original") & "µ" & record("ID") & """")
                        Else
                           sNewLine = Replace(sNewLine, """" & sTmp & """", "Translation(""" & record("Original") & "µ" & record("ID") & """)")
                        End If
                     Else
                        ' *** Already translated
                        sNewLine = Replace(sNewLine, """" & sTmp & """", """" & record("Original") & "µ" & record("ID") & """")
                     End If
                  End If
                  record.Close
                  Set record = Nothing
               End If
            End If
            err = 0

            ' *** Continue the line for possible second string
            If nJ < Len(sLine) Then
               bExceptString = False
               nPos = InStr(nJ + 1, sLine, Chr$(34))
               If nPos > 0 Then GoTo Get_String
            End If

            On Error GoTo ERROR_InternationalizeProject

         End If
Next_Line:
         sNewFile = sNewFile & sNewLine & vbCrLf
      Next
      ' *** Save the new file
      sNewFile = Trim$(sNewFile)

      If FileExist(cmp.FileNames(1)) Then
         nFile = FreeFile
         Open cmp.FileNames(1) For Output Access Write As #nFile
         Print #nFile, sNewFile
         Close #nFile

         ' *** Reload the components
         On Error Resume Next
         cmp.Reload
         On Error GoTo ERROR_InternationalizeProject

         frmExtractString.lstResult.AddItem sName & " done and saved"
      Else
         frmExtractString.lstResult.AddItem sName & " not saved"
      End If
      frmExtractString.lstResult.AddItem ""

      nComponent = nComponent + 1
   Next
   frmProgress.Progress = VBInstance.ActiveVBProject.VBComponents.Count
   frmExtractString.lstResult.AddItem VBInstance.ActiveVBProject.Name & " done"
   frmExtractString.lstResult.AddItem ""

   frmExtractString.lstResult.AddItem "Add in the Form_Load of each form the following code"
   frmExtractString.lstResult.AddItem "   Call InternationalizeForm(Me)"
   frmExtractString.lstResult.AddItem ""
   frmExtractString.lstResult.AddItem "Add also the Internationalization.bas"
   frmExtractString.lstResult.AddItem " wich implements all the needed functions"
   frmExtractString.lstResult.AddItem ""

EXIT_InternationalizeProject:
   On Error Resume Next
   DBTranslate.Close
   Set DBTranslate = Nothing

   Unload frmProgress

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_InternationalizeProject:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in InternationalizeProject", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_InternationalizeProject
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_InternationalizeProject

End Sub
