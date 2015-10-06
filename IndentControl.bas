Attribute VB_Name = "IndentControl_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/05/97
' * Time             : 17:25
' * Module Name      : IndentControl_Module
' * Module Filename  : IndentControl.bas
' **********************************************************************
' * Comments         :
' *
' **********************************************************************

Option Explicit

Dim m_sIndentType       As String
Dim m_sIndentName       As String
Dim m_vbProjectObj      As VBProject
Dim m_vbCodemoduleObj   As CodeModule
Dim m_nStart            As Long
Dim m_nEnd              As Long

Public Sub IndentProcedure()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/05/97
   ' * Time             : 17:26
   ' * Module Name      : IndentControl_Module
   ' * Module Filename  : IndentControl.bas
   ' * Procedure Name   : IndentProcedure
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane
   Dim nStartLine       As Long
   Dim nStartCol        As Long
   Dim nEndline         As Long
   Dim nEndCol          As Long
   Dim sProc            As String
   Dim vaTypes          As Variant
   Dim nI               As Integer
   Dim nTmp             As Long

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set prjProject = VBInstance.ActiveVBProject

   If prjProject Is Nothing Then Exit Sub

   Set cpCodePane = VBInstance.ActiveCodePane

   If cpCodePane Is Nothing Then Exit Sub

   If Not IsThereCode(cpCodePane.CodeModule) Then Exit Sub

   ' *** Get where the current selection is in the module
   cpCodePane.GetSelection nStartLine, nStartCol, nEndline, nEndCol

   ' *** Create an array of procedure types to check for
   vaTypes = Array(vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)

   sProc = ""

   ' *** Loop through the procedure type
   For nI = 0 To 3
      ' *** Try to get the procedure name
      sProc = cpCodePane.CodeModule.ProcOfLine(nStartLine, CLng(vaTypes(nI)))
      If sProc <> "" Then
         ' *** If we got a procedure name, find its start and end lines and quit the loop
         nTmp = cpCodePane.CodeModule.ProcStartLine(sProc, CLng(vaTypes(nI)))
         nEndline = cpCodePane.CodeModule.ProcCountLines(sProc, CLng(vaTypes(nI))) + nTmp - 1
         If err Then err.Clear
         If nTmp > 0 Then Exit For
         sProc = ""
      End If
   Next

   If sProc = "" Then Exit Sub
   nStartLine = nTmp

   ' *** Store the currently active pane
   RecoveActiveState "Store"

   ' *** Set some module-level variables to specify what to rebuild
   m_sIndentType = "Procedure"
   m_sIndentName = sProc
   Set m_vbProjectObj = prjProject
   Set m_vbCodemoduleObj = cpCodePane.CodeModule
   m_nStart = nStartLine
   m_nEnd = nEndline

   Load frmProgress
   frmProgress.Show
   frmProgress.ZOrder
   frmProgress.MessageText = "Indenting " & Chr$(13) & m_sIndentName
   Call RebuildIndents
   Unload frmProgress
   Set frmProgress = Nothing

   ' *** Restore the currently active pane
   RecoveActiveState "Restore"

End Sub

Public Sub IndentModule()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/05/97
   ' * Time             : 17:26
   ' * Module Name      : IndentControl_Module
   ' * Module Filename  : IndentControl.bas
   ' * Procedure Name   : IndentModule
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Locates the active module, checks it for code and indents it
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set prjProject = VBInstance.ActiveVBProject

   If prjProject Is Nothing Then Exit Sub

   Set cpCodePane = VBInstance.ActiveCodePane

   If cpCodePane Is Nothing Then Exit Sub

   If Not IsThereCode(cpCodePane.CodeModule) Then
      Call MsgBox("The current module does not contain any code to indent!", vbExclamation + vbOKOnly + vbDefaultButton1, "No Code")
      Exit Sub
   End If

   ' *** Store the currently active pane
   RecoveActiveState "Store"

   ' *** Set some module-level variables to specify what to rebuild
   m_sIndentType = "Module"
   m_sIndentName = cpCodePane.CodeModule.Parent.Name
   Set m_vbProjectObj = prjProject
   Set m_vbCodemoduleObj = cpCodePane.CodeModule

   Load frmProgress
   frmProgress.Show
   frmProgress.ZOrder
   frmProgress.MessageText = "Indenting " & Chr$(13) & m_sIndentName
   Call RebuildIndents
   Unload frmProgress
   Set frmProgress = Nothing

   ' *** Restore the currently active pane
   RecoveActiveState "Restore"

End Sub

Public Sub IndentProject()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/05/97
   ' * Time             : 17:26
   ' * Module Name      : IndentControl_Module
   ' * Module Filename  : IndentControl.bas
   ' * Procedure Name   : IndentProject
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Locates the active project, checks it for code and indents it
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject
   Dim cmpObj           As VBComponent
   Dim bSomeCode        As Boolean

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set prjProject = VBInstance.ActiveVBProject

   If prjProject Is Nothing Then Exit Sub

   bSomeCode = False
   For Each cmpObj In prjProject.VBComponents
      If (cmpObj.Type <> vbext_ct_RelatedDocument) And (cmpObj.Type <> vbext_ct_ResFile) Then
         If IsThereCode(cmpObj.CodeModule) Then
            bSomeCode = True
            Exit For
         End If
      End If
   Next

   ' *** If we didn't find any code, display a message and quit
   If Not bSomeCode Then
      MsgBox "The current project, " & prjProject.Name & ", does not contain any code"
      Exit Sub
   End If

   ' *** Store the currently active pane
   RecoveActiveState "Store"

   ' *** Set some module-level variables to specify what to rebuild
   m_sIndentType = "Project"
   Set m_vbProjectObj = prjProject
   Set m_vbCodemoduleObj = Nothing

   Load frmProgress
   frmProgress.Show
   frmProgress.ZOrder
   frmProgress.MessageText = "Indenting " & Chr$(13) & prjProject.Name
   Call RebuildIndents
   Unload frmProgress
   Set frmProgress = Nothing

   ' *** Restore the currently active pane
   RecoveActiveState "Restore"

End Sub

Public Function IsThereCode(vbCodemoduleObj As CodeModule) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 13/10/99
   ' * Time             : 16:38
   ' * Module Name      : IndentControl_Module
   ' * Module Filename  : IndentControl.bas
   ' * Procedure Name   : IsThereCode
   ' * Parameters       :
   ' *                    vbCodemoduleObj As CodeModule
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Long

   For nI = 1 To vbCodemoduleObj.CountOfLines
      If Len(Trim$(vbCodemoduleObj.Lines(nI, 1))) > 0 Then
         IsThereCode = True
         Exit Function
      End If
   Next

   IsThereCode = False

End Function

Private Sub RecoveActiveState(sType As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/05/97
   ' * Time             : 17:28
   ' * Module Name      : IndentControl_Module
   ' * Module Filename  : IndentControl.bas
   ' * Procedure Name   : RecoveActiveState
   ' * Parameters       :
   ' *                    sType As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' **********************************************************************

   Static nStartLine    As Long
   Static nStartCol     As Long
   Static nEndline      As Long
   Static nEndCol       As Long
   Static nTop          As Long

   On Error GoTo ErrorHandler

   With VBInstance.ActiveCodePane
      If sType = "Store" Then
         nTop = VBInstance.ActiveCodePane.TopLine
         .GetSelection nStartLine, nStartCol, nEndline, nEndCol
      Else
         .TopLine = nTop
         .SetSelection nStartLine, nStartCol, nEndline, nEndCol
      End If
   End With

   Exit Sub

ErrorHandler:

   Exit Sub

End Sub

Private Sub RebuildIndents()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/05/97
   ' * Time             : 17:29
   ' * Module Name      : IndentControl_Module
   ' * Module Filename  : IndentControl.bas
   ' * Procedure Name   : RebuildIndents
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cmp              As VBComponent
   Dim lLineCount       As Long
   Dim lLinesDone       As Long

   ' *** Work out what needs to be done
   Select Case m_sIndentType
      Case "Procedure"
         ' *** Just rebuilding a procedure, so pass the procedure name and line boundaries
         frmProgress.MessageText = "Indenting " & Chr$(13) & m_sIndentName
         DoEvents
         RebuildCodePanel m_vbCodemoduleObj, m_sIndentName, m_nStart, m_nEnd, 0, m_nEnd - m_nStart

      Case "Module"
         ' *** Rebuilding a module, so pass the module name and number of lines therein
         frmProgress.MessageText = "Indenting " & Chr$(13) & m_sIndentName
         DoEvents
         RebuildCodePanel m_vbCodemoduleObj, m_sIndentName, 1, m_vbCodemoduleObj.CountOfLines, 0, m_vbCodemoduleObj.CountOfLines - 1

      Case "Project"
         ' *** Rebuilding a project, so we need to work out how many lines there are before doing the indenting

         ' *** Loop through the components, totalling their lines
         On Error GoTo Exit_1
         For Each cmp In m_vbProjectObj.VBComponents
            If (cmp.Type <> vbext_ct_RelatedDocument) And (cmp.Type <> vbext_ct_ResFile) Then
               If IsThereCode(cmp.CodeModule) Then ' *** removed, 05/11/1999 13:21:59 :
                  lLineCount = lLineCount + cmp.CodeModule.CountOfLines
               End If
            End If
         Next
Exit_1:

         ' *** Now loop through the components to rebuild their indenting
         On Error GoTo Exit_2
         For Each cmp In m_vbProjectObj.VBComponents
            If (cmp.Type <> vbext_ct_RelatedDocument) And (cmp.Type <> vbext_ct_ResFile) Then
               If IsThereCode(cmp.CodeModule) Then

                  frmProgress.MessageText = "Indenting " & Chr$(13) & cmp.CodeModule.Parent.Name
                  DoEvents
                  ' *** Pass the module name, number of lines, how many have been done already and how many are there in total
                  RebuildCodePanel cmp.CodeModule, cmp.CodeModule.Parent.Name, 1, cmp.CodeModule.CountOfLines, lLinesDone, lLineCount - 1

                  ' *** Increment the number of lines done for next time round
                  lLinesDone = lLinesDone + cmp.CodeModule.CountOfLines
               End If
            End If
         Next
Exit_2:
   End Select

End Sub
