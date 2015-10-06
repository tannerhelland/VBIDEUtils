Attribute VB_Name = "HTML_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 15/10/99
' * Time             : 15:19
' * Module Name      : HTML_Module
' * Module Filename  : HTML.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Public Sub ExportCurrentProcedureAsHTML()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/10/98
   ' * Time             : 17:26
   ' * Module Name      : HTML_Module
   ' * Module Filename  : HTML.bas
   ' * Procedure Name   : ExportCurrentProcedureAsHTML
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Save a procedure as HTML
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ExportCurrentProcedureAsHTML

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane
   Dim nStartLine       As Long
   Dim nStartCol        As Long
   Dim nEndline         As Long
   Dim nEndCol          As Long
   Dim sProcName        As String
   Dim sCode            As String

   Dim CommonDialog1    As class_CommonDialog

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Get the active project
   Set prjProject = VBInstance.ActiveVBProject

   ' *** If we couldn't get it, quit
   If prjProject Is Nothing Then Exit Sub

   ' *** Try to find the active code pane
   Set cpCodePane = VBInstance.ActiveCodePane

   ' *** If we couldn't get it, quit
   If cpCodePane Is Nothing Then Exit Sub

   ' *** Check if the module contains any code
   If Not IsThereCode(cpCodePane.CodeModule) Then Exit Sub

   ' *** Get where the current selection is in the module
   cpCodePane.GetSelection nStartLine, nStartCol, nEndline, nEndCol

   ' *** Get the name of the procedure
   sProcName = cpCodePane.CodeModule.ProcOfLine(nStartLine, vbext_pk_Proc)

   If sProcName = "" Then Exit Sub

   ' *** Get where the current selection is in the module
   nStartLine = cpCodePane.CodeModule.ProcStartLine(sProcName, vbext_pk_Proc)
   nEndline = nStartLine + cpCodePane.CodeModule.ProcCountLines(sProcName, vbext_pk_Proc)
   sCode = cpCodePane.CodeModule.Lines(nStartLine, IIf(nEndline - nStartLine = 0, 1, nEndline - nStartLine))

   Call InitColorize
   frmIcons.Visible = False
   frmIcons.rtfColorize.TextRTF = ColorizeVBCode(sCode)

   Dim sFileName        As String
   Dim nFile            As Integer
   Dim sHTML            As String

   sHTML = rtf2html(frmIcons.rtfColorize.TextRTF, "+G+H+T=""" & sProcName & """")
   If Trim$(sHTML) = "" Then Exit Sub

   Set CommonDialog1 = New class_CommonDialog
   With CommonDialog1
      .DialogTitle = "Choose a filename to save"
      .CancelError = False
      .Flags = OFN_HIDEREADONLY + OFN_EXPLORER
      .InitDir = App.Path
      .FileName = "Exported.HTML"
      .Filter = "Internet documents (*.HTML)|*.HTML|Text files (*.TXT)|*.TXT|All Files (*.*)|*.*"
      .FilterIndex = 1
      .HookDialog = True
      .ShowSave

      sFileName = .FileName

   End With
   Set CommonDialog1 = Nothing

   If sFileName = "" Then Exit Sub

   ' *** Save the HTML
   nFile = FreeFile ' Get free file number.
   Open sFileName For Binary As #nFile  ' Open the file.
   Put #nFile, , sHTML
   Close nFile

   If gbRegistered = False Then
      frmAbout.bAbout = True
      frmAbout.Show vbModal
   End If

   Call MsgBox("The code of the procedure has been exported as HTML : " & sFileName, vbInformation + vbOKOnly + vbDefaultButton1, "Export as HTML")

EXIT_ExportCurrentProcedureAsHTML:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ExportCurrentProcedureAsHTML:
   Resume EXIT_ExportCurrentProcedureAsHTML

End Sub

Public Sub ExportCurrentModuleAsHTML()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/10/98
   ' * Time             : 17:26
   ' * Module Name      : HTML_Module
   ' * Module Filename  : HTML.bas
   ' * Procedure Name   : ExportCurrentModuleAsHTML
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Save a module as HTML
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ExportCurrentModuleAsHTML

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane
   Dim nStartLine       As Long
   Dim nStartCol        As Long
   Dim nEndline         As Long
   Dim nEndCol          As Long
   Dim sProcName        As String
   Dim sCode            As String

   Dim sFileName        As String
   Dim nFile            As Integer
   Dim sHTML            As String

   Dim sTmp             As String

   Dim CommonDialog1    As class_CommonDialog

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

   ' *** Try to find the active code pane
   Set cpCodePane = VBInstance.ActiveCodePane

   ' *** If we couldn't get it, quit
   If cpCodePane Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current module", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the module")
      Exit Sub
   End If

   ' *** Check if the module contains any code
   If Not IsThereCode(cpCodePane.CodeModule) Then
      Call MsgBox("The current module does not contain any code!", vbExclamation + vbOKOnly + vbDefaultButton1, "No Code")
      Exit Sub
   End If

   ' *** Get where the current selection is in the module
   cpCodePane.GetSelection nStartLine, nStartCol, nEndline, nEndCol

   ' *** Get the name of the module
   sProcName = cpCodePane.CodeModule.Name

   ' *** Get where the current selection is in the module
   nStartLine = 1
   nEndline = cpCodePane.CodeModule.CountOfLines + 1
   sCode = cpCodePane.CodeModule.Lines(nStartLine, IIf(nEndline - nStartLine = 0, 1, nEndline - nStartLine))

   Set CommonDialog1 = New class_CommonDialog
   With CommonDialog1
      .DialogTitle = "Choose a filename to save"
      .CancelError = False
      .Flags = OFN_HIDEREADONLY + OFN_EXPLORER
      .InitDir = App.Path
      .FileName = "Exported.HTML"
      .Filter = "Internet documents (*.HTML)|*.HTML|Text files (*.TXT)|*.TXT|All Files (*.*)|*.*"
      .FilterIndex = 1
      .HookDialog = True
      .ShowSave

      sFileName = .FileName

   End With
   Set CommonDialog1 = Nothing

   If sFileName = "" Then Exit Sub

   If FileExist(sFileName) Then
      If MsgBox("The file " & sFileName & " already exists." & Chr$(13) & "Do you want to overwrite it?", vbQuestion + vbYesNo + vbDefaultButton1, "Overwrite file") = vbNo Then GoTo EXIT_ExportCurrentModuleAsHTML
   End If

   DoEvents

   Set cHourglass = New class_Hourglass

   Load frmProgress
   frmProgress.MessageText = "Generating VB code HTML"
   frmProgress.ProgressBar.Visible = False
   frmProgress.Show
   frmProgress.ZOrder
   DoEvents

   Call InitColorize
   frmIcons.Visible = False
   frmIcons.rtfColorize.TextRTF = ColorizeVBCode(sCode)

   sHTML = rtf2html(frmIcons.rtfColorize.TextRTF, "+G+H+T=""" & sProcName & """")
   If Trim$(sHTML) = "" Then GoTo EXIT_ExportCurrentModuleAsHTML

   ' *** Replace the needed at the beginning
   sTmp = "<BODY>" & vbCrLf
   sTmp = sTmp & "<b><font face=Verdana size=4>" & sProcName & "</font></b>&nbsp;"
   sTmp = sTmp & "generated by <A HREF=""http://www.ppreview.net""> VBIDEUtils </A> Add-In for Visual Basic"
   If gbRegistered = False Then sTmp = sTmp & " <b><<Shareware Version>></b>"
   sTmp = sTmp & "</font></td><HR SIZE=1>" & vbCrLf
   sTmp = sTmp & vbCrLf
   sTmp = sTmp & "<TABLE WIDTH=""80%"" CELLPADDING=0 CELLSPACING=0 BORDER=0>" & vbCrLf
   sTmp = sTmp & "<TR>" & vbCrLf
   sTmp = sTmp & "<TD VALIGN=MIDDLE ALIGN=CENTER BGCOLOR=""#999999"">" & vbCrLf
   sTmp = sTmp & "<TABLE WIDTH=100% CELLPADDING=4 CELLSPACING=1 BORDER=0>" & vbCrLf
   sTmp = sTmp & "<TR>" & vbCrLf
   sTmp = sTmp & "<TD VALIGN=TOP BGCOLOR=#00007F><FONT SIZE=3 FACE=""ARIAL, HELVETICA"" COLOR=#FFFFFF><STRONG>VB" & vbCrLf
   sTmp = sTmp & vbCrLf
   sTmp = sTmp & "Code</STRONG></FONT></TD></TR><TR>" & vbCrLf
   sTmp = sTmp & "<TD VALIGN=TOP BGCOLOR=""#FFFFFF"">" & vbCrLf
   sTmp = sTmp & "<BR>" & vbCrLf
   sHTML = Replace(sHTML, "<BODY>", sTmp)

   ' *** Replace the needed at the end
   sTmp = "</TD></TR></TABLE>"
   sTmp = sTmp & "</TD></TR></TABLE></TD></TR></TABLE></CENTER><BR>" & vbCrLf
   sTmp = sTmp & "<b><font face=Verdana size=4>" & sProcName & "</font></b>&nbsp;"
   sTmp = sTmp & "generated by <A HREF=""http://www.ppreview.net""> VBIDEUtils </A> Add-In for Visual Basic"
   If gbRegistered = False Then sTmp = sTmp & " <b><<Shareware Version>></b>"
   sTmp = sTmp & "</font></td><HR SIZE=1>" & vbCrLf
   sTmp = sTmp & "</BODY>" & vbCrLf
   sHTML = Replace(sHTML, "</BODY>", sTmp)

   ' *** Kill file if existing
   On Error Resume Next
   Kill sFileName

   ' *** Save the HTML
   On Error GoTo ERROR_ExportCurrentModuleAsHTML
   nFile = FreeFile ' Get free file number.
   Open sFileName For Binary As #nFile  ' Open the file.
   Put #nFile, , sHTML
   Close nFile

   If gbRegistered = False Then
      frmAbout.bAbout = True
      frmAbout.Show vbModal
   End If

   Call MsgBoxTop(frmProgress.hWnd, "The code of module has been exported to HTML : " & sFileName, vbInformation + vbOKOnly + vbDefaultButton1, "Export VB Code to HTML")

EXIT_ExportCurrentModuleAsHTML:
   On Error Resume Next

   Unload frmProgress
   Set frmProgress = Nothing

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ExportCurrentModuleAsHTML:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in ExportCurrentModuleAsHTML", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_ExportCurrentModuleAsHTML
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_ExportCurrentModuleAsHTML

End Sub
