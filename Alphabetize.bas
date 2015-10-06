Attribute VB_Name = "AlphabetizeCode_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 13/10/99
' * Time             : 11:57
' * Module Name      : AlphabetizeCode_Module
' * Module Filename  : Alphabetize.bas
' **********************************************************************
' * Comments         : Alphabetize the procedures in a module
' *
' *
' **********************************************************************

Option Explicit

Const ProcUnderscore = "2"
Const ProcNoUnderscore = "1"

Public Sub AlphabetizeProcedure()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 13/10/99
   ' * Time             : 11:57
   ' * Module Name      : AlphabetizeCode_Module
   ' * Module Filename  : Alphabetize.bas
   ' * Procedure Name   : AlphabetizeProcedure
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim modCode          As CodeModule
   Dim cpCodePane       As CodePane
   Dim sProcName        As String
   Dim nProcKind        As Long
   Dim nSelectLine      As Long
   Dim nStartLine       As Long
   Dim nStartColumn     As Long
   Dim nEndline         As Long
   Dim nEndColumn       As Long
   Dim nCountOfLines    As Long
   Dim sProcText        As String
   Dim CollectedProcs   As Object
   Dim CollectedKeys()  As String
   Dim sKey             As String
   Dim nI               As Integer
   Dim nIndex           As Integer

   If MsgBox("Would you like to sort all the procedures name alphabetically", vbQuestion + vbYesNo + vbDefaultButton1, "Alphabetization the procedures") = vbNo Then
      Exit Sub
   End If

   On Error Resume Next

   Set CollectedProcs = New Collection
   ReDim CollectedKeys(0) As String

   ' *** If we couldn't get it, quit
   If VBInstance.ActiveVBProject Is Nothing Then
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

   Set modCode = cpCodePane.CodeModule

   Do While modCode.CountOfLines > modCode.CountOfDeclarationLines
      nStartLine = modCode.CountOfDeclarationLines + 1
      cpCodePane.SetSelection modCode.CountOfDeclarationLines + 1, 1, modCode.CountOfDeclarationLines + 1, 1
      cpCodePane.GetSelection nSelectLine, nStartColumn, nEndline, nEndColumn
      sProcName = modCode.ProcOfLine(nSelectLine, nProcKind)
      nCountOfLines = modCode.ProcCountLines(sProcName, nProcKind)
      sProcText = modCode.Lines(nStartLine, nCountOfLines)
      sKey = IIf(InStr(sProcName, "_"), ProcUnderscore, ProcNoUnderscore)
      sKey = sKey & sProcName & StringProcKind(nProcKind)
      CollectedProcs.Add sProcText, sKey
      ReDim Preserve CollectedKeys(0 To UBound(CollectedKeys) + 1) As String
      CollectedKeys(UBound(CollectedKeys)) = sKey
      modCode.DeleteLines nStartLine, nCountOfLines
   Loop

   Do
      nIndex = 0
      sKey = " "

      For nI = 1 To UBound(CollectedKeys)
         If LCase$(CollectedKeys(nI)) > LCase$(sKey) Then
            sKey = CollectedKeys(nI)
            nIndex = nI
         End If
      Next

      If nIndex > 0 Then
         sProcText = CollectedProcs(sKey)
         CollectedKeys(nIndex) = " "
         modCode.AddFromString sProcText
      End If

   Loop Until nIndex = 0

   cpCodePane.Window.SetFocus
   cpCodePane.SetSelection 1, 1, 1, 1

End Sub

Private Function StringProcKind(ByVal kind As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 13/10/99
   ' * Time             : 11:57
   ' * Module Name      : AlphabetizeCode_Module
   ' * Module Filename  : Alphabetize.bas
   ' * Procedure Name   : StringProcKind
   ' * Parameters       :
   ' *                    ByVal Kind As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Select Case kind
      Case vbext_pk_Get
         StringProcKind = " Get"
      Case vbext_pk_Let
         StringProcKind = " Let"
      Case vbext_pk_Set
         StringProcKind = " Set"
      Case vbext_pk_Proc
         StringProcKind = " "
   End Select
End Function
