Attribute VB_Name = "IndentRoutine_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/05/97
' * Time             : 17:29
' * Module Name      : IndentRoutine_Module
' * Module Filename  : IndentRoutine.bas
' **********************************************************************
' * Comments         :
' *   Contains the main procedure to rebuild the code's indenting
' *
' **********************************************************************
Option Explicit
Option Compare Text
Option Base 1

Global Const miTAB      As Integer = 9

Public Sub RebuildCodePanel(modCode As CodeModule, sName As String, nStartLine As Long, nEndline As Long, nProgDone As Long, nProgTotal As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/05/97
   ' * Time             : 17:30
   ' * Module Name      : IndentRoutine_Module
   ' * Module Filename  : IndentRoutine.bas
   ' * Procedure Name   : RebuildCodePanel
   ' * Parameters       :
   ' *                    modCode As CodeModule
   ' *                    sName As String
   ' *                    nStartLine As Long
   ' *                    nEndline As Long
   ' *                    nProgDone As Long
   ' *                    nProgTotal As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' *
   ' **********************************************************************

   Dim clsIndenter      As class_Indenter

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Dim sIndented()      As String
   Dim sText            As String
   Dim sTmp             As String

   Dim prjProject       As VBProject
   Dim memberObj        As Member
   Dim clsMember        As New class_Member
   Dim colMembers       As New Collection
   Dim nI               As Long
   Dim nTmp             As Long

   Dim nSecondStart     As Long

   ' *** Get all the text to indent
   sText = modCode.Lines(nStartLine, nEndline - nStartLine + 1)

   ' *** Store all the procedures attributes
   Set prjProject = VBInstance.ActiveVBProject

   If prjProject Is Nothing Then Exit Sub

   Set colMembers = New Collection
   nI = 1
   nTmp = 0
   frmProgress.MessageText = "Analysing code" & vbCrLf & sName
   frmProgress.Maximum = modCode.members.Count
   For Each memberObj In modCode.members
      frmProgress.Progress = nTmp
      nTmp = nTmp + 1
      Set clsMember = New class_Member
      On Error Resume Next
      With memberObj
         If (.Type = vbext_mt_Method) Or (.Type = vbext_mt_Property) Or (.Type = vbext_mt_Event) Then
            Debug.Print .Name
            clsMember.Name = .Name
            clsMember.Bindable = .Bindable
            clsMember.Browsable = .Browsable
            clsMember.Category = .Category
            clsMember.DefaultBind = .DefaultBind
            clsMember.Description = .Description
            clsMember.DisplayBind = .DisplayBind
            clsMember.HelpContextID = .HelpContextID
            clsMember.Hidden = .Hidden
            clsMember.PropertyPage = .PropertyPage
            clsMember.RequestEdit = .RequestEdit
            clsMember.StandardMethod = .StandardMethod
            clsMember.UIDefault = .UIDefault

            colMembers.Add clsMember, "Member" & CStr(nI)
            nI = nI + 1
         End If
      End With
   Next

   frmProgress.MessageText = "Indenting code" & vbCrLf & sName

   Set clsIndenter = New class_Indenter

   ' *** Set for progress
   clsIndenter.nProgDone = nProgDone
   nSecondStart = (nEndline - nStartLine)
   clsIndenter.nProgTotal = nSecondStart

   frmProgress.Maximum = clsIndenter.nProgTotal

   ' *** Start to indent
   sTmp = clsIndenter.IndentVBCode(sText, sIndented)

   Set clsIndenter = Nothing

   Do While right(sTmp, 2) = vbCrLf
      sTmp = left$(sTmp, Len(sTmp) - 2)
   Loop
   If (nStartLine + nProgTotal + 2 <= modCode.CountOfLines) Then
      modCode.DeleteLines nStartLine, nProgTotal + 1
   Else
      modCode.DeleteLines nStartLine, modCode.CountOfLines - nStartLine + 1
   End If
   modCode.InsertLines nStartLine, sTmp

   frmProgress.MessageText = "Verifying indentation" & vbCrLf & sName
   frmProgress.Maximum = colMembers.Count

   ' *** Set back all the attributes
   For nI = 1 To colMembers.Count
      frmProgress.Progress = nI
      Set clsMember = colMembers("Member" & CStr(nI))
      On Error Resume Next
      With modCode.members(clsMember.Name)
         Debug.Print clsMember.Name
         If clsMember.Bindable Then .Bindable = True
         If clsMember.Browsable Then .Browsable = True
         If Len(clsMember.Category) Then .Bindable = clsMember.Category
         If clsMember.DefaultBind Then .DefaultBind = True
         If Len(clsMember.Description) Then .Description = clsMember.Description
         If clsMember.DisplayBind Then .DisplayBind = True
         If clsMember.HelpContextID Then .HelpContextID = clsMember.HelpContextID
         If clsMember.Hidden Then .Hidden = True
         If Len(clsMember.PropertyPage) Then .PropertyPage = clsMember.PropertyPage
         If clsMember.RequestEdit Then .RequestEdit = True
         If clsMember.StandardMethod <= 0 Then .StandardMethod = clsMember.StandardMethod
         If clsMember.UIDefault Then .UIDefault = True
      End With

      Set clsMember = Nothing
   Next

   Set colMembers = Nothing

End Sub

Public Function IndentCode(sCode As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/05/97
   ' * Time             : 17:26
   ' * Module Name      : IndentRoutine_Module
   ' * Module Filename  : IndentRoutine.bas
   ' * Procedure Name   : IndentCode
   ' * Parameters       :
   ' *                    sCode As String
   ' **********************************************************************
   ' * Comments         :
   ' *   Indents the code
   ' *
   ' **********************************************************************

   Dim clsIndenter      As class_Indenter

   Dim sIndented()      As String
   Dim sTmp             As String

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Show the status bar user form.  The activate of the userform runs the indenting
   ' *** routine, so it can update the status bar form as it progresses.
   Load frmProgress
   frmProgress.MessageText = "Indenting code"
   DoEvents
   frmProgress.Show
   frmProgress.ZOrder

   Set clsIndenter = New class_Indenter

   ' *** Set for progress
   clsIndenter.nProgDone = 1
   clsIndenter.nProgTotal = CountTokens(sCode, Chr$(13))

   frmProgress.Maximum = clsIndenter.nProgTotal

   ' *** Start to indent
   sTmp = clsIndenter.IndentVBCode(sCode, sIndented)

   Do While right(sTmp, 2) = vbCrLf
      sTmp = left$(sTmp, Len(sTmp) - 2)
   Loop
   sTmp = sTmp & vbCrLf

   Set clsIndenter = Nothing

   Unload frmProgress
   Set frmProgress = Nothing

   IndentCode = sTmp

End Function
