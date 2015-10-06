Attribute VB_Name = "CodeToolbar_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 9/02/99
' * Time             : 14:28
' * Module Name      : CodeToolbar_Module
' * Module Filename  : CodeToolbar.bas
' **********************************************************************
' * Comments         : Generate code to create the toolbar
' *
' *
' **********************************************************************

Option Explicit

Private mcmpCurrentForm As VBComponent      'current form
Private mcolCtls        As VBControls       'form's controls

Public Sub GenerateToolbar()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 9/02/99
   ' * Time             : 14:28
   ' * Module Name      : CodeToolbar_Module
   ' * Module Filename  : CodeToolbar.bas
   ' * Procedure Name   : GenerateToolbar
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Generate the toolbar
   ' *
   ' *
   ' **********************************************************************

   Dim Ctl              As VBControl

   Dim sTmp             As String

   Dim sConfig          As String
   Dim sImageList       As String
   Dim sButton          As String

   Dim tool             As MSComctlLib.ToolBar
   Dim Button           As MSComctlLib.Button

   Dim nI               As Long
   Dim nCount           As Integer

   Dim sTotal           As String

   Dim sIndex           As String

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   On Error GoTo ERROR_GenerateToolbar

   If VBInstance.ActiveVBProject Is Nothing Then Exit Sub

   ' *** Load the component
   Set mcmpCurrentForm = VBInstance.SelectedVBComponent

   ' *** Check to see if we have a valid component
   If mcmpCurrentForm Is Nothing Then Exit Sub

   ' *** Make sure the active component is a form, user control or property page
   If (mcmpCurrentForm.Type <> vbext_ct_VBForm) And _
      (mcmpCurrentForm.Type <> vbext_ct_UserControl) And _
      (mcmpCurrentForm.Type <> vbext_ct_DocObject) And _
      (mcmpCurrentForm.Type <> vbext_ct_PropPage) Then
      Exit Sub
   End If

   Set mcolCtls = mcmpCurrentForm.Designer.VBControls

   sTotal = ""

   Load frmProgress
   frmProgress.MessageText = "Generating toolbar creation code"
   frmProgress.Maximum = mcmpCurrentForm.Designer.VBControls.Count
   nCount = 1
   frmProgress.Show
   frmProgress.ZOrder

   For Each Ctl In mcmpCurrentForm.Designer.VBControls
      frmProgress.Progress = nCount
      nCount = nCount + 1
      If Ctl.ClassName = "Toolbar" Then
         ' *** We found a toolbar
         ' *** Create the code to configure it

         Set tool = Ctl.ControlObject

         On Error Resume Next

         ' *** Set the imagelist
         If tool.index = -1 Then
            sIndex = ""
         Else
            sIndex = "(" & tool.index & ")"
         End If
         sImageList = vbTab & "Set " & tool.Name & sIndex & ".ImageList = " & tool.ImageList.Name & Chr$(vbKeyReturn)

         ' *** Create the configuration of the toolbar
         sConfig = vbTab & "With " & tool.Name & sIndex & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".Align = " & CStr(tool.Align) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".AllowCustomize = " & CStr(tool.AllowCustomize) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".Appearance = " & CStr(tool.Appearance) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".BorderStyle = " & CStr(tool.BorderStyle) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".ButtonHeight = " & tool.ImageList.Name & ".ImageHeight" & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".ButtonWidth = " & tool.ImageList.Name & ".ImageWidth" & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".DragMode = " & CStr(tool.DragMode) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".Enabled = " & CStr(tool.Enabled) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".Height = " & ChangetoDot(tool.Height) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".HelpContextID = " & CStr(tool.HelpContextID) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".MousePointer = " & CStr(tool.MousePointer) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".OLEDropMode = " & CStr(tool.OLEDropMode) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".TabIndex = " & CStr(tool.TabIndex) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".Tag = """ & CStr(tool.Tag) & """" & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".ToolTipText= """ & CStr(tool.ToolTipText) & """" & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".Top = " & ChangetoDot(tool.top) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".Visible = " & CStr(tool.Visible) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".WhatsThisHelpID = " & CStr(tool.WhatsThisHelpID) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & vbTab & ".Wrappable = " & CStr(tool.Wrappable) & Chr$(vbKeyReturn)
         sConfig = sConfig & vbTab & "End With" & Chr$(vbKeyReturn)

         ' *** Set all the buttons
         If tool.Buttons.Count > 0 Then
            sButton = vbTab & "Dim aButton" & vbTab & vbTab & "As MSComctlLib.button" & Chr$(vbKeyReturn) & Chr$(vbKeyReturn)
         Else
            sButton = ""
         End If
         For nI = 1 To tool.Buttons.Count
            Set Button = tool.Buttons(nI)

            sButton = sButton & vbTab & "Set aButton = " & tool.Name & sIndex & ".Buttons.Add(" & Button.index & ", """ & Button.Key & """, """ & Button.Caption & """, " & Button.Style & ", " & IIf(IsNumeric(Button.Image), Button.Image, """" & Button.Image & """") & ")" & Chr$(vbKeyReturn)

            sButton = sButton & vbTab & "With aButton" & Chr$(vbKeyReturn)
            sButton = sButton & vbTab & vbTab & ".Description = """ & Button.Description & """" & Chr$(vbKeyReturn)
            sButton = sButton & vbTab & vbTab & ".Enabled = " & CStr(Button.Enabled) & Chr$(vbKeyReturn)
            sButton = sButton & vbTab & vbTab & ".MixedState = " & CStr(Button.MixedState) & Chr$(vbKeyReturn)
            sButton = sButton & vbTab & vbTab & ".Tag = """ & CStr(Button.Tag) & """" & Chr$(vbKeyReturn)
            sButton = sButton & vbTab & vbTab & ".ToolTipText = """ & CStr(Button.ToolTipText) & """" & Chr$(vbKeyReturn)
            sButton = sButton & vbTab & vbTab & ".Value = " & CStr(Button.Value) & Chr$(vbKeyReturn)
            sButton = sButton & vbTab & vbTab & ".Visible = " & CStr(Button.Visible) & Chr$(vbKeyReturn)
            sButton = sButton & vbTab & "End With" & Chr$(vbKeyReturn)

         Next

         ' *** Toolbar done, add it
         If tool.index = -1 Then
            sTmp = "Private Sub Create_" & CStr(tool.Name) & "()" & Chr$(vbKeyReturn)
         Else
            sTmp = "Private Sub Create_" & CStr(tool.Name) & CStr(tool.index) & "()" & Chr$(vbKeyReturn)
         End If
         sTmp = sTmp & vbTab & "' #VBIDEUtils#************************************************************" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' * Programmer Name  : VBIDEUtils Toolbar Generator" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' * Web Site         : http://www.ppreview.net" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' * E-Mail           : removed" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' * Date             : " & Format(Now, "mm/dd/yyyy") & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' * Time             : " & Format(Now, "Short Time") & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' * Procedure Name   : CreateToolbar" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' * Parameters       :" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' **********************************************************************" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' * Comments         : Create the toolbar at runtime" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' *" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' *" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & "' **********************************************************************" & Chr$(vbKeyReturn)
         sTmp = sTmp & vbTab & Chr$(vbKeyReturn)

         ' *** Clear existing buttons
         sTmp = sTmp & vbTab & tool.Name & sIndex & ".Buttons.Clear" & Chr$(vbKeyReturn)

         ' *** Add imagelist
         sTmp = sTmp & sImageList & Chr$(vbKeyReturn)

         ' *** Create toolbar
         sTmp = sTmp & sConfig & Chr$(vbKeyReturn)

         ' *** Add buttons
         sTmp = sTmp & sButton & Chr$(vbKeyReturn)

         ' *** End procedure
         sTmp = sTmp & "End Sub" & Chr$(vbKeyReturn)

         ' *** Keep it all
         sTotal = sTotal & sTmp & Chr$(vbKeyReturn)
      End If
   Next

   If (sTotal <> "") Then
      ' *** Copy to the clipboard
      Clipboard.Clear
      Clipboard.SetText sTotal, vbCFText

      ' *** Copy done
      sTmp = "The CreateToolbar function " & Chr$(13) & "has been copied to the clipboard." & Chr$(13)
      sTmp = sTmp & "You can paste it in your code."

      Call MsgBoxTop(frmProgress.hWnd, sTmp, vbOKOnly + vbInformation, "Toolbar Code generator")
   End If

EXIT_GenerateToolbar:
   Unload frmProgress
   Set frmProgress = Nothing

   Exit Sub

ERROR_GenerateToolbar:
   'MsgBox Error

   Resume EXIT_GenerateToolbar

End Sub

Private Function ChangetoDot(nLng As Double) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 9/02/99
   ' * Time             : 14:28
   ' * Module Name      : CodeToolbar_Module
   ' * Module Filename  : CodeToolbar.bas
   ' * Procedure Name   : ChangetoDot
   ' * Parameters       :
   ' *                    nLng As Double
   ' **********************************************************************
   ' * Comments         : Cahnge all , to .
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim nI               As Long

   sTmp = Format(nLng, "0")

   For nI = 1 To Len(sTmp)
      If Mid$(sTmp, nI, 1) = "," Then Mid$(sTmp, nI, 1) = "."
   Next

   ChangetoDot = sTmp

End Function

