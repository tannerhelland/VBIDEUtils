VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVBEEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 12/10/1998
' * Time             : 20:32
' * Module Name      : clsVBEEvents
' * Module Filename  : clsVBEEvents.cls
' **********************************************************************
' * Comments         :
' * Adds items to the VBE environment to recreate the indenting
' * for the current procedure, module or project.
' *
' *
' **********************************************************************

Option Explicit

Public WithEvents cMenu As CommandBarEvents
Attribute cMenu.VB_VarHelpID = -1

Private Sub cMenu_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 12/10/1998
   ' * Time             : 20:33
   ' * Module Name      : clsVBEEvents
   ' * Module Filename  : clsVBEEvents.cls
   ' * Procedure Name   : cMenu_Click
   ' * Parameters       :
   ' *                    ByVal CommandBarControl As Object
   ' *                    handled As Boolean
   ' *                    CancelDefault As Boolean
   ' **********************************************************************
   ' * Comments         : Handles clicking on the menu item
   ' *
   ' *
   ' **********************************************************************

   Dim cHourglass       As class_Hourglass
   Static nCountReg  As Integer

   On Error Resume Next

   If CommandBarControl.Parameter = "" Then Exit Sub

   If (gbRegistered = False) And (CommandBarControl.Parameter <> "About") And (CommandBarControl.Parameter <> "Register") Then
      If ((nCountReg = 0) Or (nCountReg >= 20)) Or (gbTrial = False) Then
         nCountReg = 1
         frmAbout.bAbout = True
         frmAbout.Show vbModal
      Else
         nCountReg = nCountReg + 1
      End If
   End If

   VBInstance.ActiveCodePane.Window.SetFocus

   ' *** Do the appropriate action, depending on the menu item's parameter value
   Select Case CommandBarControl.Parameter
      Case "Proc":
         Call IndentProcedure

      Case "Mod":
         Call IndentModule

      Case "Proj":
         Call IndentProject

      Case "ModuleHeader":
         Call InsertModuleHeader

      Case "ProcHeader":
         Call InsertProcedureHeader

      Case "RemoveProcHeader":
         Call RemoveProcedureHeader

      Case "ModuleProcHeader":
         Call InsertModuleProcedureHeader

      Case "RemoveModuleProcHeader":
         Call RemoveModuleProcedureHeader

      Case "ProjProcHeader":
         Call InsertProjectHeader

      Case "CommentOut":
         Call BlockOutCode

      Case "UnComment":
         Call UnBlockOutCode

      Case "VBDoc":
         frmVBDocumentor.Show

      Case "HTMLProcedure":
         Call ExportCurrentProcedureAsHTML

      Case "HTMLModule":
         Call ExportCurrentModuleAsHTML

      Case "AddLineNumber":
         Call AddLineNumbering(False)

      Case "RemoveLineNumber":
         Call AddLineNumbering(True)

      Case "Swap=":
         Call SwapEgual

      Case "Alphabetize":
         Call AlphabetizeProcedure

      Case "Inline":
         Call InsertText(gsInline & gsDevelopper & ", " & Now & " : ")

      Case "Pending":
         Call GetPending

      Case "TabOrder":
         Call AlwaysOnTop(frmTabOrder, True)
         frmTabOrder.Show

      Case "ControlProperties":
         Call AlwaysOnTop(frmTooltips, True)
         frmTooltips.Show

      Case "Accelerators":
         Call AlwaysOnTop(frmAccelerator, True)
         frmAccelerator.Show

      Case "GenerateDLLBaseAdress":
         Call GenerateDLLBaseAdress

      Case "GenerateGUID":
         Call GenerateGUID

      Case "ClearDebug":
         ClearDebug

      Case "APIError":
         Call APIError

      Case "ShowKeyCode":
         Call AlwaysOnTop(frmKeyCode, True)
         frmKeyCode.Show

      Case "CharPicker":
         Call AlwaysOnTop(frmCharPicker, True)
         frmCharPicker.Show

      Case "ColorPicker":
         Call AlwaysOnTop(frmColorPicker, True)
         frmColorPicker.Show

      Case "ClassSpy":
         Call AlwaysOnTop(frmSpy, True)
         frmSpy.Show

      Case "MsgBox":
         Set cHourglass = New class_Hourglass
         frmMessageBox.Show

      Case "ListOLEServers":
         Call AlwaysOnTop(frmGetServers, True)
         frmGetServers.Show

      Case "IconExplorer":
         Set cHourglass = New class_Hourglass
         On Error Resume Next
         Shell GetSetting(gsREG_APP, "Execute", "VBIDEUtils1", App.Path & "\VBIDEUtils1.exe")

      Case "ActiveXExplorer":
         Set cHourglass = New class_Hourglass
         On Error Resume Next
         Shell GetSetting(gsREG_APP, "Execute", "VBIDEUtils2", App.Path & "\VBIDEUtils2.exe")

      Case "Dependencies":
         frmDependency.Show

      Case "ProjectExplorer":
         Call ShowProjectExplorer

      Case "ProjectAnalyzer"
         Call AlwaysOnTop(frmCodeAnalyse, True)
         frmCodeAnalyse.Show

      Case "DBCreator":
         Set cHourglass = New class_Hourglass
         frmDBCreator.Show

      Case "AddNewProcedure":
         frmAddNewProcedure.Show vbModal

      Case "StringExtractor":
         Call InitExtractString
         frmExtractString.Show

      Case "ToolbarCode":
         Call GenerateToolbar

      Case "SearchVBWeb":
         Call AlwaysOnTop(frmFindWeb, True)
         frmFindWeb.Show

      Case "CodeDatabase":
         Set cHourglass = New class_Hourglass
         On Error Resume Next
         Shell GetSetting(gsREG_APP, "Execute", "VBCodeDatabase", App.Path & "\VBCodeDatabase.exe")

      Case "Options":
         frmOptions.Show

      Case "About":
         frmAbout.bAbout = True
         frmAbout.Show

      Case "Register":
         frmRegister.Show vbModal

      Case "ErrorProc":
         Call InsertProcedureError

      Case "ErrorMod":
         Call InsertModuleError

      Case "ErrorProj":
         Call InsertProjectError

      Case "EnhancedErrorProc":
         Call InsertEnhancedProcedureError

      Case "EnhancedErrorMod":
         Call InsertEnhancedModuleError

      Case "EnhancedErrorProj":
         Call InsertEnhancedProjectError

      Case "ClearUnused":
         Call CloseUnusedWindows

      Case "FindReplace":
         Call ShowFindAndReplace

      Case "MouseZoom":
         Call ShowMouseZoom

      Case "ADOConnectionCreator":
         Call CreateNewADOConnection

      Case "AlignControlsLeft":
         Call AlignControlsLeft

      Case "AlignControlsRight":
         Call AlignControlsRight

      Case "AlignControlsTop":
         Call AlignControlsTop

      Case "AlignControlsBottom":
         Call AlignControlsBottom

   End Select

   Handled = True

End Sub
