VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VB IDEUtils"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 18/09/1998
' * Time             : 12:24
' * Module Name      : clsConnect
' * Module Filename  : clsConnect.cls
' **********************************************************************
' * Comments         : Adds items to the VBE environment to recreate the indenting
' *                    for the current procedure, module or project.
' *
' *                    Sets up the menu items in the VBE and handles their events.
' *                    Command bar controls in the VBE do not have a working OnAction
' *                    property, so we have to trap and respond to the command bar
' *                    events instead.  This is done by the vbeMenus class array.
' *
' *
' **********************************************************************

Option Explicit
Option Compare Text
Option Base 1

Implements IDTExtensibility

' *** Define some constants for the text of the menu items
Const msmenuVBIDEUTILS     As String = "V&BIDEUtils"
Const msVBIDEUTILS         As String = "VBIDEUtils"
Const msMENU_PROC          As String = "Indent &Procedure"
Const msMENU_MOD           As String = "Indent &Module"
Const msMENU_PROJ          As String = "Indent Pr&oject"
Const msMENU_MODULE_HEADER As String = "&Module Header"
Const msMENU_PROC_HEADER   As String = "&Procedure Header"
Const msMENU_REM_PROC_HEADER   As String = "Remove Procedure Header"
Const msMENU_MOD_PROC_HEADER   As String = "Module Procedure &Header"
Const msMENU_REM_MOD_PROC_HEADER   As String = "Remove Module Procedure &Header"
Const msMENU_PROJ_PROC_HEADER  As String = "Pr&oject Procedure Header"
Const msMENU_COMMENT_OUT   As String = "&Comment Out"
Const msMENU_UNCOMMENT     As String = "&Uncomment"
Const msMENU_INLINE        As String = "&Inline"
Const msMENU_VBDOC         As String = "&VBDoc"
Const msMENU_ADDLINENUMBER As String = "&Add line numbering"
Const msMENU_REMOVELINENUMBER  As String = "&Remove line numbering"
Const msMENU_HTML_PROC     As String = "Export &Procedure code as HTML"
Const msMENU_HTML_MOD      As String = "Export &Module code as HTML"
Const msMENU_SWAP          As String = "S&wap between '='"
Const msMENU_ALPHABETIZE   As String = "Alpha&betize procedures"
Const msMENU_PENDING       As String = "Find &pending"
Const msMENU_FINDREPLACE   As String = "&Find and Replace..."
Const msMENU_TABORDER      As String = "&Tab Order of controls..."
Const msMENU_PROPERTIES    As String = "Change &Properties of controls..."
Const msMENU_MOUSEZOOM     As String = "Mouse &Zoom window..."
Const msMENU_ACCELERATOR   As String = "Change &Accelerators of controls..."
Const msMENU_CLEAR_DEBUG   As String = "Clear &Debug Window"
Const msMENU_API_ERROR     As String = "&API Error..."
Const msMENU_KEYCODE       As String = "&Keycode constants..."
Const msMENU_CHARPICKER    As String = "&Char Picker..."
Const msMENU_COLORPICKER   As String = "C&olor Picker..."
Const msMENU_CLASSNAME     As String = "&Spy window Classname..."
Const msMENU_MSGBOX        As String = "&MessageBox Creator..."
Const msMENU_DBCREATOR     As String = "DB &Creator..."
Const msMENU_ADOCONNCREATOR  As String = "&ADO Connection String..."
Const msMENU_STRINGEXTRACT As String = "Internationalize Project..."
Const msMENU_ICONEXPLORER  As String = "&Icons Explorer..."
Const msMENU_LISTOLESERVERS As String = "&List installed OLE Servers..."
Const msMENU_ACTIVEXEXPLORER  As String = "&ActiveX Documentor..."
Const msMENU_DLLBASEADDRESS  As String = "Generate &DLL Base adress..."
Const msMENU_GUIDID        As String = "&Generate GUID..."
Const msMENU_DEPENDENCIES  As String = "Detect &Dependencies..."
Const msMENU_PROJECTEXPLORER  As String = "&Project Explorer..."
Const msMENU_PROJECTEANALYZER As String = "Project A&nalyzer..."
Const msMENU_TOOLBAR       As String = "&Generate code for toolbar creation"
Const msMENU_FORM          As String = "VBIDE&Utils Options..."
Const msMENU_CODEDATABASE  As String = "&VB Code Database..."
Const msMENU_SEARCHVBWEB   As String = "Search on VB sites on the Web..."
Const msMENU_ABOUT         As String = "About VBIDEUtils..."

Const msMENU_ALIGNCONTROLS As String = "&Align Controls"
Const msMENU_ALIGNLEFT     As String = "Align All Controls &Left"
Const msMENU_ALIGNRIGHT    As String = "Align All Controls &Right"
Const msMENU_ALIGNTOP      As String = "Align All Controls &Top"
Const msMENU_ALIGNBOTTOM   As String = "Align All Controls &Bottom"

Const msMENU_REGISTER      As String = "!!!Register VBIDEUtils!!!"

Const msUNUSED_WINDOWS     As String = "Cl&ose all unused windows"
Const msADD_NEWPROCEDURE   As String = "&Add new procedure..."

Const msVBIDEERROR         As String = "VBIDEUtils E&rrors"
Const msERROR_PROC         As String = "Add Error handler in current P&rocedure"
Const msERROR_MOD          As String = "Add Error handler in the Mo&dule"
Const msERROR_PROJ         As String = "Add Error handler in all Pro&ject"
Const msENHANCEDERROR_PROC As String = "Add Enhanced Error handler in current P&rocedure"
Const msENHANCEDERROR_MOD  As String = "Add Enhanced Error handler in the Mo&dule"
Const msENHANCEDERROR_PROJ As String = "Add Enhanced Error handler in all Pro&ject"

Const mnSTART_EDITMENU           As Integer = 1
Const mnSTART_WINDOWSMENU        As Integer = 62
Const mnSTART_DEBUGWINDOWSMENU   As Integer = mnSTART_WINDOWSMENU * 2 - 1

Const mnEND_EVENTS               As Integer = mnSTART_WINDOWSMENU * 2

Const mnSTART_WINDOW_EVENTS      As Integer = 1
Const mnEND_WINDOW_EVENTS        As Integer = mnSTART_WINDOW_EVENTS * 4

Const mnSTART_CONTROL_EVENTS     As Integer = 1
Const mnEND_CONTROL_EVENTS       As Integer = 50

Const mnSTART_ERROR_EVENTS     As Integer = 1
Const mnEND_ERROR_EVENTS       As Integer = 8

Private mCommandBarButton(mnEND_EVENTS) As Office.CommandBarControl

' *** Define an array to store the menu item click event handlers
Private vbeMenus(1 To mnEND_EVENTS) As New clsVBEEvents
Private vbeControl(1 To mnEND_CONTROL_EVENTS) As New clsVBEEvents
Private vbeErrorControl(1 To mnEND_ERROR_EVENTS) As New clsVBEEvents
Private vbeWindowMenu(1 To mnEND_WINDOW_EVENTS) As New clsVBEEvents
Private vbeErrorMenu(1 To mnEND_ERROR_EVENTS) As New clsVBEEvents

Public WithEvents PrjHandler As VBProjectsEvents          'projects event handler
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents CmpHandler As VBComponentsEvents        'components event handler
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents CtlHandler As VBControlsEvents          'controls event handler
Attribute CtlHandler.VB_VarHelpID = -1

Private mnUses          As Integer

'Public vbideSubclasser        As class_VBIDE

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : IDTExtensibility_OnAddInsUpdate
   ' * Parameters       :
   ' *                    custom() As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : IDTExtensibility_OnStartupComplete
   ' * Parameters       :
   ' *                    custom() As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 18/09/1998
   ' * Time             : 12:31
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : IDTExtensibility_OnConnection
   ' * Parameters       :
   ' *                    ByVal VBInst As Object
   ' *                    ByVal ConnectMode As vbext_ConnectMode
   ' *                    ByVal AddInInst As VBIDE.AddIn
   ' *                    custom() As Variant
   ' **********************************************************************
   ' * Comments         : this method adds the Add-In to VB
   ' *
   ' *
   ' **********************************************************************

   Dim oCtl             As CommandBarControl

   On Error GoTo ERROR_handler

   Dim nFile            As Integer

   On Error GoTo ERROR_GETFile
   nFile = FreeFile
   Open GetWindowsDirectory() & "Win.res" For Input As #nFile
   Input #nFile, gsGUID
   Close #nFile
END_GET_IT:

   On Error GoTo ERROR_handler

   gbTrial = GetTrial()

   Load frmIcons
   DoEvents

   ' *** Get the settings
   gsTemplate = GetSetting(gsREG_APP, "General", "DefaultTemplate", "")
   gsDevelopper = GetSetting(gsREG_APP, "General", "Developper", gsDefaultDevelopper)
   gsWebSite = GetSetting(gsREG_APP, "General", "WebSite", gsDefaultWebSite)
   gsEmail = GetSetting(gsREG_APP, "General", "Email", gsDefaultEmail)
   gsTelephone = GetSetting(gsREG_APP, "General", "Telephone", "")
   gsInline = GetSetting(gsREG_APP, "General", "Inline", "' *** ")
   gsErrorHandler = GetSetting(gsREG_APP, "Error", "CustomizedErrorHandler", "")
   If Trim$(gsErrorHandler) = vbNullString Then
      gsErrorHandler = "Select Case MsgBox(""Error "" & Err.Number & "": "" & Err.Description & vbCrLf & ""in {ProcedureName}"" & vbCrLf & ""The error occured at line: "" & Erl, vbAbortRetryIgnore + vbCritical, ""Error"")" & vbCrLf
      gsErrorHandler = gsErrorHandler & "   Case vbAbort" & vbCrLf
      gsErrorHandler = gsErrorHandler & "      Screen.MousePointer = vbDefault" & vbCrLf
      gsErrorHandler = gsErrorHandler & "      Resume EXIT_{ProcedureName}" & vbCrLf
      gsErrorHandler = gsErrorHandler & "   Case vbRetry" & vbCrLf
      gsErrorHandler = gsErrorHandler & "      Resume" & vbCrLf
      gsErrorHandler = gsErrorHandler & "   Case vbIgnore" & vbCrLf
      gsErrorHandler = gsErrorHandler & "      Resume Next" & vbCrLf
      gsErrorHandler = gsErrorHandler & "End Select" & vbCrLf & vbCrLf
   End If
   gsUserComment = GetSetting(gsREG_APP, "General", "UserComment", "")
   gsCommentString = GetSetting(gsREG_APP, "General", "CommentString", "*")

   ' *** Save the Vb instance
   Set VBInstance = VBInst

   ' *** Setup the popup menu Window/Debug
   Call PopupCodeMenu

   ' *** Setup the menu
   Call SetupMainMenus

   ' *** Setup the Edit menu
   Call SetupDebugMenus

   ' *** Setup the Window menu
   Call SetupWindowMenus

   ' *** Setup the Tools menu
   Call SetupToolsMenus

   ' *** Setup the commandbar
   Call SetupNewCommandBar

   Set oCtl = VBInstance.CommandBars("Add-Ins").Controls.Add(Type:=msoControlButton, Parameter:="About", temporary:=True)
   oCtl.Caption = msMENU_ABOUT

   ' *** Copy the icon to the clipboard
   DoEvents
   Clipboard.SetData frmIcons.VBIDEUtils.Image, vbCFBitmap

   ' *** Set the icon for the button
   oCtl.PasteFace
   DoEvents

   Set vbeMenus(mnEND_EVENTS).cMenu = VBInstance.Events.CommandBarEvents(oCtl)

   VBInstance.CommandBars(msVBIDEUTILS).Visible = IIf(GetSetting(gsREG_APP, "MainCommandBar", "Visible", "1") = 1, True, False)
   VBInstance.CommandBars(msVBIDEUTILS).Position = GetSetting(gsREG_APP, "MainCommandBar", "Position", msoBarTop)
   VBInstance.CommandBars(msVBIDEUTILS).RowIndex = GetSetting(gsREG_APP, "MainCommandBar", "RowIndex", msoBarRowLast)
   VBInstance.CommandBars(msVBIDEUTILS).left = GetSetting(gsREG_APP, "MainCommandBar", "Left", 0)
   VBInstance.CommandBars(msVBIDEUTILS).top = GetSetting(gsREG_APP, "MainCommandBar", "Top", 0)

   '   VBInstance.CommandBars("VBIDEError").Visible = IIf(GetSetting(gsREG_APP, "VBIDEError", "Visible", "1") = 1, True, False)
   '   VBInstance.CommandBars("VBIDEError").Position = GetSetting(gsREG_APP, "VBIDEError", "Position", msoBarBottom)
   '   VBInstance.CommandBars("VBIDEError").Left = GetSetting(gsREG_APP, "VBIDEError", "Left", 0)
   '   VBInstance.CommandBars("VBIDEError").Top = GetSetting(gsREG_APP, "VBIDEError", "Top", 0)

   Set gAddInInst = AddInInst

   If (gsTemplate = "") Then Call Add_DefaultTemplate

   ' *** Sink the project, components and controls event handler
   Set Me.PrjHandler = VBInstance.Events.VBProjectsEvents
   Set Me.CmpHandler = VBInstance.Events.VBComponentsEvents(Nothing)
   Set Me.CtlHandler = VBInstance.Events.VBControlsEvents(Nothing, Nothing)

   Call SetKeyHook

   Dim aiTmp            As VBIDE.AddIn

   If Not gwinAdvFndRpl Is Nothing Then
      ' *** Already running so just show it
      ShowFindAndReplace
      Exit Sub
   End If

   If Not gwinMouseZoom Is Nothing Then
      ' *** Already running so just show it
      ShowMouseZoom
      Exit Sub
   End If

   '   If Not gwinProjectExplorer Is Nothing Then
   '      ' *** Already running so just show it
   '      gwinProjectExplorer.left = GetSetting(gsREG_APP, "ProjectExplorer", "Left", 0)
   '      gwinProjectExplorer.top = GetSetting(gsREG_APP, "ProjectExplorer", "Top", 0)
   '      gwinProjectExplorer.Visible = GetSetting(gsREG_APP, "ProjectExplorer", "Visible", True)
   '
   '      Exit Sub
   '   End If

   If ConnectMode = vbext_cm_External Then
      Set aiTmp = VBInstance.Addins(1)
      On Error GoTo EXIT_Connect
      If aiTmp Is Nothing Then
         'app is not in the VBADDIN.INI file so it is not in the collection
         'so lets attempt to use the 1st addin in the collection just
         'to get this app running and if there are none, an error
         'will occur and this app will not run
         Set gwinAdvFndRpl = VBInstance.Windows.CreateToolWindow(VBInstance.Addins(1), "VBIDEUtils.docAdvFndRpl", AdvFndRplCaption, guidAdvFndRpl$, gdocAdvFndRpl)
         Set gwinMouseZoom = VBInstance.Windows.CreateToolWindow(VBInstance.Addins(1), "VBIDEUtils.docMouseZoom", MouseZoomCaption, guidMouseZoom$, gdocMouseZoom)
         Set gwinProjectExplorer = VBInstance.Windows.CreateToolWindow(VBInstance.Addins(1), "VBIDEUtils.docProjectExplorer", ProjectExplorerCaption, guidProjectExplorer$, gdocProjectExplorer)
      Else
         If aiTmp.Connect = False Then
            Set gwinAdvFndRpl = VBInstance.Windows.CreateToolWindow(aiTmp, "VBIDEUtils.docAdvFndRpl", AdvFndRplCaption, guidAdvFndRpl$, gdocAdvFndRpl)
            Set gwinMouseZoom = VBInstance.Windows.CreateToolWindow(aiTmp, "VBIDEUtils.docMouseZoom", MouseZoomCaption, guidMouseZoom$, gdocMouseZoom)
            Set gwinProjectExplorer = VBInstance.Windows.CreateToolWindow(aiTmp, "VBIDEUtils.docProjectExplorer", ProjectExplorerCaption, guidProjectExplorer$, gdocProjectExplorer)
         End If
      End If
   Else
      ' *** Must've been called from addin mgr
      Set gwinAdvFndRpl = VBInstance.Windows.CreateToolWindow(AddInInst, "VBIDEUtils.docAdvFndRpl", AdvFndRplCaption, guidAdvFndRpl$, gdocAdvFndRpl)
      Set gwinMouseZoom = VBInstance.Windows.CreateToolWindow(AddInInst, "VBIDEUtils.docMouseZoom", MouseZoomCaption, guidMouseZoom$, gdocMouseZoom)
      Set gwinProjectExplorer = VBInstance.Windows.CreateToolWindow(AddInInst, "VBIDEUtils.docProjectExplorer", ProjectExplorerCaption, guidProjectExplorer$, gdocProjectExplorer)

   End If

   If ConnectMode = vbext_cm_External Then
      'started from the addin toolbar
      ShowFindAndReplace
   End If

   gwinProjectExplorer.Visible = IIf(GetSetting(gsREG_APP, "ProjectExplorer", "Visible", 1) = 1, True, False)

   '   If Not gbRegistered Then
   '      Set vbideSubclasser = New class_VBIDE
   '      vbideSubclasser.HookVBIDE
   '   End If

EXIT_Connect:
   Unload frmIcons

   Exit Sub

ERROR_GETFile:
   If err = 53 Then
      On Error Resume Next
      gsGUID = CreateGUID()
      Open GetWindowsDirectory() & "Win.res" For Binary Access Write As #nFile
      Put #nFile, , gsGUID
      Close #nFile
      Call SetOtherFileDate(GetWindowsDirectory() & "Win.res")

      Dim clsRegistry      As New class_Registry
      With clsRegistry
         .ClassKey = HKEY_CLASSES_ROOT
         .SectionKey = "CLSID\" & gsGUID
         .ValueKey = ""
         .ValueType = REG_SZ
         .Value = Format(Date, "MMYYYYDD")
      End With

   End If
   Resume END_GET_IT

ERROR_handler:
   Resume Next

End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : IDTExtensibility_OnDisconnection
   ' * Parameters       :
   ' *                    ByVal RemoveMode As vbext_DisconnectMode
   ' *                    custom() As Variant
   ' **********************************************************************
   ' * Comments         : This method removes the Add-In from VB
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Dim nI               As Integer

   ' *** Ignore errors in case the menu items don't exist
   On Error Resume Next

   ' *** Save the setting
   Call SaveSetting(gsREG_APP, "MainCommandBar", "Left", VBInstance.CommandBars(msVBIDEUTILS).left)
   Call SaveSetting(gsREG_APP, "MainCommandBar", "Top", VBInstance.CommandBars(msVBIDEUTILS).top)
   Call SaveSetting(gsREG_APP, "MainCommandBar", "Type", VBInstance.CommandBars(msVBIDEUTILS).Type)
   Call SaveSetting(gsREG_APP, "MainCommandBar", "RowIndex", VBInstance.CommandBars(msVBIDEUTILS).RowIndex)
   Call SaveSetting(gsREG_APP, "MainCommandBar", "Position", VBInstance.CommandBars(msVBIDEUTILS).Position)
   Call SaveSetting(gsREG_APP, "MainCommandBar", "Visible", IIf(VBInstance.CommandBars(msVBIDEUTILS).Visible, 1, 0))
   '   Call SaveSetting(gsREG_APP, "VBIDEError", "Left", VBInstance.CommandBars("VBIDEError").Left)
   '   Call SaveSetting(gsREG_APP, "VBIDEError", "Top", VBInstance.CommandBars("VBIDEError").Top)
   '   Call SaveSetting(gsREG_APP, "VBIDEError", "Type", VBInstance.CommandBars("VBIDEError").Type)
   '   Call SaveSetting(gsREG_APP, "VBIDEError", "Position", VBInstance.CommandBars("VBIDEError").Position)
   '   Call SaveSetting(gsREG_APP, "VBIDEError", "Visible", IIf(VBInstance.CommandBars("VBIDEError").Visible, 1, 0))

   'Call SaveSetting(gsREG_APP, "ProjectExplorer", "Left", gwinProjectExplorer.left)
   'Call SaveSetting(gsREG_APP, "ProjectExplorer", "Top", gwinProjectExplorer.top)
   Call SaveSetting(gsREG_APP, "ProjectExplorer", "Visible", IIf(gwinProjectExplorer.Visible, 1, 0))

   If gbRegistered = False Then
      frmAbout.bAbout = True
      frmAbout.ZOrder
      frmAbout.Show vbModal
   End If

   Unload frmProgress
   Set frmProgress = Nothing

   Unload frmAbout
   Set frmAbout = Nothing

   Unload frmOptions
   Set frmOptions = Nothing

   Unload frmTrayIcon
   Set frmTrayIcon = Nothing

   Unload frmIcons
   Set frmIcons = Nothing

   'Unload frmShortcut
   'Set frmShortcut = Nothing

   Set gcolFind = Nothing

   Set gwinAdvFndRpl = Nothing
   Set gdocAdvFndRpl = Nothing

   Set gwinMouseZoom = Nothing
   Set gdocMouseZoom = Nothing

   Set gwinProjectExplorer = Nothing
   Set gdocProjectExplorer = Nothing

   ' *** Clear down the class handler objects
   For nI = 1 To mnEND_EVENTS
      Set vbeMenus(nI).cMenu = Nothing
   Next
   For nI = 1 To mnEND_CONTROL_EVENTS
      Set vbeControl(nI).cMenu = Nothing
   Next
   For nI = 1 To mnEND_ERROR_EVENTS
      Set vbeErrorControl(nI).cMenu = Nothing
   Next
   For nI = 1 To mnEND_WINDOW_EVENTS
      Set vbeWindowMenu(nI).cMenu = Nothing
   Next
   For nI = 1 To mnEND_ERROR_EVENTS
      Set vbeErrorMenu(nI).cMenu = Nothing
   Next

   ' *** Remove our menus from the VBE menu
   Dim cBar             As CommandBar

   Set cBar = VBInstance.CommandBars(1)

   For nI = 1 To cBar.Controls.Count
      If cBar.Controls(nI).Caption = msmenuVBIDEUTILS Then Exit For
   Next

   cBar.Controls(nI).Delete
   VBInstance.CommandBars(1).FindControl(msoControlPopup, 30165).Controls(msVBIDEERROR).Delete
   VBInstance.CommandBars(1).FindControl(msoControlPopup, 30009).Controls(msUNUSED_WINDOWS).Delete
   VBInstance.CommandBars(1).FindControl(msoControlPopup, 30007).Controls(msADD_NEWPROCEDURE).Delete
   VBInstance.CommandBars("Add-Ins").Controls(msMENU_ABOUT).Delete
   VBInstance.CommandBars("Code Window").Controls(msVBIDEUTILS).Delete
   VBInstance.CommandBars("Code Window").Controls(msMENU_SEARCHVBWEB).Delete
   VBInstance.CommandBars("Immediate Window").Controls(msVBIDEUTILS).Delete
   VBInstance.CommandBars(msVBIDEUTILS).Delete

   Call UnsetKeyHook

   '   If Not gbRegistered Then
   '      Call vbideSubclasser.unHookVBIDE
   '      Set vbideSubclasser = Nothing
   '   End If

End Sub

Private Sub SetupMainMenus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : SetupMainMenus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Adds the menus to the VBE Edit menu and sets up the
   ' * objects to trap the command bar events for the new controls
   ' *
   ' *
   ' **********************************************************************

   Dim cBar             As CommandBar
   Dim ctlMenu          As CommandBarPopup
   Dim ctlMenu2         As CommandBarPopup
   Dim ctlMenu3         As CommandBarPopup
   Dim ctlMenu4         As CommandBarPopup
   Dim ctlMenu5         As CommandBarPopup
   Dim ctlMenu6         As CommandBarPopup
   Dim ctlMenu7         As CommandBarPopup
   Dim ctlMenu8         As CommandBarPopup
   Dim ctlMenu9         As CommandBarPopup
   Dim ctlMenu10        As CommandBarPopup
   Dim ctlMenu11        As CommandBarPopup

   Dim nI               As Integer
   Dim nJ               As Integer
   Dim nK               As Integer
   Dim nL               As Integer
   Dim nM               As Integer
   Dim nO               As Integer
   Dim nP               As Integer
   Dim nQ               As Integer
   Dim nR               As Integer
   Dim nS               As Integer

   ' *** Ignore errors, so we can set objects, then check if they were set OK
   On Error Resume Next

   ' *** Get the VBE Main menu bar
   Set cBar = VBInstance.CommandBars(1)

   ' *** Find our pop-up menu
   Set ctlMenu2 = cBar.Controls(msVBIDEUTILS)

   ' *** If not found, we need to create our menus
   If ctlMenu2 Is Nothing Then

      ' *** Look for the "Add-Ins" menu item, so we can add ours above it
      For nI = 1 To cBar.Controls.Count
         Debug.Print cBar.Controls(nI).Caption, " ", cBar.Controls(nI).Id
         If cBar.Controls(nI).Id = 30038 Then Exit For
      Next

      ' *** Add our popup menu
      Set ctlMenu2 = cBar.Controls.Add(Type:=msoControlPopup, Before:=nI + 1, temporary:=True)
      ctlMenu2.Caption = msmenuVBIDEUTILS

      ' *** Now add the submenus for our indenting menu items
      With ctlMenu2.Controls
         Set ctlMenu3 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu3.Caption = "&Indentation"

         With ctlMenu3.Controls.Add(Type:=msoControlButton, Parameter:="Proc", temporary:=True)
            .Caption = msMENU_PROC
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.IndentProcedure.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu3.Controls.Add(Type:=msoControlButton, Parameter:="Mod", temporary:=True)
            .Caption = msMENU_MOD
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.IndentModule.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu3.Controls.Add(Type:=msoControlButton, Parameter:="Proj", temporary:=True)
            .Caption = msMENU_PROJ
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.IndentProject.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         Set ctlMenu4 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu4.Caption = "&Headers and comments"

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="ModuleHeader", temporary:=True)
            .Caption = msMENU_MODULE_HEADER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ModHeader.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="ProcHeader", temporary:=True)
            .Caption = msMENU_PROC_HEADER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ProcHeader.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="RemoveProcHeader", temporary:=True)
            .Caption = msMENU_REM_PROC_HEADER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ProcHeader.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="ModuleProcHeader", temporary:=True)
            .Caption = msMENU_MOD_PROC_HEADER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ModProcHeader.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="RemoveModuleProcHeader", temporary:=True)
            .Caption = msMENU_REM_MOD_PROC_HEADER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.RemoveAll.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="ProjProcHeader", temporary:=True)
            .Caption = msMENU_PROJ_PROC_HEADER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ProjProcHeader.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="CommentOut", temporary:=True)
            .Caption = msMENU_COMMENT_OUT
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.AllComment.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="UnComment", temporary:=True)
            .Caption = msMENU_UNCOMMENT
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.RemoveAll.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="Inline", temporary:=True)
            .Caption = msMENU_INLINE
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Inline.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu4.Controls.Add(Type:=msoControlButton, Parameter:="VBDoc", temporary:=True)
            .Caption = msMENU_VBDOC
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.pictVBDoc.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         Set ctlMenu5 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu5.Caption = "Expor&t code"

         With ctlMenu5.Controls.Add(Type:=msoControlButton, Parameter:="HTMLProcedure", temporary:=True)
            .Caption = msMENU_HTML_PROC
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ExportHTML.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu5.Controls.Add(Type:=msoControlButton, Parameter:="HTMLModule", temporary:=True)
            .Caption = msMENU_HTML_MOD
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ExportHTML.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         Set ctlMenu8 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu8.Caption = "Error Handler"

         With ctlMenu8.Controls.Add(Type:=msoControlButton, Parameter:="ErrorProc", temporary:=True)
            .Caption = msERROR_PROC
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ErrorProcedure.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu8.Controls.Add(Type:=msoControlButton, Parameter:="ErrorMod", temporary:=True)
            .Caption = msERROR_MOD
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ErrorModule.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu8.Controls.Add(Type:=msoControlButton, Parameter:="EnhancedErrorProc", temporary:=True)
            .BeginGroup = True
            .Caption = msENHANCEDERROR_PROC
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.EnhancedErrorProc.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu8.Controls.Add(Type:=msoControlButton, Parameter:="EnhancedErrorMod", temporary:=True)
            .Caption = msENHANCEDERROR_MOD
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.EnhancedErrorMod.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu8.Controls.Add(Type:=msoControlButton, Parameter:="EnhancedErrorProj", temporary:=True)
            .Caption = msENHANCEDERROR_PROJ
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.EnhancedErrorMod.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         Set ctlMenu6 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu6.Caption = "&Line numbering"

         With ctlMenu6.Controls.Add(Type:=msoControlButton, Parameter:="AddLineNumber", temporary:=True)
            .Caption = msMENU_ADDLINENUMBER
            DoEvents
            Clipboard.SetData frmIcons.AddLineNumbering.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu6.Controls.Add(Type:=msoControlButton, Parameter:="RemoveLineNumber", temporary:=True)
            .Caption = msMENU_REMOVELINENUMBER
            DoEvents
            Clipboard.SetData frmIcons.RemoveLineNumbering.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         Set ctlMenu11 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu11.Caption = msMENU_ALIGNCONTROLS

         With ctlMenu11.Controls.Add(Type:=msoControlButton, Parameter:="AlignControlsLeft", temporary:=True)
            .Caption = msMENU_ALIGNLEFT
            DoEvents
         End With

         With ctlMenu11.Controls.Add(Type:=msoControlButton, Parameter:="AlignControlsRight", temporary:=True)
            .Caption = msMENU_ALIGNRIGHT
            DoEvents
         End With

         With ctlMenu11.Controls.Add(Type:=msoControlButton, Parameter:="AlignControlsTop", temporary:=True)
            .Caption = msMENU_ALIGNTOP
            DoEvents
         End With

         With ctlMenu11.Controls.Add(Type:=msoControlButton, Parameter:="AlignControlsBottom", temporary:=True)
            .Caption = msMENU_ALIGNBOTTOM
            DoEvents
         End With

         Set ctlMenu7 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu7.Caption = "C&ontrols"

         With ctlMenu7.Controls.Add(Type:=msoControlButton, Parameter:="TabOrder", temporary:=True)
            .Caption = msMENU_TABORDER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.TabOrder.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu7.Controls.Add(Type:=msoControlButton, Parameter:="ControlProperties", temporary:=True)
            .Caption = msMENU_PROPERTIES
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Properties.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu7.Controls.Add(Type:=msoControlButton, Parameter:="MouseZoom", temporary:=True)
            .Caption = msMENU_MOUSEZOOM
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.MouseZoom.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu7.Controls.Add(Type:=msoControlButton, Parameter:="Accelerators", temporary:=True)
            .Caption = msMENU_ACCELERATOR
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Accelerator.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu7.Controls.Add(Type:=msoControlButton, Parameter:="GenerateDLLBaseAdress", temporary:=True)
            .Caption = msMENU_DLLBASEADDRESS
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.GenerateDLLBaseAdress.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu7.Controls.Add(Type:=msoControlButton, Parameter:="GenerateGUID", temporary:=True)
            .Caption = msMENU_GUIDID
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.GenerateGUID.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         Set ctlMenu9 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu9.Caption = "&Database"

         With ctlMenu9.Controls.Add(Type:=msoControlButton, Parameter:="DBCreator", temporary:=True)
            .Caption = msMENU_DBCREATOR

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.DBCreator.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu9.Controls.Add(Type:=msoControlButton, Parameter:="ADOConnectionCreator", temporary:=True)
            .Caption = msMENU_ADOCONNCREATOR

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ADOConnectionCreator.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         Set ctlMenu10 = ctlMenu2.Controls.Add(Type:=msoControlPopup, temporary:=True)
         ctlMenu10.Caption = "&Tools"

         With ctlMenu10.Controls.Add(Type:=msoControlButton, Parameter:="ListOLEServers", temporary:=True)
            .Caption = msMENU_LISTOLESERVERS
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ListOLEServers.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu10.Controls.Add(Type:=msoControlButton, Parameter:="IconExplorer", temporary:=True)
            .Caption = msMENU_ICONEXPLORER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.IconExplorer.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu10.Controls.Add(Type:=msoControlButton, Parameter:="ActiveXExplorer", temporary:=True)
            .Caption = msMENU_ACTIVEXEXPLORER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ActiveXExplorer.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu10.Controls.Add(Type:=msoControlButton, Parameter:="ShowKeyCode", temporary:=True)
            .Caption = msMENU_KEYCODE
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.KeyCode.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu10.Controls.Add(Type:=msoControlButton, Parameter:="CharPicker", temporary:=True)
            .Caption = msMENU_CHARPICKER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.CharPicker.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With ctlMenu10.Controls.Add(Type:=msoControlButton, Parameter:="ColorPicker", temporary:=True)
            .Caption = msMENU_COLORPICKER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ColorPicker.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="FindReplace", temporary:=True)
            .BeginGroup = True
            .Caption = msMENU_FINDREPLACE

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.FindAndReplace.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Swap=", temporary:=True)
            .BeginGroup = True
            .Caption = msMENU_SWAP

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Swap.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Alphabetize", temporary:=True)
            .Caption = msMENU_ALPHABETIZE
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Alphabetize.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Pending", temporary:=True)
            .Caption = msMENU_PENDING

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Pending.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="StringExtractor", temporary:=True)
            .Caption = msMENU_STRINGEXTRACT

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.StringExtractor.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="ClearDebug", temporary:=True)
            .Caption = msMENU_CLEAR_DEBUG

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ClearDebug.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="APIError", temporary:=True)
            .Caption = msMENU_API_ERROR
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.APIError.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="ClassSpy", temporary:=True)
            .Caption = msMENU_CLASSNAME
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ClassSpy.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="MsgBox", temporary:=True)
            .Caption = msMENU_MSGBOX
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.MsgBox.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Dependencies", temporary:=True)
            .Caption = msMENU_DEPENDENCIES
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Dependencies.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="ProjectExplorer", temporary:=True)
            .Caption = msMENU_PROJECTEXPLORER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ProjectExplorer.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="ProjectAnalyzer", temporary:=True)
            .Caption = msMENU_PROJECTEANALYZER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ProjectAnalyzer.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="ToolbarCode", temporary:=True)
            .Caption = msMENU_TOOLBAR
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ToolBar.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="SearchVBWeb", temporary:=True)
            .BeginGroup = True
            .Caption = msMENU_SEARCHVBWEB
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.VBDiamond.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="CodeDatabase", temporary:=True)
            .BeginGroup = True
            .Caption = msMENU_CODEDATABASE

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.CodeDatabase.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Options", temporary:=True)
            .BeginGroup = True
            .Caption = msMENU_FORM
            .FaceId = 222
         End With

         If gbRegistered = False Then
            With .Add(Type:=msoControlButton, Parameter:="Register", temporary:=True)
               .BeginGroup = True
               .Caption = msMENU_REGISTER
               ' *** Copy the icon to the clipboard
               DoEvents
               Clipboard.SetData frmIcons.Register.Image, vbCFBitmap
               DoEvents

               ' *** Set the icon for the button
               .PasteFace
               DoEvents
            End With
         End If

         With .Add(Type:=msoControlButton, Parameter:="About", temporary:=True)
            .BeginGroup = True
            .Caption = msMENU_ABOUT
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.VBIDEUtils.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

      End With
   End If

   ' *** Set up the command bar event handlers for the new menu items
   Set vbeMenus(mnSTART_EDITMENU).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu2.Controls(msMENU_PROC))
   For nI = 1 To ctlMenu2.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu2.Controls(nI))
   Next

   For nJ = 1 To ctlMenu3.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu3.Controls(nJ))
   Next

   For nK = 1 To ctlMenu4.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ + nK).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu4.Controls(nK))
   Next

   For nL = 1 To ctlMenu5.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ + nK + nL).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu5.Controls(nL))
   Next

   For nM = 1 To ctlMenu6.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ + nK + nL + nM).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu6.Controls(nM))
   Next

   For nO = 1 To ctlMenu7.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ + nK + nL + nM + nO).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu7.Controls(nO))
   Next

   For nP = 1 To ctlMenu8.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ + nK + nL + nM + nO + nP).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu8.Controls(nP))
   Next

   For nQ = 1 To ctlMenu9.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ + nK + nL + nM + nO + nP + nQ).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu9.Controls(nQ))
   Next

   For nR = 1 To ctlMenu10.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ + nK + nL + nM + nO + nP + nQ + nR).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu10.Controls(nR))
   Next

   For nS = 1 To ctlMenu11.Controls.Count
      Set vbeMenus(mnSTART_EDITMENU + nI + nJ + nK + nL + nM + nO + nP + nQ + nR + nS).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu11.Controls(nS))
   Next

End Sub

Private Sub SetupWindowMenus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : SetupWindowMenus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Adds the menus to the VBE Window menu and sets up the
   ' * objects to trap the command bar events for the new controls
   ' *
   ' *
   ' **********************************************************************

   Dim ctlBar           As CommandBarControl
   Dim ctlMenu          As CommandBarPopup
   Dim ctlMenu2         As CommandBarButton
   Dim nI               As Integer

   ' *** Ignore errors, so we can set objects, then check if they were set OK
   On Error Resume Next

   ' *** Get the VBE Main menu bar
   Set ctlBar = VBInstance.CommandBars(1).FindControl(msoControlPopup, 30009)

   ' *** Find our item
   Set ctlMenu2 = ctlBar.Controls(msUNUSED_WINDOWS)

   ' *** If not found, we need to create our menus
   If ctlMenu2 Is Nothing Then

      ' *** Look for the "Arrange Icons" menu item, so we can add ours after it
      For nI = 1 To ctlBar.Controls.Count
         If ctlBar.Controls(nI).Id = 1435 Then Exit For
      Next

      ' *** Add our item
      Set ctlMenu2 = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ClearUnused", Before:=nI + 1, temporary:=True)
      ctlMenu2.Caption = msUNUSED_WINDOWS
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.CloseAllWindows.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      ctlMenu2.PasteFace

      ctlMenu2.BeginGroup = True
   End If

   Set vbeWindowMenu(mnSTART_WINDOW_EVENTS).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu2)

End Sub

Private Sub SetupToolsMenus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : SetupToolsMenus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Adds the menus to the VBE Window menu and sets up the
   ' * objects to trap the command bar events for the new controls
   ' *
   ' *
   ' **********************************************************************

   Dim ctlBar           As CommandBarControl
   Dim ctlMenu          As CommandBarPopup
   Dim ctlMenu2         As CommandBarButton
   Dim nI               As Integer

   ' *** Ignore errors, so we can set objects, then check if they were set OK
   On Error Resume Next

   ' *** Get the VBE Main menu bar
   Set ctlBar = VBInstance.CommandBars(1).FindControl(msoControlPopup, 30007)

   ' *** Find our item
   Set ctlMenu2 = ctlBar.Controls(msADD_NEWPROCEDURE)

   ' *** If not found, we need to create our menus
   If ctlMenu2 Is Nothing Then

      ' *** Look for the "Arrange Icons" menu item, so we can add ours after it
      For nI = 1 To ctlBar.Controls.Count
         If ctlBar.Controls(nI).Id = 1435 Then Exit For
      Next

      nI = 0

      ' *** Add our item
      Set ctlMenu2 = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="AddNewProcedure", Before:=nI + 1, temporary:=True)
      ctlMenu2.Caption = msADD_NEWPROCEDURE
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.PropertyBuilder.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      ctlMenu2.PasteFace

      ctlMenu2.BeginGroup = True
   End If

   Set vbeWindowMenu(mnSTART_WINDOW_EVENTS + 2).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu2)

End Sub

Private Sub SetupDebugMenus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : SetupDebugMenus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Adds the menus to the VBE Debug menu and sets up the
   ' * objects to trap the command bar events for the new controls
   ' *
   ' *
   ' **********************************************************************

   Dim ctlBar           As CommandBarControl
   Dim ctlMenu          As CommandBarPopup
   Dim ctlMenu2         As CommandBarPopup
   Dim nI               As Integer

   ' *** Ignore errors, so we can set objects, then check if they were set OK
   On Error Resume Next

   ' *** Get the VBE Main menu bar
   Set ctlBar = VBInstance.CommandBars(1).FindControl(msoControlPopup, 30165)

   ' *** Find our item
   Set ctlMenu2 = ctlBar.Controls(msVBIDEERROR)

   ' *** If not found, we need to create our menus
   If ctlMenu2 Is Nothing Then

      ' *** Look for the "Run to Cursor" menu item, so we can add ours after it
      For nI = 1 To ctlBar.Controls.Count
         If ctlBar.Controls(nI).Id = 1811 Then Exit For
      Next

      ' *** Add our item
      Set ctlMenu2 = ctlBar.Controls.Add(Type:=msoControlPopup, Before:=nI + 1, temporary:=True)
      ctlMenu2.Caption = msVBIDEERROR
      ctlMenu2.BeginGroup = True

      ' *** Now add the submenus for our indenting menu items
      With ctlMenu2.Controls
         With .Add(Type:=msoControlButton, Parameter:="ErrorProc", temporary:=True)
            .Caption = msERROR_PROC
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ErrorProcedure.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
         End With

         With .Add(Type:=msoControlButton, Parameter:="ErrorMod", temporary:=True)
            .Caption = msERROR_MOD
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ErrorModule.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
         End With

         With .Add(Type:=msoControlButton, Parameter:="EnhancedErrorProc", temporary:=True)
            .BeginGroup = True
            .Caption = msENHANCEDERROR_PROC
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.EnhancedErrorProc.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
         End With

         With .Add(Type:=msoControlButton, Parameter:="EnhancedErrorMod", temporary:=True)
            .Caption = msENHANCEDERROR_MOD
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.EnhancedErrorMod.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
         End With
      End With

      With ctlMenu2
         Set vbeErrorMenu(mnSTART_ERROR_EVENTS).cMenu = VBInstance.Events.CommandBarEvents(.Controls(msERROR_PROC))
         Set vbeErrorMenu(mnSTART_ERROR_EVENTS + 1).cMenu = VBInstance.Events.CommandBarEvents(.Controls(msERROR_MOD))
         Set vbeErrorMenu(mnSTART_ERROR_EVENTS + 2).cMenu = VBInstance.Events.CommandBarEvents(.Controls(msENHANCEDERROR_PROC))
         Set vbeErrorMenu(mnSTART_ERROR_EVENTS + 3).cMenu = VBInstance.Events.CommandBarEvents(.Controls(msENHANCEDERROR_MOD))
      End With
   End If

End Sub

Private Sub PopupCodeMenu()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : PopupCodeMenu
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Adds the menus to the VBE Code Window menu and sets up the objects to trap
   ' * the command bar events for the new controls
   ' *
   ' *
   ' **********************************************************************

   Dim ctlBar           As CommandBar
   Dim ctlMenu          As CommandBarPopup
   Dim ctlMenu1         As CommandBarButton
   Dim ctlMenu2         As CommandBarPopup
   Dim nI               As Integer

   ' *** Ignore errors, so we can set objects, then check if they were set OK
   On Error Resume Next

   ' *** VBE Code Window shortcut Menu
   Set ctlBar = VBInstance.CommandBars("Code Window")

   ' *** See if our VBIDEUtils menu item exists
   Set ctlMenu2 = Nothing
   Set ctlMenu2 = ctlBar.Controls(msVBIDEUTILS)

   ' *** If not, we need to add them
   If ctlMenu2 Is Nothing Then

      ' *** Look for the "Object Browser" menu item, so we can add ours above it
      For nI = 1 To ctlBar.Controls.Count
         If ctlBar.Controls(nI).Id = 473 Then Exit For
      Next

      Set ctlMenu2 = ctlBar.Controls.Add(Type:=msoControlPopup, Before:=nI, temporary:=True)
      ctlMenu2.Caption = msVBIDEUTILS
      ctlMenu2.BeginGroup = True

      Set ctlMenu1 = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="SearchVBWeb", Before:=nI + 1, temporary:=True)
      With ctlMenu1
         .Caption = msMENU_SEARCHVBWEB

         ' *** Copy the icon to the clipboard
         DoEvents
         Clipboard.SetData frmIcons.VBDiamond.Image, vbCFBitmap
         DoEvents

         ' *** Set the icon for the button
         .PasteFace
         DoEvents
      End With

      ' *** Now add our menu items
      With ctlMenu2.Controls
         With .Add(Type:=msoControlButton, Parameter:="Proc", temporary:=True)
            .Caption = msMENU_PROC
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.IndentProcedure.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Mod", temporary:=True)
            .Caption = msMENU_MOD
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.IndentModule.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
         End With

         With .Add(Type:=msoControlButton, Parameter:="Proj", temporary:=True)
            .Caption = msMENU_PROJ
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.IndentProject.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
         End With

         With .Add(Type:=msoControlButton, Parameter:="ModuleHeader", temporary:=True)
            .BeginGroup = True
            .Caption = msMENU_MODULE_HEADER

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ModHeader.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="ProcHeader", temporary:=True)
            .Caption = msMENU_PROC_HEADER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ProcHeader.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="ModuleProcHeader", temporary:=True)
            .Caption = msMENU_MOD_PROC_HEADER
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ModProcHeader.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="CommentOut", temporary:=True)
            .Caption = msMENU_COMMENT_OUT
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.AllComment.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="UnComment", temporary:=True)
            .Caption = msMENU_UNCOMMENT
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.RemoveAll.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Inline", temporary:=True)
            .Caption = msMENU_INLINE
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Inline.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Swap=", temporary:=True)
            .BeginGroup = True
            .Caption = msMENU_SWAP

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Swap.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Alphabetize", temporary:=True)
            .Caption = msMENU_ALPHABETIZE
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Alphabetize.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="Pending", temporary:=True)
            .Caption = msMENU_PENDING

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.Pending.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

         With .Add(Type:=msoControlButton, Parameter:="ClearDebug", temporary:=True)
            .Caption = msMENU_CLEAR_DEBUG

            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ClearDebug.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With

      End With
   End If

   ' *** Set up the command bar event handlers for the new menu items
   Set vbeMenus(mnSTART_WINDOWSMENU).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu2.Controls(msMENU_PROC))
   Set vbeMenus(mnSTART_WINDOWSMENU + 1).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu1)
   For nI = 2 To ctlMenu2.Controls.Count
      Set vbeMenus(mnSTART_WINDOWSMENU + nI).cMenu = VBInstance.Events.CommandBarEvents(ctlMenu2.Controls(nI))
   Next

   ' *** VBE Code Window shortcut Menu
   Set ctlBar = VBInstance.CommandBars("Immediate Window")

   ' *** See if our VBIDEUtils menu item exists
   Set ctlMenu2 = Nothing
   Set ctlMenu2 = ctlBar.Controls(msVBIDEUTILS)

   ' *** If not, we need to add them
   If ctlMenu2 Is Nothing Then

      ' *** Look for the "Object Browser" menu item, so we can add ours above it
      For nI = 1 To ctlBar.Controls.Count
         If ctlBar.Controls(nI).Id = 473 Then Exit For
      Next

      Set ctlMenu2 = ctlBar.Controls.Add(Type:=msoControlPopup, Before:=nI, temporary:=True)
      ctlMenu2.Caption = msVBIDEUTILS

      ' *** Now add our indenting menu items (in reverse order)
      With ctlMenu2.Controls
         With .Add(Type:=msoControlButton, Parameter:="ClearDebug", temporary:=True)
            .Caption = msMENU_CLEAR_DEBUG
            ' *** Copy the icon to the clipboard
            DoEvents
            Clipboard.SetData frmIcons.ClearDebug.Image, vbCFBitmap
            DoEvents

            ' *** Set the icon for the button
            .PasteFace
            DoEvents
         End With
      End With
   End If

   With ctlMenu2
      Set vbeMenus(mnSTART_DEBUGWINDOWSMENU).cMenu = VBInstance.Events.CommandBarEvents(.Controls(msMENU_CLEAR_DEBUG))
   End With

End Sub

Private Sub SetupNewCommandBar()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : SetupNewCommandBar
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Adds the a CommandBar to the VB environment and
   ' * sets up the objects to trap the command bar events for the new controls
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_SetupNewCommandBar

   Dim ctlBar           As CommandBar
   Dim mButton          As Office.CommandBarControl

   Dim nRetry           As Integer
   nRetry = 0

   ' *** Ignore errors, so we can set objects, then check if they were set OK
   Clipboard.Clear

   ' *** Add a new CommandBar
   Set ctlBar = VBInstance.CommandBars.Add(msVBIDEUTILS, Position:=msoBarTop)

   ' *** Now add all the controls
   ' *** Indent
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="Proc", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_PROC)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.IndentProcedure.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msMENU_PROC)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS).cMenu = VBInstance.Events.CommandBarEvents(ctlBar.Controls(msMENU_PROC))

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="Mod", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_MOD)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.IndentModule.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msMENU_MOD)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 1).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** Comment
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ModuleHeader", temporary:=True)
   With mButton
      .BeginGroup = True

      .Caption = RemoveAmpersand(msMENU_MODULE_HEADER)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ModHeader.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_MODULE_HEADER)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 3).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ProcHeader", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_PROC_HEADER)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ProcHeader.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_PROC_HEADER)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 4).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ModuleProcHeader", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_MOD_PROC_HEADER)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ModProcHeader.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_MOD_PROC_HEADER)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 5).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="CommentOut", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_COMMENT_OUT)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.AllComment.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_COMMENT_OUT)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 6).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="UnComment", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_UNCOMMENT)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.RemoveAll.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_UNCOMMENT)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 7).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="Inline", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_INLINE)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.Inline.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_INLINE)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 8).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="VBDoc", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_VBDOC)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.pictVBDoc.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_VBDOC)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 9).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="Swap=", temporary:=True)
   With mButton
      .BeginGroup = True
      .Caption = RemoveAmpersand(msMENU_SWAP)

      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.Swap.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_SWAP)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 10).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="Alphabetize", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_ALPHABETIZE)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.Alphabetize.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_ALPHABETIZE)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 11).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="Pending", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_PENDING)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.Pending.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_PENDING)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 12).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="FindReplace", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_FINDREPLACE)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.FindAndReplace.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_FINDREPLACE)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 13).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** Debug
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ClearDebug", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_CLEAR_DEBUG)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ClearDebug.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_CLEAR_DEBUG)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 14).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** API Error
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="APIError", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_API_ERROR)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.APIError.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_API_ERROR)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 15).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** Properties
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ControlProperties", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_PROPERTIES)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.Properties.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_PROPERTIES)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 16).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** Properties
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="MouseZoom", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_MOUSEZOOM)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.MouseZoom.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_MOUSEZOOM)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 17).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** MessageBox
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="MsgBox", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_MSGBOX)

      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.MsgBox.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_MSGBOX)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 18).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** Search VB Web
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="SearchVBWeb", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_SEARCHVBWEB)

      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.VBDiamond.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_SEARCHVBWEB)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 19).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** DBCreator
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ProjectExplorer", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msMENU_PROJECTEXPLORER)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ProjectExplorer.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_PROJECTEXPLORER)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 20).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   '   ' *** Toolbar Code
   '   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ToolbarCode", temporary:=True)
   '   With mButton
   '      .Caption = RemoveAmpersand(msMENU_TOOLBAR)
   '      ' *** Copy the icon to the clipboard
   '         DoEvents
   '      Clipboard.SetData frmIcons.ToolBar.Image, vbCFBitmap
   '
   '      ' *** Set the icon for the button
   '      .PasteFace
   '      DoEvents
   '      .ToolTipText = RemoveAmpersand(msMENU_TOOLBAR)
   '      .Enabled = True
   '   End With
   '   Set vbeControl(mnSTART_CONTROL_EVENTS + 20).cMenu = VBInstance.Events.CommandBarEvents(mButton)
   '
   ' *** Code Database
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="CodeDatabase", temporary:=True)
   With mButton
      .BeginGroup = True
      .Caption = RemoveAmpersand(msMENU_CODEDATABASE)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.CodeDatabase.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      DoEvents
      .ToolTipText = RemoveAmpersand(msMENU_CODEDATABASE)
      .Enabled = True
   End With
   Set vbeControl(mnSTART_CONTROL_EVENTS + 21).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** Error procedure
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ErrorProc", temporary:=True)
   With mButton
      .BeginGroup = True
      .Caption = RemoveAmpersand(msERROR_PROC)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ErrorProcedure.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msERROR_PROC)
      .Enabled = True
   End With
   Set vbeErrorControl(mnSTART_ERROR_EVENTS).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ErrorMod", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msERROR_MOD)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ErrorModule.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msERROR_MOD)
      .Enabled = True
   End With
   Set vbeErrorControl(mnSTART_ERROR_EVENTS + 1).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** Enhanced Error procedure
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="EnhancedErrorProc", temporary:=True)
   With mButton
      .BeginGroup = True

      .Caption = RemoveAmpersand(msENHANCEDERROR_PROC)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.EnhancedErrorProc.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msENHANCEDERROR_PROC)
      .Enabled = True
   End With
   Set vbeErrorControl(mnSTART_ERROR_EVENTS + 3).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="EnhancedErrorMod", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msENHANCEDERROR_MOD)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.EnhancedErrorMod.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msENHANCEDERROR_MOD)
      .Enabled = True
   End With
   Set vbeErrorControl(mnSTART_ERROR_EVENTS + 4).cMenu = VBInstance.Events.CommandBarEvents(mButton)

EXIT_SetupNewCommandBar:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_SetupNewCommandBar:
   If nRetry <= 40 Then
      nRetry = nRetry + 1
      Resume
   End If

   Select Case MsgBox("Error " & err.number & ": " & err.Description & vbCrLf & "in SetupNewCommandBar " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_SetupNewCommandBar
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select

   Resume EXIT_SetupNewCommandBar

End Sub

Private Sub CtlHandler_ItemAdded(ByVal VBControl As VBIDE.VBControl)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : CtlHandler_ItemAdded
   ' * Parameters       :
   ' *                    ByVal VBControl As VBIDE.VBControl
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** This event fires when a control is added to the current form in the IDE
   If IsLoaded(frmTabOrder.Name) Then frmTabOrder.ControlAdded VBControl

End Sub

Private Sub CtlHandler_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : CtlHandler_ItemRenamed
   ' * Parameters       :
   ' *                    ByVal VBControl As VBIDE.VBControl
   ' *                    ByVal OldName As String
   ' *                    ByVal OldIndex As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** This event fires when a control is renamed on the current form in the IDE
   If IsLoaded(frmTabOrder.Name) Then frmTabOrder.ControlRenamed VBControl, OldName, OldIndex

End Sub

Private Sub CtlHandler_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : CtlHandler_ItemRemoved
   ' * Parameters       :
   ' *                    ByVal VBControl As VBIDE.VBControl
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** This event fires when a control is removed from the current form in the IDE
   If IsLoaded(frmTabOrder.Name) Then frmTabOrder.ControlRemoved VBControl

End Sub

Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : CmpHandler_ItemActivated
   ' * Parameters       :
   ' *                    ByVal VBComponent As VBIDE.VBComponent
   ' **********************************************************************
   ' * Comments         :
   ' *   This event fires when a form becomes activated in the IDE
   ' *
   ' **********************************************************************

   On Error GoTo CmpHandler_ItemActivatedErr

   If IsLoaded(frmTabOrder.Name) Then frmTabOrder.RefreshList 0

CmpHandler_ItemActivatedErr:

End Sub

Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : CmpHandler_ItemSelected
   ' * Parameters       :
   ' *                    ByVal VBComponent As VBIDE.VBComponent
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** This event fires when a form is selected in the project window

   CmpHandler_ItemActivated VBComponent

End Sub

' Listing 4: List all CommandBarControls
Sub DumpAllMenus(Optional TopMenu As String = "Menu Bar")
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : DumpAllMenus
   ' * Parameters       :
   ' *                    Optional TopMenu As String = "Menu Bar"
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cbr              As CommandBar
   Dim cbp              As CommandBarPopup

   Set cbr = VBInstance.CommandBars("Menu Bar")
   If TopMenu <> "Menu Bar" Then
      Set cbp = cbr.Controls(TopMenu)
      Call DumpMenu(cbp, 1)
   Else
      For Each cbp In cbr.Controls
         Debug.Print cbp.Caption, cbp.Id
         Call DumpMenu(cbp, 1)
      Next cbp
   End If

End Sub

Sub DumpMenu(cbp As CommandBarPopup, intLevel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 12:44
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : DumpMenu
   ' * Parameters       :
   ' *                    cbp As CommandBarPopup
   ' *                    intLevel As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cbc              As CommandBarControl
   Dim intI             As Integer

   For Each cbc In cbp.CommandBar.Controls
      ' Insert enough spaces to indent according to the
      ' level of recursion.
      For intI = 0 To intLevel
         Debug.Print "   ";
      Next
      On Error Resume Next
      Debug.Print cbc.Caption, cbc.Id, cbc.ShortcutText
      If cbc.Type = msoControlPopup Then
         ' Call this routine recursively, to document
         ' the next lower level.
         Call DumpMenu(cbc.Control, intLevel + 1)
      End If
   Next
End Sub

Private Sub SetupErrorCommandBar()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/1998
   ' * Time             : 22:06
   ' * Module Name      : clsConnect
   ' * Module Filename  : clsConnect.cls
   ' * Procedure Name   : SetupErrorCommandBar
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Adds the Erro CommandBar to the VB environment and
   ' * sets up the objects to trap the command bar events for the new controls
   ' *
   ' *
   ' **********************************************************************

   Dim ctlBar           As CommandBar
   Dim mButton          As Office.CommandBarControl

   ' *** Ignore errors, so we can set objects, then check if they were set OK
   On Error Resume Next

   ' *** Add a new CommandBar
   Set ctlBar = VBInstance.CommandBars.Add("VBIDEError", Position:=msoBarTop)

   ' *** Now add all the controls

   ' *** Error procedure
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ErrorProc", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msERROR_PROC)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ErrorProcedure.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msERROR_PROC)
      .Enabled = True
   End With
   Set vbeErrorControl(mnSTART_ERROR_EVENTS).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="ErrorMod", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msERROR_MOD)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.ErrorModule.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msERROR_MOD)
      .Enabled = True
   End With
   Set vbeErrorControl(mnSTART_ERROR_EVENTS + 1).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   ' *** Enhanced Error procedure
   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="EnhancedErrorProc", temporary:=True)
   With mButton
      .BeginGroup = True

      .Caption = RemoveAmpersand(msENHANCEDERROR_PROC)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.EnhancedErrorProc.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msENHANCEDERROR_PROC)
      .Enabled = True
   End With
   Set vbeErrorControl(mnSTART_ERROR_EVENTS + 3).cMenu = VBInstance.Events.CommandBarEvents(mButton)

   Set mButton = ctlBar.Controls.Add(Type:=msoControlButton, Parameter:="EnhancedErrorMod", temporary:=True)
   With mButton
      .Caption = RemoveAmpersand(msENHANCEDERROR_MOD)
      ' *** Copy the icon to the clipboard
      DoEvents
      Clipboard.SetData frmIcons.EnhancedErrorMod.Image, vbCFBitmap
      DoEvents

      ' *** Set the icon for the button
      .PasteFace
      .ToolTipText = RemoveAmpersand(msENHANCEDERROR_MOD)
      .Enabled = True
   End With
   Set vbeErrorControl(mnSTART_ERROR_EVENTS + 4).cMenu = VBInstance.Events.CommandBarEvents(mButton)

End Sub

Private Sub PrjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
   On Error Resume Next
   gdocAdvFndRpl.Initialize
End Sub
