VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVBDocumentor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBDoc Generator"
   ClientHeight    =   4845
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   9960
   Icon            =   "VBDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSetDescriptionAttribute 
      Caption         =   "Set descriptin to the attribute"
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   2880
      TabIndex        =   15
      Top             =   3480
      Width           =   1995
   End
   Begin VB.CheckBox chkCheck 
      Caption         =   "All/None"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2200
      Width           =   1095
   End
   Begin VB.CommandButton cmdViewDoc 
      Caption         =   "View Doc"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   2220
      Width           =   1335
   End
   Begin VB.TextBox txtDOSOutputs 
      Height          =   1455
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3240
      Width           =   4935
   End
   Begin VB.CheckBox chkSourceCode 
      Caption         =   "Source Code"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2880
      TabIndex        =   5
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CheckBox chkVarAsProperty 
      Caption         =   "Variables like properties"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   2820
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   9480
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleziona il Compilatore"
      Filter          =   "Executables(*.exe)|*.exe|"
   End
   Begin VB.TextBox txtCompiler 
      Height          =   285
      Left            =   4920
      TabIndex        =   9
      Top             =   2880
      Width           =   4455
   End
   Begin MSComctlLib.ListView lstTags 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tag Type"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Comment"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.TreeView trvComps 
      Height          =   2235
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3942
      _Version        =   393217
      Indentation     =   529
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.CheckBox chkPublicOnly 
      Caption         =   "Public member only"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   2220
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Generate"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Help compiler"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4920
      TabIndex        =   14
      Top             =   2640
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Project Components"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2220
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comments Tag"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   1065
   End
End
Attribute VB_Name = "frmVBDocumentor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : frmAddIn
' * Module Filename  : VBDoc.frm
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Purpose: This Form is the Interface with the user. The most part of this
'   code is generated automatically by the Visual Basic Add-In Wizard.<BR>
'   Other Code is added for visualization of the CHM file and for the visualization
'   of result of compilation.<BR>
'Author:    Marco Pipino
Option Explicit

Private Const REG_SZ    As Long = 1
Private Const HKEY_CLASSES_ROOT As Long = &H80000000

'Purpose: This declaration is used for read the path of application
'   for read the CHM File.<Br>
'   This information is stored in the registry in the key<BR>
'   <B>HKEY_CLASSES_ROOT\.chm\shell\open\command</B>
Private Declare Function RegOpenKey Lib "advapi32.dll" _
   Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As _
   String, phkResult As Long) As Long

'Purpose: This declaration is used for read the path of application
'   for read the CHM File.<Br>
'   This information is stored in the registry in the key<BR>
'   <B>HKEY_CLASSES_ROOT\.chm\shell\open\command</B>
Private Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long

'Purpose: This declaration is used for read the path of application
'   for read the CHM File.<Br>
'   This information is stored in the registry in the key<BR>
'   <B>HKEY_CLASSES_ROOT\.chm\shell\open\command</B>
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
   Alias "RegQueryValueExA" (ByVal hKey As Long, _
   ByVal lpValueName As String, ByVal lpReserved As Long, _
   lpType As Long, lpData As Any, lpcbData As Long) As Long

Private ItemClicked     As ListItem         'User for managing the ListBox

Private WithEvents objDos As DOSOutputs 'DOSOutput object for viewing
Attribute objDos.VB_VarHelpID = -1
'the optputs of the hhc compiler

'Purpose: Hide the form
Private Sub CancelButton_Click()
   Unload Me
End Sub

Private Sub chkCheck_Click()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : chkCheck_Click
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Long

   For nI = 1 To trvComps.Nodes.Count
      trvComps.Nodes(nI).Checked = IIf(chkCheck.Value = vbChecked, True, False)
   Next

End Sub

'Purpose: Control of the check for Technical documentation.<BR>
'   Remove comments in the code if you want remove mixed documentations.
Private Sub chkPublicOnly_Click()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : chkPublicOnly_Click
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   '    If chkPublicOnly.Value = 1 Then
   '        chkVarAsProperty.Value = 1
   '        chkVarAsProperty.Enabled = False
   '        chkSourceCode.Value = 0
   '        chkSourceCode.Enabled = False
   '    Else
   '        chkVarAsProperty.Enabled = True
   '        chkSourceCode.Enabled = True
   '    End If
End Sub

'Purpose: Select the compiler
Private Sub cmdBrowse_Click()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : cmdBrowse_Click
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   CommonDialog1.ShowOpen
   If Len(CommonDialog1.FileName) > 0 Then
      txtCompiler.Text = CommonDialog1.FileName
   End If
End Sub

'Purpose: After copilation we can view the CHM file. <BR>
'   We must read the <B>HKEY_CLASSES_ROOT\.chm\shell\open\command</B> registry key
'   for the path of the HH.exe viewer.
Private Sub cmdViewDoc_Click()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : cmdViewDoc_Click
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error GoTo cmvViewDoc_Error

   '   Dim lngRet           As Long
   '   Dim lngKey           As Long
   '   Dim lngKeyType       As Long
   '   Dim strBuffer        As String
   '   Dim lngBufferSize    As Long

   Call ExecuteWithAssociate(Me.hWnd, gProjectFolder & "\" & gProjectName & ".chm", , vbNormalFocus)

   Exit Sub

   '   strBuffer = Space(256)
   '   'Open the registry key
   '   lngRet = RegOpenKey(HKEY_CLASSES_ROOT, ".chm\shell\open\command", lngKey)
   '   If lngRet <> 0 Then
   '      GoTo cmvViewDoc_Error
   '   End If
   '
   '   'Get the key type value
   '   lngRet = RegQueryValueEx(lngKey, "", 0&, lngKeyType, ByVal strBuffer, lngBufferSize)
   '   If lngKeyType <> REG_SZ Then
   '      GoTo cmvViewDoc_Error
   '   End If
   '
   '   'Get the value of the key i.e. the path of the HH.exe application
   '   lngRet = RegQueryValueEx(lngKey, "", 0&, REG_SZ, ByVal strBuffer, lngBufferSize)
   '   If lngRet <> 0 Then
   '      GoTo cmvViewDoc_Error
   '   End If
   '
   '   'Close the key
   '   lngRet = RegCloseKey(lngKey)
   '   strBuffer = left$(strBuffer, lngBufferSize)
   '
   '   'Launch the help file
   '   Shell Replace(strBuffer, "%1", gProjectFolder & "\" & gProjectName & ".chm"), vbNormalFocus
   '   Exit Sub
cmvViewDoc_Error:
   MsgBox ("Can't open the file")
End Sub

'Purpose: Load the form and get the setting from the registry
Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : Form_Load
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   cmdViewDoc.Enabled = False
   GetSettings
End Sub

'Purpose: Unload the form
Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : Form_Unload
   ' * Purpose          :
   ' * Parameters       :
   ' *                    Cancel As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set objProject = Nothing
End Sub

'Purpose: With double click we can change the Tags for comments
Private Sub lstTags_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : lstTags_DblClick
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error Resume Next
   Dim strTemp          As String
   strTemp = ItemClicked.ListSubItems(1).Text
   strTemp = InputBox(ItemClicked.ListSubItems(2).Text, , strTemp)
   If Len(strTemp) > 0 Then
      ItemClicked.ListSubItems(1).Text = strTemp
   End If
End Sub

'Purpose: Select the item clicked
Private Sub lstTags_ItemClick(ByVal Item As MSComctlLib.ListItem)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : lstTags_ItemClick
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Item As MSComctlLib.ListItem
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set ItemClicked = Item
End Sub

'Purpose: Event of the DOSOutputs object. Fill the TextBox txtOutputs with
'   the response of the compiler.
Private Sub objDos_ReceiveOutputs(CommandOutputs As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : objDos_ReceiveOutputs
   ' * Purpose          :
   ' * Parameters       :
   ' *                    CommandOutputs As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   txtDOSOutputs.Text = txtDOSOutputs.Text & CommandOutputs
End Sub

'Purpose: Launch the generation of the documentation.<BR>
'   Create a new cProject object, save the current settings to the registry,
'   save the check values and then create the Tree of the project, then the HTML files
'   and the compile.<B>
'Remarks: See the cProject for more information
Private Sub OKButton_Click()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : OKButton_Click
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim clsHourglass     As New class_Hourglass
   Set clsHourglass = New class_Hourglass

   Dim strDosCommand    As String

   Set objProject = New cProject
   Set objProject.VBI = VBInstance
   objProject.GetSettings

   SaveSettings

   gPublicOnly = chkPublicOnly.Value
   'gInsConsts = chkCostants.Value
   gVarAsProperty = chkVarAsProperty.Value
   gSourceCode = chkSourceCode.Value

   txtDOSOutputs.Text = ""

   'Call AlwaysOnTop(frmProgress, True)
   Load frmProgress
   frmProgress.Show
   frmProgress.ZOrder

   objProject.BuildTree Me
   objProject.BuildTypesValue
   objProject.CreateHTMLFiles

   If chkSetDescriptionAttribute.Value = vbChecked Then
      frmProgress.MessageText = "Setting the descriptions for all attributes"
      objProject.SetHelpDescription
   End If

   Unload frmProgress
   Set frmProgress = Nothing

   Set objDos = New DOSOutputs
   txtDOSOutputs.Text = ""
   strDosCommand = Chr(34) & gHHCCompiler & Chr(34) & " " & Chr(34) & _
      gProjectFolder & "\VBDoc\" & gProjectName & ".hhp" & Chr(34)
   If (objDos.ExecuteCommand(strDosCommand)) Then
      cmdViewDoc.Enabled = True
   Else
      cmdViewDoc.Enabled = False
   End If

   Set objDos = Nothing
   Set objProject = Nothing
End Sub

'Purpose: Get the setting from the registry and fill the Controls on the form.
Private Sub GetSettings()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : GetSettings
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim comp             As VBComponent
   Dim lstItem          As ListItem

   gBLOCK_PURPOSE = GetSetting(gsREG_APP, "Blocks", "Purpose", "Purpose")
   gBLOCK_PROJECT = GetSetting(gsREG_APP, "Blocks", "Project", "Project:")
   gBLOCK_AUTHOR = GetSetting(gsREG_APP, "Blocks", "Author", "Author")
   gBLOCK_DATE_CREATION = GetSetting(gsREG_APP, "Blocks", "Date_Creation", "Date")
   gBLOCK_DATE_LAST_MOD = GetSetting(gsREG_APP, "Blocks", "Date_Last_Mod", "Modified")
   gBLOCK_VERSION = GetSetting(gsREG_APP, "Blocks", "Version", "Version")
   gBLOCK_EXAMPLE = GetSetting(gsREG_APP, "Blocks", "Example", "Example")
   gBLOCK_SEEALSO = GetSetting(gsREG_APP, "Blocks", "SeeAlso", "See Also")
   gBLOCK_SCREEENSHOT = GetSetting(gsREG_APP, "Blocks", "Screenshot", "Screenshot")
   gBLOCK_CODE = GetSetting(gsREG_APP, "Blocks", "Code", "Code")
   gBLOCK_TEXT = GetSetting(gsREG_APP, "Blocks", "Text", "Text")
   gBLOCK_REMARKS = GetSetting(gsREG_APP, "Blocks", "Remarks", "Comments")
   gBLOCK_NO_COMMENT = GetSetting(gsREG_APP, "Blocks", "NoComment", "''")
   gBLOCK_PARAMETER = GetSetting(gsREG_APP, "Blocks", "Parameter", "Parameters")

   gBLOCK_WEBSITE = GetSetting(gsREG_APP, "BLocks", "Web Site", "Web Site")
   gBLOCK_EMAIL = GetSetting(gsREG_APP, "BLocks", "E-Mail", "E-Mail")
   gBLOCK_TIME = GetSetting(gsREG_APP, "BLocks", "Time", "Time")
   gBLOCK_TEL = GetSetting(gsREG_APP, "BLocks", "Telephone", "Telephone")
   gBLOCK_PROCEDURE_NAME = GetSetting(gsREG_APP, "BLocks", "Procedure Name", "Procedure Name")
   gBLOCK_MODULE_NAME = GetSetting(gsREG_APP, "BLocks", "Module Name", "Module Name")
   gBLOCK_MODULE_FILE = GetSetting(gsREG_APP, "BLocks", "Module Filename", "Module Filename")

   gHHCCompiler = GetSetting(gsREG_APP, "Compiler", "Path", "C:\Program Files\HTML Help Workshop\hhc.exe")

   lstTags.ListItems.Add 1, , "PROJECT"
   lstTags.ListItems(1).ListSubItems.Add 1, , gBLOCK_PROJECT
   lstTags.ListItems(1).ListSubItems.Add 2, , "Begin and End of project section"

   lstTags.ListItems.Add 2, , "AUTHOR"
   lstTags.ListItems(2).ListSubItems.Add 1, , gBLOCK_AUTHOR
   lstTags.ListItems(2).ListSubItems.Add 2, , "Author Block"

   lstTags.ListItems.Add 3, , "DATE_CREATION"
   lstTags.ListItems(3).ListSubItems.Add 1, , gBLOCK_DATE_CREATION
   lstTags.ListItems(3).ListSubItems.Add 2, , "Date Creation Block"

   lstTags.ListItems.Add 4, , "DATE_LAST_MOD"
   lstTags.ListItems(4).ListSubItems.Add 1, , gBLOCK_DATE_LAST_MOD
   lstTags.ListItems(4).ListSubItems.Add 2, , "Date of last modify Block"

   lstTags.ListItems.Add 5, , "VERSION"
   lstTags.ListItems(5).ListSubItems.Add 1, , gBLOCK_VERSION
   lstTags.ListItems(5).ListSubItems.Add 2, , "Version of the project"

   Set lstItem = lstTags.ListItems.Add(6, , "PURPOSE")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_PURPOSE)
   Call lstItem.ListSubItems.Add(2, , "Purpose of the block")
   '   Call lstItem.ListSubItems.Add(3, , "###Description###")

   lstTags.ListItems.Add 7, , "EXAMPLE"
   lstTags.ListItems(7).ListSubItems.Add 1, , gBLOCK_EXAMPLE
   lstTags.ListItems(7).ListSubItems.Add 2, , "Example Block"

   Set lstItem = lstTags.ListItems.Add(8, , "REMARKS")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_REMARKS)
   Call lstItem.ListSubItems.Add(2, , "Remarks for the Block")
   '   Call lstItem.ListSubItems.Add(3, , "###Description###")

   lstTags.ListItems.Add 9, , "PARAMETER"
   lstTags.ListItems(9).ListSubItems.Add 1, , gBLOCK_PARAMETER
   lstTags.ListItems(9).ListSubItems.Add 2, , "Parameter Block of a member"

   Set lstItem = lstTags.ListItems.Add(10, , "CODE")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_CODE)
   Call lstItem.ListSubItems.Add(2, , "Code in a Example Block")
   '   Call lstItem.ListSubItems.Add(3, , "###Description###")

   lstTags.ListItems.Add 11, , "NO COMMENT"
   lstTags.ListItems(11).ListSubItems.Add 1, , gBLOCK_NO_COMMENT
   lstTags.ListItems(11).ListSubItems.Add 2, , "No Comments"

   Set lstItem = lstTags.ListItems.Add(12, , "WEB SITE")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_WEBSITE)
   Call lstItem.ListSubItems.Add(2, , "Web Site")

   Set lstItem = lstTags.ListItems.Add(13, , "E-Mail")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_EMAIL)
   Call lstItem.ListSubItems.Add(2, , "E-Mail")

   Set lstItem = lstTags.ListItems.Add(14, , "Time")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_TIME)
   Call lstItem.ListSubItems.Add(2, , "Time")

   Set lstItem = lstTags.ListItems.Add(15, , "Telephone")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_TEL)
   Call lstItem.ListSubItems.Add(2, , "Telephone")

   Set lstItem = lstTags.ListItems.Add(16, , "Procedure Name")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_PROCEDURE_NAME)
   Call lstItem.ListSubItems.Add(2, , "Procedure Name")
   '   Call lstItem.ListSubItems.Add(3, , "###Name###")

   Set lstItem = lstTags.ListItems.Add(17, , "Module Name")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_MODULE_NAME)
   Call lstItem.ListSubItems.Add(2, , "Module Name")
   '   Call lstItem.ListSubItems.Add(3, , "###Name###")

   Set lstItem = lstTags.ListItems.Add(18, , "Module Filename")
   Call lstItem.ListSubItems.Add(1, , gBLOCK_MODULE_FILE)
   Call lstItem.ListSubItems.Add(2, , "Module Filename")

   lstTags.ListItems.Add 19, , "SEE ALSO"
   lstTags.ListItems(19).ListSubItems.Add 1, , gBLOCK_SEEALSO
   lstTags.ListItems(19).ListSubItems.Add 2, , "See Also"

   lstTags.ListItems.Add 20, , "TEXT"
   lstTags.ListItems(20).ListSubItems.Add 1, , gBLOCK_TEXT
   lstTags.ListItems(20).ListSubItems.Add 2, , "Text"

   lstTags.ListItems.Add 21, , "SCREENSHOT"
   lstTags.ListItems(21).ListSubItems.Add 1, , gBLOCK_SCREEENSHOT
   lstTags.ListItems(21).ListSubItems.Add 2, , "Screenshot"

   txtCompiler.Text = gHHCCompiler

   For Each comp In VBInstance.ActiveVBProject.VBComponents
      If Len(comp.Name) > 0 Then
         trvComps.Nodes.Add Null, tvwChild, comp.Name, comp.Name
         'If comp.Type = vbext_ct_ClassModule Then
         '   trvComps.Nodes(comp.Name).Checked = True
         'End If
      End If
   Next

End Sub

'Purpose: Save the settings
Private Sub SaveSettings()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : SaveSettings
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   gBLOCK_PROJECT = lstTags.ListItems(1).ListSubItems(1).Text
   gBLOCK_AUTHOR = lstTags.ListItems(2).ListSubItems(1).Text
   gBLOCK_DATE_CREATION = lstTags.ListItems(3).ListSubItems(1).Text
   gBLOCK_DATE_LAST_MOD = lstTags.ListItems(4).ListSubItems(1).Text
   gBLOCK_VERSION = lstTags.ListItems(5).ListSubItems(1).Text
   gBLOCK_PURPOSE = lstTags.ListItems(6).ListSubItems(1).Text
   gBLOCK_EXAMPLE = lstTags.ListItems(7).ListSubItems(1).Text
   gBLOCK_REMARKS = lstTags.ListItems(8).ListSubItems(1).Text
   gBLOCK_PARAMETER = lstTags.ListItems(9).ListSubItems(1).Text
   gBLOCK_CODE = lstTags.ListItems(10).ListSubItems(1).Text
   gBLOCK_NO_COMMENT = lstTags.ListItems(11).ListSubItems(1).Text

   gBLOCK_WEBSITE = lstTags.ListItems(12).ListSubItems(1).Text
   gBLOCK_EMAIL = lstTags.ListItems(13).ListSubItems(1).Text
   gBLOCK_TIME = lstTags.ListItems(14).ListSubItems(1).Text
   gBLOCK_TEL = lstTags.ListItems(15).ListSubItems(1).Text
   gBLOCK_PROCEDURE_NAME = lstTags.ListItems(16).ListSubItems(1).Text
   gBLOCK_MODULE_NAME = lstTags.ListItems(17).ListSubItems(1).Text
   gBLOCK_MODULE_FILE = lstTags.ListItems(18).ListSubItems(1).Text
   gBLOCK_SEEALSO = lstTags.ListItems(19).ListSubItems(1).Text
   gBLOCK_TEXT = lstTags.ListItems(20).ListSubItems(1).Text
   gBLOCK_SCREEENSHOT = lstTags.ListItems(21).ListSubItems(1).Text

   gHHCCompiler = txtCompiler.Text

   SaveSetting gsREG_APP, "Blocks", "Project", gBLOCK_PROJECT
   SaveSetting gsREG_APP, "Blocks", "Author", gBLOCK_AUTHOR
   SaveSetting gsREG_APP, "Blocks", "Date_Creation", gBLOCK_DATE_CREATION
   SaveSetting gsREG_APP, "Blocks", "Date_Last_Mod", gBLOCK_DATE_LAST_MOD
   SaveSetting gsREG_APP, "Blocks", "Version", gBLOCK_VERSION
   SaveSetting gsREG_APP, "Blocks", "Purpose", gBLOCK_PURPOSE
   SaveSetting gsREG_APP, "Blocks", "Example", gBLOCK_EXAMPLE
   SaveSetting gsREG_APP, "Blocks", "SeeAlso", gBLOCK_SEEALSO
   SaveSetting gsREG_APP, "Blocks", "Screenshot", gBLOCK_SCREEENSHOT
   SaveSetting gsREG_APP, "Blocks", "Code", gBLOCK_CODE
   SaveSetting gsREG_APP, "Blocks", "Remarks", gBLOCK_REMARKS
   SaveSetting gsREG_APP, "Blocks", "SeeAlso", gBLOCK_SEEALSO
   SaveSetting gsREG_APP, "Blocks", "Parameter", gBLOCK_PARAMETER
   SaveSetting gsREG_APP, "Blocks", "NoComment", gBLOCK_NO_COMMENT
   SaveSetting gsREG_APP, "Blocks", "Text", gBLOCK_TEXT

   SaveSetting gsREG_APP, "BLocks", "Web Site", gBLOCK_WEBSITE
   SaveSetting gsREG_APP, "BLocks", "E-Mail", gBLOCK_EMAIL
   SaveSetting gsREG_APP, "BLocks", "Time", gBLOCK_TIME
   SaveSetting gsREG_APP, "BLocks", "Telephone", gBLOCK_TEL
   SaveSetting gsREG_APP, "BLocks", "Procedure Name", gBLOCK_PROCEDURE_NAME
   SaveSetting gsREG_APP, "BLocks", "Module Name", gBLOCK_MODULE_NAME
   SaveSetting gsREG_APP, "BLocks", "Module Filename", gBLOCK_MODULE_FILE

   SaveSetting gsREG_APP, "Compiler", "Path", gHHCCompiler
End Sub

'Purpose: Put the cusor at the end of the textbox
Private Sub txtDOSOutputs_Change()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : frmAddIn
   ' * Module Filename  : VBDoc.frm
   ' * Procedure Name   : txtDOSOutputs_Change
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   txtDOSOutputs.SelStart = Len(txtDOSOutputs.Text)
End Sub

