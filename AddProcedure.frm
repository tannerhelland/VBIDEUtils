VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddNewProcedure 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Procedure Assistant"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   Icon            =   "AddProcedure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameTab 
      ForeColor       =   &H00800000&
      Height          =   4215
      Index           =   3
      Left            =   2160
      TabIndex        =   40
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkHidden 
         Alignment       =   1  'Right Justify
         Caption         =   "Hidden procedure"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Is this an hidden procedure?"
         Top             =   2040
         Width           =   1640
      End
      Begin VB.TextBox tbHelpContextID 
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Text            =   "Text1"
         ToolTipText     =   "don't forget his context ID :)"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox tbDescription 
         Height          =   1005
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Text            =   "Text1"
         ToolTipText     =   "Set the description for this new procedure..."
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Help Context ID :"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "&Description :"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   615
      Left            =   3000
      Picture         =   "AddProcedure.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Generate the code"
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   3960
      Picture         =   "AddProcedure.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Oops, forget it..."
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame frameTab 
      Height          =   4215
      Index           =   2
      Left            =   1560
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Move &Down"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "move this argument to the next place"
         Top             =   3480
         Width           =   735
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "Move &Up"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Move this argument to the previous place"
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton cmdRemoveArgument 
         Caption         =   "Remove"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Remove the selected argument"
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton cmdAddArgument 
         Caption         =   "Add"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Add this new argument"
         Top             =   2040
         Width           =   735
      End
      Begin VB.ListBox lstParameters 
         Height          =   2010
         Left            =   960
         TabIndex        =   19
         ToolTipText     =   "List of all created argument"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Frame Frame6 
         Caption         =   "Create a new argument"
         ForeColor       =   &H00800000&
         Height          =   1815
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   4335
         Begin VB.TextBox tbDefaultValue 
            Height          =   285
            Left            =   1200
            TabIndex        =   17
            Text            =   "Text1"
            ToolTipText     =   "Does it have a default value?"
            Top             =   1430
            Width           =   3015
         End
         Begin VB.CheckBox chkOptional 
            Caption         =   "&Optional"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2280
            TabIndex        =   16
            ToolTipText     =   "Is it optional?"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkByValue 
            Caption         =   "By &Value"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Do you pass it by value? "
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkArray 
            Caption         =   "&Array"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1200
            TabIndex        =   15
            ToolTipText     =   "Is this an array?"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.ComboBox cbDatatype 
            Height          =   315
            Left            =   960
            Sorted          =   -1  'True
            TabIndex        =   13
            Text            =   "Combo1"
            ToolTipText     =   "What kind of argument it is?"
            Top             =   680
            Width           =   3255
         End
         Begin VB.TextBox tbNameArgument 
            Height          =   285
            Left            =   960
            TabIndex        =   12
            Text            =   "Text1"
            ToolTipText     =   "Give the name of the new argument"
            Top             =   320
            Width           =   3255
         End
         Begin VB.Label lbDefaultValue 
            AutoSize        =   -1  'True
            Caption         =   "&Default Value :"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1470
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Data type :"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "&Name :"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame frameTab 
      ForeColor       =   &H00800000&
      Height          =   4215
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Width           =   4575
      Begin VB.CheckBox chkAddEnhancedError 
         Alignment       =   1  'Right Justify
         Caption         =   "Add Enhanced Error handler"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Add the enhanced error handler"
         Top             =   3315
         Width           =   2415
      End
      Begin VB.CheckBox chkAddStandardError 
         Alignment       =   1  'Right Justify
         Caption         =   "Add Standard Error handler"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Add the standard error handler"
         Top             =   2955
         Width           =   2415
      End
      Begin VB.CheckBox chkAddProcedureHeader 
         Alignment       =   1  'Right Justify
         Caption         =   "Add Procedure Header"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Add the procedure header"
         Top             =   2595
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.ComboBox cbType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "Combo1"
         ToolTipText     =   "Set the return type for the functions"
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Frame frameProperty 
         Caption         =   "Property"
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   1455
         Left            =   3000
         TabIndex        =   32
         ToolTipText     =   "In case of property, read/only?"
         Top             =   600
         Width           =   1455
         Begin VB.CheckBox chkWrite 
            Caption         =   "&Write"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkRead 
            Caption         =   "&Read"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Value           =   1  'Checked
            Width           =   1215
         End
      End
      Begin VB.Frame frameScope 
         Caption         =   "Scope"
         ForeColor       =   &H00000080&
         Height          =   1455
         Left            =   1560
         TabIndex        =   31
         ToolTipText     =   "Set the scope for your new procedure..."
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton optScope 
            Caption         =   "&Private"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optScope 
            Caption         =   "P&ublic"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optScope 
            Caption         =   "Fr&iend"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Frame frameType 
         Caption         =   "Type"
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Is this a procedure, function or property?"
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton optType 
            Caption         =   "&Property"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton optType 
            Caption         =   "&Function"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optType 
            Caption         =   "&Sub"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox tbNameProcedure 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "Set the name of the procedure/function/property"
         Top             =   210
         Width           =   3615
      End
      Begin VB.Label lbType 
         AutoSize        =   -1  'True
         Caption         =   "Return type :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2205
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "&Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   4695
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8281
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            Object.ToolTipText     =   "General infos about the procedure creation"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Arguments"
            Key             =   "Arguments"
            Object.ToolTipText     =   "Set the arguments for the procedure"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Description"
            Key             =   "Description"
            Object.ToolTipText     =   "Set the description and other attributes of the procedure"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddNewProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 09/11/1999
' * Time             : 10:26
' * Module Name      : frmAddNewProcedure
' * Module Filename  : AddProcedure.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Private mnCurFrame      As Integer

Private Sub chkAddEnhancedError_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 12:40
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : chkAddEnhancedError_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If (chkAddStandardError.Value = vbChecked) And (chkAddEnhancedError.Value = vbChecked) Then chkAddStandardError.Value = vbUnchecked

End Sub

Private Sub chkAddstandardError_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 12:41
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : chkAddstandardError_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If (chkAddStandardError.Value = vbChecked) And (chkAddEnhancedError.Value = vbChecked) Then chkAddEnhancedError.Value = vbUnchecked

End Sub

Private Sub cmdAddArgument_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 14:42
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : cmdAddArgument_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *  Add the parameter to the list
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdAddArgument_Click

   Dim sTmp             As String

   sTmp = ""

   If Trim$(tbNameArgument.Text) = "" Then Exit Sub

   If chkOptional.Value = vbChecked Then sTmp = sTmp & "Optional "
   If chkByValue.Value = vbChecked Then sTmp = sTmp & "ByVal "

   sTmp = sTmp & Trim$(tbNameArgument.Text)

   If chkArray.Value = vbChecked Then
      sTmp = sTmp & "() "
   Else
      sTmp = sTmp & " "
   End If

   If Trim$(cbDatatype.Text) <> "" Then sTmp = sTmp & "As " & Trim$(cbDatatype.Text) & " "

   If chkOptional.Value = vbChecked And Trim$(tbDefaultValue.Text) <> "" Then
      Select Case UCase$(Trim$(cbDatatype.Text))
         Case "STRING":
            sTmp = sTmp & "= """ & Trim$(tbDefaultValue.Text) & """ "
         Case Else:
            sTmp = sTmp & "= " & Trim$(tbDefaultValue.Text) & " "
      End Select
   End If

   lstParameters.AddItem sTmp

EXIT_cmdAddArgument_Click:
   tbNameArgument.Text = ""

   chkByValue.Value = vbUnchecked
   chkOptional.Value = vbUnchecked

   cbDatatype.Text = "Variant"

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdAddArgument_Click:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdAddArgument_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdAddArgument_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_cmdAddArgument_Click

End Sub

Private Sub cmdCancel_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 14:42
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : cmdCancel_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call UnloadEffect(Me)
   Unload Me

End Sub

Private Sub cmdMoveDown_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 14:51
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : cmdMoveDown_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim nI               As Long

   On Error Resume Next

   If (lstParameters.ListIndex = -1) Or (lstParameters.ListIndex = lstParameters.ListCount - 1) Then Exit Sub

   nI = lstParameters.ListIndex
   sTmp = lstParameters.List(nI)

   lstParameters.RemoveItem nI
   lstParameters.AddItem sTmp, nI + 1
   lstParameters.ListIndex = nI + 1

End Sub

Private Sub cmdMoveUp_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 14:50
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : cmdMoveUp_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim nI               As Long

   On Error Resume Next

   If lstParameters.ListIndex < 1 Then Exit Sub

   nI = lstParameters.ListIndex
   sTmp = lstParameters.List(nI)

   lstParameters.RemoveItem nI
   lstParameters.AddItem sTmp, nI - 1
   lstParameters.ListIndex = nI - 1

End Sub

Private Sub cmdOK_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 12:00
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : cmdOk_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' * Generate the created procedure
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdOk_Click

   Dim nI               As Long
   Dim sTmp             As String

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim nProcKind        As vbext_ProcKind

   Dim nNbrPass         As Integer

   If Trim$(tbNameProcedure.Text) = "" Then Exit Sub

   On Error Resume Next

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

   nNbrPass = 0

   If optType(2).Value Then
      If chkRead.Value = vbChecked Then nNbrPass = nNbrPass + 1
      If chkWrite.Value = vbChecked Then nNbrPass = nNbrPass + 1
   End If

   If nNbrPass = 0 Then
      If optType(2).Value Then chkRead.Value = vbChecked
      nNbrPass = 1
   End If

   sTmp = vbCrLf
Next_Pass:

   ' *** Select scope
   If optScope(0).Value Then sTmp = sTmp & "Private "
   If optScope(1).Value Then sTmp = sTmp & "Public "
   If optScope(2).Value Then sTmp = sTmp & "Friend "

   ' *** Select type
   If optType(0).Value Then
      sTmp = sTmp & "Sub "
      nProcKind = vbext_pk_Proc
   End If
   If optType(1).Value Then
      sTmp = sTmp & "Function "
      nProcKind = vbext_pk_Proc
   End If
   If optType(2).Value Then
      sTmp = sTmp & "Property "

      If chkRead.Value = vbChecked Then
         nProcKind = vbext_pk_Get
         sTmp = sTmp & "Get "
         chkRead.Value = vbUnchecked

         ' *** Clear the parameters
         lstParameters.Clear

      ElseIf chkWrite.Value = vbChecked Then
         nProcKind = vbext_pk_Let
         sTmp = sTmp & "Let "
         chkWrite.Value = vbUnchecked

         ' *** Add the parameter
         lstParameters.Clear
         lstParameters.AddItem "NewValue As " & cbType.Text
      End If

   End If

   ' *** Add the name
   sTmp = sTmp & tbNameProcedure.Text & "("

   ' *** Add the parameters
   For nI = 0 To lstParameters.ListCount - 1
      sTmp = sTmp & lstParameters.List(nI) & ", "
   Next
   If right(sTmp, 2) = ", " Then sTmp = left$(sTmp, Len(sTmp) - 2)
   sTmp = sTmp & ") "

   ' *** Add the return
   If cbType.Text <> "" Then
      If optType(2).Value Then
         If nProcKind <> vbext_pk_Let Then
            sTmp = sTmp & "As " & cbType.Text
         End If
      ElseIf optType(1).Value Then
         sTmp = sTmp & "As " & cbType.Text
      End If

   End If

   sTmp = sTmp & vbCrLf
   sTmp = sTmp & vbCrLf

   ' *** Add the End Sub
   If optType(0).Value Then sTmp = sTmp & "End Sub "
   If optType(1).Value Then sTmp = sTmp & "End Function "
   If optType(2).Value Then sTmp = sTmp & "End Property "

   ' *** Add this procedure
   cpCodePane.CodeModule.AddFromString sTmp
   nI = cpCodePane.CodeModule.ProcBodyLine(tbNameProcedure.Text, nProcKind)
   cpCodePane.SetSelection nI, 1, nI, 1
   cpCodePane.CodeModule.members(1).Description = tbDescription.Text
   cpCodePane.CodeModule.members(1).HelpContextID = tbHelpContextID.Text
   cpCodePane.CodeModule.members(1).Hidden = IIf(chkHidden.Value = vbChecked, True, False)

   ' *** Add the procedure Header
   If chkAddProcedureHeader.Value = vbChecked Then Call InsertProcedureHeader

   ' *** Add the standard error handler
   If chkAddStandardError.Value = vbChecked Then Call InsertProcedureError

   ' *** Add the enhanced error handler
   If chkAddEnhancedError.Value = vbChecked Then Call InsertEnhancedProcedureError

   If nNbrPass > 1 Then
      nNbrPass = nNbrPass - 1
      sTmp = ""
      GoTo Next_Pass
   End If

EXIT_cmdOk_Click:
   Call InitAll

   Call cmdCancel_Click

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdOk_Click:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdOk_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdOk_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_cmdOk_Click

End Sub

Private Sub cmdRemoveArgument_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 14:48
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : cmdRemoveArgument_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *  Remove the selected parameter
   ' *
   ' **********************************************************************

   On Error Resume Next

   If lstParameters.ListIndex = -1 Then Exit Sub

   lstParameters.RemoveItem lstParameters.ListIndex

End Sub

Private Sub optType_Click(index As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 12:01
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : optType_Click
   ' * Parameters       :
   ' *                    Index As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Select Case index
      Case 0:
         ' *** Procedure
         frameScope.Enabled = True
         optScope(2).Enabled = True
         frameProperty.Enabled = False

         cbType.Text = ""
         lbType.Enabled = False
         cbType.Enabled = False

      Case 1:
         ' *** Function
         frameScope.Enabled = True
         optScope(2).Enabled = True
         frameProperty.Enabled = False

         lbType.Enabled = True
         cbType.Enabled = True

      Case 2:
         ' *** Property, disable friend and enable property
         If optScope(2).Value = True Then optScope(1).Value = True
         optScope(2).Enabled = False
         frameScope.Enabled = True
         frameProperty.Enabled = True

         lbType.Enabled = True
         cbType.Enabled = True

   End Select

End Sub

Private Sub tabstrip_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : tabstrip_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   If TabStrip.SelectedItem.index = mnCurFrame Then Exit Sub       ' No need to change frame.

   ' *** Otherwise, hide old frame, show new.
   frametab(TabStrip.SelectedItem.index).left = 240
   frametab(TabStrip.SelectedItem.index).top = 480
   frametab(TabStrip.SelectedItem.index).Visible = True
   frametab(mnCurFrame).Visible = False
   frametab(TabStrip.SelectedItem.index).ZOrder

   ' *** Set mnCurFrame to new value.
   mnCurFrame = TabStrip.SelectedItem.index

End Sub

Private Sub InitAll()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:48
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : InitAll
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Long

   ' *** Init all textbox, checkbox...
   tbNameProcedure.Text = ""
   tbNameArgument.Text = ""
   tbDefaultValue.Text = ""
   tbDescription.Text = ""
   tbHelpContextID.Text = "0"

   ' *** Init combo box
   cbType.Clear
   cbType.AddItem "Boolean"
   cbType.AddItem "Byte"
   cbType.AddItem "Collection"
   cbType.AddItem "Currency"
   cbType.AddItem "Date"
   cbType.AddItem "Double"
   cbType.AddItem "Integer"
   cbType.AddItem "Long"
   cbType.AddItem "Object"
   cbType.AddItem "Single"
   cbType.AddItem "String"
   cbType.AddItem "Variant"
   cbType.Text = "Variant"

   cbDatatype.Clear
   For nI = 0 To cbType.ListCount - 1
      cbDatatype.AddItem cbType.List(nI)
   Next
   cbDatatype.Text = "Variant"

   lstParameters.Clear

End Sub

Private Sub chkOptional_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 11:56
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : chkOptional_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If chkOptional.Value = vbChecked Then
      tbDefaultValue.Enabled = True
      lbDefaultValue.Enabled = True
   Else
      tbDefaultValue.Enabled = False
      lbDefaultValue.Enabled = False
   End If

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   On Error Resume Next

   Call InitTooltips

   Call InitAll

   frametab(2).Visible = False
   Set TabStrip.SelectedItem = TabStrip.Tabs("General")

   ' *** Get the active project
   Set prjProject = VBInstance.ActiveVBProject

   ' *** If we couldn't get it, quit
   If prjProject Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current project", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the project")
      Call UnloadEffect(Me)
      Unload Me
      Exit Sub
   End If

   ' *** Try to find the active code pane
   Set cpCodePane = VBInstance.ActiveCodePane

   ' *** If we couldn't get it, quit
   If cpCodePane Is Nothing Then
      'Call MsgBoxTop(Me.hwnd, "Could not identify current module", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the module")
      Call UnloadEffect(Me)
      Unload Me
      Exit Sub
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : Form_Unload
   ' * Parameters       :
   ' *                    Cancel As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Set clsTooltips = Nothing

End Sub

Private Sub tbDefaultValue_GotFocus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 16:34
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : tbDefaultValue_GotFocus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   SelectAll

End Sub

Private Sub tbDescription_GotFocus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 16:33
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : tbDescription_GotFocus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call SelectAll

End Sub

Private Sub tbHelpContextID_GotFocus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 16:33
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : tbHelpContextID_GotFocus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call SelectAll

End Sub

Private Sub tbHelpContextID_KeyPress(KeyAscii As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 15:48
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : tbHelpContextID_KeyPress
   ' * Parameters       :
   ' *                    KeyAscii As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0

End Sub

Private Sub tbNameArgument_GotFocus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 16:34
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : tbNameArgument_GotFocus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   SelectAll

End Sub

Private Sub tbNameProcedure_GotFocus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/11/1999
   ' * Time             : 16:34
   ' * Module Name      : frmAddNewProcedure
   ' * Module Filename  : AddProcedure.frm
   ' * Procedure Name   : tbNameProcedure_GotFocus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   SelectAll

End Sub

Private Sub InitTooltips(Optional Flags As TooltipFlagConstants)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/11/1999
   ' * Time             : 11:15
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
   ' * Procedure Name   : InitTooltips
   ' * Parameters       :
   ' *                    Optional Flags As TooltipFlagConstants
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer

   On Error Resume Next

   With clsTooltips
      .Create Me, ttfBalloon
      .MaxTipWidth = 70
      .Icon = itInfoIcon
      .Title = "VBIDEUtils"
      For nI = 0 To Me.Controls.Count
         If Trim$(Me.Controls(nI).ToolTipText) <> "" Then
            If err = 0 Then
               .AddTool Me.Controls(nI), tfTransparent, Me.Controls(nI).ToolTipText
               Me.Controls(nI).ToolTipText = ""
            Else
               err.Clear
            End If
         End If
      Next
   End With

End Sub
