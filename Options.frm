VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VBIDEUtils Options"
   ClientHeight    =   5490
   ClientLeft      =   2175
   ClientTop       =   1890
   ClientWidth     =   9120
   ControlBox      =   0   'False
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frametab 
      Caption         =   "Indenting Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4095
      Index           =   2
      Left            =   4320
      TabIndex        =   31
      Top             =   3120
      Visible         =   0   'False
      Width           =   8655
      Begin VB.VScrollBar vSpace 
         Height          =   320
         Left            =   1680
         Max             =   0
         Min             =   9
         TabIndex        =   11
         Top             =   3670
         Width           =   255
      End
      Begin VB.CheckBox cbIndentDim 
         Caption         =   "Align all dim/global variables..."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         ToolTipText     =   "Align all variable on the 'As'"
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.TextBox tbSpaceNum 
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "3"
         ToolTipText     =   "Set how much space for 1 tab.... 3 is good :)"
         Top             =   3690
         Width           =   255
      End
      Begin RichTextLib.RichTextBox rtfColor 
         Height          =   495
         Left            =   2040
         TabIndex        =   36
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         _Version        =   393217
         TextRTF         =   $"Options.frx":0442
      End
      Begin RichTextLib.RichTextBox rtfCodeExample 
         Height          =   2895
         Left            =   120
         TabIndex        =   35
         ToolTipText     =   "Sample of indenting"
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5106
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"Options.frx":04C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox cbIndentCmt 
         Caption         =   "Indent comments to align with code"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         ToolTipText     =   "Align all the comments to the code..."
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox cbIndentCase 
         Caption         =   "Indent within ""Select Case"" lines"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         ToolTipText     =   "Indent all the code within a Select Case"
         Top             =   3360
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.OptionButton obUseTabs 
         Caption         =   "Tabs"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         ToolTipText     =   "Do you prefer to use tabs?"
         Top             =   3360
         Width           =   1035
      End
      Begin VB.OptionButton obUseSpaces 
         Caption         =   "Spaces:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         ToolTipText     =   "Indent using spaces"
         Top             =   3720
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.CheckBox cbIndentProc 
         Caption         =   "Indent everything within a procedure"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         ToolTipText     =   "Indent all within a procedure/function"
         Top             =   3360
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Use:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   3360
         Width           =   330
      End
   End
   Begin VB.Frame frametab 
      Caption         =   "Comment template"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4095
      Index           =   3
      Left            =   2760
      TabIndex        =   33
      Top             =   3120
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cmdDeleteTemplate 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   7560
         TabIndex        =   23
         ToolTipText     =   "Delete a template"
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdSaveTemplate 
         Caption         =   "&Save"
         Height          =   315
         Left            =   6720
         TabIndex        =   22
         ToolTipText     =   "Save a template"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbCommentString 
         Height          =   315
         Left            =   2280
         MaxLength       =   1
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "*"
         ToolTipText     =   "Set the string to be used in the comments"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cbTemplate 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Do you want to load/save a template? Set the name here"
         Top             =   680
         Width           =   5055
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "<<"
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         ToolTipText     =   "Remve all fields from the template"
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<"
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         ToolTipText     =   "Remove a field from the template"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         ToolTipText     =   "Add a field in the template"
         Top             =   1680
         Width           =   975
      End
      Begin VB.ListBox lstDestination 
         Height          =   2790
         Left            =   5160
         TabIndex        =   19
         ToolTipText     =   "Fields used in the template"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ListBox lstSource 
         Height          =   2790
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Fields not used in the template"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "String used in the comments"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Template Name"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame frametab 
      Caption         =   "User Comment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4095
      Index           =   4
      Left            =   7560
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox tbUserComment 
         Height          =   3735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   41
         ToolTipText     =   "Set the user comment"
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.Frame frametab 
      Caption         =   "Error handler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4095
      Index           =   5
      Left            =   -1800
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame Frame1 
         Caption         =   "Add the following function"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   120
         TabIndex        =   49
         Top             =   3360
         Width           =   8415
         Begin VB.CommandButton cmdAddResumeExit 
            Caption         =   "Resume Exit"
            Height          =   255
            Left            =   5880
            TabIndex        =   47
            ToolTipText     =   "Add a resume exit in the error handler"
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdAddProjectName 
            Caption         =   "Project Name"
            Height          =   255
            Left            =   3960
            TabIndex        =   46
            ToolTipText     =   "Add the project name in the error handler"
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton clmdAddModuleName 
            Caption         =   "Module Name"
            Height          =   255
            Left            =   2040
            TabIndex        =   45
            ToolTipText     =   "Add the module name in the error handler"
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdAddProcedureName 
            Caption         =   "Procedure Name"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            ToolTipText     =   "Add the procedure name in the error handler"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox tbCustomizeErrorHandler 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Text            =   "Options.frx":0546
         ToolTipText     =   "Set the standard error handler"
         Top             =   480
         Width           =   8415
      End
      Begin VB.Label Label13 
         Caption         =   "Customize Error Handler"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame frametab 
      Caption         =   "General configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4095
      Index           =   1
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Width           =   8655
      Begin VB.TextBox tbInline 
         Height          =   315
         Left            =   1680
         MaxLength       =   35
         TabIndex        =   5
         ToolTipText     =   "Set a particular inline comment..."
         Top             =   2160
         Width           =   5655
      End
      Begin VB.TextBox tbDevelopper 
         Height          =   315
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   1
         ToolTipText     =   "Set the developer name used in the comments"
         Top             =   240
         Width           =   5655
      End
      Begin VB.TextBox tbWebSite 
         Height          =   315
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "Set the web site name used in the comments"
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox tbEMail 
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Set the e-mail name used in the comments"
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox tbTelephone 
         Height          =   315
         Left            =   1680
         MaxLength       =   35
         TabIndex        =   4
         ToolTipText     =   "Set the phone number name used in the comments"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox tbComment 
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         ToolTipText     =   "Set your prefered comment"
         Top             =   2640
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.Label Label11 
         Caption         =   "Web site"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Inline comment"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2205
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Prefered comment"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Developer name"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "E-Mail"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Telephone"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1725
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   7080
      Picture         =   "Options.frx":054C
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Ok, I accept all the options"
      Top             =   4800
      Width           =   855
   End
   Begin MSComctlLib.TabStrip tabstrip 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8070
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General Options"
            Key             =   "GENERAL"
            Object.ToolTipText     =   "General options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Indenting Options"
            Key             =   "INDENTING"
            Object.ToolTipText     =   "Indenting options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Comment Template"
            Key             =   "COMMENT"
            Object.ToolTipText     =   "Comment template"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&User Comment"
            Key             =   "USERCOMMENT"
            Object.ToolTipText     =   "Use user comment"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Error handler"
            Key             =   "ERROR"
            Object.ToolTipText     =   "Error handler"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   8040
      Picture         =   "Options.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Oops, forget it..."
      Top             =   4800
      Width           =   855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 2/10/98
' * Time             : 14:36
' * Module Name      : frmOptions
' * Module Filename  : Options.frm
' **********************************************************************
' * Comments         : Set the options for the VBIDEUtils addins
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

' *** Dimension an array to hold the code for the example procedure
Dim asCodeLines(1 To 16) As String

Private mnCurFrame      As Integer

Private Sub clmdAddModuleName_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : clmdAddModuleName_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tbCustomizeErrorHandler.SetFocus
   SendKeys "{{}ModuleName{}}"

End Sub

Private Sub cmdAddProcedureName_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdAddProcedureName_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tbCustomizeErrorHandler.SetFocus
   SendKeys "{{}ProcedureName{}}"

End Sub

Private Sub cmdAddProjectName_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdAddProjectName_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tbCustomizeErrorHandler.SetFocus
   SendKeys "{{}ProjectName{}}"

End Sub

Private Sub cmdAddResumeExit_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdAddResumeExit_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tbCustomizeErrorHandler.SetFocus
   SendKeys "Resume EXIT_{{}ProcedureName{}}:"

End Sub

Private Sub Form_Activate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : Form_Activate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   ' *** Define the example procedure code lines
   asCodeLines(1) = "' *** Example Procedure"
   asCodeLines(2) = "Public Sub SetRedraw(frm As Form, bRedraw As Boolean)"
   asCodeLines(3) = ""
   asCodeLines(4) = "' *** VBIDEUtils"
   asCodeLines(5) = "MsgBox ""Copyright © 2001 Me"""
   asCodeLines(6) = ""
   asCodeLines(7) = "If bDoYouKnowPrintPreviewOCX = False Then"
   asCodeLines(8) = "' *** Visit http://www.ppreview.net"
   asCodeLines(9) = "Select Case X"
   asCodeLines(10) = "Case 1"
   asCodeLines(11) = "' *** If you have any comments or suggestions,"
   asCodeLines(12) = "MsgBox ""Contact removed"""
   asCodeLines(13) = "End Select"
   asCodeLines(14) = "End If"
   asCodeLines(15) = ""
   asCodeLines(16) = "End Sub"

   ' *** Read the options from the registry
   tbSpaceNum.Text = GetSetting(gsREG_APP, "Indent", "IndentSpaces", "3")
   obUseTabs.Value = (GetSetting(gsREG_APP, "Indent", "UseTabs", "N") = "Y")
   cbIndentProc.Value = IIf(GetSetting(gsREG_APP, "Indent", "IndentProc", "N") = "Y", 1, 0)
   cbIndentCmt.Value = IIf(GetSetting(gsREG_APP, "Indent", "IndentCmt", "Y") = "Y", 1, 0)
   cbIndentCase.Value = IIf(GetSetting(gsREG_APP, "Indent", "IndentCase", "N") = "Y", 1, 0)
   cbIndentDim.Value = IIf(GetSetting(gsREG_APP, "Indent", "IndentVariable", "N") = "Y", 1, 0)

   ' *** Update the code box
   UpdateCodeListBox

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Set the textbox

   On Error Resume Next

   Call InitTooltips

   mnCurFrame = -1
   tbDevelopper.Text = gsDevelopper
   tbWebSite.Text = gsWebSite
   tbEMail.Text = gsEmail
   tbTelephone.Text = gsTelephone
   tbInline.Text = gsInline
   tbComment.Text = gsComment
   tbUserComment.Text = gsUserComment
   tbCommentString.Text = gsCommentString
   tbCustomizeErrorHandler.Text = gsErrorHandler

   ' *** Fill the Procedure Header possibility
   Fill_ProcedureHeader

   ' *** Fill the combo with template
   Fill_ComboTemplate

   vSpace.Value = tbSpaceNum.Text

   Me.Show
   DoEvents

End Sub

Private Sub obUseTabs_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:12
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : obUseTabs_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Handles clicking on the Use Tabs option button
   ' *
   ' *
   ' **********************************************************************

   UpdateCodeListBox

End Sub

Private Sub obUseSpaces_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:13
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : obUseSpaces_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Handles clicking on the Use Spaces option button
   ' *
   ' *
   ' **********************************************************************

   ' *** Default to use 3 spaces if nothing there already
   If tbSpaceNum = "" Then tbSpaceNum = "3"

   UpdateCodeListBox

End Sub

Private Sub tabstrip_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
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

Private Sub tbSpaceNum_Change()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:13
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : tbSpaceNum_Change
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Handles changes in the Space number edit box, validating the entry
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   'Set the option button to use spaces
   obUseSpaces = True

   'Validate the item in the edit box
   If tbSpaceNum = "" Then
      obUseTabs = True
   ElseIf Not IsNumeric(tbSpaceNum) Then
      tbSpaceNum = ""
      obUseTabs = True
   Else
      tbSpaceNum = CStr(Abs(Int(CDbl(tbSpaceNum))))
   End If

   UpdateCodeListBox

End Sub

Private Sub cbIndentProc_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:13
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cbIndentProc_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Handles checking the "Indent everything within procedure" check box
   ' *
   ' *
   ' **********************************************************************

   UpdateCodeListBox

End Sub

Private Sub cbIndentCmt_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:14
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cbIndentCmt_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Handles checking the "Indent comments to align with code" check box
   ' *
   ' *
   ' **********************************************************************

   UpdateCodeListBox
End Sub

Private Sub cbIndentCase_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:14
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cbIndentCase_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Handles checking the "Indent within Select Case blocks" check box
   ' *
   ' *
   ' **********************************************************************

   UpdateCodeListBox
End Sub

Private Sub cmdOK_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:14
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdOK_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Handles clicking the OK button.  Stores the current options in the registry
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   ' *** Store the current options in the registy
   gsDevelopper = tbDevelopper.Text
   gsWebSite = tbWebSite.Text
   gsEmail = tbEMail.Text
   gsTelephone = tbTelephone.Text
   gsInline = tbInline.Text
   gsComment = tbComment.Text
   gsUserComment = tbUserComment.Text
   gsCommentString = tbCommentString.Text
   gsErrorHandler = IndentCode(tbCustomizeErrorHandler.Text)

   SaveSetting gsREG_APP, "General", "Developper", gsDevelopper
   SaveSetting gsREG_APP, "General", "WebSite", gsWebSite
   SaveSetting gsREG_APP, "General", "Email", gsEmail
   SaveSetting gsREG_APP, "General", "Telephone", gsTelephone
   SaveSetting gsREG_APP, "General", "Inline", gsInline
   SaveSetting gsREG_APP, "General", "Comment", gsComment
   SaveSetting gsREG_APP, "General", "UserComment", gsUserComment
   SaveSetting gsREG_APP, "General", "CommentString", gsCommentString

   SaveSetting gsREG_APP, "Error", "CustomizedErrorHandler", gsErrorHandler

   SaveSetting gsREG_APP, "Indent", "UseTabs", IIf(obUseTabs, "Y", "N")
   SaveSetting gsREG_APP, "Indent", "IndentSpaces", tbSpaceNum.Text
   SaveSetting gsREG_APP, "Indent", "IndentProc", IIf(cbIndentProc = 1, "Y", "N")
   SaveSetting gsREG_APP, "Indent", "IndentCmt", IIf(cbIndentCmt = 1, "Y", "N")
   SaveSetting gsREG_APP, "Indent", "IndentCase", IIf(cbIndentCase = 1, "Y", "N")
   SaveSetting gsREG_APP, "Indent", "IndentVariable", IIf(cbIndentDim = 1, "Y", "N")

   If Trim$(cbTemplate.Text) <> vbNullString Then
      gsTemplate = "Template_" & Trim$(cbTemplate.Text)
   Else
      gsTemplate = "Template_Default"
   End If
   SaveSetting gsREG_APP, "General", "DefaultTemplate", gsTemplate

   Call UnloadEffect(Me)
   Unload Me

End Sub

Private Sub cmdCancel_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:16
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdCancel_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Handles clicking the Cancel button.  Just unloads the form
   ' *
   ' *
   ' **********************************************************************

   Call UnloadEffect(Me)
   Unload Me

End Sub

Private Sub UpdateCodeListBox()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 14:16
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : UpdateCodeListBox
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Common routine to work out the indenting in
   ' * the example code and put it in the list box.
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer
   Dim sPad             As String
   Dim sTmp             As String

   On Error Resume Next

   ' *** Remove any spacing that might be there already
   For nI = 1 To 16
      asCodeLines(nI) = Trim$(asCodeLines(nI))
   Next

   ' *** Get how much to pad by
   If obUseTabs Then
      sPad = Space(4)
   Else
      If tbSpaceNum = "" Then
         sPad = ""
      Else
         sPad = Space(CInt(tbSpaceNum))
      End If
   End If

   ' *** Always indent the If
   For nI = 8 To 13
      asCodeLines(nI) = sPad & asCodeLines(nI)
   Next

   ' *** Always indent the Select
   asCodeLines(11) = sPad & asCodeLines(11)
   asCodeLines(12) = sPad & asCodeLines(12)

   ' *** Do we extra indent the Select Case?
   If cbIndentCase = 1 Then
      asCodeLines(10) = sPad & asCodeLines(10)
      asCodeLines(11) = sPad & asCodeLines(11)
      asCodeLines(12) = sPad & asCodeLines(12)
   End If

   ' *** Do we indent comments?
   If cbIndentCmt <> 1 Then
      asCodeLines(4) = Trim$(asCodeLines(4))
      asCodeLines(8) = Trim$(asCodeLines(8))
      asCodeLines(11) = Trim$(asCodeLines(11))
   End If

   ' *** Do we indent the entire procedure?
   If cbIndentProc = 1 Then
      For nI = 3 To 14
         asCodeLines(nI) = sPad & asCodeLines(nI)
      Next
   End If

   ' *** Put the procedure code in the list box.
   sTmp = ""
   For nI = 1 To 16
      sTmp = sTmp & asCodeLines(nI) & vbCrLf
   Next

   rtfColor.SelStart = 0
   rtfColor.SelLength = Len(rtfColor.TextRTF)
   rtfColor.SelColor = vbBlack
   rtfColor.TextRTF = ""
   rtfColor.SelStart = 0
   rtfColor.Text = sTmp

   ' *** Colorize the code
   Call InitColorize
   rtfCodeExample.TextRTF = ColorizeVBCode(rtfColor.Text)

End Sub

Private Sub Fill_ComboTemplate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : Fill_ComboTemplate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Get all the saved templates

   Dim allSettings()    As String
   Dim nI               As Integer
   Dim sTmp             As String

   On Error Resume Next

   cbTemplate.Clear

   Call GetAllSettings(gsREG_APP, "Template", allSettings)
   If (IsEmpty(allSettings) = False) Then
      For nI = LBound(allSettings, 1) To UBound(allSettings, 1)
         sTmp = CStr(allSettings(nI))
         sTmp = right$(sTmp, Len(sTmp) - Len("Template_"))
         cbTemplate.AddItem sTmp
      Next
   End If

   gsTemplate = GetSetting(gsREG_APP, "General", "DefaultTemplate", "Template_Default")
   sTmp = right$(gsTemplate, Len(gsTemplate) - Len("Template_"))

   ' *** Check if at least one template
   If nI = 0 Then
      ' *** Add a default Template
      Add_DefaultTemplate

      If (CBLFind(cbTemplate, sTmp, 0) = -1) Then cbTemplate.AddItem gsTemplate

   End If

   ' *** Select the template
   cbTemplate.ListIndex = CBLFind(cbTemplate, sTmp, 0)

End Sub

Private Sub Fill_ProcedureHeader()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : Fill_ProcedureHeader
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' * Fill the Procedure Header possibility
   ' *
   ' **********************************************************************

   lstSource.Clear

   lstSource.AddItem "Author"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentProgrammerName

   lstSource.AddItem "Programmer Web Site"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentWebSite

   lstSource.AddItem "Programmer E-Mail"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentEMail

   lstSource.AddItem "Programmer Tel"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentTel

   lstSource.AddItem "Date" 'Format(Now, "dd/mm/yyyy")
   lstSource.ItemData(lstSource.NewIndex) = twaCommentDate

   lstSource.AddItem "Time" 'Format(Now, "hh:mm")
   lstSource.ItemData(lstSource.NewIndex) = twaCommentTime

   lstSource.AddItem "Project name" 'gVBInstance.Name
   lstSource.ItemData(lstSource.NewIndex) = twaCommentProjectName

   lstSource.AddItem "Module name" 'gVBInstance.ActiveCodePane.CodeModule.Parent.Name
   lstSource.ItemData(lstSource.NewIndex) = twaCommentModuleName

   lstSource.AddItem "Module filename" 'gVBInstance.ActiveCodePane.CodeModule.Parent.FileNames
   lstSource.ItemData(lstSource.NewIndex) = twaCommentModuleFileName

   lstSource.AddItem "Procedure name" 'gVBInstance.ActiveCodePane.CodeModule.ProcOfLine(nStartLine, vbext_pk_Proc)
   lstSource.ItemData(lstSource.NewIndex) = twaCommentProcedureName

   lstSource.AddItem "Purpose"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentPurpose

   lstSource.AddItem "Procedure Parameters" 'gVBInstance.ActiveCodePane.CodeModule.Lines(gVBInstance.ActiveCodePane.CodeModule.ProcBodyLine(gVBInstance.ActiveCodePane.CodeModule.ProcOfLine(nStartLine, vbext_pk_Proc), vbext_pk_Proc), 1)
   lstSource.ItemData(lstSource.NewIndex) = twaCommentProcedureParameters

   lstSource.AddItem "Prefered comment"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentPrefered

   lstSource.AddItem "Screenshot"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentScreenshot

   lstSource.AddItem "Sample"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentSample

   lstSource.AddItem "See Also"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentSeeAlso

   lstSource.AddItem "History"
   lstSource.ItemData(lstSource.NewIndex) = twaCommentHistory

End Sub

Private Sub ReadTemplate(sTemplate As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/21/2002
   ' * Time             : 20:49
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : ReadTemplate
   ' * Purpose          :
   ' * Parameters       :
   ' *                    sTemplate As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' *
   ' * Screenshot       :
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Read a template for particular items

   Dim nI               As Integer
   Dim nItem            As Long

   If left(sTemplate, Len("Template_")) <> "Template_" Then
      sTemplate = "Template_" & sTemplate
   End If

   lstDestination.Clear

   ' *** Get all infos
   Fill_ProcedureHeader

   nI = 1
   nItem = CInt(GetSetting(gsREG_APP, sTemplate, CStr(nI), "0"))
   Do While nItem <> 0
      ' *** Select the right one and add it
      lstSource.ListIndex = CBLFindCode(lstSource, nItem)
      cmdAdd_Click

      ' *** Get next item
      nI = nI + 1
      nItem = CInt(GetSetting(gsREG_APP, sTemplate, CStr(nI), "0"))
   Loop

End Sub

Private Sub SaveTemplate(sTemplate As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : SaveTemplate
   ' * Parameters       :
   ' *                    sTemplate As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Save a template for particular items

   Dim nI               As Integer

   sTemplate = "Template_" & sTemplate

   ' *** Delete previous entries
   On Error Resume Next

   DeleteSection gsREG_APP, sTemplate

   For nI = 1 To lstDestination.ListCount
      Call SaveSetting(gsREG_APP, sTemplate, CStr(nI), lstDestination.ItemData(nI - 1))
   Next

   SaveSetting gsREG_APP, "Template", sTemplate, sTemplate

   ' *** Save as default template
   SaveSetting gsREG_APP, "General", "DefaultTemplate", sTemplate

   gsTemplate = sTemplate

End Sub

Private Sub cbTemplate_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cbTemplate_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Read the current template

   ReadTemplate cbTemplate.Text

End Sub

Private Sub cmdAdd_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdAdd_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Add One line

   If (lstSource.ListIndex = -1) Then Exit Sub

   lstDestination.AddItem lstSource.List(lstSource.ListIndex)
   lstDestination.ItemData(lstDestination.NewIndex) = lstSource.ItemData(lstSource.ListIndex)

   lstSource.RemoveItem lstSource.ListIndex

End Sub

Private Sub cmdDeleteTemplate_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdDeleteTemplate_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Delete this template

   On Error Resume Next
   DeleteSection gsREG_APP, "Template_" & Trim$(cbTemplate.Text)
   DeleteKeyValue gsREG_APP, "Template", "Template_" & Trim$(cbTemplate.Text)
   Call cbTemplate.RemoveItem(CBLFind(cbTemplate, cbTemplate.Text, 0))

   gsTemplate = "Template_Default"
   SaveSetting gsREG_APP, "General", "DefaultTemplate", gsTemplate

   ' *** Fill the combo with template
   Fill_ProcedureHeader
   Fill_ComboTemplate

   'If (Trim$(cbTemplate.Text) = "") Then
   '   Add_DefaultTemplate
   '   If (CBLFind(cbTemplate, "Default", 0) = -1) Then
   '      cbTemplate.AddItem gsTemplate
   '   Else
   '      cbTemplate.ListIndex = CBLFind(cbTemplate, "Default", 0)
   '      Call cbTemplate_Click
   '   End If
   'End If

End Sub

Private Sub cmdRemove_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdRemove_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Remove One line

   If (lstDestination.ListIndex = -1) Then Exit Sub

   lstSource.AddItem lstDestination.List(lstDestination.ListIndex)
   lstSource.ItemData(lstSource.NewIndex) = lstDestination.ItemData(lstDestination.ListIndex)

   lstDestination.RemoveItem lstDestination.ListIndex

End Sub

Private Sub cmdRemoveAll_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdRemoveAll_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Remove all

   lstDestination.Clear
   Fill_ProcedureHeader

End Sub

Private Sub cmdSaveTemplate_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : cmdSaveTemplate_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Save this template
   SaveTemplate Trim$(cbTemplate.Text)

End Sub

Private Sub lstDestination_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : lstDestination_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   cmdRemove_Click

End Sub

Private Sub lstSource_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : lstSource_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   cmdAdd_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
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

Private Sub InitTooltips(Optional Flags As TooltipFlagConstants)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/11/1999
   ' * Time             : 11:15
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
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

Private Sub vSpace_Change()
   ' #VBIDEUtils#************************************************************
   ' * Author           : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/21/2002
   ' * Time             : 20:49
   ' * Module Name      : frmOptions
   ' * Module Filename  : Options.frm
   ' * Procedure Name   : vSpace_Change
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' *
   ' * Screenshot       :
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   tbSpaceNum.Text = vSpace.Value

End Sub
