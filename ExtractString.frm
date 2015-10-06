VERSION 5.00
Object = "{D1DAC785-7BF2-42C1-9915-A540451B87F2}#1.1#0"; "VBIDEUtils1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExtractString 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Project internationalization"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10695
   Icon            =   "ExtractString.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frametab 
      Caption         =   "Internationalization of the project"
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
      Height          =   6615
      Index           =   4
      Left            =   600
      TabIndex        =   46
      Top             =   1200
      Width           =   10215
      Begin VB.CommandButton cmdBrowse2 
         Height          =   285
         Left            =   3360
         Picture         =   "ExtractString.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Browse to find the database"
         Top             =   480
         Width           =   375
      End
      Begin VB.ListBox lstResult 
         Height          =   4740
         Left            =   120
         TabIndex        =   50
         ToolTipText     =   "This is the result of the internationalization done by the VBIDEUtils"
         Top             =   1200
         Width           =   9975
      End
      Begin VB.TextBox tbDatabaseTranslation 
         Height          =   285
         Left            =   840
         TabIndex        =   47
         ToolTipText     =   "Indicates the database to use"
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton cmdInternationalizeProject 
         Caption         =   "&Internationalize"
         Height          =   1035
         Left            =   3960
         Picture         =   "ExtractString.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Launch the internationalization of your VB project, don't forget to make a backup before"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Database"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   525
         Width           =   690
      End
   End
   Begin VB.PictureBox picSideBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   255
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   9000
      Picture         =   "ExtractString.frx":3116
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Bye bye string extractor"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame frametab 
      Caption         =   "Translate strings"
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
      Height          =   6615
      Index           =   3
      Left            =   480
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton cmdSaveToResource 
         Caption         =   "Save &Resource"
         Height          =   555
         Left            =   8640
         Picture         =   "ExtractString.frx":3420
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Save all the translations to a resource file"
         Top             =   240
         Width           =   1335
      End
      Begin vbAcceleratorGrid.vbalGrid grdGrid 
         Height          =   5655
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "Edit all the translations, move the arrow keys, double-click with mouse or hit enter key to edit"
         Top             =   840
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   9975
         GridLines       =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         GridLineColor   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Editable        =   -1  'True
         DisableIcons    =   -1  'True
         Begin VB.TextBox tbEdit 
            Height          =   375
            Left            =   960
            MultiLine       =   -1  'True
            TabIndex        =   42
            Text            =   "ExtractString.frx":356A
            Top             =   600
            Visible         =   0   'False
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmdSaveTranslation 
         Caption         =   "Save translations"
         Height          =   555
         Left            =   7320
         Picture         =   "ExtractString.frx":3570
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Save all the translations to the database"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdGetTranslation 
         Caption         =   "Get translation"
         Height          =   555
         Left            =   6000
         Picture         =   "ExtractString.frx":36BA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Read all the translations from the database"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox tbDatabase 
         Height          =   285
         Left            =   3000
         TabIndex        =   35
         ToolTipText     =   "Indicates the database to use"
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox tbLanguage 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   34
         Text            =   "2"
         ToolTipText     =   "Indicate the number of desired languages"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowse 
         Height          =   285
         Left            =   5520
         Picture         =   "ExtractString.frx":3804
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Browse to find the database"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Database"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2280
         TabIndex        =   40
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number of languages :"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   405
         Width           =   1605
      End
   End
   Begin VB.Frame frametab 
      Caption         =   "Result of the string extraction"
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
      Height          =   6615
      Index           =   2
      Left            =   360
      TabIndex        =   29
      Top             =   720
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton cmdLoadResource 
         Caption         =   "&Load a resource file"
         Height          =   975
         Left            =   4560
         Picture         =   "ExtractString.frx":3D36
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Load a resource file"
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveResource 
         Caption         =   "Save to &Resource file"
         Height          =   975
         Left            =   4560
         Picture         =   "ExtractString.frx":4178
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Save all the extracted string to a resource file"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveDB 
         Caption         =   "Save to &DB"
         Height          =   975
         Left            =   4560
         Picture         =   "ExtractString.frx":45BA
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Save all the extracted string to a database"
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveString 
         Caption         =   "&Save strings to text file"
         Height          =   975
         Left            =   4560
         Picture         =   "ExtractString.frx":5074
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Save all the extracted string to a text file"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddExtracted 
         Caption         =   "&>"
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Add a non extracted string to the extracted list"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemoveExtracted 
         Caption         =   "&<"
         Height          =   375
         Left            =   4560
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Remove an extracted string from the extracted list"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearAllExtracted 
         Caption         =   "&Clear All"
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Clear all the extracted list"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Caption         =   "Strings found but not extracted"
         ForeColor       =   &H00004000&
         Height          =   6255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   4215
         Begin VB.ListBox lstNotExtracted 
            Height          =   5910
            Left            =   120
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "List of all non extracted strings"
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Extracted strings"
         ForeColor       =   &H00004000&
         Height          =   6255
         Left            =   5880
         TabIndex        =   30
         Top             =   240
         Width           =   4215
         Begin VB.ListBox lstExtracted 
            Height          =   5910
            Left            =   120
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "list of all extracted strings"
            Top             =   240
            Width           =   3975
         End
      End
   End
   Begin VB.Frame frametab 
      Caption         =   "Setup string disqualifiers"
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
      Height          =   6615
      Index           =   1
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   10215
      Begin VB.CommandButton cmdExtracts 
         Caption         =   "&Extract all strings"
         Height          =   615
         Left            =   8760
         Picture         =   "ExtractString.frx":5B2E
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Launch the extraction. It could take a few minutes, depending on the size of the project"
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox tbLength 
         Height          =   285
         Left            =   3720
         MaxLength       =   1
         TabIndex        =   15
         Text            =   "2"
         ToolTipText     =   "Set the minimal number of characters needed to be a string"
         Top             =   6000
         Width           =   375
      End
      Begin VB.CheckBox chkGetMinimumSize 
         Caption         =   "Disqualify string having a legnth less than"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Enable/disable disqualifying strings having less than..."
         Top             =   6000
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.CheckBox chkDisqualifyUpperCaseOnly 
         Caption         =   "Disqualify ""UPPERCASE ONLY"" string"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Disqualify the use of uppercase only strings"
         Top             =   5640
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.Frame Frame4 
         Caption         =   "Entire line disqualification"
         ForeColor       =   &H00004000&
         Height          =   5175
         Left            =   6840
         TabIndex        =   28
         Top             =   240
         Width           =   3255
         Begin VB.CommandButton cmdAddNewEntireLine 
            Caption         =   "Add new"
            Height          =   315
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Add a new line disqualification"
            Top             =   4320
            Width           =   1455
         End
         Begin VB.CommandButton cmdRemoveEntireLine 
            Caption         =   "Remove"
            Height          =   315
            Left            =   1680
            TabIndex        =   11
            ToolTipText     =   "Remove the selected line disqualification"
            Top             =   4320
            Width           =   1455
         End
         Begin VB.ListBox lstEntireLine 
            Height          =   3960
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "If the line disqualifier is found on a line, the line is ignored"
            Top             =   240
            Width           =   3015
         End
         Begin VB.CheckBox chkEntireLineDisqualification 
            Caption         =   "Enable entire line disqualification"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Enabled/disable the use of line disqualification"
            Top             =   4800
            Value           =   1  'Checked
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "String content disqualification"
         ForeColor       =   &H00004000&
         Height          =   5175
         Left            =   3480
         TabIndex        =   27
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox chkContentDisqualification 
            Caption         =   "Enable string content disqualification"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Enabled/disable the use of content disqualification"
            Top             =   4800
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.ListBox lstStringContent 
            Height          =   3960
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "List of strings content disqualification, if the string contents one of those disqualifiers, the string is ignored"
            Top             =   240
            Width           =   3015
         End
         Begin VB.CommandButton cmdRemoveStringContent 
            Caption         =   "Remove"
            Height          =   315
            Left            =   1680
            TabIndex        =   7
            ToolTipText     =   "Remove the selected content disqualification"
            Top             =   4320
            Width           =   1455
         End
         Begin VB.CommandButton cmdAddNewStringContent 
            Caption         =   "Add new"
            Height          =   315
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Add a new content disqualification"
            Top             =   4320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Assignment disqualification"
         ForeColor       =   &H00004000&
         Height          =   5175
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox chkAssignmentDisqualification 
            Caption         =   "Enable assignment disqualification"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Enabled/disable the use of assignment disqualification"
            Top             =   4800
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.CommandButton cmdAddNewAssignment 
            Caption         =   "Add new"
            Height          =   315
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Add a new assignment disqualification"
            Top             =   4320
            Width           =   1455
         End
         Begin VB.CommandButton cmdRemoveAssignment 
            Caption         =   "Remove"
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            ToolTipText     =   "Remove the selected assignment disqualification"
            Top             =   4320
            Width           =   1455
         End
         Begin VB.ListBox lstAssignment 
            Height          =   3960
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   1
            ToolTipText     =   "List of assignment disqualification, if the disqualifier is before the string, the string is ignored"
            Top             =   240
            Width           =   3015
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   12515
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Setup"
            Key             =   "Setup"
            Object.ToolTipText     =   "Setup string disqualifiers"
            ImageVarType    =   8
            ImageKey        =   "Setup"
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Result"
            Key             =   "Result"
            Object.ToolTipText     =   "Result of the string extraction"
            ImageVarType    =   8
            ImageKey        =   "Result"
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Translation"
            Key             =   "Translation"
            Object.ToolTipText     =   "Translate the string into other language"
            ImageVarType    =   8
            ImageKey        =   "Translation"
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Internationalize Project"
            Key             =   "TranslateProject"
            Object.ToolTipText     =   "Translate the current project"
            ImageVarType    =   8
            ImageKey        =   "Translate"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8388736
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExtractString.frx":5C78
            Key             =   "Setup"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExtractString.frx":6212
            Key             =   "Result"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExtractString.frx":652C
            Key             =   "Translation"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExtractString.frx":6846
            Key             =   "Translate"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExtractString.frx":6B60
            Key             =   "Infos"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmExtractString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 14/10/99
' * Time             : 13:33
' * Module Name      : frmExtractString
' * Module Filename  : ExtractString.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Private mnCurFrame      As Integer

Private blstChange      As Boolean

Private bgrdChange      As Boolean

Private nMaxTranslation As Integer

Private WithEvents clsPopupMenu As cPopupMenu
Attribute clsPopupMenu.VB_VarHelpID = -1

Private WithEvents CommonDialog1 As class_CommonDialog
Attribute CommonDialog1.VB_VarHelpID = -1

Private Sub commonDialog1_InitDialog(ByVal hDlg As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/03/2000
   ' * Time             : 15:29
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : commonDialog1_InitDialog
   ' * Parameters       :
   ' *                    ByVal hDlg As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call CommonDialog1.CentreDialog(hDlg, Me)

End Sub

Private Sub cmdAddExtracted_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 16:21
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdAddExtracted_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Long

   For nI = lstNotExtracted.ListCount - 1 To 0 Step -1
      If lstNotExtracted.Selected(nI) = True Then
         Call AddItemToList(Me, lstExtracted, lstNotExtracted.List(nI))
         Call lstNotExtracted.RemoveItem(nI)
         lstNotExtracted.ListIndex = nI
      End If
   Next

   'If lstNotExtracted.ListIndex <> -1 Then
   '   Call AddItemToList(Me, lstExtracted, lstNotExtracted.List(lstNotExtracted.ListIndex))
   '   Call lstNotExtracted.RemoveItem(lstNotExtracted.ListIndex)
   'End If

End Sub

Private Sub cmdAddNewAssignment_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 15:59
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdAddNewAssignment_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sNew             As String

   sNew = InputBox("Enter a new assignment disqualification", "New assignment disqualification")

   If Trim$(sNew) <> "" Then
      blstChange = True
      Call AddItemToList(Me, lstAssignment, sNew)
   End If

End Sub

Private Sub cmdAddNewEntireLine_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 16:00
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdAddNewEntireLine_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sNew             As String

   sNew = InputBox("Enter a new entire line disqualification", "New entire line disqualification")

   If Trim$(sNew) <> "" Then
      blstChange = True
      Call AddItemToList(Me, lstEntireLine, sNew)
   End If

End Sub

Private Sub cmdAddNewStringContent_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 16:00
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdAddNewStringContent_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sNew             As String

   sNew = InputBox("Enter a new string content disqualification", "New string content disqualification")

   If Trim$(sNew) <> "" Then
      blstChange = True
      Call AddItemToList(Me, lstStringContent, sNew)
   End If

End Sub

Private Sub cmdBrowse_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 11:48
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdBrowse_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set CommonDialog1 = New class_CommonDialog

   If bgrdChange Then
      Call cmdSaveTranslation_Click
   End If

   CommonDialog1.DialogTitle = "Choose the database to use"
   CommonDialog1.DefaultExt = "*.mdb"
   If VBInstance.ActiveVBProject Is Nothing Then
      CommonDialog1.FileName = "ExtractString.mdb"
   Else
      CommonDialog1.FileName = VBInstance.ActiveVBProject.Name & ".mdb"
   End If
   CommonDialog1.Filter = "Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
   CommonDialog1.InitDir = App.Path
   CommonDialog1.Flags = OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_EXPLORER
   CommonDialog1.CancelError = False
   CommonDialog1.HookDialog = True
   CommonDialog1.ShowOpen

   If CommonDialog1.FileName = "" Then Exit Sub

   tbDatabase.Text = CommonDialog1.FileName

End Sub

Private Sub cmdBrowse2_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 11:48
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdBrowse2_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set CommonDialog1 = New class_CommonDialog

   If bgrdChange Then
      Call cmdSaveTranslation_Click
   End If

   CommonDialog1.DialogTitle = "Choose the database to use for internationalization"
   CommonDialog1.DefaultExt = "*.mdb"
   If VBInstance.ActiveVBProject Is Nothing Then
      CommonDialog1.FileName = "ExtractString.mdb"
   Else
      CommonDialog1.FileName = VBInstance.ActiveVBProject.Name & ".mdb"
   End If
   CommonDialog1.Filter = "Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
   CommonDialog1.InitDir = App.Path
   CommonDialog1.Flags = OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_EXPLORER
   CommonDialog1.CancelError = False
   CommonDialog1.HookDialog = True
   CommonDialog1.ShowOpen

   If CommonDialog1.FileName = "" Then Exit Sub

   tbDatabaseTranslation.Text = CommonDialog1.FileName

End Sub

Private Sub cmdClearAllExtracted_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 16:23
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdClearAllExtracted_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If MsgBoxTop(Me.hWnd, "Are you sure to clear all the extracted string?", vbQuestion + vbYesNo + vbDefaultButton1, "Clear all extracted string") = vbYes Then
      lstNotExtracted.Clear
      lstExtracted.Clear
   End If

End Sub

Private Sub cmdExit_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 14:08
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdExit_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If bgrdChange Then
      Call cmdSaveTranslation_Click
   End If

   If blstChange = True Then
      If MsgBoxTop(Me.hWnd, "Do you want to save the changes made to the lists", vbQuestion + vbYesNo + vbDefaultButton1, "Save changes") = vbYes Then
         Call SaveList
      End If
   End If

   Call UnloadEffect(Me)
   Unload Me

End Sub

Private Sub cmdExtracts_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 14:08
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdExtracts_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Long

   If bgrdChange Then
      Call cmdSaveTranslation_Click
   End If

   If blstChange = True Then
      If MsgBoxTop(Me.hWnd, "Do you want to save the changes made to the lists", vbQuestion + vbYesNo + vbDefaultButton1, "Save changes") = vbYes Then
         Call SaveList
      End If
      blstChange = False
   End If

   Call SaveSetting(gsREG_APP, "ExtractString", "ExceptLine", IIf(chkEntireLineDisqualification.Value = vbChecked, "Y", "N"))
   Call SaveSetting(gsREG_APP, "ExtractString", "ExceptAssign", IIf(chkAssignmentDisqualification.Value = vbChecked, "Y", "N"))
   Call SaveSetting(gsREG_APP, "ExtractString", "ExceptString", IIf(chkContentDisqualification.Value = vbChecked, "Y", "N"))

   Call SaveSetting(gsREG_APP, "ExtractString", "ExceptAllUpperCase", IIf(chkDisqualifyUpperCaseOnly.Value = vbChecked, "Y", "N"))

   Call SaveSetting(gsREG_APP, "ExtractString", "IsGetMinimumSize", IIf(chkGetMinimumSize.Value = vbChecked, "Y", "N"))
   Call SaveSetting(gsREG_APP, "ExtractString", "NumGetMinimumSize", tbLength.Text)

   lstNotExtracted.Clear
   lstExtracted.Clear

   ' *** Reload all
   Call InitExtractString

   ' *** Extract all the strings
   Call ExtractString

   ' *** Add them to the listbox
   For nI = 1 To colExtracted.Count
      Call AddItemToList(Me, lstExtracted, colExtracted(nI))
   Next

   For nI = 1 To colRefused.Count
      Call AddItemToList(Me, lstNotExtracted, colRefused(nI))
   Next

   Set TabStrip.SelectedItem = TabStrip.Tabs("Result")

End Sub

Private Sub cmdInternationalizeProject_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/12/1999
   ' * Time             : 14:42
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdInternationalizeProject_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   lstResult.Clear

   If FileExist(tbDatabaseTranslation.Text) = False Then
      Call MsgBoxTop(Me.hWnd, "The specified translation database does not exists", vbCritical, "Internationalize the project")
      Exit Sub
   End If

   If VBInstance.ActiveVBProject Is Nothing Then Exit Sub

   If VBInstance.ActiveVBProject.Saved = False Then
      If MsgBoxTop(Me.hWnd, "You need to save the project." & Chr$(13) & "Do you want to save it and continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Save VB Project") = vbNo Then Exit Sub
      Call VBInstance.ActiveVBProject.SaveAs(VBInstance.ActiveVBProject.FileName)
   End If

   If MsgBoxTop(Me.hWnd, "Do you want to internationalize your VB Project?" & Chr$(13) & "Have you done a backup before?", vbQuestion + vbYesNo + vbDefaultButton1, "Internationalize Project") = vbNo Then Exit Sub

   ' *** Reload all
   Call InitExtractString

   ' *** Internationalize the project
   Call InternationalizeProject(tbDatabaseTranslation.Text)

End Sub

Private Sub cmdLoadResource_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/12/1999
   ' * Time             : 12:28
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdLoadResource_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdLoadResource_Click

   Dim sFileName        As String

   Dim nFile            As Integer
   Dim sFile            As String

   Dim sLine            As String

   Dim nPos             As String
   Dim sTmp             As String

   Set CommonDialog1 = New class_CommonDialog

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   With CommonDialog1
      .DialogTitle = "Choose a resource file to load"
      .CancelError = False
      .Flags = OFN_HIDEREADONLY + OFN_EXPLORER
      .InitDir = App.Path
      .FileName = "*.rc"
      .Filter = "Resource File (*.rc)|*.rc|Text File (*.txt)|*.txt|All Files (*.*)|*.*"
      .FilterIndex = 1
      .HookDialog = True
      .ShowOpen

   End With

   If CommonDialog1.FileName = "" Then Exit Sub

   sFileName = CommonDialog1.FileName

   If FileExist(sFileName) = False Then
      Call MsgBoxTop(Me.hWnd, "The specified file does not exists", vbCritical, "Import a resource")
      GoTo EXIT_cmdLoadResource_Click
   End If

   nFile = FreeFile
   Open sFileName For Input Access Read As #nFile
   sFile = Input(LOF(nFile), nFile)
   Close #nFile

   Call SetRedraw(Me, False)

   ' *** Clear the list
   lstExtracted.Clear
   sFile = Replace(sFile, vbTab, "  ")
   sLine = GetALine(sFile)
   Do While Len(sFile) > 0
      ' *** find the first "
      nPos = InStr(sLine, """")

      If nPos > 0 Then
         sTmp = Trim$(left(sLine, nPos - 1))
         If IsNumeric(sTmp) Then
            ' *** String line
            sTmp = Mid$(sLine, nPos + 1)
            sTmp = Trim$(sTmp)
            sTmp = Replace(sTmp, """""", """")
            If right(sTmp, 1) = """" Then sTmp = left$(sTmp, Len(sTmp) - 1)
            Call AddItemToList(Me, lstExtracted, sTmp)
         End If
      End If

      sLine = GetALine(sFile)
   Loop
   Call SetRedraw(Me, True)

   Call MsgBoxTop(Me.hWnd, "The resource file " & sFileName & " has been imported!", vbInformation + vbOKOnly + vbDefaultButton1, "Import resource file")

EXIT_cmdLoadResource_Click:
   Call SetRedraw(Me, True)
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdLoadResource_Click:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdLoadResource_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdLoadResource_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_cmdLoadResource_Click

End Sub

Private Sub cmdRemoveAssignment_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 15:58
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdRemoveAssignment_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If lstAssignment.ListIndex > -1 Then
      blstChange = True
      lstAssignment.RemoveItem lstAssignment.ListIndex
   End If

End Sub

Private Sub cmdRemoveEntireLine_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 15:58
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdRemoveEntireLine_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If lstEntireLine.ListIndex > -1 Then
      blstChange = True
      lstEntireLine.RemoveItem lstEntireLine.ListIndex
   End If

End Sub

Private Sub cmdRemoveExtracted_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 16:22
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdRemoveExtracted_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Long
   Dim nOld             As Long

   nOld = -1

   For nI = lstExtracted.ListCount - 1 To 0 Step -1
      If lstExtracted.Selected(nI) = True Then
         Call AddItemToList(Me, lstNotExtracted, lstExtracted.List(nI))
         Call lstExtracted.RemoveItem(nI)
         nOld = nI
      End If
   Next

   On Error Resume Next
   lstExtracted.Selected(nOld) = True

   'If lstExtracted.ListIndex <> -1 Then
   '   Call AddItemToList(Me, lstNotExtracted, lstExtracted.List(lstExtracted.ListIndex))
   '   Call lstExtracted.RemoveItem(lstExtracted.ListIndex)
   '   On Error Resume Next
   '   lstExtracted.Selected(lstExtracted.ListIndex) = True
   'End If

End Sub

Private Sub cmdRemoveStringContent_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 15:58
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdRemoveStringContent_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If lstStringContent.ListIndex > -1 Then
      blstChange = True
      lstStringContent.RemoveItem lstStringContent.ListIndex
   End If

End Sub

Private Sub cmdSaveDB_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 09:46
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdSaveDB_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdSaveDB_Click

   Dim sFileName        As String

   Dim nI               As Integer

   Dim nMax             As Long

   Dim DBLocal          As Database
   Dim record           As Recordset

   Dim bOk              As Boolean

   Set CommonDialog1 = New class_CommonDialog

   Dim cHourglass       As class_Hourglass

   If Not gbRegistered Then
      Call MsgBoxTop(Me.hWnd, "In the Shareware Version, you are not allowed to export to a database", vbExclamation + vbOKOnly + vbDefaultButton1, "Shareware Version")
      Exit Sub
   End If

   With CommonDialog1
      .DialogTitle = "Choose a DB to save"
      .CancelError = False
      .Flags = OFN_HIDEREADONLY + OFN_EXPLORER
      .InitDir = App.Path
      If VBInstance.ActiveVBProject Is Nothing Then
         .FileName = "ExtractString.mdb"
      Else
         .FileName = VBInstance.ActiveVBProject.Name & ".mdb"
      End If
      .Filter = "Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
      .FilterIndex = 1
      .HookDialog = True
      .ShowSave

   End With

   If CommonDialog1.FileName = "" Then Exit Sub

   sFileName = CommonDialog1.FileName

   nMaxTranslation = 2

   ' *** Generate the database
   On Error Resume Next
   If FileExist(sFileName) Then
      If MsgBoxTop(Me.hWnd, "Do you want to complete the existing database?" & Chr$(13) & "If you select no, the DB will be overwrited!", vbQuestion + vbYesNo + vbDefaultButton1, "Save to a database") = vbNo Then Kill sFileName
   End If
   On Error GoTo ERROR_cmdSaveDB_Click

   Set cHourglass = New class_Hourglass
   If FileExist(sFileName) = False Then
      bOk = GenerateDatabaseTranslation(sFileName)
   Else
      bOk = True
   End If

   If bOk Then
      Set DBLocal = OpenDatabase(sFileName)
      Set record = DBLocal.OpenRecordset("Translation")
      For nI = 0 To lstExtracted.ListCount - 1
         On Error Resume Next
         record.AddNew
         record("Original") = left$(lstExtracted.List(nI), 250)
         record.Update
      Next
      record.Close
      Set record = Nothing

      DBLocal.Close
      Set DBLocal = Nothing

      Call MsgBoxTop(Me.hWnd, "The database " & sFileName & " has been generated!", vbInformation + vbOKOnly + vbDefaultButton1, "Save extracted strings")

   End If

EXIT_cmdSaveDB_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdSaveDB_Click:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdSaveDB_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdSaveDB_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_cmdSaveDB_Click

End Sub

Private Sub cmdSaveResource_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/12/1999
   ' * Time             : 11:22
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdSaveResource_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdSaveResource_Click

   Dim sFileName        As String

   Dim nFile            As Integer
   Dim nI               As Integer

   Dim nMax             As Long

   Set CommonDialog1 = New class_CommonDialog

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   With CommonDialog1
      .DialogTitle = "Choose a filename to save"
      .CancelError = False
      .Flags = OFN_HIDEREADONLY + OFN_EXPLORER
      .InitDir = App.Path
      If VBInstance.ActiveVBProject Is Nothing Then
         .FileName = "ExtractString.rc"
      Else
         .FileName = VBInstance.ActiveVBProject.Name & ".rc"
      End If
      .Filter = "Resource File (*.rc)|*.rc|Text File (*.txt)|*.txt|All Files (*.*)|*.*"
      .FilterIndex = 1
      .HookDialog = True
      .ShowSave

   End With

   If CommonDialog1.FileName = "" Then Exit Sub

   sFileName = CommonDialog1.FileName

   nFile = FreeFile
   Open sFileName For Output Access Write As #nFile
   If gbRegistered Then
      nMax = lstExtracted.ListCount - 1
   Else
      Call MsgBoxTop(Me.hWnd, "In the Shareware Version, you are only allowed to export 10 strings", vbExclamation + vbOKOnly + vbDefaultButton1, "Shareware Version")

      If lstExtracted.ListCount - 1 > 10 Then
         nMax = 10
      Else
         nMax = lstExtracted.ListCount - 1
      End If
   End If

   Print #nFile, "// #VBIDEUtils#************************************************************"
   Print #nFile, "// * Generated by VBIDEUtils at http://http://www.ppreview.net"
   Print #nFile, "// *"
   Print #nFile, "// * Programmer Name  : " & gsDevelopper
   Print #nFile, "// * Web Site         : " & gsWebSite
   Print #nFile, "// * E-Mail           : " & gsEmail
   Print #nFile, "// * Date             : " & Date
   Print #nFile, "// * Time             : " & time
   Print #nFile, "// **********************************************************************"
   Print #nFile, "// * Comments         :"
   Print #nFile, "// *"
   If Not (VBInstance.ActiveVBProject Is Nothing) Then
      Print #nFile, "// * Project       : " & VBInstance.ActiveVBProject.Name
      Print #nFile, "// *                 " & VBInstance.ActiveVBProject.Description
   End If
   Print #nFile, "// *"
   Print #nFile, "// **********************************************************************"
   Print #nFile, "STRINGTABLE DISCARDABLE"
   Print #nFile, "BEGIN"
   For nI = 0 To nMax
      Print #nFile, PadL(CStr(nI), 6) & vbTab & vbTab & """" & Replace(lstExtracted.List(nI), """", """""") & """"
   Next
   Print #nFile, "END"
   Print #nFile, ""
   Print #nFile, "// **********************************************************************"
   Print #nFile, "// * That's all folks......"
   Print #nFile, "// **********************************************************************"
   Close #nFile

   Call MsgBoxTop(Me.hWnd, "The string file " & sFileName & " has been generated!", vbInformation + vbOKOnly + vbDefaultButton1, "Save extracted strings")

EXIT_cmdSaveResource_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdSaveResource_Click:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdSaveResource_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdSaveResource_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_cmdSaveResource_Click

End Sub

Private Sub cmdSaveString_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 08:59
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdSaveString_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdSaveString_Click

   Dim sFileName        As String

   Dim nFile            As Integer
   Dim nI               As Integer

   Dim nMax             As Long

   Set CommonDialog1 = New class_CommonDialog

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   With CommonDialog1
      .DialogTitle = "Choose a filename to save"
      .CancelError = False
      .Flags = OFN_HIDEREADONLY + OFN_EXPLORER
      .InitDir = App.Path
      .FileName = "ExtractString.txt"
      .Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
      .FilterIndex = 1
      .HookDialog = True
      .ShowSave

   End With

   If CommonDialog1.FileName = "" Then Exit Sub

   sFileName = CommonDialog1.FileName

   nFile = FreeFile
   Open sFileName For Output Access Write As #nFile
   If gbRegistered Then
      nMax = lstExtracted.ListCount - 1
   Else
      Call MsgBoxTop(Me.hWnd, "In the Shareware Version, you are only allowed to export 10 strings", vbExclamation + vbOKOnly + vbDefaultButton1, "Shareware Version")

      If lstExtracted.ListCount - 1 > 10 Then
         nMax = 10
      Else
         nMax = lstExtracted.ListCount - 1
      End If
   End If

   For nI = 0 To nMax
      Print #nFile, lstExtracted.List(nI)
   Next
   Close #nFile

   Call MsgBoxTop(Me.hWnd, "The string file " & sFileName & " has been generated!", vbInformation + vbOKOnly + vbDefaultButton1, "Save extracted strings")

EXIT_cmdSaveString_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdSaveString_Click:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdSaveString_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdSaveString_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_cmdSaveString_Click

End Sub

Private Sub cmdSaveToResource_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/12/1999
   ' * Time             : 13:08
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdSaveToResource_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdSaveToResource_Click

   Dim sFileName        As String

   Dim nFile            As Integer
   Dim nI               As Long
   Dim nJ               As Long

   Dim nMax             As Long
   Dim nLangues         As Long
   Dim nNextLangue      As Long

   Set CommonDialog1 = New class_CommonDialog

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   With CommonDialog1
      .DialogTitle = "Choose a filename to save"
      .CancelError = False
      .Flags = OFN_HIDEREADONLY + OFN_EXPLORER
      .InitDir = App.Path
      If VBInstance.ActiveVBProject Is Nothing Then
         .FileName = "ExtractString.rc"
      Else
         .FileName = VBInstance.ActiveVBProject.Name & ".rc"
      End If
      .Filter = "Resource File (*.rc)|*.rc|Text File (*.txt)|*.txt|All Files (*.*)|*.*"
      .FilterIndex = 1
      .HookDialog = True

      .ShowSave

   End With

   If CommonDialog1.FileName = "" Then Exit Sub

   sFileName = CommonDialog1.FileName

   nFile = FreeFile
   Open sFileName For Output Access Write As #nFile
   If gbRegistered Then
      nMax = grdGrid.Rows
   Else
      Call MsgBoxTop(Me.hWnd, "In the Shareware Version, you are only allowed to export 10 strings", vbExclamation + vbOKOnly + vbDefaultButton1, "Shareware Version")

      If lstExtracted.ListCount - 1 > 10 Then
         nMax = 10
      Else
         nMax = grdGrid.Rows - 1
      End If
   End If
   nLangues = tbLanguage.Text
   nNextLangue = CLng("1" & String(Len(CStr(nMax * 10)), "0"))

   Print #nFile, "// #VBIDEUtils#************************************************************"
   Print #nFile, "// * Generated by VBIDEUtils at http://http://www.ppreview.net"
   Print #nFile, "// *"
   Print #nFile, "// * Programmer Name  : " & gsDevelopper
   Print #nFile, "// * Web Site         : " & gsWebSite
   Print #nFile, "// * E-Mail           : " & gsEmail
   Print #nFile, "// * Date             : " & Date
   Print #nFile, "// * Time             : " & time
   Print #nFile, "// **********************************************************************"
   Print #nFile, "// * Comments         :"
   Print #nFile, "// *"
   If Not (VBInstance.ActiveVBProject Is Nothing) Then
      Print #nFile, "// * Project          : " & VBInstance.ActiveVBProject.Name
      Print #nFile, "// *                    " & VBInstance.ActiveVBProject.Description
   End If
   Print #nFile, "// *"
   Print #nFile, "// **********************************************************************"
   Print #nFile, "//"
   If nLangues > 1 Then
      Print #nFile, "// * There are " & CStr(nLangues) & " languages defined"
      Print #nFile, "//"
      For nJ = 1 To nLangues
         Print #nFile, "// Language " & nJ & " begins at " & CStr((nJ - 1) * nNextLangue + 1)
      Next
      Print #nFile, "//"
   End If
   Print #nFile, "// Language number 1"
   Print #nFile, "STRINGTABLE DISCARDABLE"
   Print #nFile, "BEGIN"
   For nJ = 1 To nLangues
      For nI = 1 To nMax
         Print #nFile, PadL(CStr((nJ - 1) * nNextLangue + grdGrid.CellText(nI, 1)), 6) & vbTab & vbTab & """" & Replace(grdGrid.CellText(nI, nJ + 1), """", """""") & """"
      Next
      If nJ < nLangues Then
         Print #nFile, "END"
         Print #nFile, ""
         Print #nFile, "// Language number " & CStr(nJ + 1)
         Print #nFile, "STRINGTABLE DISCARDABLE"
         Print #nFile, "BEGIN"
      End If
   Next
   Print #nFile, "END"
   Print #nFile, ""
   Print #nFile, "// **********************************************************************"
   Print #nFile, "// * That's all folks......"
   Print #nFile, "// **********************************************************************"
   Close #nFile

   Call MsgBoxTop(Me.hWnd, "The string file " & sFileName & " has been generated!", vbInformation + vbOKOnly + vbDefaultButton1, "Save extracted strings")

EXIT_cmdSaveToResource_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdSaveToResource_Click:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdSaveToResource_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdSaveToResource_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_cmdSaveToResource_Click

End Sub

Private Sub cmdSaveTranslation_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 14:03
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdSaveTranslation_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim DBLocal          As Database
   Dim record           As Recordset

   Dim nI               As Long
   Dim nJ               As Long

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdSaveTranslation_Click

   If Not gbRegistered Then
      Call MsgBoxTop(Me.hWnd, "In the Shareware Version, you are not allowed to save to a database", vbExclamation + vbOKOnly + vbDefaultButton1, "Shareware Version")
      Exit Sub
   End If

   If bgrdChange Then
      If MsgBoxTop(Me.hWnd, "Do you want to save the translation to the DB?", vbQuestion + vbYesNo + vbDefaultButton1, "Save translations") = vbNo Then GoTo EXIT_cmdSaveTranslation_Click

      nMaxTranslation = grdGrid.Columns - 2

      If Trim$(tbDatabase.Text) = "" Then
         Call MsgBoxTop(Me.hWnd, "You need to specify a translation database", vbExclamation + vbOKOnly + vbDefaultButton1, "Save translation database")
         GoTo ERROR_cmdSaveTranslation_Click
      End If

      ' *** Generate the database
      On Error Resume Next
      If FileExist(tbDatabase.Text) Then Kill tbDatabase.Text
      On Error GoTo ERROR_cmdSaveTranslation_Click

      If GenerateDatabaseTranslation(tbDatabase.Text) Then
         Set DBLocal = OpenDatabase(tbDatabase.Text)
         Set record = DBLocal.OpenRecordset("Translation")
         For nI = 1 To grdGrid.Rows
            record.AddNew
            record("ID") = grdGrid.CellText(nI, 1)
            record("Original") = grdGrid.CellText(nI, 2)
            For nJ = 3 To grdGrid.Columns
               record("Translation" & CStr(nJ - 2)) = grdGrid.CellText(nI, nJ)
            Next
            record.Update
         Next
         record.Close
         Set record = Nothing

         DBLocal.Close
         Set DBLocal = Nothing

         bgrdChange = False
         Call MsgBoxTop(Me.hWnd, "The database " & tbDatabase.Text & " has been saved!", vbInformation + vbOKOnly + vbDefaultButton1, "Save extracted strings")
      End If

   End If

EXIT_cmdSaveTranslation_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdSaveTranslation_Click:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdSaveTranslation_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdSaveTranslation_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_cmdSaveTranslation_Click

End Sub

Private Sub grdGrid_ColumnWidthChanging(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:48
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : grdGrid_ColumnWidthChanging
   ' * Parameters       :
   ' *                    ByVal lCol As Long
   ' *                    ByVal lWidth As Long
   ' *                    bCancel As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lLeft            As Long
   Dim lTop             As Long
   Dim lHeight          As Long

   If tbEdit.Visible Then
      grdGrid.CellBoundary grdGrid.SelectedRow, lCol, lLeft, lTop, lWidth, lHeight
      tbEdit.Width = lWidth
   End If

End Sub

Private Sub grdGrid_ColumnWidthStartChange(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:48
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : grdGrid_ColumnWidthStartChange
   ' * Parameters       :
   ' *                    ByVal lCol As Long
   ' *                    ByVal lWidth As Long
   ' *                    bCancel As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lLeft            As Long
   Dim lTop             As Long
   Dim lHeight          As Long

   If tbEdit.Visible Then
      grdGrid.CellBoundary grdGrid.SelectedRow, lCol, lLeft, lTop, lWidth, lHeight
      tbEdit.Width = lWidth
   End If

End Sub

Private Sub grdGrid_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:34
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : grdGrid_RequestEdit
   ' * Parameters       :
   ' *                    ByVal lRow As Long
   ' *                    ByVal lCol As Long
   ' *                    ByVal iKeyAscii As Integer
   ' *                    bCancel As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lLeft            As Long
   Dim lTop             As Long
   Dim lWidth           As Long
   Dim lHeight          As Long

   If lCol > 2 Then
      grdGrid.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
      If Not IsMissing(grdGrid.CellText(lRow, lCol)) Then
         tbEdit.Text = grdGrid.CellFormattedText(lRow, lCol)
         tbEdit.SelStart = 0
         tbEdit.SelLength = Len(tbEdit.Text)
         SendKeys Chr$(iKeyAscii)

      Else
         tbEdit.Text = ""
      End If
      Set tbEdit.Font = grdGrid.CellFont(lRow, lCol)
      tbEdit.Move lLeft, lTop, lWidth, lHeight
      tbEdit.Visible = True
      tbEdit.ZOrder
      tbEdit.SetFocus
   End If

End Sub

Private Sub lstExtracted_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 14:37
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : lstExtracted_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call cmdRemoveExtracted_Click

End Sub

Private Sub lstNotExtracted_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 14:37
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : lstNotExtracted_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call cmdAddExtracted_Click

End Sub

Private Sub tbEdit_KeyDown(KeyCode As Integer, Shift As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:36
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : tbEdit_KeyDown
   ' * Parameters       :
   ' *                    KeyCode As Integer
   ' *                    Shift As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyDown) Then
      ' *** Commit edit
      grdGrid.CellText(grdGrid.SelectedRow, grdGrid.SelectedCol) = tbEdit.Text
      tbEdit.Visible = False
      grdGrid.SetFocus
      If grdGrid.SelectedRow < grdGrid.Rows Then grdGrid.SelectedRow = grdGrid.SelectedRow + 1
      KeyCode = 0
   ElseIf (KeyCode = vbKeyUp) Then
      ' *** Commit edit
      grdGrid.CellText(grdGrid.SelectedRow, grdGrid.SelectedCol) = tbEdit.Text
      tbEdit.Visible = False
      grdGrid.SetFocus
      If grdGrid.SelectedRow > 1 Then grdGrid.SelectedRow = grdGrid.SelectedRow - 1
      KeyCode = 0
   ElseIf (KeyCode = vbKeyEscape) Then
      ' *** Cancel edit
      tbEdit.Visible = False
      grdGrid.SetFocus
   Else
      bgrdChange = True
   End If

End Sub

Private Sub tbEdit_LostFocus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:36
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : tbEdit_LostFocus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tbEdit.Visible = False
   grdGrid.CancelEdit

End Sub

Public Sub tabstrip_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 20/09/99
   ' * Time             : 12:50
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
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
   frametab(TabStrip.SelectedItem.index).Visible = True
   frametab(TabStrip.SelectedItem.index).ZOrder
   frametab(mnCurFrame).Visible = False

   ' *** Set mnCurFrame to new value.
   mnCurFrame = TabStrip.SelectedItem.index

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 13:33
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Long

   On Error Resume Next

   Call InitTooltips

   blstChange = False
   bgrdChange = False

   For nI = 1 To TabStrip.Tabs.Count
      frametab(nI).Move TabStrip.left + 120, TabStrip.top + 360, TabStrip.Width - 240, TabStrip.Height - 360 - 120
   Next

   ' *** Fill all the listbox
   lstAssignment.Clear
   For nI = 1 To colExceptAssign.Count
      Call AddItemToList(Me, lstAssignment, colExceptAssign(nI))
   Next

   lstStringContent.Clear
   For nI = 1 To colExceptString.Count
      Call AddItemToList(Me, lstStringContent, colExceptString(nI))
   Next

   lstEntireLine.Clear
   For nI = 1 To colExceptLine.Count
      Call AddItemToList(Me, lstEntireLine, colExceptLine(nI))
   Next

   chkAssignmentDisqualification.Value = IIf(gbExceptAssign, vbChecked, vbUnchecked)
   chkContentDisqualification.Value = IIf(gbExceptString, vbChecked, vbUnchecked)
   chkEntireLineDisqualification.Value = IIf(gbExceptLine, vbChecked, vbUnchecked)
   chkDisqualifyUpperCaseOnly.Value = IIf(gbExceptAllUpperCase, vbChecked, vbUnchecked)
   chkGetMinimumSize.Value = IIf(gbGetMinimumSize, vbChecked, vbUnchecked)
   tbLength.Text = gnGetMinimumSize

   mnCurFrame = 0
   Set TabStrip.SelectedItem = TabStrip.Tabs("Setup")

   ' *** Init the listview for translation
   Call InitGrid

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 13:33
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
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

Private Sub SaveList()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 16:03
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : SaveList
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_SaveList

   Dim nFile            As Integer
   Dim nI               As Integer

   nFile = FreeFile
   Open App.Path & "\Extract1.txt" For Output Access Write As #nFile
   For nI = 0 To lstEntireLine.ListCount - 1
      Print #nFile, lstEntireLine.List(nI)
   Next
   Close #nFile

   nFile = FreeFile
   Open App.Path & "\Extract2.txt" For Output Access Write As #nFile
   For nI = 0 To lstAssignment.ListCount - 1
      Print #nFile, lstAssignment.List(nI)
   Next
   Close #nFile

   nFile = FreeFile
   Open App.Path & "\Extract3.txt" For Output Access Write As #nFile
   For nI = 0 To lstStringContent.ListCount - 1
      Print #nFile, lstStringContent.List(nI)
   Next
   Close #nFile

EXIT_SaveList:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_SaveList:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in SaveList", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_SaveList
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_SaveList

End Sub

Public Function GenerateDatabaseTranslation(sDestDBPath As String, Optional sDestDBPassword As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 9:44:33
   ' * Procedure Name   : GenerateDatabaseTranslation
   ' * Parameters       :
   ' *                    sDestDBPath As String
   ' *                    Optional sDestDBPassword As String
   ' **********************************************************************
   ' * Comments         : Create a new database by code
   ' *   This Database has been created using VBIDEUtils
   ' *
   ' **********************************************************************

   Dim DB               As Database

   On Error GoTo ERROR_GenerateDatabaseTranslation

   GenerateDatabaseTranslation = False

   ' *** Create the database
   If Trim$(sDestDBPassword) <> "" Then
      Set DB = Workspaces(0).CreateDatabase(sDestDBPath, dbLangGeneral & ";pwd=" & sDestDBPassword)
   Else
      Set DB = Workspaces(0).CreateDatabase(sDestDBPath, dbLangGeneral)
   End If

   ' *** Create each table
   ' *** Table Translation
   If CreateTableTranslation(DB) = False Then
      GenerateDatabaseTranslation = False
      DB.Close
      Set DB = Nothing
      Exit Function
   End If

   GenerateDatabaseTranslation = True

EXIT_GenerateDatabaseTranslation:
   On Error Resume Next
   ' *** Close the database :)
   DB.Close
   Set DB = Nothing
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_GenerateDatabaseTranslation:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in GenerateDatabaseTranslation", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         GenerateDatabaseTranslation = False
   '         Resume EXIT_GenerateDatabaseTranslation:
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   GenerateDatabaseTranslation = False
   Resume EXIT_GenerateDatabaseTranslation

End Function

Private Function CreateTableTranslation(DB As Database) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 9:44:34
   ' * Procedure Name   : CreateTableTranslation
   ' * Parameters       :
   ' *                    DB As Database
   ' **********************************************************************
   ' * Comments         : This table has been created using VBIDEUtils
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_CreateTableTranslation

   Dim table            As TableDef
   Dim fld              As Field
   Dim idx              As index

   Dim nI               As Integer

   CreateTableTranslation = False

   ' *** Create the table
   Set table = DB.CreateTableDef("Translation")

   ' *** Create the field ID
   Set fld = table.CreateField("ID", dbLong)
   With fld
      .Attributes = 17
      .Required = False
      .OrdinalPosition = 1
      .Size = 4
   End With
   table.Fields.Append fld
   table.Fields.Refresh

   ' *** Create the field Original
   Set fld = table.CreateField("Original", 10)
   With fld
      .Attributes = 2
      .Required = False
      .OrdinalPosition = 2
      .Size = 250
      .AllowZeroLength = False
   End With
   table.Fields.Append fld
   table.Fields.Refresh

   ' *** Create the field Translation
   For nI = 1 To nMaxTranslation
      Set fld = table.CreateField("Translation" & CStr(nI), 10)
      With fld
         .Attributes = 2
         .Required = False
         .OrdinalPosition = 2 + nI
         .Size = 250
         .AllowZeroLength = True
      End With
      table.Fields.Append fld
      table.Fields.Refresh
   Next

   ' *** Create the Index ID
   Set idx = table.CreateIndex
   With idx
      .Name = "ID"
      .Primary = False
      .Unique = False
      .Required = False
      .Clustered = False
      .IgnoreNulls = False
      Set fld = .CreateField("ID")
      .Fields.Append fld
   End With
   table.Indexes.Append idx
   table.Indexes.Refresh

   ' *** Create the Index Original
   Set idx = table.CreateIndex
   With idx
      .Name = "Original"
      .Primary = False
      .Unique = True
      .Required = False
      .Clustered = False
      .IgnoreNulls = False
      Set fld = .CreateField("Original")
      .Fields.Append fld
   End With
   table.Indexes.Append idx
   table.Indexes.Refresh

   ' *** Create the Index PrimaryKey
   Set idx = table.CreateIndex
   With idx
      .Name = "PrimaryKey"
      .Primary = True
      .Unique = True
      .Required = True
      .Clustered = False
      .IgnoreNulls = False
      Set fld = .CreateField("ID")
      .Fields.Append fld
   End With
   table.Indexes.Append idx
   table.Indexes.Refresh

   DB.TableDefs.Append table
   DB.TableDefs.Refresh

   ' *** This table is done :)
   CreateTableTranslation = True

EXIT_CreateTableTranslation:
   Set idx = Nothing
   Set fld = Nothing
   Set table = Nothing

   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_CreateTableTranslation:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in CreateTableTranslation", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_CreateTableTranslation:
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_CreateTableTranslation

End Function

Private Sub tbLanguage_KeyPress(KeyAscii As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 11:46
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : tbLanguage_KeyPress
   ' * Parameters       :
   ' *                    KeyAscii As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0

End Sub

Private Sub tbLength_KeyPress(KeyAscii As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 11:44
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : tbLength_KeyPress
   ' * Parameters       :
   ' *                    KeyAscii As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0

End Sub

Private Sub InitGrid()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 11:42
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : InitGrid
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer

   On Error Resume Next

   With grdGrid
      ' *** Turn redraw off for speed:
      .Redraw = False

      For nI = CLng(.Columns) To 1 Step -1
         .RemoveColumn (nI)
      Next

      .BackColor = &H80000018

      .AddColumn "ID", "ID", ecgHdrTextALignLeft, , 30
      .AddColumn "Original", "Original", ecgHdrTextALignLeft, , 200
      For nI = 1 To CLng(tbLanguage.Text)
         .AddColumn "Translation" & CStr(nI), "Language " & CStr(nI), ecgHdrTextALignLeft, , 200
      Next

      .SetHeaders

      ' *** Ensure the grid will draw!
      .Redraw = True

   End With

End Sub

Private Sub cmdGetTranslation_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 11:57
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : cmdGetTranslation_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdGetTranslation_Click

   Dim DBLocal          As Database
   Dim record           As Recordset

   Dim nI               As Long
   Dim nJ               As Long

   Dim nColor           As Long
   Dim sTmp             As String

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   If bgrdChange Then
      Call cmdSaveTranslation_Click
   End If

   If FileExist(tbDatabase.Text) = False Then
      Call MsgBoxTop(Me.hWnd, "You need to specify a translation database", vbExclamation + vbOKOnly + vbDefaultButton1, "Read translation database")
      GoTo EXIT_cmdGetTranslation_Click
   End If

   ' *** Init the grid
   Call InitGrid

   ' *** Turn redraw off for speed:
   grdGrid.Redraw = False

   ' *** Clear the grid
   grdGrid.Clear

   ' *** Read the database
   Set DBLocal = OpenDatabase(tbDatabase.Text)
   Set record = DBLocal.OpenRecordset("Select * From Translation Order by ID", DAO.dbOpenDynaset)
   If record.EOF = False Then
      ' *** Fill the Grid
      record.MoveLast
      record.MoveFirst

      ' *** Set the number of rows
      grdGrid.Rows = record.RecordCount

      ' *** Fill the grid
      For nI = 1 To record.RecordCount
         nColor = &H80000018
         grdGrid.CellDetails nI, 1, record("ID"), DT_RIGHT, , vbWhite, vbRed
         grdGrid.CellDetails nI, 2, ReadRecordSet(record, "Original"), DT_LEFT, , nColor, vbBlue
         For nJ = 1 To CLng(tbLanguage.Text)
            If nColor = &H80000018 Then
               nColor = &HE0E0E0
            Else
               nColor = &H80000018
            End If

            sTmp = ""
            sTmp = ReadRecordSet(record, "Translation" & nJ)
            grdGrid.CellDetails nI, nJ + 2, sTmp, DT_LEFT, , nColor, vbBlack
         Next

         record.MoveNext
      Next

   End If
   record.Close
   Set record = Nothing

   DBLocal.Close
   Set DBLocal = Nothing

EXIT_cmdGetTranslation_Click:
   grdGrid.Redraw = True
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdGetTranslation_Click:
   If err = 3059 Then Resume EXIT_cmdGetTranslation_Click
   If err = 3078 Then Resume EXIT_cmdGetTranslation_Click
   If err = 3265 Then Resume Next

   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in cmdGetTranslation_Click", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_cmdGetTranslation_Click
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select
   '
   Resume EXIT_cmdGetTranslation_Click

End Sub

Private Sub grdGrid_ColumnClick(ByVal lCol As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 8/02/99
   ' * Time             : 13:59
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : grdGrid_ColumnClick
   ' * Parameters       :
   ' *                    ByVal lCol As Long
   ' **********************************************************************
   ' * Comments         : Sort by a specific column
   ' *
   ' *
   ' **********************************************************************

   Dim sTag             As String

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   With grdGrid.SortObject
      .Clear
      .SortColumn(1) = lCol

      sTag = grdGrid.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(1) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(1) = CCLOrderDescending
      End If
      grdGrid.ColumnTag(lCol) = sTag

      ' *** Sort by text
      .SortType(1) = CCLSortString

   End With
   grdGrid.Sort

End Sub

Private Sub DisplayStringInfos(nX As Single, nY As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/04/99
   ' * Time             : 13:15
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : DisplayStringInfos
   ' * Parameters       :
   ' *                    nX As Single
   ' *                    nY As Single
   ' **********************************************************************
   ' * Comments         : Display a popupmenu
   ' *
   ' *
   ' **********************************************************************

   Dim nItem            As Long
   Dim nIndex           As Long
   Dim nI               As Integer

   Set clsPopupMenu = New cPopupMenu

   ' *** Initialise the Image List:
   clsPopupMenu.ImageList = ImageList1
   clsPopupMenu.GradientHighlight = True

   ' *** Initialise the hWndOwner (you must do this before showing a menu):
   clsPopupMenu.hWndOwner = Me.hWnd

   'nIconIndex = ImageToolbar.ListImages("Infos").Index - 1
   'nItem = clsPopupMenu.AddItem("&Show where the string is in the, , , , nIconIndex, , , "Infos")
   'clsPopupMenu.OwnerDraw(nItem) = True

   nItem = clsPopupMenu.AddItem("-")
   clsPopupMenu.OwnerDraw(nItem) = True

   ' Firstly, evaluate the menu item's height in the main menu:
   Dim lHeight          As Long
   lHeight = 0
   For nI = 1 To clsPopupMenu.Count
      ' Check if item is in the main menu:
      If (clsPopupMenu.hMenu(nI) = clsPopupMenu.hMenu(1)) Then
         ' Add the item:
         lHeight = lHeight + clsPopupMenu.MenuItemHeight(nI)
      End If
   Next

   ' We use a PictureBox to hold the side logo here for convenience,
   ' however, you could use CreateCompatibleDC and CreateCompatibleBitmap
   ' to create a memory DC to hold this to avoid having the extra control.
   picSideBar.Height = lHeight * Screen.TwipsPerPixelY
   picSideBar.Width = 250

   ' Draw a gradient into it.  Here I stole the code directly from the
   ' SideLogo/Fonts at any angle project for simplicity:
   Dim c                As New class_Logo
   With c
      .DrawingObject = picSideBar
      .StartColor = vbBlue
      .EndColor = vbBlack
      .Caption = "VBDiamond"
      .Draw
   End With

   nIndex = clsPopupMenu.ShowPopupMenu(nX, nY)

   If (nIndex > 0) Then
      ' *** Item selected
      'Call TreatToolbar(clsPopupMenu.ItemKey(nIndex))

   End If
   Set clsPopupMenu = Nothing

End Sub

Private Sub DisplayTranslateGrid(nX As Single, nY As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/04/99
   ' * Time             : 13:15
   ' * Module Name      : frmExtractString
   ' * Module Filename  : ExtractString.frm
   ' * Procedure Name   : DisplayTranslateGrid
   ' * Parameters       :
   ' *                    nX As Single
   ' *                    nY As Single
   ' **********************************************************************
   ' * Comments         : Display a popupmenu
   ' *
   ' *
   ' **********************************************************************

   Dim nIndex           As Long
   Dim nI               As Integer
   Dim nItem            As Long

   Set clsPopupMenu = New cPopupMenu

   ' *** Initialise the Image List:
   clsPopupMenu.ImageList = ImageList1
   clsPopupMenu.GradientHighlight = True

   ' *** Initialise the hWndOwner (you must do this before showing a menu):
   clsPopupMenu.hWndOwner = Me.hWnd

   'nIconIndex = ImageToolbar.ListImages(Toolbar1.Buttons(nI).Image).Index - 1
   'nItem = clsPopupMenu.AddItem(Toolbar1.Buttons(nI).ToolTipText, , , , nIconIndex, , , Toolbar1.Buttons(nI).Key)
   'clsPopupMenu.OwnerDraw(nItem) = True

   nItem = clsPopupMenu.AddItem("-")
   clsPopupMenu.OwnerDraw(nItem) = True

   ' Firstly, evaluate the menu item's height in the main menu:
   Dim lHeight          As Long
   lHeight = 0
   For nI = 1 To clsPopupMenu.Count
      ' Check if item is in the main menu:
      If (clsPopupMenu.hMenu(nI) = clsPopupMenu.hMenu(1)) Then
         ' Add the item:
         lHeight = lHeight + clsPopupMenu.MenuItemHeight(nI)
      End If
   Next

   ' We use a PictureBox to hold the side logo here for convenience,
   ' however, you could use CreateCompatibleDC and CreateCompatibleBitmap
   ' to create a memory DC to hold this to avoid having the extra control.
   picSideBar.Height = lHeight * Screen.TwipsPerPixelY
   picSideBar.Width = 250

   ' Draw a gradient into it.  Here I stole the code directly from the
   ' SideLogo/Fonts at any angle project for simplicity:
   Dim c                As New class_Logo
   With c
      .DrawingObject = picSideBar
      .StartColor = vbBlue
      .EndColor = vbBlack
      .Caption = "VBDiamond"
      .Draw
   End With

   nIndex = clsPopupMenu.ShowPopupMenu(nX, nY)

   If (nIndex > 0) Then
      ' *** Item selected
      'Call TreatToolbar(clsPopupMenu.ItemKey(nIndex))

   End If
   Set clsPopupMenu = Nothing

End Sub

Private Sub clsPopupMenu_DrawItem(ByVal hdc As Long, ByVal lMenuIndex As Long, lLeft As Long, lTop As Long, lRight As Long, lBottom As Long, ByVal bSelected As Boolean, ByVal bChecked As Boolean, ByVal bDisabled As Boolean, bDoDefault As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 20/09/99
   ' * Time             : 12:50
   ' * Module Name      : frmVBCodeDatabase
   ' * Module Filename  : VBIDEDatabase.frm
   ' * Procedure Name   : clsPopupMenu_DrawItem
   ' * Parameters       :
   ' *                    ByVal hdc As Long
   ' *                    ByVal lMenuIndex As Long
   ' *                    lLeft As Long
   ' *                    lTop As Long
   ' *                    lRight As Long
   ' *                    lBottom As Long
   ' *                    ByVal bSelected As Boolean
   ' *                    ByVal bChecked As Boolean
   ' *                    ByVal bDisabled As Boolean
   ' *                    bDoDefault As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lW               As Long
   ' The DrawItem event for Owner Draw menu items either allows you
   ' to draw the entire item, or just to do some new drawing then
   ' let the standard method do its stuff.  This is useful if you
   ' want to add a graphic to the left or right of the menu item.

   ' Here we draw the relevant part of the side bar
   ' logo to the left of the menu then offset the
   ' left position so the rest of the menu draws
   ' after it:
   lW = picSideBar.Width \ Screen.TwipsPerPixelX
   BitBlt hdc, lLeft, lTop, lW, lBottom - lTop, picSideBar.hdc, 0, lTop, vbSrcCopy
   lLeft = lLeft + lW + 1
   bDoDefault = True
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
