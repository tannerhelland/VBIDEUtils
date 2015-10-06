VERSION 5.00
Object = "{D1DAC785-7BF2-42C1-9915-A540451B87F2}#1.1#0"; "VBIDEUtils1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCodeAnalyse 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Code Analysis"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5610
   Icon            =   "CodeAnalyse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefreshProcedure 
      Height          =   330
      Left            =   330
      Picture         =   "CodeAnalyse.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Analyse all the procedure/functions...."
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdCollapse 
      Height          =   330
      Left            =   1050
      Picture         =   "CodeAnalyse.frx":2AE4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Collapse all the structure"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.Frame frame 
      Caption         =   "Results of the analysis"
      ForeColor       =   &H8000000D&
      Height          =   5415
      Index           =   4
      Left            =   840
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      Begin vbAcceleratorGrid.vbalGrid grdGrid 
         Height          =   2535
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "List of controls and their accelerators"
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4471
         RowMode         =   -1  'True
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
      End
   End
   Begin VB.Frame frame 
      Caption         =   "Analysis of procedures/functions"
      ForeColor       =   &H8000000D&
      Height          =   5415
      Index           =   3
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   5055
      Begin MSComctlLib.TreeView tvUnusedProcedure 
         Height          =   5055
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Show all procedure/functions... Dead one have no childs"
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8916
         _Version        =   393217
         Indentation     =   26
         LabelEdit       =   1
         LineStyle       =   1
         PathSeparator   =   "|"
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Frame frame 
      Caption         =   "Dead Variables"
      ForeColor       =   &H8000000D&
      Height          =   5415
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   5055
      Begin MSComctlLib.TreeView tvUnusedVariables 
         Height          =   5055
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Show all the unused and dead code...."
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8916
         _Version        =   393217
         Indentation     =   26
         LabelEdit       =   1
         LineStyle       =   1
         PathSeparator   =   "|"
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   330
      Left            =   0
      Picture         =   "CodeAnalyse.frx":2C2E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Refresh the analysis"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.Frame frame 
      Caption         =   "Project structure"
      ForeColor       =   &H8000000D&
      Height          =   5415
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   5055
      Begin MSComctlLib.TreeView tvResults 
         Height          =   5055
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Project structure"
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8916
         _Version        =   393217
         Indentation     =   26
         LabelEdit       =   1
         LineStyle       =   1
         PathSeparator   =   "|"
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdExpand 
      Height          =   330
      Left            =   720
      Picture         =   "CodeAnalyse.frx":2D30
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Expand all the structure"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5895
      Left            =   0
      TabIndex        =   11
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   10398
      TabWidthStyle   =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Analysis"
            Key             =   "Analysis"
            Object.ToolTipText     =   "Analysis of the project"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dead Variables"
            Key             =   "DeadVariables"
            Object.ToolTipText     =   "Display all the dead variables"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dead Procedures"
            Key             =   "DeadProcedure"
            Object.ToolTipText     =   "Display all the dead procedures/functions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Statistics"
            Key             =   "Statistics"
            Object.ToolTipText     =   "Statistics on the project"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":2E7A
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":3054
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":322E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":3548
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":3862
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":3A3C
            Key             =   "UserDocument"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":3C16
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":3DF0
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":3FCA
            Key             =   "Lines"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":42E4
            Key             =   "Procedure"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":45FE
            Key             =   "Function"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":4918
            Key             =   "Related"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":4C32
            Key             =   "OCX"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":4F4C
            Key             =   "Declaration"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":5266
            Key             =   "ActiveXDLL"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":5580
            Key             =   "ActiveXDocument"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":589A
            Key             =   "ActiveXExe"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":5BB4
            Key             =   "EXE"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":5ECE
            Key             =   "Variable"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":60A8
            Key             =   "Dead"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeAnalyse.frx":63C2
            Key             =   "Property"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCodeAnalyse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 28/10/1999
' * Time             : 11:42
' * Module Name      : frmCodeAnalyse
' * Module Filename  : CodeAnalyse.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Dim WithEvents clsProject As class_Project
Attribute clsProject.VB_VarHelpID = -1
Private colResult       As Collection
Private mvarTotalLines  As Long
Private projCollection  As Collection
Private clsAnalysis     As class_Analyze
Private mvarProjectCnt  As Long

Private sProjectName    As String
Private colProjectName  As Collection

Private bCountBlanks    As Boolean
Private bCountComments  As Boolean

Private mnCurFrame      As Integer

Private clsSearchEngine As New class_Search

Private Sub Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/11/1999
   ' * Time             : 13:16
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   bCountBlanks = True
   Set colResult = Nothing
   Set colResult = New Collection
   colResult.Add "Total Number of Lines+TOTLINES"
   colResult.Add "Total Number of Classes+TOTCLASSES"
   colResult.Add "Total Number of Forms+TOTFORMS"
   colResult.Add "Total Number of Modules+TOTMODS"
   colResult.Add "Total Number of User Controls+TOTUSERCONTROLS"
   colResult.Add "Total Number of User Documents+TOTUSERDOCS"
   colResult.Add "Total Number of Functions+TOTFUNCTIONS"
   colResult.Add "Total Number of Properties+TOTPROPS"
   colResult.Add "Total Number of Subroutines+TOTSUBS"
   colResult.Add "Smallest Project+SMALLEST"
   colResult.Add "Lines in Smallest Project+LINESINSMALLEST"
   colResult.Add "Largest Project+LARGEST"
   colResult.Add "Lines in Largest Project+LINESINLARGEST"

   colResult.Add "Average Lines per Project+AVERAGE"
   colResult.Add "Average Lines per Function+AVGFUNCTION"
   colResult.Add "Average Lines per Subroutine+AVGSUB"
   colResult.Add "Average Lines per Property+AVGPROPERTY"
   colResult.Add "Total Lines in Functions+TOTLINESFUNCTIONS"
   colResult.Add "Total Lines in Subroutines+TOTLINESSUBS"
   colResult.Add "Total Lines in Properties+TOTLINESPROPS"

   colResult.Add "Total Variables+TOTVARIABLES"
   colResult.Add "Total Unused Variables+TOTUNUSEDVARIABLES"

   Set tvResults.ImageList = ImageList1

End Sub

Private Sub LaunchAnalyse()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : LaunchAnalyse
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_LaunchAnalyse

   Dim rc               As Boolean
   Dim nI               As Long

   ' *** If we couldn't get it, quit
   If VBInstance.ActiveVBProject Is Nothing Then Exit Sub

   Set colProjectName = Nothing
   Set colProjectName = New Collection

   sProjectName = VBInstance.ActiveVBProject.FileName

   If sProjectName = "" Then
      Me.Visible = False
      Call MsgBoxTop(Me.hWnd, "You need to save the project", vbCritical, "Analyze project")
      Me.Visible = True
      Exit Sub
   End If

   For nI = 1 To VBInstance.VBProjects.Count
      colProjectName.Add VBInstance.VBProjects(nI).FileName
   Next

   Set projCollection = New Collection
   mvarTotalLines = 0

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Refresh

   frmProgress.MessageText = "Analysing procedures"
   Load frmProgress
   Me.Hide
   frmProgress.Show
   frmProgress.ZOrder
   'Call AlwaysOnTop(frmProgress, True)

   Set clsProject = New class_Project
   clsProject.CountBlanks = bCountBlanks
   clsProject.CountComments = bCountComments

   ' *** Get all unused
   tvUnusedVariables.Nodes.Clear
   Set tvUnusedVariables.ImageList = ImageList1

   Call tvResults.Nodes.Clear
   Call InitGrid

   frmProgress.Maximum = colProjectName.Count
   For nI = 1 To colProjectName.Count
      frmProgress.Progress = nI
      sProjectName = colProjectName(nI)

      Set clsProject = Nothing
      Set clsProject = New class_Project
      frmProgress.MessageText = "Processing " & sProjectName
      Refresh
      rc = clsProject.Process(sProjectName)

      If rc = True Then
         Call projCollection.Add(clsProject)
         Call UpdateTree(clsProject)
      Else
         Me.Visible = False
         MsgBox "Unable to Process " & sProjectName & Chr$(13) & "Error Message : " & clsProject.LastErrorMessage, vbCritical, "Analysis Failed"
         Me.Visible = True
      End If
   Next

   Refresh

   Call Analyze

   Unload frmProgress
   Set frmProgress = Nothing

EXIT_LaunchAnalyse:
   Me.Show
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_LaunchAnalyse:
   Resume EXIT_LaunchAnalyse

End Sub

Private Sub Analyze()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : Analyze
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set clsAnalysis = New class_Analyze

   Call clsAnalysis.Analyze(projCollection)

   Call FillGrid

End Sub

Private Function GetData(ByVal sKey As String) As Variant
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : GetData
   ' * Parameters       :
   ' *                    ByVal key As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   With clsAnalysis
      Select Case sKey
         Case "TOTLINES"
            GetData = .TotalLines
         Case "TOTCLASSES"
            GetData = .TotalClasses
         Case "TOTFORMS"
            GetData = .TotalForms
         Case "TOTMODS"
            GetData = .TotalMods
         Case "TOTUSERCONTROLS"
            GetData = .TotalUserControls
         Case "TOTUSERDOCS"
            GetData = .TotalUserDocuments
         Case "TOTFUNCTIONS"
            GetData = .TotalFunctions
         Case "TOTPROPS"
            GetData = .TotalProperties
         Case "TOTSUBS"
            GetData = .TotalSubs
         Case "SMALLEST"
            GetData = .SmallestProject
         Case "LARGEST"
            GetData = .LargestProject
         Case "AVERAGE"
            GetData = .AverageProject
         Case "AVGFUNCTION"
            GetData = .AverageFunction
         Case "AVGSUB"
            GetData = .AverageSub
         Case "AVGPROPERTY"
            GetData = .AverageProperty
         Case "TOTLINESFUNCTIONS"
            GetData = .LinesInFunctions
         Case "TOTLINESSUBS"
            GetData = .LinesInSubroutines
         Case "TOTLINESPROPS"
            GetData = .LinesInProperties
         Case "LINESINLARGEST"
            GetData = .LinesInLargest
         Case "LINESINSMALLEST"
            GetData = .LinesInSmallest
         Case "TOTVARIABLES"
            GetData = .TotalVariables
         Case "TOTUNUSEDVARIABLES"
            GetData = .TotalUnusedVariables

      End Select

   End With
End Function

Private Sub UpdateTree(clsProject As class_Project)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : UpdateTree
   ' * Parameters       :
   ' *                    clsProject As class_Project
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim clsClass         As class_Class
   Dim clsSub           As class_Sub
   Dim tvNode           As Node
   Dim aNode            As Node
   Dim aVariable        As Node
   Dim clsNode          As Node
   Dim frmNode          As Node
   Dim modNode          As Node
   Dim ctlNode          As Node
   Dim docNode          As Node
   Dim resNode          As Node
   Dim tmpNode          As Node
   Dim nClass           As Long
   Dim nForm            As Long
   Dim nModule          As Long
   Dim nUserDocument    As Long
   Dim nControl         As Long
   Dim nResource        As Long
   Dim sKey             As String
   Dim nI               As Integer

   On Error GoTo ERROR_UpdateTree

   mvarProjectCnt = mvarProjectCnt + 1
   Call tvResults.Nodes.Add(, , clsProject.ProjectName, clsProject.ProjectName, "Project")
   Call tvUnusedVariables.Nodes.Add(, , clsProject.ProjectName, clsProject.ProjectName, "Project")

   sKey = ImageList1.ListImages("EXE").Key

   Select Case UCase$(clsProject.ProjectType)
      Case "EXE":
         sKey = ImageList1.ListImages("EXE").Key
      Case "CONTROL":
         sKey = ImageList1.ListImages("OCX").Key
      Case "OLEDLL":
         sKey = ImageList1.ListImages("ActiveXDLL").Key
      Case "OLEEXE":
         sKey = ImageList1.ListImages("ActiveXExe").Key
   End Select

   Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Project Type : " & clsProject.ProjectType, sKey)
   Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Project Path : " & clsProject.ProjectPath, "Path")
   Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Version : " & clsProject.msMajorVer & "." & clsProject.msMinorVer & "." & clsProject.msRevisionVer, "Property")
   Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Title : " & clsProject.msTitle, "Property")

   If Trim$(clsProject.msDescription) <> "" Then Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Description : " & clsProject.msDescription, "Property")
   If Trim$(clsProject.msVersionComments) <> "" Then Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Comments : " & clsProject.msVersionComments, "Property")
   If Trim$(clsProject.msVersionCompanyName) <> "" Then Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Company Name : " & clsProject.msVersionCompanyName, "Property")
   If Trim$(clsProject.msVersionLegalTrademarks) <> "" Then Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Legal Trademarks : " & clsProject.msVersionLegalTrademarks, "Property")
   Call tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, , "Project Total Lines : " & clsProject.TotalLines, "Lines")

   Set clsNode = tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, clsProject.ProjectName & "CLASSES", "Classes", "Class")
   Set frmNode = tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, clsProject.ProjectName & "FORMS", "Forms", "Form")
   Set modNode = tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, clsProject.ProjectName & "MODULES", "Modules", "Module")
   Set ctlNode = tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, clsProject.ProjectName & "USERCONTROL", "Controls", "OCX")
   Set docNode = tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, clsProject.ProjectName & "Form", "Form", "Form")
   Set resNode = tvResults.Nodes.Add(clsProject.ProjectName, tvwChild, clsProject.ProjectName & "RESOURCE", "Resource", "Related")

   For Each clsClass In clsProject.Classes
      Select Case clsClass.ClassType
         Case "Class":
            Set tvNode = tvResults.Nodes.Add(clsProject.ProjectName & "CLASSES", tvwChild, , clsClass.ClassName, "Close")
            tvNode.ExpandedImage = "Open"
            nClass = nClass + 1

         Case "Form":
            Set tvNode = tvResults.Nodes.Add(clsProject.ProjectName & "FORMS", tvwChild, , clsClass.ClassName, "Close")
            tvNode.ExpandedImage = "Open"
            nForm = nForm + 1

         Case "Module":
            Set tvNode = tvResults.Nodes.Add(clsProject.ProjectName & "MODULES", tvwChild, , clsClass.ClassName, "Close")
            tvNode.ExpandedImage = "Open"
            nModule = nModule + 1

         Case "UserControl":
            Set tvNode = tvResults.Nodes.Add(clsProject.ProjectName & "USERCONTROL", tvwChild, , clsClass.ClassName, "Close")
            tvNode.ExpandedImage = "Open"
            nControl = nControl + 1

         Case "Form":
            Set tvNode = tvResults.Nodes.Add(clsProject.ProjectName & "Form", tvwChild, , clsClass.ClassName, "Close")
            tvNode.ExpandedImage = "Open"
            nUserDocument = nUserDocument + 1

         Case "ResFile32":
            Set tvNode = tvResults.Nodes.Add(clsProject.ProjectName & "RESOURCE", tvwChild, , clsClass.ClassName, "Close")
            tvNode.ExpandedImage = "Open"
            nResource = nResource + 1
            GoTo NEXT_TRIP
      End Select

      ' Call tvResults.Nodes.Add(tvNode.index, tvwChild, , "Number of Lines : " & clsClass.NumLines, "Lines")

      Call tvResults.Nodes.Add(tvNode.index, tvwChild, clsClass.ClassName & "_DECLARATIONS_" & mvarProjectCnt, "Declarations", "Close")
      tvResults.Nodes(clsClass.ClassName & "_DECLARATIONS_" & mvarProjectCnt).ExpandedImage = "Open"

      Call tvResults.Nodes.Add(tvNode.index, tvwChild, clsClass.ClassName & "_FUNCTIONS_" & mvarProjectCnt, "Functions", "Close")
      tvResults.Nodes(clsClass.ClassName & "_FUNCTIONS_" & mvarProjectCnt).ExpandedImage = "Open"

      Call tvResults.Nodes.Add(tvNode.index, tvwChild, clsClass.ClassName & "_SUBS_" & mvarProjectCnt, "Subroutines", "Close")
      tvResults.Nodes(clsClass.ClassName & "_SUBS_" & mvarProjectCnt).ExpandedImage = "Open"

      'If clsClass.ClassType <> "Module" Then
      Call tvResults.Nodes.Add(tvNode.index, tvwChild, clsClass.ClassName & "_PROPERTIES_" & mvarProjectCnt, "Properties", "Close")
      tvResults.Nodes(clsClass.ClassName & "_DECLARATIONS_" & mvarProjectCnt).ExpandedImage = "Open"
      'End If

      For Each clsSub In clsClass.Subs
         Select Case clsSub.SubType
            Case "Function"
               Set aNode = tvResults.Nodes.Add(clsClass.ClassName & "_FUNCTIONS_" & mvarProjectCnt, tvwChild, , clsSub.SubName, "Function")
            Case "Declarations"
               Set aNode = tvResults.Nodes.Add(clsClass.ClassName & "_DECLARATIONS_" & mvarProjectCnt, tvwChild, , "Variable Declarations", "Declaration")
            Case "Property"
               Set aNode = tvResults.Nodes.Add(clsClass.ClassName & "_PROPERTIES_" & mvarProjectCnt, tvwChild, , clsSub.SubName, "Property")
            Case "Sub"
               Set aNode = tvResults.Nodes.Add(clsClass.ClassName & "_SUBS_" & mvarProjectCnt, tvwChild, , clsSub.SubName, "Procedure")
         End Select

         ' *** Add variables
         If clsSub.mcolVariable.Count > 0 Then
            Set aVariable = tvResults.Nodes.Add(aNode.index, tvwChild, , "Variables", "Variable")
            For nI = 1 To clsSub.mcolVariable.Count
               Call tvResults.Nodes.Add(aVariable.index, tvwChild, , clsSub.mcolVariable(nI), "Variable")
            Next
         End If

         ' *** Add unused variables
         If clsSub.mcolUnusedVar.Count > 0 Then
            Set tmpNode = tvUnusedVariables.Nodes.Add(clsProject.ProjectName, tvwChild, clsClass.ClassName, clsClass.ClassName, "Close")
            Set tmpNode = tvUnusedVariables.Nodes.Add(clsClass.ClassName, tvwChild, , aNode.Text, aNode.Image)

            Set aVariable = tvResults.Nodes.Add(aNode.index, tvwChild, , "Unused variables", "Variable")
            For nI = 1 To clsSub.mcolUnusedVar.Count
               Call tvResults.Nodes.Add(aVariable.index, tvwChild, , clsSub.mcolUnusedVar(nI), "Dead")
               Call tvUnusedVariables.Nodes.Add(tmpNode.index, tvwChild, , clsSub.mcolUnusedVar(nI), "Dead")
            Next
            Set tmpNode = Nothing
         End If

         'Call tvResults.Nodes.Add(aNode.index, tvwChild, , "Number of Lines : " & clsSub.NumLines, "Lines")

      Next

      ' *** Remove all empty
      On Error Resume Next
      If tvResults.Nodes(clsClass.ClassName & "_DECLARATIONS_" & mvarProjectCnt).Children = 0 Then tvResults.Nodes.Remove (clsClass.ClassName & "_DECLARATIONS_" & mvarProjectCnt)
      If tvResults.Nodes(clsClass.ClassName & "_FUNCTIONS_" & mvarProjectCnt).Children = 0 Then tvResults.Nodes.Remove (clsClass.ClassName & "_FUNCTIONS_" & mvarProjectCnt)
      If tvResults.Nodes(clsClass.ClassName & "_SUBS_" & mvarProjectCnt).Children = 0 Then tvResults.Nodes.Remove (clsClass.ClassName & "_SUBS_" & mvarProjectCnt)
      If tvResults.Nodes(clsClass.ClassName & "_SUBS_" & mvarProjectCnt).Children = 0 Then tvResults.Nodes.Remove (clsClass.ClassName & "_SUBS_" & mvarProjectCnt)
      If tvResults.Nodes(clsClass.ClassName & "_PROPERTIES_" & mvarProjectCnt).Children = 0 Then tvResults.Nodes.Remove (clsClass.ClassName & "_PROPERTIES_" & mvarProjectCnt)
      On Error GoTo ERROR_UpdateTree

NEXT_TRIP:
   Next

   If nClass > 0 Then
      clsNode.Text = "Classes (" & nClass & ")"
   Else
      Call tvResults.Nodes.Remove(clsNode.Key)
   End If

   If nForm > 0 Then
      frmNode.Text = "Forms (" & nForm & ")"
   Else
      Call tvResults.Nodes.Remove(frmNode.Key)
   End If

   If nModule > 0 Then
      modNode.Text = "Modules (" & nModule & ")"
   Else
      Call tvResults.Nodes.Remove(modNode.Key)
   End If

   If nControl > 0 Then
      ctlNode.Text = "Controls (" & nControl & ")"
   Else
      Call tvResults.Nodes.Remove(ctlNode.Key)
   End If

   If nUserDocument > 0 Then
      docNode.Text = "Form (" & nUserDocument & ")"
   Else
      Call tvResults.Nodes.Remove(docNode.Key)
   End If

   If nResource > 0 Then
      resNode.Text = "Resource (" & nResource & ")"
   Else
      Call tvResults.Nodes.Remove(resNode.Key)
   End If

   Exit Sub

ERROR_UpdateTree:
   Debug.Print Error
   Resume Next
   Resume

End Sub

Private Sub cmdCollapse_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 13:45
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : cmdCollapse_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Dim nI               As Integer

   If TabStrip.SelectedItem.Key = "Analysis" Then
      tvResults.Visible = False
      For nI = 1 To tvResults.Nodes.Count
         tvResults.Nodes(nI).Expanded = False
      Next
      tvResults.Visible = True

   ElseIf TabStrip.SelectedItem.Key = "DeadVariables" Then
      tvUnusedVariables.Visible = False
      For nI = 1 To tvUnusedVariables.Nodes.Count
         tvUnusedVariables.Nodes(nI).Expanded = False
      Next
      tvUnusedVariables.Visible = True

   ElseIf TabStrip.SelectedItem.Key = "DeadProcedure" Then
      tvUnusedProcedure.Visible = False
      For nI = 1 To tvUnusedProcedure.Nodes.Count
         tvUnusedProcedure.Nodes(nI).Expanded = False
      Next
      tvUnusedProcedure.Visible = True

   End If

End Sub

Private Sub cmdExpand_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 13:45
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : cmdExpand_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Dim nI               As Integer

   If TabStrip.SelectedItem.Key = "Analysis" Then
      tvResults.Visible = False
      For nI = 1 To tvResults.Nodes.Count
         tvResults.Nodes(nI).Expanded = True
      Next
      tvResults.Visible = True

   ElseIf TabStrip.SelectedItem.Key = "DeadVariables" Then
      tvUnusedVariables.Visible = False
      For nI = 1 To tvUnusedVariables.Nodes.Count
         tvUnusedVariables.Nodes(nI).Expanded = True
      Next
      tvUnusedVariables.Visible = True

   ElseIf TabStrip.SelectedItem.Key = "DeadProcedure" Then
      tvUnusedProcedure.Visible = False
      For nI = 1 To tvUnusedProcedure.Nodes.Count
         tvUnusedProcedure.Nodes(nI).Expanded = True
      Next
      tvUnusedProcedure.Visible = True

   End If

End Sub

Public Sub cmdRefresh_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/11/1999
   ' * Time             : 16:53
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : cmdRefresh_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   LaunchAnalyse

End Sub

Private Sub cmdRefreshProcedure_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 23/02/2000
   ' * Time             : 12:42
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : cmdRefreshProcedure_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdRefreshProcedure_Click

   Dim nI               As Integer
   Dim sTmp             As String
   Dim sTmp2            As String
   Dim sCallers         As String
   Dim sPart2           As String

   Dim sMember          As String

   Dim nCount           As Integer

   Dim nCallers         As Long

   Dim tvNode           As Node
   Dim tvProcedure      As Node
   Dim tvCaller         As Node

   Dim vbComponentObj   As VBComponent
   Dim vbMemberObj      As Member

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Me.Enabled = False
   frmProgress.MessageText = "Analysing procedures"
   Load frmProgress
   Me.Hide
   frmProgress.Show
   frmProgress.ZOrder
   'Call AlwaysOnTop(frmProgress, True)

   Set TabStrip.SelectedItem = TabStrip.Tabs("Analysis")
   tvUnusedProcedure.Visible = False

   ' *** Get all unused
   tvUnusedProcedure.Nodes.Clear
   Set tvUnusedProcedure.ImageList = ImageList1

   Set clsSearchEngine = Nothing
   Set clsSearchEngine = New class_Search

   ' *** Get all the functions name
   clsSearchEngine.ScanForFunctionNames

   ' *** Scan again to locate uses of each function
   clsSearchEngine.ScanForFunctionUse

   frmProgress.MessageText = "Phase 3/3"
   frmProgress.Maximum = VBInstance.ActiveVBProject.VBComponents.Count
   frmProgress.Progress = 1
   nCount = 1

   ' *** Get all procedures...
   For Each vbComponentObj In VBInstance.ActiveVBProject.VBComponents
      frmProgress.MessageText = "Phase 3/3" & vbCrLf & vbComponentObj.Name
      DoEvents
      frmProgress.Progress = nCount
      nCount = nCount + 1

      For Each vbMemberObj In vbComponentObj.CodeModule.members
         ' *** The member type tells us if this is a function or a variable
         Select Case vbMemberObj.Type
            Case vbext_mt_Method, vbext_mt_Event, vbext_mt_Property
               ' *** Get the list of callers
               If vbMemberObj.Type <> vbext_mt_Event Then
                  sMember = LCase$(vbMemberObj.Name)
                  If Not (sMember Like "*_load") And _
                     Not (sMember Like "*_unload") And _
                     Not (sMember Like "*_resize") And _
                     Not (sMember Like "*_click") And _
                     Not (sMember Like "*_dblclick") And _
                     Not (sMember Like "*_buttonclick") And _
                     Not (sMember Like "*_change") And _
                     Not (sMember Like "*_key*") And _
                     Not (sMember Like "*_mouse*") And _
                     Not (sMember Like "*_link*") And _
                     Not (sMember Like "*_ole*") And _
                     Not (sMember Like "*_drag*") And _
                     Not (sMember Like "*_lostfocus") And _
                     Not (sMember Like "*_gotfocus") And _
                     Not (sMember Like "*_terminate") And _
                     Not (sMember Like "*_initialize") And _
                     Not (sMember Like "*_activate") Then
                     'Not (sMember Like "*_timer") And
                     sCallers = clsSearchEngine.BuildMenu(vbComponentObj.Name & "!" & vbMemberObj.Name)
                     If left$(sCallers, 1) = "@" Then sCallers = Mid$(sCallers, 2)

                     Set tvProcedure = tvUnusedProcedure.Nodes.Add(, , vbMemberObj.Name, vbMemberObj.Name, "Close")
                     If LenB(sCallers) > 0 Then tvProcedure.ExpandedImage = "Open"
                     tvProcedure.Tag = VBInstance.ActiveVBProject.Name & "|" & vbComponentObj.Name & "|" & vbMemberObj.Name

                     nCallers = CountTokens(sCallers, "@")

                     If nCallers < 100 Then
                        For nI = 1 To nCallers
                           sTmp = GetToken(sCallers, "@", nI)

                           ' *** Get the part 2
                           nI = nI + 1
                           sPart2 = GetToken(sCallers, "@", nI)

                           ' *** Get the caller
                           sTmp2 = GetToken(sPart2, "~", 1)

                           ' *** Add this caller in the treeview
                           On Error Resume Next
                           Set tvCaller = tvUnusedProcedure.Nodes.Add(tvProcedure.index, tvwChild, sTmp & "|" & sTmp2, sTmp2, "Procedure")
                           tvCaller.Tag = sPart2
                           If err <> 0 Then err.Clear
                        Next
                     End If

                  End If
               End If

            Case vbext_mt_Variable
               ' Debug.Print vbTab & vbMemberObj.Name & vbTab & "Variable"
               Call tvUnusedProcedure.Nodes.Add(tvNode.index, tvwChild, , vbMemberObj.Name, "Declaration")
            Case vbext_mt_Const
               ' Debug.Print vbTab & vbMemberObj.Name & vbTab & "Constant"
               Call tvUnusedProcedure.Nodes.Add(tvNode.index, tvwChild, , vbMemberObj.Name, "Declaration")
         End Select

      Next
   Next

EXIT_cmdRefreshProcedure_Click:
   On Error Resume Next

   Unload frmProgress
   Set frmProgress = Nothing

   Me.Enabled = True
   tvUnusedProcedure.Visible = True
   tvUnusedProcedure.Sorted = True
   tvUnusedProcedure.Refresh

   Set TabStrip.SelectedItem = TabStrip.Tabs("DeadProcedure")

   Me.Show
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdRefreshProcedure_Click:
   Resume Next
   'Resume EXIT_cmdRefreshProcedure_Click
   Resume

End Sub

Private Sub tabstrip_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/11/1999
   ' * Time             : 14:59
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : tabstrip_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   If TabStrip.SelectedItem.index = mnCurFrame Then Exit Sub       ' No need to change fra

   ' *** Otherwise, hide old frame, show new.
   Frame(TabStrip.SelectedItem.index).Visible = True
   Frame(TabStrip.SelectedItem.index).ZOrder
   Frame(mnCurFrame).Visible = False

   ' *** Set mnCurFrame to new value.
   mnCurFrame = TabStrip.SelectedItem.index

End Sub

Private Sub tvResults_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 13:45
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : tvResults_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tvResults.SelectedItem.Expanded = Not tvResults.SelectedItem.Expanded

   ' *** Search the string
   Call SearchString(tvResults.SelectedItem.FullPath)

End Sub

Private Sub tvUnusedVariables_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 13:45
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : tvUnusedVariables_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tvUnusedVariables.SelectedItem.Expanded = Not tvUnusedVariables.SelectedItem.Expanded

   ' *** Search the string
   Call SearchString(tvUnusedVariables.SelectedItem.FullPath)

End Sub

Private Sub tvUnusedProcedure_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 13:45
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : tvUnusedProcedure_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tvUnusedProcedure.SelectedItem.Expanded = Not tvUnusedProcedure.SelectedItem.Expanded

   ' *** Search the string
   Call SearchString(tvUnusedProcedure.SelectedItem.Tag)

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 30/11/1999
   ' * Time             : 16:53
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   mnCurFrame = 1

   Call InitTooltips

   Call Form_Resize

   Call InitGrid
   Call Initialize

End Sub

Private Sub InitGrid()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/02/99
   ' * Time             : 12:04
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : InitGrid
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Init the grid
   ' *
   ' *
   ' **********************************************************************

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   On Error Resume Next

   ' *** Init the grid
   With grdGrid
      ' *** Turn redraw off for speed:
      .Redraw = False

      .BackColor = &H80000018

      .AddColumn "Title", "Title", ecgHdrTextALignLeft, , 170
      .AddColumn "Details", "Details", ecgHdrTextALignRight, , 80

      .SetHeaders

      ' *** Ensure the grid will draw!
      .Redraw = True

   End With

End Sub

Private Sub FillGrid()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/02/99
   ' * Time             : 12:05
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : FillGrid
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Fill the Grid with the values
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_FillGrid

   Dim nI               As Long

   Dim sTWA()           As String

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Turn redraw off for speed:
   grdGrid.Redraw = False

   ' *** Clear the grid
   grdGrid.Clear

   ' *** Set the number of rows
   grdGrid.Rows = colResult.Count

   ' *** Fill the grid
   For nI = 1 To grdGrid.Rows
      Call Main_Module.Split(colResult(nI), sTWA, "+")

      ' *** Add one row in the grid
      With grdGrid
         .CellDetails nI, 1, sTWA(1), , , vbWhite, &H80&
         .CellDetails nI, 2, GetData(sTWA(2)), DT_RIGHT, , &H80000018
      End With
   Next

   'For nI = 0 To grdGrid.Rows
   '   If grdGrid.ColumnWidth(1) < grdGrid.EvaluateTextWidth(nI, 1) + 10 Then grdGrid.ColumnWidth(1) = grdGrid.EvaluateTextWidth(nI, 1) + 10
   '   If grdGrid.ColumnWidth(2) < grdGrid.EvaluateTextWidth(nI, 2) + 10 Then grdGrid.ColumnWidth(2) = grdGrid.EvaluateTextWidth(nI, 2) + 10
   'Next

   ' *** Turn redraw on
   grdGrid.Redraw = True

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_FillGrid:

   Exit Sub

End Sub

Private Sub SearchString(sFullPath As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 13:45
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : SearchString
   ' * Parameters       :
   ' *                    sFullPath As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nPos             As Integer
   Dim nPos1            As Integer
   Dim aComponent       As VBComponent

   Dim projet           As VBProject

   Dim sTWA()           As String

   Dim sTmp2            As String

   Dim nStart           As Long
   Dim nLast            As Long

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Remove all uneeded
   sFullPath = Replace(sFullPath, "|Subroutines|", "|")
   sFullPath = Replace(sFullPath, "|Functions|", "|")
   sFullPath = Replace(sFullPath, "|Lines|", "|")
   sFullPath = Replace(sFullPath, "|Declarations|", "|")
   sFullPath = Replace(sFullPath, "|Properties|", "|")
   sFullPath = Replace(sFullPath, "|Functions|", "|")

   nPos = 99
   Do While nPos > 0
      nPos = InStr(sFullPath, "(")
      If nPos > 0 Then
         nPos1 = InStr(sFullPath, ")")
         sFullPath = left$(sFullPath, nPos - 1) & Mid$(sFullPath, nPos1 + 1)
      End If
   Loop
   sFullPath = Replace(sFullPath, "|Forms |", "|")
   sFullPath = Replace(sFullPath, "|Classes |", "|")
   sFullPath = Replace(sFullPath, "|Modules |", "|")
   sFullPath = Replace(sFullPath, "|Controls |", "|")
   sFullPath = Replace(sFullPath, "|UserDocument |", "|")
   sFullPath = Replace(sFullPath, "|Resource |", "|")
   sFullPath = Replace(sFullPath, "|Variable Declarations|", "|")
   sFullPath = Replace(sFullPath, "|Unused variables|", "|")

   Dim rc               As Boolean
   Dim nI               As Long

   ' *** Get all items
   Call Main_Module.Split(sFullPath, sTWA, "|")

   ' *** Search the project
   If sTWA(1) = "" Then Exit Sub
   If VBInstance.ActiveVBProject Is Nothing Then
      Me.Visible = False
      Call MsgBoxTop(Me.hWnd, "Could not identify current project", vbExclamation + vbOKOnly + vbDefaultButton1, "Indentify the project")
      Me.Visible = True
      Exit Sub
   End If

   For nI = 1 To VBInstance.VBProjects.Count
      If VBInstance.VBProjects(nI).Name = sTWA(1) Then Exit For
   Next

   If nI > VBInstance.VBProjects.Count Then Exit Sub
   Set projet = VBInstance.VBProjects(nI)

   ' *** Get the module/form/class...
   If UBound(sTWA) < 2 Then Exit Sub
   On Error Resume Next
   sTmp2 = projet.VBComponents(sTWA(2)).Name
   If sTmp2 = "" Then
      For nI = 1 To projet.VBComponents.Count
         If projet.VBComponents(nI).Type = vbext_ct_MSForm Then
            If GetFileName(projet.VBComponents(nI).FileNames(1)) = GetFileName(sTWA(2)) Then Exit For
         ElseIf projet.VBComponents(nI).Type = vbext_ct_VBForm Then
            If GetFileName(projet.VBComponents(nI).FileNames(1)) = GetFileName(sTWA(2)) Then Exit For
         ElseIf projet.VBComponents(nI).Type = vbext_ct_StdModule Then
            If projet.VBComponents(nI).Name = sTWA(2) Then Exit For
         ElseIf projet.VBComponents(nI).Type = vbext_ct_UserControl Then
            If GetFileName(projet.VBComponents(nI).FileNames(1)) = GetFileName(sTWA(2)) Then Exit For
         ElseIf projet.VBComponents(nI).Type = vbext_ct_DocObject Then
            If GetFileName(projet.VBComponents(nI).FileNames(1)) = GetFileName(sTWA(2)) Then Exit For
         ElseIf projet.VBComponents(nI).Type = vbext_ct_ClassModule Then
            If projet.VBComponents(nI).Name = sTWA(2) Then Exit For
         End If
      Next
      If nI > projet.VBComponents.Count Then Exit Sub
      Set aComponent = projet.VBComponents(nI)
   Else
      Set aComponent = projet.VBComponents(sTWA(2))
   End If
   ' *** Activate it
   If aComponent.HasOpenDesigner = False Then aComponent.Activate

   ' *** Find the function
   If UBound(sTWA) < 3 Then Exit Sub

   ' *** Remove potential blank
   nPos = InStr(sTWA(3), " ")
   If nPos > 0 Then sTWA(3) = Mid$(sTWA(3), nPos + 1)

   nStart = 1
   If sTWA(3) <> "Variables" Then
      If aComponent.CodeModule.members(sTWA(3)).Name = sTWA(3) Then
      End If
      If err > 0 Then
         err.Clear
         Exit Sub
      End If

      On Error Resume Next
      nStart = aComponent.CodeModule.ProcBodyLine(sTWA(3), 0)
      If err = 0 Then GoTo OK_Last

      nStart = aComponent.CodeModule.ProcBodyLine(sTWA(3), 1)
      If err = 0 Then GoTo OK_Last

      nStart = aComponent.CodeModule.ProcBodyLine(sTWA(3), 2)
      If err = 0 Then GoTo OK_Last

      nStart = aComponent.CodeModule.ProcBodyLine(sTWA(3), 3)
      If err = 0 Then GoTo OK_Last

OK_Last:
      ' *** Show the function
      On Error Resume Next
      nLast = aComponent.CodeModule.ProcCountLines(sTWA(3), 0)
      If err = 0 Then GoTo OK

      nLast = aComponent.CodeModule.ProcCountLines(sTWA(3), 1)
      If err = 0 Then GoTo OK

      nLast = aComponent.CodeModule.ProcCountLines(sTWA(3), 2)
      If err = 0 Then GoTo OK

      nLast = aComponent.CodeModule.ProcCountLines(sTWA(3), 3)

   Else
      nLast = aComponent.CodeModule.CountOfDeclarationLines

   End If

OK:
   nLast = nStart + nLast

   If UBound(sTWA) < 4 Then
      ' *** Activate the function
      Call aComponent.CodeModule.CodePane.SetSelection(nStart, 1, nStart, 1)

      Exit Sub
   End If
   If (sTWA(4) = "Variables") Then
      If UBound(sTWA) < 5 Then
         ' *** Activate the function
         Call aComponent.CodeModule.CodePane.SetSelection(aComponent.CodeModule.members(sTWA(3)).CodeLocation, 1, aComponent.CodeModule.members(sTWA(3)).CodeLocation, 1)

         Exit Sub
      End If
      sTmp2 = sTWA(5)
   Else
      sTmp2 = sTWA(4)
   End If

   ' *** Find the string
   Dim nLine            As Long
   Dim nColumn          As Long
   nLine = nStart
   nColumn = 1
   sTmp2 = " " & sTmp2 & " "
   Call aComponent.CodeModule.Find(sTmp2, nLine, nColumn, nLast + 1, 1)
   If nLine = nStart Then
      sTmp2 = " " & Trim$(sTmp2)
      Call aComponent.CodeModule.Find(sTmp2, nLine, nColumn, nLast + 1, 1)
   End If

   If (nLine <> nStart) And (nColumn <> 1) Then
      ' *** Activate the function
      aComponent.CodeModule.CodePane.Show
      aComponent.CodeModule.CodePane.SetSelection nLine, nColumn + 1, nLine, nColumn + Len(Trim$(sTmp2)) + 1
   Else
      ' *** Activate the function
      Call aComponent.CodeModule.CodePane.SetSelection(aComponent.CodeModule.members(nI).CodeLocation, 1, aComponent.CodeModule.members(nI).CodeLocation, 1)

   End If

End Sub

Private Sub Form_Resize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 13:45
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : Form_Resize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Resize

   On Error Resume Next

   Dim nI               As Integer

   TabStrip.Move TabStrip.left, TabStrip.top, ScaleWidth, ScaleHeight - 100 - cmdRefresh.Height

   For nI = 1 To 4
      Frame(nI).Move TabStrip.left + 120, TabStrip.top + 360, TabStrip.Width - 240, TabStrip.Height - 360 - 120
   Next

   tvResults.Move tvResults.left, tvResults.top, Frame(1).Width - tvResults.left * 2, Frame(1).Height - tvResults.top * 2
   grdGrid.Move grdGrid.left, grdGrid.top, Frame(1).Width - grdGrid.left * 2, Frame(1).Height - grdGrid.top * 2
   tvUnusedVariables.Move tvUnusedVariables.left, tvUnusedVariables.top, Frame(1).Width - tvUnusedVariables.left * 2, Frame(1).Height - tvUnusedVariables.top * 2
   tvUnusedProcedure.Move tvUnusedProcedure.left, tvUnusedProcedure.top, Frame(1).Width - tvUnusedProcedure.left * 2, Frame(1).Height - tvUnusedProcedure.top * 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 23/02/2000
   ' * Time             : 12:42
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
   ' * Procedure Name   : Form_Unload
   ' * Parameters       :
   ' *                    Cancel As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Set clsSearchEngine = Nothing
   Set projCollection = Nothing
   Set colProjectName = Nothing
   Set colResult = Nothing
   Set clsTooltips = Nothing

End Sub

Private Sub InitTooltips(Optional Flags As TooltipFlagConstants)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/11/1999
   ' * Time             : 11:15
   ' * Module Name      : frmCodeAnalyse
   ' * Module Filename  : CodeAnalyse.frm
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
      .CreateHwnd Me.hWnd, ttfBalloon
      .MaxTipWidth = 70
      .Icon = itInfoIcon
      .Title = "VBIDEUtils"
      For nI = 0 To Me.Controls.Count
         If Trim$(Me.Controls(nI).ToolTipText) <> "" Then
            If err = 0 Then
               .AddToolHwnd Me.Controls(nI), Me.hWnd, tfTransparent, Me.Controls(nI).ToolTipText
               Me.Controls(nI).ToolTipText = ""
            Else
               err.Clear
            End If
         End If
      Next
   End With

End Sub

