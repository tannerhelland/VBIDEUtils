VERSION 5.00
Object = "{D1DAC785-7BF2-42C1-9915-A540451B87F2}#1.1#0"; "VBIDEUtils1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDependency 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detect all the dependency files"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11295
   Icon            =   "Dependency.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGetFromCurrentProject 
      Caption         =   "Detect dependencies for current &project"
      Height          =   855
      Left            =   9480
      Picture         =   "Dependency.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Detect dependencies of the current VB project"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   7680
      Picture         =   "Dependency.frx":1274
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel the detection"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdGetDependencies 
      Caption         =   "&Detect dependencies from EXE"
      Height          =   855
      Left            =   5880
      Picture         =   "Dependency.frx":157E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click the button to detect all dependencies"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox tbFile 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Indicates the EXE/DLL/OCX to work"
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   5400
      Picture         =   "Dependency.frx":1B08
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Browse to find the needed EXE/DLL/OCX"
      Top             =   340
      Width           =   375
   End
   Begin vbAcceleratorGrid.vbalGrid grdGrid 
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Here is the list of dependencies"
      Top             =   1080
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7858
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dependency.frx":1F4A
            Key             =   "OCX"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Dependency.frx":2124
            Key             =   "DLL"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Executable to check"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   405
      Width           =   1575
   End
End
Attribute VB_Name = "frmDependency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 26/11/1999
' * Time             : 14:52
' * Module Name      : frmDependency
' * Module Filename  : Dependency.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Private WithEvents CommonDialog1 As class_CommonDialog
Attribute CommonDialog1.VB_VarHelpID = -1

Private m_ColDep        As Collection

Private Sub commonDialog1_InitDialog(ByVal hDlg As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/03/2000
   ' * Time             : 15:28
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
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

Private Sub cmdCancel_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
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

Private Sub cmdGetDependencies_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:15
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
   ' * Procedure Name   : cmdGetDependencies_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdGetDependencies_Click

   Call GetDepencies(tbFile.Text)

   Call FillGrid

EXIT_cmdGetDependencies_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdGetDependencies_Click:
   Resume EXIT_cmdGetDependencies_Click

End Sub

Private Sub cmdGetFromCurrentProject_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
   ' * Procedure Name   : cmdGetFromCurrentProject_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdGetFromCurrentProject_Click

   If VBInstance.ActiveVBProject Is Nothing Then
      Exit Sub
   End If

   Dim nI               As Long

   Dim sFile            As String

   Dim clsFileVersion   As class_FileVersion

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Turn redraw off for speed:
   grdGrid.Redraw = False

   ' *** Clear the grid
   grdGrid.Clear

   ' *** Set the number of rows
   grdGrid.Rows = VBInstance.ActiveVBProject.References.Count

   ' *** Fill the grid
   For nI = 1 To VBInstance.ActiveVBProject.References.Count
      Set clsFileVersion = New class_FileVersion

      On Error Resume Next
      clsFileVersion.FullPathName = VBInstance.ActiveVBProject.References(nI).FullPath
      If err <> 0 Then
         clsFileVersion.FullPathName = VBInstance.ActiveVBProject.References(nI).Name
      End If
      On Error GoTo ERROR_cmdGetFromCurrentProject_Click

      ' *** Add one row in the grid
      With grdGrid
         If clsFileVersion.FileType = "DLL" Then
            .CellDetails nI, 1, , DT_CENTER, ImageList1.ListImages("OCX").index - 1
         Else
            .CellDetails nI, 1, , DT_CENTER, ImageList1.ListImages("DLL").index - 1
         End If
         .CellDetails nI, 2, clsFileVersion.FullPathName, , , vbWhite, &H80&
         .CellDetails nI, 3, VBInstance.ActiveVBProject.References(nI).Major & "." & VBInstance.ActiveVBProject.References(nI).Minor, DT_MODIFYSTRING
         .CellDetails nI, 4, VBInstance.ActiveVBProject.References(nI).Description, DT_MODIFYSTRING
      End With

      Set clsFileVersion = Nothing

   Next

   For nI = 1 To grdGrid.Rows
      If grdGrid.ColumnWidth(2) < grdGrid.EvaluateTextWidth(nI, 2) + 10 Then grdGrid.ColumnWidth(2) = grdGrid.EvaluateTextWidth(nI, 2) + 10
      If grdGrid.ColumnWidth(3) < grdGrid.EvaluateTextWidth(nI, 3) + 10 Then grdGrid.ColumnWidth(3) = grdGrid.EvaluateTextWidth(nI, 3) + 10
      If grdGrid.ColumnWidth(4) < grdGrid.EvaluateTextWidth(nI, 4) + 10 Then grdGrid.ColumnWidth(4) = grdGrid.EvaluateTextWidth(nI, 4) + 10
   Next

   ' *** Turn redraw on
   grdGrid.Redraw = True

EXIT_cmdGetFromCurrentProject_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdGetFromCurrentProject_Click:
   Resume EXIT_cmdGetFromCurrentProject_Click

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Call InitTooltips

   Call InitGrid

   If Not (VBInstance.ActiveVBProject Is Nothing) Then
      tbFile.Text = VBInstance.ActiveVBProject.BuildFileName
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
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

Private Sub cmdBrowse_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/99
   ' * Time             : 11:38
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
   ' * Procedure Name   : cmdBrowse_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set CommonDialog1 = New class_CommonDialog

   CommonDialog1.DialogTitle = "Choose the executable to check"
   CommonDialog1.DefaultExt = "*.exe"
   CommonDialog1.FileName = "*.exe"
   CommonDialog1.Filter = "Executable (*.exe)|*.exe|OCX (*.ocx)|*.ocx|DLL (*.dll)|*.dll|All Files (*.*)|*.*"
   CommonDialog1.InitDir = App.Path
   CommonDialog1.Flags = OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_EXPLORER
   CommonDialog1.CancelError = False
   CommonDialog1.HookDialog = True
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then Exit Sub

   tbFile.Text = CommonDialog1.FileName

End Sub

Public Function ReplaceIt(sOriginal As String, sItem As String, sReplace As String, bReplaceAll As Boolean) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:06
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
   ' * Procedure Name   : ReplaceIt
   ' * Parameters       :
   ' *                    sOriginal As String
   ' *                    sItem As String
   ' *                    sReplace As String
   ' *                    bReplaceAll As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sLeft            As String
   Dim sRight           As String
   Dim sTmp             As String

   If InStr(sOriginal, sItem) = False Then
      ReplaceIt = sOriginal
      Exit Function
   End If
   If bReplaceAll = False Then
      sLeft = left$(sOriginal, InStr(sOriginal, sItem) - 1)
      sRight = right$(sOriginal, (Len(sOriginal) - Len(sLeft) - Len(sItem)))
      ReplaceIt = sLeft & sReplace & sRight
      Exit Function
   End If

   sTmp = sOriginal
   Do Until InStr(sTmp, sItem) = 0
      sLeft = left$(sTmp, InStr(sTmp, sItem) - 1)
      sRight = right$(sTmp, (Len(sTmp) - Len(sLeft) - Len(sItem)))
      sTmp = sLeft & sReplace & sRight
   Loop
   ReplaceIt = sTmp

End Function

Private Sub GetDepencies(sFileName As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:06
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
   ' * Procedure Name   : GetDepencies
   ' * Parameters       :
   ' *                    sFilename As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetDepencies

   Dim sAllFile         As String
   Dim nPos             As Long
   Dim nStartPos        As Long
   Dim nEndPos          As Long
   Dim nPos1            As Long
   Dim nPos2            As Long
   Dim nChar            As Integer
   Dim sFile            As String

   Dim nI               As Long

   If sFileName = "" Then Exit Sub

   Set m_ColDep = Nothing
   Set m_ColDep = New Collection

   ' *** Opens file and stores it in a string
   Open sFileName For Binary As #1
   sAllFile = Space(LOF(1))
   Get #1, , sAllFile
   Close #1

   ' *** Sets the search point at the first character
   nPos = 1
   nPos1 = InStr(nPos, sAllFile, ".dll")
   nPos2 = InStr(nPos, sAllFile, ".DLL")

   ' *** Keeps going until no more are found
   Do While (nPos1 <> 0) Or (nPos2 <> 0)
      If nPos1 < nPos2 Then
         nEndPos = nPos1
      Else
         nEndPos = nPos2
      End If
      If nPos1 = 0 Then nEndPos = nPos2
      If nEndPos = 0 Then nEndPos = nPos1

      ' *** Sets the beginning position 8 characters behind the
      ' *** end. No system files are going to be over 8 characters
      nStartPos = nEndPos - 8

      ' *** Sets the search point (not the same as starting
      ' *** position) 4 characters ahead (don't want
      ' *** to find the same file again so make it start one char
      ' *** past what we're at not
      nPos = nEndPos + 4

      ' *** This cycles through each character and if it finds a
      ' *** space or null character, it sets the starting position
      ' *** accordingly
      nI = 1
      Do While True
         nChar = Asc(Mid$(sAllFile, nEndPos - nI, 1))
         If ((nChar > Asc("z")) Or (nChar < Asc(" "))) And (nChar <> Asc("~")) Then
            nStartPos = nStartPos + (9 - nI)
            Exit Do
         End If
         nI = nI + 1
      Loop

      ' *** Gets the filename and puts it in lowercase
      sFile = LCase$(Mid$(sAllFile, nStartPos, (nEndPos - nStartPos) + 4))
      m_ColDep.Add sFile, sFile
Skip_Dll:
      nPos1 = InStr(nPos, sAllFile, ".dll")
      nPos2 = InStr(nPos, sAllFile, ".DLL")
   Loop

   ' *** Starts over and does the same for .ocx
   nPos = 1
   nPos1 = InStr(nPos, sAllFile, ".ocx")
   nPos2 = InStr(nPos, sAllFile, ".OCX")

   Do While (nPos1 <> 0) Or (nPos2 <> 0)
      If nPos1 < nPos2 Then
         nEndPos = nPos1
      Else
         nEndPos = nPos2
      End If
      If nPos1 = 0 Then nEndPos = nPos2
      If nEndPos = 0 Then nEndPos = nPos1

      nStartPos = nEndPos - 8
      nPos = nEndPos + 4

      nI = 1
      Do While True
         nChar = Asc(Mid$(sAllFile, nEndPos - nI, 1))
         If ((nChar > Asc("z")) Or (nChar < Asc(" "))) And (nChar <> Asc("~")) Then
            nStartPos = nStartPos + (9 - nI)
            Exit Do
         End If
         nI = nI + 1
      Loop

      sFile = LCase$(Mid$(sAllFile, nStartPos, (nEndPos - nStartPos) + 4))
      m_ColDep.Add sFile, sFile
Skip_Ocx:
      nPos1 = InStr(nPos, sAllFile, ".ocx")
      nPos2 = InStr(nPos, sAllFile, ".OCX")
   Loop

EXIT_GetDepencies:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GetDepencies:
   If err = 457 Then Resume Next
   Resume EXIT_GetDepencies

End Sub

Private Sub InitGrid()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/02/99
   ' * Time             : 12:04
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
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

      .ImageList = ImageList1.hImageList
      .AddColumn "Icone", "", ecgHdrTextALignCentre, , 20, , True, , False
      .AddColumn "File", "File", ecgHdrTextALignLeft, , 80
      .AddColumn "Version", "Version", ecgHdrTextALignLeft, , 80
      .AddColumn "Comment", "Comment", ecgHdrTextALignLeft, , 80

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
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
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

   Dim sFile            As String

   Dim clsFileInfo      As class_FileInfo
   Dim clsFileVersion   As class_FileVersion

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Turn redraw off for speed:
   grdGrid.Redraw = False

   ' *** Clear the grid
   grdGrid.Clear

   ' *** Set the number of rows
   grdGrid.Rows = m_ColDep.Count

   ' *** Fill the grid
   For nI = 1 To m_ColDep.Count
      Set clsFileInfo = New class_FileInfo

      sFile = GetFile(m_ColDep(nI))
      If sFile = "" Then sFile = m_ColDep(nI)

      If InStr(sFile, ":") = 0 Then
         If FileExist(GetSystemDirectory() & sFile) Then
            sFile = GetSystemDirectory() & sFile
         End If
      End If
      clsFileInfo.FullPathName = sFile
      sFile = LCase$(clsFileInfo.FullPathName)

      Set clsFileVersion = New class_FileVersion
      clsFileVersion.FullPathName = clsFileInfo.FullPathName

      ' *** Add one row in the grid
      With grdGrid
         If clsFileVersion.FileType = "DLL" Then
            .CellDetails nI, 1, , DT_CENTER, ImageList1.ListImages("DLL").index - 1
         Else
            .CellDetails nI, 1, , DT_CENTER, ImageList1.ListImages("OCX").index - 1
         End If
         .CellDetails nI, 2, clsFileInfo.FullPathName, , , vbWhite, &H80&
         .CellDetails nI, 3, clsFileVersion.ProductVersion, DT_MODIFYSTRING
         .CellDetails nI, 4, clsFileVersion.FileDescription, DT_MODIFYSTRING
      End With

      Set clsFileVersion = Nothing
      Set clsFileInfo = Nothing

   Next

   For nI = 1 To grdGrid.Rows
      If grdGrid.ColumnWidth(2) < grdGrid.EvaluateTextWidth(nI, 2) + 10 Then grdGrid.ColumnWidth(2) = grdGrid.EvaluateTextWidth(nI, 2) + 10
      If grdGrid.ColumnWidth(3) < grdGrid.EvaluateTextWidth(nI, 3) + 10 Then grdGrid.ColumnWidth(3) = grdGrid.EvaluateTextWidth(nI, 3) + 10
      If grdGrid.ColumnWidth(4) < grdGrid.EvaluateTextWidth(nI, 4) + 10 Then grdGrid.ColumnWidth(4) = grdGrid.EvaluateTextWidth(nI, 4) + 10
   Next

   ' *** Turn redraw on
   grdGrid.Redraw = True

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_FillGrid:

   Exit Sub

End Sub

Private Sub grdGrid_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 29/11/1999
   ' * Time             : 11:15
   ' * Module Name      : frmDependency
   ' * Module Filename  : Dependency.frm
   ' * Procedure Name   : grdGrid_DblClick
   ' * Parameters       :
   ' *                    ByVal lRow As Long
   ' *                    ByVal lCol As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tbFile.Text = grdGrid.CellText(lRow, 2)

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
