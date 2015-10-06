VERSION 5.00
Begin VB.Form frmDBCreator 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DB Creator code generation"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   Icon            =   "DBCreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdADOCopyClipboard 
      Caption         =   "Generate &ADO to Clipboard"
      Height          =   735
      Left            =   2760
      Picture         =   "DBCreator.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Generate the database to the clipboard"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdDAOCopyClipboard 
      Caption         =   "Generate &DAO to Clipboard"
      Height          =   735
      Left            =   1320
      Picture         =   "DBCreator.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Generate the database to the clipboard"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   2760
      Picture         =   "DBCreator.frx":2A36
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Oops, forget it..."
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4080
      Picture         =   "DBCreator.frx":2D40
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Browse to find the database"
      Top             =   210
      Width           =   375
   End
   Begin VB.TextBox tbPassword 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Optional password to access the database"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox tbDatabase 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Indicates the database to use"
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Database"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   280
      Width           =   1095
   End
End
Attribute VB_Name = "frmDBCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' *Programmer Name  : removed
' *Web Site         : http://www.ppreview.net
' *E-Mail           : removed
' *Date             : 17/09/99
' *Time             : 11:48
' *Module Name      : frmDBCreator
' *Module Filename  : DBCreator.frm
' **********************************************************************
' *Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Private msGenerated     As String

Private WithEvents CommonDialog1 As class_CommonDialog
Attribute CommonDialog1.VB_VarHelpID = -1

Private Sub cmdADOCopyClipboard_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/99
   ' * Time             : 11:48
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : cmdDAOCopyClipboard
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdADOCopyClipboard_Click

   Dim sTmp             As String

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   msGenerated = ""

   If Trim$(tbDatabase.Text) = "" Then Exit Sub

   If FileExist(tbDatabase.Text) = False Then
      Call MsgBoxTop(Me.hWnd, "The specified database has not been found", vbExclamation + vbOKOnly + vbDefaultButton1, "Database not found")
      Exit Sub
   End If

   Call GenerateADOCode(tbDatabase.Text, tbPassword.Text)

EXIT_cmdADOCopyClipboard_Click:

   If (msGenerated <> "") Then
      ' *** Copy to the clipboard
      Clipboard.Clear
      Clipboard.SetText msGenerated, vbCFText

      Me.Visible = False

      ' *** Copy done
      sTmp = "The DBCreator function " & Chr$(13) & "has generated the code and copied it to the clipboard." & Chr$(13)
      sTmp = sTmp & "You can paste it in your code."

      Call MsgBoxTop(Me.hWnd, sTmp, vbOKOnly + vbInformation, "DBCreator Code generator")
   End If

   Call cmdCancel_Click

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdADOCopyClipboard_Click:
   'MsgBox "Error in ERROR_cmdADOCopyClipboard_Click : " & Error
   Resume EXIT_cmdADOCopyClipboard_Click

End Sub

Private Sub commonDialog1_InitDialog(ByVal hDlg As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/03/2000
   ' * Time             : 15:26
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
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

Private Sub cmdBrowse_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/99
   ' * Time             : 11:38
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : cmdBrowse_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set CommonDialog1 = New class_CommonDialog

   CommonDialog1.DialogTitle = "Choose the database to use"
   CommonDialog1.DefaultExt = "*.mdb"
   CommonDialog1.FileName = "*.mdb"
   CommonDialog1.Filter = "Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
   CommonDialog1.InitDir = App.Path
   CommonDialog1.Flags = OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_EXPLORER
   CommonDialog1.CancelError = False
   CommonDialog1.HookDialog = True
   Call AlwaysOnTop(Me, False)
   CommonDialog1.ShowOpen
   Call AlwaysOnTop(Me, True)
   If CommonDialog1.FileName = "" Then Exit Sub

   tbDatabase.Text = CommonDialog1.FileName

End Sub

Private Sub cmdDAOCopyClipboard_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/09/99
   ' * Time             : 11:48
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : cmdDAOCopyClipboard
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_cmdDAOCopyClipboard_Click

   Dim sTmp             As String

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   msGenerated = ""

   If Trim$(tbDatabase.Text) = "" Then Exit Sub

   If FileExist(tbDatabase.Text) = False Then
      Call MsgBoxTop(Me.hWnd, "The specified database has not been found", vbExclamation + vbOKOnly + vbDefaultButton1, "Database not found")
      Exit Sub
   End If

   Call GenerateCode(tbDatabase.Text, tbPassword.Text)

EXIT_cmdDAOCopyClipboard_Click:

   If (msGenerated <> "") Then
      ' *** Copy to the clipboard
      Clipboard.Clear
      Clipboard.SetText msGenerated, vbCFText

      Me.Visible = False

      ' *** Copy done
      sTmp = "The DBCreator function " & Chr$(13) & "has generated the code and copied it to the clipboard." & Chr$(13)
      sTmp = sTmp & "You can paste it in your code."

      Call MsgBoxTop(Me.hWnd, sTmp, vbOKOnly + vbInformation, "DBCreator Code generator")
   End If

   Call cmdCancel_Click

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_cmdDAOCopyClipboard_Click:
   'MsgBox "Error in ERROR_cmdDAOCopyClipboard_Click : " & Error
   Resume EXIT_cmdDAOCopyClipboard_Click

End Sub

Private Sub GenerateCode(sDatabase As String, Optional sPassword As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 13/09/1999
   ' * Time             : 14:40
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GenerateCode
   ' * Parameters       :
   ' *                    sDatabase As String
   ' *                    Optional sPassword As String
   ' **********************************************************************
   ' * Comments         : Generate the code to create the database
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GenerateCode

   Dim DB               As Database

   Dim nI               As Long

   ' *** Main declaration
   msGenerated = "Public Function GenerateDatabase(sDestDBPath as String, Optional sDestDBPassword as String) as Boolean" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtils#************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Programmer Name  : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Web Site         : http://www.ppreview.net" & vbCrLf
   msGenerated = msGenerated & "   ' * E-Mail           : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Date             : " & Date & vbCrLf
   msGenerated = msGenerated & "   ' * Time             : " & time & vbCrLf
   msGenerated = msGenerated & "   ' * Procedure Name   : GenerateDatabase" & vbCrLf
   msGenerated = msGenerated & "   ' * Parameters       :" & vbCrLf
   msGenerated = msGenerated & "   ' *                    sDestDBPath As String" & vbCrLf
   msGenerated = msGenerated & "   ' *                    Optional sDestDBPassword As String" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Comments         : Create a new database by code" & vbCrLf
   msGenerated = msGenerated & "   ' *   This Database has been created using VBIDEUtils" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Dim DB               As Database" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   On Error Goto ERROR_GenerateDatabase" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateDatabase = False" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' *** Create the database" & vbCrLf
   msGenerated = msGenerated & "   If Trim$(sDestDBPassword) <> """" Then" & vbCrLf
   msGenerated = msGenerated & "      Set DB = Workspaces(0).CreateDatabase(sDestDBPath, dbLangGeneral & "";pwd="" & sDestDBPassword)" & vbCrLf
   msGenerated = msGenerated & "   Else" & vbCrLf
   msGenerated = msGenerated & "      Set DB = Workspaces(0).CreateDatabase(sDestDBPath, dbLangGeneral)" & vbCrLf
   msGenerated = msGenerated & "   End If" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

   ' *** Open the database
   Set DB = OpenDatabase(sDatabase, False, False, ";PWD=" & sPassword)

   ' *** Call the creation of each table
   If DB.TableDefs.Count > 0 Then msGenerated = msGenerated & "   ' *** Create each table" & vbCrLf
   For nI = 0 To DB.TableDefs.Count - 1
      If left(UCase$(DB.TableDefs(nI).Name), 2) <> "MS" Then
         msGenerated = msGenerated & "   ' *** Table " & DB.TableDefs(nI).Name & "" & vbCrLf
         msGenerated = msGenerated & "   If CreateTable" & DB.TableDefs(nI).Name & "(DB) = False Then" & vbCrLf
         msGenerated = msGenerated & "      GenerateDatabase = False" & vbCrLf
         msGenerated = msGenerated & "      DB.Close" & vbCrLf
         msGenerated = msGenerated & "      Set DB = Nothing" & vbCrLf
         msGenerated = msGenerated & "      Exit Function" & vbCrLf
         msGenerated = msGenerated & "   End If" & vbCrLf
         msGenerated = msGenerated & "" & vbCrLf
      End If
   Next

   ' *** Call the creation of each querydefs
   If DB.QueryDefs.Count > 0 Then msGenerated = msGenerated & "   ' *** Create each query" & vbCrLf
   For nI = 0 To DB.QueryDefs.Count - 1
      msGenerated = msGenerated & "   ' *** Query " & DB.QueryDefs(nI).Name & "" & vbCrLf
      msGenerated = msGenerated & "   If CreateQuery" & Replace(DB.QueryDefs(nI).Name, " ", "_") & "(DB) = False Then" & vbCrLf
      msGenerated = msGenerated & "      GenerateDatabase = False" & vbCrLf
      msGenerated = msGenerated & "      DB.Close" & vbCrLf
      msGenerated = msGenerated & "      Set DB = Nothing" & vbCrLf
      msGenerated = msGenerated & "      Exit Function" & vbCrLf
      msGenerated = msGenerated & "   End If" & vbCrLf
      msGenerated = msGenerated & "" & vbCrLf
   Next

   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateDatabase = True" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "EXIT_GenerateDatabase:" & vbCrLf
   msGenerated = msGenerated & "   On Error Resume Next" & vbCrLf
   msGenerated = msGenerated & "   ' *** Close the database :)" & vbCrLf
   msGenerated = msGenerated & "   DB.Close" & vbCrLf
   msGenerated = msGenerated & "   Set DB = Nothing" & vbCrLf
   msGenerated = msGenerated & "   Exit Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "ERROR_GenerateDatabase:" & vbCrLf
   msGenerated = msGenerated & "   Select Case MsgBox(""Error "" & Err.Number & "": "" & Err.Description & vbCrLf & ""in GenerateDatabase"" & vbCrLf & ""The error occured at line: "" & Erl, vbAbortRetryIgnore + vbCritical, ""Error"")" & vbCrLf
   msGenerated = msGenerated & "      Case vbAbort" & vbCrLf
   msGenerated = msGenerated & "         Screen.MousePointer = vbDefault" & vbCrLf
   msGenerated = msGenerated & "         GenerateDatabase = False" & vbCrLf
   msGenerated = msGenerated & "         Resume EXIT_GenerateDatabase:" & vbCrLf
   msGenerated = msGenerated & "      Case vbRetry" & vbCrLf
   msGenerated = msGenerated & "         Resume" & vbCrLf
   msGenerated = msGenerated & "      Case vbIgnore" & vbCrLf
   msGenerated = msGenerated & "         Resume Next" & vbCrLf
   msGenerated = msGenerated & "   End Select" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateDatabase = False" & vbCrLf
   msGenerated = msGenerated & "   Resume EXIT_GenerateDatabase" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "End Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

   ' *** Generate code for all tables
   For nI = 0 To DB.TableDefs.Count - 1
      Call GenerateTable(DB.TableDefs(nI))
   Next

   ' *** Generate code for all tables
   For nI = 0 To DB.QueryDefs.Count - 1
      Call GenerateQuery(DB.QueryDefs(nI))
   Next

EXIT_GenerateCode:
   On Error Resume Next
   DB.Close
   Set DB = Nothing
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GenerateCode:
   'Call MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description, vbAbortRetryIgnore + vbCritical, "Error")
   Resume EXIT_GenerateCode

End Sub

Private Sub GenerateTable(table As TableDef)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 13/09/1999
   ' * Time             : 14:51
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GenerateTable
   ' * Parameters       :
   ' *                    table As TableDef
   ' **********************************************************************
   ' * Comments         : Generate the code for a table
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GenerateTable

   Dim nI               As Long
   Dim nJ               As Long
   Dim sTable           As String
   Dim fld              As Field
   Dim idx              As index

   ' *** Get table name
   sTable = table.Name

   If left(UCase$(sTable), 2) = "MS" Then Exit Sub

   msGenerated = msGenerated & "Private Function CreateTable" & sTable & "(DB as Database) As Boolean" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtils#************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Programmer Name  : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Web Site         : http://www.ppreview.net" & vbCrLf
   msGenerated = msGenerated & "   ' * E-Mail           : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Date             : " & Date & vbCrLf
   msGenerated = msGenerated & "   ' * Time             : " & time & vbCrLf
   msGenerated = msGenerated & "   ' * Procedure Name   : CreateTable" & sTable & "" & vbCrLf
   msGenerated = msGenerated & "   ' * Parameters       :" & vbCrLf
   msGenerated = msGenerated & "   ' *                    DB As Database" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Comments         : This table has been created using VBIDEUtils" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "   On Error GoTo ERROR_CreateTable" & sTable & "" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Dim table            As TableDef" & vbCrLf
   msGenerated = msGenerated & "   Dim fld              As Field" & vbCrLf
   msGenerated = msGenerated & "   Dim idx              As Index" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   CreateTable" & sTable & " = False" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' *** Create the table" & vbCrLf
   msGenerated = msGenerated & "   Set table = DB.CreateTableDef(""" & sTable & """)" & vbCrLf

   ' *** Create the fields
   For nI = 0 To table.Fields.Count - 1
      Set fld = table.Fields(nI)
      msGenerated = msGenerated & "   ' *** Create the field " & fld.Name & vbCrLf
      msGenerated = msGenerated & "   Set fld = table.CreateField(""" & fld.Name & """, " & fld.Type & ")" & vbCrLf
      msGenerated = msGenerated & "   With fld" & vbCrLf
      msGenerated = msGenerated & "      .Attributes = " & fld.Attributes & vbCrLf
      msGenerated = msGenerated & "      .Required = " & IIf(fld.Required, "True", "False") & vbCrLf
      msGenerated = msGenerated & "      .OrdinalPosition = " & fld.OrdinalPosition & vbCrLf
      msGenerated = msGenerated & "      .Size = " & fld.Size & vbCrLf
      If (fld.Type = dbChar) Or _
         (fld.Type = dbMemo) Or _
         (fld.Type = dbText) Then
         msGenerated = msGenerated & "      .AllowZeroLength = " & IIf(fld.AllowZeroLength, "True", "False") & vbCrLf
      End If
      msGenerated = msGenerated & "   End With" & vbCrLf
      msGenerated = msGenerated & "   table.Fields.Append fld" & vbCrLf
      msGenerated = msGenerated & "   table.Fields.Refresh" & vbCrLf
      msGenerated = msGenerated & "" & vbCrLf
      Set fld = Nothing
   Next

   ' *** Create the index
   For nI = 0 To table.Indexes.Count - 1
      Set idx = table.Indexes(nI)
      msGenerated = msGenerated & "   ' *** Create the Index " & idx.Name & vbCrLf
      msGenerated = msGenerated & "   Set idx = table.CreateIndex" & vbCrLf
      msGenerated = msGenerated & "   With idx" & vbCrLf
      msGenerated = msGenerated & "      .Name = """ & idx.Name & """" & vbCrLf
      msGenerated = msGenerated & "      .Primary = " & IIf(idx.Primary, "True", "False") & vbCrLf
      msGenerated = msGenerated & "      .Unique = " & IIf(idx.Unique, "True", "False") & vbCrLf
      msGenerated = msGenerated & "      .Required = " & IIf(idx.Required, "True", "False") & vbCrLf
      msGenerated = msGenerated & "      .Clustered = " & IIf(idx.Clustered, "True", "False") & vbCrLf
      msGenerated = msGenerated & "      .IgnoreNulls = " & IIf(idx.IgnoreNulls, "True", "False") & vbCrLf

      For nJ = 0 To idx.Fields.Count - 1
         msGenerated = msGenerated & "      Set fld = .CreateField(""" & idx.Fields(nJ).Name & """)" & vbCrLf
         msGenerated = msGenerated & "      .Fields.Append fld" & vbCrLf
      Next

      msGenerated = msGenerated & "   End With" & vbCrLf
      msGenerated = msGenerated & "   table.Indexes.Append idx" & vbCrLf
      msGenerated = msGenerated & "   table.Indexes.Refresh" & vbCrLf
      msGenerated = msGenerated & "" & vbCrLf

      Set idx = Nothing
   Next

   msGenerated = msGenerated & "   DB.TableDefs.Append table" & vbCrLf
   msGenerated = msGenerated & "   DB.TableDefs.Refresh" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' *** This table is done :)" & vbCrLf
   msGenerated = msGenerated & "   CreateTable" & sTable & " = True" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "EXIT_CreateTable" & sTable & ":" & vbCrLf
   msGenerated = msGenerated & "   Set idx = Nothing" & vbCrLf
   msGenerated = msGenerated & "   Set fld = Nothing" & vbCrLf
   msGenerated = msGenerated & "   Set table = Nothing" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Exit Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "ERROR_CreateTable" & sTable & ":" & vbCrLf
   msGenerated = msGenerated & "   Select Case MsgBox(""Error "" & Err.Number & "": "" & Err.Description & vbCrLf & ""in CreateTable" & sTable & """ & vbCrLf & ""The error occured at line: "" & Erl, vbAbortRetryIgnore + vbCritical, ""Error"")" & vbCrLf
   msGenerated = msGenerated & "      Case vbAbort" & vbCrLf
   msGenerated = msGenerated & "         Screen.MousePointer = vbDefault" & vbCrLf
   msGenerated = msGenerated & "         Resume EXIT_CreateTable" & sTable & ":" & vbCrLf
   msGenerated = msGenerated & "      Case vbRetry" & vbCrLf
   msGenerated = msGenerated & "         Resume" & vbCrLf
   msGenerated = msGenerated & "      Case vbIgnore" & vbCrLf
   msGenerated = msGenerated & "         Resume Next" & vbCrLf
   msGenerated = msGenerated & "   End Select" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Resume EXIT_CreateTable" & sTable & "" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "End Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

EXIT_GenerateTable:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GenerateTable:
   'MsgBox "Error in ERROR_GenerateTable : " & Error
   Resume EXIT_GenerateTable

End Sub

Private Sub GenerateQuery(qry As QueryDef)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 13/09/1999
   ' * Time             : 14:51
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GenerateQuery
   ' * Parameters       :
   ' *                    qry As QueryDef
   ' **********************************************************************
   ' * Comments         : Generate the code for a querydef
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GenerateQuery

   Dim sTable           As String
   Dim sTable2          As String

   ' *** Get table name
   sTable = Replace(qry.Name, " ", "_")
   sTable2 = qry.Name

   If left(UCase$(sTable), 2) = "MS" Then Exit Sub

   msGenerated = msGenerated & "Private Function CreateQuery" & sTable & "(DB as Database) As Boolean" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtils#************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Programmer Name  : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Web Site         : http://www.ppreview.net" & vbCrLf
   msGenerated = msGenerated & "   ' * E-Mail           : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Date             : " & Date & vbCrLf
   msGenerated = msGenerated & "   ' * Time             : " & time & vbCrLf
   msGenerated = msGenerated & "   ' * Procedure Name   : CreateQuery" & sTable & "" & vbCrLf
   msGenerated = msGenerated & "   ' * Parameters       :" & vbCrLf
   msGenerated = msGenerated & "   ' *                    DB As Database" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Comments         : This query has been created using VBIDEUtils" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "   On Error GoTo ERROR_CreateQuery" & sTable & "" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Dim qry              As QueryDef" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   CreateQuery" & sTable & " = False" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' *** Create the table" & vbCrLf
   msGenerated = msGenerated & "   Set qry = DB.CreateQueryDef(""" & sTable2 & """, """ & Replace(Replace(qry.SQL, """", """"""), vbCrLf, " ") & """) " & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   CreateQuery" & sTable & " = True" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "EXIT_CreateQuery" & sTable & ":" & vbCrLf
   msGenerated = msGenerated & "   Set qry = Nothing" & vbCrLf
   msGenerated = msGenerated & "   Exit Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "ERROR_CreateQuery" & sTable & ":" & vbCrLf
   msGenerated = msGenerated & "   Select Case MsgBox(""Error "" & Err.Number & "": "" & Err.Description & vbCrLf & ""in CreateQuery" & sTable & """ & vbCrLf & ""The error occured at line: "" & Erl, vbAbortRetryIgnore + vbCritical, ""Error"")" & vbCrLf
   msGenerated = msGenerated & "      Case vbAbort" & vbCrLf
   msGenerated = msGenerated & "         Screen.MousePointer = vbDefault" & vbCrLf
   msGenerated = msGenerated & "         Resume EXIT_CreateQuery" & sTable & ":" & vbCrLf
   msGenerated = msGenerated & "      Case vbRetry" & vbCrLf
   msGenerated = msGenerated & "         Resume" & vbCrLf
   msGenerated = msGenerated & "      Case vbIgnore" & vbCrLf
   msGenerated = msGenerated & "         Resume Next" & vbCrLf
   msGenerated = msGenerated & "   End Select" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Resume EXIT_CreateQuery" & sTable & "" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "End Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

EXIT_GenerateQuery:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GenerateQuery:
   'MsgBox "Error in ERROR_GenerateQuery : " & Error
   Resume EXIT_GenerateQuery

End Sub

Private Sub cmdCancel_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
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

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Call InitTooltips

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:51
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
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

Private Sub GenerateADOCode(sDatabase As String, Optional sPassword As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 13/09/1999
   ' * Time             : 14:40
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GenerateADOCode
   ' * Parameters       :
   ' *                    sDatabase As String
   ' *                    Optional sPassword As String
   ' **********************************************************************
   ' * Comments         : Generate the ADO code to create the database
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GenerateADOCode

   Dim oConnection      As ADODB.Connection
   Dim oCatalog         As ADOX.Catalog

   Set oConnection = New ADODB.Connection

   oConnection.Provider = "Microsoft.Jet.OLEDB.4.0"
   oConnection.Mode = adModeRead
   oConnection.CursorLocation = adUseClient
   oConnection.Properties("Data Source") = sDatabase
   oConnection.Properties("Jet OLEDB:Database Password") = sPassword
   oConnection.Open

   Set oCatalog = New ADOX.Catalog
   oCatalog.ActiveConnection = oConnection

   ' *** Main declaration
   msGenerated = "Public Function GenerateDatabase(sDestDBPath as String, Optional sDestDBPassword as String) as Boolean" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtils#************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Programmer Name  : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Web Site         : http://www.ppreview.net" & vbCrLf
   msGenerated = msGenerated & "   ' * E-Mail           : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Date             : " & Date & vbCrLf
   msGenerated = msGenerated & "   ' * Time             : " & time & vbCrLf
   msGenerated = msGenerated & "   ' * Procedure Name   : GenerateDatabase" & vbCrLf
   msGenerated = msGenerated & "   ' * Parameters       :" & vbCrLf
   msGenerated = msGenerated & "   ' *                    sDestDBPath As String" & vbCrLf
   msGenerated = msGenerated & "   ' *                    Optional sDestDBPassword As String" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Comments         : Create a new database by code" & vbCrLf
   msGenerated = msGenerated & "   ' *   This Database has been created using VBIDEUtils" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Dim oCatalog       As ADOX.Catalog" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   On Error Goto ERROR_GenerateDatabase" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateDatabase = False" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' *** Create the database" & vbCrLf
   msGenerated = msGenerated & "   Set oCatalog = New ADOX.Catalog" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   oCatalog.Create ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source="" & sDestDBPath & "";Jet OLEDB:Database Password="" & sDestDBPassword & "";" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

   msGenerated = msGenerated & "   Call GenerateTables(oCatalog)" & vbCrLf
   msGenerated = msGenerated & "   Call GenerateIndexes(oCatalog)" & vbCrLf
   msGenerated = msGenerated & "   Call GenerateKeys(oCatalog)" & vbCrLf

   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateDatabase = True" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "EXIT_GenerateDatabase:" & vbCrLf
   msGenerated = msGenerated & "   On Error Resume Next" & vbCrLf
   msGenerated = msGenerated & "   ' *** Close the database :)" & vbCrLf
   msGenerated = msGenerated & "   Set oCatalog = Nothing" & vbCrLf
   msGenerated = msGenerated & "   Exit Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "ERROR_GenerateDatabase:" & vbCrLf
   msGenerated = msGenerated & "   Select Case MsgBox(""Error "" & Err.Number & "": "" & Err.Description & vbCrLf & ""in GenerateDatabase"" & vbCrLf & ""The error occured at line: "" & Erl, vbAbortRetryIgnore + vbCritical, ""Error"")" & vbCrLf
   msGenerated = msGenerated & "      Case vbAbort" & vbCrLf
   msGenerated = msGenerated & "         Screen.MousePointer = vbDefault" & vbCrLf
   msGenerated = msGenerated & "         GenerateDatabase = False" & vbCrLf
   msGenerated = msGenerated & "         Resume EXIT_GenerateDatabase:" & vbCrLf
   msGenerated = msGenerated & "      Case vbRetry" & vbCrLf
   msGenerated = msGenerated & "         Resume" & vbCrLf
   msGenerated = msGenerated & "      Case vbIgnore" & vbCrLf
   msGenerated = msGenerated & "         Resume Next" & vbCrLf
   msGenerated = msGenerated & "   End Select" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateDatabase = False" & vbCrLf
   msGenerated = msGenerated & "   Resume EXIT_GenerateDatabase" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "End Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

   Call GenerateTables(oCatalog)
   Call GenerateIndexes(oCatalog)
   Call GenerateKeys(oCatalog)

EXIT_GenerateADOCode:
   On Error Resume Next
   Set oCatalog = Nothing

   oConnection.Close
   Set oConnection = Nothing

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GenerateADOCode:
   'Call MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description, vbAbortRetryIgnore + vbCritical, "Error")
   Resume EXIT_GenerateADOCode

End Sub

Private Sub GenerateTables(oCatalog As ADOX.Catalog)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/02/2001
   ' * Time             : 19:33
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GenerateTables
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer
   Dim nJ               As Integer

   msGenerated = msGenerated & "Private Function GenerateTables(oCatalog As ADOX.Catalog) as Boolean" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtils#************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Programmer Name  : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Web Site         : http://www.ppreview.net" & vbCrLf
   msGenerated = msGenerated & "   ' * E-Mail           : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Date             : " & Date & vbCrLf
   msGenerated = msGenerated & "   ' * Time             : " & time & vbCrLf
   msGenerated = msGenerated & "   ' * Procedure Name   : GenerateTables" & vbCrLf
   msGenerated = msGenerated & "   ' * Parameters       :" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Comments         : Generate all the tables" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   On Error Goto ERROR_GenerateTables" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateTables = False" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

   msGenerated = msGenerated & "   Dim oTable           As ADOX.Table" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Set oTable = New ADOX.Table" & vbCrLf

   msGenerated = msGenerated & "" & vbCrLf

   For nI = 0 To oCatalog.Tables.Count - 1
      If oCatalog.Tables(nI).Type = "TABLE" Then
         With oCatalog.Tables(nI)
            msGenerated = msGenerated & "   ' *** Generating Table : " & .Name & vbCrLf
            msGenerated = msGenerated & "   Set oTable = New ADOX.Table" & vbCrLf
            msGenerated = msGenerated & "   oTable.Name = """ & Replace(.Name, """", """""") & """" & vbCrLf
            msGenerated = msGenerated & "   oTable.ParentCatalog = oCatalog" & vbCrLf

            For nJ = 0 To .Columns.Count - 1
               If left$(.Columns(nJ).Name, 2) <> "s_" Then
                  msGenerated = msGenerated & "   oTable.Columns.Append """ & .Columns(nJ).Name & """, " & GetType(.Columns(nJ).Type) & ", " & .Columns(nJ).DefinedSize & vbCrLf
                  '                  Dim nK As Integer
                  '                  For nK = 0 To .Columns(nJ).Properties.Count - 1
                  '                     Debug.Print .Columns(nJ).Properties(nK).Name & " / " & .Columns(nJ).Properties(nK).value
                  '                     Debug.Print " ---" & .Columns(nJ).Properties(nK).Name & " / " & .Columns(nJ).Properties(nK).Attributes
                  '                     Debug.Print " ***" & .Columns(nJ).Properties(nK).Name & " / " & .Columns(nJ).Properties(nK).Type
                  '                  Next

                  'Some Properties
                  If .Columns(nJ).Properties("AutoIncrement").Value Then
                     msGenerated = msGenerated & "      oTable.Columns(""" & .Columns(nJ).Name & """).Properties(""AutoIncrement"").Value = True" & vbCrLf
                  End If
                  If Len(.Columns(nJ).Properties("Description").Value) > 0 Then
                     msGenerated = msGenerated & "      oTable.Columns(""" & .Columns(nJ).Name & """).Properties(""Description"").Value = """ & .Columns(nJ).Properties("Description").Value & """" & vbCrLf
                  End If
                  If Not .Columns(nJ).Properties("Nullable").Value = False Then
                     msGenerated = msGenerated & "      oTable.Columns(""" & .Columns(nJ).Name & """).Properties(""Nullable"").Value = False" & vbCrLf
                  End If
                  If Len(.Columns(nJ).Properties("Default").Value) > 0 Then
                     msGenerated = msGenerated & "      oTable.Columns(""" & .Columns(nJ).Name & """).Properties(""Default"").Value = """ & .Columns(nJ).Properties("Default").Value & """" & vbCrLf
                  End If
                  If Len(.Columns(nJ).Properties("Jet OLEDB:Allow Zero Length").Value) > 0 Then
                     msGenerated = msGenerated & "      oTable.Columns(""" & .Columns(nJ).Name & """).Properties(""Default"").Value = """ & .Columns(nJ).Properties("Jet OLEDB:Allow Zero Length").Value & """" & vbCrLf
                  End If

               End If
            Next

            msGenerated = msGenerated & "   oCatalog.Tables.Append oTable" & vbCrLf
            msGenerated = msGenerated & vbCrLf
         End With
      End If
   Next

   msGenerated = msGenerated & "   Set oTable = Nothing" & vbCrLf
   msGenerated = msGenerated & vbCrLf
   msGenerated = msGenerated & "   GenerateTables = True" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "EXIT_GenerateTables:" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Exit Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "ERROR_GenerateTables:" & vbCrLf
   msGenerated = msGenerated & "   Select Case MsgBox(""Error "" & Err.Number & "": "" & Err.Description & vbCrLf & ""in GenerateTables"" & vbCrLf & ""The error occured at line: "" & Erl, vbAbortRetryIgnore + vbCritical, ""Error"")" & vbCrLf
   msGenerated = msGenerated & "      Case vbAbort" & vbCrLf
   msGenerated = msGenerated & "         Screen.MousePointer = vbDefault" & vbCrLf
   msGenerated = msGenerated & "         GenerateTables = False" & vbCrLf
   msGenerated = msGenerated & "         Resume EXIT_GenerateTables:" & vbCrLf
   msGenerated = msGenerated & "      Case vbRetry" & vbCrLf
   msGenerated = msGenerated & "         Resume" & vbCrLf
   msGenerated = msGenerated & "      Case vbIgnore" & vbCrLf
   msGenerated = msGenerated & "         Resume Next" & vbCrLf
   msGenerated = msGenerated & "   End Select" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateTables = False" & vbCrLf
   msGenerated = msGenerated & "   Resume EXIT_GenerateTables" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "End Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
End Sub

Private Sub GenerateIndexes(oCatalog As ADOX.Catalog)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/02/2001
   ' * Time             : 19:33
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GenerateIndexes
   ' * Parameters       :
   ' *                    oCatalog As ADOX.Catalog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer
   Dim nJ               As Integer
   Dim nK               As Integer

   msGenerated = msGenerated & "Private Function GenerateIndexes(oCatalog As ADOX.Catalog) as Boolean" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtils#************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Programmer Name  : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Web Site         : http://www.ppreview.net" & vbCrLf
   msGenerated = msGenerated & "   ' * E-Mail           : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Date             : " & Date & vbCrLf
   msGenerated = msGenerated & "   ' * Time             : " & time & vbCrLf
   msGenerated = msGenerated & "   ' * Procedure Name   : GenerateIndexes" & vbCrLf
   msGenerated = msGenerated & "   ' * Parameters       :" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Comments         : Generate all the indexes needed" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   On Error Goto ERROR_GenerateIndexes" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateIndexes = False" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

   msGenerated = msGenerated & "   Dim oIndex           As ADOX.Index"
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Set oIndex = New ADOX.Index"

   msGenerated = msGenerated & "" & vbCrLf

   For nI = 0 To oCatalog.Tables.Count - 1
      For nJ = 0 To oCatalog.Tables(nI).Indexes.Count - 1
         If (left$(oCatalog.Tables(nI).Indexes(nJ).Name, 2) <> "s_") _
            And (left$(oCatalog.Tables(nI).Indexes(nJ).Name, 4) <> "MSys") _
            And (left$(oCatalog.Tables(nI).Name, 4) <> "MSys") _
            And (left$(oCatalog.Tables(nI).Indexes(nJ).Name, 1) <> "{") _
            And (right$(oCatalog.Tables(nI).Indexes(nJ).Name, 1) <> "}") Then
            With oCatalog.Tables(nI).Indexes(nJ)
               msGenerated = msGenerated & "   ' *** Generating Index : " & .Name & vbCrLf
               msGenerated = msGenerated & "   Set oIndex = New ADOX.Index" & vbCrLf
               msGenerated = msGenerated & "   oIndex.Name = """ & Replace(.Name, """", """""") & """" & vbCrLf

               For nK = 0 To .Columns.Count - 1
                  msGenerated = msGenerated & "   oIndex.Columns.Append """ & .Columns(nK).Name & vbCrLf
               Next
               msGenerated = msGenerated & "   oIndex.PrimaryKey = " & IIf(.PrimaryKey, "True", "False") & vbCrLf
               msGenerated = msGenerated & "   oIndex.Unique = " & IIf(.Unique, "True", "False") & vbCrLf
               msGenerated = msGenerated & "   oIndex.Clustered = " & IIf(.Clustered, "True", "False") & vbCrLf
               msGenerated = msGenerated & "   oIndex.IndexNulls = " & GetIndexNulls(.IndexNulls) & vbCrLf
               msGenerated = msGenerated & "   oCatalog.Tables(""" & oCatalog.Tables(nI).Name & """).Indexes.Append oIndex" & vbCrLf
               msGenerated = msGenerated & vbCrLf
            End With
         End If
      Next
   Next
   msGenerated = msGenerated & vbCrLf

   msGenerated = msGenerated & "   Set oIndex = Nothing" & vbCrLf
   msGenerated = msGenerated & vbCrLf
   msGenerated = msGenerated & "   GenerateIndexes = True" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "EXIT_GenerateIndexes:" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Exit Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "ERROR_GenerateIndexes:" & vbCrLf
   msGenerated = msGenerated & "   Select Case MsgBox(""Error "" & Err.Number & "": "" & Err.Description & vbCrLf & ""in GenerateIndexes"" & vbCrLf & ""The error occured at line: "" & Erl, vbAbortRetryIgnore + vbCritical, ""Error"")" & vbCrLf
   msGenerated = msGenerated & "      Case vbAbort" & vbCrLf
   msGenerated = msGenerated & "         Screen.MousePointer = vbDefault" & vbCrLf
   msGenerated = msGenerated & "         GenerateIndexes = False" & vbCrLf
   msGenerated = msGenerated & "         Resume EXIT_GenerateIndexes:" & vbCrLf
   msGenerated = msGenerated & "      Case vbRetry" & vbCrLf
   msGenerated = msGenerated & "         Resume" & vbCrLf
   msGenerated = msGenerated & "      Case vbIgnore" & vbCrLf
   msGenerated = msGenerated & "         Resume Next" & vbCrLf
   msGenerated = msGenerated & "   End Select" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateIndexes = False" & vbCrLf
   msGenerated = msGenerated & "   Resume EXIT_GenerateIndexes" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "End Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
End Sub

Private Function GetIndexNulls(ByVal Value As ADOX.AllowNullsEnum) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/02/2001
   ' * Time             : 19:59
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GetIndexNulls
   ' * Parameters       :
   ' *                    ByVal Value As ADOX.AllowNullsEnum
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Select Case Value
      Case adIndexNullsAllow: GetIndexNulls = "adIndexNullsAllow"
      Case adIndexNullsDisallow: GetIndexNulls = "adIndexNullsDisallow"
      Case adIndexNullsIgnore: GetIndexNulls = "adIndexNullsIgnore"
      Case adIndexNullsIgnoreAny: GetIndexNulls = "adIndexNullsIgnoreAny"
      Case Else: GetIndexNulls = Value
   End Select

End Function

Private Sub GenerateKeys(oCatalog As ADOX.Catalog)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/02/2001
   ' * Time             : 19:33
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GenerateKeys
   ' * Parameters       :
   ' *                    oCatalog As ADOX.Catalog
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nI               As Integer
   Dim nJ               As Integer
   Dim nK               As Integer

   msGenerated = msGenerated & "Private Function GenerateKeys(oCatalog As ADOX.Catalog) as Boolean" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtils#************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Programmer Name  : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Web Site         : http://www.ppreview.net" & vbCrLf
   msGenerated = msGenerated & "   ' * E-Mail           : removed" & vbCrLf
   msGenerated = msGenerated & "   ' * Date             : " & Date & vbCrLf
   msGenerated = msGenerated & "   ' * Time             : " & time & vbCrLf
   msGenerated = msGenerated & "   ' * Procedure Name   : GenerateKeys" & vbCrLf
   msGenerated = msGenerated & "   ' * Parameters       :" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "   ' * Comments         : Generate all the keys" & vbCrLf
   msGenerated = msGenerated & "   ' *" & vbCrLf
   msGenerated = msGenerated & "   ' **********************************************************************" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   On Error Goto ERROR_GenerateKeys" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateKeys = False" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf

   msGenerated = msGenerated & "   Dim oKey           As ADOX.Key"
   msGenerated = msGenerated & "" & vbCrLf

   For nI = 0 To oCatalog.Tables.Count - 1
      For nJ = 0 To oCatalog.Tables(nI).Keys.Count - 1
         If left$(oCatalog.Tables(nI).Keys(nJ).Name, 2) <> "s_" Then
            If oCatalog.Tables(nI).Keys(nJ).Type = adKeyForeign Then
               With oCatalog.Tables(nI).Keys(nJ)
                  msGenerated = msGenerated & "   ' *** Generating Key : " & .Name & vbCrLf
                  msGenerated = msGenerated & "   Set oKey = New ADOX.Key" & vbCrLf
                  msGenerated = msGenerated & "   With oKey" & vbCrLf
                  msGenerated = msGenerated & "      .Name = """ & Replace(.Name, """", """""") & """" & vbCrLf
                  msGenerated = msGenerated & "      .Type = " & GetKeyType(.Type) & vbCrLf
                  msGenerated = msGenerated & "      .UpdateRule = " & GetUpdateRule(.UpdateRule) & vbCrLf
                  msGenerated = msGenerated & "      .RelatedTable = """ & .RelatedTable & """" & vbCrLf
                  msGenerated = msGenerated & "      .UpdateRule = " & GetUpdateRule(.UpdateRule) & vbCrLf

                  For nK = 0 To .Columns.Count - 1
                     msGenerated = msGenerated & "      .Columns.Append """ & .Columns(nK).Name & """" & vbCrLf
                     msGenerated = msGenerated & "      .Columns(""" & .Columns(nK).Name & """).RelatedColumn = """ & .Columns(nK).RelatedColumn & """" & vbCrLf
                  Next
                  msGenerated = msGenerated & "   End With" & vbCrLf
                  msGenerated = msGenerated & "   oCatalog.Tables(""" & oCatalog.Tables(nI).Name & """).Keys.Append oKey" & vbCrLf
                  msGenerated = msGenerated & vbCrLf

               End With
            End If
         End If
      Next
   Next
   msGenerated = msGenerated & vbCrLf

   msGenerated = msGenerated & "   Set oKey = Nothing" & vbCrLf
   msGenerated = msGenerated & vbCrLf
   msGenerated = msGenerated & "   GenerateKeys = True" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "EXIT_GenerateKeys:" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   Exit Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   ' #VBIDEUtilsERROR#" & vbCrLf
   msGenerated = msGenerated & "ERROR_GenerateKeys:" & vbCrLf
   msGenerated = msGenerated & "   Select Case MsgBox(""Error "" & Err.Number & "": "" & Err.Description & vbCrLf & ""in GenerateKeys"" & vbCrLf & ""The error occured at line: "" & Erl, vbAbortRetryIgnore + vbCritical, ""Error"")" & vbCrLf
   msGenerated = msGenerated & "      Case vbAbort" & vbCrLf
   msGenerated = msGenerated & "         Screen.MousePointer = vbDefault" & vbCrLf
   msGenerated = msGenerated & "         GenerateKeys = False" & vbCrLf
   msGenerated = msGenerated & "         Resume EXIT_GenerateKeys:" & vbCrLf
   msGenerated = msGenerated & "      Case vbRetry" & vbCrLf
   msGenerated = msGenerated & "         Resume" & vbCrLf
   msGenerated = msGenerated & "      Case vbIgnore" & vbCrLf
   msGenerated = msGenerated & "         Resume Next" & vbCrLf
   msGenerated = msGenerated & "   End Select" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "   GenerateKeys = False" & vbCrLf
   msGenerated = msGenerated & "   Resume EXIT_GenerateKeys" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
   msGenerated = msGenerated & "End Function" & vbCrLf
   msGenerated = msGenerated & "" & vbCrLf
End Sub

Private Function GetKeyType(ByVal Value As ADOX.KeyTypeEnum) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/02/2001
   ' * Time             : 20:17
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GetKeyType
   ' * Parameters       :
   ' *                    ByVal Value As ADOX.KeyTypeEnum
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Select Case Value
      Case adKeyForeign: GetKeyType = "adKeyForeign"
      Case adKeyPrimary: GetKeyType = "adKeyPrimary"
      Case adKeyUnique: GetKeyType = "adKeyUnique"
      Case Else: GetKeyType = Value
   End Select
End Function

Private Function GetUpdateRule(ByVal Value As ADOX.RuleEnum) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/02/2001
   ' * Time             : 20:17
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : cUpdateRule
   ' * Parameters       :
   ' *                    ByVal Value As ADOX.RuleEnum
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Select Case Value
      Case adRINone: GetUpdateRule = "adRINone"
      Case adRICascade: GetUpdateRule = "adRICascade"
      Case adRISetNull: GetUpdateRule = "adRISetNull"
      Case adRISetDefault: GetUpdateRule = "adRISetDefault"
      Case Else: GetUpdateRule = Value
   End Select
End Function

Private Function GetType(ByVal Value As ADOX.DataTypeEnum) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/02/2001
   ' * Time             : 20:19
   ' * Module Name      : frmDBCreator
   ' * Module Filename  : DBCreator.frm
   ' * Procedure Name   : GetType
   ' * Parameters       :
   ' *                    ByVal Value As ADOX.DataTypeEnum
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Select Case Value
      Case adTinyInt: GetType = "adTinyInt"
      Case adSmallInt: GetType = "adSmallInt"
      Case adInteger: GetType = "adInteger"
      Case adBigInt: GetType = "adBigInt"
      Case adUnsignedTinyInt: GetType = "adUnsignedTinyInt"
      Case adUnsignedSmallInt: GetType = "adUnsignedSmallInt"
      Case adUnsignedInt: GetType = "adUnsignedInt"
      Case adUnsignedBigInt: GetType = "adUnsignedBigInt"
      Case adSingle: GetType = "adSingle"
      Case adDouble: GetType = "adDouble"
      Case adCurrency: GetType = "adCurrency"
      Case adDecimal: GetType = "adDecimal"
      Case adNumeric: GetType = "adNumeric"
      Case adBoolean: GetType = "adBoolean"
      Case adUserDefined: GetType = "adUserDefined"
      Case adVariant: GetType = "adVariant"
      Case adGUID: GetType = "adGUID"
      Case adDate: GetType = "adDate"
      Case adDBDate: GetType = "adDBDate"
      Case adDBTime: GetType = "adDBTime"
      Case adDBTimeStamp: GetType = "adDBTimeStamp"
      Case adBSTR: GetType = "adBSTR"
      Case adChar: GetType = "adChar"
      Case adVarChar: GetType = "adVarChar"
      Case adLongVarChar: GetType = "adLongVarChar"
      Case adWChar: GetType = "adWChar"
      Case adVarWChar: GetType = "adVarWChar"
      Case adLongVarWChar: GetType = "adLongVarWChar"
      Case adBinary: GetType = "adBinary"
      Case adVarBinary: GetType = "adVarBinary"
      Case adLongVarBinary: GetType = "adLongVarBinary"
      Case Else: GetType = Value
   End Select
End Function

