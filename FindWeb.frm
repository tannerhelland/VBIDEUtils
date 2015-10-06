VERSION 5.00
Begin VB.Form frmFindWeb 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search the Web using VBDiamond"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5490
   ControlBox      =   0   'False
   Icon            =   "FindWeb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton mnuAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Do you want to know more about us?"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cbFind 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Search code, tips, file... on the web, and all this related to VB"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Cancel it, go back to the editor"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      ToolTipText     =   "Start the search"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   $"FindWeb.frx":058A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sample :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "(wav or bmp) and resource file                  Print+Preview"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Search What : "
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   160
      Width           =   1080
   End
End
Attribute VB_Name = "frmFindWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 08/11/1999
' * Time             : 10:56
' * Module Name      : frmFindWeb
' * Module Filename  :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Private m_nCurrentFindRow As Long
Private m_bFound        As Boolean

Private Sub cbFind_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : frmFindWeb
   ' * Module Filename  :
   ' * Procedure Name   : cbFind_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Enable the button

   cmdFindNext.Enabled = True

End Sub

Private Sub cbFind_KeyPress(KeyAscii As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : frmFindWeb
   ' * Module Filename  :
   ' * Procedure Name   : cbFind_KeyPress
   ' * Parameters       :
   ' *                    KeyAscii As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Enable the button

   cmdFindNext.Enabled = True

End Sub

Private Sub cmdCancel_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : frmFindWeb
   ' * Module Filename  :
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

Public Sub cmdFindNext_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : frmFindWeb
   ' * Module Filename  :
   ' * Procedure Name   : cmdFindNext_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *   Search the Web
   ' *
   ' **********************************************************************

   Dim sFind            As String
   Dim sURL             As String
   Dim sMatch           As String

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Keep the text to search
   sFind = cbFind.Text

   If sFind = "" Then Exit Sub

   ' *** Add the value to search in the combo
   If (CBLFind(cbFind, cbFind.Text, 0) = -1) Then
      If gcolFind Is Nothing Then
         Set gcolFind = New Collection
      End If
      gcolFind.Add sFind, sFind
   End If

   Me.Visible = False

   ' *** Replace all blanks
   sFind = Replace(sFind, " and ", "+")
   sFind = Replace(sFind, " ", "+")
   sFind = Replace(sFind, "(", "%28")
   sFind = Replace(sFind, ")", "%29")

   ' *** Create the URL
   'sURL = "http://www.codehound.com/vb/results/results.asp?S=1&Q=" & sFind

   If InStr(sFind, "+") > 0 Then
      sMatch = "allwords"
   Else
      sMatch = "anywords"
   End If
   sFind = Replace(sFind, "+", " ")

   sURL = "http://www.ppreview.net/Sources/SourcesSearch.asp?tbsearch=" & sFind & "&optSearch=title&cbMatch=" & sMatch

   ' *** Launch the search
   ShellExecute 0&, "", sURL, "", "", vbNormalFocus

   Call UnloadEffect(Me)
   Unload Me

End Sub

Public Sub Form_Activate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : frmFindWeb
   ' * Module Filename  :
   ' * Procedure Name   : Form_Activate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Give the focus to the find combo

   On Error Resume Next
   cbFind.SetFocus

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : frmFindWeb
   ' * Module Filename  :
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** fill the fields ***

   On Error Resume Next

   Dim nI               As Integer
   Dim prjProject       As VBProject
   Dim cpCodePane       As CodePane

   Dim nStartLine       As Long
   Dim nStartColumn     As Long
   Dim nEndline         As Long
   Dim nEndColumn       As Long

   Dim sCode            As String

   Dim sTmp             As String

   Call InitTooltips

   sTmp = ""

   ' *** Get the active project
   Set prjProject = VBInstance.ActiveVBProject

   ' *** If we couldn't get it, quit
   If Not (prjProject Is Nothing) Then
      ' *** Try to find the active code pane
      Set cpCodePane = VBInstance.ActiveCodePane

      ' *** If we couldn't get it, quit
      If Not (cpCodePane Is Nothing) Then
         cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn
         sCode = cpCodePane.CodeModule.Lines(nStartLine, IIf(nEndline - nStartLine = 0, 1, nEndline - nStartLine))
         sCode = Replace(sCode, ".", " ")
         sCode = Replace(sCode, "(", " ")
         sCode = Replace(sCode, ")", " ")
         sCode = Replace(sCode, ",", " ")
         sCode = Replace(sCode, ":", " ")

         nI = nStartColumn
         Do While nI > 0
            If Mid$(sCode, nI, 1) <> " " Then
               sTmp = Mid$(sCode, nI, 1) & sTmp
            Else
               Exit Do
            End If
            nI = nI - 1
         Loop
         nI = nStartColumn + 1
         Do While nI < Len(sCode) + 1
            If Mid$(sCode, nI, 1) <> " " Then
               sTmp = sTmp & Mid$(sCode, nI, 1)
            Else
               Exit Do
            End If
            nI = nI + 1
         Loop

      End If

   End If

   ' *** Continue
   cbFind.Clear
   If Not (gcolFind Is Nothing) Then
      For nI = 1 To gcolFind.Count
         cbFind.AddItem gcolFind(nI)
      Next
   End If

   ' *** Empty the find combo
   cbFind.Text = sTmp
   cmdFindNext.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : frmFindWeb
   ' * Module Filename  :
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

Private Sub mnuAbout_Click()

   frmAbout.bAbout = True
   frmAbout.Show vbModal

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
