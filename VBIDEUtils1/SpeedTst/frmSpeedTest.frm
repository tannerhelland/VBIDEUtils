VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{017E002E-D7CC-11D2-8E21-44B10AC10000}#4.0#0"; "vbalGrid.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8115
   ClientLeft      =   2460
   ClientTop       =   2130
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   6645
   Begin MSFlexGridLib.MSFlexGrid grdMS 
      Height          =   1935
      Left            =   180
      TabIndex        =   6
      Top             =   4680
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   3413
      _Version        =   65541
   End
   Begin vbAcceleratorGrid.vbalGrid grdTest 
      Height          =   2295
      Left            =   180
      TabIndex        =   3
      Top             =   2280
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   4048
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtRows 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "1000"
      Top             =   6960
      Width           =   1515
   End
   Begin ComctlLib.ListView lvwTest 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblMSGrid 
      Caption         =   "Label1"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   7860
      Width           =   1575
   End
   Begin VB.Label lblGrid 
      Caption         =   "Label1"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   7620
      Width           =   1575
   End
   Begin VB.Label lblLvw 
      Caption         =   "Label1"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   7380
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Private Const mcCOLS = 10
Private m_iRows As Long

Private Sub TestListView()
Dim i As Long
Dim j As Long
Dim lT As Long
Dim itmX As ListItem
   lT = timeGetTime
   With lvwTest
      .ListItems.Clear
      .ColumnHeaders.Clear
      For i = 1 To mcCOLS
         .ColumnHeaders.Add , , "Col" & i
      Next i
      For i = 1 To m_iRows
         Set itmX = .ListItems.Add(, , "Row" & i & ";Col 1")
         For j = 2 To mcCOLS
            itmX.SubItems(j - 1) = "Row" & i & ";Col" & j - 1
         Next
      Next
   End With
   lblLvw = timeGetTime - lT
End Sub
Private Sub TestGrid()
Dim i As Long
Dim j As Long
Dim lT As Long
   lT = timeGetTime
   With grdTest
      .Redraw = False
      .Clear True
      For i = 1 To mcCOLS
         .AddColumn , "Col" & i
      Next i
      .Rows = m_iRows
      For i = 1 To m_iRows
         For j = 1 To mcCOLS
            .CellText(i, j) = "Row" & i & ";Col" & j
         Next
      Next
      .Redraw = True
   End With
   lblGrid = timeGetTime - lT
End Sub
Private Sub TestGridMS()
Dim i As Long
Dim j As Long
Dim lT As Long
   lT = timeGetTime
   With grdMS
      .Redraw = False
      .Cols = mcCOLS
      .FixedCols = 0
      .Rows = m_iRows + 1
      .Row = 0
      For i = 1 To mcCOLS
         .Col = i - 1
         .Text = "Col" & i
      Next i
      .FixedRows = 1
      For i = 1 To m_iRows
         For j = 1 To mcCOLS
            .Row = i
            .Col = j - 1
            .Text = "Row" & i & ";Col" & j
         Next
      Next
      .Redraw = True
   End With
   lblMSGrid = timeGetTime - lT
End Sub


Private Sub cmdTest_Click()
   
   m_iRows = CLng(txtRows.Text)
   timeBeginPeriod 1
   
   TestListView
   TestGrid
   TestGridMS
   
   timeEndPeriod 1
   
End Sub

