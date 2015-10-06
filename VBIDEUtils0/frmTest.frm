VERSION 5.00
Begin VB.Form frmSSubTmr 
   Caption         =   "vbAccelerator SSubTmr Tester"
   ClientHeight    =   2925
   ClientLeft      =   3870
   ClientTop       =   2850
   ClientWidth     =   5085
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   5085
   Begin VB.CheckBox Check2 
      Caption         =   "Timer 2"
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   660
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Timer 1"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Label lblSubClass 
      Caption         =   $"frmTest.frx":014A
      Height          =   735
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   4995
   End
   Begin VB.Label lblTimer 
      Caption         =   "All Code Timer Test:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4995
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1500
      TabIndex        =   3
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1500
      TabIndex        =   2
      Top             =   300
      Width           =   1035
   End
End
Attribute VB_Name = "frmSSubTmr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_t1 As CTimer
Attribute m_t1.VB_VarHelpID = -1
Private WithEvents m_t2 As CTimer
Attribute m_t2.VB_VarHelpID = -1

Implements ISubClass
Private m_emr As EMsgResponse

Private Const WM_SIZE = &H5
Private Const WM_LBUTTONDOWN = &H201


Private Sub Check1_Click()
Dim lI As Long
    If (Check1.Value <> 0) Then
        lI = 500
    Else
        lI = -1
    End If
    m_t1.Interval = lI
End Sub

Private Sub Check2_Click()
Dim lI As Long
    If (Check2.Value <> 0) Then
        lI = 500
    Else
        lI = -1
    End If
    m_t2.Interval = lI
End Sub

Private Sub Form_Load()
    AttachMessage Me, Me.hWnd, WM_LBUTTONDOWN
    AttachMessage Me, Me.hWnd, WM_SIZE
    
    Set m_t1 = New CTimer
    m_t1.Interval = 500
    Set m_t2 = New CTimer
    m_t2.Interval = 500
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DetachMessage Me, Me.hWnd, WM_LBUTTONDOWN
    DetachMessage Me, Me.hWnd, WM_SIZE
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
   Debug.Print CurrentMessage
    ISubClass_MsgResponse = m_emr
End Property

Private Function ISubClass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Debug.Print "Got Message"
    m_emr = emrPostProcess
End Function

Private Sub m_t1_ThatTime()
    Label1.Caption = Format$(Now, "hh:nn:ss")
End Sub

Private Sub m_t2_ThatTime()
    Label2.Caption = Format$(Now, "hh:nn:ss")
End Sub
