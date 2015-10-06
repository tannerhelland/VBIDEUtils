VERSION 5.00
Begin VB.Form frmSpy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Spy windows"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4995
   Icon            =   "Spy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3360
      Top             =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Caption :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1605
      Width           =   630
   End
   Begin VB.Label lbCaption 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      ToolTipText     =   "Class name of the window under the cursor"
      Top             =   1560
      Width           =   3795
   End
   Begin VB.Label lbClassName 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "Class name of the window under the cursor"
      Top             =   1080
      Width           =   3795
   End
   Begin VB.Label lbHandle 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Handle of the windows under the mouse cursor"
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Class Name:"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1125
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Handle:"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   645
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Move the cursor over a window.  It's handle and Class Name will be shown below."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4755
   End
End
Attribute VB_Name = "frmSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 09/02/2000
' * Time             : 14:40
' * Module Name      : docClassSpy
' * Module Filename  : ClassName.dob
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

' ClassSpy Sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This shows several APIs - mainly how to get the class name.
' This is handy as a utility so that you can later use the FindWindow API call.

Private Type POINTAPI
   x                    As Long
   y                    As Long
End Type

Private clsTooltips     As New class_Tooltips

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 15:00
   ' * Module Name      : frmSpy
   ' * Module Filename  :
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
   ' * Date             : 09/02/2000
   ' * Time             : 15:00
   ' * Module Name      : frmSpy
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

Private Sub Timer1_Timer()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 14:40
   ' * Module Name      : frmSpy
   ' * Module Filename  :
   ' * Procedure Name   : Timer1_Timer
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim P                As POINTAPI
   Dim nRet             As Long
   Dim nLen             As Long
   Dim nHandle          As Long
   Dim sTmp             As String

   nRet = GetCursorPos(P)
   nHandle = WindowFromPoint(P.x, P.y)
   sTmp = Space$(128)
   nRet = GetClassName(nHandle, sTmp, 128)
   sTmp = left$(sTmp, nRet)
   lbHandle.Caption = nHandle
   lbClassName.Caption = sTmp

   nLen = GetWindowTextLength(nHandle)
   sTmp = String$(nLen + 1, Chr$(0))
   nRet = GetWindowText(nHandle, sTmp, nLen + 1)
   lbCaption.Caption = sTmp

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
