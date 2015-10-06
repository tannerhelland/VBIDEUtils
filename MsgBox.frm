VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Messagebox Assistant"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8625
   Icon            =   "MsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   "&Function / Procedure"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   3600
      TabIndex        =   32
      Top             =   4920
      Width           =   3255
      Begin VB.OptionButton optFunction 
         Caption         =   "Function"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optProcedure 
         Caption         =   "Procedure"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   735
      Left            =   7080
      Picture         =   "MsgBox.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Bye bye"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopyClipboard 
      Caption         =   "&Copy to Clipboard"
      Height          =   735
      Left            =   7080
      Picture         =   "MsgBox.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Copy the code to the clipboard"
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "T&est"
      Height          =   735
      Left            =   7080
      Picture         =   "MsgBox.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Make a try for your new created message box"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Options"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "Some more options?"
      Top             =   4920
      Width           =   3255
      Begin VB.OptionButton optSystemModal 
         Caption         =   "System modal"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optApplicationModal 
         Caption         =   "Application modal"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "&Default button"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   30
      ToolTipText     =   "Wich button is the default one?"
      Top             =   4080
      Width           =   6735
      Begin VB.OptionButton optDefaultButton 
         Caption         =   "Default Button 3"
         Height          =   375
         Index           =   2
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optDefaultButton 
         Caption         =   "Default Button 2"
         Height          =   375
         Index           =   1
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optDefaultButton 
         Caption         =   "Default Button 1"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "&Buttons to use"
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Select the needed buttons..."
      Top             =   2760
      Width           =   6735
      Begin VB.OptionButton optButtons 
         Caption         =   "Retry and Cancel"
         Height          =   375
         Index           =   5
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optButtons 
         Caption         =   "Yes and No"
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optButtons 
         Caption         =   "Yes, No, and Cancel "
         Height          =   375
         Index           =   4
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optButtons 
         Caption         =   "Abort, Retry, and Ignore"
         Height          =   375
         Index           =   2
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optButtons 
         Caption         =   "OK and Cancel"
         Height          =   375
         Index           =   1
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optButtons 
         Caption         =   "Ok Only"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Prompt"
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   1800
      TabIndex        =   28
      Top             =   840
      Width           =   5055
      Begin VB.TextBox tbPrompt 
         Height          =   1485
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "MsgBox.frx":09E0
         ToolTipText     =   "Enter the prompt text"
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Title"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   1800
      TabIndex        =   27
      Top             =   0
      Width           =   5055
      Begin VB.TextBox tbTitle 
         Height          =   285
         Left            =   120
         MaxLength       =   80
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "Set the title for the message box"
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose the &icon"
      ForeColor       =   &H8000000D&
      Height          =   2655
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Choose the desired icon for the message box"
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton optStop 
         Height          =   195
         Left            =   720
         TabIndex        =   7
         Top             =   2160
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "MsgBox.frx":09E6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   26
         Top             =   2040
         Width           =   480
      End
      Begin VB.OptionButton optExclamation 
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "MsgBox.frx":0E28
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   25
         Top             =   840
         Width           =   480
      End
      Begin VB.OptionButton optInformation 
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "MsgBox.frx":126A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   24
         Top             =   240
         Width           =   480
      End
      Begin VB.OptionButton optInterrogation 
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "MsgBox.frx":16AC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   23
         Top             =   1440
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:51
' * Module Name      : frmMessageBox
' * Module Filename  : MsgBox.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Private Sub cmdCopyClipboard_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 3/02/99
   ' * Time             : 15:10
   ' * Module Name      : frmMessageBox
   ' * Module Filename  : MsgBox.frm
   ' * Procedure Name   : cmdCopyClipboard_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Create the Message box and copy it to the
   ' *
   ' *
   ' **********************************************************************

   Dim sIcon            As String
   Dim sButton          As String
   Dim sDefaultButton   As String

   Dim sMsgbox          As String

   ' *** Get the right icon
   If optInformation.Value Then
      sIcon = "vbInformation"
   ElseIf optExclamation.Value Then
      sIcon = "vbExclamation"
   ElseIf optInterrogation.Value Then
      sIcon = "vbQuestion"
   ElseIf optStop.Value Then
      sIcon = "vbCritical"
   End If

   ' *** Get the right buttons
   If optButtons(0).Value Then
      sButton = "vbOKOnly"
   ElseIf optButtons(1).Value Then
      sButton = "vbOKCancel"
   ElseIf optButtons(2).Value Then
      sButton = "vbAbortRetryIgnore"
   ElseIf optButtons(3).Value Then
      sButton = "vbYesNo"
   ElseIf optButtons(4).Value Then
      sButton = "vbYesNoCancel"
   ElseIf optButtons(5).Value Then
      sButton = "vbRetryCancel"
   End If

   ' *** Get the default button
   If optDefaultButton(0).Value Then
      sDefaultButton = "vbDefaultButton1"
   ElseIf optDefaultButton(1).Value Then
      sDefaultButton = "vbDefaultButton2"
   ElseIf optDefaultButton(2).Value Then
      sDefaultButton = "vbDefaultButton3"
   End If

   If Not gbRegistered Then tbPrompt.Text = tbPrompt.Text & "  <Shareware Version>"

   ' *** Create the message box
   sMsgbox = "Call MsgBox("
   sMsgbox = sMsgbox & """" & TreatText(tbPrompt.Text) & """, "
   If Trim$(sIcon <> "") Then sMsgbox = sMsgbox & sIcon & " + "
   If Trim$(sButton <> "") Then sMsgbox = sMsgbox & sButton & " + "
   If Trim$(sButton <> "") Then sMsgbox = sMsgbox & sDefaultButton & " + "

   If right(sMsgbox, 2) = "+ " Then sMsgbox = left$(sMsgbox, Len(sMsgbox) - 2)
   sMsgbox = sMsgbox & ", "

   sMsgbox = sMsgbox & """" & TreatText(tbTitle.Text) & """)"

   ' *** Copy to the clipboard
   Clipboard.Clear
   Clipboard.SetText sMsgbox, vbCFText

   ' *** Copy done
   sMsgbox = "The created message box " & Chr$(13) & "has been copied to the clipboard." & Chr$(13)
   sMsgbox = sMsgbox & "You can paste it in your code."

   Call MsgBoxTop(Me.hWnd, sMsgbox, vbOKOnly + vbInformation, "Messagebox Creator")

   ' *** Exit
   Call cmdExit_Click

End Sub

Private Sub cmdExit_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 3/02/99
   ' * Time             : 15:31
   ' * Module Name      : frmMessageBox
   ' * Module Filename  : MsgBox.frm
   ' * Procedure Name   : cmdExit_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Exit
   ' *
   ' *
   ' **********************************************************************

   Call UnloadEffect(Me)
   Unload Me

End Sub

Private Sub cmdTest_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 3/02/99
   ' * Time             : 14:45
   ' * Module Name      : frmMessageBox
   ' * Module Filename  : MsgBox.frm
   ' * Procedure Name   : cmdTest_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Test the message box
   ' *
   ' *
   ' **********************************************************************

   Dim nIcon            As Long
   Dim nButton          As Long
   Dim nDefaultButton   As Long

   ' *** Get the right icon
   If optInformation.Value Then
      nIcon = vbInformation
   ElseIf optExclamation.Value Then
      nIcon = vbExclamation
   ElseIf optInterrogation.Value Then
      nIcon = vbQuestion
   ElseIf optStop.Value Then
      nIcon = vbCritical
   End If

   ' *** Get the right buttons
   If optButtons(0).Value Then
      nButton = vbOKOnly
   ElseIf optButtons(1).Value Then
      nButton = vbOKCancel
   ElseIf optButtons(2).Value Then
      nButton = vbAbortRetryIgnore
   ElseIf optButtons(3).Value Then
      nButton = vbYesNo
   ElseIf optButtons(4).Value Then
      nButton = vbYesNoCancel
   ElseIf optButtons(5).Value Then
      nButton = vbRetryCancel
   End If

   ' *** Get the default button
   If optDefaultButton(0).Value Then
      nDefaultButton = vbDefaultButton1
   ElseIf optDefaultButton(1).Value Then
      nDefaultButton = vbDefaultButton2
   ElseIf optDefaultButton(2).Value Then
      nDefaultButton = vbDefaultButton3
   End If

   ' *** Test of the message box
   MsgBox tbPrompt.Text, nIcon + nButton + nDefaultButton, tbTitle.Text

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 3/02/99
   ' * Time             : 14:59
   ' * Module Name      : frmMessageBox
   ' * Module Filename  : MsgBox.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Init the values
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Call InitTooltips

   tbTitle.Text = ""
   tbPrompt.Text = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 3/02/99
   ' * Time             : 14:59
   ' * Module Name      : frmMessageBox
   ' * Module Filename  : MsgBox.frm
   ' * Procedure Name   : Form_Unload
   ' * Parameters       :
   ' *                    Cancel As Integer
   ' **********************************************************************
   ' * Comments         : Exit
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Set clsTooltips = Nothing

End Sub

Private Function TreatText(sText As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 6/10/98
   ' * Time             : 15:18
   ' * Module Name      : frmMessageBox
   ' * Module Filename  : MsgBox.frm
   ' * Procedure Name   : TreatText
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         : Convert all double quotes to double double quotes
   ' *  Change all Returns in chr$(13)
   ' *
   ' **********************************************************************

   Dim nPos             As Integer
   Dim sTmp             As String

   ' *** Remove all double quotes
   sTmp = sText
   nPos = InStr(1, sTmp, """")
   While nPos <> 0
      sTmp = Mid$(sTmp, 1, nPos - 1) & """" & """" & Mid$(sTmp, nPos + 1)
      nPos = InStr(nPos + 2, sTmp, """")
   Wend

   ' *** Change all the #13
   nPos = InStr(1, sTmp, Chr$(13) & Chr$(10))
   While nPos <> 0
      sTmp = Mid$(sTmp, 1, nPos - 1) & """ & Chr$(13) & " & """" & Mid$(sTmp, nPos + 2)
      nPos = InStr(nPos + 2, sTmp, Chr$(13) & Chr$(10))
   Wend

   TreatText = sTmp

End Function

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
