VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3525
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   570
      TabIndex        =   0
      ToolTipText     =   "Enter the name you used in the KeyGen here"
      Top             =   225
      Width           =   2865
   End
   Begin VB.TextBox txtSoftwareKey 
      Height          =   330
      Left            =   570
      TabIndex        =   1
      ToolTipText     =   "Paste the Software Key from KeyGen into here"
      Top             =   840
      Width           =   2865
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Register"
      Default         =   -1  'True
      Height          =   330
      Left            =   600
      TabIndex        =   2
      Top             =   1215
      Width           =   1410
   End
   Begin VB.CommandButton cmbCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2040
      TabIndex        =   3
      Top             =   1215
      Width           =   1410
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Registration Name"
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
      Height          =   195
      Left            =   570
      TabIndex        =   6
      Top             =   0
      Width           =   2025
   End
   Begin VB.Label lblSoftwareKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software License Key"
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
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   600
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   2
      Left            =   2100
      TabIndex        =   4
      Top             =   1635
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   0
      Picture         =   "frmRegister.frx":0CCA
      Top             =   135
      Width           =   495
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 07/14/2001
' * Time             : 22:01
' * Module Name      : frmRegister
' * Module Filename  : frmRegister.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private Sub cmbCancel_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 07/14/2001
   ' * Time             : 22:01
   ' * Module Name      : frmRegister
   ' * Module Filename  : frmRegister.frm
   ' * Procedure Name   : cmbCancel_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Unload Me
End Sub

Private Sub cmdOK_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 07/14/2001
   ' * Time             : 22:01
   ' * Module Name      : frmRegister
   ' * Module Filename  : frmRegister.frm
   ' * Procedure Name   : cmdOK_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   gbRegistered = False

End Sub
