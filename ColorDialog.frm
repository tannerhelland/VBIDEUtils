VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ChooseColor(nColor As Long) As Long

   Dim hDlg             As class_ColorDialog
   Dim pColor           As Long

   Set hDlg = New class_ColorDialog

   pColor = nColor

   With hDlg
      .hWndOwner = Me.hWnd
      .Flags = cdlgCCDefault
      .color = pColor
      pColor = .color
   End With

   If pColor >= 0 Then Me.BackColor = pColor

   Set hDlg = Nothing

   ChooseColor = pColor

End Function
