VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGetServers 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "List of all OLE Servers on your System"
   ClientHeight    =   4620
   ClientLeft      =   2835
   ClientTop       =   3735
   ClientWidth     =   8475
   Icon            =   "OLEGetServers.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Height          =   330
      Left            =   120
      Picture         =   "OLEGetServers.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Refresh the list of controls and their accelerators"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   7329
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmGetServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 06/09/2001
' * Time             : 00:20
' * Module Name      : frmGetServers
' * Module Filename  : fGetServers.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Public Sub GetOLEServersInfos()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/09/2001
   ' * Time             : 00:20
   ' * Module Name      : frmGetServers
   ' * Module Filename  : fGetServers.frm
   ' * Procedure Name   : GetOLEServersInfos
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetOLEServersInfos

   Dim cR               As New class_Registry
   Dim sSect()          As String
   Dim iSectCount       As Long
   Dim iSect            As Long
   Dim sClassID         As String
   Dim sName            As String
   Dim itmX             As ListItem

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Me.Enabled = False

   frmProgress.MessageText = "Indenting"
   frmProgress.bCancel = True
   Load frmProgress
   frmProgress.Show
   frmProgress.ZOrder

   SetRedraw Me, False

   lvwItems.ListItems.Clear
   cR.ClassKey = HKEY_CLASSES_ROOT
   If (cR.EnumerateSections(sSect(), iSectCount)) Then
      frmProgress.Maximum = iSectCount

      For iSect = 1 To iSectCount
         If gbCancelProgress Then Exit For

         frmProgress.Progress = iSect

         cR.SectionKey = sSect(iSect)
         sName = cR.Value

         frmProgress.MessageText = "Server : " & sName
         DoEvents

         cR.SectionKey = sSect(iSect) & "\CLSID"
         sClassID = cR.Value
         If (sClassID <> "") Then
            Set itmX = lvwItems.ListItems.Add(, , sName)
            itmX.SubItems(1) = sSect(iSect)
            itmX.SubItems(2) = sClassID
            cR.SectionKey = "CLSID\" & sClassID & "\LocalServer32"
            itmX.SubItems(3) = cR.Value
            cR.SectionKey = "CLSID\" & sClassID & "\InProcServer32"
            itmX.SubItems(4) = cR.Value
            cR.SectionKey = "CLSID\" & sClassID & "\ThreadingModel"
            itmX.SubItems(5) = cR.Value
            cR.SectionKey = "CLSID\" & sClassID & "\Version"
            itmX.SubItems(6) = cR.Value
            cR.SectionKey = "CLSID\" & sClassID & "\VersionIndependentProgID"
            itmX.SubItems(7) = cR.Value

            If itmX.SubItems(3) <> vbNullString Then
               If FileExist(itmX.SubItems(3)) Then
                  itmX.SubItems(8) = True
               Else
                  itmX.SubItems(8) = False
               End If
            ElseIf itmX.SubItems(4) <> vbNullString Then
               If FileExist(itmX.SubItems(4)) Then
                  itmX.SubItems(8) = True
               Else
                  itmX.SubItems(8) = False
               End If
            End If

         End If
      Next

   Else
      MsgBox "Can't get sections in HKEY_CLASSES_ROOT", vbInformation
   End If

EXIT_GetOLEServersInfos:
   SetRedraw Me, True
   Me.Enabled = True

   Unload frmProgress
   Set frmProgress = Nothing

   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GetOLEServersInfos:
   Resume EXIT_GetOLEServersInfos

End Sub

Private Sub cmdRefresh_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/09/2001
   ' * Time             : 10:07
   ' * Module Name      : frmGetServers
   ' * Module Filename  : fGetServers.frm
   ' * Procedure Name   : cmdRefresh_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call GetOLEServersInfos

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/09/2001
   ' * Time             : 00:20
   ' * Module Name      : frmGetServers
   ' * Module Filename  : fGetServers.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   lvwItems.ColumnHeaders.Add , "USERNAME", "User Name"
   lvwItems.ColumnHeaders.Add , "CLASSNAME", "Class Name"
   lvwItems.ColumnHeaders.Add , "CLASSID", "ClassID"
   lvwItems.ColumnHeaders.Add , "EXECUTABLE", "Exe"
   lvwItems.ColumnHeaders.Add , "IPEXECUTABLE", "InProc Exe"
   lvwItems.ColumnHeaders.Add , "THREADINGMODEL", "Threading Model"
   lvwItems.ColumnHeaders.Add , "VERSION", "Version"
   lvwItems.ColumnHeaders.Add , "VERSIONIND", "V.Ind Prog ID"
   lvwItems.ColumnHeaders.Add , "FILEEXISTS", "Is File exists"
   lvwItems.View = lvwReport

   Me.Show
   DoEvents

   GetOLEServersInfos

End Sub

Private Sub Form_Resize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 06/09/2001
   ' * Time             : 10:07
   ' * Module Name      : frmGetServers
   ' * Module Filename  : fGetServers.frm
   ' * Procedure Name   : Form_Resize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   lvwItems.Move 100, 360, Me.Width - 300, Me.Height - 760

End Sub
