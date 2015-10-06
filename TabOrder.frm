VERSION 5.00
Begin VB.Form frmTabOrder 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "TabOrder Assistant"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2745
   Icon            =   "TabOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstTabIndex 
      DragIcon        =   "TabOrder.frx":014A
      Height          =   285
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "List of controls and their tab order"
      Top             =   345
      Width           =   2025
   End
   Begin VB.CommandButton cmdApply 
      Height          =   330
      Left            =   400
      Picture         =   "TabOrder.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Apply all the modifications"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdUp 
      Height          =   330
      Left            =   735
      Picture         =   "TabOrder.frx":068E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Move this control up in the tab order"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      Height          =   330
      Left            =   1065
      Picture         =   "TabOrder.frx":0790
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move this control down in the tab order"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdTopToBottom 
      Height          =   330
      Left            =   1395
      Picture         =   "TabOrder.frx":0892
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Refresh the list from top to bottom"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdLeftToRight 
      Height          =   330
      Left            =   1725
      Picture         =   "TabOrder.frx":0994
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Refresh list from left to right"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   330
      Left            =   65
      Picture         =   "TabOrder.frx":0A96
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Refresh the list"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
End
Attribute VB_Name = "frmTabOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 28/10/1999
' * Time             : 11:42
' * Module Name      : docTabOrder
' * Module Filename  : Taborder.dob
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Dim mcmpCurrentForm     As VBComponent      'current form
Dim mcolCtls            As VBControls       'form's controls

Private clsTooltips     As New class_Tooltips

'refresh types
Const NEWFORM = 0
Const TOPTOBOTTOM = 1
Const LEFTTORIGHT = 2
Const REFRESHCTLS = 3

Private Sub cmdApply_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : cmdApply_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo cmdApply_ClickErr

   Dim i                As Integer
   Dim sTmp             As String
   Dim nCtlArrIndex     As Integer

   If InRunMode(VBInstance) Then Exit Sub

   Screen.MousePointer = vbHourglass
   For i = 0 To lstTabIndex.ListCount - 1
      GetNameAndIndex lstTabIndex.List(i), sTmp, nCtlArrIndex
      If nCtlArrIndex >= 0 Then
         'set the new tab index
         mcmpCurrentForm.Designer.VBControls.Item(sTmp, nCtlArrIndex).Properties!TabIndex = i
      Else
         'set the new tab index
         mcmpCurrentForm.Designer.VBControls.Item(sTmp).Properties!TabIndex = i
      End If
   Next
   Screen.MousePointer = vbDefault
   Exit Sub

cmdApply_ClickErr:
   '   If MsgBoxTop(Me.hwnd, Err.Description & vbCrLf & "Resume?", vbYesNo) = vbYes Then
   '      Resume Next
   '   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdLeftToRight_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : cmdLeftToRight_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RefreshList LEFTTORIGHT
End Sub

Private Sub cmdTopToBottom_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : cmdTopToBottom_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RefreshList TOPTOBOTTOM
End Sub

Private Sub cmdRefresh_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : cmdRefresh_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RefreshList REFRESHCTLS

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 16/02/2000
   ' * Time             : 14:58
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : Form_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Call InitTooltips

   Call cmdRefresh_Click

   Call Form_Resize

End Sub

Private Sub Form_Show()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : Form_Show
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'load the strings from the resource file
   cmdUp.ToolTipText = "Move the current control up in the list."
   cmdDown.ToolTipText = "Move the current control down in the list."
   cmdTopToBottom.ToolTipText = "List controls in top to bottom order."
   cmdLeftToRight.ToolTipText = "List controls in left to right order."
   cmdRefresh.ToolTipText = "Refresh the list from the form."
   cmdApply.ToolTipText = "Apply the changes to the form."

   Call Form_Resize

End Sub

Private Sub Form_Resize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : Form_Resize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   lstTabIndex.Width = ScaleWidth - (lstTabIndex.left * 2)
   lstTabIndex.Height = ScaleHeight - (cmdApply.Height + 100)
End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error Resume Next

   Set clsTooltips = Nothing

End Sub

Private Sub lstTabIndex_DragDrop(Source As Control, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : lstTabIndex_DragDrop
   ' * Parameters       :
   ' *                    Source As Control
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' * this sub moves the dragged item to a new location
   ' * based on the Y coordinate where it was dropped
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim nListIndex       As Integer
   Dim nPos             As Integer
   Dim i                As Integer

   With lstTabIndex
      nListIndex = .ListIndex
      If Source = lstTabIndex Then
         If nListIndex >= 0 Then
            sTmp = .Text
            nPos = (y \ TextHeight(sTmp)) + .TopIndex
            'check for the last item
            If nPos > .ListCount Then
               nPos = .ListCount
            End If
            .AddItem sTmp, nPos
            If nListIndex > nPos Then
               .RemoveItem nListIndex + 1
            Else
               .RemoveItem nListIndex
            End If
         End If
      End If
   End With

End Sub

Sub lstTabIndex_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : lstTabIndex_MouseMove
   ' * Parameters       :
   ' *                    button As Integer
   ' *                    Shift As Integer
   ' *                    x As Single
   ' *                    Y As Single
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If Button = vbLeftButton Then lstTabIndex.Drag
End Sub

Private Sub cmdUp_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : cmdUp_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next
   Dim nItem            As Integer

   With lstTabIndex
      If .ListIndex < 0 Then Exit Sub
      nItem = .ListIndex
      If nItem = 0 Then Exit Sub  'can't move 1st item up
      'move item up
      .AddItem .Text, nItem - 1
      'remove old item
      .RemoveItem nItem + 1
      'select the item that was just moved
      .Selected(nItem - 1) = True
   End With
End Sub

Private Sub cmdDown_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : cmdDown_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next
   Dim nItem            As Integer

   With lstTabIndex
      If .ListIndex < 0 Then Exit Sub
      nItem = .ListIndex
      If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
      'move item down
      .AddItem .Text, nItem + 2
      'remove old item
      .RemoveItem nItem
      'select the item that was just moved
      .Selected(nItem + 1) = True
   End With
End Sub

Function ControlName(Ctl As VBIDE.VBControl) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : ControlName
   ' * Parameters       :
   ' *                    ctl As VBIDE.VBControl
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' * this function returns a value to put in the listbox
   ' * for a control. It appends the Caption property
   ' * if it exists and is not null. It also appends
   ' * the control array index if the control is a
   ' * member of a control array
   ' *
   ' **********************************************************************

   On Error Resume Next

   Dim sTmp             As String
   Dim sCaption         As String
   Dim i                As Integer

   sTmp = Ctl.Properties!Name
   sCaption = Ctl.Properties!Caption
   'will be null if there isn't one

   i = Ctl.Properties!index
   If i >= 0 Then
      sTmp = sTmp & "(" & i & ")"
   End If

   If Len(sCaption) > 0 Then
      ControlName = sTmp & " - '" & sCaption & "'"
   Else
      ControlName = sTmp
   End If

   err.Clear

End Function

Public Sub RefreshList(nType As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : RefreshList
   ' * Parameters       :
   ' *                    nType As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo RefreshListErr

   Dim i                As Integer
   Dim Ctl              As VBControl
   Dim sTmp             As String
   Dim ti               As Integer
   Dim sCtlName         As String
   Dim nCtlArrIndex     As Integer

   If InRunMode(VBInstance) Then Exit Sub

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   ' *** Clear the list control
   lstTabIndex.Clear

   If VBInstance.ActiveVBProject Is Nothing Then Exit Sub
   If nType = NEWFORM Then
      If mcmpCurrentForm Is VBInstance.SelectedVBComponent Then
         ' *** Same one as we have now
         Exit Sub
      End If
   End If

   ' *** Load the component
   Set mcmpCurrentForm = VBInstance.SelectedVBComponent

   ' *** Check to see if we have a valid component
   If mcmpCurrentForm Is Nothing Then Exit Sub

   ' *** Make sure the active component is a form, user control or property page
   If (mcmpCurrentForm.Type <> vbext_ct_VBForm) And _
      (mcmpCurrentForm.Type <> vbext_ct_UserControl) And _
      (mcmpCurrentForm.Type <> vbext_ct_DocObject) And _
      (mcmpCurrentForm.Type <> vbext_ct_PropPage) Then
      Exit Sub
   End If

   Set mcolCtls = mcmpCurrentForm.Designer.VBControls

   For Each Ctl In mcmpCurrentForm.Designer.VBControls
      ' *** Try to get the tabindex
      On Error Resume Next
      ti = Ctl.Properties!TabIndex
      If err Then
         ' *** Doesn't have a tabindex
         err.Clear
         GoTo SkipIt
      End If
      On Error GoTo RefreshListErr

      sTmp = ControlName(Ctl)

      ' *** Find out where it goes in the list
      Select Case nType
         Case NEWFORM, REFRESHCTLS
            For i = 0 To lstTabIndex.ListCount - 1
               If ti < lstTabIndex.ItemData(i) Then
                  Exit For
               End If
            Next

         Case TOPTOBOTTOM
            ' *** Rearrange from top to bottom
            For i = lstTabIndex.ListCount To 1 Step -1
               GetNameAndIndex lstTabIndex.List(i - 1), sCtlName, nCtlArrIndex
               If nCtlArrIndex >= 0 Then
                  ' *** Control array member
                  If Ctl.Properties!top > mcolCtls(sCtlName, nCtlArrIndex).Properties!top Then
                     ' *** It is above the current list item
                     Exit For
                  ElseIf Ctl.Properties!top = mcolCtls(sCtlName, nCtlArrIndex).Properties!top Then
                     ' *** It is at the same top position so see if it is farther left
                     If Ctl.Properties!left > mcolCtls(sCtlName, nCtlArrIndex).Properties!left Then
                        Exit For
                     End If
                  End If
               Else
                  If Ctl.Properties!top > mcolCtls(sCtlName).Properties!top Then
                     Exit For
                  ElseIf Ctl.Properties!top = mcolCtls(sCtlName).Properties!top Then
                     If Ctl.Properties!left > mcolCtls(sCtlName).Properties!left Then
                        Exit For
                     End If
                  End If
               End If
            Next

         Case LEFTTORIGHT
            ' *** Rearrange from left to right
            For i = lstTabIndex.ListCount To 1 Step -1
               GetNameAndIndex lstTabIndex.List(i - 1), sCtlName, nCtlArrIndex
               If nCtlArrIndex >= 0 Then
                  ' *** Control array member
                  If Ctl.Properties!left > mcolCtls(sCtlName, nCtlArrIndex).Properties!left Then
                     Exit For
                  ElseIf Ctl.Properties!left = mcolCtls(sCtlName, nCtlArrIndex).Properties!left Then
                     If Ctl.Properties!top > mcolCtls(sCtlName, nCtlArrIndex).Properties!top Then
                        Exit For
                     End If
                  End If
               Else
                  If Ctl.Properties!left > mcolCtls(sCtlName).Properties!left Then
                     Exit For
                  ElseIf Ctl.Properties!left = mcolCtls(sCtlName).Properties!left Then
                     If Ctl.Properties!top > mcolCtls(sCtlName).Properties!top Then
                        Exit For
                     End If
                  End If
               End If
            Next

      End Select

      ' *** Add it to the list
      lstTabIndex.AddItem sTmp, i
      lstTabIndex.ItemData(lstTabIndex.NewIndex) = ti
      lstTabIndex.Refresh

SkipIt:
   Next

   Exit Sub
RefreshListErr:
   'MsgBox Err.Description
   Exit Sub

End Sub

Public Sub ControlRemoved(Ctl As VBControl)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : ControlRemoved
   ' * Parameters       :
   ' *                    ctl As VBControl
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim i                As Integer

   sTmp = ControlName(Ctl)
   For i = 0 To lstTabIndex.ListCount - 1
      If lstTabIndex.List(i) = sTmp Then
         'remove it from the list
         lstTabIndex.RemoveItem i
         Exit Sub
      End If
   Next

End Sub

Public Sub ControlAdded(Ctl As VBControl)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : ControlAdded
   ' * Parameters       :
   ' *                    ctl As VBControl
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim i                As Integer

   'try to get the tabindex
   On Error Resume Next
   i = Ctl.Properties!TabIndex
   If err Then
      err.Clear
      'doesn't have a tabindex
      Exit Sub
   End If
   lstTabIndex.AddItem ControlName(Ctl)

End Sub

Public Sub ControlRenamed(Ctl As VBControl, sOldName As String, lOldIndex As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : ControlRenamed
   ' * Parameters       :
   ' *                    ctl As VBControl
   ' *                    sOldName As String
   ' *                    lOldIndex As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next
   Dim sTmp             As String
   Dim i                As Integer

   If lOldIndex >= 0 Then
      sOldName = sOldName & "(" & lOldIndex & ")"
   End If

   sTmp = ControlName(Ctl)
   For i = 0 To lstTabIndex.ListCount - 1
      If left$(lstTabIndex.List(i), Len(sOldName)) = sOldName Then
         'remove it from the list
         lstTabIndex.RemoveItem i
         'add it back with the new name
         lstTabIndex.AddItem sTmp, i
         Exit Sub
      End If
   Next
   err.Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : Form_KeyDown
   ' * Parameters       :
   ' *                    KeyCode As Integer
   ' *                    Shift As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Pass the keystrokes onto the IDE
   HandleKeyDown Me, KeyCode, Shift

End Sub

Sub GetNameAndIndex(sListItem As String, sName As String, nIndex As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : docTabOrder
   ' * Module Filename  : Taborder.dob
   ' * Procedure Name   : GetNameAndIndex
   ' * Parameters       :
   ' *                    sListItem As String
   ' *                    SName As String
   ' *                    nIndex As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nPos             As Integer
   Dim nPos2            As Integer
   Dim sTmp             As String

   ' *** Strip off the caption if there is one
   nPos = InStr(sListItem, " ")
   If nPos > 0 Then
      sTmp = left$(sListItem, nPos - 1)
   Else
      sTmp = sListItem
   End If

   ' *** Now check for an index
   nPos = InStr(sTmp, "(")
   If nPos > 0 Then
      ' *** Control has an index so we need to
      ' *** Strip it off and save it
      nPos2 = InStr(sTmp, ")")
      nIndex = Val(Mid$(sTmp, nPos + 1, nPos2 - nPos))
      sName = left$(sTmp, nPos - 1)
   Else
      nIndex = -1
      sName = sTmp
   End If

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
      .CreateHwnd Me.hWnd, ttfBalloon
      .MaxTipWidth = 70
      .Icon = itInfoIcon
      .Title = "VBIDEUtils"
      For nI = 0 To Me.Controls.Count
         If Trim$(Me.Controls(nI).ToolTipText) <> "" Then
            If err = 0 Then
               .AddToolHwnd Me.Controls(nI), Me.hWnd, tfTransparent, Me.Controls(nI).ToolTipText
               Me.Controls(nI).ToolTipText = ""
            Else
               err.Clear
            End If
         End If
      Next
   End With

End Sub

