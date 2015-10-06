VERSION 5.00
Object = "{D1DAC785-7BF2-42C1-9915-A540451B87F2}#1.1#0"; "VBIDEUtils1.ocx"
Begin VB.Form frmAccelerator 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Accelerator Assistant"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8295
   Icon            =   "Accelerator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefresh 
      Height          =   330
      Left            =   0
      Picture         =   "Accelerator.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Refresh the list of controls and their accelerators"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.OptionButton optLetter 
      Caption         =   "A"
      Enabled         =   0   'False
      Height          =   330
      Index           =   65
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   220
   End
   Begin vbAcceleratorGrid.vbalGrid grdGrid 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "List of controls and their accelerators"
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4683
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "frmAccelerator"
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
' * Module Name      : frmAccelerator
' * Module Filename  : Accelerator.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Dim mcmpCurrentForm     As VBComponent      'current form
Dim mcolCtls            As VBControls       'form's controls
Private m_nEditCol      As Integer
Private bGridClick      As Boolean

Private Sub cmdRefresh_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : cmdRefresh_Click
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Call RefreshList

End Sub

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   Call InitTooltips

   ' *** Create all the optionbox for the accelerator keys
   Dim nI               As Integer
   Dim nIndex           As Integer
   Dim nLeft            As Integer

   nIndex = 1
   nLeft = optLetter(Asc("A")).left + optLetter(Asc("A")).Width
   optLetter(Asc("A")).BackColor = &H8000000F
   optLetter(Asc("A")).ForeColor = vbBlack
   optLetter(Asc("A")).Visible = False
   optLetter(Asc("A")).Enabled = True

   ' *** Load letters
   For nI = Asc("B") To Asc("Z")
      Load optLetter(nI)
      ' *** Set location of new option button
      optLetter(nI).Caption = Chr$(nI)
      optLetter(nI).top = optLetter(Asc("A")).top
      optLetter(nI).left = nLeft
      optLetter(nI).BackColor = &H8000000F
      optLetter(nI).ForeColor = vbBlack
      nLeft = nLeft + optLetter(Asc("A")).Width
      optLetter(nI).Visible = False
      optLetter(nI).Enabled = True
   Next

   ' *** Load numbers
   For nI = Asc("0") To Asc("9")
      Load optLetter(nI)
      ' *** Set location of new option button
      optLetter(nI).Caption = Chr$(nI)
      optLetter(nI).top = optLetter(Asc("A")).top
      optLetter(nI).left = nLeft
      optLetter(nI).BackColor = &H8000000F
      optLetter(nI).ForeColor = vbBlack
      nLeft = nLeft + optLetter(Asc("A")).Width
      optLetter(nI).Visible = False
      optLetter(nI).Enabled = True
   Next

   cmdRefresh.ToolTipText = "Refresh the list from the form."

   Me.Show
   DoEvents

   Call RefreshList

End Sub

Private Sub Form_Resize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * E-Mail           : Me
   ' * Procedure Name   : Form_Resize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   grdGrid.Move grdGrid.left, grdGrid.top, ScaleWidth - (grdGrid.left * 2), ScaleHeight - (grdGrid.top)

End Sub

Function ControlName(Ctl As VBIDE.VBControl) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : ControlName
   ' * Parameters       :
   ' *                    Ctl As VBIDE.VBControl
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

Public Sub RefreshList()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : RefreshList
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo RefreshListErr

   Dim Ctl              As VBControl
   Dim Ctl2             As VBControl
   Dim Ctl3             As VBControl
   Dim sControlName     As String
   Dim sIndex           As String
   Dim sInfo            As String
   Dim sTmp             As String
   Dim nI               As Integer
   Dim nTabIndex        As Integer
   Dim nTmp             As Integer

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   If InRunMode(VBInstance) Then Exit Sub

   If VBInstance.ActiveVBProject Is Nothing Then Exit Sub

   ' *** load the component
   Set mcmpCurrentForm = VBInstance.SelectedVBComponent

   ' *** check to see if we have a valid component
   If mcmpCurrentForm Is Nothing Then
      Exit Sub
   End If

   ' *** make sure the active component is a form, user control or property page
   If (mcmpCurrentForm.Type <> vbext_ct_VBForm) And _
      (mcmpCurrentForm.Type <> vbext_ct_UserControl) And _
      (mcmpCurrentForm.Type <> vbext_ct_DocObject) And _
      (mcmpCurrentForm.Type <> vbext_ct_PropPage) Then
      Exit Sub
   End If

   ' *** Init the grid
   Call InitGrid

   On Error Resume Next
   For nI = Asc("0") To Asc("Z")
      optLetter(nI).Value = False
      optLetter(nI).Enabled = False
      If optLetter(nI).Visible Then
         If optLetter(nI).ForeColor = vbBlack Then
            optLetter(nI).Visible = False
         End If
      End If
   Next

   ' *** Turn redraw off for speed:
   grdGrid.Redraw = False

   ' *** Clear the grid
   grdGrid.Clear

   Set mcolCtls = mcmpCurrentForm.Designer.VBControls
   For Each Ctl In mcmpCurrentForm.Designer.VBControls
      sControlName = Ctl.Properties!Name
      sIndex = Ctl.Properties!index

      ' *** Try to get the caption
      On Error Resume Next
      sInfo = Ctl.Properties!Caption
      If err Then
         ' *** Doesn't have a caption
         err.Clear
         ' *** Go to next control
         GoTo SkipIt

      End If
      On Error GoTo RefreshListErr
      grdGrid.Rows = grdGrid.Rows + 1
      grdGrid.CellDetails grdGrid.Rows, 1, sInfo, DT_LEFT, , &H80000018, vbBlue

      ' *** Get the accelerator key
      sTmp = ""
      For nI = 1 To Len(sInfo)
         If Mid$(sInfo, nI, 1) = "&" Then
            If nI < Len(sInfo) Then
               If Mid$(sInfo, nI + 1, 1) = "&" Then
                  nI = nI + 1
               Else
                  sTmp = UCase$(Mid$(sInfo, nI + 1, 1))
                  On Error Resume Next
                  If optLetter(Asc(sTmp)).ForeColor = vbBlack Then
                     ' *** Selected letter
                     optLetter(Asc(sTmp)).ForeColor = vbBlue
                     optLetter(Asc(sTmp)).Enabled = False
                     optLetter(Asc(sTmp)).Visible = True
                  ElseIf optLetter(Asc(sTmp)).ForeColor = vbBlue Then
                     ' *** Conflict letter
                     optLetter(Asc(sTmp)).ForeColor = vbRed
                     optLetter(Asc(sTmp)).Enabled = False
                     optLetter(Asc(sTmp)).Visible = True
                  End If
                  optLetter(Asc(sTmp)).Value = False
                  On Error GoTo RefreshListErr
                  Exit For
               End If
            End If
         End If
      Next
      grdGrid.CellDetails grdGrid.Rows, 2, sTmp, DT_CENTER, , vbWhite, vbRed

      ' *** Set the control name
      If sIndex <> "-1" Then
         grdGrid.CellDetails grdGrid.Rows, 3, sControlName & "(" & sIndex & ")", DT_LEFT, , vbWhite, vbBlue
      Else
         grdGrid.CellDetails grdGrid.Rows, 3, sControlName, DT_LEFT, , vbWhite, vbBlue
      End If

      ' *** Set the tabindex
      On Error Resume Next
      nTmp = Ctl.Properties!TabIndex
      If err Then
         ' *** Doesn't have a tabindex
         err.Clear
      Else
         grdGrid.CellDetails grdGrid.Rows, 4, nTmp, DT_CENTER, , &H80000018, vbRed
      End If
      On Error GoTo RefreshListErr

      ' *** Get the next tabindex control
      If Ctl.ClassName = "Label" Then
         On Error Resume Next
         nTabIndex = Ctl.Properties!TabIndex
         Set Ctl3 = Nothing
         For Each Ctl2 In mcmpCurrentForm.Designer.VBControls
            If Ctl2.ClassName <> "Label" Then
               nTmp = Ctl2.Properties!TabIndex
               If err Then
                  ' *** Doesn't have a tabindex
                  err.Clear
               ElseIf Ctl2.Properties!TabIndex > nTabIndex Then
                  If Ctl3 Is Nothing Then
                     Set Ctl3 = Ctl2
                  Else
                     If Ctl2.Properties!TabIndex < Ctl3.Properties!TabIndex Then
                        Set Ctl3 = Ctl2
                     End If
                  End If
               End If
            End If
         Next
         If Not (Ctl3 Is Nothing) Then grdGrid.CellDetails grdGrid.Rows, 5, Ctl3.Properties!Name, DT_LEFT, , &H80000018, vbBlue
         On Error GoTo RefreshListErr
      End If
      grdGrid.Redraw = True
      grdGrid.Redraw = False

SkipIt:
   Next

Exit_RefreshList:
   grdGrid.Width = ScaleWidth - (grdGrid.left * 2)
   grdGrid.Height = ScaleHeight - (cmdRefresh.Height + 100)

   grdGrid.Redraw = True

   Exit Sub
RefreshListErr:
   'MsgBox Err.Description
   GoTo Exit_RefreshList

End Sub

Private Sub GetNameAndIndex(sListItem As String, sName As String, nIndex As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : GetNameAndIndex
   ' * Parameters       :
   ' *                    sListItem As String
   ' *                    sName As String
   ' *                    nIndex As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *  This sub extracts the control name and index
   ' *  from the formatted list item
   ' *
   ' **********************************************************************

   Dim nPos             As Integer
   Dim nPos2            As Integer
   Dim sTmp             As String

   'strip off the caption if there is one
   nPos = InStr(sListItem, " ")
   If nPos > 0 Then
      sTmp = left$(sListItem, nPos - 1)
   Else
      sTmp = sListItem
   End If

   'now check for an index
   nPos = InStr(sTmp, "(")
   If nPos > 0 Then
      'control has an index so we need to
      'strip it off and save it
      nPos2 = InStr(sTmp, ")")
      nIndex = Val(Mid$(sTmp, nPos + 1, nPos2 - nPos))
      sName = left$(sTmp, nPos - 1)
   Else
      nIndex = -1
      sName = sTmp
   End If

End Sub

Private Sub InitGrid()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 11:42
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : InitGrid
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   With grdGrid
      ' *** Turn redraw off for speed:
      .Redraw = False
      .Clear

      .RemoveColumn (1)
      .RemoveColumn (2)

      .BackColor = &H80000018

      .AddColumn "Caption", "Caption", ecgHdrTextALignLeft, , 80
      .AddColumn "Accelerator", "Accelerator", ecgHdrTextALignCentre, , 70
      .AddColumn "Control Name", "Control Name", ecgHdrTextALignLeft, , 100
      .AddColumn "Tab Index", "Tab Index", ecgHdrTextALignCentre, , 65
      .AddColumn "Next Tab Stop", "Next Tab Stop", ecgHdrTextALignLeft, , 100

      .SetHeaders

      ' *** Ensure the grid will draw!
      .Redraw = True

   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 10/02/2000
   ' * Time             : 10:36
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
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
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
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

Private Sub grdGrid_ColumnClick(ByVal lCol As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 8/02/99
   ' * Time             : 13:59
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : grdGrid_ColumnClick
   ' * Parameters       :
   ' *                    ByVal lCol As Long
   ' **********************************************************************
   ' * Comments         : Sort by a specific column
   ' *
   ' *
   ' **********************************************************************

   Dim sTag             As String

   Dim cHourglass       As class_Hourglass
   Set cHourglass = New class_Hourglass

   With grdGrid.SortObject
      .Clear
      .SortColumn(1) = lCol

      sTag = grdGrid.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(1) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(1) = CCLOrderDescending
      End If
      grdGrid.ColumnTag(lCol) = sTag

      ' *** Sort by text
      .SortType(1) = CCLSortString

   End With
   grdGrid.Sort

End Sub

Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/29/2001
   ' * Time             : 10:52
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : grdGrid_KeyDown
   ' * Parameters       :
   ' *                    KeyCode As Integer
   ' *                    Shift As Integer
   ' *                    bDoDefault As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If KeyCode = vbKeyDelete Then
      ' *** Delete the current shortcut

      Dim sCaption         As String
      Dim sTmp             As String
      Dim sIndex           As String

      If bGridClick Then Exit Sub

      ' *** Get the caption and existing accelerator
      sCaption = grdGrid.CellFormattedText(grdGrid.SelectedRow, 1)
      sTmp = grdGrid.CellFormattedText(grdGrid.SelectedRow, 2)

      If sTmp <> "" Then
         ' *** Remove all previous accelerator
         sCaption = Replace(sCaption, "&&", "VBIDAMOND IS THE BEST")
         sCaption = Replace(sCaption, "&", "")
         sCaption = Replace(sCaption, "VBIDAMOND IS THE BEST", "&&")
      End If

      ' *** Update the grid
      grdGrid.CellText(grdGrid.SelectedRow, 1) = sCaption
      grdGrid.CellText(grdGrid.SelectedRow, 2) = ""
      err.Clear

      ' *** Update the letters
      Call grdGrid_SelectionChange(grdGrid.SelectedRow, 1)

      ' *** Update the control
      sTmp = grdGrid.CellFormattedText(grdGrid.SelectedRow, 3)
      sIndex = GetStringBetweenTags(sTmp, "(", ")")
      sTmp = Replace(sTmp, "(" & sIndex & ")", "")
      Set mcolCtls = mcmpCurrentForm.Designer.VBControls
      If Not IsNumeric(sIndex) Then
         mcolCtls(sTmp).Properties!Caption = sCaption
      Else
         mcolCtls(sTmp, sIndex).Properties!Caption = sCaption
      End If

   End If

End Sub

Private Sub grdGrid_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/29/2001
   ' * Time             : 10:52
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : grdGrid_SelectionChange
   ' * Parameters       :
   ' *                    ByVal lRow As Long
   ' *                    ByVal lCol As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim nI               As Integer
   Dim sCaption         As String

   sCaption = grdGrid.CellFormattedText(lRow, 1)

   bGridClick = True

   SetRedraw Me, False

   On Error Resume Next
   For nI = Asc("0") To Asc("Z")
      optLetter(nI).Value = False
      optLetter(nI).Enabled = False
      optLetter(nI).Visible = False
      optLetter(nI).ForeColor = vbBlack
   Next

   ' *** Set the accelerator keys
   For nI = 1 To grdGrid.Rows
      sTmp = grdGrid.CellFormattedText(nI, 2)
      On Error Resume Next
      If optLetter(Asc(sTmp)).ForeColor = vbBlack Then
         ' *** Selected letter
         optLetter(Asc(sTmp)).ForeColor = vbBlue
         optLetter(Asc(sTmp)).Enabled = False
         optLetter(Asc(sTmp)).Visible = True
      ElseIf optLetter(Asc(sTmp)).ForeColor = vbBlue Then
         ' *** Conflict letter
         optLetter(Asc(sTmp)).ForeColor = vbRed
         optLetter(Asc(sTmp)).Enabled = False
         optLetter(Asc(sTmp)).Visible = True
      End If
      optLetter(Asc(sTmp)).Value = False
   Next

   ' *** Get the accelerator key$
   sTmp = ""
   For nI = 1 To Len(sCaption)
      If Mid$(sCaption, nI, 1) = "&" Then
         If nI < Len(sCaption) Then
            If Mid$(sCaption, nI + 1, 1) = "&" Then
               nI = nI + 1
            Else
               sTmp = UCase$(Mid$(sCaption, nI + 1, 1))
               optLetter(Asc(sTmp)).Value = True
               optLetter(Asc(sTmp)).Visible = True
            End If
         End If
      Else
         On Error Resume Next
         optLetter(Asc(UCase$(Mid$(sCaption, nI, 1)))).Enabled = True
         optLetter(Asc(UCase$(Mid$(sCaption, nI, 1)))).Visible = True
      End If
   Next

   SetRedraw Me, True

   bGridClick = False

End Sub

Private Sub optLetter_Click(index As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/29/2001
   ' * Time             : 10:52
   ' * Module Name      : frmAccelerator
   ' * Module Filename  : Accelerator.frm
   ' * Procedure Name   : optLetter_Click
   ' * Parameters       :
   ' *                    index As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sCaption         As String
   Dim nI               As Integer
   Dim sTmp             As String
   Dim sIndex           As String

   If bGridClick Then Exit Sub

   ' *** Get the caption and existing accelerator
   sCaption = grdGrid.CellFormattedText(grdGrid.SelectedRow, 1)
   sTmp = grdGrid.CellFormattedText(grdGrid.SelectedRow, 2)

   If sTmp <> "" Then
      ' *** Remove all previous accelerator
      sCaption = Replace(sCaption, "&&", "VBIDAMOND IS THE BEST")
      sCaption = Replace(sCaption, "&", "")
      sCaption = Replace(sCaption, "VBIDAMOND IS THE BEST", "&&")
   End If

   For nI = 1 To Len(sCaption)
      If UCase$(Mid$(sCaption, nI, 1)) = optLetter(index).Caption Then
         If nI = 1 Then
            sCaption = "&" & sCaption
         Else
            sCaption = left$(sCaption, nI - 1) & "&" & right(sCaption, Len(sCaption) - nI + 1)
         End If
         Exit For
      End If
   Next

   ' *** Update the grid
   grdGrid.CellText(grdGrid.SelectedRow, 1) = sCaption
   grdGrid.CellText(grdGrid.SelectedRow, 2) = optLetter(index).Caption
   If optLetter(index).ForeColor = vbBlue Then
      optLetter(index).ForeColor = vbRed
   ElseIf optLetter(index).ForeColor = vbBlack Then
      optLetter(index).ForeColor = vbBlue
   End If
   err.Clear

   ' *** Update the letters
   Call grdGrid_SelectionChange(grdGrid.SelectedRow, 1)

   ' *** Update the control
   sTmp = grdGrid.CellFormattedText(grdGrid.SelectedRow, 3)
   sIndex = GetStringBetweenTags(sTmp, "(", ")")
   sTmp = Replace(sTmp, "(" & sIndex & ")", "")
   Set mcolCtls = mcmpCurrentForm.Designer.VBControls
   If Not IsNumeric(sIndex) Then
      mcolCtls(sTmp).Properties!Caption = sCaption
   Else
      mcolCtls(sTmp, sIndex).Properties!Caption = sCaption
   End If

End Sub

