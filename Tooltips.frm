VERSION 5.00
Object = "{D1DAC785-7BF2-42C1-9915-A540451B87F2}#1.1#0"; "VBIDEUtils1.ocx"
Begin VB.Form frmTooltips 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Control properties assistant"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9390
   Icon            =   "Tooltips.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefresh 
      Height          =   330
      Left            =   0
      Picture         =   "Tooltips.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Refresh the list of controls and all properties"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin vbAcceleratorGrid.vbalGrid grdGrid 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "List of controls and properties. Double click with the mouse to edit a property"
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4683
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
      Begin VB.TextBox tbEdit 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "Tooltips.frx":0544
         Top             =   600
         Visible         =   0   'False
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmTooltips"
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
' * Module Name      : frmTooltips
' * Module Filename  :
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

Private Sub cmdRefresh_Click()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
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
   ' * Module Name      : frmTooltips
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

   Call RefreshList

   cmdRefresh.ToolTipText = "Refresh the list from the form."

End Sub

Private Sub Form_Resize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
   ' * E-Mail           : Me
   ' * Procedure Name   : Form_Resize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   grdGrid.Width = ScaleWidth - (grdGrid.left * 2)
   grdGrid.Height = ScaleHeight - (cmdRefresh.Height + 100)
End Sub

Function ControlName(Ctl As VBIDE.VBControl) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
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

Public Sub RefreshList()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
   ' * Procedure Name   : RefreshList
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo RefreshListErr

   Dim Ctl              As VBControl
   Dim sControlName     As String
   Dim sInfo            As String
   Dim nInfo            As Long

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

   ' *** Turn redraw off for speed:
   grdGrid.Redraw = False

   ' *** Clear the grid
   grdGrid.Clear

   Set mcolCtls = mcmpCurrentForm.Designer.VBControls
   For Each Ctl In mcmpCurrentForm.Designer.VBControls
      sControlName = ControlName(Ctl)
      grdGrid.Rows = grdGrid.Rows + 1

      ' *** Add it to the grid
      grdGrid.CellDetails grdGrid.Rows, 1, sControlName, DT_LEFT, , vbWhite, vbRed

      ' *** Try to get the caption
      On Error Resume Next
      sInfo = Ctl.Properties!Caption
      If err Then
         ' *** Doesn't have a caption
         err.Clear
      End If
      On Error GoTo RefreshListErr
      grdGrid.CellDetails grdGrid.Rows, 2, sInfo, DT_LEFT, , &H80000018, vbBlue

      ' *** Try to get the backcolor
      On Error Resume Next
      nInfo = Ctl.Properties!BackColor
      If err Then
         ' *** Doesn't have a backcolor
         err.Clear
      Else
         grdGrid.CellDetails grdGrid.Rows, 3, "", DT_LEFT, , nInfo, nInfo
      End If
      On Error GoTo RefreshListErr

      ' *** Try to get the ToolTipText
      On Error Resume Next
      sInfo = Ctl.Properties!ToolTipText
      If err Then
         ' *** Doesn't have a ToolTipText
         err.Clear
      End If
      On Error GoTo RefreshListErr
      grdGrid.CellDetails grdGrid.Rows, 4, sInfo, DT_LEFT, , &H80000018, vbBlue

      ' *** Try to get the HelpID
      On Error Resume Next
      nInfo = Ctl.Properties!HelpContextID
      If err Then
         ' *** Doesn't have a HelpID
         err.Clear
      Else
         grdGrid.CellDetails grdGrid.Rows, 5, nInfo, DT_LEFT, , &H80000018, vbGreen
      End If
      On Error GoTo RefreshListErr

SkipIt:
   Next
   grdGrid.Width = ScaleWidth - (grdGrid.left * 2)
   grdGrid.Height = ScaleHeight - (cmdRefresh.Height + 100)

   grdGrid.Redraw = True

   Exit Sub
RefreshListErr:
   'MsgBox Err.Description
   GoTo SkipIt

End Sub

Sub GetNameAndIndex(sListItem As String, sName As String, nIndex As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 28/10/1999
   ' * Time             : 11:42
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
   ' * Procedure Name   : GetNameAndIndex
   ' * Parameters       :
   ' *                    sListItem As String
   ' *                    SName As String
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
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
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

      .AddColumn "Control", "Control", ecgHdrTextALignLeft, , 200
      .AddColumn "Caption", "Caption", ecgHdrTextALignLeft, , 200
      .AddColumn "BackColor", "BackColor", ecgHdrTextALignLeft, , 70
      .AddColumn "Tooltips", "Tooltips", ecgHdrTextALignLeft, , 200
      .AddColumn "HelpID", "HelpID", ecgHdrTextALignLeft, , 50

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
   ' * Module Name      : frmTooltips
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

Private Sub tbEdit_KeyDown(KeyCode As Integer, Shift As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:36
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
   ' * Procedure Name   : tbEdit_KeyDown
   ' * Parameters       :
   ' *                    KeyCode As Integer
   ' *                    Shift As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_tbEdit_KeyDown

   Dim sTmp             As String
   Dim nCtlArrIndex     As Integer

   If (KeyCode = vbKeyReturn) Then
      ' *** Commit edit
      grdGrid.CellText(grdGrid.SelectedRow, grdGrid.SelectedCol) = tbEdit.Text

      If m_nEditCol = 2 Then
         ' **** Set the new Caption
         On Error Resume Next
         GetNameAndIndex grdGrid.Cell(grdGrid.SelectedRow, 1).Text, sTmp, nCtlArrIndex
         mcmpCurrentForm.Designer.VBControls.Item(sTmp, nCtlArrIndex).Properties!Caption = grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).Text
         If err.number <> 0 Then grdGrid.CellText(grdGrid.SelectedRow, m_nEditCol).Text = ""
      End If

      If m_nEditCol = 4 Then
         ' **** Set the new tooltips
         On Error Resume Next
         GetNameAndIndex grdGrid.Cell(grdGrid.SelectedRow, 1).Text, sTmp, nCtlArrIndex
         mcmpCurrentForm.Designer.VBControls.Item(sTmp, nCtlArrIndex).Properties!ToolTipText = grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).Text
         If err.number <> 0 Then grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).Text = ""
      End If

      If m_nEditCol = 5 Then
         ' **** Set the new HepContextID
         On Error Resume Next
         GetNameAndIndex grdGrid.Cell(grdGrid.SelectedRow, 1).Text, sTmp, nCtlArrIndex
         mcmpCurrentForm.Designer.VBControls.Item(sTmp, nCtlArrIndex).Properties!HelpContextID = grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).Text
         If err.number <> 0 Then grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).Text = ""
      End If

      tbEdit.Visible = False
      grdGrid.SetFocus
   ElseIf (KeyCode = vbKeyEscape) Then
      ' *** Cancel edit
      tbEdit.Visible = False
      grdGrid.SetFocus
   Else
   End If

EXIT_tbEdit_KeyDown:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_tbEdit_KeyDown:
   Resume EXIT_tbEdit_KeyDown

End Sub

Private Sub tbEdit_KeyPress(KeyAscii As Integer)

   If grdGrid.SelectedCol = 5 Then
      Call InputNumeric(KeyAscii, tbEdit, False)
   End If

End Sub

Private Sub tbEdit_LostFocus()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:36
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
   ' * Procedure Name   : tbEdit_LostFocus
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   tbEdit.Visible = False
   grdGrid.CancelEdit

End Sub

Private Sub grdGrid_ColumnWidthChanging(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:48
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
   ' * Procedure Name   : grdGrid_ColumnWidthChanging
   ' * Parameters       :
   ' *                    ByVal lCol As Long
   ' *                    ByVal lWidth As Long
   ' *                    bCancel As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lLeft            As Long
   Dim lTop             As Long
   Dim lHeight          As Long

   If tbEdit.Visible Then
      grdGrid.CellBoundary grdGrid.SelectedRow, lCol, lLeft, lTop, lWidth, lHeight
      tbEdit.Width = lWidth
   End If

End Sub

Private Sub grdGrid_ColumnWidthStartChange(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:48
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
   ' * Procedure Name   : grdGrid_ColumnWidthStartChange
   ' * Parameters       :
   ' *                    ByVal lCol As Long
   ' *                    ByVal lWidth As Long
   ' *                    bCancel As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lLeft            As Long
   Dim lTop             As Long
   Dim lHeight          As Long

   If tbEdit.Visible Then
      grdGrid.CellBoundary grdGrid.SelectedRow, lCol, lLeft, lTop, lWidth, lHeight
      tbEdit.Width = lWidth
   End If

End Sub

Private Sub grdGrid_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/10/99
   ' * Time             : 13:34
   ' * Module Name      : frmTooltips
   ' * Module Filename  :
   ' * Procedure Name   : grdGrid_RequestEdit
   ' * Parameters       :
   ' *                    ByVal lRow As Long
   ' *                    ByVal lCol As Long
   ' *                    ByVal iKeyAscii As Integer
   ' *                    bCancel As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lLeft            As Long
   Dim lTop             As Long
   Dim lWidth           As Long
   Dim lHeight          As Long

   Dim sTmp             As String
   Dim nCtlArrIndex     As Integer
   Dim nColor           As Long

   If lCol > 1 Then
      grdGrid.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
      m_nEditCol = lCol
      If m_nEditCol = 3 Then
         ' *** Choose the color
         frmColor.Move Me.left + lLeft, Me.top + lTop
         Call AlwaysOnTop(frmColor, True)
         nColor = grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).BackColor
         nColor = frmColor.ChooseColor(nColor)
         If nColor = -1 Then Exit Sub
         On Error Resume Next
         If grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).BackColor <> nColor Then grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).BackColor = nColor
         If m_nEditCol = 3 Then
            ' **** Set the new BackColor
            On Error Resume Next
            GetNameAndIndex grdGrid.Cell(grdGrid.SelectedRow, 1).Text, sTmp, nCtlArrIndex
            mcmpCurrentForm.Designer.VBControls.Item(sTmp, nCtlArrIndex).Properties!BackColor = grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).BackColor
            If err.number <> 0 Then grdGrid.Cell(grdGrid.SelectedRow, m_nEditCol).BackColor = &H80000018
         End If

      Else
         If Not IsMissing(grdGrid.CellText(lRow, lCol)) Then
            If grdGrid.SelectedCol = 5 Then Call InputNumeric(iKeyAscii, tbEdit, False)
            If iKeyAscii < 32 Then
               tbEdit.Text = grdGrid.CellFormattedText(lRow, lCol)
               tbEdit.SelStart = 0
               tbEdit.SelLength = Len(tbEdit.Text)
            Else
               tbEdit.Text = Chr$(iKeyAscii)
               tbEdit.SelStart = 9999
            End If
         Else
            tbEdit.Text = ""
         End If
         Set tbEdit.Font = grdGrid.CellFont(lRow, lCol)
         tbEdit.Move lLeft, lTop, lWidth, lHeight
         tbEdit.Visible = True
         tbEdit.ZOrder
         tbEdit.SetFocus
      End If
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

