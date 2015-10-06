VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKeyCode 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Keycode Assistant"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5145
   Icon            =   "KeyCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvViewKeyCode 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "List of all the keycode. Double-click with the mouse to put in the clipboard the constant..."
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2778
      View            =   3
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
Attribute VB_Name = "frmKeyCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 09/02/2000
' * Time             : 15:05
' * Module Name      : frmKeyCode
' * Module Filename  : KeyCode.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private clsTooltips     As New class_Tooltips

Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const LVM_FIRST = &H1000                   '// ListView messages
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54) '// optional wParam == mask

Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_FULLROWSELECT = &H20         '// applies to report mode only

Private Sub Form_Load()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 15:07
   ' * Module Name      : frmKeyCode
   ' * Module Filename  : KeyCode.frm
   ' * Procedure Name   : Form_Load
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Call InitTooltips

   lvViewKeyCode.left = 0
   lvViewKeyCode.top = 0

   Call Form_Resize
   Call Fill_Listview

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 15:07
   ' * Module Name      : frmKeyCode
   ' * Module Filename  : KeyCode.frm
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

Private Sub lvViewKeyCode_DblClick()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 15:05
   ' * Module Name      : frmKeyCode
   ' * Module Filename  : KeyCode.frm
   ' * Procedure Name   : lvViewKeyCode_DblClick
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Call Clipboard.Clear
   Call Clipboard.SetText(lvViewKeyCode.SelectedItem.Text, vbCFText)

End Sub

Private Sub Form_Resize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 15:05
   ' * Module Name      : frmKeyCode
   ' * Module Filename  : KeyCode.frm
   ' * Procedure Name   : Form_Resize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   lvViewKeyCode.Width = ScaleWidth - (lvViewKeyCode.left * 2)
   lvViewKeyCode.Height = ScaleHeight

End Sub

Private Sub Fill_Listview()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 09/02/2000
   ' * Time             : 15:05
   ' * Module Name      : frmKeyCode
   ' * Module Filename  : KeyCode.frm
   ' * Procedure Name   : Fill_Listview
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Fill in the listview

   Dim li               As ListItem
   Dim nStyle           As Long

   Call lvViewKeyCode.ColumnHeaders.Add(, "Constant", "Constant", "1300")
   Call lvViewKeyCode.ColumnHeaders.Add(, "Value", "Value", "600")
   Call lvViewKeyCode.ColumnHeaders.Add(, "Description", "Description", "2200")

   Set li = lvViewKeyCode.ListItems.Add(, "VbKeyLButton", " VbKeyLButton"): li.SubItems(1) = "0x1": li.SubItems(2) = "Left mouse button"
   Set li = lvViewKeyCode.ListItems.Add(, "VbKeyRButton", " VbKeyRButton"): li.SubItems(1) = "0x2": li.SubItems(2) = "Right mouse button"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyCancel", " vbKeyCancel"): li.SubItems(1) = "0x3": li.SubItems(2) = "CANCEL key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyMButton", " vbKeyMButton"): li.SubItems(1) = "0x4": li.SubItems(2) = "Middle mouse button"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyBack", " vbKeyBack"): li.SubItems(1) = "0x8": li.SubItems(2) = "BACKSPACE key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyTab", " vbKeyTab"): li.SubItems(1) = "0x9": li.SubItems(2) = "TAB key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyClear", " vbKeyClear"): li.SubItems(1) = "0xC": li.SubItems(2) = "CLEAR key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyReturn", " vbKeyReturn"): li.SubItems(1) = "0xD": li.SubItems(2) = "ENTER key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyShift", " vbKeyShift"): li.SubItems(1) = "0x10": li.SubItems(2) = "SHIFT key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyControl", " vbKeyControl"): li.SubItems(1) = "0x11": li.SubItems(2) = "CTRL key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyMenu", " vbKeyMenu"): li.SubItems(1) = "0x12": li.SubItems(2) = "MENU key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyPause", " vbKeyPause"): li.SubItems(1) = "0x13": li.SubItems(2) = "PAUSE key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyCapital", " vbKeyCapital"): li.SubItems(1) = "0x14": li.SubItems(2) = "CAPS LOCK key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyEscape", " vbKeyEscape"): li.SubItems(1) = "0x1B": li.SubItems(2) = "ESC key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeySpace", " vbKeySpace"): li.SubItems(1) = "0x20": li.SubItems(2) = "SPACEBAR key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyPageUp", " vbKeyPageUp"): li.SubItems(1) = "0x21": li.SubItems(2) = "PAGE UP key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyPageDown", " vbKeyPageDown"): li.SubItems(1) = "0x22": li.SubItems(2) = "PAGE DOWN key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyEnd", " vbKeyEnd"): li.SubItems(1) = "0x23": li.SubItems(2) = "END key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyHome", " vbKeyHome"): li.SubItems(1) = "0x24": li.SubItems(2) = "HOME key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyLeft", " vbKeyLeft"): li.SubItems(1) = "0x25": li.SubItems(2) = "LEFT ARROW key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyUp", " vbKeyUp"): li.SubItems(1) = "0x26": li.SubItems(2) = "UP ARROW key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyRight", " vbKeyRight"): li.SubItems(1) = "0x27": li.SubItems(2) = "RIGHT ARROW key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyDown", " vbKeyDown"): li.SubItems(1) = "0x28": li.SubItems(2) = "DOWN ARROW key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeySelect", " vbKeySelect"): li.SubItems(1) = "0x29": li.SubItems(2) = "SELECT key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyPrint", " vbKeyPrint"): li.SubItems(1) = "0x2A": li.SubItems(2) = "PRINT SCREEN key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyExecute", " vbKeyExecute"): li.SubItems(1) = "0x2B": li.SubItems(2) = "EXECUTE key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeySnapshot", " vbKeySnapshot"): li.SubItems(1) = "0x2C": li.SubItems(2) = "SNAPSHOT key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyInsert", " vbKeyInsert"): li.SubItems(1) = "0x2D": li.SubItems(2) = "INSERT key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyDelete", " vbKeyDelete"): li.SubItems(1) = "0x2E": li.SubItems(2) = "DELETE key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyHelp", " vbKeyHelp"): li.SubItems(1) = "0x2F": li.SubItems(2) = "HELP key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumlock", " vbKeyNumlock"): li.SubItems(1) = "0x90": li.SubItems(2) = "NUM LOCK key"

   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyA", " vbKeyA"): li.SubItems(1) = "65": li.SubItems(2) = "A key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyB", " vbKeyB"): li.SubItems(1) = "66": li.SubItems(2) = "B key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyC", " vbKeyC"): li.SubItems(1) = "67": li.SubItems(2) = "C key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyD", " vbKeyD"): li.SubItems(1) = "68": li.SubItems(2) = "D key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyE", " vbKeyE"): li.SubItems(1) = "69": li.SubItems(2) = "E key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF", " vbKeyF"): li.SubItems(1) = "70": li.SubItems(2) = "F key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyG", " vbKeyG"): li.SubItems(1) = "71": li.SubItems(2) = "G key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyH", " vbKeyH"): li.SubItems(1) = "72": li.SubItems(2) = "H key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyI", " vbKeyI"): li.SubItems(1) = "73": li.SubItems(2) = "I key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyJ", " vbKeyJ"): li.SubItems(1) = "74": li.SubItems(2) = "J key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyK", " vbKeyK"): li.SubItems(1) = "75": li.SubItems(2) = "K key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyL", " vbKeyL"): li.SubItems(1) = "76": li.SubItems(2) = "L key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyM", " vbKeyM"): li.SubItems(1) = "77": li.SubItems(2) = "M key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyN", " vbKeyN"): li.SubItems(1) = "78": li.SubItems(2) = "N key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyO", " vbKeyO"): li.SubItems(1) = "79": li.SubItems(2) = "O key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyP", " vbKeyP"): li.SubItems(1) = "80": li.SubItems(2) = "P key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyQ", " vbKeyQ"): li.SubItems(1) = "81": li.SubItems(2) = "Q key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyR", " vbKeyR"): li.SubItems(1) = "82": li.SubItems(2) = "R key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyS", " vbKeyS"): li.SubItems(1) = "83": li.SubItems(2) = "S key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyT", " vbKeyT"): li.SubItems(1) = "84": li.SubItems(2) = "T key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyU", " vbKeyU"): li.SubItems(1) = "85": li.SubItems(2) = "U key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyV", " vbKeyV"): li.SubItems(1) = "86": li.SubItems(2) = "V key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyW", " vbKeyW"): li.SubItems(1) = "87": li.SubItems(2) = "W key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyX", " vbKeyX"): li.SubItems(1) = "88": li.SubItems(2) = "X key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyY", " vbKeyY"): li.SubItems(1) = "89": li.SubItems(2) = "Y key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyZ", " vbKeyZ"): li.SubItems(1) = "90": li.SubItems(2) = "Z key"

   Set li = lvViewKeyCode.ListItems.Add(, "vbKey0", " vbKey0"): li.SubItems(1) = "48": li.SubItems(2) = "0 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey1", " vbKey1"): li.SubItems(1) = "49": li.SubItems(2) = "1 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey2", " vbKey2"): li.SubItems(1) = "50": li.SubItems(2) = "2 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey3", " vbKey3"): li.SubItems(1) = "51": li.SubItems(2) = "3 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey4", " vbKey4"): li.SubItems(1) = "52": li.SubItems(2) = "4 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey5", " vbKey5"): li.SubItems(1) = "53": li.SubItems(2) = "5 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey6", " vbKey6"): li.SubItems(1) = "54": li.SubItems(2) = "6 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey7", " vbKey7"): li.SubItems(1) = "55": li.SubItems(2) = "7 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey8", " vbKey8"): li.SubItems(1) = "56": li.SubItems(2) = "8 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKey9", " vbKey9"): li.SubItems(1) = "57": li.SubItems(2) = "9 key"

   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad0", " vbKeyNumpad0"): li.SubItems(1) = "0x60": li.SubItems(2) = "0 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad1", " vbKeyNumpad1"): li.SubItems(1) = "0x61": li.SubItems(2) = "1 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad2", " vbKeyNumpad2"): li.SubItems(1) = "0x62": li.SubItems(2) = "2 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad3", " vbKeyNumpad3"): li.SubItems(1) = "0x63": li.SubItems(2) = "3 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad4", " vbKeyNumpad4"): li.SubItems(1) = "0x64": li.SubItems(2) = "4 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad5", " vbKeyNumpad5"): li.SubItems(1) = "0x65": li.SubItems(2) = "5 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad6", " vbKeyNumpad6"): li.SubItems(1) = "0x66": li.SubItems(2) = "6 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad7", " vbKeyNumpad7"): li.SubItems(1) = "0x67": li.SubItems(2) = "7 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad8", " vbKeyNumpad8"): li.SubItems(1) = "0x68": li.SubItems(2) = "8 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyNumpad9", " vbKeyNumpad9"): li.SubItems(1) = "0x69": li.SubItems(2) = "9 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyMultiply", " vbKeyMultiply"): li.SubItems(1) = "0x6A": li.SubItems(2) = "MULTIPLICATION SIGN (*) key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyAdd", " vbKeyAdd"): li.SubItems(1) = "0x6B": li.SubItems(2) = "PLUS SIGN (+) key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeySeparator", " vbKeySeparator"): li.SubItems(1) = "0x6C": li.SubItems(2) = "ENTER key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeySubtract", " vbKeySubtract"): li.SubItems(1) = "0x6D": li.SubItems(2) = "MINUS SIGN (–) key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyDecimal", " vbKeyDecimal"): li.SubItems(1) = "0x6E": li.SubItems(2) = "DECIMAL POINT (.) key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyDivide", " vbKeyDivide"): li.SubItems(1) = "0x6F": li.SubItems(2) = "DIVISION SIGN (/) key"

   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF1", " vbKeyF1"): li.SubItems(1) = "0x70": li.SubItems(2) = "F1 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF2", " vbKeyF2"): li.SubItems(1) = "0x71": li.SubItems(2) = "F2 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF3", " vbKeyF3"): li.SubItems(1) = "0x72": li.SubItems(2) = "F3 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF4", " vbKeyF4"): li.SubItems(1) = "0x73": li.SubItems(2) = "F4 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF5", " vbKeyF5"): li.SubItems(1) = "0x74": li.SubItems(2) = "F5 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF6", " vbKeyF6"): li.SubItems(1) = "0x75": li.SubItems(2) = "F6 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF7", " vbKeyF7"): li.SubItems(1) = "0x76": li.SubItems(2) = "F7 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF8", " vbKeyF8"): li.SubItems(1) = "0x77": li.SubItems(2) = "F8 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF9", " vbKeyF9"): li.SubItems(1) = "0x78": li.SubItems(2) = "F9 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF10", " vbKeyF10"): li.SubItems(1) = "0x79": li.SubItems(2) = "F10 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF11", " vbKeyF11"): li.SubItems(1) = "0x7A": li.SubItems(2) = "F11 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF12", " vbKeyF12"): li.SubItems(1) = "0x7B": li.SubItems(2) = "F12 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF13", " vbKeyF13"): li.SubItems(1) = "0x7C": li.SubItems(2) = "F13 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF14", " vbKeyF14"): li.SubItems(1) = "0x7D": li.SubItems(2) = "F14 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF15", " vbKeyF15"): li.SubItems(1) = "0x7E": li.SubItems(2) = "F15 key"
   Set li = lvViewKeyCode.ListItems.Add(, "vbKeyF16", " vbKeyF16"): li.SubItems(1) = "0x7F": li.SubItems(2) = "F16 key"

   nStyle = SendMessageByLong(lvViewKeyCode.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
   SendMessageByLong lvViewKeyCode.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, nStyle Or LVS_EX_FULLROWSELECT

   lvViewKeyCode.SelectedItem = lvViewKeyCode.ListItems("VbKeyLButton")

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
