VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form APIfrm 
   Caption         =   "API VIEWER BY VENKATESH ©"
   ClientHeight    =   5850
   ClientLeft      =   2460
   ClientTop       =   1815
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "APIfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5190
      Left            =   45
      TabIndex        =   2
      Top             =   240
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   9155
      _Version        =   327681
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Declares:"
      TabPicture(0)   =   "APIfrm.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Constants:"
      TabPicture(1)   =   "APIfrm.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Types:"
      TabPicture(2)   =   "APIfrm.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Found:"
      TabPicture(3)   =   "APIfrm.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView4"
      Tab(3).ControlCount=   1
      Begin ComctlLib.ListView ListView4 
         Height          =   4665
         Left            =   -74895
         TabIndex        =   6
         Top             =   420
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ImageList4"
         SmallIcons      =   "ImageList4"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ComctlLib.ListView ListView3 
         Height          =   4680
         Left            =   -74880
         TabIndex        =   5
         Top             =   390
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   8255
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ImageList3"
         SmallIcons      =   "ImageList3"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ComctlLib.ListView ListView2 
         Height          =   4635
         Left            =   -74880
         TabIndex        =   4
         Top             =   450
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   8176
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   105
         TabIndex        =   3
         Top             =   420
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   8281
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         NumItems        =   0
      End
   End
   Begin ComctlLib.ProgressBar apiprog 
      Height          =   105
      Left            =   795
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   185
      _Version        =   327682
      Appearance      =   1
      Max             =   10000
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   370
      SimpleText      =   "s"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
      Font            =   "APIfrm.frx":037A
   End
   Begin ComctlLib.ImageList ImageList4 
      Left            =   5745
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList img 
      Left            =   6645
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "APIfrm.frx":039F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "APIfrm.frx":0775
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "APIfrm.frx":0B27
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   5865
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   5130
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4365
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSepView 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "Vie&w"
         Begin VB.Menu mnulist 
            Caption         =   "&List"
         End
         Begin VB.Menu mnudetails 
            Caption         =   "&Details"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSepCopy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Code &Finder"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSepFav 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFavs 
         Caption         =   "&Addtofavorites"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuViewFavs 
         Caption         =   "View &Favorites"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSepPri 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprivate 
         Caption         =   "Pr&ivate"
         Checked         =   -1  'True
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuPublic 
         Caption         =   "P&ublic"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuforOpt 
         Caption         =   "Format &Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSepVB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVB 
         Caption         =   "&Back To VB"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSepHelp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "&Help"
         Begin VB.Menu mnuAbt 
            Caption         =   "&About"
         End
      End
   End
End
Attribute VB_Name = "APIfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itmX As ListItem, imgX As ListImage
Dim loadcnst As Boolean, loadtyp As Boolean, hprev As Boolean
Dim nStyle As Long, nHeader As Long
Dim qd As QueryDef
Private Sub Form_Activate()
   modREG.WriteAppName
   mnuView.Visible = True
   mnuSepView.Visible = True
   If hprev = False Then
      hprev = True
      Set apirs = apidb.OpenRecordset("Declares", dbOpenDynaset)
      ListView1.ColumnHeaders.Add 1, , "Declares", ListView1.Width / 4, lvwColumnLeft
      ListView1.ColumnHeaders.Add 2, , "Code", ListView1.Width / 1, lvwColumnLeft
      ListView1.View = lvwReport
      Set imgX = ImageList1.ListImages.Add(1, , LoadPicture(App.path & "\declare.bmp"))
      ListView1.Icons = ImageList1
      ListView1.SmallIcons = ImageList1
      apirs.MoveFirst
      modAPI.showprogressinstatusbar True, apiprog, APIfrm
      While Not apirs.EOF
         Screen.MousePointer = vbHourglass
         Set itmX = ListView1.ListItems.Add(, , CStr(apirs("Name")))
         itmX.Icon = 1   ' Set an icon from ImageList1.
         itmX.SmallIcon = 1  ' Set an icon from ImageList2.
         apiprog.value = apirs.PercentPosition * 100
         APIfrm.StatusBar1.Panels(1).Text = "Loading Functions...."
         APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
         If Not IsNull(apirs("FullName")) Then
            itmX.SubItems(1) = CStr(apirs("FullName"))
         End If
         apirs.MoveNext
      Wend
      Screen.MousePointer = vbDefault
      apiprog.value = 0
      apiprog.Visible = False
      APIfrm.StatusBar1.Panels(1).Text = "Declares.."
      APIfrm.StatusBar1.Panels.Add 2
      APIfrm.StatusBar1.Panels(2).Text = apirs.RecordCount & " " & "Declares"
      APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
      '----------------~ For  changing the listview style ~-------------------------
      ' get the current ListView style
      nStyle = SendMessage(ListView1.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0)
      nStyle = SendMessage(ListView1.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 41)
      nHeader = SendMessage(ListView1.hwnd, LVM_GETHEADER, 0, ByVal 0&)
      nStyle = GetWindowLong(nHeader, GWL_STYLE)
      nStyle = nStyle And (Not HDS_BUTTONS)
      Call SetWindowLong(nHeader, GWL_STYLE, nStyle)
      '----------------~ For  changing the listview style ~-------------------------
      APIfrm.AddorRemove "Add"
      APIfrm.StatusBar1.Panels(3).Bevel = sbrInset
      APIfrm.StatusBar1.Panels(3).Text = ListView1.ListItems(1).Text
      APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
   End If
End Sub
Private Sub Form_Load()
   Dim ll, lt As Long, retval1 As Boolean
   lt = (Screen.Height \ 2 - Me.Height \ 2) \ Screen.TwipsPerPixelY
   ll = (Screen.Width \ 2 - Me.Width \ 2) \ Screen.TwipsPerPixelX
   APIfrm.StatusBar1.Panels.Add 1
   APIfrm.StatusBar1.Panels(1).Text = "Declares"
   APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
   '==============================~For setting the menu~==============================
   Dim i As Integer
   Dim hMenu, hSubMenu, menuID, x
   hMenu = GetMenu(Me.hwnd)
   '-----------------~ For the first menu item ~---------------------------------------------------------------------------
   hSubMenu = GetSubMenu(hMenu, 0) '1 for "Other" menu etcetera
   For i = 1 To 2
      menuID = GetMenuItemID(hSubMenu, i - 1)
      x = SetMenuItemBitmaps(hMenu, menuID, MF_BYCOMMAND, img.ListImages(i).Picture, img.ListImages(i).Picture)
   Next
   menuID = GetMenuItemID(hSubMenu, 5)
   x = SetMenuItemBitmaps(hMenu, menuID, MF_BYCOMMAND, img.ListImages(3).Picture, img.ListImages(3).Picture)
   '==============================~For setting the menu~==============================

End Sub
Private Sub ListView1_GotFocus()
   APIfrm.StatusBar1.Panels(3).Text = ListView1.ListItems(1).Text
   APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
   APIfrm.StatusBar1.Panels(3).Text = Item.Text
   APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
End Sub
Private Sub ListView1_LostFocus()
   APIfrm.StatusBar1.Panels(3).Text = ""
End Sub
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      PopupMenu mnupopup
   End If
End Sub
Private Sub ListView1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
   DefaultCursors = True
End Sub
Private Sub ListView2_GotFocus()
   APIfrm.StatusBar1.Panels(3).Text = ListView2.ListItems(1).Text
   APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
End Sub
Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
   APIfrm.StatusBar1.Panels(3).Text = Item.Text
   APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
End Sub
Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      PopupMenu mnupopup
   End If
End Sub
Private Sub ListView2_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
   DefaultCursors = True
End Sub
Private Sub ListView3_GotFocus()
   APIfrm.StatusBar1.Panels(1).Text = ListView3.ListItems(1).Text
   APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
End Sub
Private Sub ListView3_ItemClick(ByVal Item As ComctlLib.ListItem)
   APIfrm.StatusBar1.Panels(3).Text = Item.Text
   APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
End Sub
Private Sub ListView3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      PopupMenu mnupopup
   End If
End Sub
Private Sub ListView4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      PopupMenu mnupopup
   End If
End Sub
Private Sub mnuAbt_Click()
   Dim x As Long
   x = ShellAbout(Me.hwnd, "API TEXT VIEWER For", "BY VENKATESH SOFTWARE SOLUTIONS©", 0)
End Sub
Private Sub mnuCopy_Click()
   Dim finque As QueryDef, typefin As Recordset
   Select Case SSTab1.Caption
      Case "Declares:":
         Set finque = apidb.CreateQueryDef("")
         finque.SQL = "Select FullName from Declares where Name=" & "'" & ListView1.SelectedItem.Text & "'"
         Set apirs = finque.OpenRecordset(dbOpenDynaset)
         Clipboard.Clear
         If pubcli = True Then
            apifunc = "Public" & Space(1)
         ElseIf pricli = True Then
            apifunc = "Private" & Space(1)
         ElseIf defcli = True Then
            apifunc = "Private" & Space(1)
         End If
         apifunc = apifunc & apirs.Fields(0)
         Clipboard.SetText apifunc, vbCFText
      Case "Constants:":
         Set finque = apidb.CreateQueryDef("")
         finque.SQL = "Select FullName from Constants where Name=" & "'" & ListView2.SelectedItem.Text & "'"
         Set apicnrs = finque.OpenRecordset(dbOpenDynaset)
         If pubcli = True Then
            apicons = "Public" & Space(1)
         ElseIf pricli = True Then
            apicons = "Private" & Space(1)
         ElseIf defcli = True Then
            apifunc = "Private" & Space(1)
         End If
         apicons = apicons & apicnrs.Fields(0)
         Clipboard.Clear
         Clipboard.SetText apicons, vbCFText
      Case "Types:":
         If pubcli = True Then
            typ = "Public Type " & ListView3.SelectedItem.Text & vbCrLf
         ElseIf pricli = True Then
            typ = "Private Type " & ListView3.SelectedItem.Text & vbCrLf
         ElseIf defcli = True Then
            typ = "Private Type " & ListView3.SelectedItem.Text & vbCrLf
         End If
         Set finque = apidb.CreateQueryDef("")
         finque.SQL = "Select ID from Types where Name=" & "'" & ListView3.SelectedItem.Text & "'"
         Set apityp = finque.OpenRecordset(dbOpenDynaset)
         Set typefin = apidb.OpenRecordset("TypeItems", dbOpenDynaset)
         typefin.MoveFirst
         Do While Not typefin.EOF
            If typefin("TypeID") = apityp.Fields(0) Then
               typ = typ & typefin("TypeItem") & vbCrLf
            End If
            typefin.MoveNext
         Loop
         typ = typ & vbCrLf & "End Type"
         MsgBox Trim(typ)
         Clipboard.Clear
         Clipboard.SetText typ, vbCFText
   End Select
End Sub
Private Sub mnudetails_Click()
   Screen.MousePointer = vbHourglass
   Select Case SSTab1.Caption
      Case "Declares:":
         mnulist.Checked = False
         mnudetails.Checked = True
         ListView1.View = lvwReport
         Screen.MousePointer = vbDefault
      Case "Constants:"
         mnulist.Checked = False
         mnudetails.Checked = True
         ListView2.View = lvwReport
         Screen.MousePointer = vbDefault
      Case "Types:":
         mnulist.Checked = False
         mnudetails.Checked = True
         ListView3.View = lvwReport
         Screen.MousePointer = vbDefault
      Case "Found:"
         mnulist.Checked = False
         mnudetails.Checked = True
         ListView4.View = lvwReport
         Screen.MousePointer = vbDefault
   End Select
End Sub
Private Sub mnuFind_Click()
   Load FINDfrm
   FINDfrm.Show
End Sub
Private Sub mnulist_Click()
   Screen.MousePointer = vbHourglass
   Select Case SSTab1.Caption
      Case "Declares:":
         mnulist.Checked = True
         mnudetails.Checked = False
         ListView1.View = lvwList
         Screen.MousePointer = vbDefault
      Case "Constants:":
         mnulist.Checked = True
         mnudetails.Checked = False
         ListView2.View = lvwList
         Screen.MousePointer = vbDefault
      Case "Types:":
         mnulist.Checked = True
         mnudetails.Checked = False
         ListView3.View = lvwList
         Screen.MousePointer = vbDefault
      Case "Found:"
         mnulist.Checked = True
         mnudetails.Checked = False
         ListView4.View = lvwList
         Screen.MousePointer = vbDefault
   End Select
End Sub
Private Sub mnuPreview_Click()
   Dim finque As QueryDef, typefin As Recordset
   Select Case SSTab1.Caption
      Case "Declares:":
         Set finque = apidb.CreateQueryDef("")
         finque.SQL = "Select FullName from Declares where Name=" & "'" & ListView1.SelectedItem.Text & "'"
         Set apirs = finque.OpenRecordset(dbOpenDynaset)
         If pubcli = True Then
            apifunc = "Public" & Space(1)
         ElseIf pricli = True Then
            apifunc = "Private" & Space(1)
         ElseIf defcli = True Then
            apifunc = "Private" & Space(1)
         End If
         apifunc = apifunc & apirs.Fields(0)
         Load VIEWfrm
         VIEWfrm.Show
         VIEWfrm.APItext.Text = apifunc
      Case "Constants:":
         Set finque = apidb.CreateQueryDef("")
         finque.SQL = "Select FullName from Constants where Name=" & "'" & ListView2.SelectedItem.Text & "'"
         Set apicnrs = finque.OpenRecordset(dbOpenDynaset)
         If pubcli = True Then
            apicons = "Public" & Space(1)
         ElseIf pricli = True Then
            apicons = "Private" & Space(1)
         ElseIf defcli = True Then
            apifunc = "Private" & Space(1)
         End If
         apicons = apicons & apicnrs.Fields(0)
         Load VIEWfrm
         VIEWfrm.Show
         VIEWfrm.APItext.Text = apicons
      Case "Types:":
         If pubcli = True Then
            typ = "Public Type " & ListView3.SelectedItem.Text & vbCrLf
         ElseIf pricli = True Then
            typ = "Private Type " & ListView3.SelectedItem.Text & vbCrLf
         ElseIf defcli = True Then
            typ = "Private Type " & ListView3.SelectedItem.Text & vbCrLf
         End If
         Set finque = apidb.CreateQueryDef("")
         finque.SQL = "Select ID from Types where Name=" & "'" & ListView3.SelectedItem.Text & "'"
         Set apityp = finque.OpenRecordset(dbOpenDynaset)
         Set typefin = apidb.OpenRecordset("TypeItems", dbOpenDynaset)
         typefin.MoveFirst
         Do While Not typefin.EOF
            If typefin("TypeID") = apityp.Fields(0) Then
               typ = typ & vbCrLf & typefin("TypeItem")
            End If
            typefin.MoveNext
         Loop
         typ = typ & vbCrLf & "End Type"
         Load VIEWfrm
         VIEWfrm.Show
         VIEWfrm.APItext.Text = typ
      Case "Found:":
         If FINDfrm.declcat = True Then    '1st if
            If pubcli = True Then
               typ = "Public" & ListView4.SelectedItem.Text & vbCrLf
            ElseIf pricli = True Then
               typ = "Private" & ListView4.SelectedItem.Text & vbCrLf
            ElseIf defcli = True Then
               typ = "Private" & ListView4.SelectedItem.Text & vbCrLf
            End If
            Set finque = apidb.CreateQueryDef("")
            finque.SQL = "Select FullName from Declares where Name=" & "'" & ListView4.SelectedItem.Text & "'"
            Set apirs = finque.OpenRecordset(dbOpenDynaset)
            If pubcli = True Then
               apifunc = "Public" & Space(1)
            ElseIf pricli = True Then
               apifunc = "Private" & Space(1)
            ElseIf defcli = True Then
               apifunc = "Private" & Space(1)
            End If     'end of if for pubcli under declcat
            apifunc = apifunc & apirs.Fields(0)
            Load VIEWfrm
            VIEWfrm.Show
            VIEWfrm.APItext.Text = apifunc
            FINDfrm.declcat = False
         ElseIf FINDfrm.conscat = True Then
            If pubcli = True Then
               typ = "Public" & ListView4.SelectedItem.Text & vbCrLf
            ElseIf pricli = True Then
               typ = "Private" & ListView4.SelectedItem.Text & vbCrLf
            ElseIf defcli = True Then
               typ = "Private" & ListView4.SelectedItem.Text & vbCrLf
            End If
            Set finque = apidb.CreateQueryDef("")
            finque.SQL = "Select FullName from Constants where Name=" & "'" & ListView4.SelectedItem.Text & "'"
            Set apicnrs = finque.OpenRecordset(dbOpenDynaset)
            If pubcli = True Then
               apicons = "Public" & Space(1)
            ElseIf pricli = True Then
               apicons = "Private" & Space(1)
            ElseIf defcli = True Then
               apicons = "Private" & Space(1)
            End If     'end of if for pubcli under declcat
            apicons = apicons & apicnrs.Fields(0)
            Load VIEWfrm
            VIEWfrm.Show
            VIEWfrm.APItext.Text = apicons
            FINDfrm.conscat = False
         ElseIf FINDfrm.typcat = True Then
            If pubcli = True Then
               typ = "Public Type " & ListView4.SelectedItem.Text & vbCrLf
            ElseIf pricli = True Then
               typ = "Private Type " & ListView4.SelectedItem.Text & vbCrLf
            ElseIf defcli = True Then
               typ = "Private Type " & ListView4.SelectedItem.Text & vbCrLf
            End If
            Set finque = apidb.CreateQueryDef("")
            finque.SQL = "Select ID from Types where Name=" & "'" & ListView4.SelectedItem.Text & "'"
            Set apityp = finque.OpenRecordset(dbOpenDynaset)
            Set typefin = apidb.OpenRecordset("TypeItems", dbOpenDynaset)
            typefin.MoveFirst
            Do While Not typefin.EOF
               If typefin("TypeID") = apityp.Fields(0) Then
                  typ = typ & vbCrLf & typefin("TypeItem")
               End If
               typefin.MoveNext
            Loop
            typ = typ & vbCrLf & "End Type"
            Load VIEWfrm
            VIEWfrm.Show
            VIEWfrm.APItext.Text = typ
         End If
   End Select
End Sub
Private Sub mnuprivate_Click()
   mnuprivate.Checked = True
   If mnuPublic.Checked = True Then
      mnuPublic.Checked = False
      mnuprivate.Checked = True
   End If
   pricli = True
   pubcli = False
   defcli = True
End Sub
Private Sub mnuPublic_Click()
   mnuPublic.Checked = True
   If mnuprivate.Checked = True Then
      mnuPublic.Checked = True
      mnuprivate.Checked = False
   End If
   pubcli = True
   pricli = False
   defcli = False
End Sub

Private Sub mnuVB_Click()
   SwitchTo hwnd, "Microsoft Visual Basic"
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case SSTab1.Caption
      Case "Declares:"
         mnuView.Visible = True
         mnuSepView.Visible = True
         APIfrm.StatusBar1.Panels(1).Text = "Declares..."
         APIfrm.StatusBar1.Panels(2).Text = apirs.RecordCount & "  " & "Declares"
         If ListView1.View = lvwReport Then
            mnudetails.Checked = True
            mnulist.Checked = False
         ElseIf ListView1.View = lvwList Then
            mnulist.Checked = True
            mnudetails.Checked = False
         End If
      Case "Constants:":
         mnuView.Visible = True
         mnuSepView.Visible = True
         If loadcnst = False Then
            loadcnst = True
            APIfrm.AddorRemove "Remove"
            APIfrm.StatusBar1.Panels.Remove 2
            Set apidb = OpenDatabase(App.path & "\Win32api.MDB")
            'Set apidb = OpenDatabase("C:\Program Files\Microsoft Visual Studio\Common\Tools\Winapi\Win32api.MDB")
            Set apicnrs = apidb.OpenRecordset("Constants", dbOpenDynaset)
            ListView2.ColumnHeaders.Add , , "Constants", ListView1.Width / 3, lvwColumnLeft
            ListView2.ColumnHeaders.Add , , "Code", ListView1.Width / 1, lvwColumnLeft
            ListView2.View = lvwReport
            Set imgX = ImageList2.ListImages.Add(1, , LoadPicture(App.path & "\const.bmp"))
            ListView2.Icons = ImageList2
            ListView2.SmallIcons = ImageList2
            apicnrs.MoveFirst
            modAPI.showprogressinstatusbar True, apiprog, APIfrm
            While Not apicnrs.EOF
               Screen.MousePointer = vbHourglass
               Set itmX = ListView2.ListItems.Add(, , CStr(apicnrs("Name")))
               itmX.Icon = 1   ' Set an icon from ImageList1.
               itmX.SmallIcon = 1  ' Set an icon from ImageList2.
               apiprog.value = apicnrs.PercentPosition * 100
               APIfrm.StatusBar1.Panels(1).Text = "Loading Constants...."
               APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
               If Not IsNull(apicnrs("Name")) Then
                  itmX.SubItems(1) = CStr(apicnrs("FullName"))
               End If
               apicnrs.MoveNext
            Wend
            Screen.MousePointer = vbDefault
            apiprog.value = 0
            apiprog.Visible = False
            APIfrm.StatusBar1.Panels(1).Text = "Constants.."
            APIfrm.StatusBar1.Panels.Add 2
            APIfrm.StatusBar1.Panels(2).Text = apicnrs.RecordCount & " " & "Constants"
            APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
            '----------------~ For  changing the listview style ~-------------------------
            ' get the current ListView style
            nStyle = SendMessage(ListView2.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0)
            nStyle = SendMessage(ListView2.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 41)
            nHeader = SendMessage(ListView2.hwnd, LVM_GETHEADER, 0, ByVal 0&)
            nStyle = GetWindowLong(nHeader, GWL_STYLE)
            nStyle = nStyle And (Not HDS_BUTTONS)
            Call SetWindowLong(nHeader, GWL_STYLE, nStyle)
            '----------------~ For  changing the listview style ~-------------------------
            APIfrm.AddorRemove "Add"
            APIfrm.StatusBar1.Panels(3).Text = ListView1.ListItems(1).Text
            APIfrm.StatusBar1.Panels(3).Bevel = sbrInset
            APIfrm.StatusBar1.Panels(3).AutoSize = sbrConstents
         Else
            APIfrm.StatusBar1.Panels(1).Text = "Constants.."
            'APIfrm.StatusBar1.Panels.Add 2
            APIfrm.StatusBar1.Panels(2).Text = apicnrs.RecordCount & " " & "Constants"
            APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
         End If
         If ListView2.View = lvwReport Then
            mnudetails.Checked = True
            mnulist.Checked = False
         ElseIf ListView2.View = lvwList Then
            mnulist.Checked = True
            mnudetails.Checked = False
         End If
         ListView2.SetFocus
      Case "Types:":
         If loadtyp = False Then
            loadtyp = True
            APIfrm.AddorRemove "Remove"
            APIfrm.StatusBar1.Panels.Remove 2
            Set apidb = OpenDatabase(App.path & "\Win32api.MDB")
            'Set apidb = OpenDatabase("C:\Program Files\Microsoft Visual Studio\Common\Tools\Winapi\Win32api.MDB")
            Set apityp = apidb.OpenRecordset("Types", dbOpenDynaset)
            ListView3.ColumnHeaders.Add , , "Types", ListView1.Width / 1, lvwColumnLeft
            ListView3.View = lvwList
            Set imgX = ImageList3.ListImages.Add(1, , LoadPicture(App.path & "\type.bmp"))
            ListView3.Icons = ImageList3
            ListView3.SmallIcons = ImageList3
            apityp.MoveFirst
            modAPI.showprogressinstatusbar True, apiprog, APIfrm
            While Not apityp.EOF
               Screen.MousePointer = vbHourglass
               Set itmX = ListView3.ListItems.Add(, , CStr(apityp("Name")))
               itmX.Icon = 1   ' Set an icon from ImageList1.
               itmX.SmallIcon = 1  ' Set an icon from ImageList2.
               apiprog.value = apityp.PercentPosition * 100
               APIfrm.StatusBar1.Panels(1).Text = "Loading Types...."
               APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
               apityp.MoveNext
            Wend
            Screen.MousePointer = vbDefault
            apiprog.value = 0
            apiprog.Visible = False
            APIfrm.StatusBar1.Panels(1).Text = "Types.."
            APIfrm.StatusBar1.Panels.Add 2
            APIfrm.StatusBar1.Panels(2).Text = apityp.RecordCount & " " & "Types"
            APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
            '----------------~ For  changing the listview style ~-------------------------
            ' get the current ListView style
            nStyle = SendMessage(ListView3.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0)
            nStyle = SendMessage(ListView3.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 41)
            nHeader = SendMessage(ListView3.hwnd, LVM_GETHEADER, 0, ByVal 0&)
            nStyle = GetWindowLong(nHeader, GWL_STYLE)
            nStyle = nStyle And (Not HDS_BUTTONS)
            Call SetWindowLong(nHeader, GWL_STYLE, nStyle)
            '----------------~ For  changing the listview style ~-------------------------
            APIfrm.AddorRemove "Add"
            APIfrm.StatusBar1.Panels(3).Text = ListView3.ListItems(1).Text
            APIfrm.StatusBar1.Panels(3).Bevel = sbrInset
            APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
         Else
            APIfrm.StatusBar1.Panels(1).Text = "Types.."
            APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
            APIfrm.StatusBar1.Panels(2).Text = apityp.RecordCount & " " & "Types"
            APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
         End If
         mnuSepView.Visible = False
         mnuView.Visible = False
         ListView3.SetFocus
      Case "Found:"
         If APIfrm.ListView4.ListItems.count > 0 Then
            APIfrm.StatusBar1.Panels(1).Text = APIfrm.ListView4.ListItems(1).Text
            APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
         Else
            APIfrm.StatusBar1.Panels(1).Text = ""
            APIfrm.StatusBar1.Panels(2).Text = ""
            APIfrm.StatusBar1.Panels(3).Text = ""
         End If
   End Select
End Sub
Public Function AddorRemove(whatodo As String)
   Select Case whatodo
      Case "Add":
         APIfrm.StatusBar1.Panels.Add 3
      Case "Remove":
         APIfrm.StatusBar1.Panels.Remove 3
   End Select
End Function
Private Sub SSTab1_GotFocus()
   On Error Resume Next
   Select Case SSTab1.Tab
      Case 0:
         APIfrm.StatusBar1.Panels(1).Text = "Declares.."
         APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
         APIfrm.StatusBar1.Panels(2).Text = apirs.RecordCount
         APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
         APIfrm.StatusBar1.Panels(3).Text = APIfrm.ListView1.SelectedItem.Text
         APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
      Case 1:
         APIfrm.StatusBar1.Panels(1).Text = "Constants.."
         APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
         APIfrm.StatusBar1.Panels(2).Text = apicnrs.RecordCount
         APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
         APIfrm.StatusBar1.Panels(3).Text = APIfrm.ListView2.SelectedItem.Text
         APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
      Case 2:
         APIfrm.StatusBar1.Panels(1).Text = "Types.."
         APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
         APIfrm.StatusBar1.Panels(2).Text = apityp.RecordCount
         APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
         APIfrm.StatusBar1.Panels(3).Text = APIfrm.ListView3.SelectedItem.Text
         APIfrm.StatusBar1.Panels(3).AutoSize = sbrContents
      Case 3:
         If APIfrm.ListView4.ListItems.count > 0 Then
            APIfrm.StatusBar1.Panels(1).Text = "Last Query" & Space(1) & APIfrm.ListView4.SelectedItem.Text
            APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
         Else
            APIfrm.StatusBar1.Panels(1).Text = ""
            APIfrm.StatusBar1.Panels(2).Text = ""
            APIfrm.StatusBar1.Panels(3).Text = ""
         End If
   End Select
End Sub
