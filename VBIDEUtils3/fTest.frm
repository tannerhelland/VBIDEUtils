VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMenuTest 
   Caption         =   "vbAccelerator Popup Menu Component"
   ClientHeight    =   4635
   ClientLeft      =   4245
   ClientTop       =   2715
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   5415
   Begin VB.CommandButton cmdTest2 
      Caption         =   "&Test"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   8820
      Width           =   1155
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   8400
      Width           =   1155
   End
   Begin VB.CommandButton cmdAccelTest 
      Caption         =   "&Accelerator"
      Height          =   375
      Left            =   4020
      TabIndex        =   5
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Checks"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   660
      Width           =   1155
   End
   Begin VB.CommandButton cmdVBAccel 
      Height          =   375
      Left            =   1380
      Picture         =   "fTest.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Connect to vbAccelerator - the VB Programmer's Resource"
      Top             =   660
      Width           =   1275
   End
   Begin VB.ListBox lstStatus 
      Height          =   3375
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Right click to get an Edit popup menu"
      Top             =   1080
      Width           =   5355
   End
   Begin VB.CommandButton cmdNewMenu 
      Caption         =   "&Show Menu"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Click to show a demonstration menu with sub levels."
      Top             =   660
      Width           =   1215
   End
   Begin VB.PictureBox picBackground 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   1800
      Picture         =   "fTest.frx":08D1
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   1920
   End
   Begin ComctlLib.ImageList ilsIcons16 
      Left            =   120
      Top             =   4500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   43
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":C913
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":CC2D
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":CF47
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":D261
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":D57B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":D895
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":DBAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":DEC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":E1E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":E4FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":E817
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":EB31
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":EE4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":F165
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":F47F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":F799
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":FAB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":FDCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":100E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":10401
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":1071B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":10A35
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":10D4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":11069
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":11383
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":1169D
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":119B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":11CD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":11FEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":12305
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":1261F
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":12939
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":12C53
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":12F6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":13287
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":135A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":138BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":13BD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":13EEF
            Key             =   "Web"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":14209
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":14523
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":1483D
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fTest.frx":14B57
            Key             =   "vbAccelerator"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Click one of the buttons below or Right Click in the list box to demonstrate unlimited Popup-menus with icons."
      Height          =   555
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   5295
   End
   Begin VB.Menu mnuF0MAIN 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   3
      End
   End
   Begin VB.Menu mnuEditTOP 
      Caption         =   "&Edit"
      Index           =   0
      Begin VB.Menu mnuEdit 
         Caption         =   "Cu&t"
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&vbAccelerator on the Web"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Add vbAccelerator Active &Channel..."
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About.."
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMenuTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cP As cPopupMenu
Attribute cP.VB_VarHelpID = -1
Private Const mcWEBSITE = -&H8000&
Private Sub Status(ByVal sMsg As String)
   lstStatus.AddItem sMsg
   lstStatus.ListIndex = lstStatus.NewIndex
End Sub


Private Sub cmdAccelTest_Click()
Dim iIndex As Long
   
   With cP
      Set .BackgroundPicture = picBackground.Picture
      .Restore "AccelTest"
      iIndex = .IndexForKey("mnuAccel(3)")
      '.Enabled(iIndex) = Not (.Enabled(iIndex))
      iIndex = .ShowPopupMenu( _
         cmdAccelTest.Left, cmdAccelTest.Top + cmdAccelTest.Height, _
         cmdAccelTest.Left, cmdAccelTest.Top, cmdAccelTest.Left + cmdAccelTest.Width, cmdAccelTest.Top + cmdAccelTest.Height _
            )
      .Store "AccelTest"
      .ClearBackgroundPicture
   End With
End Sub

Private Sub cmdCheck_Click()
Dim iIndex As Long

   With cP
      .Restore "CheckTest"
      iIndex = .ShowPopupMenu( _
         cmdCheck.Left, cmdCheck.Top + cmdCheck.Height, _
         cmdCheck.Left, cmdCheck.Top, cmdCheck.Left + cmdCheck.Width, cmdCheck.Top + cmdCheck.Height _
            )
      If iIndex > 0 Then
         If InStr(.ItemKey(iIndex), "Option") <> 0 Then
            .GroupToggle iIndex
         ElseIf InStr(.ItemKey(iIndex), "Check") <> 0 Then
            .Checked(iIndex) = Not (.Checked(iIndex))
         End If
         .Store "CheckTest"
      End If
   End With
End Sub

Private Sub cmdNewMenu_Click()
Dim iIndex As Long
   With cP
      .Restore "Demo"
      .Caption(1) = "Test Modify Caption"
      .Caption(6) = "Test Modify Caption 2"
      .Default(6) = True
      iIndex = .ShowPopupMenu( _
         cmdNewMenu.Left, cmdNewMenu.Top + cmdNewMenu.Height, _
         cmdNewMenu.Left, cmdNewMenu.Top, cmdNewMenu.Left + cmdNewMenu.Width, cmdNewMenu.Top + cmdNewMenu.Height _
            )
      If (iIndex > 0) Then
         Status "Selected Item=" & iIndex
         If (.ItemKey(iIndex) = "CHECK") Then
           .Checked(iIndex) = Not (.Checked(iIndex))
           .Store "Demo"
         End If
      End If
   End With
End Sub

Private Sub cmdTest_Click()
Dim i As Long
Dim j As Long
Dim k As Long
Dim n As Long
   With cP
      .Clear
      For i = 1 To 10
         i = .AddItem("Test" & i, , , , , , , "Test" & i)
      Next i
      For i = 1 To 10
         k = .InsertItem("InsTest" & i, "Test3", , , , , , "Test" & j)
         If i = 3 Then
            For j = 1 To 10
               .AddItem "SubTest" & j, , , k, , , , "SubTest" & j
            Next j
         End If
      Next i
      k = .InsertItem("InsTOP", "Test1", , , , , , "InsTOP")
      k = .AddItem("InsTopSub 1", , , k, , , , "InsTopSub1")
      For j = 1 To 4
         .InsertItem "InsTopSub " & j + 1, "InsTopSub1", , , , , , "InsTopSub" & j + 1
      Next j
      k = .InsertItem("InsBOTTOM", "Test10", , , , , , "InsBOTTOM")
      For j = 1 To 5
         .AddItem "InsBottom" & j, , , k, , , , "InsBottom" & j
      Next j
      
      .ShowPopupMenu 0, 0
      
      .ClearSubMenusOfItem "InsTOP"
      k = .IndexForKey("InsTOP")
      For j = 1 To 24
         .AddItem "InsTopSub " & j, , , k, , , , "InsTopSub" & j
      Next j
      .ClearSubMenusOfItem "InsBOTTOM"
      k = .IndexForKey("InsTopSub20")
      For j = 1 To 24
         i = .AddItem("InsTopSubSub " & j, , , k, , , , "InsTopSubSub" & j)
         If j Mod 5 = 0 Then
            For n = 1 To Rnd * 8 + 4
               .AddItem "Testing" & n, , , i
            Next n
         End If
      Next j
      
      .ShowPopupMenu 0, 0
      
      
      
      
   End With
End Sub

Private Sub cmdTest2_Click()
   With cP
      .RestoreFromFile , "C:\Stevemac\VB\Controls\vbalTbar\Menu.dat"
      .Restore "Main"
      .ShowPopupMenu 0, 0
   End With
End Sub

Private Sub cmdVBAccel_Click()
Dim iIndex As Long
   With cP
      .Restore "vbAccelerator"
      iIndex = .ShowPopupMenu( _
         cmdVBAccel.Left, cmdVBAccel.Top + cmdVBAccel.Height, _
         cmdVBAccel.Left, cmdVBAccel.Top, cmdVBAccel.Left + cmdVBAccel.Width, cmdVBAccel.Top + cmdVBAccel.Height _
            )
      If (iIndex > 0) Then
         Status "Selected Item=" & iIndex
         If (.ItemKey(iIndex) = "Web") Then
            mnuHelp_Click 0
         ElseIf (.ItemKey(iIndex) = "Channel") Then
            mnuHelp_Click 1
         ElseIf (.ItemData(iIndex) = mcWEBSITE) Then
            Screen.MousePointer = vbHourglass
            ShellEx .ItemKey(iIndex)
            Screen.MousePointer = vbDefault
         End If
      End If
   End With
End Sub

Private Sub cP_Click(ItemNumber As Long)
   Status "Clicked Item=" & ItemNumber
End Sub

Private Sub cP_InitPopupMenu(ParentItemNumber As Long)
   Status "InitPopupMenu with Parent= " & ParentItemNumber
End Sub

Private Sub cP_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
   Status "Highlighted  Item=" & ItemNumber & ", Enabled=" & bEnabled & ", Separator = " & bSeparator
End Sub

Private Sub cP_MenuExit()
   Status "Menu Exited."
End Sub


Private Sub Form_Load()
   Set cP = New cPopupMenu
   ' Make sure the ImageList has icons before setting
   ' this if it is a MS ImageList:
   cP.ImageList = ilsIcons16
   ' Make sure you set this up before trying any menus
   '
   cP.hWndOwner = Me.hWnd
   
   ' Cool!
   cP.GradientHighlight = True
   
   ' Create some menus and store them:
   CreateMenus
   
End Sub
Private Sub CreateMenus()
Dim i As Long
Dim j As Long
Dim iIndex As Long
Dim lIcon As Long
Dim sKey As String
Dim sCap As String
   
   ' Create the demo menu:
   With cP
      .Clear
      For i = 1 To 10
       If (i = 6) Or (i = 7) Then sKey = "CHECK" Else sKey = ""
         iIndex = .AddItem("Test " & i, , i, , i + 3, ((i = 6) Or (i = 7)), ((i Mod 3) <> 0), sKey)
         If (i = 5) Then
            For j = 1 To 30
               sCap = "SubMenu Test" & j
               If ((j - 1) Mod 10) = 0 And j > 1 Then
                  sCap = "|" & sCap
               End If
               .AddItem sCap, , , iIndex, j + 10
            Next j
         End If
         If (i = 4) Or (i = 5) Then
            .AddItem "-"
         End If
      Next i
      .Store "Demo"
      
      ' Create the edit menu:
      .Clear
      .AddItem "Cu&t" & vbTab & "Ctrl+X", , , , ilsIcons16.ListImages("CUT").Index - 1, , , "Cut"
      .AddItem "&Copy" & vbTab & "Ctrl+C", , , , ilsIcons16.ListImages("COPY").Index - 1, , , "Copy"
      .AddItem "&Paste" & vbTab & "Ctrl+V", , , , ilsIcons16.ListImages("PASTE").Index - 1, , False, "Paste"
      .Store "Edit"
      
      ' Create the vbAccelerator menu:
      .Clear
      .AddItem "vbAccelerator"
      .Header(1) = True
      lIcon = ilsIcons16.ListImages("vbAccelerator").Index - 1
      .AddItem "&vbAccelerator on the Web..." & vbTab & "F1", , , , lIcon, , , "Web"
      .Default(2) = True
      lIcon = ilsIcons16.ListImages("Web").Index - 1
      .AddItem "Add vbAccelerator Active &Channel...", , mcWEBSITE, , lIcon, , , "Channel"
      .AddItem "-Other sites"
      i = .AddItem("VB Sites", , , , lIcon)
      .AddItem "-VB Sites", , , i
      .AddItem "Goffredo's VB Page", , mcWEBSITE, i, lIcon, , , "http://www.cs.utexas.edu/users/gglaze/vb.htm"
      .AddItem "Advanced Visual Basic WebBoard", , mcWEBSITE, i, lIcon, , , "http://webboard.duke.net:8080/~avb/"
      .AddItem "VBNet", , mcWEBSITE, i, lIcon, , , "http://www.mvps.org/mvps"
      .AddItem "CCRP", , mcWEBSITE, i, lIcon, , , "http://www.mvps.org/ccrp"
      .AddItem "DevX", , mcWEBSITE, i, lIcon, , , "http://www.devx.com/"
      i = .AddItem("Technology", , , , lIcon)
      .AddItem "-Games", , , i
      .AddItem "Dave's Classics", , mcWEBSITE, i, lIcon, , , "http://www.davesclassics.com/"
      .AddItem "Future Gamer", , mcWEBSITE, i, lIcon, , , "http://www.futuregamer.com/"
      .AddItem "-Web Site Building", , , i
      .AddItem "Builder.com", , mcWEBSITE, i, lIcon, , , "http://www.builder.com/"
      .AddItem "The Web Design Resource", , mcWEBSITE, i, lIcon, , , "http://www.pageresource.com/"
      .AddItem "Web Review", , mcWEBSITE, i, lIcon, , , "http://www.webreview.com/"
      .AddItem "-Downloads", , , i
      .AddItem "Tucows", , mcWEBSITE, i, lIcon, , , "http://tucows.cableinet.net/"
      .AddItem "WinFiles.com", , mcWEBSITE, i, lIcon, , , "http://www.winfiles.com/"
      i = .AddItem("Searching and Other", , , , lIcon)
      .AddItem "-Pick'n'Mix", , , i
      .AddItem "The SCHWA Corporation", , mcWEBSITE, i, lIcon, , , "http://www.theschwacorporation.com/"
      .AddItem "Art Cars", , mcWEBSITE, i, lIcon, , , "http://www.artcars.com/"
      .AddItem "The Onion", , mcWEBSITE, i, lIcon, , , "http://www.theonion.com/"
      .AddItem "Virtues of a Programmer", i, mcWEBSITE, i, lIcon, , , "http://www.hhhh.org/wiml/virtues.html"
      .AddItem "-Search", , , i
      .AddItem "HotBot", , mcWEBSITE, i, lIcon, , , "http://www.hotbot.com/"
      .AddItem "DogPile", , mcWEBSITE, i, lIcon, , , "http://www.dogpile.com/"
      .Store "vbAccelerator"
      
      .Clear
      .AddItem "First Check", , , , , True, , "Check1"
      .AddItem "Second Check", , , , , , , "Check2"
      .AddItem "Third Check", , , , , , , "Check3"
      .AddItem "-"
      i = .AddItem("First Option", , , , , , , "Option1")
      .RadioCheck(i) = True
      .AddItem "Second Option", , , , , , , "Option2"
      .AddItem "Third Option", , , , , , , "Option3"
      .AddItem "Fourth Option", , , , , , , "Option4"
      .AddItem "-"
      .AddItem "&vbAccelerator on the Web...", , , , lIcon, , , "Web"
      .Store "CheckTest"
      
      .Clear
      .AddItem "&Back" & vbTab & "Alt+Left Arrow", , , , , , , "mnuAccel(0)"
      .AddItem "&Next" & vbTab & "Alt+Right Arrow", , , , , , , "mnuAccel(1)"
      .AddItem "-"
      .AddItem "&Home Page" & vbTab & "Alt+Home", , , , , , , "mnuAccel(3)"
      .AddItem "&Search the Web", , , , , , , "mnuAccel(4)"
      .AddItem "-"
      .AddItem "&Mail", , , , , , , "mnuAccel(6)"
      .AddItem "&News", , , , , , , "mnuAccel(7)"
      .AddItem "My &Computer", , , , , , , "mnuAccel(8)"
      .AddItem "A&ddress Book", , , , , , , "mnuAccel(9)"
      .AddItem "Ca&lendar", , , , , , , "mnuAccel(10)"
      .AddItem "&Internet Call", , , , , , , "mnuAccel(11)"
      .Store "AccelTest"
   End With
   
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstStatus.Move lstStatus.Left, lstStatus.Top, Me.ScaleWidth - lstStatus.Left * 2, Me.ScaleHeight - lstStatus.Top - 4 * Screen.TwipsPerPixelY
End Sub

Private Sub lstStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (Button And vbRightButton) = vbRightButton Then
      Dim iIndex As Long
         cP.Restore "Edit"
         cP.Enabled(cP.IndexForKey("Paste")) = Clipboard.GetFormat(vbCFText)
         iIndex = cP.ShowPopupMenu( _
            X + lstStatus.Left, Y + lstStatus.Top)
         If (iIndex > 0) Then
            Status "Clicked " & iIndex
         End If
   End If
End Sub

Private Sub mnuEdit_Click(Index As Integer)
   MsgBox "Clicked VB Menu " & mnuEdit(Index).Caption
End Sub

Private Sub mnuFile_Click(Index As Integer)
Dim sFile As String
   Select Case Index
   Case 0
      ' Demonstrates Deserialising menu:
      sFile = App.Path & "\Test.Dat"
      cP.RestoreFromFile , sFile
   Case 1
      ' Demonstrates Serialising menu:
      sFile = App.Path & "\Test.Dat"
      cP.StoreToFile , sFile
   Case 3
      Unload Me
   End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
   Select Case Index
   Case 0
      ' vbAccelerator!
      Screen.MousePointer = vbHourglass
      ShellEx "http://vbaccelerator.com", , , , , Me.hWnd
      Screen.MousePointer = vbDefault
   Case 1
      ' Add vbAccelerator Active Channel
      Screen.MousePointer = vbHourglass
      ShellEx "http://vbaccelerator.com/vbaccel.cdf", , , , , Me.hWnd
      Screen.MousePointer = vbDefault
   Case 3
      ' About
      frmAbout.Show vbModal, Me
   End Select
End Sub
