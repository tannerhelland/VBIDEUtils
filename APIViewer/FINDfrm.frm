VERSION 5.00
Begin VB.Form FINDfrm 
   Caption         =   "To Find Declares,Constants,Types..."
   ClientHeight    =   1500
   ClientLeft      =   2775
   ClientTop       =   3720
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FINDfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6555
   Begin VB.Frame Frame1 
      Caption         =   "Enter The Name Of the Declare,Constant,Type you want to find:"
      Height          =   1395
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   6435
      Begin VB.ComboBox finwha 
         Height          =   315
         Left            =   1020
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   3165
      End
      Begin VB.CommandButton fincan 
         Caption         =   "C&ancel"
         Height          =   360
         Left            =   4560
         TabIndex        =   6
         Top             =   870
         Width           =   1635
      End
      Begin VB.CommandButton finapi 
         Caption         =   "&Find"
         Default         =   -1  'True
         Height          =   360
         Left            =   4560
         TabIndex        =   5
         Top             =   390
         Width           =   1635
      End
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   105
         TabIndex        =   2
         Top             =   750
         Width           =   4065
         Begin VB.OptionButton opt2 
            Caption         =   "All Category"
            Height          =   240
            Left            =   2085
            TabIndex        =   4
            Top             =   225
            Width           =   1815
         End
         Begin VB.OptionButton opt1 
            Caption         =   "In Category"
            Height          =   195
            Left            =   75
            TabIndex        =   3
            Top             =   240
            Width           =   1800
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Find What:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   795
      End
   End
End
Attribute VB_Name = "FINDfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public opt1cli As Boolean, opt2cli As Boolean, declcat As Boolean, conscat As Boolean, typcat As Boolean, decfnd As Boolean, cnstfnd As Boolean, typfnd As Boolean, defnd As Boolean, cofnd As Boolean, tyfnd As Boolean
Private Sub finapi_Click()
   Static cmcount As Integer
   Dim valunam As String
   decfnd = True
   cnstfnd = True
   typfnd = True
   defnd = True
   cofnd = True
   tyfnd = True
   found = 0
   Dim itmX As ListItem, imgX As ListImage
   If Len(finwha.Text) <> 0 Then
      If opt1cli = True Then
         Select Case APIfrm.SSTab1.Caption
            Case "Declares:"
               declcat = True
               APIfrm.ListView4.ColumnHeaders(1).Text = "Declares:"
               APIfrm.ListView4.ColumnHeaders(2).Text = "Code"
               If APIfrm.ListView4.ListItems.count > 0 Then
                  APIfrm.ListView4.ListItems.Clear
               End If
               Set apirs = apidb.OpenRecordset("Declares", dbOpenDynaset)
               apirs.MoveFirst
               While Not apirs.EOF
                  If apirs("Name") = finwha.Text Then
                     Set itmX = APIfrm.ListView4.ListItems.Add(, , CStr(apirs("Name")))
                     itmX.Icon = 1   ' Set an icon from ImageList1.
                     itmX.SmallIcon = 1  ' Set an icon from ImageList2.
                     If Not IsNull(apirs("FullName")) Then
                        itmX.SubItems(1) = CStr(apirs("FullName"))
                     End If
                     hkey = HKEY_LOCAL_MACHINE
                     subkey = "SOFTWARE\"
                     newkey = "API TEXT VIEWER BY VENKATESH" & Chr(174) & "\" & "MRUAPIs"
                     retval = RegOpenKeyEx(hkey, subkey & newkey, 0, KEY_ALL_ACCESS, phkResult)
                     If retval = ERROR_SUCCESS Then
                        cmcount = cmcount + 1
                        nvn = "String" & CStr(cmcount)
                        value = FINDfrm.finwha.Text
                        retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
                        RegCloseKey phkResult
                        APIfrm.StatusBar1.Panels(1).Text = "Last Query:" & Space(2) & APIfrm.ListView4.ListItems(1).Text
                        APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
                        APIfrm.StatusBar1.Panels(2).Text = "Declares:"
                        APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
                        APIfrm.StatusBar1.Panels(3).Text = ""
                     End If
                     Call fincan_Click
                     APIfrm.SSTab1.Tab = 3
                     Exit Sub
                  Else
                     defnd = False
                  End If
                  apirs.MoveNext
               Wend
               If defnd = False Then
                  MsgBox "No matches found for" & Space(2) & Trim(finwha.Text), vbCritical + vbSystemModal, App.Title
                  Unload Me
               End If
            Case "Constants:"
               conscat = True
               APIfrm.ListView4.ColumnHeaders(1).Text = "Constants:"
               APIfrm.ListView4.ColumnHeaders(2).Text = "Code:"
               If APIfrm.ListView4.ListItems.count > 0 Then
                  APIfrm.ListView4.ListItems.Clear
               End If
               Set apicnrs = apidb.OpenRecordset("Constants", dbOpenDynaset)
               apicnrs.MoveFirst
               While Not apicnrs.EOF
                  If Trim(apicnrs("Name")) = Trim(finwha.Text) Then
                     Set itmX = APIfrm.ListView4.ListItems.Add(, , CStr(apicnrs("Name")))
                     itmX.Icon = 2   ' Set an icon from ImageList1.
                     itmX.SmallIcon = 2  ' Set an icon from ImageList2.
                     If Not IsNull(apicnrs("FullName")) Then
                        itmX.SubItems(1) = CStr(apicnrs("FullName"))
                     End If
                     hkey = HKEY_LOCAL_MACHINE
                     subkey = "SOFTWARE\"
                     newkey = "API TEXT VIEWER BY VENKATESH" & Chr(174) & "\" & "MRUAPIs"
                     retval = RegOpenKeyEx(hkey, subkey & newkey, 0, KEY_ALL_ACCESS, phkResult)
                     If retval = ERROR_SUCCESS Then
                        cmcount = cmcount + 1
                        nvn = "String" & CStr(cmcount)
                        value = FINDfrm.finwha.Text
                        retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
                        RegCloseKey phkResult
                        APIfrm.StatusBar1.Panels(1).Text = "Last Query:" & Space(2) & APIfrm.ListView4.ListItems(1).Text
                        APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
                        APIfrm.StatusBar1.Panels(2).Text = "Constants:"
                        APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
                        APIfrm.StatusBar1.Panels(3).Text = ""
                     End If
                     APIfrm.SSTab1.Tab = 3
                     Call fincan_Click
                     Exit Sub
                  Else
                     cofnd = False
                  End If
                  apicnrs.MoveNext
               Wend
               If cofnd = False Then
                  MsgBox "Nomatches found for" & Space(2) & Trim(finwha.Text), vbCritical + vbSystemModal, App.Title
                  Unload Me
               End If
            Case "Types:"
               typcat = True
               APIfrm.ListView4.ColumnHeaders(1).Text = "Types:"
               'APIfrm.ListView4.ColumnHeaders(2).Width = 0
               If APIfrm.ListView4.ListItems.count > 0 Then
                  APIfrm.ListView4.ListItems.Clear
               End If
               Set apityp = apidb.OpenRecordset("Types", dbOpenDynaset)
               apityp.MoveFirst
               While Not apityp.EOF
                  If Trim(apityp("Name")) = Trim(finwha.Text) Then
                     Set itmX = APIfrm.ListView4.ListItems.Add(, , CStr(apityp("Name")))
                     itmX.Icon = 2   ' Set an icon from ImageList1.
                     itmX.SmallIcon = 2  ' Set an icon from ImageList2.
                     hkey = HKEY_LOCAL_MACHINE
                     subkey = "SOFTWARE\"
                     newkey = "API TEXT VIEWER BY VENKATESH" & Chr(174) & "\" & "MRUAPIs"
                     retval = RegOpenKeyEx(hkey, subkey & newkey, 0, KEY_ALL_ACCESS, phkResult)
                     If retval = ERROR_SUCCESS Then
                        cmcount = cmcount + 1
                        nvn = "String" & CStr(cmcount)
                        value = FINDfrm.finwha.Text
                        retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
                        RegCloseKey phkResult
                        APIfrm.StatusBar1.Panels(1).Text = "Last Query:" & Space(2) & APIfrm.ListView4.ListItems(1).Text
                        APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
                        APIfrm.StatusBar1.Panels(2).Text = "Types:"
                        APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
                        APIfrm.StatusBar1.Panels(3).Text = ""
                     End If
                     Call fincan_Click
                     APIfrm.SSTab1.Tab = 3
                     Exit Sub
                  Else
                     tyfnd = False
                  End If
                  apityp.MoveNext
               Wend
               If tyfnd = False Then
                  MsgBox "No matches found for" & Space(2) & Trim(finwha.Text), vbCritical + vbSystemModal, App.Title
                  Unload Me
               End If
         End Select
      ElseIf opt2cli = True Then
         If decfnd = True Then
            Set apirs = apidb.OpenRecordset("Declares", dbOpenDynaset)
            apirs.MoveFirst
            While Not apirs.EOF
               If apirs("Name") = finwha.Text Then
                  declcat = True
                  decfnd = True
                  APIfrm.ListView4.ColumnHeaders(1).Text = "Declares:"
                  APIfrm.ListView4.ColumnHeaders(2).Text = "Code"
                  If APIfrm.ListView4.ListItems.count > 0 Then
                     APIfrm.ListView4.ListItems.Clear
                  End If
                  Set itmX = APIfrm.ListView4.ListItems.Add(, , CStr(apirs("Name")))
                  itmX.Icon = 1   ' Set an icon from ImageList1.
                  itmX.SmallIcon = 1  ' Set an icon from ImageList2.
                  If Not IsNull(apirs("FullName")) Then
                     itmX.SubItems(1) = CStr(apirs("FullName"))
                  End If
                  hkey = HKEY_LOCAL_MACHINE
                  subkey = "SOFTWARE\"
                  newkey = "API TEXT VIEWER BY VENKATESH" & Chr(174) & "\" & "MRUAPIs"
                  retval = RegOpenKeyEx(hkey, subkey & newkey, 0, KEY_ALL_ACCESS, phkResult)
                  If retval = ERROR_SUCCESS Then
                     cmcount = cmcount + 1
                     nvn = "String" & CStr(cmcount)
                     value = FINDfrm.finwha.Text
                     retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
                     RegCloseKey phkResult
                     APIfrm.StatusBar1.Panels(1).Text = "Last Query:" & Space(2) & APIfrm.ListView4.ListItems(1).Text
                     APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
                     APIfrm.StatusBar1.Panels(2).Text = "Declares:"
                     APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
                     APIfrm.StatusBar1.Panels(3).Text = ""
                  End If
                  Call fincan_Click
                  APIfrm.SSTab1.Tab = 3
                  Exit Sub
               Else
                  decfnd = False
               End If
               apirs.MoveNext
            Wend
         End If
         If decfnd = False Then
            Set apicnrs = apidb.OpenRecordset("Constants", dbOpenDynaset)
            apicnrs.MoveFirst
            While Not apicnrs.EOF
               If Trim(apicnrs("Name")) = Trim(finwha.Text) Then
                  conscat = True
                  APIfrm.ListView4.ColumnHeaders(1).Text = "Constants:"
                  APIfrm.ListView4.ColumnHeaders(2).Text = "Code:"
                  If APIfrm.ListView4.ListItems.count > 0 Then
                     APIfrm.ListView4.ListItems.Clear
                  End If
                  Set itmX = APIfrm.ListView4.ListItems.Add(, , CStr(apicnrs("Name")))
                  itmX.Icon = 2   ' Set an icon from ImageList1.
                  itmX.SmallIcon = 2  ' Set an icon from ImageList2.
                  If Not IsNull(apicnrs("FullName")) Then
                     itmX.SubItems(1) = CStr(apicnrs("FullName"))
                  End If
                  hkey = HKEY_LOCAL_MACHINE
                  subkey = "SOFTWARE\"
                  newkey = "API TEXT VIEWER BY VENKATESH" & Chr(174) & "\" & "MRUAPIs"
                  retval = RegOpenKeyEx(hkey, subkey & newkey, 0, KEY_ALL_ACCESS, phkResult)
                  If retval = ERROR_SUCCESS Then
                     cmcount = cmcount + 1
                     nvn = "String" & CStr(cmcount)
                     value = FINDfrm.finwha.Text
                     retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
                     RegCloseKey phkResult
                     APIfrm.StatusBar1.Panels(1).Text = "Last Query:" & Space(2) & APIfrm.ListView4.ListItems(1).Text
                     APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
                     APIfrm.StatusBar1.Panels(2).Text = "Constants:"
                     APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
                     APIfrm.StatusBar1.Panels(3).Text = ""
                  End If
                  cnstfnd = True
                  Call fincan_Click
                  APIfrm.SSTab1.Tab = 3
                  Exit Sub
               Else
                  cnstfnd = False
               End If
               apicnrs.MoveNext
            Wend
         End If
         If cnstfnd = False Then
            typcat = True
            APIfrm.ListView4.ColumnHeaders(1).Text = "Types:"
            'APIfrm.ListView4.ColumnHeaders(2).Width = 0
            If APIfrm.ListView4.ListItems.count > 0 Then
               APIfrm.ListView4.ListItems.Clear
            End If
            Set apityp = apidb.OpenRecordset("Types", dbOpenDynaset)
            apityp.MoveFirst
            While Not apityp.EOF
               If Trim(apityp("Name")) = Trim(finwha.Text) Then
                  Set itmX = APIfrm.ListView4.ListItems.Add(, , CStr(apityp("Name")))
                  itmX.Icon = 2   ' Set an icon from ImageList1.
                  itmX.SmallIcon = 2  ' Set an icon from ImageList2.
                  hkey = HKEY_LOCAL_MACHINE
                  subkey = "SOFTWARE\"
                  newkey = "API TEXT VIEWER BY VENKATESH" & Chr(174) & "\" & "MRUAPIs"
                  retval = RegOpenKeyEx(hkey, subkey & newkey, 0, KEY_ALL_ACCESS, phkResult)
                  If retval = ERROR_SUCCESS Then
                     cmcount = cmcount + 1
                     nvn = "String" & CStr(cmcount)
                     value = FINDfrm.finwha.Text
                     retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
                     RegCloseKey phkResult
                     APIfrm.StatusBar1.Panels(1).Text = "Last Query:" & Space(2) & APIfrm.ListView4.ListItems(1).Text
                     APIfrm.StatusBar1.Panels(1).AutoSize = sbrContents
                     APIfrm.StatusBar1.Panels(2).Text = "Types:"
                     APIfrm.StatusBar1.Panels(2).AutoSize = sbrContents
                     APIfrm.StatusBar1.Panels(3).Text = ""
                  End If
                  Call fincan_Click
                  APIfrm.SSTab1.Tab = 3
                  Exit Sub
               Else
                  typfnd = False
               End If
               apityp.MoveNext
            Wend
         End If
         If typfnd = False Then
            MsgBox "No matches found for" & Space(2) & Trim(finwha.Text), vbCritical + vbSystemModal, App.Title
            Unload Me
         End If
      End If    'for checking which category clicked
   End If    'for checking the length
End Sub
Private Sub fincan_Click()
   Unload Me
End Sub
Private Sub finwha_KeyPress(KeyAscii As Integer)
   'Select Case keyacsii
   'Case 48 To 57
   'Case 65 To 92
   'Case Else
   'KeyAscii = 0
   'End Select
End Sub
Private Sub Form_Load()
   Dim ll, lt As Long, retval1 As Boolean
   lt = (Screen.Height \ 2 - Me.Height \ 2) \ Screen.TwipsPerPixelY
   ll = (Screen.Width \ 2 - Me.Width \ 2) \ Screen.TwipsPerPixelX
   retval1 = SetWindowPos(Me.hwnd, HWND_TOPMOST, ll, lt, 0&, 0&, SWP_NOSIZE)
   opt1.value = True
   opt1cli = True
   Me.Show
   modREG.GetMRU
   If ffldoce = False Then
      APIfrm.ListView4.ColumnHeaders.Add 1, , , APIfrm.ListView4.Width / 4, lvwColumnLeft
      APIfrm.ListView4.ColumnHeaders.Add 2, , , APIfrm.ListView4.Width / 1, lvwColumnLeft
      Set imgX = APIfrm.ImageList4.ListImages.Add(1, , LoadPicture(App.path & "\declare.bmp"))
      Set imgX = APIfrm.ImageList4.ListImages.Add(2, , LoadPicture(App.path & "\const.bmp"))
      APIfrm.ListView4.Icons = APIfrm.ImageList4
      APIfrm.ListView4.SmallIcons = APIfrm.ImageList4
      ffldoce = True
   End If
   finwha.SetFocus
End Sub
Private Sub opt1_Click()
   opt1cli = True
   opt2cli = False
End Sub
Private Sub opt2_Click()
   opt2cli = True
   opt1cli = False
End Sub
