Attribute VB_Name = "modAPI"
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 55
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 54
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const GWL_STYLE = (-16)
Public Const HDS_BUTTONS = &H2
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_FILE_EXISTS = 80&
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public path As String, ret As Long, finpath As String
'=====================~ For finding the file ~================================================================================================================================================================
Public Const MAX_PATH = 260
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'=====================~ For finding the file ~================================================================================================================================================================
Public apidb As Database
Public typ As String, apicons As String, apifunc As String, pubcli As Boolean, pricli As Boolean, defcli As Boolean
Public apirs As Recordset
Public apicnrs As Recordset
Public apityp As Recordset
Public ffldoce As Boolean
Public Sub showprogressinstatusbar(ByVal bshowprogressbar As Boolean, pro As Control, frmnam As Form)
   Dim trc As RECT
   If bshowprogressbar Then
      SendMessage APIfrm.StatusBar1.hwnd, SB_GETRECT, 1, trc
      With trc
         .Top = (.Top * Screen.TwipsPerPixelY)
         .Left = (.Left * Screen.TwipsPerPixelX)
         .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
         .Right = (.Right * Screen.TwipsPerPixelX) - .Left
      End With
      With pro
         SetParent .hwnd, APIfrm.StatusBar1.hwnd
         .Move trc.Left, trc.Top, trc.Right, trc.Bottom
         .Visible = True
         .value = 0
      End With
   Else
      SetParent pro.hwnd, frmnam.hwnd
      pro.Visible = False
   End If
End Sub
Public Sub Main()
   Dim hkey As Long, subkey As String, newkey As String
   On Error GoTo errhan
   Set apidb = OpenDatabase(App.path & "\Win32api.MDB")
   'Set apidb = OpenDatabase("C:\Program Files\Microsoft Visual Studio\Common\Tools\Winapi\Win32api.MDB")
   '----------------------~ For Writing Appln name to registry ~-----------------------------------------
   '----------------------~ For Writing Appln name to registry ~-----------------------------------------
   Load APIfrm
   APIfrm.Show
   defcli = True
   Exit Sub
errhan:
   Select Case Err.Number
      Case 3044:
         MsgBox Err.Description & vbCrLf & "Please make the file from VB 6.0 Studio Tools >> API TEXT VIEWER..." & vbCrLf & "This will now self destruct", vbCritical + vbSystemModal, App.Title
         End
   End Select
End Sub
Public Function SwitchTo(ByVal Handle&, ByVal sWindowText$)
   ' switches to any running program; use IsProgramRunning first
   ' Handle is a handle to your own program (YourForm.hwnd)
   Dim NewHandle&
   If IsProgramRunning(Handle, sWindowText) Then
      NewHandle = GetHandle(Handle, sWindowText)
      If NewHandle = 0 Then Exit Function
      SetForegroundWindow NewHandle
   End If
End Function
Public Function IsProgramRunning(mHandle&, FileName$) As Boolean
   Dim bFound: bFound = False ' let's be pessimistic and say it's not
   ' gets the current tasklist, which is very useful given SendKeys
   Dim CurrWnd As Long
   Dim Length As Long
   Dim TaskName As String
   Dim Parent As Long
   Dim lPos As Long
   Dim TestWin&

   CurrWnd = GetWindow(mHandle, GW_HWNDFIRST)

   While CurrWnd <> 0

      If CurrWnd <> mHandle And (0 <> IsWindowVisible(CurrWnd)) And _
         (0 = GetWindow(CurrWnd, GW_OWNER)) Then

         Length = GetWindowTextLength(CurrWnd)
         TaskName = Space$(Length + 1)
         Length = GetWindowText(CurrWnd, TaskName, Length + 1)
         TaskName = Left$(TaskName, Len(TaskName) - 1)

         If Length > 0 Then
            ' _________________________________________
            Dim pos&
            pos = InStr(1, TaskName, FileName, vbTextCompare)
            If pos > 0 Then

               ' HERE 'S WHERE YOU FIND ALL THE RUNNING, VISIBLE PROGRAMS!
               ' YOU CAN DO ANYTHING YOU WANT HERE
               ' in this case we're just gonna find out if an instance of a program is running

               bFound = True            ' uh yep, we found it
               IsProgramRunning = True
               Exit Function
            End If
         End If
         ' _______________________________________
      End If
      CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
      DoEvents
   Wend
   IsProgramRunning = bFound
End Function
Public Function GetHandle(mHandle&, WindowText$) As Long
   ' gets the handle of a VISIBLE program (program on task bar)
   ' works with partial window names
   Dim CurrWnd As Long
   Dim Length As Long
   Dim TaskName As String
   Dim Parent As Long
   Dim lPos As Long
   Dim TestWin&
   CurrWnd = GetWindow(mHandle, GW_HWNDFIRST)
   While CurrWnd <> 0
      If CurrWnd <> mHandle And (0 <> IsWindowVisible(CurrWnd)) And _
         (0 = GetWindow(CurrWnd, GW_OWNER)) Then

         Length = GetWindowTextLength(CurrWnd)
         TaskName = Space$(Length + 1)
         Length = GetWindowText(CurrWnd, TaskName, Length + 1)
         TaskName = Left$(TaskName, Len(TaskName) - 1)

         If Length > 0 Then
            ' _________________________________________
            Dim pos&
            pos = InStr(1, TaskName, WindowText, vbTextCompare)
            If pos > 0 Then

               ' HERE 'S WHERE YOU FIND ALL THE RUNNING, VISIBLE PROGRAMS!
               ' YOU CAN DO ANYTHING YOU WANT HERE

               ' uh yep, we found it , so let's leave!!!!
               GetHandle = CurrWnd
               Exit Function

            End If
         End If
         ' _______________________________________
      End If
      CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
      DoEvents
   Wend
End Function

