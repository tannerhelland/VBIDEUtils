Attribute VB_Name = "Shell_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:53
' * Module Name      : Shell_Module
' * Module Filename  : Execute.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

' File and Disk functions.
Private Const DRIVE_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_UNKNOWN = 0    'Unknown, or unable to be determined.

Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long

Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const UNIQUE_NAME = &H0

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_HIDE = 0             ' = vbHide
Private Const SW_SHOWNORMAL = 1       ' = vbNormal
Private Const SW_SHOWMINIMIZED = 2    ' = vbMinimizeFocus
Private Const SW_SHOWMAXIMIZED = 3    ' = vbMaximizedFocus
Private Const SW_SHOWNOACTIVATE = 4   ' = vbNormalNoFocus
Private Const SW_MINIMIZE = 6         ' = vbMinimizedNofocus

Private Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Type SHFILEOPSTRUCT
   hWnd                 As Long
   wFunc                As Long
   pFrom                As String
   pTo                  As String
   fFlags               As Integer
   fAborted             As Boolean
   hNameMaps            As Long
   sProgress            As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_SILENT = &H4
Private Const FOF_NOCONFIRMATION = &H10

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type STARTUPINFO
   cb                   As Long
   lpReserved           As String
   lpDesktop            As String
   lpTitle              As String
   dwX                  As Long
   dwY                  As Long
   dwXSize              As Long
   dwYSize              As Long
   dwXCountChars        As Long
   dwYCountChars        As Long
   dwFillAttribute      As Long
   dwFlags              As Long
   wShowWindow          As Integer
   cbReserved2          As Integer
   lpReserved2          As Long
   hStdInput            As Long
   hStdOutput           As Long
   hStdError            As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess             As Long
   hThread              As Long
   dwProcessID          As Long
   dwThreadID           As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function FindExecutableA Lib "shell32.dll" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function SetVolumeLabelA Lib "kernel32" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long

' *** Get running applications
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Const GW_CHILD = 5
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_MAX = 5
Const GW_OWNER = 4

' *** Infos about files
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const OFS_MAXPATHNAME = 128
Public Const OF_READ = &H0

Type OFSTRUCT
   cBytes               As Byte
   fFixedDisk           As Byte
   nErrCode             As Integer
   Reserved1            As Integer
   Reserved2            As Integer
   szPathName(OFS_MAXPATHNAME) As Byte
End Type

Type FILETIME
   dwLowDateTime        As Long
   dwHighDateTime       As Long
End Type

Type BY_HANDLE_FILE_INFORMATION
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   dwVolumeSerialNumber As Long
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   nNumberOfLinks       As Long
   nFileIndexHigh       As Long
   nFileIndexLow        As Long
End Type

' *** RegisterAsServiceProcess
Public Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Public Function FindExecutable(s As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : FindExecutable
   ' * Parameters       :
   ' *                    s As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '  Finds the executable associated with a file
   '
   '  Returns "" if no file is found.

   Dim i                As Integer
   Dim s2               As String

   s2 = String$(256, 32) & Chr$(0)

   i = FindExecutableA(s & Chr$(0), "", s2)

   If i > 32 Then
      FindExecutable = left$(s2, InStr(s2, Chr$(0)) - 1)
   Else
      FindExecutable = ""
   End If

End Function

Public Function ShellDelete(ParamArray vntFileName() As Variant) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : ShellDelete
   ' * Parameters       :
   ' *                    ParamArray vntFileName() As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   '  Deletes a single file, or an array of files to the trashcan.
   '
   Dim i                As Integer
   Dim sFileNames       As String
   Dim SHFileOp         As SHFILEOPSTRUCT

   For i = LBound(vntFileName) To UBound(vntFileName)
      sFileNames = sFileNames & vntFileName(i) & vbNullChar
   Next

   sFileNames = sFileNames & vbNullChar

   With SHFileOp
      .wFunc = FO_DELETE
      .pFrom = sFileNames
      .fFlags = FOF_ALLOWUNDO + FOF_SILENT + FOF_NOCONFIRMATION
   End With

   i = SHFileOperation(SHFileOp)

   If i = 0 Then
      ShellDelete = True
   Else
      ShellDelete = False
   End If

End Function

Public Function ShellWait(cCommandLine As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 2/10/98
   ' * Time             : 09:48
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : ShellWait
   ' * Parameters       :
   ' *                    cCommandLine As String
   ' **********************************************************************
   ' * Comments         :
   ' * Runs a command as the Shell command does but waits for the command
   ' * to finish before returning.  Note: The full path and filename extention
   ' * is required.
   ' * You might want to use Environ$("COMSPEC") & " /c " & command
   ' * if you wish to run it under the command shell (and thus it)
   ' * will search the path etc...
   ' *
   ' * returns false if the shell failed
   ' *
   ' **********************************************************************

   Dim NameOfProc       As PROCESS_INFORMATION
   Dim NameStart        As STARTUPINFO
   Dim i                As Long

   NameStart.cb = Len(NameStart)
   i = CreateProcessA(0&, cCommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)

   If i <> 0 Then
      Call WaitForSingleObject(NameOfProc.hProcess, INFINITE)
      Call CloseHandle(NameOfProc.hProcess)
      ShellWait = True
   Else
      ShellWait = False
   End If

End Function

Public Sub MonitorProcess(sProcess As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : MonitorProcess
   ' * Parameters       :
   ' *                    sProcess As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** check if another application is still running

   Dim pId              As Long
   Dim pHnd             As Long

   pId = Shell(sProcess, vbNormalFocus)

   pHnd = OpenProcess(SYNCHRONIZE, 0, pId) ' get Process Handle
   If pHnd <> 0 Then
      Call WaitForSingleObject(pHnd, INFINITE) ' Wait until shelled prog ends
      Call CloseHandle(pHnd)
   End If

End Sub

Public Function ExecuteWait(s As String, Optional Param As Variant) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : ExecuteWait
   ' * Parameters       :
   ' *                    s As String
   ' *                    Optional Param As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   '                    As the Execute function but waits for the process to finish before
   '  returning
   '
   '  returns true on success.

   Dim s2               As String

   s2 = FindExecutable(s)

   If s2 <> "" Then
      ExecuteWait = ShellWait(s2 & IIf(IsMissing(Param), " ", " " & CStr(Param) & " ") & s)
   Else
      ExecuteWait = False
   End If

End Function

Public Function AddBackslash(s As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : AddBackslash
   ' * Parameters       :
   ' *                    s As String
   ' **********************************************************************
   ' * Comments         :
   ' *  Add a backslash if the string doesn't have one already.
   ' *
   ' **********************************************************************

   If Len(s) > 0 Then
      If right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If

End Function

Public Function ExecuteWithAssociate(ByVal hWnd As Long, sExecute As String, Optional Param As Variant = "", Optional windowstyle As Variant) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 9/10/98
   ' * Time             : 11:48
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : ExecuteWithAssociate
   ' * Parameters       :
   ' *                    ByVal hwnd As Long
   ' *                    sExecute As String
   ' *                    Optional Param As Variant = ""
   ' *                    Optional windowstyle As Variant
   ' **********************************************************************
   ' * Comments         :
   ' * Executes a file with it's associated program.
   ' *   windowstyle uses the same constants as the Shell function:
   ' *      vbHide   0
   ' *      vbNormalFocus  1
   ' *      vbMinimizedFocus  2
   ' *      vbMaximizedFocus  3
   ' *      vbNormalNoFocus   4
   ' *      vbMinimizedNoFocus   6
   ' *
   ' *  returns true on success
   ' *
   ' *
   ' **********************************************************************

   Dim i                As Long

   If IsMissing(windowstyle) Then windowstyle = vbNormalFocus

   i = ShellExecute(0, vbNullString, sExecute, vbNullString, vbNullString, CLng(windowstyle))
   If i > 32 Then
      ExecuteWithAssociate = True
   Else
      ExecuteWithAssociate = False
   End If

End Function

Public Function GetFile(s As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetFile
   ' * Parameters       :
   ' *                    s As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   '  Returns the file portion of a file + pathname
   '

   Dim i                As Integer
   Dim j                As Integer

   i = 0
   j = 0

   i = InStr(s, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, s, "\")
   Loop

   If j = 0 Then
      GetFile = ""
   Else
      GetFile = right$(s, Len(s) - j)
   End If

End Function

Public Function GetPath(s As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetPath
   ' * Parameters       :
   ' *                    s As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   '  Returns the path portion of a file + pathname
   '

   Dim i                As Integer
   Dim j                As Integer

   i = 0
   j = 0

   i = InStr(s, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, s, "\")
   Loop

   If j = 0 Then
      GetPath = ""
   Else
      GetPath = left$(s, j)
   End If
End Function

Public Function GetShortPathName(longpath As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetShortPathName
   ' * Parameters       :
   ' *                    longpath As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim s                As String
   Dim i                As Long

   i = Len(longpath) + 1
   s = String$(i, 0)
   GetShortPathNameA longpath, s, i

   GetShortPathName = left$(s, InStr(s, Chr$(0)) - 1)
End Function

Public Function GetSystemDirectory() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetSystemDirectory
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *  Return the system directory of windows
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim i                As Integer

   i = GetSystemDirectoryA("", 0)
   sTmp = Space$(i)
   Call GetSystemDirectoryA(sTmp, i)
   sTmp = left$(sTmp, i - 1)
   If right$(sTmp, 1) <> "\" Then
      GetSystemDirectory = sTmp + "\"
   Else
      GetSystemDirectory = sTmp
   End If

End Function

Public Function GetTempFileName() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/11/98
   ' * Time             : 11:16
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetTempFileName
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Returns a unique tempfile name.
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim sTmp2            As String

   sTmp2 = GetTempPath
   sTmp = Space$(Len(sTmp2) + 256)
   Call GetTempFileNameA(sTmp2, App.EXEName, UNIQUE_NAME, sTmp)
   GetTempFileName = left$(sTmp, InStr(sTmp, Chr$(0)) - 1)

End Function

Public Function GetTempPath() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 17/11/98
   ' * Time             : 11:16
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetTempPath
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Returns the path to the temp directory.
   ' *
   ' *
   ' **********************************************************************

   Dim sTmp             As String
   Dim i                As Integer

   i = GetTempPathA(0, "")
   sTmp = Space$(i)

   Call GetTempPathA(i, sTmp)
   GetTempPath = AddBackslash(left$(sTmp, i - 1))

End Function

Public Function GetWindowsDirectory() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetWindowsDirectory
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   '  Returns the windows directory.
   '
   Dim s                As String
   Dim i                As Integer
   i = GetWindowsDirectoryA("", 0)
   s = Space$(i)
   Call GetWindowsDirectoryA(s, i)
   GetWindowsDirectory = AddBackslash(left$(s, i - 1))

End Function

Public Function RemoveBackslash(s As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : RemoveBackslash
   ' * Parameters       :
   ' *                    s As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   '  Removes the backslash from the string if it has one.
   '
   Dim i                As Integer
   i = Len(s)
   If i <> 0 Then
      If right$(s, 1) = "\" Then
         RemoveBackslash = left$(s, i - 1)
      Else
         RemoveBackslash = s
      End If
   Else
      RemoveBackslash = ""
   End If
End Function

Public Function sDriveType(sDrive As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : sDriveType
   ' * Parameters       :
   ' *                    sDrive As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   ' Returns the drive type if possible.
   '
   Dim lRet             As Long

   lRet = GetDriveTypeA(sDrive & ":\")
   Select Case lRet
      Case 0
         'sDriveType = "Cannot be determined!"
         sDriveType = "Unknown"

      Case 1
         'sDriveType = "The root directory does not exist!"
         sDriveType = "Unknown"
      Case DRIVE_CDROM:
         sDriveType = "CD-ROM Drive"

      Case DRIVE_REMOVABLE:
         sDriveType = "Removable Drive"

      Case DRIVE_FIXED:
         sDriveType = "Fixed Drive"

      Case DRIVE_REMOTE:
         sDriveType = "Remote Drive"
   End Select

End Function

Public Function GetDriveType(sDrive As String) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetDriveType
   ' * Parameters       :
   ' *                    sDrive As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lRet             As Long
   lRet = GetDriveTypeA(sDrive & ":\")

   If lRet = 1 Then
      lRet = 0
   End If

   GetDriveType = lRet

End Function

Public Sub GetRunningApplications(nHwnd As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : GetRunningApplications
   ' * Parameters       :
   ' *                    nHwnd As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Show all the running applications

   Dim lLgthChild       As Long
   Dim sNameChild       As String
   Dim lLgthOwner       As Long
   Dim sNameOwner       As String
   Dim lHwnd            As Long
   Dim lHwnd2           As Long
   Dim lProssId         As Long

   Const vbTextCompare = 1

   lHwnd = GetWindow(nHwnd, GW_HWNDFIRST)
   While lHwnd <> 0
      lHwnd2 = GetWindow(lHwnd, GW_OWNER)
      lLgthOwner = GetWindowTextLength(lHwnd2)
      sNameOwner = String$(lLgthOwner + 1, Chr$(0))
      lLgthOwner = GetWindowText(lHwnd2, sNameOwner, lLgthOwner + 1)

      If lLgthOwner <> 0 Then
         sNameOwner = left$(sNameOwner, InStr(1, sNameOwner, Chr$(0), vbTextCompare) - 1)
         Call GetWindowThreadProcessId(lHwnd2, lProssId)
         'Debug.Print sNameOwner, lProssId
      End If

      lLgthChild = GetWindowTextLength(lHwnd)
      sNameChild = String$(lLgthChild + 1, Chr$(0))
      lLgthChild = GetWindowText(lHwnd, sNameChild, lLgthChild + 1)
      If lLgthChild <> 0 Then
         sNameChild = left$(sNameChild, InStr(1, sNameChild, Chr$(0), vbTextCompare) - 1)
         Call GetWindowThreadProcessId(lHwnd, lProssId)
         'Debug.Print sNameChild, lProssId
      End If

      lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
      DoEvents

   Wend

End Sub

Public Function RegisterAsServiceProcess(bRegister As Boolean) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 25/09/98
   ' * Time             : 11:36
   ' * Module Name      : Shell_Module
   ' * Module Filename  : Execute.bas
   ' * Procedure Name   : RegisterAsServiceProcess
   ' * Parameters       :
   ' *                    bRegister As Boolean
   ' **********************************************************************
   ' * Comments         : This code allows the programer to register
   ' * the application as a service process.
   ' * To put it in simpler terms, a service process does not appear in the "End Task"
   ' * dialog box. Using this call does not let the user of your application terminate
   ' * the program using Ctrl-Alt-Del.
   ' *
   ' **********************************************************************

   RegisterAsServiceProcess = RegisterServiceProcess(GetCurrentProcessId, IIf(bRegister, 1, 0)) = 1

End Function
