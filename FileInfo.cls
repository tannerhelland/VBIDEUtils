VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_FileInfo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 26/11/1999
' * Time             : 16:33
' * Module Name      : class_FileInfo
' * Module Filename  : FileInfo.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit
'
' API declarations
'
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, lpFilePart As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal nBufferLength As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'
' API constants.
'
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
'
' File attribute constants.
'
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
'
' SHGetFileInfo constants.
'
Private Const SHGFI_ICON = &H100                         '  get icon
Private Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Private Const SHGFI_TYPENAME = &H400                     '  get type name
Private Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Private Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Private Const SHGFI_EXETYPE = &H2000                     '  return exe type
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Private Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Private Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const SHGFI_OPENICON = &H2                       '  get open icon
Private Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Private Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Private Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
'
' CreateFile constants
'
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
'
' API structures.
'
Private Type FILETIME
   dwLowDateTime        As Long
   dwHighDateTime       As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * MAX_PATH
   cAlternate           As String * 14
End Type

Private Type SYSTEMTIME
   wYear                As Integer
   wMonth               As Integer
   wDayOfWeek           As Integer
   wDay                 As Integer
   wHour                As Integer
   wMinute              As Integer
   wSecond              As Integer
   wMilliseconds        As Integer
End Type

Private Type SHFILEINFO
   hIcon                As Long                       '  out: icon
   iIcon                As Long                       '  out: icon index
   dwAttributes         As Long                '  out: SFGAO_ flags
   szDisplayName        As String * MAX_PATH  '  out: display name (or path)
   szTypeName           As String * 80           '  out: type name
End Type
'
' Member variables.
'
Private m_PathName      As String
Private m_Name          As String
Private m_Path          As String
Private m_Extension     As String
Private m_DisplayName   As String
Private m_TypeName      As String
Private m_hIcon         As Long
Private m_PathNameShort As String
Private m_NameShort     As String
Private m_PathShort     As String
Private m_FileExists    As Boolean
Private m_PathExists    As Boolean
Private m_FileSize      As Long
Private m_FileSizeHigh  As Long
Private m_CompFileSize  As Long
Private m_CompFileSizeHigh As Long
Private m_Attributes    As Long
Private m_tmCreation    As Double
Private m_tmAccess      As Double
Private m_tmWrite       As Double
'
' Enumerated constants
'
Private Enum FileTimes
   ftCreationTime = 0
   ftLastAccessTime = 1
   ftLastWriteTime = 2
End Enum

' ********************************************
'  Initialize and Terminate
' ********************************************
Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   ' All member variables can be left to defaults.
   '
End Sub

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '
   ' Just need to clear the icon copy to be
   ' completely tidy.
   '
   If m_hIcon Then
      Call DestroyIcon(m_hIcon)
   End If
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Let FullPathName(ByVal NewVal As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : FullPathName
   ' * Parameters       :
   ' *                    ByVal NewVal As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Buffer           As String
   Dim nFilePart        As Long
   Dim nRet             As Long
   '
   ' Retrieve fully qualified path/name specs.
   '
   Buffer = Space(MAX_PATH)
   nRet = GetFullPathName(NewVal, Len(Buffer), Buffer, nFilePart)
   If nRet Then
      m_PathName = left$(Buffer, nRet)
      Refresh
   End If
End Property

Public Property Get FullPathName() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : FullPathName
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns fully-qualified path/name spec.
   FullPathName = m_PathName
End Property

Public Property Get FileName() As String
   ' Returns filename only.
   FileName = m_Name
End Property

Public Property Get FilePath() As String
   ' Returns fully-qualified pathname only.
   FilePath = m_Path
End Property

Public Property Get FileExtension() As String
   ' Returns the file's extension only.
   FileExtension = m_Extension
End Property

Public Property Get ShortPathName() As String
   ' Returns fully-qualified *short* path/name spec.
   ShortPathName = m_PathNameShort
End Property

Public Property Get ShortName() As String
   ' Returns *short* filename only.
   ShortName = m_NameShort
End Property

Public Property Get ShortPath() As String
   ' Returns *short* fully-qualified pathname only.
   ShortPath = m_PathShort
End Property

Public Property Get DisplayName() As String
   ' Returns the "display" name for the file, not necessarily
   ' proper-cased, but as Explorer shows it.
   DisplayName = m_DisplayName
End Property

Public Property Get TypeName() As String
   ' Returns the string that describes the file's type.
   TypeName = m_TypeName
End Property

Public Property Get FileExists() As Boolean
   ' Returns whether file exists.
   FileExists = m_FileExists
End Property

Public Property Get PathExists() As Boolean
   ' Returns whether path exists.
   PathExists = m_PathExists
End Property

Public Property Get FileSize() As Long
   ' Return size of file.
   FileSize = m_FileSize
End Property

Public Property Get FileSizeHigh() As Long
   ' Returns high dword of filesize to support files > 2Gb.
   FileSizeHigh = m_FileSizeHigh
End Property

Public Property Get CompressedFileSize() As Long
   ' Return actual size of file.
   CompressedFileSize = m_CompFileSize
End Property

Public Property Get CompressedFileSizeHigh() As Long
   ' Returns high dword of actual filesize to support files > 2Gb.
   CompressedFileSizeHigh = m_CompFileSizeHigh
End Property

Public Property Let CreationTime(ByVal NewVal As Double)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : CreationTime
   ' * Parameters       :
   ' *                    ByVal NewVal As Double
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Try setting new timestamp.
   If SetTime(NewVal, ftCreationTime) Then
      Me.Refresh
   End If
End Property

Public Property Get CreationTime() As Double
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : CreationTime
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns date/time of file creation.
   CreationTime = m_tmCreation
End Property

Public Property Let LastAccessTime(ByVal NewVal As Double)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : LastAccessTime
   ' * Parameters       :
   ' *                    ByVal NewVal As Double
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Try setting new timestamp.
   If SetTime(NewVal, ftLastAccessTime) Then
      Me.Refresh
   End If
End Property

Public Property Get LastAccessTime() As Double
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : LastAccessTime
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns date/time of last access.
   LastAccessTime = m_tmAccess
End Property

Public Property Let ModifyTime(ByVal NewVal As Double)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : ModifyTime
   ' * Parameters       :
   ' *                    ByVal NewVal As Double
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Try setting new timestamp.
   If SetTime(NewVal, ftLastWriteTime) Then
      Me.Refresh
   End If
End Property

Public Property Get ModifyTime() As Double
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : ModifyTime
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns date/time of last write.
   ModifyTime = m_tmWrite
End Property

Public Property Let Attributes(ByVal NewVal As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : Attributes
   ' * Parameters       :
   ' *                    ByVal NewVal As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Attempt to set new attributes if not set already.
   If NewVal <> m_Attributes Then
      On Error Resume Next
      If SetAttr(NewVal) Then
         On Error GoTo 0
         Me.Refresh
      End If
   End If
End Property

Public Property Get Attributes() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : Attributes
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns entire set of attribute flags.
   Attributes = m_Attributes
End Property

Public Property Let attrReadOnly(ByVal NewVal As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrReadOnly
   ' * Parameters       :
   ' *                    ByVal NewVal As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim NewAttr          As Long
   ' Calculate new attribute value.
   If NewVal Then
      NewAttr = m_Attributes Or FILE_ATTRIBUTE_READONLY
   Else
      NewAttr = m_Attributes And Not FILE_ATTRIBUTE_READONLY
   End If
   ' Attempt to set new attribute.
   Me.Attributes = NewAttr
End Property

Public Property Get attrReadOnly() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrReadOnly
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns whether file has ReadOnly attribute.
   attrReadOnly = (m_Attributes And FILE_ATTRIBUTE_READONLY)
End Property

Public Property Let attrHidden(ByVal NewVal As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrHidden
   ' * Parameters       :
   ' *                    ByVal NewVal As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim NewAttr          As Long
   ' Calculate new attribute value.
   If NewVal Then
      NewAttr = m_Attributes Or FILE_ATTRIBUTE_HIDDEN
   Else
      NewAttr = m_Attributes And Not FILE_ATTRIBUTE_HIDDEN
   End If
   ' Attempt to set new attribute.
   Me.Attributes = NewAttr
End Property

Public Property Get attrHidden() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrHidden
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns whether file has Hidden attribute.
   attrHidden = (m_Attributes And FILE_ATTRIBUTE_HIDDEN)
End Property

Public Property Let attrSystem(ByVal NewVal As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrSystem
   ' * Parameters       :
   ' *                    ByVal NewVal As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim NewAttr          As Long
   ' Calculate new attribute value.
   If NewVal Then
      NewAttr = m_Attributes Or FILE_ATTRIBUTE_SYSTEM
   Else
      NewAttr = m_Attributes And Not FILE_ATTRIBUTE_SYSTEM
   End If
   ' Attempt to set new attribute.
   Me.Attributes = NewAttr
End Property

Public Property Get attrSystem() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrSystem
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns whether file has System attribute.
   attrSystem = (m_Attributes And FILE_ATTRIBUTE_SYSTEM)
End Property

Public Property Let attrArchive(ByVal NewVal As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrArchive
   ' * Parameters       :
   ' *                    ByVal NewVal As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim NewAttr          As Long
   ' Calculate new attribute value.
   If NewVal Then
      NewAttr = m_Attributes Or FILE_ATTRIBUTE_ARCHIVE
   Else
      NewAttr = m_Attributes And Not FILE_ATTRIBUTE_ARCHIVE
   End If
   ' Attempt to set new attribute.
   Me.Attributes = NewAttr
End Property

Public Property Get attrArchive() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrArchive
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns whether file has Archive attribute.
   attrArchive = (m_Attributes And FILE_ATTRIBUTE_ARCHIVE)
End Property

Public Property Let attrTemporary(ByVal NewVal As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrTemporary
   ' * Parameters       :
   ' *                    ByVal NewVal As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Cannot change Temporary attribute with normal methods
   ' (This must be set with CreateFile?), but include the
   ' stub routine here just in case anyone tries.
End Property

Public Property Get attrTemporary() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrTemporary
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns whether file has Temporary attribute.
   attrTemporary = (m_Attributes And FILE_ATTRIBUTE_TEMPORARY)
End Property

Public Property Let attrCompressed(ByVal NewVal As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrCompressed
   ' * Parameters       :
   ' *                    ByVal NewVal As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Cannot change Compressed attribute with normal methods,
   ' but should have a stub routine here just in case anyone
   ' tries.
End Property

Public Property Get attrCompressed() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : attrCompressed
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Returns whether file has Compressed attribute.
   attrCompressed = (m_Attributes And FILE_ATTRIBUTE_COMPRESSED)
End Property

Public Property Get hIcon() As Long
   ' Returns handle to display icon.
   hIcon = m_hIcon
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Sub Refresh()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : Refresh
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim hSearch          As Long
   Dim wfd              As WIN32_FIND_DATA
   Dim Buffer           As String
   Dim nRet             As Long
   Dim i                As Long
   Dim sfi              As SHFILEINFO
   '
   ' Check for existence of file.
   '
   hSearch = FindFirstFile(m_PathName, wfd)
   If hSearch <> INVALID_HANDLE_VALUE Then
      Call FindClose(hSearch)
      '
      ' Assign file data to member variables.
      '
      m_FileExists = True
      m_PathExists = True
      m_FileSize = wfd.nFileSizeLow
      m_FileSizeHigh = wfd.nFileSizeHigh
      m_Attributes = wfd.dwFileAttributes
      m_tmCreation = FileTimeToDouble(wfd.ftCreationTime, True)
      m_tmAccess = FileTimeToDouble(wfd.ftLastAccessTime, True)
      m_tmWrite = FileTimeToDouble(wfd.ftLastWriteTime, True)
      '
      ' Assign file/path data to member variables.
      '
      m_Name = TrimNull(wfd.cFileName)
      For i = Len(m_PathName) To 1 Step -1
         If Mid$(m_PathName, i, 1) = "\" Then
            m_Path = ProperCasePath(left(m_PathName, i))
            If right(m_Path, 1) <> "\" Then m_Path = m_Path & "\"
            Exit For
         End If
      Next i
      m_PathName = m_Path & m_Name
      '
      ' Extract extension from filename.
      '
      If InStr(m_Name, ".") Then
         For i = Len(m_Name) To 1 Step -1
            If Mid$(m_Name, i, 1) = "." Then
               m_Extension = Mid$(m_Name, i + 1)
               Exit For
            End If
         Next i
      Else
         m_Extension = ""
      End If
      '
      ' Short name same as long, if cAlternate element empty.
      '
      If InStr(wfd.cAlternate, vbNullChar) = 1 Then
         m_NameShort = UCase$(m_Name)
      Else
         m_NameShort = TrimNull(wfd.cAlternate)
      End If
      '
      ' Retrieve short path name.
      '
      Buffer = Space(MAX_PATH)
      nRet = GetShortPathName(m_PathName, Buffer, Len(Buffer))
      If nRet Then
         m_PathNameShort = left$(Buffer, nRet)
         m_PathShort = left$(m_PathNameShort, Len(m_PathNameShort) - Len(m_NameShort))
      End If
      '
      ' Retrieve compressed size.
      '
      m_CompFileSize = GetCompressedFileSize(m_PathName, m_CompFileSizeHigh)
      '
      ' Get icon and descriptive text.
      '
      If m_hIcon Then
         Call DestroyIcon(m_hIcon)
         m_hIcon = 0
      End If
      nRet = SHGetFileInfo(m_PathName, 0&, sfi, Len(sfi), _
         SHGFI_ICON Or SHGFI_DISPLAYNAME Or SHGFI_TYPENAME)
      m_DisplayName = TrimNull(sfi.szDisplayName)
      m_TypeName = TrimNull(sfi.szTypeName)
      m_hIcon = sfi.hIcon
      '
      ' Confirm displayable typename.
      '
      If Trim$(m_TypeName) = "" Then
         m_TypeName = Trim$(UCase$(m_Extension) & " File")
      End If
   Else
      '
      ' Assign applicable data to member variables.
      '
      m_FileExists = False
   End If
End Sub

Public Function FormatFileDate(ByVal dt As Double) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : FormatFileDate
   ' * Parameters       :
   ' *                    ByVal dt As Double
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   FormatFileDate = Format(dt, "long date") & " " & _
      Format(dt, "long time")
End Function

Public Function FormatFileSize(ByVal Size As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : FormatFileSize
   ' * Parameters       :
   ' *                    ByVal Size As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim sRet             As String
   Const KB& = 1024
   Const MB& = KB * KB
   ' Return size of file in kilobytes.
   If Size < KB Then
      sRet = Format(Size, "#,##0") & " bytes"
   Else
      Select Case Size \ KB
         Case Is < 10
            sRet = Format(Size / KB, "0.00") & "KB"
         Case Is < 100
            sRet = Format(Size / KB, "0.0") & "KB"
         Case Is < 1000
            sRet = Format(Size / KB, "0") & "KB"
         Case Is < 10000
            sRet = Format(Size / MB, "0.00") & "MB"
         Case Is < 100000
            sRet = Format(Size / MB, "0.0") & "MB"
         Case Is < 1000000
            sRet = Format(Size / MB, "0") & "MB"
         Case Is < 10000000
            sRet = Format(Size / MB / KB, "0.00") & "GB"
      End Select
      sRet = sRet & " (" & Format(Size, "#,##0") & " bytes)"
   End If
   FormatFileSize = sRet
End Function

' ********************************************
'  Private Methods
' ********************************************
Private Function DoubleToFileTime(ftDbl As Double, Optional Universalize As Boolean = True) As FILETIME
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : DoubleToFileTime
   ' * Parameters       :
   ' *                    ftDbl As Double
   ' *                    Optional Universalize As Boolean = True
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim ft               As FILETIME
   Dim st               As SYSTEMTIME
   '
   ' Convert double to systemtime structure.
   '
   With st
      .wYear = Year(ftDbl)
      .wMonth = Month(ftDbl)
      .wDay = Day(ftDbl)
      .wDayOfWeek = Weekday(ftDbl) - 1
      .wHour = Hour(ftDbl)
      .wMinute = Minute(ftDbl)
      .wSecond = Second(ftDbl)
   End With
   '
   ' Convert systemtime to filetime structure.
   '
   Call SystemTimeToFileTime(st, ft)
   '
   ' Convert local time to UTC time, if requested.
   '
   If Universalize Then
      Call LocalFileTimeToFileTime(ft, DoubleToFileTime)
   Else
      DoubleToFileTime = ft
   End If
End Function

Private Function FileTimeToDouble(ftUTC As FILETIME, Localize As Boolean) As Double
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : FileTimeToDouble
   ' * Parameters       :
   ' *                    ftUTC As FILETIME
   ' *                    Localize As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim ft               As FILETIME
   Dim st               As SYSTEMTIME
   Dim d                As Double
   Dim t                As Double
   '
   ' Convert to local filetime, if necessary.
   '
   If Localize Then
      Call FileTimeToLocalFileTime(ftUTC, ft)
   Else
      ft = ftUTC
   End If
   '
   ' Convert to system time structure.
   '
   Call FileTimeToSystemTime(ft, st)
   '
   ' Convert to VB-style date (double).
   '
   FileTimeToDouble = DateSerial(st.wYear, st.wMonth, st.wDay) + _
      TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function

Private Function SetAttr(NewAttr As Long) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : SetAttr
   ' * Parameters       :
   ' *                    NewAttr As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Nothing fancy, just set new attribute and return
   If m_FileExists Then
      SetAttr = SetFileAttributes(m_PathName, NewAttr)
   End If
End Function

Private Function SetTime(NewTime As Double, WhichTime As FileTimes) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : SetTime
   ' * Parameters       :
   ' *                    NewTime As Double
   ' *                    WhichTime As FileTimes
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim ft               As FILETIME
   Dim hFile            As Long
   '
   ' Bail if no file exists.
   '
   If m_FileExists = False Then Exit Function
   '
   ' Convert passed time (presumably local) to UTC.
   '
   ft = DoubleToFileTime(NewTime, True)
   '
   ' Get a handle on existing file so we can change times.
   '
   hFile = CreateFile(m_PathName, _
      GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
      ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
   '
   ' If we were able to open file, change it's timestamp.
   '
   If hFile <> INVALID_HANDLE_VALUE Then
      Select Case WhichTime
         Case ftCreationTime
            SetTime = SetFileTime(hFile, ft, ByVal 0&, ByVal 0&)
         Case ftLastAccessTime
            SetTime = SetFileTime(hFile, ByVal 0&, ft, ByVal 0&)
         Case ftLastWriteTime
            SetTime = SetFileTime(hFile, ByVal 0&, ByVal 0&, ft)
      End Select
      Call CloseHandle(hFile)
   End If
End Function

Private Function ProperCasePath(ByVal PathIn As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : ProperCasePath
   ' * Parameters       :
   ' *                    ByVal PathIn As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim hSearch          As Long
   Dim wfd              As WIN32_FIND_DATA
   Dim PathOut          As String
   Dim i                As Long
   '
   ' Trim trailing backslash, unless root dir.
   '
   If right(PathIn, 1) = "\" Then
      If right(PathIn, 2) <> ":\" Then
         PathIn = left$(PathIn, Len(PathIn) - 1)
      Else
         ProperCasePath = UCase$(PathIn)
         Exit Function
      End If
   End If
   '
   ' Check for UNC share and return just that,
   ' if that's all that's left of PathIn.
   '
   If InStr(PathIn, "\\") = 1 Then
      i = InStr(3, PathIn, "\")
      If i > 0 Then
         If InStr(i + 1, PathIn, "\") = 0 Then
            ProperCasePath = PathIn
            Exit Function
         End If
      End If
   End If
   '
   ' Insure that path portion of string uses the
   ' same case as the real pathname.
   '
   If InStr(PathIn, "\") Then
      For i = Len(PathIn) To 1 Step -1
         If Mid$(PathIn, i, 1) = "\" Then
            '
            ' Found end of previous directory.
            ' Recurse back up into path.
            '
            PathOut = ProperCasePath(left(PathIn, i - 1)) & "\"
            '
            ' Use FFF to proper-case current directory.
            '
            hSearch = FindFirstFile(PathIn, wfd)
            If hSearch <> INVALID_HANDLE_VALUE Then
               Call FindClose(hSearch)
               If wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                  ProperCasePath = PathOut & TrimNull(wfd.cFileName)
               End If
            End If
            '
            ' Bail out of loop.
            '
            Exit For
         End If
      Next i
   Else
      '
      ' Just a drive letter and colon,
      ' upper-case and return.
      '
      ProperCasePath = UCase$(PathIn)
   End If
End Function

Private Function TrimNull(ByVal strIn As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/1999
   ' * Time             : 16:32
   ' * Module Name      : class_FileInfo
   ' * Module Filename  : FileInfo.cls
   ' * Procedure Name   : TrimNull
   ' * Parameters       :
   ' *                    ByVal StrIn As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nul              As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(strIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = left$(strIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim$(strIn)
   End Select
End Function

