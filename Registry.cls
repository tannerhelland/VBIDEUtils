VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : Steve McMahon
' * Web Site         : http://www.dogma.demon.co.uk/
' * E-Mail           : steve@dogma.demon.co.uk
' * Date             : 25/09/98
' * Time             : 14:49
' * Module Name      : class_Registry
' * Module Filename  : Registry.cls
' **********************************************************************
' * Comments         : This class is an easy, self-contained way to get
' *  complete access to the Windows registry.
' *  Simple methods allow you to create, enumerate and delete keys
' *  and values in the registry, without restriction.
' *  You can even read/write binary data to the registry.
' *
' * Example :
' *  get a String Value from the Registry
' *    Dim clsRegistry As New class_Registry
' *    With clsRegistry
' *       .ClassKey = HKEY_LOCAL_MACHINE
' *       .SectionKey = "Software\PrintPreview"
' *       .ValueKey = "Version"
' *       .ValueType = REG_SZ
' *       sTip = .Value
' *    End With
' *
' *  Save a Form's position to the Registry
' *    Dim clsRegistry As New class_Registry
' *    With clsRegistry
' *       .ClassKey = HKEY_CURRENT_USER
' *       ' You don't need to check if this key already exists
' *       ' - the class will create it for you
' *       .SectionKey = "Software\" & App.EXEName & "\" & frmThis.name
' *       .ValueKey = "Maximized"
' *       .ValueType = REG_DWORD
' *       .Value = (frmThis.WindowState = vbMaximized)
' *       If (frmThis.WindowState <> vbMaximized)
' *           .ValueKey = "Left"
' *           .Value = frmThis.Left
' *           .ValueKey = "Top"
' *           .Value = frmThis.Top
' *           .ValueKey = "Width"
' *           .Value = frmThis.Width
' *           .ValueKey = "Height"
' *           .Value = frmThis.Height
' *       End If
' *    End With
' *
' **********************************************************************

Option Explicit

' *** Registry Specific Access Rights
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F

' *** Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

' *** Key creation/open disposition
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

' *** masks for the predefined standard access types
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF

' *** Define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234 ' ***   dderror
Private Const ERROR_NO_MORE_ITEMS = 259

' *** Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
   nLength              As Long
   lpSecurityDescriptor As Long
   bInheritHandle       As Boolean
End Type

Private Type FILETIME
   dwLowDateTime        As Long
   dwHighDateTime       As Long
End Type

' *** Registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "Advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExStr Lib "Advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "Advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "Advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "Advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueExStr Lib "Advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "Advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "Advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, szData As Byte, ByRef lpcbData As Long) As Long

Private Declare Function RegCreateKeyEx Lib "Advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

' ***  Other declares:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Public Enum ERegistryClassConstants
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum

Public Enum ERegistryValueTypes
   ' *** Predefined Value Types
   REG_NONE = (0)                         ' *** No value type
   REG_SZ = (1)                           ' *** Unicode nul terminated string
   REG_EXPAND_SZ = (2)                    ' *** Unicode nul terminated string w/enviornment var
   REG_BINARY = (3)                       ' *** Free form binary
   REG_DWORD = (4)                        ' *** 32-bit number
   REG_DWORD_LITTLE_ENDIAN = (4)          ' *** 32-bit number (same as REG_DWORD)
   REG_DWORD_BIG_ENDIAN = (5)             ' *** 32-bit number
   REG_LINK = (6)                         ' *** Symbolic Link (unicode)
   REG_MULTI_SZ = (7)                     ' *** Multiple Unicode strings
   REG_RESOURCE_LIST = (8)                ' *** Resource list in the resource map
   REG_FULL_RESOURCE_DESCRIPTOR = (9)     ' *** Resource list in the hardware description
   REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum

Private m_hClassKey     As Long
Private m_sSectionKey   As String
Private m_sValueKey     As String
Private m_vValue        As Variant
Private m_sSetValue     As String
Private m_vDefault      As Variant
Private m_eValueType    As ERegistryValueTypes

Public Property Get KeyExists() As Boolean
   ' *** KeyExists = bCheckKeyExists( _
   ' ***                 m_hClassKey, _
   ' ***                 m_sSectionKey _
   ' ***             )

   Dim hKey             As Long

   If RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, hKey) = ERROR_SUCCESS Then
      KeyExists = True
      RegCloseKey hKey
   Else
      KeyExists = False
   End If

End Property

Public Function CreateKey() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : CreateKey
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim tSA              As SECURITY_ATTRIBUTES
   Dim hKey             As Long
   Dim lCreate          As Long
   Dim e                As Long

   ' *** Open or Create the key
   e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tSA, hKey, lCreate)
   If e Then
      err.Raise 26001, App.EXEName & ".cRegistry", "Failed to create registry Key: ' *** " & m_sSectionKey
   Else
      CreateKey = (e = ERROR_SUCCESS)
      ' *** Close the key
      RegCloseKey hKey
   End If

End Function

Public Function DeleteKey() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : DeleteKey
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim e                As Long
   e = RegDeleteKey(m_hClassKey, m_sSectionKey)
   If e Then
      err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: ' *** " & m_hClassKey & "' *** ,Section: ' *** " & m_sSectionKey
   Else
      DeleteKey = (e = ERROR_SUCCESS)
   End If

End Function

Public Function DeleteValue() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : DeleteValue
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim e                As Long
   Dim hKey             As Long

   e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ALL_ACCESS, hKey)
   If e Then
      err.Raise 26001, App.EXEName & ".cRegistry", "Failed to open key ' *** " & m_hClassKey & "' *** ,Section: ' *** " & m_sSectionKey & "' ***  for delete access"
   Else
      e = RegDeleteValue(hKey, m_sValueKey)
      If e Then
         err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: ' *** " & m_hClassKey & "' *** ,Section: ' *** " & m_sSectionKey & "' *** ,Key: ' *** " & m_sValueKey
      Else
         DeleteValue = (e = ERROR_SUCCESS)
      End If
   End If

End Function

Public Property Get Value() As Variant
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : Value
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim vValue           As Variant
   Dim cData            As Long
   Dim sData            As String
   Dim ordType          As Long
   Dim e                As Long
   Dim hKey             As Long

   e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
   ' *** Api Raise If e

   e = RegQueryValueExLong(hKey, m_sValueKey, 0&, ordType, 0&, cData)
   If e And e <> ERROR_MORE_DATA Then
      Value = m_vDefault
      Exit Property
   End If

   m_eValueType = ordType
   Select Case ordType
      Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
         Dim iData            As Long
         e = RegQueryValueExLong(hKey, m_sValueKey, 0&, ordType, iData, cData)
         vValue = CLng(iData)

      Case REG_DWORD_BIG_ENDIAN  ' ***  Unlikely, but you never know
         Dim dwData           As Long
         e = RegQueryValueExLong(hKey, m_sValueKey, 0&, ordType, dwData, cData)
         vValue = SwapEndian(dwData)

      Case REG_SZ, REG_MULTI_SZ ' ***  Same thing to Visual Basic
         sData = String$(cData - 1, 0)
         e = RegQueryValueExStr(hKey, m_sValueKey, 0&, ordType, sData, cData)
         vValue = sData

      Case REG_EXPAND_SZ
         sData = String$(cData - 1, 0)
         e = RegQueryValueExStr(hKey, m_sValueKey, 0&, ordType, sData, cData)
         vValue = ExpandEnvStr(sData)

         ' ***  Catch REG_BINARY and anything else
      Case Else
         Dim abData()         As Byte
         ReDim abData(cData)
         e = RegQueryValueExByte(hKey, m_sValueKey, 0&, ordType, abData(0), cData)
         vValue = abData

   End Select
   Value = vValue

End Property

Public Property Let Value(ByVal vValue As Variant)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : Value
   ' * Parameters       :
   ' *                    ByVal vValue As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim ordType          As Long
   Dim c                As Long
   Dim hKey             As Long
   Dim e                As Long
   Dim lCreate          As Long
   Dim tSA              As SECURITY_ATTRIBUTES

   ' *** Open or Create the key
   e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tSA, hKey, lCreate)

   If e Then
      'Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: ' *** " & m_hClassKey & "' *** ,Section: ' *** " & m_sSectionKey & "' *** ,Key: ' *** " & m_sValueKey & "' ***  to value: ' *** " & m_vValue & "' *** "
   Else

      Select Case m_eValueType
         Case REG_BINARY
            If (VarType(vValue) = vbArray + vbByte) Then
               Dim ab()             As Byte
               ab = vValue
               ordType = REG_BINARY
               c = UBound(ab) - LBound(ab) - 1
               e = RegSetValueExByte(hKey, m_sValueKey, 0&, ordType, ab(0), c)
            Else
               err.Raise 26001
            End If
         Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
            If (VarType(vValue) = vbInteger) Or (VarType(vValue) = vbLong) Then
               Dim i                As Long
               i = vValue
               ordType = REG_DWORD
               e = RegSetValueExLong(hKey, m_sValueKey, 0&, ordType, i, 4)
            End If
         Case REG_SZ, REG_EXPAND_SZ
            Dim s                As String, iPos As Long
            s = vValue & ""
            ordType = REG_SZ
            ' ***  Assume anything with two non-adjacent percents is expanded string
            iPos = InStr(s, "%")
            If iPos Then
               If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
            End If
            c = Len(s) + 1
            e = RegSetValueExStr(hKey, m_sValueKey, 0&, ordType, s, c)

            ' ***  User should convert to a compatible type before calling
         Case Else
            e = ERROR_INVALID_DATA

      End Select

      If Not e Then
         m_vValue = vValue
      Else
         err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: ' *** " & m_hClassKey & "' *** ,Section: ' *** " & m_sSectionKey & "' *** ,Key: ' *** " & m_sValueKey & "' ***  to value: ' *** " & m_vValue & "' *** "
      End If

      ' *** Close the key
      RegCloseKey hKey

   End If

End Property

Public Function EnumerateValues(ByRef sKeyNames() As String, ByRef iKeyCount As Long) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : EnumerateValues
   ' * Parameters       :
   ' *                    ByRef sKeyNames() As String
   ' *                    ByRef iKeyCount As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lResult          As Long
   Dim hKey             As Long
   Dim sName            As String
   Dim lNameSize        As Long
   Dim lIndex           As Long
   Dim cJunk            As Long
   Dim cNameMax         As Long
   Dim ft               As Currency

   ' ***  Log "EnterEnumerateValues"

   iKeyCount = 0
   Erase sKeyNames()

   lIndex = 0
   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
   If (lResult = ERROR_SUCCESS) Then
      ' ***  Log "OpenedKey:" & m_hClassKey & "," & m_sSectionKey
      lResult = RegQueryInfoKey(hKey, "", cJunk, 0, cJunk, cJunk, cJunk, cJunk, cNameMax, cJunk, cJunk, ft)
      Do While lResult = ERROR_SUCCESS

         ' *** Set buffer space
         lNameSize = cNameMax + 1
         sName = String$(lNameSize, 0)
         If (lNameSize = 0) Then lNameSize = 1

         ' ***  Log "Requesting Next Value"

         ' *** Get value name:
         lResult = RegEnumValue(hKey, lIndex, sName, lNameSize, 0&, 0&, 0&, 0&)

         ' ***  Log "RegEnumValue returned:" & lResult
         If (lResult = ERROR_SUCCESS) Then

            ' ***  Although in theory you can also retrieve the actual
            ' ***  value and type here, I found it always (ultimately) resulted in
            ' ***  a GPF, on Win95 and NT.  Why?  Can anyone help?

            sName = left$(sName, lNameSize)
            ' ***  Log "Enumerated value:" & sName

            iKeyCount = iKeyCount + 1
            ReDim Preserve sKeyNames(1 To iKeyCount) As String
            sKeyNames(iKeyCount) = sName
         End If
         lIndex = lIndex + 1
      Loop
   End If
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If

   ' ***  Log "Exit Enumerate Values"
   EnumerateValues = True
   Exit Function

EnumerateValuesError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   err.Raise vbObjectError + 1048 + 26003, App.EXEName & ".cRegistry", err.Description
   Exit Function

End Function

Public Function EnumerateSections(ByRef sSect() As String, ByRef iSectCount As Long) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : EnumerateSections
   ' * Parameters       :
   ' *                    ByRef sSect() As String
   ' *                    ByRef iSectCount As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim lResult          As Long
   Dim hKey             As Long
   Dim szBuffer         As String
   Dim lBuffSize        As Long
   Dim lIndex           As Long
   Dim iPos             As Long

   On Error GoTo EnumerateSectionsError

   iSectCount = 0
   Erase sSect
   ' ***
   lIndex = 0

   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
   Do While lResult = ERROR_SUCCESS
      ' *** Set buffer space
      szBuffer = String$(255, 0)
      lBuffSize = Len(szBuffer)

      ' *** Get next value
      lResult = RegEnumKey(hKey, lIndex, szBuffer, lBuffSize)

      If (lResult = ERROR_SUCCESS) Then
         iSectCount = iSectCount + 1
         ReDim Preserve sSect(1 To iSectCount) As String
         iPos = InStr(szBuffer, Chr$(0))
         If (iPos > 0) Then
            sSect(iSectCount) = left$(szBuffer, iPos - 1)
         Else
            sSect(iSectCount) = left$(szBuffer, lBuffSize)
         End If
      End If

      lIndex = lIndex + 1
   Loop
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   EnumerateSections = True
   Exit Function

EnumerateSectionsError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   err.Raise vbObjectError + 1048 + 26002, App.EXEName & ".cRegistry", err.Description
   Exit Function

End Function

Public Sub CreateEXEAssociation(ByVal sExePath As String, ByVal sClassName As String, ByVal sClassDescription As String, ByVal sAssociation As String, Optional ByVal lDefaultIconIndex As Long = -1)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : CreateEXEAssociation
   ' * Parameters       :
   ' *                    ByVal sExePath As String
   ' *                    ByVal sClassName As String
   ' *                    ByVal sClassDescription As String
   ' *                    ByVal sAssociation As String
   ' *                    Optional ByVal lDefaultIconIndex As Long = -1
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ClassKey = HKEY_CLASSES_ROOT
   SectionKey = "." & sAssociation
   ValueKey = ""
   Value = sClassName
   SectionKey = "." & sAssociation & "\shell\open\command"
   ValueKey = ""
   Value = sExePath & " ""%1"""

   SectionKey = sClassName
   ValueKey = ""
   Value = sClassDescription
   SectionKey = sClassName & "\shell\open\command"
   ValueKey = sExePath & " ""%1"""
   If lDefaultIconIndex > -1 Then
      SectionKey = sClassName & "\DefaultIcon"
      ValueKey = ""
      Value = sExePath & "," & CStr(lDefaultIconIndex)
   End If

End Sub

Public Property Get ValueType() As ERegistryValueTypes
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : ValueType
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ValueType = m_eValueType

End Property

Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : ValueType
   ' * Parameters       :
   ' *                    ByVal eValueType As ERegistryValueTypes
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_eValueType = eValueType

End Property

Public Property Get ClassKey() As ERegistryClassConstants
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : ClassKey
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ClassKey = m_hClassKey

End Property

Public Property Let ClassKey(ByVal eKey As ERegistryClassConstants)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : ClassKey
   ' * Parameters       :
   ' *                    ByVal eKey As ERegistryClassConstants
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_hClassKey = eKey

End Property

Public Property Get SectionKey() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : SectionKey
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   SectionKey = m_sSectionKey

End Property

Public Property Let SectionKey(ByVal sSectionKey As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : SectionKey
   ' * Parameters       :
   ' *                    ByVal sSectionKey As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_sSectionKey = sSectionKey

End Property

Public Property Get ValueKey() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : ValueKey
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ValueKey = m_sValueKey

End Property

Public Property Let ValueKey(ByVal sValueKey As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : ValueKey
   ' * Parameters       :
   ' *                    ByVal sValueKey As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_sValueKey = sValueKey

End Property

Public Property Get Default() As Variant
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : Default
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Default = m_vDefault

End Property

Public Property Let Default(ByVal vDefault As Variant)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : Default
   ' * Parameters       :
   ' *                    ByVal vDefault As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_vDefault = vDefault

End Property

Private Function SwapEndian(ByVal dw As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : SwapEndian
   ' * Parameters       :
   ' *                    ByVal dw As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
   CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
   CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
   CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1

End Function

Private Function ExpandEnvStr(sData As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : ExpandEnvStr
   ' * Parameters       :
   ' *                    sData As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim c                As Long, s As String

   ' ***  Get the length
   s = "" ' ***  Needed to get around Windows 95 limitation
   c = ExpandEnvironmentStrings(sData, s, c)
   ' ***  Expand the string
   s = String$(c - 1, 0)
   c = ExpandEnvironmentStrings(sData, s, c)
   ExpandEnvStr = s

End Function

Public Function GetRegistrySetting(nKeyRoot As ERegistryClassConstants, sKeyName As String, sSubKeyRef As String, sDefault As String, Optional vValueType As ERegistryValueTypes = REG_EXPAND_SZ) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : GetRegistrySetting
   ' * Parameters       :
   ' *                    nKeyRoot As ERegistryClassConstants
   ' *                    sKeyName As String
   ' *                    sSubKeyRef As String
   ' *                    sDefault As String
   ' *                    Optional vValueType As ERegistryValueTypes = REG_EXPAND_SZ
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Return the value of a settings in the registry

   ClassKey = nKeyRoot
   SectionKey = sKeyName
   ValueKey = sSubKeyRef
   Default = sDefault
   ValueType = vValueType

   GetRegistrySetting = Value

End Function

Public Sub SetRegistrySetting(nKeyRoot As ERegistryClassConstants, sSubKeyRef As String, sKeyName As String, sValue As String, Optional vValueType As ERegistryValueTypes = REG_EXPAND_SZ)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:58
   ' * Module Name      : class_Registry
   ' * Module Filename  : Registry.cls
   ' * Procedure Name   : SetRegistrySetting
   ' * Parameters       :
   ' *                    nKeyRoot As ERegistryClassConstants
   ' *                    sSubKeyRef As String
   ' *                    sKeyName As String
   ' *                    sValue As String
   ' *                    Optional vValueType As ERegistryValueTypes = REG_EXPAND_SZ
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Set a value in the registry

   ClassKey = nKeyRoot
   SectionKey = sSubKeyRef
   ValueKey = sKeyName
   ValueType = vValueType
   Value = sValue

End Sub

