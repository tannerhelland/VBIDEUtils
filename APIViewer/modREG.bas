Attribute VB_Name = "modREG"
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const SYNCHRONIZE = &H100000
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long        ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Public retval As Long, hkey As Long, subkey As String, newkey As String, phkResult As Long, SA As SECURITY_ATTRIBUTES, create As Long, nvn As String, value As String, value1 As Integer, szbuffer As String, doi As String, tod As String, dadi As Long, id As Date, user As String, size As Long, pwd As String, pass1 As String
Public Function addaslash(instring As String) As String
   If Mid(instring, Len(instring), 1) <> "\" Then
      addaslash = instring & "\"
   Else
      addaslash = instring
   End If
End Function
Public Sub WriteAppName()
   '--------------~writing the Appln. Name on to the registry~-----------------------------------------------------------
   hkey = HKEY_LOCAL_MACHINE
   subkey = "SOFTWARE\"
   newkey = "API TEXT VIEWER BY VENKATESH" & Chr(174)
   retval = RegCreateKeyEx(hkey, subkey & newkey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, create)
   subkey = "SOFTWARE\"
   newkey = "API TEXT VIEWER BY VENKATESH" & Chr(174)
   subkey = addaslash(subkey & newkey)
   newkey = "Company Name"
   retval = RegCreateKeyEx(hkey, subkey & newkey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, create)
   nvn = "Product Done By"
   value = "VENKATESH SOFTWARE SOLUTIONS ©,CHENNAI,INDIA."
   retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
   nvn = "Software Required"
   value = "Best Optmized Performance in WINNT,MS-VISUAL BASIC 6.0"
   retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
   nvn = "Licence"
   value = "To be distributed with code free"
   retval = RegSetValueEx(phkResult, nvn, 0, REG_SZ, value, CLng(Len(value) + 1))
   newkey = "MRUAPIs"
   retval = RegCreateKeyEx(hkey, subkey & newkey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, create)
   '--------------~writing the Appln. Name on to the registry~-----------------------------------------------------------
End Sub
Public Sub GetMRU()
   Dim szbuffer As String, databuff As String, ldbs As Long, count As Integer
   hkey = HKEY_LOCAL_MACHINE
   szbuffer = "SOFTWARE\API TEXT VIEWER BY VENKATESH®" & "\MRUAPIs"
   databuff = Space(255)
   ldbs = Len(databuff)
   Do
      retval = RegOpenKeyEx(hkey, szbuffer, 0, KEY_ALL_ACCESS, phkResult)
      count = count + 1
      If retval = ERROR_SUCCESS Then
         retval = RegQueryValueEx(phkResult, "String" & CStr(count), 0, 0, databuff, Len(databuff))
         If databuff = "" Then Exit Do
         If retval = ERROR_SUCCESS Then
            FINDfrm.finwha.AddItem databuff
         Else
            Exit Do
         End If
      End If
   Loop Until databuff = ""
   RegCloseKey hkey
   RegCloseKey phkResult
End Sub
