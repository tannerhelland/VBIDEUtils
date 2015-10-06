Attribute VB_Name = "GUID_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 22/10/98
' * Time             : 15:50
' * Module Name      : GUID_Module
' * Module Filename  : guid.bas
' **********************************************************************
' * Comments         : Create a Globally Unique Identifier (GUID)
' *
' * Samples:
' *  {3201047B-FA1C-11D0-B3F9-004445535400}
' *  {0547C3D5-FA24-11D0-B3F9-004445535400}
' *
' **********************************************************************

Option Explicit
DefLng A-Z

Private Type GUID
   Data1                As Long
   Data2                As Integer
   Data3                As Integer
   Data4(0 To 7)  As String * 1
End Type

Declare Function CoCreateGuid Lib "ole32.dll" (tGUIDStructure As GUID) As Long

Const mciLen As Integer = 4

Public Function CreateGUID() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 26/11/98
   ' * Time             : 15:27
   ' * Module Name      : GUID_Module
   ' * Module Filename  : guid.bas
   ' * Procedure Name   : CreateGUID
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Create a Globally Unique Identifier (GUID)
   ' *
   ' *
   ' **********************************************************************

   Dim sGUID            As String       'store result here
   Dim tGUID            As GUID         'get into this structure
   If CoCreateGuid(tGUID) = 0 Then 'use API to get the GUID
      With tGUID              'build return string
         sGUID = "{" & PadLeft(Hex(.Data1), mciLen * 2) & "-"
         sGUID = sGUID & PadLeft(Hex(.Data2), mciLen) & "-"
         sGUID = sGUID & PadLeft(Hex(.Data3), mciLen) & "-"
         sGUID = sGUID & FormatGUIDData4(.Data4())
      End With
      sGUID = sGUID & "}"     'ending brace
      CreateGUID = sGUID
   End If

End Function

Private Function FormatGUIDData4(aryData4() As String * 1) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : GUID_Module
   ' * Module Filename  : guid.bas
   ' * Procedure Name   : FormatGUIDData4
   ' * Parameters       :
   ' *                    aryData4() As String * 1
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim i                As Integer      'loop thru the array
   Dim sGUID            As String       'store result here
   Dim sTemp1           As String       'first part here
   Dim sTemp2           As String       'second part here

   For i = LBound(aryData4()) To UBound(aryData4())   'process string array
      If i < 2 Then           'first part
         sTemp1 = sTemp1 & Hex(Asc(aryData4(i)))
      Else                    'second part
         sTemp2 = sTemp2 & Hex(Asc(aryData4(i)))
      End If
   Next
   sGUID = PadLeft(sTemp1, mciLen) & "-" & PadLeft(sTemp2, mciLen * 3) 'pad left with zeros
   FormatGUIDData4 = sGUID                     'return what we created

End Function

Private Function PadLeft(sString As String, iLen As Integer) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:52
   ' * Module Name      : GUID_Module
   ' * Module Filename  : guid.bas
   ' * Procedure Name   : PadLeft
   ' * Parameters       :
   ' *                    sString As String
   ' *                    iLen As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' Pad with left zeros if needed
   Dim sTemp            As String
   sTemp = right$(String$(iLen, "0") & sString, iLen)
   PadLeft = sTemp

End Function
