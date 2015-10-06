Attribute VB_Name = "APIError_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 12/10/1998
' * Time             : 20:20
' * Module Name      : APIError_Module
' * Module Filename  : APIError.bas
' **********************************************************************
' * Comments         :
' * Used to get error messages directly from the
' *   system instead of hard-coding them
' *
' *
' **********************************************************************

Option Explicit

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' *** Status Codes
Public Const INVALID_HANDLE_VALUE = -1&
Public Const ERROR_SUCCESS = 0&

Public Function ReturnAPIError(ErrorCode As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 12/10/1998
   ' * Time             : 20:21
   ' * Module Name      : APIError_Module
   ' * Module Filename  : APIError.bas
   ' * Procedure Name   : ReturnAPIError
   ' * Parameters       :
   ' *                    ErrorCode As Long
   ' **********************************************************************
   ' * Comments         :
   ' * Takes an API error number, and returns
   ' * a descriptive text string of the error
   ' *
   ' **********************************************************************

   Dim sBuffer          As String

   ' *** Allocate the string, then get the system to
   ' *** tell us the error message associated with
   ' *** this error number

   sBuffer = String(256, 0)
   FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, ErrorCode, 0&, sBuffer, Len(sBuffer), 0&

   ' *** Strip the last null, then the last CrLf pair if it exists

   sBuffer = left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   If right$(sBuffer, 2) = Chr$(13) & Chr$(10) Then
      sBuffer = Mid$(sBuffer, 1, Len(sBuffer) - 2)
   End If

   ReturnAPIError = sBuffer

End Function

Public Sub APIError()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 12/10/1998
   ' * Time             : 20:35
   ' * Module Name      : APIError_Module
   ' * Module Filename  : APIError.bas
   ' * Procedure Name   : APIError
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' * Takes an API error number, and returns
   ' * a descriptive text string of the error
   ' *
   ' **********************************************************************

   Dim sError           As String

   On Error GoTo ERROR_APIError

   sError = InputBox("Enter the error number" & Chr$(13) & "you want to know", "Returns API error string for an API error number")

   If IsNumeric(sError) = False Then Exit Sub

   MsgBox ReturnAPIError(CLng(sError)), vbInformation + vbOKOnly, "Error n°" & sError

   Exit Sub

ERROR_APIError:
   'MsgBox "Error n°" & sError & vbCrLf & " Invalid error number" & vbCrLf & "You have to give another one", vbCritical + vbOKOnly, "Error n°" & sError

End Sub
