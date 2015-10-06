VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_GeneralUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 03/11/1999
' * Time             : 15:01
' * Module Name      : GeneralUtilities
' * Module Filename  : GeneralUtilities.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Public Function GetDirectoryName(ByVal fullyQualifiedFileName As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : GeneralUtilities
   ' * Module Filename  : GeneralUtilities.cls
   ' * Procedure Name   : GetDirectoryName
   ' * Parameters       :
   ' *                    ByVal fullyQualifiedFileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim done             As Boolean
   Dim startPos         As Long

   On Error GoTo errHandle
   startPos = 1
   done = False

   Do Until done
      Location = InStr(startPos, fullyQualifiedFileName, "\")

      If Location = 0 Then
         done = True
      Else
         startPos = Location + 1
      End If
   Loop

   GetDirectoryName = left$(fullyQualifiedFileName, Len(fullyQualifiedFileName) - (Len(fullyQualifiedFileName) - (startPos - 1)))

   Exit Function
errHandle:

   Call handleError.handleError(err, "GeneralUtilities Class : GetDirectoryName Method")

End Function

Function FormatNumber(number As Double) As Variant
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : GeneralUtilities
   ' * Module Filename  : GeneralUtilities.cls
   ' * Procedure Name   : FormatNumber
   ' * Parameters       :
   ' *                    number As Double
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo errHandle

   If InStr(1, CStr(number), ".") = 0 Then
      FormatNumber = (CStr(number) & ".0")
   Else
      FormatNumber = number
   End If

   Exit Function

errHandle:

   Call handleError.handleError(err, "GeneralUtilities Class : FormatNumber Method")

End Function

Public Function DetermineMonthName(monthNum As Long) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : GeneralUtilities
   ' * Module Filename  : GeneralUtilities.cls
   ' * Procedure Name   : DetermineMonthName
   ' * Parameters       :
   ' *                    monthNum As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo errHandle

   Select Case monthNum
      Case 1
         DetermineMonthName = "January"
      Case 2
         DetermineMonthName = "February"
      Case 3
         DetermineMonthName = "March"
      Case 4
         DetermineMonthName = "April"
      Case 5
         DetermineMonthName = "May"
      Case 6
         DetermineMonthName = "June"
      Case 7
         DetermineMonthName = "July"
      Case 8
         DetermineMonthName = "August"
      Case 9
         DetermineMonthName = "September"
      Case 10
         DetermineMonthName = "October"
      Case 11
         DetermineMonthName = "November"
      Case 12
         DetermineMonthName = "December"
      Case Else
         DetermineMonthName = "Invalid parameter passed to determinemonthname"
   End Select

   Exit Function
errHandle:

   Call handleError.handleError(err, "General utilities Class : DetermineMonthName Method")

End Function

Public Function DetermineMonthNum(monthName As String) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : GeneralUtilities
   ' * Module Filename  : GeneralUtilities.cls
   ' * Procedure Name   : DetermineMonthNum
   ' * Parameters       :
   ' *                    monthName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo errHandle

   Select Case UCase$(monthName)
      Case "JANUARY"
         DetermineMonthNum = 1
      Case "FEBRUARY"
         DetermineMonthNum = 2
      Case "MARCH"
         DetermineMonthNum = 3
      Case "ARPIL"
         DetermineMonthNum = 4
      Case "MAY"
         DetermineMonthNum = 5
      Case "JUNE"
         DetermineMonthNum = 6
      Case "JULY"
         DetermineMonthNum = 7
      Case "AUGUST"
         DetermineMonthNum = 8
      Case "SEPTEMBER"
         DetermineMonthNum = 9
      Case "OCTOBER"
         DetermineMonthNum = 10
      Case "NOVEMBER"
         DetermineMonthNum = 11
      Case "DECEMBER"
         DetermineMonthNum = 12

   End Select

   Exit Function

errHandle:

   Call handleError.handleError(err, "GeneralUtilities Class : DetermineMonthNum Method")

End Function

Public Function GetFileName(fullyQualifiedFileName As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : GeneralUtilities
   ' * Module Filename  : GeneralUtilities.cls
   ' * Procedure Name   : GetFileName
   ' * Parameters       :
   ' *                    fullyQualifiedFileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim startPos         As Long

   On Error GoTo errHandle

   'Pull the file name from a fully qualified file name
   startPos = 1
   Do Until InStr(startPos, fullyQualifiedFileName, "\") = 0
      startPos = InStr(startPos, fullyQualifiedFileName, "\") + 1
   Loop

   'Startpos is now the first letter of the file name

   GetFileName = right$(fullyQualifiedFileName, Len(fullyQualifiedFileName) - (startPos - 1)) 'Subtract one from start pos to capture the entire filename

   Exit Function

errHandle:

   Call handleError.handleError(err, "GeneralUtilities Class : GetFileName Method")

End Function
