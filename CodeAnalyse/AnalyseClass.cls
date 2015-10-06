VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 03/11/1999
' * Time             : 15:01
' * Module Name      : class_Class
' * Module Filename  : class_Class.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

'local variable(s) to hold property value(s)
Private mvarClassType   As String 'local copy
Private mvarClassName   As String 'local copy
Private mvarNumLines    As Long 'local copy
Private mvarNumSubs     As Variant 'local copy
Private mvarNumFunctions As Long 'local copy
Private mvarNumProperties As Long 'local copy
Private mvarCountComments As Boolean 'local copy
Private mvarCountSpace  As Boolean 'local copy
Private mvarSubs        As Collection 'local copy
Private msLastErrorMessage As String
Private mvarIsFunction  As Boolean
Public Event Process(StatusMsg As String)

Public Property Get LastErrorMessage() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : LastErrorMessage
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   LastErrorMessage = msLastErrorMessage
End Property

Public Property Set Subs(ByVal vData As Collection)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : Subs
   ' * Parameters       :
   ' *                    ByVal vData As Collection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning an Object to the property, on the left side of a Set statement.
   'Syntax: Set x.Subs = Form1
   Set mvarSubs = vData
End Property

Public Property Get Subs() As Collection
   'used when retrieving value of a property, on the right side of an assignment.
   Set Subs = mvarSubs
End Property

Public Property Let CountSpace(ByVal vData As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : CountSpace
   ' * Parameters       :
   ' *                    ByVal vData As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.CountSpace = 5
   mvarCountSpace = vData
End Property

Public Property Get CountSpace() As Boolean
   'used when retrieving value of a property, on the right side of an assignment.
   CountSpace = mvarCountSpace
End Property

Public Property Let CountComments(ByVal vData As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : CountComments
   ' * Parameters       :
   ' *                    ByVal vData As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.CountComments = 5
   mvarCountComments = vData
End Property

Public Property Get CountComments() As Boolean
   'used when retrieving value of a property, on the right side of an assignment.
   CountComments = mvarCountComments
End Property

Public Property Get NumProperties() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : NumProperties
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   NumProperties = mvarNumProperties
End Property

Public Property Get NumFunctions() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : NumFunctions
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   NumFunctions = mvarNumFunctions
End Property

Public Property Get NumSubs() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : NumSubs
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   NumSubs = mvarNumSubs

End Property

Public Property Let NumLines(ByVal vData As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : NumLines
   ' * Parameters       :
   ' *                    ByVal vData As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.NumLines = 5
   mvarNumLines = vData
End Property

Public Property Get NumLines() As Long
   'used when retrieving value of a property, on the right side of an assignment.
   NumLines = mvarNumLines
End Property

Public Property Let ClassName(ByVal vData As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : ClassName
   ' * Parameters       :
   ' *                    ByVal vData As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.ClassName = 5
   mvarClassName = vData
End Property

Public Property Get ClassName() As String
   'used when retrieving value of a property, on the right side of an assignment.
   ClassName = mvarClassName
End Property

Public Property Let ClassType(ByVal vData As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : ClassType
   ' * Parameters       :
   ' *                    ByVal vData As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.ClassType = 5
   mvarClassType = vData
End Property

Public Property Get ClassType() As String
   'used when retrieving value of a property, on the right side of an assignment.
   ClassType = mvarClassType
End Property

Public Function Read(ByVal fileToRead As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : Read
   ' * Parameters       :
   ' *                    ByVal fileToRead As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Read

   Dim nFile            As Long
   Dim clsSub           As class_Sub
   Dim sLine            As String
   Dim sKeyWords()      As String
   Dim sTWA()           As String
   Dim sTmp             As String
   Dim nTmp             As String
   Dim nPos             As Long
   Dim nPos1            As Long
   Dim nI               As Long
   Dim bExit            As Boolean

   Dim sFunction        As String

   If Dir(fileToRead) = "" Then
      Read = False
      msLastErrorMessage = fileToRead & " Does not Exist"
      Exit Function
   End If

   nFile = FreeFile
   Open fileToRead For Input As #nFile

   bExit = False

   Set clsSub = New class_Sub
   clsSub.SubType = "Declarations"
   RaiseEvent Process("Processing " & fileToRead)
   Line Input #nFile, sLine
   Do While True
      ' *** Read the first word on the line
      If sLine <> "" Then
         Call Main_Module.Split(Trim$(sLine), sKeyWords(), " ")
      Else
         sKeyWords(1) = ""
      End If
      Select Case UCase$(sKeyWords(1))
         Case "ATTRIBUTE":
            If UCase$(sKeyWords(2)) = "VB_NAME" Then
               ClassName = Replace(sKeyWords(UBound(sKeyWords)), """", "")
            End If
         Case "PRIVATE", "PUBLIC"
            ' *** Check to see if we are in a function
            If Not mvarIsFunction Then
               'If uCase$(sKeyWords(2)) = "DECLARE" Then
               '   sTmp = uCase$(sKeyWords(3))
               '   nTmp = 3
               'Else
               sTmp = UCase$(sKeyWords(2))
               nTmp = 2
               'End If

               ' *** Check to See if we are reading a Function
               Select Case sTmp
                  Case "FUNCTION", "SUB", "PROPERTY"   ' *** This is a function
                     mvarIsFunction = True
                     sFunction = ""
                     Call mvarSubs.Add(clsSub)
                     Set clsSub = New class_Sub
                     If UCase$(sKeyWords(2)) = "PROPERTY" Then
                        Call Main_Module.Split(Trim$(sKeyWords(4)), sTWA, "(")
                        clsSub.SubName = sKeyWords(3) & " " & Trim$(sTWA(1))
                     Else
                        Call Main_Module.Split(Trim$(sKeyWords(3)), sTWA, "(")
                        clsSub.SubName = Trim$(sTWA(1))
                     End If
                     clsSub.SubType = sKeyWords(2)
                     clsSub.NumLines = clsSub.NumLines + 1
                  Case Else  ' *** This is just a variable declaration
                     If (left(Trim$(sLine), 1) = "'" And mvarCountComments = False) Or (sLine = "" And mvarCountSpace = False) Then
                     Else
                        sLine = Trim$(sLine)
                        sTmp = LCase$(sLine)
                        nPos = InStr(sTmp, "as ")
                        If (nPos > 0) And (left$(sLine, 1) <> "'") And (left$(sTmp, 4) <> "rem ") Then
                           sLine = Replace(sLine, ",", " ")
                           sLine = Replace(sLine, "            ", " ")
                           sLine = Replace(sLine, "           ", " ")
                           sLine = Replace(sLine, "          ", " ")
                           sLine = Replace(sLine, "         ", " ")
                           sLine = Replace(sLine, "        ", " ")
                           sLine = Replace(sLine, "       ", " ")
                           sLine = Replace(sLine, "      ", " ")
                           sLine = Replace(sLine, "     ", " ")
                           sLine = Replace(sLine, "    ", " ")
                           sLine = Replace(sLine, "   ", " ")
                           sLine = Replace(sLine, "  ", " ")

                           ' *** Remove all array
                           nPos = InStr(sLine, "(")
                           If nPos > 0 Then
                              nPos1 = InStr(sLine, ")")
                              sLine = left$(sLine, nPos - 1) & Mid$(sLine, nPos1 + 1)
                           End If

                           Call Main_Module.Split(sLine, sTWA(), " ")
                           If ((LCase$(sTWA(1)) = "dim") Or _
                              (LCase$(sTWA(1)) = "public") Or _
                              (LCase$(sTWA(1)) = "private") Or _
                              (LCase$(sTWA(1)) = "global")) And _
                              (LCase$(sTWA(2)) <> "declare") And _
                              (LCase$(sTWA(2)) <> "const") And _
                              (LCase$(sTWA(2)) <> "withevents") And _
                              (LCase$(sTWA(2)) <> "event") Then
                              For nI = 1 To UBound(sTWA)
                                 If (LCase$(sTWA(nI)) <> "dim") And _
                                    (LCase$(sTWA(nI)) <> "public") And _
                                    (LCase$(sTWA(nI)) <> "private") And _
                                    (LCase$(sTWA(nI)) <> "global") Then
                                    If (LCase$(sTWA(nI)) = "as") And (nI = UBound(sTWA) - 1) Then
                                       Exit For
                                    Else
                                       If (left(sTWA(nI), 1) = "'") Then Exit For
                                       If (LCase$(sTWA(nI)) = "as") Then
                                          If nI + 1 < UBound(sTWA) Then
                                             If LCase$(sTWA(nI + 1)) = "new" Then nI = nI + 1
                                          End If
                                          nI = nI + 1
                                       Else
                                          ' *** Add the variable
                                          clsSub.mcolVariable.Add sTWA(nI), sTWA(nI)
                                       End If
                                    End If
                                 End If
                              Next
                           End If
                        End If

                        clsSub.NumLines = clsSub.NumLines + 1
                     End If
               End Select
            Else  ' *** We are in a function
               If (left(Trim$(sLine), 1) = "'" And mvarCountComments = False) Or (sLine = "" And mvarCountSpace = False) Then
               Else
                  clsSub.NumLines = clsSub.NumLines + 1
               End If
               clsSub.NumLines = clsSub.NumLines + 1
            End If
         Case Else
            If mvarIsFunction Then
               Select Case UCase$(sKeyWords(1))
                  Case "END":  ' *** The function is ending
                     If UBound(sKeyWords) > 1 Then
                        If (UCase$(sKeyWords(2)) = "FUNCTION") Or (UCase$(sKeyWords(2)) = "SUB") Or (UCase$(sKeyWords(2)) = "PROPERTY") Then
                           mvarIsFunction = False
                           clsSub.NumLines = clsSub.NumLines + 1

                           ' *** Verify all the variables
                           For nI = 1 To clsSub.mcolVariable.Count
                              nPos = InStr(sFunction, clsSub.mcolVariable(nI))
                              If nPos = 0 Then
                                 clsSub.mcolUnusedVar.Add clsSub.mcolVariable(nI), clsSub.mcolVariable(nI)
                              End If
                           Next

                           sFunction = ""
                        End If
                     End If

                  Case "DIM": ' *** add the variable to the list of variable
                     sLine = Trim$(sLine)
                     sTmp = LCase$(sLine)
                     nPos = InStr(sTmp, "as ")
                     If nPos > 0 Then
                        sLine = Replace(sLine, ",", " ")
                        sLine = Replace(sLine, "            ", " ")
                        sLine = Replace(sLine, "           ", " ")
                        sLine = Replace(sLine, "          ", " ")
                        sLine = Replace(sLine, "         ", " ")
                        sLine = Replace(sLine, "        ", " ")
                        sLine = Replace(sLine, "       ", " ")
                        sLine = Replace(sLine, "      ", " ")
                        sLine = Replace(sLine, "     ", " ")
                        sLine = Replace(sLine, "    ", " ")
                        sLine = Replace(sLine, "   ", " ")
                        sLine = Replace(sLine, "  ", " ")

                        ' *** Remove all array
                        nPos = InStr(sLine, "(")
                        If nPos > 0 Then
                           nPos1 = InStr(sLine, ")")
                           sLine = left$(sLine, nPos - 1) & Mid$(sLine, nPos1 + 1)
                        End If

                        Call Main_Module.Split(sLine, sTWA(), " ")
                        For nI = 1 To UBound(sTWA)
                           If LCase$(sTWA(nI)) <> "dim" Then
                              If (LCase$(sTWA(nI)) = "as") And (nI = UBound(sTWA) - 1) Then
                                 Exit For
                              Else
                                 If (left(sTWA(nI), 1) = "'") Or (left(sTWA(nI), 1) = "*") Then Exit For
                                 If (LCase$(sTWA(nI)) = "as") Then
                                    If nI + 1 < UBound(sTWA) Then
                                       If LCase$(sTWA(nI + 1)) = "new" Then nI = nI + 1
                                    End If
                                    nI = nI + 1
                                 Else
                                    ' *** Add the variable
                                    clsSub.mcolVariable.Add sTWA(nI), sTWA(nI)
                                 End If
                              End If
                           End If
                        Next
                     End If

                  Case Else
                     ' *** Keep the code of the function
                     If (((left$(Trim$(sLine), 1) = "'") Or (LCase$(left$(Trim$(sLine), 3)) = "rem")) And mvarCountComments = False) Or (sLine = "" And mvarCountSpace = False) Then  'This is a comment line
                     Else
                        sFunction = sFunction & sLine & vbCrLf

                        clsSub.NumLines = clsSub.NumLines + 1
                     End If
               End Select
            Else
               Select Case UCase$(sKeyWords(1))
                  Case "FUNCTION", "SUB", "PROPERTY"
                     sFunction = ""
                     mvarIsFunction = True
                     Call mvarSubs.Add(clsSub)
                     Set clsSub = New class_Sub

                     If UCase$(sKeyWords(1)) = "PROPERTY" Then
                        Call Main_Module.Split(Trim$(sKeyWords(3)), sTWA, "(")
                        clsSub.SubName = sKeyWords(2) & " " & Trim$(sTWA(1))
                     Else
                        Call Main_Module.Split(Trim$(sKeyWords(2)), sTWA, "(")
                        clsSub.SubName = Trim$(sTWA(1))
                     End If
                     clsSub.SubType = sKeyWords(1)
                     clsSub.NumLines = clsSub.NumLines + 1
                  Case Else
                     If (left(sLine, 1) = "'" And mvarCountComments = False) Or (sLine = "" And mvarCountSpace = False) Then
                     Else
                        clsSub.NumLines = clsSub.NumLines + 1
                     End If
               End Select
            End If
      End Select
      If bExit Then Exit Do
      Line Input #nFile, sLine
      bExit = EOF(nFile)
   Loop
   Call mvarSubs.Add(clsSub)  ' *** Add the last sub routine in
   Close #nFile
   Read = True

   ' *** Loop through
   For Each clsSub In Me.Subs
      Select Case clsSub.SubType
         Case "Function"
            mvarNumFunctions = mvarNumFunctions + 1

         Case "Property"
            mvarNumProperties = mvarNumProperties + 1

         Case "Sub"
            mvarNumSubs = mvarNumSubs + 1
      End Select
      mvarNumLines = mvarNumLines + clsSub.NumLines
   Next

EXIT_Read:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_Read:
   '   Select Case MsgBoxTop(Me.hwnd, "Error " & Err.number & ": " & Err.Description & vbCrLf & "in Read", vbAbortRetryIgnore + vbCritical, "Error")
   '      Case vbAbort
   '         msLastErrorMessage = Err.Description
   '         Close #nFile
   '         Read = False
   '
   '         Screen.MousePointer = vbDefault
   '         Resume EXIT_Read
   '      Case vbRetry
   '         Resume
   '      Case vbIgnore
   '         Resume Next
   '   End Select

   Resume EXIT_Read
   Resume

End Function

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   mvarIsFunction = False
   Set mvarSubs = New Collection
End Sub

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Class
   ' * Module Filename  : class_Class.cls
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set mvarSubs = Nothing
End Sub

