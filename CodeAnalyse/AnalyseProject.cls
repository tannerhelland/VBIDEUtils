VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Project"
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
' * Module Name      : class_Project
' * Module Filename  : cntProject.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private mvarProjectName As String 'local copy
Private mvarProjectPath As String 'local copy
Private mvarProjectType As String 'local copy
Private mvarCountComments As Boolean 'local copy
Private mvarCountBlanks As Boolean 'local copy
Private mvarTotalLines  As Long 'local copy
Private mvarClasses     As Collection 'local copy
Private msLastErrorMessage As String
Public msMajorVer       As String
Public msMinorVer       As String
Public msRevisionVer    As String
Public msTitle          As String
Public msExeName32      As String
Public msDescription    As String
Public msVersionComments As String
Public msVersionCompanyName As String
Public msVersionLegalTrademarks As String
Public msVersionProductName As String

Private WithEvents clsClass As class_Class
Attribute clsClass.VB_VarHelpID = -1

Public Event StatusMsg(Msg As String)

Public Property Set Classes(ByVal vData As Collection)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : Classes
   ' * Parameters       :
   ' *                    ByVal vData As Collection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning an Object to the property, on the left side of a Set statement.
   'Syntax: Set x.Classes = Form1
   Set mvarClasses = vData
End Property

Public Property Get Classes() As Collection
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 16:37
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : Classes
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   Set Classes = mvarClasses
End Property

Public Property Let TotalLines(ByVal vData As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : TotalLines
   ' * Parameters       :
   ' *                    ByVal vData As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.TotalLines = 5
   mvarTotalLines = vData
End Property

Public Property Get TotalLines() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 16:37
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : TotalLines
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   TotalLines = mvarTotalLines
End Property

Public Property Get LastErrorMessage() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : LastErrorMessage
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   LastErrorMessage = msLastErrorMessage
End Property

Public Function Process(Project As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : Process
   ' * Parameters       :
   ' *                    project As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo errHandle

   Dim nFile            As Long

   Dim sKeyWords()      As String
   Dim rc               As Boolean
   Dim sLine            As String
   Dim genUts           As class_GeneralUtilities
   Dim dirName          As String
   Dim sFileToProcess   As String
   Dim clsclass_Class   As class_Class

   Dim sTWA()           As String

   Process = False
   If Dir(Project) = "" Then
      Process = False
      msLastErrorMessage = "Project File Does Not Exist"
      Exit Function
   End If

   nFile = FreeFile
   Set genUts = New class_GeneralUtilities
   dirName = genUts.GetDirectoryName(Project)
   mvarProjectPath = dirName
   Open Project For Input As #nFile
   Line Input #nFile, sLine
   RaiseEvent StatusMsg("Process Started")
   Do While Not EOF(nFile)
      If sLine <> "" Then
         Call Main_Module.Split(sLine, sKeyWords, "=")
      Else
         sKeyWords(0) = ""
      End If
      'keyWord = left$(sLine, InStr(1, sLine, "=") - 1)
      'sKeyValue = Mid$(sLine, InStr(1, sLine, "="), Len(sLine) - InStr(1, sLine, "="))
      Debug.Print sLine
      Select Case UCase$(sKeyWords(1))
         Case "TYPE"
            mvarProjectType = sKeyWords(2)

         Case "MAJORVER"
            msMajorVer = sKeyWords(2)

         Case "MINORVER"
            msMinorVer = sKeyWords(2)

         Case "REVISIONVER"
            msRevisionVer = sKeyWords(2)

         Case "TITLE"
            msTitle = sKeyWords(2)

         Case "EXENAME32"
            msExeName32 = sKeyWords(2)

         Case "DESCRIPTION"
            msDescription = sKeyWords(2)

         Case "VERSIONCOMMENTS"
            msVersionComments = sKeyWords(2)

         Case "VERSIONCOMPANYNAME"
            msVersionCompanyName = sKeyWords(2)

         Case "VERSIONLEGALTRADEMARKS"
            msVersionLegalTrademarks = sKeyWords(2)

         Case "VERSIONPRODUCTNAME"
            msVersionProductName = sKeyWords(2)

         Case "NAME"
            mvarProjectName = Replace(sKeyWords(2), """", "")

         Case "CLASS", "FORM", "MODULE", "USERDOCUMENT", "USERCONTROL", "RESFILE32"
            Set clsClass = New class_Class
            clsClass.ClassType = sKeyWords(1)
            clsClass.CountComments = Me.CountComments
            clsClass.CountSpace = Me.CountBlanks
            Call Main_Module.Split(Trim$(sKeyWords(2)), sTWA, ";")
            clsClass.ClassName = Trim$(sTWA(1))

            If UCase$(sKeyWords(1)) = "FORM" Then
               sFileToProcess = dirName & sKeyWords(2)

            ElseIf UCase$(sKeyWords(1)) = "USERDOCUMENT" Then
               sFileToProcess = dirName & sKeyWords(2)

            ElseIf UCase$(sKeyWords(1)) = "USERCONTROL" Then
               sFileToProcess = dirName & sKeyWords(2)

            ElseIf UCase$(sKeyWords(1)) = "RESFILE32" Then
               sFileToProcess = dirName & Replace(sKeyWords(2), """", "")

            Else
               sFileToProcess = dirName & Trim$(sTWA(2))
            End If

            If UCase$(sKeyWords(1)) <> "RESFILE32" Then rc = clsClass.Read(sFileToProcess)

            If Not rc Then
               Process = False
               msLastErrorMessage = clsClass.LastErrorMessage
               MsgBox clsClass.ClassName & " " & msLastErrorMessage, , "Failed"
               GoTo cleanUp
            End If
            Call Classes.Add(clsClass, clsClass.ClassName)
      End Select
      Line Input #nFile, sLine
   Loop

   For Each clsclass_Class In Me.Classes
      mvarTotalLines = mvarTotalLines + clsclass_Class.NumLines
   Next clsclass_Class
   Process = True

cleanUp:
   Close #nFile

   Exit Function
errHandle:
   Process = False
   msLastErrorMessage = err.Description

   Close #nFile

End Function

Public Property Let CountBlanks(ByVal vData As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : CountBlanks
   ' * Parameters       :
   ' *                    ByVal vData As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.CountBlanks = 5
   mvarCountBlanks = vData
End Property

Public Property Get CountBlanks() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 16:37
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : CountBlanks
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   CountBlanks = mvarCountBlanks
End Property

Public Property Let CountComments(ByVal vData As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
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
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 16:36
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : CountComments
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   CountComments = mvarCountComments
End Property

Public Property Get ProjectType() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : ProjectType
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   ProjectType = mvarProjectType
End Property

Public Property Let ProjectPath(ByVal vData As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : ProjectPath
   ' * Parameters       :
   ' *                    ByVal vData As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.ProjectPath = 5
   mvarProjectPath = vData
End Property

Public Property Get ProjectPath() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 16:36
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : ProjectPath
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   ProjectPath = mvarProjectPath
End Property

Public Property Let ProjectName(ByVal vData As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : ProjectName
   ' * Parameters       :
   ' *                    ByVal vData As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.ProjectName = 5
   mvarProjectName = vData
End Property

Public Property Get ProjectName() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 16:37
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : ProjectName
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   ProjectName = mvarProjectName
End Property

Private Sub clsClass_Process(StatusMsg As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : clsClass_Process
   ' * Parameters       :
   ' *                    StatusMsg As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   RaiseEvent StatusMsg(StatusMsg)
End Sub

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set Classes = New Collection

End Sub

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Project
   ' * Module Filename  : cntProject.cls
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set Classes = Nothing
End Sub
