VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Analyze"
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
' * Date             : 23/05/2000
' * Time             : 10:53
' * Module Name      : class_Analyze
' * Module Filename  : MainAnalyze.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

' *** local variable(s) to hold property value(s)
Private mvarTotalLines  As Long 'local copy
Private mvarTotalClasses As Long 'local copy
Private mvarTotalForms  As Long 'local copy
Private mvarTotalMods   As Long 'local copy
Private mvarTotalUserControls As Long 'local copy
Private mvarTotalUserDocuments As Long 'local copy
Private mvarTotalFunctions As Long 'local copy
Private mvarTotalProperties As Long 'local copy
Private mvarTotalSubs   As Long 'local copy
Private mvarSmallestProject As String 'local copy
Private mvarLargestProject As String 'local copy
Private mvarLinesInSmallest As Long 'local copy
Private mvarLinesInLargest As Long 'local copy
Private mvarAverageProject As Double 'local copy
Private mvarAverageFunction As Double 'local copy
Private mvarAverageSub  As Double 'local copy
Private mvarAverageProperty As Double 'local copy
Private mvarLinesInFunctions As Long 'local copy
Private mvarLinesInSubroutines As Long 'local copy
Private mvarLinesInProperties As Long 'local copy
Private mvarVariables   As Long
Private mvarUnusedVariables As Long

Public Sub Analyze(Projects As Collection)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 23/05/2000
   ' * Time             : 10:53
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : Analyze
   ' * Parameters       :
   ' *                    Projects As Collection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim clsProject       As class_Project
   Dim clsClass         As class_Class
   Dim clsSub           As class_Sub

   If Projects Is Nothing Then Exit Sub

   For Each clsProject In Projects
      For Each clsClass In clsProject.Classes
         Select Case clsClass.ClassType
            Case "Form"
               mvarTotalForms = mvarTotalForms + 1
            Case "Class"
               mvarTotalClasses = mvarTotalClasses + 1
            Case "Module"
               mvarTotalMods = mvarTotalMods + 1
            Case "UserControl"
               mvarTotalUserControls = mvarTotalUserControls + 1
            Case "UserDocument"
               mvarTotalUserDocuments = mvarTotalUserDocuments + 1
         End Select

         For Each clsSub In clsClass.Subs
            Select Case clsSub.SubType
               Case "Function"
                  mvarLinesInFunctions = mvarLinesInFunctions + clsSub.NumLines
               Case "Sub"
                  mvarLinesInSubroutines = mvarLinesInSubroutines + clsSub.NumLines
               Case "Property"
                  mvarLinesInProperties = mvarLinesInProperties + clsSub.NumLines
            End Select
            mvarVariables = mvarVariables + clsSub.mcolVariable.Count
            mvarUnusedVariables = mvarUnusedVariables + clsSub.mcolUnusedVar.Count
         Next
         mvarTotalFunctions = mvarTotalFunctions + clsClass.NumFunctions
         mvarTotalSubs = mvarTotalSubs + clsClass.NumSubs
         mvarTotalProperties = mvarTotalProperties + clsClass.NumProperties

      Next

      If clsProject.TotalLines < mvarLinesInSmallest Or mvarLinesInSmallest = 0 Then
         mvarLinesInSmallest = clsProject.TotalLines
         mvarSmallestProject = clsProject.ProjectName
      End If
      If clsProject.TotalLines > mvarLinesInLargest Then
         mvarLargestProject = clsProject.ProjectName
         mvarLinesInLargest = clsProject.TotalLines
      End If
      mvarTotalLines = mvarTotalLines + clsProject.TotalLines
   Next

   If Projects.Count > 0 Then mvarAverageProject = CLng(mvarTotalLines / Projects.Count)
   If mvarTotalFunctions > 0 Then mvarAverageFunction = CLng(mvarLinesInFunctions / mvarTotalFunctions)
   If mvarTotalSubs > 0 Then mvarAverageSub = CLng(mvarLinesInSubroutines / mvarTotalSubs)
   If mvarTotalProperties > 0 Then mvarAverageProperty = CLng(mvarLinesInProperties / mvarTotalProperties)

End Sub

Public Property Get LinesInProperties() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : LinesInProperties
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   LinesInProperties = mvarLinesInProperties
End Property

Public Property Get LinesInSubroutines() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : LinesInSubroutines
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   LinesInSubroutines = mvarLinesInSubroutines
End Property

Public Property Get LinesInFunctions() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : LinesInFunctions
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   LinesInFunctions = mvarLinesInFunctions
End Property

Public Property Get AverageProperty() As Double
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : AverageProperty
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   AverageProperty = mvarAverageProperty
End Property

Public Property Get AverageSub() As Double
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : AverageSub
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   AverageSub = mvarAverageSub
End Property

Public Property Get AverageFunction() As Double
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : AverageFunction
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   AverageFunction = mvarAverageFunction
End Property

Public Property Get AverageProject() As Double
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : AverageProject
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   AverageProject = mvarAverageProject
End Property

Public Property Get LinesInLargest() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : LinesInLargest
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   LinesInLargest = mvarLinesInLargest
End Property

Public Property Get LinesInSmallest() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : LinesInSmallest
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   LinesInSmallest = mvarLinesInSmallest
End Property

Public Property Get LargestProject() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : LargestProject
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   LargestProject = mvarLargestProject
End Property

Public Property Get SmallestProject() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : SmallestProject
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   SmallestProject = mvarSmallestProject
End Property

Public Property Get TotalSubs() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalSubs
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   TotalSubs = mvarTotalSubs
End Property

Public Property Get TotalProperties() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalProperties
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   TotalProperties = mvarTotalProperties
End Property

Public Property Let TotalFunctions(ByVal vData As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalFunctions
   ' * Parameters       :
   ' *                    ByVal vData As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.TotalFunctions = 5
   mvarTotalFunctions = vData
End Property

Public Property Get TotalFunctions() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 23/05/2000
   ' * Time             : 10:53
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalFunctions
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   'used when retrieving value of a property, on the right side of an assignment.
   TotalFunctions = mvarTotalFunctions
End Property

Public Property Get TotalMods() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalMods
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   TotalMods = mvarTotalMods
End Property

Public Property Get TotalForms() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalForms
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   TotalForms = mvarTotalForms
End Property

Public Property Get TotalClasses() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalClasses
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   TotalClasses = mvarTotalClasses
End Property

Public Property Get TotalUserControls() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalUserControls
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   TotalUserControls = mvarTotalUserControls
End Property

Public Property Get TotalUserDocuments() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalUserDocuments
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when retrieving value of a property, on the right side of an assignment.
   TotalUserDocuments = mvarTotalUserDocuments
End Property

Public Property Get TotalLines() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
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

Public Property Get TotalVariables() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalVariables
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   TotalVariables = mvarVariables
End Property

Public Property Get TotalUnusedVariables() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Analyze
   ' * Module Filename  : MainAnalyze.cls
   ' * Procedure Name   : TotalUnusedVariables
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   TotalUnusedVariables = mvarUnusedVariables
End Property
