VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : cToc
' * Module Filename  : VBDoccToc.cls
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Purpose: Used only for create the TOC of the CHM file.
'Author:    Marco Pipino
Option Explicit

Private objProject      As cProject  'Private variable for storing the cProject object
Private strText         As String       'Text of the TOC

'Purpose: Create a single folder for each module
Private Sub CreateModules(comp As cModule)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : CreateModules
   ' * Purpose          :
   ' * Parameters       :
   ' *                    comp As cModule
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   'Create the TOC for each Module

   Dim strPath          As String
   Dim strModName       As String

   Dim TempProperties   As OrderedColl
   Dim TempObj          As Object

   strPath = gProjectFolder & "\VBDoc\"

   strModName = comp.Name

   If comp.ComponentType = vbext_ct_StdModule Then
      HHCAddBookInit strModName & " Module"
   ElseIf comp.ComponentType = vbext_ct_ClassModule Then
      HHCAddBookInit strModName & " Object"
   ElseIf comp.ComponentType = vbext_ct_VBForm Then
      HHCAddBookInit strModName & " Forms"
   ElseIf comp.ComponentType = vbext_ct_UserControl Then
      HHCAddBookInit strModName & " User Controls"
   ElseIf comp.ComponentType = vbext_ct_ActiveXDesigner Then
      HHCAddBookInit strModName & " ActiveX Designer"
   ElseIf comp.ComponentType = vbext_ct_VBMDIForm Then
      HHCAddBookInit strModName & " MDI Form"
   End If

   HHCAddTopic "Overview", strPath & strModName & "_Overview.htm"
   If objProject.Consts.Count(strModName) > 0 Then
      HHCAddTopic "Consts", strPath & strModName & "_Consts.htm"
   End If

   If objProject.Declarations.Count(strModName) > 0 Then
      HHCAddBookInit "Declarations"
      For Each TempObj In objProject.Declarations
         If TempObj.RefModule = strModName Then
            HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Declaration.htm"
         End If
      Next
      HHCAddBookEnd
   End If

   If objProject.Implementations.Count(strModName) > 0 Then
      HHCAddBookInit "Implements"
      For Each TempObj In objProject.Implementations
         If TempObj.RefModule = strModName Then
            HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Implements.htm"
         End If
      Next
      HHCAddBookEnd
   End If

   If objProject.Methods.Count(strModName) > 0 Then
      HHCAddBookInit "Methods"
      For Each TempObj In objProject.Methods
         If TempObj.RefModule = strModName Then
            HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Method.htm"
         End If
      Next
      HHCAddBookEnd
   End If

   If gVarAsProperty Then
      Set TempProperties = New OrderedColl
      For Each TempObj In objProject.Properties
         If TempObj.RefModule = strModName Then
            TempProperties.Add TempObj, TempObj.Name
         End If
      Next
      For Each TempObj In objProject.Variables
         If TempObj.RefModule = strModName Then
            TempProperties.Add TempObj, TempObj.Name
         End If
      Next

      If TempProperties.Count > 0 Then
         HHCAddBookInit "Properties"
         For Each TempObj In TempProperties
            HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Property.htm"
         Next
         Set TempProperties = Nothing
         HHCAddBookEnd
      End If
   Else
      If objProject.Properties.Count(strModName) > 0 Then
         HHCAddBookInit "Properties"
         For Each TempObj In objProject.Properties
            If TempObj.RefModule = strModName Then
               HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Property.htm"
            End If
         Next
         HHCAddBookEnd
      End If
      If objProject.Variables.Count(strModName) > 0 Then
         HHCAddBookInit "Variables"
         For Each TempObj In objProject.Variables
            If TempObj.RefModule = strModName Then
               HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Property.htm"
            End If
         Next
         HHCAddBookEnd
      End If
   End If

   If objProject.Events.Count(strModName) > 0 Then
      HHCAddBookInit "Events"
      For Each TempObj In objProject.Events
         If TempObj.RefModule = strModName Then
            HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Event.htm"
         End If
      Next
      HHCAddBookEnd
   End If

   If objProject.Enums.Count(strModName) > 0 Then
      HHCAddBookInit "Enums"
      For Each TempObj In objProject.Enums
         If TempObj.RefModule = strModName Then
            HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Enum.htm"
         End If
      Next
      HHCAddBookEnd
   End If

   If objProject.Types.Count(strModName) > 0 Then
      HHCAddBookInit "UDT's"
      For Each TempObj In objProject.Types
         If TempObj.RefModule = strModName Then
            HHCAddTopic TempObj.Name, strPath & strModName & "_" & TempObj.Name & "_Type.htm"
         End If
      Next
      HHCAddBookEnd
   End If

   HHCAddBookEnd  'Chiusura del Module
End Sub

'Purpose: Add to strText the End of the Book
Private Sub HHCAddBookEnd()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : HHCAddBookEnd
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   strText = strText & "</ul>" & vbCrLf
End Sub

'Purpose: Add to strText the Begin of the Book
Private Sub HHCAddBookInit(BookValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : HHCAddBookInit
   ' * Purpose          :
   ' * Parameters       :
   ' *                    BookValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   strText = strText & "<li><object type=""text/sitemap"">" & vbCrLf
   strText = strText & "<param name=""Name"" value=""" + BookValue + """>" & vbCrLf
   strText = strText & "</object>" & vbCrLf
   strText = strText & "<ul>" & vbCrLf
End Sub

'Purpose: Add a Topic to the HHC File
Private Sub HHCAddTopic(topicValue As String, topicFile As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : HHCAddTopic
   ' * Purpose          :
   ' * Parameters       :
   ' *                    topicValue As String
   ' *                    topicFile As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   strText = strText & "<li><object type=""text/sitemap"">" & vbCrLf
   strText = strText & "<param name=""Name"" value=""" + topicValue + """>" & vbCrLf
   strText = strText & "<param name=""Local"" value=""" + topicFile + """>" & vbCrLf
   strText = strText & "</object>" & vbCrLf
End Sub

'Purpose: Add the tail of the Book
Private Sub HHCCoda()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : HHCCoda
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   strText = strText & "</ul>" & vbCrLf
   strText = strText & "</html>" & vbCrLf
End Sub

'Purpose: Add the Head of the Book
Private Sub HHCTestata()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : HHCTestata
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   strText = strText & "<html>" & vbCrLf
   strText = strText & "<!-- Sitemap 1.0 -->" & vbCrLf
   strText = strText & "<object type=""text/site properties"">" & vbCrLf
   strText = strText & "</object>" & vbCrLf
   strText = strText & "<ul>" & vbCrLf & vbCrLf
End Sub

'Purpose: Create the complete TOC
Public Sub CreateTOC()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : CreateTOC
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim TempProperties   As OrderedColl
   Dim TempObj          As Object

   Dim strPath          As String

   strPath = gProjectFolder & "\VBDoc\"

   HHCTestata
   HHCAddTopic "Intro", strPath & "Intro.htm"
   HHCAddTopic "References", strPath & "References.htm"

   If objProject.Classes.Count > 0 Then
      HHCAddBookInit "Objects"
      For Each TempObj In objProject.Classes
         CreateModules TempObj
      Next
      HHCAddBookEnd
   End If

   If objProject.Modules.Count > 0 Then
      HHCAddBookInit "Modules"
      For Each TempObj In objProject.Modules
         CreateModules TempObj
      Next
      HHCAddBookEnd
   End If

   If objProject.Forms.Count > 0 Then
      HHCAddBookInit "Forms"
      For Each TempObj In objProject.Forms
         CreateModules TempObj
      Next
      HHCAddBookEnd
   End If

   If objProject.UserControls.Count > 0 Then
      HHCAddBookInit "User Controls"
      For Each TempObj In objProject.UserControls
         CreateModules TempObj
      Next
      HHCAddBookEnd
   End If

   If objProject.ActiveXDesigners.Count > 0 Then
      HHCAddBookInit "ActiveX Designers"
      For Each TempObj In objProject.ActiveXDesigners
         CreateModules TempObj
      Next
      HHCAddBookEnd
   End If

   If objProject.MDIForms.Count > 0 Then
      HHCAddBookInit "MDI Forms"
      For Each TempObj In objProject.ActiveXDesigners
         CreateModules TempObj
      Next
      HHCAddBookEnd
   End If

   If objProject.Declarations.Count > 0 Then
      HHCAddBookInit "Declarations"
      For Each TempObj In objProject.Declarations
         HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Declaration.htm"
      Next
      HHCAddBookEnd
   End If

   If objProject.Implementations.Count > 0 Then
      HHCAddBookInit "Implements"
      For Each TempObj In objProject.Implementations
         HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Implements.htm"
      Next
      HHCAddBookEnd
   End If

   If objProject.Methods.Count > 0 Then
      HHCAddBookInit "Methods"
      For Each TempObj In objProject.Methods
         HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Method.htm"
      Next
      HHCAddBookEnd
   End If

   If gVarAsProperty Then
      Set TempProperties = New OrderedColl
      For Each TempObj In objProject.Properties
         TempProperties.Add TempObj, TempObj.Name & "_" & TempObj.RefModule, True
      Next
      For Each TempObj In objProject.Variables
         TempProperties.Add TempObj, TempObj.Name & "_" & TempObj.RefModule, True
      Next
      If TempProperties.Count > 0 Then
         HHCAddBookInit "Properties"
         For Each TempObj In TempProperties
            HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Property.htm"
         Next
         Set TempProperties = Nothing
         HHCAddBookEnd
      End If
   Else
      If objProject.Properties.Count > 0 Then
         HHCAddBookInit "Properties"
         For Each TempObj In objProject.Properties
            HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Property.htm"
         Next
         HHCAddBookEnd
      End If

      If objProject.Variables.Count > 0 Then
         HHCAddBookInit "Variables"
         For Each TempObj In objProject.Variables
            HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Property.htm"
         Next
         HHCAddBookEnd
      End If

   End If

   If objProject.Events.Count > 0 Then
      HHCAddBookInit "Events"
      For Each TempObj In objProject.Events
         HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Event.htm"
      Next
      HHCAddBookEnd
   End If

   If objProject.Enums.Count > 0 Then
      HHCAddBookInit "Enums"
      For Each TempObj In objProject.Enums
         HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Enum.htm"
      Next
      HHCAddBookEnd
   End If
   If objProject.Types.Count > 0 Then
      HHCAddBookInit "UDT's"
      For Each TempObj In objProject.Types
         HHCAddTopic TempObj.Name & " (" & TempObj.RefModule & ")", strPath & TempObj.RefModule & "_" & TempObj.Name & "_Type.htm"
      Next
      HHCAddBookEnd
   End If
   HHCCoda

   WriteTextFile strPath & gProjectName & ".hhc", strText

End Sub

'Purpose: Set the project property
Public Property Set Project(ByVal vProject As cProject)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : Project
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal vProject As cProject
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set objProject = vProject
End Property

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cToc
   ' * Module Filename  : VBDoccToc.cls
   ' * Procedure Name   : Class_Initialize
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   strText = ""
End Sub
