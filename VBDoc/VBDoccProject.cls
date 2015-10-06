VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2008
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : cProject
' * Module Filename  : VBDoccProject.cls
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Purpose: This is the core object of this application. After the parsing it contains
'all object of the project and the it creates the documentation.
'Author:    Marco Pipino
Option Explicit

Public Classes          As OrderedColl          'Ordered collection of Classes modules
Public Modules          As OrderedColl          'Ordered collection of Standard modules
Public Forms            As OrderedColl          'Ordered collection of Form
Public UserControls     As OrderedColl          'Ordered collection of User Controls
Public ActiveXDesigners As OrderedColl          'Ordered collection of ActiveX Designers
Public MDIForms         As OrderedColl          'Ordered collection of MDIForms

Public Properties       As OrderedColl          'Ordered collection of Properties
Public Variables        As OrderedColl          'Ordered collection of Variables
Public Methods          As OrderedColl          'Ordered collection of Methods
Public Events           As OrderedColl          'Ordered collection of Events
Public Consts           As OrderedColl          'Ordered collection of Consts
Public Enums            As OrderedColl          'Ordered collection of Enumerations
Public Types            As OrderedColl          'Ordered collection of UDT's
Public References       As OrderedColl          'Ordered collection of References
Public Declarations     As OrderedColl          'Ordered collection of Declares
Public Implementations  As OrderedColl          'Ordered collection of Implements

Public VBI              As VBIDE.VBE            'The IDE of Visual Basic

Public Name             As String               'The name of the project
Public Description      As String               'The purpose of the project
Public Author           As String               'The Author of the project
Public DateBuild        As String               'The Date of creation
Public DateLastMod      As String               'The date of last modification
Public Version          As String               'The version
Public Example          As String               'Examples of project
Public SeeAlso          As String               'See Also
Public Screenshot       As String

'Templates
Private HTML_TEMPLATE_PROJECT_OVERVIEW As String
Private HTML_TEMPLATE_REFERENCE As String
Private HTML_TEMPLATE_REFERENCES As String
Private HTML_TEMPLATE_MODULE_OVERVIEW As String
Private HTML_TEMPLATE_METHOD As String
Private HTML_TEMPLATE_PROPERTY As String
Private HTML_TEMPLATE_VARIABLE As String
Private HTML_TEMPLATE_EVENT As String
Private HTML_TEMPLATE_ENUM As String
Private HTML_TEMPLATE_TYPE As String
Private HTML_TEMPLATE_PARAMETER As String
Private HTML_TEMPLATE_DECLARATION As String
Private HTML_TEMPLATE_IMPLEMENTS As String
Private HTML_TEMPLATE_SOURCECODE As String
Private HTML_TEMPLATE_CONSTS As String
Private HTML_TEMPLATE_ENUM_MEMBER As String
Private HTML_TEMPLATE_TYPE_MEMBER As String
Private HTML_TEMPLATE_PROPERTY_VALUE As String
Private HTML_TEMPLATE_VARIABLE_VALUE As String
Private HTML_TEMPLATE_CONST As String

'Purpose: Create all the ordered collection of the project
Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : Class_Initialize
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set Classes = New OrderedColl
   Set Modules = New OrderedColl
   Set Forms = New OrderedColl
   Set UserControls = New OrderedColl
   Set ActiveXDesigners = New OrderedColl
   Set MDIForms = New OrderedColl

   Set Properties = New OrderedColl
   Set Methods = New OrderedColl
   Set Events = New OrderedColl
   Set Consts = New OrderedColl
   Set Properties = New OrderedColl
   Set Variables = New OrderedColl
   Set Enums = New OrderedColl
   Set Types = New OrderedColl
   Set References = New OrderedColl
   Set Declarations = New OrderedColl
   Set Implementations = New OrderedColl
End Sub

'Purpose: The core of the application.<BR>
'Create a tree with each object (members, submemers) for each module
Public Sub BuildTree(myForm As Object)
Attribute BuildTree.VB_HelpID = 2019
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : BuildTree
   ' * Purpose          :
   ' * Parameters       :
   ' *                    myForm As Object
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim comp             As VBComponent
   Dim objClass         As cModule
   Dim objModule        As cModule
   Dim objForm          As cModule
   Dim objUserControl   As cModule
   Dim objActiveXDesigner As cModule
   Dim objMDIForm       As cModule

   'Set the default name for project
   Name = gProjectName

   For Each comp In VBI.ActiveVBProject.VBComponents
      frmProgress.MessageText = "Analyzing" & Chr$(13) & comp.Name

      Select Case comp.Type
         Case vbext_ct_ClassModule
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
               Set objClass = New cModule
               objClass.ComponentType = comp.Type
               objClass.Parse comp, Me
               Classes.Add objClass, objClass.Name
            End If
         Case vbext_ct_StdModule
            If myForm.trvComps.Nodes(comp.Name).Checked = True _
               And Not gPublicOnly Then
               Set objModule = New cModule
               objModule.ComponentType = comp.Type
               objModule.Parse comp, Me
               Modules.Add objModule, objModule.Name
            End If
         Case vbext_ct_VBForm
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
               Set objForm = New cModule
               objForm.ComponentType = comp.Type
               objForm.Parse comp, Me
               Forms.Add objForm, objForm.Name
            End If
         Case vbext_ct_UserControl
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
               Set objUserControl = New cModule
               objUserControl.ComponentType = comp.Type
               objUserControl.Parse comp, Me
               UserControls.Add objUserControl, objUserControl.Name
            End If
         Case vbext_ct_ActiveXDesigner
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
               Set objActiveXDesigner = New cModule
               objActiveXDesigner.ComponentType = comp.Type
               objActiveXDesigner.Parse comp, Me
               ActiveXDesigners.Add objActiveXDesigner, objActiveXDesigner.Name
            End If
         Case vbext_ct_VBMDIForm
            If myForm.trvComps.Nodes(comp.Name).Checked = True Then
               Set objMDIForm = New cModule
               objMDIForm.ComponentType = comp.Type
               objMDIForm.Parse comp, Me
               MDIForms.Add objMDIForm, objMDIForm.Name
            End If
      End Select
   Next
End Sub

'Purpose: Get Settings for application from registry. Read the template files
Public Sub GetSettings()
Attribute GetSettings.VB_HelpID = 2020
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : GetSettings
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error GoTo ErrorHandler

   gProjectFolder = left$(VBI.ActiveVBProject.FileName, InStrRev(VBI.ActiveVBProject.FileName, "\") - 1)

   gProjectName = right$(VBI.ActiveVBProject.FileName, Len(VBI.ActiveVBProject.FileName) - InStrRev(VBI.ActiveVBProject.FileName, "\"))
   gProjectName = left$(gProjectName, Len(gProjectName) - 4)
   gProjectFolder = left$(VBI.ActiveVBProject.FileName, InStrRev(VBI.ActiveVBProject.FileName, "\") - 1)

   MkDir gProjectFolder & "\VBDoc"

   gBLOCK_PURPOSE = GetSetting(gsREG_APP, "Blocks", "Purpose", "Purpose")
   gBLOCK_PROJECT = GetSetting(gsREG_APP, "Blocks", "Project Description", "Project Description")
   gBLOCK_AUTHOR = GetSetting(gsREG_APP, "Blocks", "Author", "Author")
   gBLOCK_DATE_CREATION = GetSetting(gsREG_APP, "Blocks", "Date_Creation", "Date")
   gBLOCK_DATE_LAST_MOD = GetSetting(gsREG_APP, "Blocks", "Date_Last_Mod", "Modified")
   gBLOCK_VERSION = GetSetting(gsREG_APP, "Blocks", "Version", "Version")
   gBLOCK_EXAMPLE = GetSetting(gsREG_APP, "Blocks", "Example", "Example")
   gBLOCK_SEEALSO = GetSetting(gsREG_APP, "Blocks", "SeeAlso", "See Also")
   gBLOCK_SCREEENSHOT = GetSetting(gsREG_APP, "Blocks", "Screenshot", "Screenshot")
   gBLOCK_CODE = GetSetting(gsREG_APP, "Blocks", "Code", "Code")
   gBLOCK_TEXT = GetSetting(gsREG_APP, "Blocks", "Text", "Text")
   gBLOCK_REMARKS = GetSetting(gsREG_APP, "Blocks", "Remarks", "Comments")
   gBLOCK_PARAMETER = GetSetting(gsREG_APP, "Blocks", "Parameter", "Parameters")
   gBLOCK_NO_COMMENT = GetSetting(gsREG_APP, "BLocks", "NoComment", "''")

   gBLOCK_WEBSITE = GetSetting(gsREG_APP, "BLocks", "Web Site", "Web Site")
   gBLOCK_EMAIL = GetSetting(gsREG_APP, "BLocks", "E-Mail", "E-Mail")
   gBLOCK_TIME = GetSetting(gsREG_APP, "BLocks", "Time", "Time")
   gBLOCK_TEL = GetSetting(gsREG_APP, "BLocks", "Telephone", "Telephone")
   gBLOCK_PROCEDURE_NAME = GetSetting(gsREG_APP, "BLocks", "Procedure Name", "Procedure Name")
   gBLOCK_MODULE_NAME = GetSetting(gsREG_APP, "BLocks", "Module Name", "Module Name")
   gBLOCK_MODULE_FILE = GetSetting(gsREG_APP, "BLocks", "Module Filename", "Module Filename")

   gHHCCompiler = GetSetting(gsREG_APP, "Compiler", "Path", "C:\Program Files\HTML Help Workshop\hhc.exe")

   HTML_TEMPLATE_MODULE_OVERVIEW = ReadTextFile(App.Path & "\VBDocTemplate\TemplateModule.htm")
   HTML_TEMPLATE_METHOD = ReadTextFile(App.Path & "\VBDocTemplate\TemplateMethod.htm")
   HTML_TEMPLATE_PROPERTY = ReadTextFile(App.Path & "\VBDocTemplate\TemplateProperty.htm")
   HTML_TEMPLATE_VARIABLE = ReadTextFile(App.Path & "\VBDocTemplate\TemplateVariable.htm")
   HTML_TEMPLATE_EVENT = ReadTextFile(App.Path & "\VBDocTemplate\TemplateEvent.htm")
   HTML_TEMPLATE_ENUM = ReadTextFile(App.Path & "\VBDocTemplate\TemplateEnum.htm")
   HTML_TEMPLATE_TYPE = ReadTextFile(App.Path & "\VBDocTemplate\TemplateType.htm")
   HTML_TEMPLATE_CONSTS = ReadTextFile(App.Path & "\VBDocTemplate\TemplateConsts.htm")
   HTML_TEMPLATE_SOURCECODE = ReadTextFile(App.Path & "\VBDocTemplate\TemplateSourceCode.htm")
   HTML_TEMPLATE_PROJECT_OVERVIEW = ReadTextFile(App.Path & "\VBDocTemplate\TemplateProject.htm")
   HTML_TEMPLATE_REFERENCES = ReadTextFile(App.Path & "\VBDocTemplate\TemplateReferences.htm")
   HTML_TEMPLATE_DECLARATION = ReadTextFile(App.Path & "\VBDocTemplate\TemplateDeclaration.htm")
   HTML_TEMPLATE_IMPLEMENTS = ReadTextFile(App.Path & "\VBDocTemplate\TemplateImplements.htm")

   gHHPTemplate = ReadTextFile(App.Path & "\VBDocTemplate\TemplateHHP.hhp")

   HTML_TEMPLATE_PARAMETER = "<TR VALIGN=top><TD width=27%><I>###Name###</I></TD><TD width=73%>###IsOptional### ###Description###</TD></TR>"
   HTML_TEMPLATE_ENUM_MEMBER = "<TR vAlign=top><TD>###Name###</TD><TD>###Value###</TD><TD width=73%>###Description###</TD></TR>"
   HTML_TEMPLATE_TYPE_MEMBER = "<TR vAlign=top><TD>###Name###</TD><TD>###Type###</TD><TD width=73%>###Description###</TD></TR>"
   HTML_TEMPLATE_PROPERTY_VALUE = "<TR VALIGN=top><TD width=27%><I>###TypeValue###</I></TD><TD width=73%>###ReturnType###.</TD></TR>"
   HTML_TEMPLATE_VARIABLE_VALUE = "<TR VALIGN=top><TD width=27%><I>###TypeValue###</I></TD><TD width=73%>###ReturnType###.</TD></TR>"
   HTML_TEMPLATE_CONST = "<TR vAlign=top><TD>###Scope###</TD><TD><b>###Name###</b></TD><TD>###Type###</TD><TD><b>###Value###</b></TD><TD>###Description###</TD></TR>"

   HTML_TEMPLATE_REFERENCE = "<TR vAlign=top><TD width=30%><B>###FileName###</B></TD><TD width=20%><B>###Version###</B></TD><TD width=*><B>###Description###</B></TD></TR>"

   Exit Sub
ErrorHandler:
   'if the error is Path/File access error i.e. The directory exist
   If err.number = 75 Then
      Resume Next
   End If
End Sub

'Puprpose: Create the HTML References for the project
Private Sub CreateHTMLReferences()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : CreateHTMLReferences
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim temp             As String
   Dim templibref       As String
   Dim refFile          As String

   Dim ref              As Reference
   For Each ref In VBI.ActiveVBProject.References
      temp = temp & HTML_TEMPLATE_REFERENCE
      temp = Replace(temp, "###Version###", CStr(ref.Major) & "." & CStr(ref.Minor))
      templibref = right$(ref.FullPath, Len(ref.FullPath) - InStrRev(ref.FullPath, "\"))
      templibref = right$(ref.FullPath, Len(ref.FullPath) - InStrRev(ref.FullPath, "\"))
      Do While InStr(1, templibref, ".") = 0
         templibref = right$(ref.FullPath, Len(ref.FullPath) - InStrRev(ref.FullPath, "\", Len(ref.FullPath) - Len(templibref) - 1))
      Loop
      templibref = Replace(templibref, ".oca", ".ocx")
      temp = Replace(temp, "###FileName###", templibref)
      temp = Replace(temp, "###Description###", RemoveCRLF(ref.Description))
   Next

   refFile = HTML_TEMPLATE_REFERENCES
   refFile = Replace(refFile, "###References###", temp)
   WriteTextFile gProjectFolder & "\VBDoc\References.htm", refFile
End Sub

'Purpose: Create the project overview
Private Sub CreateHTMLProjectOverview()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : CreateHTMLProjectOverview
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim strHTMLProject   As String
   'Name = gProjectName

   strHTMLProject = HTML_TEMPLATE_PROJECT_OVERVIEW
   strHTMLProject = Replace(strHTMLProject, "###Name###", Name)
   strHTMLProject = Replace(strHTMLProject, "###Description###", Description)
   strHTMLProject = Replace(strHTMLProject, "###Version###", IIf(Len(Version) > 0, "Version: " & Version, ""))
   strHTMLProject = Replace(strHTMLProject, "###DateRelease###", IIf(Len(DateLastMod) > 0, "Release: " & DateLastMod, ""))
   strHTMLProject = Replace(strHTMLProject, "###Author###", IIf(Len(Author) > 0, "Author: " & Author, ""))
   strHTMLProject = Replace(strHTMLProject, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<p><b> Example</b></p>" & ReplaceCRLF(Example), ""))
   strHTMLProject = Replace(strHTMLProject, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<p><b> See Also</b></p>" & ReplaceCRLF(SeeAlso), ""))

   Dim sTmpScreenshot   As String
   sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
   strHTMLProject = Replace(strHTMLProject, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

   WriteTextFile gProjectFolder & "\VBDoc\" & "Intro.htm", strHTMLProject
End Sub

'Purpose: Create all HTML files
Public Sub CreateHTMLFiles()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : CreateHTMLFiles
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim objCurr          As Object
   Dim objTOC           As cToc
   Dim strPath          As String
   Dim objMemb          As cMember

   strPath = gProjectFolder & "\VBDoc\"
   'Create the TOC object
   Set objTOC = New cToc
   Set objTOC.Project = Me
   objTOC.CreateTOC
   Set objTOC = Nothing

   frmProgress.MessageText = "Generating"

   CreateHTMLProjectOverview
   CreateHTMLReferences
   CreateHHP

   For Each objCurr In Classes
      frmProgress.MessageText = "Generating Class " & Chr$(13) & objCurr.Name
      objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
         strPath & objCurr.Name & "_Overview.htm", gPublicOnly
      'If gInsConsts Then
      objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
         strPath & objCurr.Name & "_Consts.htm"
      'End If
      If gSourceCode Then
         Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
      End If
   Next

   For Each objCurr In Modules
      frmProgress.MessageText = "Generating Module " & Chr$(13) & objCurr.Name
      objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
         strPath & objCurr.Name & "_Overview.htm", gPublicOnly
      'If gInsConsts Then
      objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
         strPath & objCurr.Name & "_Consts.htm"
      'End If
      If gSourceCode Then
         Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
      End If
   Next

   For Each objCurr In Forms
      frmProgress.MessageText = "Generating Form " & Chr$(13) & objCurr.Name
      objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
         strPath & objCurr.Name & "_Overview.htm", gPublicOnly
      'If gInsConsts Then
      objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
         strPath & objCurr.Name & "_Consts.htm"
      'End If
      If gSourceCode Then
         Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
      End If
   Next

   For Each objCurr In UserControls
      frmProgress.MessageText = "Generating UserControl " & Chr$(13) & objCurr.Name
      objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
         strPath & objCurr.Name & "_Overview.htm", gPublicOnly
      'If gInsConsts Then
      objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
         strPath & objCurr.Name & "_Consts.htm"
      'End If
      If gSourceCode Then
         Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
      End If
   Next

   For Each objCurr In ActiveXDesigners
      frmProgress.MessageText = "Generating ActiveXDesigner " & Chr$(13) & objCurr.Name
      objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
         strPath & objCurr.Name & "_Overview.htm", gPublicOnly
      'If gInsConsts Then
      objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
         strPath & objCurr.Name & "_Consts.htm"
      'End If
      If gSourceCode Then
         Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
      End If
   Next

   For Each objCurr In MDIForms
      frmProgress.MessageText = "Generating MDIForm " & Chr$(13) & objCurr.Name
      objCurr.CreateHTML Me, HTML_TEMPLATE_MODULE_OVERVIEW, _
         strPath & objCurr.Name & "_Overview.htm", gPublicOnly
      'If gInsConsts Then
      objCurr.CreateHTMLConsts Me, HTML_TEMPLATE_CONSTS, HTML_TEMPLATE_CONST, _
         strPath & objCurr.Name & "_Consts.htm"
      'End If
      If gSourceCode Then
         Call CreateHTMLCode(VBI.ActiveVBProject.VBComponents(objCurr.Name).CodeModule, objCurr.Name)
      End If
   Next

   For Each objMemb In Methods
      objMemb.CreateHTMLMethod HTML_TEMPLATE_METHOD, HTML_TEMPLATE_PARAMETER, _
         strPath & objMemb.RefModule & "_" & objMemb.Name & "_Method.htm", _
         gPublicOnly
   Next

   For Each objMemb In Properties
      objMemb.CreateHTMLProperty HTML_TEMPLATE_PROPERTY, HTML_TEMPLATE_PROPERTY_VALUE, _
         strPath & objMemb.RefModule & "_" & objMemb.Name & "_Property.htm", _
         gPublicOnly
   Next

   For Each objMemb In Events
      objMemb.CreateHTMLEvent HTML_TEMPLATE_EVENT, HTML_TEMPLATE_PARAMETER, _
         strPath & objMemb.RefModule & "_" & objMemb.Name & "_Event.htm"
   Next

   For Each objMemb In Variables
      objMemb.CreateHTMLVariable HTML_TEMPLATE_VARIABLE, HTML_TEMPLATE_VARIABLE_VALUE, _
         strPath & objMemb.RefModule & "_" & objMemb.Name & "_Property.htm", _
         gPublicOnly
   Next

   For Each objMemb In Enums
      objMemb.CreateHTMLEnum HTML_TEMPLATE_ENUM, HTML_TEMPLATE_ENUM_MEMBER, _
         strPath & objMemb.RefModule & "_" & objMemb.Name & "_Enum.htm", _
         gPublicOnly
   Next

   For Each objMemb In Types
      objMemb.CreateHTMLType HTML_TEMPLATE_TYPE, HTML_TEMPLATE_TYPE_MEMBER, _
         strPath & objMemb.RefModule & "_" & objMemb.Name & "_Type.htm", _
         gPublicOnly
   Next

   For Each objMemb In Declarations
      objMemb.CreateHTMLDeclaration HTML_TEMPLATE_DECLARATION, HTML_TEMPLATE_PARAMETER, _
         strPath & objMemb.RefModule & "_" & objMemb.Name & "_Declaration.htm", _
         gPublicOnly
   Next

   For Each objMemb In Implementations
      objMemb.CreateHTMLImplements HTML_TEMPLATE_IMPLEMENTS, _
         strPath & objMemb.RefModule & "_" & objMemb.Name & "_Implements.htm"
   Next

End Sub

'Purpose: Create the hhp file from the template. It insert all files HTML created
Private Sub CreateHHP()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : CreateHHP
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim strHHPTemplate   As String
   Dim FileList         As String
   Dim objCurr          As Object
   Dim strPath          As String

   FileList = ""
   strPath = gProjectFolder & "\VBDoc\"
   strHHPTemplate = gHHPTemplate

   strHHPTemplate = Replace(strHHPTemplate, "###Title###", gProjectName)
   strHHPTemplate = Replace(strHHPTemplate, "###ProjectTitle###", strPath & gProjectName)
   strHHPTemplate = Replace(strHHPTemplate, "###CompiledFile###", gProjectFolder & "\" & gProjectName)
   strHHPTemplate = Replace(strHHPTemplate, "###IntroFile###", gProjectFolder & "\VBDoc\Intro.htm")

   For Each objCurr In Methods
      FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Method.htm" & vbCrLf
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In Properties
      FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Property.htm" & vbCrLf
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In Events
      FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Event.htm" & vbCrLf
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In Variables
      FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Property.htm" & vbCrLf
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In Enums
      FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Enum.htm" & vbCrLf
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In Types
      FileList = FileList & strPath & objCurr.RefModule & "_" & objCurr.Name & "_Type.htm" & vbCrLf
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In Classes
      FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
      If Consts.Count(objCurr.Name) > 0 Then
         FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
      End If
      If gSourceCode Then
         FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
      End If
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In Modules
      FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
      If Consts.Count(objCurr.Name) > 0 Then
         FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
      End If
      If Not gSourceCode Then
         FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
      End If
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In Forms
      FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
      If Consts.Count(objCurr.Name) > 0 Then
         FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
      End If
      If Not gSourceCode Then
         FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
      End If
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In UserControls
      FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
      If Consts.Count(objCurr.Name) > 0 Then
         FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
      End If
      If Not gSourceCode Then
         FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
      End If
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In ActiveXDesigners
      FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
      If Consts.Count(objCurr.Name) > 0 Then
         FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
      End If
      If Not gSourceCode Then
         FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
      End If
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   For Each objCurr In MDIForms
      FileList = FileList & strPath & objCurr.Name & "_Overview.htm" & vbCrLf
      If Consts.Count(objCurr.Name) > 0 Then
         FileList = FileList & strPath & objCurr.Name & "_Consts.htm" & vbCrLf
      End If
      If Not gSourceCode Then
         FileList = FileList & strPath & objCurr.Name & "_Code.htm" & vbCrLf
      End If
      If Len(objCurr.Screenshot) > 2 Then
         FileList = FileList & strPath & objCurr.Screenshot & vbCrLf
      End If
   Next

   strHHPTemplate = Replace(strHHPTemplate, "###FileList###", FileList)
   WriteTextFile strPath & gProjectName & ".hhp", strHHPTemplate
End Sub

Public Sub SetHelpDescription()
   ' #VBIDEUtils#************************************************************
   ' * Author           : removed
   ' * Date             : 10/01/2002
   ' * Time             : 15:51
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : SetHelpDescription
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' *
   ' * Screenshot       :
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   On Error GoTo ERROR_SetHelpDescription

   Dim objCurr          As Object
   Dim comp             As VBComponent
   Dim sTmp             As String

   For Each objCurr In Methods
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Properties
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Variables
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Events
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Consts
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Types
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In References
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Declarations
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Implementations
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.RefModule = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.CodeModule.members(objCurr.Name).Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Classes
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.Name = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Forms
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.Name = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In MDIForms
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.Name = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In Modules
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.Name = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In UserControls
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.Name = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.Description = sTmp
            Exit For
         End If
      Next
   Next

   For Each objCurr In ActiveXDesigners
      For Each comp In VBI.ActiveVBProject.VBComponents
         If objCurr.Name = comp.Name Then
            sTmp = Trim$(objCurr.Description)
            sTmp = RemoveCRLF(sTmp)
            If Len(sTmp) > 2 Then comp.Description = sTmp
            Exit For
         End If
      Next
   Next

EXIT_SetHelpDescription:
   Exit Sub

ERROR_SetHelpDescription:
   Resume EXIT_SetHelpDescription

End Sub

'Purpose: Build a collection of Type Files user for recognize if the value of a
'member is as object or a not.
Public Function BuildTypesValue() As Boolean
Attribute BuildTypesValue.VB_HelpID = 2128
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : BuildTypesValue
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim objMember        As cMember
   On Error Resume Next
   Set gTypeValues = New Collection

   gTypeValues.Add "BOOLEAN", "BOOLEAN"
   gTypeValues.Add "BYTE", "BYTE"
   gTypeValues.Add "INTEGER", "INTEGER"
   gTypeValues.Add "LONG", "LONG"
   gTypeValues.Add "CURRENCY", "CURRENCY"
   gTypeValues.Add "SINGLE", "SINGLE"
   gTypeValues.Add "DOUBLE", "DOUBLE"
   gTypeValues.Add "DECIMAL", "DECIMAL"
   gTypeValues.Add "DATE", "DATE"
   gTypeValues.Add "STRING", "STRING"
   gTypeValues.Add "VARIANT", "VARIANT"
   gTypeValues.Add "ANY", "ANY"
   For Each objMember In Enums
      gTypeValues.Add UCase$(objMember.Name), UCase$(objMember.Name)
   Next
   For Each objMember In Types
      gTypeValues.Add UCase$(objMember.Name), UCase$(objMember.Name)
   Next
End Function

'Purpose: create the HTML for modules source code
Public Sub CreateHTMLCode(VBcode As CodeModule, ModuleName As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cProject
   ' * Module Filename  : VBDoccProject.cls
   ' * Procedure Name   : CreateHTMLCode
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    ModuleName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim i                As Integer
   Dim strCode          As String
   Dim strCodeModule    As String

   strCode = HTML_TEMPLATE_SOURCECODE
   strCodeModule = ""
   For i = 1 To VBcode.CountOfLines
      strCodeModule = strCodeModule & HTMLCodeLine(VBcode.Lines(i, 1)) & vbCrLf
   Next
   strCode = Replace(strCode, "###Name###", ModuleName)
   strCode = Replace(strCode, "###Code###", strCodeModule)
   WriteTextFile gProjectFolder & "\VBDoc\" & ModuleName & "_Code.htm", strCode
End Sub

