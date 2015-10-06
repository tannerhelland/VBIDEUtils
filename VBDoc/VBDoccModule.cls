VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : cModule
' * Module Filename  : VBDoccModule.cls
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Purpose: This class store the module object of each type, i.e. we can store
'   forms, module, class module etc.
'Author:    Marco Pipino
Option Explicit

Public Name             As String                  'The name of the module
Public Author           As String                  'The author of the module
Public DateCreation     As String                  'The date of creation of the module
Public DateLastMod      As String                  'The date of the lase modification of the module
Public Description      As String                  'The Purpose of the module
Public Remarks          As String                  'Remarks for this module
Public Example          As String                  'Example for this module
Public SeeAlso          As String
Public Screenshot       As String
Public WebSite          As String
Public EMail            As String
Public ModuleName       As String
Public ModuleFile       As String

Public ComponentType    As vbext_ComponentType 'Type of module

'Purpose: This enumeration serve for the parsing of the declaration module.<BR>
'   Because the VBIDE do not recognize the enumerations, the UDT's and the Implements
'   block.
Private Enum EnumBlockType
   Block_NONE
   Block_ENUM
   Block_TYPE
   Block_IMPLEMENTS
End Enum

'Purpose: This method create the HTML page for the Consts of the module
'Parameter: objProject the cProject object
'Parameter: TemplateConts the template for the HTML Table
'Parameter: TemplateConst the template for the HTML Row of Table
'Parameter: FileName Output file name
Public Sub CreateHTMLConsts(objProject As cProject, _
   TemplateConsts As String, _
   TemplateConst As String, _
   FileName As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : CreateHTMLConsts
   ' * Purpose          :
   ' * Parameters       :
   ' *                    objProject As cProject
   ' *                    TemplateConsts As String
   ' *                    TemplateConst As String
   ' *                    FileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim HTML_Consts      As String
   Dim objMemb          As cMember

   HTML_Consts = ""
   For Each objMemb In objProject.Consts
      'If PublicOnly we must write only the public consts
      If objMemb.RefModule = Name And _
         Not (objMemb.Scope <> vbext_Public And gPublicOnly) Then
         HTML_Consts = HTML_Consts & TemplateConst
         HTML_Consts = Replace(HTML_Consts, "###Scope###", ScopeToString(objMemb.Scope))
         HTML_Consts = Replace(HTML_Consts, "###Name###", objMemb.Name)
         HTML_Consts = Replace(HTML_Consts, "###Type###", objMemb.TypeValue)
         HTML_Consts = Replace(HTML_Consts, "###Value###", objMemb.Value)
         HTML_Consts = Replace(HTML_Consts, "###Description###", objMemb.Description)
      End If
   Next
   HTML_Consts = Replace(TemplateConsts, "###Members###", HTML_Consts)
   HTML_Consts = Replace(HTML_Consts, "###ObjectName###", Name)
   HTML_Consts = Replace(HTML_Consts, "###Remarks###", "")

   WriteTextFile FileName, HTML_Consts
End Sub

'Purpose: This method create all the HTML file in order to create the
'   CHM file.<BR>It read the templete file and replace the value for the
'   current object.
Public Sub CreateHTML(objProject As cProject, ByVal Template As String, _
   FileName As String, PublicOnly As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : CreateHTML
   ' * Purpose          :
   ' * Parameters       :
   ' *                    objProject As cProject
   ' *                    ByVal Template As String
   ' *                    FileName As String
   ' *                    PublicOnly As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim HTML_Properties  As String
   Dim HTML_Variables   As String
   Dim HTML_Methods     As String
   Dim HTML_Events      As String
   Dim HTML_Enumerations As String
   Dim HTML_UDTs        As String
   Dim HTML_Interfaces  As String
   Dim HTML_Consts      As String
   Dim HTML_Declarations As String
   Dim HTML_Implements  As String
   Dim HTML_Source_Code As String
   Dim i                As Integer

   Dim CurrObj          As Object
   Dim TempProperties   As OrderedColl

   HTML_Properties = ""
   HTML_Variables = ""
   HTML_Methods = ""
   HTML_Events = ""
   HTML_Enumerations = ""
   HTML_UDTs = ""
   HTML_Interfaces = ""
   HTML_Declarations = ""

   HTML_Methods = GetMenuHeader("Methods")
   i = 0
   For Each CurrObj In objProject.Methods
      If CurrObj.RefModule = Name Then
         i = i + 1
         HTML_Methods = HTML_Methods & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Method.htm" & """>" & vbCrLf
      End If
   Next
   HTML_Methods = HTML_Methods & GetMenuFooter(i)

   If Not gVarAsProperty Then
      HTML_Variables = GetMenuHeader("Variables")
      i = 0
      For Each CurrObj In objProject.Variables
         If CurrObj.RefModule = Name Then
            i = i + 1
            HTML_Variables = HTML_Variables & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Property.htm" & """>" & vbCrLf
         End If
      Next
      HTML_Variables = HTML_Variables & GetMenuFooter(i)

      HTML_Properties = GetMenuHeader("Properties")
      i = 0
      For Each CurrObj In objProject.Properties
         If CurrObj.RefModule = Name Then
            i = i + 1
            HTML_Properties = HTML_Properties & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Property.htm" & """>" & vbCrLf
         End If
      Next
      HTML_Properties = HTML_Properties & GetMenuFooter(i)
   Else
      Set TempProperties = New OrderedColl
      For Each CurrObj In objProject.Properties
         If CurrObj.RefModule = Name Then
            TempProperties.Add CurrObj, CurrObj.Name
         End If
      Next
      For Each CurrObj In objProject.Variables
         If CurrObj.RefModule = Name Then
            TempProperties.Add CurrObj, CurrObj.Name
         End If
      Next

      HTML_Properties = GetMenuHeader("Properties")
      i = 0
      For Each CurrObj In TempProperties
         i = i + 1
         HTML_Properties = HTML_Properties & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Property.htm" & """>" & vbCrLf
      Next
      HTML_Properties = HTML_Properties & GetMenuFooter(i)
      Set TempProperties = Nothing
   End If

   HTML_Events = GetMenuHeader("Events")
   i = 0
   For Each CurrObj In objProject.Events
      If CurrObj.RefModule = Name Then
         i = i + 1
         HTML_Events = HTML_Events & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Event.htm" & """>" & vbCrLf
      End If
   Next
   HTML_Events = HTML_Events & GetMenuFooter(i)

   HTML_Enumerations = GetMenuHeader("Enums")
   i = 0
   For Each CurrObj In objProject.Enums
      If CurrObj.RefModule = Name Then
         i = i + 1
         HTML_Enumerations = HTML_Enumerations & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Enum.htm" & """>" & vbCrLf
      End If
   Next
   HTML_Enumerations = HTML_Enumerations & GetMenuFooter(i)

   HTML_UDTs = GetMenuHeader("UDT's")
   i = 0
   For Each CurrObj In objProject.Types
      If CurrObj.RefModule = Name Then
         i = i + 1
         HTML_UDTs = HTML_UDTs & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Type.htm" & """>" & vbCrLf
      End If
   Next
   HTML_UDTs = HTML_UDTs & GetMenuFooter(i)

   If Not PublicOnly Then
      HTML_Declarations = GetMenuHeader("Declarations")
      i = 0
      For Each CurrObj In objProject.Declarations
         If CurrObj.RefModule = Name Then
            i = i + 1
            HTML_Declarations = HTML_Declarations & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Declaration.htm" & """>" & vbCrLf
         End If
      Next
      HTML_Declarations = HTML_Declarations & GetMenuFooter(i)

      HTML_Implements = GetMenuHeader("Implements")
      i = 0
      For Each CurrObj In objProject.Implementations
         If CurrObj.RefModule = Name Then
            i = i + 1
            HTML_Implements = HTML_Implements & "<PARAM NAME=""Item" & CStr(i) & """ VALUE=""" & CStr(CurrObj.Name) & ";" & Name & "_" & CurrObj.Name & "_Implements.htm" & """>" & vbCrLf
         End If
      Next
      HTML_Implements = HTML_Implements & GetMenuFooter(i)
   End If

   HTML_Consts = GetMenuHeader("Consts")
   i = 0
   For Each CurrObj In objProject.Consts
      If CurrObj.RefModule = Name Then i = i + 1
   Next
   If i > 0 Then
      HTML_Consts = HTML_Consts & "<PARAM NAME=""Item1"" VALUE=""" & Name & " Consts;" & Name & "_Consts.htm" & """>" & vbCrLf
   End If
   HTML_Consts = HTML_Consts & GetMenuFooter(i)

   If gSourceCode Then
      HTML_Source_Code = GetMenuHeader("Source Code")
      If gSourceCode Then
         HTML_Source_Code = HTML_Source_Code & "<PARAM NAME=""Item1"" VALUE=""" & Name & " Source Code;" & Name & "_Code.htm" & """>" & vbCrLf
         i = 1
      Else
         i = 0
      End If
      HTML_Source_Code = HTML_Source_Code & GetMenuFooter(i)
   End If

   If Not gVarAsProperty Then
      Template = Replace(Template, "###HTML_Variables###", HTML_Variables & "&nbsp;&nbsp; &nbsp;")
   Else
      Template = Replace(Template, "###HTML_Variables###", "")
      '###HTML_Variables###

   End If
   Template = Replace(Template, "###HTML_Methods###", HTML_Methods & "&nbsp;&nbsp; &nbsp;")
   Template = Replace(Template, "###HTML_Properties###", HTML_Properties & "&nbsp;&nbsp; &nbsp;")
   Template = Replace(Template, "###HTML_Events###", HTML_Events & "&nbsp;&nbsp; &nbsp;")
   Template = Replace(Template, "###HTML_Enumerations###", HTML_Enumerations & "&nbsp;&nbsp; &nbsp;")
   Template = Replace(Template, "###HTML_UDTs###", HTML_UDTs & "&nbsp;&nbsp; &nbsp;")
   Template = Replace(Template, "###HTML_Declarations###", HTML_Declarations & "&nbsp;&nbsp; &nbsp;")
   Template = Replace(Template, "###HTML_Implements###", HTML_Implements & "&nbsp;&nbsp; &nbsp;")
   Template = Replace(Template, "###HTML_Consts###", HTML_Consts & "&nbsp;&nbsp; &nbsp;")
   Template = Replace(Template, "###HTML_SourceCode###", HTML_Source_Code & "&nbsp;&nbsp; &nbsp;")

   Template = Replace(Template, "###Name###", Name)
   Template = Replace(Template, "###Type###", TypeToString)
   Template = Replace(Template, "###Description###", ReplaceCRLF(Description))

   If Len(Author) > 0 Then Author = "Author: " & Author & "<BR>"
   Template = Replace(Template, "###Author###", Author)

   If Len(DateCreation) > 0 Then DateCreation = "Creation " & DateCreation & "<BR>"
   Template = Replace(Template, "###Date###", DateCreation)

   If Len(DateLastMod) > 0 Then DateLastMod = "Modified: " & DateLastMod & "<BR>"
   Template = Replace(Template, "###DateLastMod###", DateLastMod)

   '   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<b>Remarks</b><br>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###Methods###", HTML_Methods)
   Template = Replace(Template, "###Properties###", HTML_Properties)
   Template = Replace(Template, "###Events###", HTML_Events)
   Template = Replace(Template, "###Enumerations###", HTML_Enumerations)
   Template = Replace(Template, "###UDTs###", HTML_UDTs)
   Template = Replace(Template, "###Declarations###", HTML_Declarations)
   Template = Replace(Template, "###Interfaces###", HTML_Interfaces)

   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<b>Remarks</b><br><br>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<b>Example</b><br><br>" & ReplaceCRLF(Example), ""))
   Template = Replace(Template, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<b>See Also</b><br><br>" & ReplaceCRLF(SeeAlso), ""))

   Dim sTmpScreenshot   As String
   sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
   Template = Replace(Template, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

   WriteTextFile FileName, Template

End Sub

'Purpose: This method parse all code for the module
Public Sub Parse(comp As VBComponent, objProject As cProject)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : Parse
   ' * Purpose          :
   ' * Parameters       :
   ' *                    comp As VBComponent
   ' *                    objProject As cProject
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   ParseModule comp, objProject
   ParseStatements comp.CodeModule, objProject, gPublicOnly
   ParseComments comp.CodeModule, objProject
End Sub

'Purpose: This methid parse the declaration section in order to generate the
'Enum, UDT's and Implements block.
Private Sub ParseStatements(VBcode As CodeModule, _
   objProject As cProject, PublicOnly As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : ParseStatements
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    objProject As cProject
   ' *                    PublicOnly As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error GoTo errore
   Dim Block            As EnumBlockType

   Dim Scope            As vbext_Scope
   Dim IsStatic         As Boolean

   Dim i                As Integer
   Dim intDecLines      As Integer
   Dim intBeginBlock    As Integer

   Dim temp             As String
   Dim objMember        As cMember

   intDecLines = VBcode.CountOfDeclarationLines

   Block = Block_NONE
   For i = 1 To intDecLines
      temp = VBcode.Lines(i, 1)
      Select Case Block
         Case Block_NONE
            Scope = ScopeOfMemb(temp, vbext_Public)
            IsStatic = IsStaticMemb(temp)

            If FirstLeftPart(temp, "Enum ", False, True) Then
               Block = Block_ENUM
               intBeginBlock = i
            ElseIf FirstLeftPart(temp, "Type ", False, True) Then
               Block = Block_TYPE
               intBeginBlock = i
            ElseIf FirstLeftPart(temp, "Implements ", False, True) Then
               Block = Block_IMPLEMENTS
               intBeginBlock = i
            End If

         Case Block_IMPLEMENTS
            If Not FirstLeftPart(temp, "'", True, False) Then
               Set objMember = New cMember
               objMember.RefModule = Name
               objMember.ParseImplements VBcode, Name, vbext_mt_Method, intBeginBlock
               objMember.MemberType = memb_type_Implements
               If Not gPublicOnly Then
                  objProject.Implementations.Add objMember, objMember.Name & "_" & Name, True
               Else
                  Set objMember = Nothing
               End If
               Block = Block_NONE
            End If

         Case Block_ENUM
            If FirstLeftPart(temp, "End", False, True) Then
               If FirstLeftPart(temp, "Enum", False, True) Then
                  Set objMember = New cMember
                  objMember.Scope = Scope
                  objMember.RefModule = Name
                  objMember.MemberType = memb_type_Enum

                  objMember.ParseEnum VBcode, Name, intBeginBlock, i - 1
                  If Not PublicOnly Or PublicOnly And objMember.Scope = vbext_Public Then
                     objProject.Enums.Add objMember, objMember.Name & "_" & Name, True
                  Else
                     Set objMember = Nothing
                  End If
                  Block = Block_NONE
               End If
            End If

         Case Block_TYPE
            If FirstLeftPart(temp, "End", False, True) Then
               If FirstLeftPart(temp, "Type", False, True) Then
                  Set objMember = New cMember
                  objMember.Scope = Scope
                  objMember.RefModule = Name
                  objMember.MemberType = memb_type_UDTs

                  objMember.ParseType VBcode, Name, intBeginBlock, i - 1
                  If Not PublicOnly Or PublicOnly And objMember.Scope = vbext_Public Then
                     objProject.Types.Add objMember, objMember.Name & "_" & Name, True
                  Else
                     Set objMember = Nothing
                  End If
                  Block = Block_NONE
               End If
            End If
      End Select
   Next
   Exit Sub
errore:
   Resume

End Sub

'Purpose: Execute the Parse of the Comments in order to create
'the description of the Module and the Project.
Private Sub ParseComments(VBcode As CodeModule, objProject As cProject)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : ParseComments
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    objProject As cProject
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim i                As Integer

   Dim strTemp          As String

   i = 1
   strTemp = VBcode.Lines(i, 1)
   Do While FirstLeftPart(strTemp, "'", True, True)
      If FirstLeftPart(strTemp, gBLOCK_PROJECT, False, True) Then
         If Len(Trim$(strTemp)) > 0 Then objProject.Name = Trim$(strTemp)
         'objProject.Description = strTemp & " " & vbCrLf & NextComments(VBcode, i)
         i = i + 1
         Do While left$(VBcode.Lines(i, 1), 1) = "'" And Not _
            FirstLeftPart(UCase$(VBcode.Lines(i, 1)), "'" & gBLOCK_PROJECT, False, False)
            strTemp = VBcode.Lines(i, 1)
            If FirstLeftPart(strTemp, gBLOCK_AUTHOR, False, True) Then
               objProject.Author = objProject.Author & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
            ElseIf FirstLeftPart(strTemp, gBLOCK_DATE_CREATION, False, True) Then
               objProject.DateBuild = objProject.DateBuild & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
            ElseIf FirstLeftPart(strTemp, gBLOCK_DATE_LAST_MOD, False, True) Then
               objProject.DateLastMod = objProject.DateLastMod & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
            ElseIf FirstLeftPart(strTemp, gBLOCK_VERSION, False, True) Then
               objProject.Version = objProject.Version & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
            ElseIf FirstLeftPart(strTemp, gBLOCK_EXAMPLE, False, True) Then
               objProject.Example = objProject.Example & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i, True)
            ElseIf FirstLeftPart(strTemp, gBLOCK_SEEALSO, False, True) Then
               objProject.SeeAlso = objProject.SeeAlso & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i, True)
            ElseIf FirstLeftPart(strTemp, gBLOCK_SCREEENSHOT, False, True) Then
               objProject.Screenshot = objProject.Screenshot & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i, True)
            Else
               LeftPart strTemp, "'", True, True
               objProject.Description = objProject.Description & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
               objProject.Description = RemoveCRLF(objProject.Description)
            End If
            i = i + 1
         Loop
         strTemp = ""
      End If
      i = i + 1
      strTemp = VBcode.Lines(i, 1)
      Do While FirstLeftPart(strTemp, "'", False, False)
         If FirstLeftPart(strTemp, gBLOCK_AUTHOR, False, True) Then
            Author = Author & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
         ElseIf FirstLeftPart(strTemp, gBLOCK_WEBSITE, False, True) Then
            WebSite = WebSite & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i) ' ***
         ElseIf FirstLeftPart(strTemp, gBLOCK_EMAIL, False, True) Then
            EMail = EMail & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i) ' ***
         ElseIf FirstLeftPart(strTemp, gBLOCK_MODULE_NAME, False, True) Then
            ModuleName = ModuleName & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)  ' ***
         ElseIf FirstLeftPart(strTemp, gBLOCK_MODULE_FILE, False, True) Then
            ModuleName = ModuleName & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i) ' ***
         ElseIf FirstLeftPart(strTemp, gBLOCK_TIME, False, True) Then
            ' *** Ignore
         ElseIf FirstLeftPart(strTemp, gBLOCK_TEL, False, True) Then
            ' *** Ignore

         ElseIf FirstLeftPart(strTemp, gBLOCK_DATE_CREATION, False, True) Then
            DateCreation = DateCreation & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
         ElseIf FirstLeftPart(strTemp, gBLOCK_DATE_LAST_MOD, False, True) Then
            DateLastMod = DateLastMod & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
         ElseIf FirstLeftPart(strTemp, gBLOCK_EXAMPLE, False, True) Then
            Example = Example & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i, True)
            If Len(ReplaceCRLF(Example)) > 0 Then
               Example = "<font face=""Courier New"">" & Example & "</font>"
            End If
         ElseIf FirstLeftPart(strTemp, gBLOCK_SEEALSO, False, True) Then
            SeeAlso = SeeAlso & strTemp & vbCrLf & NextComments(VBcode, i, True)
         ElseIf FirstLeftPart(strTemp, gBLOCK_SCREEENSHOT, False, True) Then
            Screenshot = Screenshot & strTemp & vbCrLf & NextComments(VBcode, i, True)
         ElseIf FirstLeftPart(strTemp, gBLOCK_PURPOSE, False, True) Then
            If strTemp <> vbNullString Then Description = Description & strTemp & " " & vbCrLf
            Description = Description & NextComments(VBcode, i)
            Description = Replace(Description, vbCrLf, " ")
            Description = Replace(Description, "  ", " ")
            Description = Replace(Description, "  ", " ")
            Description = Replace(Description, "  ", " ")
            Description = Trim$(Replace(Description, "  ", " "))
            Description = RemoveCRLF(Description)
         ElseIf FirstLeftPart(strTemp, gBLOCK_REMARKS, False, True) Then
            If strTemp <> vbNullString Then Remarks = Remarks & strTemp & " " & vbCrLf
            Remarks = Remarks & NextComments(VBcode, i)
            Remarks = Replace(Remarks, vbCrLf, " ")
            Remarks = Replace(Remarks, "  ", " ")
            Remarks = Replace(Remarks, "  ", " ")
            Remarks = Replace(Remarks, "  ", " ")
            Remarks = Trim$(Replace(Remarks, "  ", " "))
            Remarks = RemoveCRLF(Remarks)
         ElseIf FirstLeftPart(strTemp, gBLOCK_NO_COMMENT, False, True) Then
            'There are no commands for no comments
         Else
            LeftPart strTemp, "'", True, True
            Description = Description & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, i)
         End If
         i = i + 1
         strTemp = VBcode.Lines(i, 1)
      Loop
   Loop

   If Len(Description) = 0 Then Description = Remarks

End Sub

'Purpose: Parse the module in order to recognize all the note block for
'the VBIDE object
Private Sub ParseModule(comp As VBComponent, objProject As cProject)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : ParseModule
   ' * Purpose          :
   ' * Parameters       :
   ' *                    comp As VBComponent
   ' *                    objProject As cProject
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim memb             As Member
   Dim objMember        As cMember

   Name = comp.Name
   If right$(Name, 1) = "%" Then
      Name = left$(Name, Len(Name) - 1)
   End If
   ComponentType = comp.Type

   For Each memb In comp.CodeModule.members
      If memb.Scope = vbext_Public Or (memb.Scope <> vbext_Public And Not gPublicOnly) Then
         Select Case memb.Type
            Case vbext_mt_Method
               'This section is valid for methods and declare
               Set objMember = New cMember
               objMember.RefModule = Name
               objMember.ParseMethod comp.CodeModule, memb
               If objMember.MemberType = memb_type_Declare Then
                  objProject.Declarations.Add objMember, objMember.Name & "_" & Name, True
               Else
                  objProject.Methods.Add objMember, objMember.Name & "_" & Name, True
               End If

            Case vbext_mt_Property
               Set objMember = New cMember
               objMember.Name = memb.Name
               objMember.IsStatic = memb.Static
               objMember.RefModule = Name
               objMember.Scope = memb.Scope
               objMember.ParseProperty comp.CodeModule, memb
               objMember.MemberType = memb_type_Property
               objProject.Properties.Add objMember, objMember.Name & "_" & Name, True

            Case vbext_mt_Event
               Set objMember = New cMember
               objMember.Name = memb.Name
               objMember.RefModule = Name
               objMember.Scope = memb.Scope
               objMember.MemberType = memb_type_Event

               objMember.ParseEvent comp.CodeModule, memb
               objProject.Events.Add objMember, objMember.Name & "_" & Name, True

            Case vbext_mt_Const
               'If gInsConsts Then
               Set objMember = New cMember
               objMember.Name = memb.Name
               objMember.Scope = memb.Scope
               objMember.RefModule = Name
               objMember.MemberType = memb_type_Const
               objMember.ParseConst comp.CodeModule, memb
               objProject.Consts.Add objMember, objMember.Name & "_" & Name, True
               'End If

            Case vbext_mt_Variable
               Set objMember = New cMember
               objMember.Name = memb.Name
               objMember.Scope = memb.Scope
               objMember.MemberType = memb_type_Variable
               objMember.RefModule = Name
               objMember.ParseVariable comp.CodeModule, memb
               objProject.Variables.Add objMember, objMember.Name & "_" & Name, True
         End Select

         If right$(objMember.Name, 1) = "%" Then
            objMember.Name = left$(objMember.Name, Len(objMember.Name) - 1)
         End If
      End If
   Next
End Sub

'Purpose: Generate a string relative to the value of the Type of component
Private Function TypeToString() As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : TypeToString
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Select Case ComponentType
      Case vbext_ct_ClassModule
         TypeToString = "Object"
      Case vbext_ct_StdModule
         TypeToString = "Module"
      Case vbext_ct_VBForm
         TypeToString = "Form"
      Case vbext_ct_UserControl
         TypeToString = "User Control"
      Case vbext_ct_ActiveXDesigner
         TypeToString = "ActiveX Designer"
      Case vbext_ct_VBMDIForm
         TypeToString = "MDI Form"
      Case Else
         TypeToString = "Module"
   End Select
End Function

'Purpose: Get the menu haeder for the HTML file
Private Function GetMenuHeader(strType As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : GetMenuHeader
   ' * Purpose          :
   ' * Parameters       :
   ' *                    strType As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim strTemp          As String
   strTemp = "<OBJECT TYPE=""application/x-oleobject"" CLASSID=""clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11"" ID=""M"" VIEWASTEXT>" & vbCrLf
   strTemp = strTemp & "<PARAM NAME=""Command"" value="""
   strTemp = strTemp & "Related Topics, MENU"
   strTemp = strTemp & """ > " & vbCrLf
   strTemp = strTemp & "<PARAM NAME=""Command"" value=""ALink"">"
   'strTemp = strTemp & "<PARAM NAME=""Font"" value="",,,,underline"">"
   strTemp = strTemp & "<PARAM NAME=""Text"" VALUE=""Text:" & strType & """>" & vbCrLf
   strTemp = strTemp & "<PARAM NAME=""Flags"" VALUE=""1"">" & vbCrLf
   GetMenuHeader = strTemp
End Function

'Purpose: Get the menu footer for the HTML file
'Parameter: intItems if 0 the color of the menu is gray else it is blue
Private Function GetMenuFooter(intItems As Integer) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cModule
   ' * Module Filename  : VBDoccModule.cls
   ' * Procedure Name   : GetMenuFooter
   ' * Purpose          :
   ' * Parameters       :
   ' *                    intItems As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim strTemp          As String
   strTemp = "<PARAM NAME=""Font"" VALUE=""Helvetica,13,," & IIf(intItems > 0, "#0000ff", "#808080") & "," & IIf(intItems > 0, "UNDERLINE", "PLAIN") & """>" & vbCrLf
   strTemp = strTemp & "</OBJECT>" & vbCrLf
   GetMenuFooter = strTemp
End Function

