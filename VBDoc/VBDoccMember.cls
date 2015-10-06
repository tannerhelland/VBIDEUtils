VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : cMember
' * Module Filename  : VBDoccMember.cls
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Purpose: This class contain information about the each members of the module
'Author:    Marco Pipino

Option Explicit

Public RefModule        As String               'The module of the member
Public MemberType       As EnumMemberType       'Indicate if the member is a method, a property,
'an event, etc..
Public Name             As String               'The name of the meber

Public Author           As String               'The author of the member
Public DateCreation     As String               'The date of creation of the member
Public DateLastMod      As String               'The date of the lase modification of the member
Public Description      As String               'Purpose of the member
Public Remarks          As String               'Remorks for the member
Public Example          As String               'One or more examples for the member
Public SeeAlso          As String
Public Screenshot       As String
Public WebSite          As String
Public EMail            As String
Public ModuleName       As String
Public ModuleFile       As String

Public Scope            As vbext_Scope          'The scope of the member (i.e. public,
'private or friend)
Public IsStatic         As Boolean              'Indicate if the member is static
Public TypeValue        As String               'The type of member

Public Value            As String               'The value of member valid only for Consts
Public Library          As String               'The library of the declare
Public Alias            As String               'The Alias of the declare
Public SubMembers       As Collection           'This is a <B>NOT</B> ordered collection

'because the order it's is important for
'parameters, enum members and type members
Public PropertyType     As EnumPropertyType 'Only for property, indicate if it's Read
'Only, Write Only or Read and Write

'Purpose: Enumerate the type of member
Public Enum EnumMemberType
   memb_type_Method
   memb_type_Property
   memb_type_Const
   memb_type_Declare
   memb_type_Enum
   memb_type_Event
   memb_type_Implements
   memb_type_UDTs
   memb_type_Variable
End Enum

'Purpose: Enumerate the type of property
Public Enum EnumPropertyType
   Prop_Type_ReadOnly
   Prop_Type_ReadWrite
   Prop_Type_WriteOnly
End Enum

'Purpose: Class initialization, create only the SubMembers Colletion <B>
'not</B> ordered.
Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : Class_Initialize
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set SubMembers = New Collection
End Sub

'Purpose: Destroy the SubMembers collection.
Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : Class_Terminate
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set SubMembers = Nothing
End Sub

'Purpose: Parse the method and declare definitions and the its comments i.e.<BR><BR>
'for Declarations Syntax 1:<BR>
'[Public | Private] Declare Sub name Lib "libname" [Alias "aliasname"] [([arglist])]<BR>
'Syntax 2:<BR>
'[Public | Private] Declare Function name Lib "libname" [Alias "aliasname"] [([arglist])] [As type]<BR><BR>
'for Functions:<BR>
'[Public | Private | Friend] [Static] Function name [(arglist)] [As type]<BR><BR>
'for Subs:<BR>
'[Private | Public | Friend] [Static] Sub name [(arglist)]
Public Sub ParseMethod(VBcode As CodeModule, VBMemb As Member)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseMethod
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    VBMemb As Member
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error Resume Next
   Dim MethodDef        As String

   Dim intParentesi     As Integer
   Dim i                As Integer
   Dim strTemp          As String
   Dim strComment       As String
   Dim intCodeLocation  As Integer

   MethodDef = GetDef(VBcode, VBMemb, intCodeLocation, strComment)

   ScopeOfMemb MethodDef
   Scope = VBMemb.Scope

   IsStatic = IsStaticMemb(MethodDef)
   If FirstLeftPart(MethodDef, "Declare", False, True) Then
      MemberType = memb_type_Declare
   Else
      MemberType = memb_type_Method
   End If

   If FirstLeftPart(MethodDef, "Sub", False, True) Then
      TypeValue = "None"
   ElseIf FirstLeftPart(MethodDef, "Function", False, True) Then
      TypeValue = "Variant"
   End If

   If MemberType = memb_type_Declare Then
      'if it's a declaration
      strTemp = LeftPart(MethodDef, "(", True, True)
      Name = LeftPart(strTemp, "Lib", False, True, True)
      Library = LeftPart(strTemp, "Alias", False, True, True)

      Library = left$(Library, Len(Library) - 1)
      Library = right$(Library, Len(Library) - 1)
      Alias = strTemp
      Alias = left$(Alias, Len(Alias) - 1)
      Alias = right$(Alias, Len(Alias) - 1)

   Else
      'It's a method
      Name = LeftPart(MethodDef, "(", True, True)
   End If

   intParentesi = 0
   strTemp = ""
   For i = 1 To Len(MethodDef)
      If intParentesi = 0 And Mid$(MethodDef, i, 1) = ")" Then
         Exit For
      ElseIf Mid$(MethodDef, i, 1) = ")" Then
         intParentesi = intParentesi - 1
      ElseIf Mid$(MethodDef, i, 1) = "(" Then
         intParentesi = intParentesi + 1
      Else
         strTemp = strTemp & Mid$(MethodDef, i, 1)
      End If
   Next

   'Generate Parameter
   SetParameter (Trim$(strTemp))

   MethodDef = Trim$(right$(MethodDef, Len(MethodDef) - Len(strTemp) - 1))
   If UCase$(TypeValue) = "VARIANT" And _
      FirstLeftPart(MethodDef, "AS", False, True) Then
      TypeValue = Trim$(MethodDef)
   End If

   ParseDescriptions VBcode, GetFirstComment(VBcode, VBMemb.Name, VBMemb.Type, intCodeLocation)

   Description = Description & " " & strComment

End Sub

'Purpose: Parse descriptions for all members
Private Sub ParseDescriptions(VBcode As CodeModule, CodeLocation As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseDescriptions
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    CodeLocation As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next

   Dim strTemp          As String
   Dim tempParamDesc    As String
   Dim tempParamName    As String

   strTemp = VBcode.Lines(CodeLocation, 1)
   '   Do While FirstLeftPart(strTemp, "'", True, True)
   '      If FirstLeftPart(strTemp, gBLOCK_REMARKS, False, True) Then
   '         Remarks = Remarks & " " & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation)
   '      ElseIf FirstLeftPart(strTemp, gBLOCK_NO_COMMENT, False, True) Then
   '         'There are no commands for the no comments tag
   '      ElseIf FirstLeftPart(strTemp, gBLOCK_EXAMPLE, False, True) Then
   '         Example = Example & " " & strTemp & vbCrLf & NextComments(VBcode, CodeLocation, True)
   '      ElseIf FirstLeftPart(strTemp, gBLOCK_PARAMETER, False, True) Then
   '         tempParamDesc = strTemp & " " & vbCrLf & NextComments(VBcode, CodeLocation)
   '         tempParamName = LeftPart(tempParamDesc, " ", True, True, True)
   '         SubMembers(tempParamName).Description = tempParamDesc
   '      ElseIf FirstLeftPart(strTemp, gBLOCK_PURPOSE, False, True) Then
   '         Description = Description & " " & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation)
   '      End If
   '      CodeLocation = CodeLocation + 1
   '      strTemp = VBcode.Lines(CodeLocation, 1)
   '   Loop

   Do While FirstLeftPart(strTemp, "'", True, True)
      If FirstLeftPart(strTemp, gBLOCK_AUTHOR, False, True) Then
         Author = Author & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation)
      ElseIf FirstLeftPart(strTemp, gBLOCK_WEBSITE, False, True) Then
         WebSite = WebSite & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation) ' ***
      ElseIf FirstLeftPart(strTemp, gBLOCK_EMAIL, False, True) Then
         EMail = EMail & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation) ' ***
      ElseIf FirstLeftPart(strTemp, gBLOCK_MODULE_NAME, False, True) Then
         ModuleName = ModuleName & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation)  ' ***
      ElseIf FirstLeftPart(strTemp, gBLOCK_MODULE_FILE, False, True) Then
         ModuleName = ModuleName & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation) ' ***
      ElseIf FirstLeftPart(strTemp, gBLOCK_TIME, False, True) Then
         ' *** Ignore
      ElseIf FirstLeftPart(strTemp, gBLOCK_TEL, False, True) Then
         ' *** Ignore

      ElseIf FirstLeftPart(strTemp, gBLOCK_PARAMETER, False, True) Then
         tempParamDesc = Trim$(strTemp & " " & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextCommentsParameters(VBcode, CodeLocation))
         Do While (Len(tempParamDesc) > 0) And (left$(tempParamDesc, 10) <> "' ********")
            DoEvents
            tempParamName = LeftPart(tempParamDesc, " ", True, True, True)
            If LCase$(tempParamName) = "optional" Then
               tempParamName = LeftPart(tempParamDesc, " ", True, True, True)
            End If
            SubMembers(tempParamName).Description = tempParamDesc
            strTemp = VBcode.Lines(CodeLocation, 1)
            tempParamDesc = Trim$(NextCommentsParameters(VBcode, CodeLocation))
         Loop

      ElseIf FirstLeftPart(strTemp, gBLOCK_DATE_CREATION, False, True) Then
         DateCreation = DateCreation & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation)
      ElseIf FirstLeftPart(strTemp, gBLOCK_DATE_LAST_MOD, False, True) Then
         DateLastMod = DateLastMod & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation)
      ElseIf FirstLeftPart(strTemp, gBLOCK_EXAMPLE, False, True) Then
         Example = Example & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation, True)
         If Len(ReplaceCRLF(Example)) > 0 Then
            Example = "<font face=""Courier New"">" & Example & "</font>"
         End If
      ElseIf FirstLeftPart(strTemp, gBLOCK_SEEALSO, False, True) Then
         SeeAlso = Trim$(SeeAlso & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation, True))

      ElseIf FirstLeftPart(strTemp, gBLOCK_SCREEENSHOT, False, True) Then
         Screenshot = Trim$(Screenshot & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation, True))

      ElseIf FirstLeftPart(strTemp, gBLOCK_PURPOSE, False, True) Then
         If strTemp <> vbNullString Then Description = Description & strTemp & " " & vbCrLf
         Description = Description & NextComments(VBcode, CodeLocation)
         Description = Replace(Description, vbCrLf, " ")
         Description = Replace(Description, "  ", " ")
         Description = Replace(Description, "  ", " ")
         Description = Replace(Description, "  ", " ")
         Description = Trim$(Replace(Description, "  ", " "))
         Description = RemoveCRLF(Description)
      ElseIf FirstLeftPart(strTemp, gBLOCK_REMARKS, False, True) Then
         If strTemp <> vbNullString Then Remarks = Remarks & strTemp & " " & vbCrLf
         Remarks = Remarks & NextComments(VBcode, CodeLocation)
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
         Description = Description & IIf(Len(strTemp) > 0, strTemp & vbCrLf, "") & NextComments(VBcode, CodeLocation)
      End If
      CodeLocation = CodeLocation + 1
      strTemp = VBcode.Lines(CodeLocation, 1)
   Loop

   If Len(Description) = 0 Then Description = Remarks

End Sub

'Purpose: Parse a string of type <BR>
'[Optional] [ByVal | ByRef] [ParamArray] varname[( )] [As type]
'Remarks: The ParamArrya keyword generate error!!!!!
Private Sub SetParameter(argList As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : SetParameter
   ' * Purpose          :
   ' * Parameters       :
   ' *                    argList As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim temp             As String
   Dim cpar             As cSubMember
   Do While Len(argList) > 0
      temp = LeftPart(argList, ",", True, True, True)
      Set cpar = New cSubMember
      cpar.ParseParameter (temp)
      SubMembers.Add cpar, cpar.Name
   Loop
End Sub

'Purpose: This function used by the Events,Enum,UDTs and Declaration Classes
'   retrieve the first line of comment before the definition
'Remarks: It's important a blank line between a declaration af a member
'   and the next initial comment for another member
'Parameter: VBcode      The Code Module for the member
'Parameter: sMethodName The name of the member
'Parameter: intLine     The Line number of the definition
Private Function GetFirstComment(VBcode As CodeModule, _
   sMethodName As String, nType As vbext_MemberType, intLine As Integer) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : GetFirstComment
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    sMethodName As String
   ' *                    nType As vbext_MemberType
   ' *                    intLine As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error GoTo GetFirstComment_Error

   Dim bFound           As Boolean

   bFound = False
   GetFirstComment = intLine - 1
   Do While FirstLeftPart(VBcode.Lines(GetFirstComment - 1, VBcode.ProcCountLines(sMethodName, vbext_pk_Proc)), "'", True, False)
      GetFirstComment = GetFirstComment - 1
      bFound = True
   Loop

   If Not bFound Then
      GetFirstComment = intLine + 2
      Do While right$(VBcode.Lines(GetFirstComment - 1, 1), 1) = "_"
         GetFirstComment = GetFirstComment + 1
      Loop
      GetFirstComment = GetFirstComment + 1
      Do While FirstLeftPart(VBcode.Lines(GetFirstComment - 1, VBcode.ProcCountLines(sMethodName, vbext_pk_Proc)), "'", True, False)
         GetFirstComment = GetFirstComment - 1
         bFound = True
      Loop
   End If

   Exit Function
GetFirstComment_Error:

End Function

'Purpose: this complex function count the numbeer of line that terminate with
'the Underscore char.<BR>
'The CodeLocation property of the Member object of VBIDE fails
'when encounters this type of lines
Private Function PrevUnderScore(VBModule As CodeModule, _
   CodeLocation As Integer) As Integer
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : PrevUnderScore
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBModule As CodeModule
   ' *                    CodeLocation As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim i                As Integer
   Dim intTemp          As Integer
   Dim strTemp          As String
   Dim intRowToScroll   As Integer
   Dim intTempUnderScore As Integer

   PrevUnderScore = 0
   intTemp = 0
   intRowToScroll = CodeLocation
   Do
      intTempUnderScore = 0
      For i = intTemp + 1 To intTemp + intRowToScroll
         strTemp = Trim$(VBModule.Lines(i, 1))
         If right$(strTemp, 1) = "_" Then
            intTempUnderScore = intTempUnderScore + 1
         End If
      Next
      PrevUnderScore = PrevUnderScore + intTempUnderScore
      intTemp = intTemp + intRowToScroll
      intRowToScroll = intTempUnderScore
   Loop While intTempUnderScore > 0
End Function

'Purpose: Parse the events declarations and its descriptions<BR>
'[<B>Public</B>] <B>Event</B> procedurename [<B>(</B>arglist<B>)</B>]
Public Sub ParseEvent(VBcode As CodeModule, VBMemb As Member)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseEvent
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    VBMemb As Member
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim EventDef         As String

   Dim intParentesi     As Integer
   Dim i                As Integer
   Dim strTemp          As String
   Dim strComment       As String
   Dim intCodeLocation  As Integer

   EventDef = GetDef(VBcode, VBMemb, intCodeLocation, strComment)
   ScopeOfMemb EventDef
   Call LeftPart(EventDef, "(", True, True, True)

   intParentesi = 0
   strTemp = ""
   For i = 1 To Len(EventDef)
      If intParentesi = 0 And Mid$(EventDef, i, 1) = ")" Then
         Exit For
      ElseIf Mid$(EventDef, i, 1) = ")" Then
         intParentesi = intParentesi - 1
      ElseIf Mid$(EventDef, i, 1) = "(" Then
         intParentesi = intParentesi + 1
      End If
      strTemp = strTemp & Mid$(EventDef, i, 1)
   Next
   SetParameter (Trim$(strTemp))

   LeftPart EventDef, "'", True, True, True

   ParseDescriptions VBcode, GetFirstComment(VBcode, VBMemb.Name, VBMemb.Type, intCodeLocation)
   Description = Description & " " & EventDef
   Description = Trim$(Description)
   Description = RemoveCRLF(Description)
   If left$(Description, 3) = "***" Then Description = Trim$(Mid$(Description, 4))
End Sub

'Purpose: Get the definition of a member even if on more code lines.
'Remarks: Returs the correct CodeLocation for the declaration of member in the
'byref parameter <I>CodeLocation</I>.<BR>
'It use the PrevUnderScore function for obtain the correct CodeLocation.<BR>
'The <I>strComment</I> byref parameter returns the comment after the definition.
Private Function GetDef(VBcode As CodeModule, _
   VBMemb As Member, CodeLocation As Integer, _
   Optional strCommentAfterDef As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : GetDef
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    VBMemb As Member
   ' *                    CodeLocation As Integer
   ' *                    Optional strCommentAfterDef As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim temp             As String
   Dim i                As Integer
   Dim CodLoc           As Integer

   CodLoc = VBMemb.CodeLocation + PrevUnderScore(VBcode, VBMemb.CodeLocation - 1)
   i = 0

   temp = RemoveDoubleBlank(VBcode.Lines(CodLoc, 1))
   Do While right$(temp, 1) = "_"
      i = i + 1
      temp = RemoveDoubleBlank(left$(temp, Len(temp) - 1)) _
         & " " & RemoveDoubleBlank(VBcode.Lines(CodLoc + i, 1))
   Loop

   GetDef = LeftPart(temp, "'", True, True, True)
   strCommentAfterDef = temp
   CodeLocation = CodLoc '+ i
End Function

'Purpose: Parse the variable code and its comments.<BR>

Public Sub ParseVariable(VBcode As CodeModule, VBMemb As Member)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseVariable
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    VBMemb As Member
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim VariableDef      As String
   Dim i                As Integer
   Dim strComment       As String
   Dim intCodeLocation  As Integer

   VariableDef = GetDef(VBcode, VBMemb, intCodeLocation, strComment)

   Call LeftPart(VariableDef, " AS ", False, True, True)
   If Len(VariableDef) > 0 Then
      TypeValue = VariableDef
   Else
      TypeValue = "Variant"
   End If
   ParseDescriptions VBcode, GetFirstComment(VBcode, VBMemb.Name, VBMemb.Type, intCodeLocation)
   Description = Description & " " & strComment & " " & vbCrLf & NextComments(VBcode, intCodeLocation)
   Description = Trim$(Description)
   Description = RemoveCRLF(Description)
   If left$(Description, 3) = "***" Then Description = Trim$(Mid$(Description, 4))
End Sub

'Purpose: Parsing of the Const. When we define two or more consts on the same line this method
'FAILS.
Public Sub ParseConst(VBcode As CodeModule, VBMemb As Member)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseConst
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    VBMemb As Member
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim ConstDef         As String
   Dim i                As Integer
   Dim strComment       As String
   Dim intCodeLocation  As Integer

   ConstDef = GetDef(VBcode, VBMemb, intCodeLocation, strComment)

   ScopeOfMemb ConstDef

   FirstLeftPart ConstDef, "CONST", False, True
   FirstLeftPart ConstDef, VBMemb.Name, False, True

   If FirstLeftPart(ConstDef, "AS", False, True) Then
      TypeValue = LeftPart(ConstDef, "=", False, True, True)
   Else
      TypeValue = "Variant"
      Call LeftPart(ConstDef, "=", False, True, True)
   End If

   Value = LeftPart(ConstDef, "'", True, True)

   If TypeValue = "Variant" Then
      If IsNumeric(Value) Then
         TypeValue = "Long"
      ElseIf UCase$(Value) = "TRUE" Or UCase$(Value) = "FALSE" Then
         TypeValue = "Boolean"
      ElseIf IsDate(Value) Then
         TypeValue = "Date"
      Else
         TypeValue = "String"
      End If
   End If

   'Description = ConstDef
   ParseDescriptions VBcode, GetFirstComment(VBcode, VBMemb.Name, VBMemb.Type, intCodeLocation)
   Description = Description & " " & strComment
   Description = Trim$(Description)
   Description = RemoveCRLF(Description)
   If left$(Description, 3) = "***" Then Description = Trim$(Mid$(Description, 4))
End Sub

'Purpose: Parse the property block. <B> The code location is on the Get
'block if exist, Let o Set else.</B>
Public Sub ParseProperty(VBcode As CodeModule, VBMemb As Member)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseProperty
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    VBMemb As Member
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   'Viene definito se la proprietà è di sola lettura, sola scrittura
   'o lettura e scrittura.

   Dim strTemp          As String
   Dim strDefProp       As String
   Dim intCodeLocation  As Integer
   Dim strComment       As String

   'Recognize if the property is ReadOnly, WriteOnly o Read-Write
   If GetPropertyFirstLine(VBcode, VBMemb, vbext_pk_Get) > 0 Then
      If GetPropertyFirstLine(VBcode, VBMemb, vbext_pk_Let) > 0 Or _
         GetPropertyFirstLine(VBcode, VBMemb, vbext_pk_Set) > 0 Then
         PropertyType = Prop_Type_ReadWrite
      Else
         PropertyType = Prop_Type_ReadOnly
      End If
   Else
      PropertyType = Prop_Type_WriteOnly
   End If

   strDefProp = GetDef(VBcode, VBMemb, intCodeLocation, strComment)

   strTemp = LeftPart(strDefProp, " AS ", False, True, True)

   TypeValue = LeftPart(strDefProp, ")", True, True, True)
   If TypeValue = "" Then TypeValue = "Variant"

   ParseDescriptions VBcode, GetPropertyFirstLine(VBcode, VBMemb, vbext_pk_Get)
   Description = Description & " " & strComment
   Description = Trim$(Description)
   Description = RemoveCRLF(Description)
   If left$(Description, 3) = "***" Then Description = Trim$(Mid$(Description, 4))

End Sub

'Purpose: Return the first code line of the block of property
'0 if the block does not exist
'Parameter: VBCode The code module
'Parameter: VBMemb The member property object
'Parameter: PropType The type of property
Private Function GetPropertyFirstLine(VBcode As CodeModule, _
   VBMemb As Member, PropType As vbext_ProcKind) As Long
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : GetPropertyFirstLine
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    VBMemb As Member
   ' *                    PropType As vbext_ProcKind
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   On Error Resume Next
   GetPropertyFirstLine = 0
   GetPropertyFirstLine = VBcode.ProcStartLine(VBMemb.Name, PropType)

   Do While right$(VBcode.Lines(GetPropertyFirstLine - 1, 1), 1) = "_"
      GetPropertyFirstLine = GetPropertyFirstLine + 1
   Loop
   GetPropertyFirstLine = GetPropertyFirstLine + 2
   If FirstLeftPart(VBcode.Lines(GetPropertyFirstLine, VBcode.ProcCountLines(VBMemb.Name, PropType)), "'", True, False) Then

   End If

End Function

'Purpose: Parse Implements syntax
'Parameter: VBCode The code module
Public Sub ParseImplements(VBcode As CodeModule, sName As String, nType As vbext_MemberType, BeginLine As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseImplements
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    sName As String
   ' *                    nType As vbext_MemberType
   ' *                    BeginLine As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim i                As Integer
   Dim strTemp          As String

   strTemp = VBcode.Lines(BeginLine, 1)

   ScopeOfMemb strTemp

   FirstLeftPart strTemp, "Implements", False, True

   Name = LeftPart(strTemp, "'")
   ParseDescriptions VBcode, GetFirstComment(VBcode, sName, nType, BeginLine)
   Description = Trim$(Description)
   Description = RemoveCRLF(Description)
   If left$(Description, 3) = "***" Then Description = Trim$(Mid$(Description, 4))
End Sub

'Purpose: Parse an enumeration
'Parameter: VBCode The code module
Public Sub ParseEnum(VBcode As CodeModule, _
   sName As String, ByVal BeginLine As Integer, _
   EndLine As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseEnum
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    sName As String
   ' *                    ByVal BeginLine As Integer
   ' *                    EndLine As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   On Error Resume Next
   Dim i                As Integer
   Dim strTemp          As String
   Dim lngCurrMemberValue As Long

   strTemp = VBcode.Lines(BeginLine, 1)

   ScopeOfMemb strTemp
   FirstLeftPart strTemp, "Enum", False, True

   Name = LeftPart(strTemp, "'", True, True, True)
   ParseDescriptions VBcode, GetFirstComment(VBcode, sName, vbext_mt_Method, BeginLine) + 2

   lngCurrMemberValue = -1
   For i = BeginLine + 1 To EndLine
      strTemp = VBcode.Lines(i, 1)
      If Not FirstLeftPart(strTemp, "'", True, False) Then
         ParseEnumMember VBcode, strTemp, lngCurrMemberValue, i
      End If
   Next
   Description = Trim$(Description)
   Description = RemoveCRLF(Description)
   If left$(Description, 3) = "***" Then Description = Trim$(Mid$(Description, 4))

End Sub

'Purpose: Parse Members of an enumeration
'Parameter: VBCode The code module
Private Function ParseEnumMember(VBcode As CodeModule, strLine As String, _
   currValue As Long, CurrLine As Integer) As Long
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseEnumMember
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    strLine As String
   ' *                    currValue As Long
   ' *                    CurrLine As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim objEnumMember    As cSubMember
   Dim defLine          As String

   defLine = LeftPart(strLine, "'", True, True, True)
   Set objEnumMember = New cSubMember
   objEnumMember.Name = LeftPart(defLine, "=", False, True, True)
   If Len(defLine) > 0 Then
      objEnumMember.Value = defLine
      currValue = cLngP(defLine)
   Else
      currValue = currValue + 1
      objEnumMember.Value = currValue
   End If
   objEnumMember.Description = strLine & " " & vbCrLf & NextComments(VBcode, CurrLine)
   objEnumMember.Description = Trim$(objEnumMember.Description)
   objEnumMember.Description = RemoveCRLF(objEnumMember.Description)
   If left$(objEnumMember.Description, 3) = "***" Then objEnumMember.Description = Trim$(Mid$(objEnumMember.Description, 4))
   SubMembers.Add objEnumMember, objEnumMember.Name
End Function

'Purpose: Parse the UDT's block in a module
'Parameter: VBCode The code module
'Parameter: BeginLine The first line of the block
'Parameter: EndLine The last line of the block
Public Sub ParseType(VBcode As CodeModule, _
   sName As String, BeginLine As Integer, _
   EndLine As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseType
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    sName As String
   ' *                    BeginLine As Integer
   ' *                    EndLine As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim i                As Integer
   Dim strTemp          As String

   strTemp = VBcode.Lines(BeginLine, 1)

   ScopeOfMemb strTemp
   FirstLeftPart strTemp, "Type", False, True

   Name = LeftPart(strTemp, "'", True, True, True)
   ParseDescriptions VBcode, GetFirstComment(VBcode, sName, vbext_mt_Method, BeginLine)

   For i = BeginLine + 1 To EndLine
      strTemp = VBcode.Lines(i, 1)
      If Not FirstLeftPart(strTemp, "'", True, False) Then
         ParseTypeMember VBcode, strTemp, i
      End If
   Next
   Description = Trim$(Description)
   Description = RemoveCRLF(Description)
   If left$(Description, 3) = "***" Then Description = Trim$(Mid$(Description, 4))

End Sub

'Purpose: Parse the type submember
Public Sub ParseTypeMember(VBcode As CodeModule, _
   strLine As String, _
   CurrLine As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ParseTypeMember
   ' * Purpose          :
   ' * Parameters       :
   ' *                    VBcode As CodeModule
   ' *                    strLine As String
   ' *                    CurrLine As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim objTypeMember    As cSubMember
   Dim strComment       As String

   strComment = strLine
   strLine = LeftPart(strComment, "'", False, True)
   Set objTypeMember = New cSubMember
   objTypeMember.TypeValue = "Variant"
   objTypeMember.Name = LeftPart(strLine, " As ", False, True, True)
   If Len(strLine) > 0 Then objTypeMember.TypeValue = strLine
   objTypeMember.Description = strComment & " " & vbCrLf & NextComments(VBcode, CurrLine)
   objTypeMember.Description = Trim$(objTypeMember.Description)
   objTypeMember.Description = RemoveCRLF(objTypeMember.Description)
   If left$(objTypeMember.Description, 3) = "***" Then objTypeMember.Description = Trim$(Mid$(objTypeMember.Description, 4))
   SubMembers.Add objTypeMember, objTypeMember.Name
End Sub

'Purpose: Create HTML for Declarations
Public Sub CreateHTMLDeclaration(ByVal Template As String, _
   TemplateParamenter As String, _
   FileName As String, _
   PublicOnly As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : CreateHTMLDeclaration
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Template As String
   ' *                    TemplateParamenter As String
   ' *                    FileName As String
   ' *                    PublicOnly As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim CurrParam        As cSubMember
   Dim Syntax           As String
   Dim params           As String

   If PublicOnly Then
      Exit Sub
   End If
   Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
   Template = Replace(Template, "###Name###", Name)
   Description = Trim$(Description)
   If Len(Description) = 0 Then Description = Name
   Description = RemoveCRLF(Description)
   Template = Replace(Template, "###Description###", ReplaceCRLF(Description))
   Template = Replace(Template, "###Library###", Library)

   If Len(Alias) > 0 Then Alias = "<p><B>Alias</B></p><p>" & Alias & "</P>"
   Template = Replace(Template, "###Alias###", Alias)

   Template = Replace(Template, "###ReturnTypeMethod###", IIf(isAnObject(TypeValue), "As <b>" & TypeValue & "</b> " & "Object", IIf(TypeValue = "None", "<b>None</b>", "As <b>" & TypeValue & "</b> " & "Value")))
   Template = Replace(Template, "###ReturnType###", IIf(isAnObject(TypeValue), "As <b>" & TypeValue & "</b> " & "Object", IIf(TypeValue = "None", "<b>None</b>", "As <b>" & TypeValue & "</b> " & "Value")))
   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###ObjectName###", RefModule)
   'Syntax
   Syntax = "<b>" & Name & "</b> "
   If TypeValue <> "None" Then
      Syntax = Syntax & "<b>(</b>"
   End If
   params = ""
   Syntax = Syntax & "<i>"
   For Each CurrParam In SubMembers
      Syntax = Syntax & " " & IIf(CurrParam.IsOptional, "[", "") & CurrParam.Name & IIf(CurrParam.IsOptional, "]", "") & ","
      params = params & TemplateParamenter
      params = Replace(params, "###Name###", CurrParam.Name)
      params = Replace(params, "###Type###", IIf(isAnObject(CurrParam.TypeValue), "As <b>" & CurrParam.TypeValue & "</b> Object.", "As <b>" & CurrParam.TypeValue & "</b> Value."))
      params = Replace(params, "###Description###", Trim$(Replace(CurrParam.Description, "As ", "")))
      params = Replace(params, "###IsOptional###", IIf(CurrParam.IsOptional, "Optional.", "Required."))
   Next
   If right$(Syntax, 1) = "," Then Syntax = left$(Syntax, Len(Syntax) - 1)
   Syntax = Syntax & "</i>"
   If TypeValue <> "None" Then
      Syntax = Syntax & "<b>)</b>"
   End If
   Template = Replace(Template, "###Syntax###", Syntax)
   Template = Replace(Template, "###Parameters###", params)

   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<p><b> Example</b></p>" & ReplaceCRLF(Example), ""))
   Template = Replace(Template, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<p><b> See Also</b></p>" & ReplaceCRLF(SeeAlso), ""))
   Dim sTmpScreenshot   As String
   sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
   Template = Replace(Template, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

   WriteTextFile FileName, Template
End Sub

'Purpose: create HTML For Enumerations
Public Sub CreateHTMLEnum(ByVal Template As String, TemplateMember As String, _
   FileName As String, PublicOnly As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : CreateHTMLEnum
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Template As String
   ' *                    TemplateMember As String
   ' *                    FileName As String
   ' *                    PublicOnly As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim CurrMember       As cSubMember
   Dim members          As String

   If PublicOnly Then
      Template = Replace(Template, "###Scope###", "")
   Else
      Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
   End If
   Template = Replace(Template, "###Name###", Name)
   Template = Replace(Template, "###Description###", ReplaceCRLF(Description))
   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
   members = ""
   For Each CurrMember In SubMembers
      members = members & TemplateMember
      members = Replace(members, "###Name###", CurrMember.Name)
      members = Replace(members, "###Value###", CurrMember.Value)
      members = Replace(members, "###Description###", CurrMember.Description)
   Next
   Template = Replace(Template, "###Members###", members)
   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<p><b> Example</b></p>" & ReplaceCRLF(Example), ""))
   Template = Replace(Template, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<p><b> See Also</b></p>" & ReplaceCRLF(SeeAlso), ""))
   Dim sTmpScreenshot   As String
   sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
   Template = Replace(Template, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

   WriteTextFile FileName, Template
End Sub

'Purpose: Create HTML for Events
Public Sub CreateHTMLEvent(ByVal Template As String, _
   TemplateParameter As String, _
   FileName As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : CreateHTMLEvent
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Template As String
   ' *                    TemplateParameter As String
   ' *                    FileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim CurrParam        As cSubMember
   Dim Syntax           As String
   Dim params           As String
   If gPublicOnly Then
      Template = Replace(Template, "###Scope###", "")
   Else
      Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
   End If

   If Len(Description) = 0 Then Description = Name

   Template = Replace(Template, "###Name###", Name)
   Template = Replace(Template, "###Description###", ReplaceCRLF(Description))
   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###ObjectName###", RefModule)
   'Syntax
   Syntax = "<b>Private Sub </b><i>Object_</i><b>" & Name & " (</b> "
   params = ""
   Syntax = Syntax & "<i>"
   For Each CurrParam In SubMembers
      Syntax = Syntax & " " & IIf(CurrParam.IsOptional, "[", "") & CurrParam.Name & IIf(CurrParam.IsOptional, "]", "") & ","
      params = params & TemplateParameter
      params = Replace(params, "###Name###", CurrParam.Name)
      params = Replace(params, "###Type###", IIf(isAnObject(CurrParam.TypeValue), "As <b>" & CurrParam.TypeValue & "</b> Object.", "As <b>" & CurrParam.TypeValue & "</b> Value."))
      params = Replace(params, "###Description###", Trim$(Replace(CurrParam.Description, "As ", "")))
      params = Replace(params, "###IsOptional###", IIf(CurrParam.IsOptional, "Optional.", "Required."))
   Next
   If right$(Syntax, 1) = "," Then Syntax = left$(Syntax, Len(Syntax) - 1)
   Syntax = Syntax & "</i>"
   Syntax = Syntax & "<b> )</b>"
   Template = Replace(Template, "###Syntax###", Syntax)
   Template = Replace(Template, "###Parameters###", params)
   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<p><b> Example</b></p>" & ReplaceCRLF(Example), ""))
   Template = Replace(Template, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<p><b> See Also</b></p>" & ReplaceCRLF(SeeAlso), ""))
   Dim sTmpScreenshot   As String
   sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
   Template = Replace(Template, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

   WriteTextFile FileName, Template
End Sub

'Purpose: Create HTML for Implements
Public Sub CreateHTMLImplements(ByVal Template As String, FileName As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : CreateHTMLImplements
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Template As String
   ' *                    FileName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Template = Replace(Template, "###Name###", Name)
   Template = Replace(Template, "###Description###", ReplaceCRLF(Description))
   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###ObjectName###", RefModule)
   Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
   Template = Replace(Template, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<p><b> Example</b></p>" & ReplaceCRLF(Example), ""))
   Template = Replace(Template, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<p><b> See Also</b></p>" & ReplaceCRLF(SeeAlso), ""))
   Dim sTmpScreenshot   As String
   sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
   Template = Replace(Template, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

   WriteTextFile FileName, Template
End Sub

'Purpose: Create HTML for Methods
Public Sub CreateHTMLMethod(ByVal Template As String, TemplateParamenter As String, FileName As String, PublicOnly As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : CreateHTMLMethod
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Template As String
   ' *                    TemplateParamenter As String
   ' *                    FileName As String
   ' *                    PublicOnly As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim CurrParam        As cSubMember
   Dim Syntax           As String
   Dim params           As String

   If (PublicOnly And Scope = vbext_Public) Or Not PublicOnly Then
      If PublicOnly Then
         Template = Replace(Template, "###Scope###", "")
      Else
         Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
      End If
      Template = Replace(Template, "###Name###", Name)
      Template = Replace(Template, "###Description###", ReplaceCRLF(Description))
      Template = Replace(Template, "###ReturnTypeMethod###", IIf(isAnObject(TypeValue), _
         IIf(TypeValue = "None", "", "As <b>" & TypeValue & "</b> " & "Object"), _
         IIf(TypeValue = "None", "", "As <b>" & TypeValue & "</b> " & "Value")))

      Template = Replace(Template, "###ReturnType###", IIf(isAnObject(TypeValue), _
         IIf(TypeValue = "None", "<b>None</b>", "As <b>" & TypeValue & "</b> " & "Object"), _
         IIf(TypeValue = "None", "<b>None</b>", "As <b>" & TypeValue & "</b> " & "Value")))
      Template = Replace(Template, "###ObjectName###", RefModule)
      'Syntax
      Syntax = "<i>Object.</i><b>" & Name & "</b> "
      If TypeValue <> "None" Then
         Syntax = Syntax & "<b>(</b>"
      End If
      params = ""
      Syntax = Syntax & "<i>"
      For Each CurrParam In SubMembers
         Syntax = Syntax & " " & IIf(CurrParam.IsOptional, "[", "") & CurrParam.Name & IIf(CurrParam.IsOptional, "]", "") & ","
         params = params & TemplateParamenter
         params = Replace(params, "###Name###", CurrParam.Name)
         'params = Replace(params, "###Type###", IIf(isAnObject(CurrParam.TypeValue), "As <b>" & CurrParam.TypeValue & "</b> Object.", "As <b>" & CurrParam.TypeValue & "</b> Value."))
         params = Replace(params, "###Type###", "")
         params = Replace(params, "###Description###", Trim$(Replace(CurrParam.Description, "As ", "")))
         params = Replace(params, "###IsOptional###", IIf(CurrParam.IsOptional, "Optional.", "Required."))
      Next
      If right$(Syntax, 1) = "," Then Syntax = left$(Syntax, Len(Syntax) - 1)
      Syntax = Syntax & "</i>"
      If TypeValue <> "None" Then
         Syntax = Syntax & "<b>)</b>"
      End If

      ' *** Handle all See Also
      SeeAlso = Trim$(SeeAlso)
      If Len(SeeAlso) > 0 Then

         Dim objCurr          As Object
         Dim objCurr2         As Object
         Dim nI               As Long
         Dim sTmp             As String
         Dim sLink            As String
         Dim sURL             As String
         Dim sModule          As String
         Dim sAllSeeAlso      As String

         SeeAlso = Replace(SeeAlso, vbCr, "")
         SeeAlso = Replace(SeeAlso, vbLf, "")
         SeeAlso = Replace(SeeAlso, " ", "")
         SeeAlso = Replace(SeeAlso, ";", ",")

         ' *** Parse all the SeeAlso
         sAllSeeAlso = ""
         For nI = 1 To CountTokens(SeeAlso, ",")
            sLink = GetToken(SeeAlso, ",", nI)
            If InStr(sLink, ".") > 0 Then
               sModule = GetToken(sLink, ".", 1)
            Else
               sModule = RefModule
            End If
            sURL = ""
            If Len(sModule) > 0 Then
               If InStr(sLink, ".") > 0 Then
                  sLink = GetToken(sLink, ".", 2)
               End If

               ' *** Find the module
               For Each objCurr In objProject.Classes
                  If LCase$(objCurr.Name) = LCase$(sModule) Then
                     ' *** Class found
                     ' *** Search in members

                     For Each objCurr2 In objProject.Methods
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Method.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Properties
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Property.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Consts
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Consts.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Enums
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Enum.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Events
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Event.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Implementations
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Implements.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Types
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Type.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Variables
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Declaration.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                  End If
               Next

               If sURL <> "" Then
                  sTmp = GetLink(sLink, sURL)
               Else
                  sTmp = sLink
               End If

               sAllSeeAlso = sAllSeeAlso & ", " & sTmp

            Else
               sAllSeeAlso = sAllSeeAlso & ", " & sLink

            End If

         Next

         If left$(sAllSeeAlso, 2) = ", " Then
            sAllSeeAlso = Mid$(sAllSeeAlso, 3)
         End If

         SeeAlso = sAllSeeAlso
      End If

      Template = Replace(Template, "###Syntax###", Syntax)
      Template = Replace(Template, "###Parameters###", params)
      Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
      Template = Replace(Template, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<p><b> Example</b></p>" & ReplaceCRLF(Example), ""))
      Template = Replace(Template, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<p><b> See Also</b></p>" & ReplaceCRLF(SeeAlso), ""))
      Dim sTmpScreenshot   As String
      sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
      Template = Replace(Template, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

      WriteTextFile FileName, Template
   End If
End Sub

'Purpose: Create HTML for propery
Public Sub CreateHTMLProperty(ByVal Template As String, _
   TemplateValue As String, _
   FileName As String, _
   PublicOnly As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : CreateHTMLProperty
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Template As String
   ' *                    TemplateValue As String
   ' *                    FileName As String
   ' *                    PublicOnly As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim Syntax           As String

   If (gPublicOnly And Scope = vbext_Public) Or Not gPublicOnly Then
      If PublicOnly Then
         Template = Replace(Template, "###Scope###", "")
      Else
         Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
      End If
      Template = Replace(Template, "###PropValue###", TemplateValue)
      Template = Replace(Template, "###Name###", Name)
      Template = Replace(Template, "###Description###", ReplaceCRLF(Description))
      Template = Replace(Template, "###ReturnType###", IIf(isAnObject(TypeValue), "As <b>" & TypeValue & "</b> " & "Object", "As <b>" & TypeValue & "</b> " & "Value") & ". " & ReadWrite)
      Template = Replace(Template, "###ObjectName###", RefModule)
      Template = Replace(Template, "###TypeValue###", TypeValue)

      'Syntax
      Syntax = ""
      If PropertyType = Prop_Type_ReadWrite And isAnObject(TypeValue) Then
         Syntax = "<b>[Set] </b>"
      ElseIf PropertyType = Prop_Type_WriteOnly And isAnObject(TypeValue) Then
         Syntax = "<b>Set </b>"
      End If

      Syntax = Syntax & "<i>Object.</i><b>" & Name & "</b> "

      If PropertyType = Prop_Type_WriteOnly Then
         Syntax = Syntax & " = <i>" & TypeValue & "</i>"
      ElseIf PropertyType = Prop_Type_ReadWrite Then
         Syntax = Syntax & "[ = <i>" & TypeValue & "</i> ]"
      End If
      Template = Replace(Template, "###Syntax###", Syntax)

      ' *** Handle all See Also
      SeeAlso = Trim$(SeeAlso)
      If Len(SeeAlso) > 0 Then

         Dim objCurr          As Object
         Dim objCurr2         As Object
         Dim nI               As Long
         Dim sTmp             As String
         Dim sLink            As String
         Dim sURL             As String
         Dim sModule          As String
         Dim sAllSeeAlso      As String

         SeeAlso = Replace(SeeAlso, vbCr, "")
         SeeAlso = Replace(SeeAlso, vbLf, "")
         SeeAlso = Replace(SeeAlso, " ", "")
         SeeAlso = Replace(SeeAlso, ";", ",")

         ' *** Parse all the SeeAlso
         sAllSeeAlso = ""
         For nI = 1 To CountTokens(SeeAlso, ",")
            sLink = GetToken(SeeAlso, ",", nI)
            If InStr(sLink, ".") > 0 Then
               sModule = GetToken(sLink, ".", 1)
            Else
               sModule = RefModule
            End If
            sURL = ""
            If Len(sModule) > 0 Then
               If InStr(sLink, ".") > 0 Then
                  sLink = GetToken(sLink, ".", 2)
               End If

               ' *** Find the module
               For Each objCurr In objProject.Classes
                  If LCase$(objCurr.Name) = LCase$(sModule) Then
                     ' *** Class found
                     ' *** Search in members

                     For Each objCurr2 In objProject.Methods
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Method.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Properties
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Property.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Consts
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Consts.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Enums
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Enum.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Events
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Event.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Implementations
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Implements.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Types
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Type.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                     For Each objCurr2 In objProject.Variables
                        If (objCurr2.RefModule = sModule) And (objCurr2.Name = sLink) Then
                           sURL = sModule & "_" & sLink & "_Declaration.htm"
                           Exit For
                        End If
                     Next
                     If Len(sURL) > 0 Then Exit For

                  End If
               Next

               If sURL <> "" Then
                  sTmp = GetLink(sLink, sURL)
               Else
                  sTmp = sLink
               End If

               sAllSeeAlso = sAllSeeAlso & ", " & sTmp

            Else
               sAllSeeAlso = sAllSeeAlso & ", " & sLink

            End If

         Next

         If left$(sAllSeeAlso, 2) = ", " Then
            sAllSeeAlso = Mid$(sAllSeeAlso, 3)
         End If

         SeeAlso = sAllSeeAlso
      End If

      Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
      Template = Replace(Template, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<p><b> Example</b></p>" & ReplaceCRLF(Example), ""))
      Template = Replace(Template, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<p><b> See Also</b></p>" & ReplaceCRLF(SeeAlso), ""))
      Dim sTmpScreenshot   As String
      sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
      Template = Replace(Template, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

      WriteTextFile FileName, Template
   End If
End Sub

'Purpose: Create HTML for UDT's
Public Sub CreateHTMLType(ByVal Template As String, _
   TemplateMember As String, _
   FileName As String, _
   PublicOnly As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : CreateHTMLType
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Template As String
   ' *                    TemplateMember As String
   ' *                    FileName As String
   ' *                    PublicOnly As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim CurrMember       As cSubMember
   Dim members          As String

   If (PublicOnly And Scope = vbext_Public) Or Not PublicOnly Then
      If PublicOnly Then
         Template = Replace(Template, "###Scope###", "")
      Else
         Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
      End If
      Template = Replace(Template, "###Name###", Name)
      Template = Replace(Template, "###Description###", ReplaceCRLF(Description))
      Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
      members = ""
      For Each CurrMember In SubMembers
         members = members & TemplateMember
         members = Replace(members, "###Name###", CurrMember.Name)
         members = Replace(members, "###Type###", CurrMember.TypeValue)
         members = Replace(members, "###Description###", CurrMember.Description)
      Next
      Template = Replace(Template, "###Members###", members)
      Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
      Template = Replace(Template, "###Example###", IIf(Len(ReplaceCRLF(Example)) > 0, "<p><b> Example</b></p>" & ReplaceCRLF(Example), ""))
      Template = Replace(Template, "###SeeAlso###", IIf(Len(ReplaceCRLF(SeeAlso)) > 0, "<p><b> See Also</b></p>" & ReplaceCRLF(SeeAlso), ""))
      Dim sTmpScreenshot   As String
      sTmpScreenshot = Trim$(Replace(Replace(Screenshot, vbCr, ""), vbLf, ""))
      Template = Replace(Template, "###Screenshot###", IIf(Len(sTmpScreenshot) > 0, "<p><b> Screenshot</b></p><p align=""center""><img src=""" & sTmpScreenshot & """></p>", ""))

      WriteTextFile FileName, Template
   End If
End Sub

'Purpose: Create HTML for Variables<BR>
'IMPORTANT: Each public variabile is a read and write property
Public Sub CreateHTMLVariable(ByVal Template As String, _
   TemplateValue As String, _
   FileName As String, _
   PublicOnly As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : CreateHTMLVariable
   ' * Purpose          :
   ' * Parameters       :
   ' *                    ByVal Template As String
   ' *                    TemplateValue As String
   ' *                    FileName As String
   ' *                    PublicOnly As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   '###VarProp###
   Dim Syntax           As String

   If (PublicOnly And Scope = vbext_Public) Or Not gPublicOnly Then
      If PublicOnly Then
         Template = Replace(Template, "###Scope###", "")
      Else
         Template = Replace(Template, "###Scope###", ScopeToString(Scope) & " ")
      End If

      'If is a Developer documentation create a Variable page
      Template = Replace(Template, "###VarProp###", IIf(gVarAsProperty Or gPublicOnly, "Property", "Variable"))

      If Len(Description) = 0 Then Description = Name

      Template = Replace(Template, "###PropValue###", TemplateValue)
      Template = Replace(Template, "###Name###", Name)
      Template = Replace(Template, "###Description###", ReplaceCRLF(Description))
      Template = Replace(Template, "###ReturnType###", _
         IIf(isAnObject(TypeValue), "As <b>" & TypeValue & "</b> " & "Object", "As <b>" & TypeValue & "</b> " & "Value") & _
         IIf(gVarAsProperty Or gPublicOnly, ". The property is Read-Write", ""))
      Template = Replace(Template, "###Remarks###", IIf(Len(ReplaceCRLF(Remarks)) > 0, "<p><b> Remarks</b></p>" & ReplaceCRLF(Remarks), ""))
      Template = Replace(Template, "###ObjectName###", RefModule)
      Template = Replace(Template, "###TypeValue###", TypeValue)

      'Syntax
      Syntax = ""
      If isAnObject(TypeValue) Then
         Syntax = "<b>[Set] </b>"
      End If

      Syntax = Syntax & "<i>Object.</i><b>" & Name & "</b> "
      Syntax = Syntax & "[ = <i>" & TypeValue & "</i> ]"

      Template = Replace(Template, "###Syntax###", Syntax)
      Template = Replace(Template, "###Example###", "")

      WriteTextFile FileName, Template
   End If
End Sub

'Purpose: returns a string value that describes the type of property. <BR>
'Used for creation of the HTML file of a property.
Private Function ReadWrite() As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : ReadWrite
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   If PropertyType = Prop_Type_ReadOnly Then
      ReadWrite = "The Property is Read Only"
   ElseIf PropertyType = Prop_Type_ReadWrite Then
      ReadWrite = "The Property is Read-Write"
   Else
      ReadWrite = "The Property is Write Only"
   End If
End Function

Private Function GetLink(sName As String, sURL As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : cMember
   ' * Module Filename  : VBDoccMember.cls
   ' * Procedure Name   : GetLink
   ' * Purpose          :
   ' * Parameters       :
   ' *                    sName As String
   ' *                    sURL As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim strTemp          As String
   'strTemp = "<OBJECT TYPE=""application/x-oleobject"" CLASSID=""clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11"" ID=""" & sName & """>" & vbCrLf
   'strTemp = strTemp & "<PARAM NAME=""Command"" value=""ALink"">" & vbCrLf
   'strTemp = strTemp & "<PARAM name=""DefaultTopic"" value=""_topic_not_found.htm"">" & vbCrLf
   'strTemp = strTemp & "<PARAM name=""Item1"" value="""">" & vbCrLf
   'strTemp = strTemp & "<PARAM name=""Item2"" value=""" & sURL & """>" & vbCrLf
   'strTemp = strTemp & "<PARAM NAME=""Font"" VALUE=""Helvetica,9,,#0000ff,UNDERLINE>" & vbCrLf
   'strTemp = strTemp & "</OBJECT>" & vbCrLf
   'strTemp = strTemp & "<a href=JavaScript:" & sName & ".Click()>" & sName & "</a>" & vbCrLf
   strTemp = strTemp & "<a href=""" & sURL & """>" & sName & "</a>"
   GetLink = strTemp
End Function

