VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 15/12/1999
' * Time             : 16:13
' * Module Name      : class_Search
' * Module Filename  : CSearch.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private colFunctionList As Collection
Private colWhoCallsMe   As Collection
Private m_ScanDone      As Boolean

Public Function BuildMenu(sKey As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:11
   ' * Module Name      : class_Search
   ' * Module Filename  : Search.cls
   ' * Procedure Name   : BuildMenu
   ' * Parameters       :
   ' *                    sKey As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim colWhoCallMe     As Collection
   Dim clsFuncData      As class_FunctionData

   ' *** Get the collection of who called this routine
   On Error Resume Next
   err.Clear
   Set colWhoCallMe = colWhoCallsMe.Item(sKey)

   ' *** Err # 5 is thrown if the key does not exist
   BuildMenu = ""
   If err.number = 0 Then
      On Error Resume Next

      ' *** Walk the collection and build a menu to display
      ' *** ENHANCEMENT   Limit the number of menu items based on screen real estate
      For Each clsFuncData In colWhoCallMe
         BuildMenu = BuildMenu & "@" & clsFuncData.FuncName & "@" & clsFuncData.Module & "~!" & clsFuncData.FuncName & "~!" & clsFuncData.LineNumber
      Next
   Else
      ' *** If the collection does not exist then the funcion is never used
      If err.number = 5 Then
         BuildMenu = ""
      End If
   End If

End Function

Public Sub ScanForFunctionNames()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_Search
   ' * Module Filename  : Search.cls
   ' * Procedure Name   : ScanForFunctionNames
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim vbComponentObj   As VBComponent
   Dim vbMemberObj      As Member
   Dim sKey             As String
   Dim clsFunctionData  As class_FunctionData

   Dim nCount           As Integer

   On Error Resume Next

   '
   ' *** Clear out any current members in the collection
   Set colFunctionList = Nothing
   Set colFunctionList = New Collection

   frmProgress.MessageText = "Phase 1/3"
   frmProgress.Maximum = VBInstance.ActiveVBProject.VBComponents.Count
   nCount = 1

   For Each vbComponentObj In VBInstance.ActiveVBProject.VBComponents
      frmProgress.Progress = nCount
      nCount = nCount + 1

      ' Debug.Print vbComponentObj.Name
      For Each vbMemberObj In vbComponentObj.CodeModule.members
         ' *** The member type tells us if this is a function or a variable
         Select Case vbMemberObj.Type
            Case vbext_mt_Method, vbext_mt_Event, vbext_mt_Property
               '
               ' *** Add it to our collection

               ' The key is made up of the module name and the function name
               sKey = vbComponentObj.Name + "!" + vbMemberObj.Name
               ' Debug.Print vbTab + vbTab + sKey

               Set clsFunctionData = New class_FunctionData
               clsFunctionData.FuncName = vbMemberObj.Name
               clsFunctionData.Module = vbComponentObj.Name
               colFunctionList.Add clsFunctionData, sKey
            Case vbext_mt_Variable
               ' Debug.Print vbTab & vbMemberObj.Name & vbTab & "Variable"
            Case vbext_mt_Const
               ' Debug.Print vbTab & vbMemberObj.Name & vbTab & "Constant"
         End Select
      Next
   Next

End Sub

Private Function IsInsideAComment(strLine As String, strFuncName As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_Search
   ' * Module Filename  : Search.cls
   ' * Procedure Name   : IsInsideAComment
   ' * Parameters       :
   ' *                    strLine As String
   ' *                    strFuncName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nIndex           As Integer
   Dim nLength          As Integer
   Dim bQCount          As Boolean
   Dim bytSource()      As Byte

   ' *** Convert to a byte array
   bytSource = strLine

   ' *** Find it's length for walking the length
   nLength = Len(strLine)

   For nIndex = 0 To nLength - 1
      ' *** If the character is a quote,
      If bytSource(nIndex) = 34 Then bQCount = Not bQCount

      ' *** bQCount is True then we are inside a double quote so we ignore the sigle quote
      If bQCount = False And bytSource(nIndex) = 39 Then
         ' *** we found a single quote which not inside of double quotes
         ' *** then it must be a comment character
         If InStr(strLine, strFuncName) > nIndex Then
            ' *** If the function name we are looking for is beyond the comment, return True
            IsInsideAComment = True
         End If
         Exit Function
      End If

   Next

End Function

Public Sub ScanForFunctionUse()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_Search
   ' * Module Filename  : Search.cls
   ' * Procedure Name   : ScanForFunctionUse
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim vbComponentObj   As VBComponent
   Dim vbMemberObj      As Member
   Dim vbCodemoduleObj  As CodeModule
   Dim nProcedureType   As Long
   Dim nCount           As Integer

   On Error GoTo ERROR_ScanForFunctionUse

   frmProgress.MessageText = "Phase 2/3"
   frmProgress.Maximum = VBInstance.ActiveVBProject.VBComponents.Count
   nCount = 1

   ' *** We are going to loop through the components of the active project
   ' *** Inside each of them we will look at individual code modules
   For Each vbComponentObj In VBInstance.ActiveVBProject.VBComponents
      frmProgress.Progress = nCount
      frmProgress.MessageText = "Phase 2/3" & vbCrLf & vbComponentObj.Name
      DoEvents
      nCount = nCount + 1

      Debug.Print nCount & " - " & vbComponentObj.Name
      Set vbCodemoduleObj = vbComponentObj.CodeModule

      If Not (vbComponentObj Is Nothing) Then
         If Not (vbComponentObj.CodeModule Is Nothing) Then
            ' *** The members of a CodeModule are the Functions, Subs and variable declarations
            For Each vbMemberObj In vbComponentObj.CodeModule.members
               ' Debug.Print vbComponentObj.Name & "|" & vbMemberObj.Name
               SearchForFunction vbCodemoduleObj, vbMemberObj, nProcedureType
            Next
         End If
      End If
   Next

   Exit Sub

ERROR_ScanForFunctionUse:
   Resume Next
   Resume

End Sub

Sub SearchForFunction(vbCodemoduleObj As CodeModule, vbMemberObj As Member, nProcedureType As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_Search
   ' * Module Filename  : Search.cls
   ' * Procedure Name   : SearchForFunction
   ' * Parameters       :
   ' *                    vbCodemoduleObj As CodeModule
   ' *                    vbMemberObj As Member
   ' *                    nProcedureType As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nBodyStart       As Long
   Dim nStartLine       As Long
   Dim nLineCount       As Long
   Dim nProcedureLineCount As Long
   Dim sProcedureCode   As String
   Dim sProcedureLine   As String
   Dim nIndexCounter    As Integer
   Dim j                As Long
   Dim colCallList      As Collection
   Dim clsFuncData      As class_FunctionData
   Dim clsData          As class_FunctionData
   Dim sFunctionName    As String
   Dim sFunctionArray() As String
   Dim sKey             As String

   On Error GoTo ERROR_SearchForFunction

   ' *** if the procedure type is not defined, skip out
   nProcedureType = -1
   Select Case vbMemberObj.Type
      Case vbext_mt_Method, vbext_mt_Event
         nProcedureType = vbext_pk_Proc

      Case vbext_mt_Property

         ' *** This all necessary due to the fact that we have to knwo the proper type of Get/Let to use
         ' *** We are just getting one line now because we are doing this just to test and probe
         sProcedureCode = vbCodemoduleObj.Lines(vbCodemoduleObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Get), 1)
         If err.number = 35 Then ' Sub or Function not defined
            err.Clear
            sProcedureCode = vbCodemoduleObj.Lines(vbCodemoduleObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Let), 1)
            If err.number = 35 Then ' Sub or Function not defined
               err.Clear
               sProcedureCode = vbCodemoduleObj.Lines(vbCodemoduleObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Set), 1)
               If err.number = 0 Then
                  nProcedureType = vbext_pk_Set
               Else
                  ' could not determine the type...
                  nProcedureType = -1
               End If
            Else
               nProcedureType = vbext_pk_Let
            End If
         Else
            nProcedureType = vbext_pk_Get
         End If

   End Select

   ' *** If we could not determine the type or it is not a type we wnat, skip on outta here
   If nProcedureType = -1 Then Exit Sub

   ' *** We are going to inspect each function as a whole for the presence
   ' *** of any procedure calls.  To do so we need to get information about
   ' *** the code for each procedure.

   ' *** First we will get the starting lines and line count of the procedure
   nBodyStart = vbCodemoduleObj.ProcBodyLine(vbMemberObj.Name, nProcedureType)
   nStartLine = vbCodemoduleObj.ProcStartLine(vbMemberObj.Name, nProcedureType)
   nLineCount = vbCodemoduleObj.ProcCountLines(vbMemberObj.Name, nProcedureType)

   ' *** ProcCountLInes returns the number of lines INCLUDING any comments before the actual
   ' *** start of the code.  We don't want those so we calculate a new line count which only
   ' *** includes the actual body of code
   nProcedureLineCount = nLineCount - (nBodyStart - nStartLine)

   ' *** But, as we look through the code for occurances of functions, the entry point (as MS calls it)
   ' *** will cause false hits so lets skip the first line
   nBodyStart = nBodyStart + 1

   ' *** And we might as well skip the End Sub line
   nProcedureLineCount = nProcedureLineCount - 2

   ' *** Now, let's actually get the code
   ' *** obj.Lines (startline, count)
   sProcedureCode = vbCodemoduleObj.Lines(nBodyStart, nProcedureLineCount)

   ' *** Look in the string and see if any function exist in the procedure

   ' *** This redim will remove any old data since we do not use Preserve
   ReDim sFunctionArray(1, 0) As String

   ' *** colFunctionList contains a list of all of the functions found in the project
   ' *** Here we go through that list to see it any of them occcur in the current procedure.
   ' *** As we find a procedure that contains a call, we add it to our sFunctionArray array.
   ' *** After we are finished we will go through the functions in sFunctionArray and search
   ' *** them line-by-line to find the exact match
   For Each clsData In colFunctionList
      ' *** Save some cycle time by using a local variable instead of a property
      sFunctionName = clsData.FuncName

      ' *** In a collection I cannot get the Key value so I rebuild the key here for later use
      sKey = clsData.Module + "!" + clsData.FuncName

      sProcedureCode = Trim$(sProcedureCode)

      ' *** Look if it exists
      ' *** the Like statement can be faster than InStr
      ' *** the Like statement also provides for the use of wild cards where InStr does not
      If InStrB(sProcedureCode, sFunctionName) Then 'sProcedureCode Like "*" & sFunctionName & "*" Then

         ' *** Ignore assignments to return a value
         If Not InStrB(sProcedureCode, sFunctionName & " = ") Then 'Not sProcedureCode Like "*" & sFunctionName & " = *" Then

            ' *** Somewhere in this procedure a call is made to a declared function
            ' *** Store it into an array for detailed scanning later
            ReDim Preserve sFunctionArray(1, UBound(sFunctionArray, 2) + 1)
            sFunctionArray(0, UBound(sFunctionArray, 2)) = sFunctionName
            sFunctionArray(1, UBound(sFunctionArray, 2)) = sKey
         End If
      End If
   Next

   ' *** Now lets go through the function in line by line detail,
   ' *** looking only for those functions that appear in the array
   On Error Resume Next
   If UBound(sFunctionArray) > 0 Then
      For nIndexCounter = nBodyStart To nBodyStart + nProcedureLineCount

         ' *** Get the individual line from the CodeModule
         sProcedureLine = Trim$(vbCodemoduleObj.Lines(nIndexCounter, 1))
         For j = 1 To UBound(sFunctionArray, 2)

            '================================================================================================
            '================================================================================================
            '
            ' *** Here is where we actually do the comparisons to see if the function actually appears in a line
            ' *** Add as many additional conditions here as you see fit
            '
            '================================================================================================
            '================================================================================================

            ' *** Is it on this line?
            'If sProcedureLine Like "*" & sFunctionArray(0, j) & "*" Then
            If InStrB(sProcedureLine, sFunctionArray(0, j)) Then
               If Not IsInsideAComment(sProcedureLine, sFunctionArray(0, j)) Then
                  ' *** Ignore if the procedure is actually a label
                  'If Not sProcedureLine Like sFunctionArray(0, j) & ":" Then
                  If Not InStrB(sProcedureLine, sFunctionArray(0, j) & ":") Then
                     ' *** Or part of an On Error Go to
                     If Not sProcedureLine Like "On Error *" & sFunctionArray(0, j) & "*" Then
                        ' *** And finally, make sure the name is not part of another name
                        ' *** The LCase is used here to allow for the [!a-z] condition.  This check makes
                        ' *** sure that function name we are looking for is not a subset of another string
                        ' *** granted, if numbers appear before or after the sFunctionname string then
                        ' *** this check will fail.  But since I wrote thhs pass and I do not
                        ' *** use numbers in my prcedure names, this does not cause me any problem.
                        'If LCase$(sProcedureLine) Like "[!a-z]" & LCase$(sFunctionArray(0, j)) & "[!a-z]" Then
                        If LCase$(sProcedureLine) Like "*" & LCase$(sFunctionArray(0, j)) & "*" Then                           '
                           ' *** Create a new object to be added to the collection and then fill the data
                           Set clsFuncData = New class_FunctionData
                           clsFuncData.FuncName = sFunctionArray(0, j) 'vbMemberObj.Name
                           clsFuncData.Module = vbCodemoduleObj.Parent.Name
                           clsFuncData.LineNumber = nIndexCounter
                           '
                           ' *** Cheating is allowed.  Since there is no Exist method in a collection
                           ' *** I assume that the key exists and handle the error if it doesn't
                           sKey = sFunctionArray(1, j)
                           err.Clear
                           Set colCallList = colWhoCallsMe.Item(sKey)

                           If err.number <> 0 Then
                              err.Clear
                              ' *** Did it fail because there is nothing in colWhoCallsMe
                              If IsEmpty(colWhoCallsMe) Then
                              Else
                                 ' *** Or because the key does not exist
                                 ' *** Let's add a member that has the required key
                                 Set colCallList = New Collection
                                 colCallList.Add clsFuncData, CStr(Rnd(time))
                                 ' *** Now add this collection to the original collection
                                 colWhoCallsMe.Add colCallList, sKey
                              End If
                           Else
                              colCallList.Add clsFuncData, CStr(Rnd(time))
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         Next
      Next
   End If

   Exit Sub

ERROR_SearchForFunction:
   Exit Sub
   Resume Next
   Resume

End Sub

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_Search
   ' * Module Filename  : Search.cls
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Set colFunctionList = New Collection
   Set colWhoCallsMe = New Collection

End Sub
