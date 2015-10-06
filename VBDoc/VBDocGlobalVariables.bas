Attribute VB_Name = "GlobalVariables"
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : GlobalVariables
' * Module Filename  : VBDocGlobalVariables.bas
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Project: Documentation Generator
'This Add-in for Visual Basic 6.0 generate a documentation in CHM format
'   directly from code comments.<BR>
'   It's possible to generate a technical documentation for ActiveX components
'   or a complete developer documentation, for the management of Visual Basic
'   projects.<BR>
'   This Full <I>Auto-Documentation</I> shows how to generate a correct documentation.<BR>
'   In order to create the CHM files is needed the HTML HELP Compliler.<BR>
'   The HELP HTML WorkShop with the hhc.exe file included is freeware and
'   avaible (4.00 MB) at Microsoft site<BR>
'   <A HREF="http://msdn.microsoft.com/library/en-us/htmlhelp/html/hwMicrosoftHTMLHelpDownloads.asp">
'   http://msdn.microsoft.com/library/en-us/htmlhelp/html/hwMicrosoftHTMLHelpDownloads.asp</A><BR>
'   Other requirements are:<BR>
'   Microsoft Office in order to add a menu to Visual Basic IDE
'   (Automatically called from the Add-In Wizard Creation of Visual Basic.
'   If you don't have MSOffice you must create an add-in and serach for code
'   where the add-in insert the menu item)<BR>
'   <BR><BR><IMG src="ScreenShoot1.jpg"><BR><BR>
'   You can personalize the Tags comment, choose the modules you want insert into
'   documentation.<BR>
'   If you want create a technical documentation you must check
'   <B>Public member only</B><BR>
'   Clearly, you can't show the private member or source code.<BR>
'   The check <B>Variables as Property</B> show all variable like property.
'   This only it's necessary in techical documentation but not in
'   developer decomentation.<BR>
'   In order to use this add-in you <B>must</B> save the current project
'   and select the CHM compiler (hhc.exe file).<BR>
'   All changes you made are registered in the registry and saved for next use!<BR>
'   I parse all declaration and nearly all case  but ...  nothing is perfect,
'   I don't have parsed the ParamArray <BR>
'   In order to write the project documentation you must write the Project Tag
'   in the first line of a whichever module. After you can write the documentation
'   for the module after another Project Tag and the next line with the Purpose Tag.
'   (see the Global variables module for a valid example) <BR>
'   In order to write the module documentation you must write the Purpose Tag
'   in the first line of module.<BR>
'   When you write a comment for a member it's important that have no blank
'   line between the last comment and the declaration.<BR>
'   For comment variables, const, UDTs member and Enum member you can write it
'   after the declaration on the same line using also the next line of code
'   with no Tag.<BR>
'   <B>IMPORTANT:</B> This application don't recognize multiple
'   declaration on the same code line.<BR>
'   <B>YOU CAN ADD EVERY HTML TAG TO YOUR COMMENTS FOR A BETTER VISUALIZATION</B><BR>
'   There are a little of comments 'cause I'm Italian and I don't speak currently the
'   english language :-( <BR><BR>
'   Please send me your feedback and report the bugs that you encounter.<BR><BR>
'Author: <B>Marco Pipino</B><BR> <A HREF="mailto:marcopipino@libero.it"> marcopipino@libero.it</A>
'Example:
'An example of this application is the following code<BR>
'Code:
'Code:
'Code:'Purpose: This function do something
'Code:'Parameter:Param1 Descr of Param1
'Code:'Parameter:Param2 Descr of Param2
'Code:'Remarks: Remarks ....
'Code:Public Function SampleFunction(Param1 As String, Param2 As Long) as Integer
'Code:    Dim a As Integer
'Code:    ........
'Code:End Function
'<BR>You can change the tags an write your code in this mode, setting the
'   no comment tag as <I><B>##</B></I>, Purpose tag as <I><B>#Scope</B></I><BR><BR>
'Code:'###########################################
'Code:'#Scope This function do something
'Code:'#Parameter Param1 Descr of Param1
'Code:'#Parameter Param2 Descr of Param2
'Code:'#Remarks Remarks ....
'Code:'     Second line fo remarks
'Code:'###########################################
'Code:Public Function SampleFunction(Param1 As String, Param2 As Long) as Integer
'Code:      ..........
'Code:End Function
'
'Project:
'Purpose: This module contain the global variable for the project.
Option Explicit

Global gBLOCK_PROJECT   As String             'The tag for Project comments
Global gBLOCK_AUTHOR    As String              'The tag for Author
Global gBLOCK_DATE_CREATION As String       'The tag for Date of creation
Global gBLOCK_DATE_LAST_MOD As String       'The tag for Date of last modification
Global gBLOCK_VERSION   As String             'The tag for Version
Global gBLOCK_PURPOSE   As String
Global gBLOCK_REMARKS   As String
Global gBLOCK_PARAMETER As String
Global gBLOCK_EXAMPLE   As String
Global gBLOCK_SEEALSO   As String
Global gBLOCK_SCREEENSHOT As String
Global gBLOCK_CODE      As String
Global gBLOCK_TEXT      As String
Global gBLOCK_NO_COMMENT As String
Global gBLOCK_WEBSITE   As String
Global gBLOCK_EMAIL     As String
Global gBLOCK_TIME      As String
Global gBLOCK_TEL       As String
Global gBLOCK_PROCEDURE_NAME As String
Global gBLOCK_MODULE_NAME As String
Global gBLOCK_MODULE_FILE As String

Global gHHPTemplate     As String               'HHP File template
Global gHHCCompiler     As String               'Path of the HHC Compiler

Global gTypeValues      As Collection            'Collection of TYpe Value

Global gProjectFolder   As String             'Project Folder
Global gProjectName     As String               'Project Name
Global gPublicOnly      As Boolean               'indicate if user want a user documentation
'or a developer documentation
Global gVarAsProperty   As Boolean            'Determine if view variables like property
Global gSourceCode      As Boolean               'Determine if insert module source code

Public objProject       As cProject          'The cProject object
