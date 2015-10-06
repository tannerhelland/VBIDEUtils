Attribute VB_Name = "ADOConnection_Module"
Option Explicit

Public Sub CreateNewADOConnection()
   ' #VBIDEUtils#************************************************************
   ' * Author           : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 02/28/2001
   ' * Time             : 22:31
   ' * Module Name      : ADOConnection_Module
   ' * Module Filename  : ADOConnection.bas
   ' * Procedure Name   : CreateNewADOConnection
   ' * Purpose          : Create easily an ADO connection string
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' * Create easily an ADO connection string and copies it to the clipboard
   ' *
   ' * Example          :
   ' * Call CreateNewADOConnection
   ' *
   ' * See Also         :
   ' *
   ' * History          :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_CreateNewADOConnection

   Dim objDataLink      As New DataLinks
   Dim sConnection      As String
   Dim sMsgbox          As String

   sConnection = objDataLink.PromptNew

   ' *** Copy to the clipboard
   Clipboard.Clear
   Clipboard.SetText sConnection, vbCFText

   ' *** Copy done
   sMsgbox = sConnection & Chr$(13) & "has been copied to the clipboard." & Chr$(13)
   sMsgbox = sMsgbox & "You can paste it in your code."

   Call MsgBox(sMsgbox, vbOKOnly + vbInformation, "ADO Connection String Creator")

EXIT_CreateNewADOConnection:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_CreateNewADOConnection:
   Resume EXIT_CreateNewADOConnection

End Sub
