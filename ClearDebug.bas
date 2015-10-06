Attribute VB_Name = "ClearDebug_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:51
' * Module Name      : ClearDebug_Module
' * Module Filename  : ClearDebug.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Public Sub ClearDebug()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 12/10/1998
   ' * Time             : 20:03
   ' * Module Name      : ClearDebug_Module
   ' * Module Filename  : ClearDebug.bas
   ' * Procedure Name   : ClearDebug
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         : Clear the debug window
   ' *
   ' *
   ' **********************************************************************

   Dim pWindow          As VBIDE.Window

   Set pWindow = VBInstance.Windows("Immediate")

   If pWindow Is Nothing Then Exit Sub

   If pWindow.Visible = True Then
      pWindow.SetFocus
      SendKeys ("^({Home})"), True
      SendKeys ("^(+({End}))"), True
      SendKeys ("{Del}"), True
   End If
   Set pWindow = Nothing

End Sub
