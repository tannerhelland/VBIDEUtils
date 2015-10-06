VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Hourglass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 7/05/98
' * Time             : 11:03
' * Module Name      : class_Hourglass
' * Module Filename  : Hourglass.cls
' **********************************************************************
' * Comments         : Add easily mouse hourglass to you processing code
' * witouth any care about setting back the hourglass to previous pointer.
' *
' * You have to do as follow :
' *    Private Sub YourFunction()
' *
' *    Dim cHourglass    as New class_Hourglass
' *
' *    ' .... Do stuff
' *    ' .... The hourglass will be reset to previous pointer once
' *    ' .... you quit thefunction
' *
' *    End Sub
' *
' **********************************************************************

Option Explicit

Private nOldMousePointer As Integer

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : class_Hourglass
   ' * Module Filename  : Hourglass.cls
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Show the Hourglass

   nOldMousePointer = Screen.MousePointer
   Screen.MousePointer = vbHourglass

End Sub

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:56
   ' * Module Name      : class_Hourglass
   ' * Module Filename  : Hourglass.cls
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :  Show previous pointer
   ' *
   ' *
   ' **********************************************************************

   Screen.MousePointer = nOldMousePointer

End Sub
