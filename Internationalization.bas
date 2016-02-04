Attribute VB_Name = "Internationalization_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 08/11/1999
' * Time             : 12:30
' **********************************************************************
' * Comments         : Translate string from resource string
' *
' *
' **********************************************************************
Option Explicit

Global gnLanguage       As Long

Global Const ENGLISH_LANG = 0
Global Const FRENCH_LANG = 10000
Global Const SPANISH_LANG = 15000
Global Const GERMAN_LANG = 20000
Global Const DANISH_LANG = 25000

Public Sub InternationalizeForm(theForm As Form)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 12:07
   ' * Module Name      : Internationalize
   ' * Module Filename  : Internationalize.BAS
   ' * Procedure Name   : InternationalizeForm
   ' * Parameters       :
   ' *                    theForm As Form
   ' **********************************************************************
   ' * Comments         : Translate all the controls on a form
   ' * We use the ID after the µ
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_InternationalizeForm

   Dim nI               As Long
   Dim nJ               As Long

   ' *** Translate the caption of the form
   theForm.Caption = Translation(CStr(theForm.Caption))

   For nI = 0 To theForm.Controls.Count - 1
      ' *** Translate all the captions
      On Error Resume Next
      If TypeOf theForm.Controls(nI) Is Toolbar Then
         For nJ = 1 To theForm.Controls(nI).Buttons.Count
            With theForm.Controls(nI).Buttons(nJ)
               .Caption = Translation(.Caption)
               .ToolTipText = Translation(.ToolTipText)
            End With
         Next
      End If

      theForm.Controls(nI).Caption = Translation(CStr(theForm.Controls(nI).Caption))
      theForm.Controls(nI).ToolTipText = Translation(theForm.Controls(nI).ToolTipText)

      If TypeOf theForm.Controls(nI) Is TabStrip Then
         With theForm.Controls(nI)
            For nJ = 1 To .Tabs.Count
               .Tabs(nJ).Caption = Translation(CStr(.Tabs(nJ).Caption))
            Next
         End With
      End If

   Next

EXIT_InternationalizeForm:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_InternationalizeForm:

   Resume EXIT_InternationalizeForm

End Sub

Public Function Translation(sText As String) As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 12:07
   ' * Module Name      : Internationalize
   ' * Module Filename  : Internationalize.BAS
   ' * Procedure Name   : Translation
   ' * Parameters       :
   ' *                    sText As String
   ' **********************************************************************
   ' * Comments         :
   ' * Translate a text using the ID in the string after the µ
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Translation

   Dim sTmp             As String
   Dim sID              As String
   Dim nPos             As Long

   Translation = sText

   sTmp = ""

   ' *** Search the µ
   nPos = InStr(sText, "µ")

   ' *** No Translation found
   If (nPos = 0) Then
      Translation = sText
      Resume EXIT_Translation
   End If

   sID = right$(sText, Len(sText) - nPos)

   ' *** No identifiant found
   If (IsNumeric(sID) = False) Then
      Translation = left$(sText, nPos - 1)
      Resume EXIT_Translation
   End If

   sTmp = LoadResString(gnLanguage + CLng(sID))

   ' *** No Translation found
   If (sTmp = "") Then
      Translation = left$(sText, nPos - 1)
      Resume EXIT_Translation
   End If

   Translation = sTmp

EXIT_Translation:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_Translation:
   If (IsNumeric(sID) = True) Then
      Translation = left$(sText, nPos - 1)
      Resume EXIT_Translation
   End If
   Resume EXIT_Translation

End Function
