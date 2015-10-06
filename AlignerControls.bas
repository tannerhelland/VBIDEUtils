Attribute VB_Name = "Aligner_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 03/25/2001
' * Time             : 13:25
' * Module Name      : Aligner_Module
' * Module Filename  : AlignerControls.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Public Sub AlignControlsLeft()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/25/2001
   ' * Time             : 13:24
   ' * Module Name      : Aligner_Module
   ' * Module Filename  : AlignerControls.bas
   ' * Procedure Name   : AlignControlsLeft
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_AlignControlsLeft

   Dim AllControls
   Dim Control          As Object
   Dim nControlLeft     As Long

   nControlLeft = 0

   ' *** Get collection containing all selected controls on form
   Set AllControls = VBInstance.SelectedVBComponent.Designer.SelectedVBControls

   ' *** For each control on the active form...
   For Each Control In AllControls
      ' *** Find the farthest left
      With Control.Properties
         If nControlLeft = 0 Then
            nControlLeft = .Item("Left")
         Else
            If .Item("left") < nControlLeft Then nControlLeft = .Item("Left")
         End If
      End With
   Next Control

   ' *** For each control on the active form...
   For Each Control In AllControls
      ' *** ...align control
      With Control.Properties
         .Item("left") = nControlLeft
      End With
   Next Control

EXIT_AlignControlsLeft:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_AlignControlsLeft:
   Resume EXIT_AlignControlsLeft

End Sub

Public Sub AlignControlsRight()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/25/2001
   ' * Time             : 13:24
   ' * Module Name      : Aligner_Module
   ' * Module Filename  : AlignerControls.bas
   ' * Procedure Name   : AlignControlsRight
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_AlignControlsRight

   Dim AllControls
   Dim Control          As Object
   Dim nControlRight    As Long

   nControlRight = 0

   ' *** Get collection containing all selected controls on form
   Set AllControls = VBInstance.SelectedVBComponent.Designer.SelectedVBControls

   ' *** For each control on the active form...
   For Each Control In AllControls
      ' *** Find the farthest Right
      With Control.Properties
         If nControlRight = 0 Then
            nControlRight = .Item("Left")
         Else
            If .Item("Left") > nControlRight Then nControlRight = .Item("Left")
         End If
      End With
   Next Control

   ' *** For each control on the active form...
   For Each Control In AllControls
      ' *** ...align control
      With Control.Properties
         .Item("Left") = nControlRight
      End With
   Next Control

EXIT_AlignControlsRight:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_AlignControlsRight:
   Resume EXIT_AlignControlsRight

End Sub

Public Sub AlignControlsTop()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/25/2001
   ' * Time             : 13:24
   ' * Module Name      : Aligner_Module
   ' * Module Filename  : AlignerControls.bas
   ' * Procedure Name   : AlignControlsTop
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_AlignControlsTop

   Dim AllControls
   Dim Control          As Object
   Dim nControlTop      As Long

   nControlTop = 0

   ' *** Get collection containing all selected controls on form
   Set AllControls = VBInstance.SelectedVBComponent.Designer.SelectedVBControls

   ' *** For each control on the active form...
   For Each Control In AllControls
      ' *** Find the farthest Top
      With Control.Properties
         If nControlTop = 0 Then
            nControlTop = .Item("Top")
         Else
            If .Item("Top") < nControlTop Then nControlTop = .Item("Top")
         End If
      End With
   Next Control

   ' *** For each control on the active form...
   For Each Control In AllControls
      ' *** ...align control
      With Control.Properties
         .Item("Top") = nControlTop
      End With
   Next Control

EXIT_AlignControlsTop:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_AlignControlsTop:
   Resume EXIT_AlignControlsTop

End Sub

Public Sub AlignControlsBottom()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/25/2001
   ' * Time             : 13:24
   ' * Module Name      : Aligner_Module
   ' * Module Filename  : AlignerControls.bas
   ' * Procedure Name   : AlignControlsBottom
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_AlignControlsBottom

   Dim AllControls
   Dim Control          As Object
   Dim nControlBottom   As Long

   nControlBottom = 0

   ' *** Get collection containing all selected controls on form
   Set AllControls = VBInstance.SelectedVBComponent.Designer.SelectedVBControls

   ' *** For each control on the active form...
   For Each Control In AllControls
      ' *** Find the farthest Bottom
      With Control.Properties
         If nControlBottom = 0 Then
            nControlBottom = .Item("Top")
         Else
            If .Item("Top") > nControlBottom Then nControlBottom = .Item("Top")
         End If
      End With
   Next Control

   ' *** For each control on the active form...
   For Each Control In AllControls
      ' *** ...align control
      With Control.Properties
         .Item("Top") = nControlBottom
      End With
   Next Control

EXIT_AlignControlsBottom:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_AlignControlsBottom:
   Resume EXIT_AlignControlsBottom

End Sub
