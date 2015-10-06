VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_ColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Programmer Name  : Klaus H. Probst
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 14/02/2000
' * Time             : 11:26
' * Module Name      : class_ColorDialog
' * Module Filename  :
' **********************************************************************
' * Comments         :
' * Procedures for accessing the Color Common Dialog services through the Windows API
' * without the overhead of the VB Common Dialog OCX control.
' *
' *  Notes:
' *
' *     The .Color property is the standard entry point to this class.
' *     After you've initialized all relevant property members, just assign
' *     the value of the Color property to a variable. The class will open
' *     the dialog and return the selected color, or -1 if the user canceled
' *     or there was an error.
' *     The following illustrates how you might use the class:
' *
' *         Dim hDlg As CColorDialog
' *         Dim pColor As Long
' *
' *         Set hDlg = New CColorDialog
' *
' *         pColor = vbYellow
' *
' *         With hDlg
' *
' *             .hwndOwner = Me.hWnd
' *             .Flags = cdlgCCDefault
' *             .Color = pColor
' *             pColor = .Color
' *
' *         End With
' *
' *         If pColor >=0 Then Command1.BackColor = pColor
' *
' *         Set hDlg = Nothing
' *
' *     You can use the CustomColors[] property to set and return the values of
' *     the 16 colors that can be customized by the user while in the dialog. The
' *     class initializes these colors to white by default.
' *     The class maps OLE_COLORs to standard RGB colors, so it's possible to pass
' *     values such as the VB system color constants. OLE_COLORs are just RGB values
' *     with some support built in for system colors and so on. You can pass either
' *     OLE_COLORs or Long values to the dialog's properties.
' *     NOTE: When the user cancels and the dialog the Color property will return -1.
' *     If you try to assign that to, say, Form.BackColor, you will get an
' *     "Invalid property value" error. However, there's no other way
' *     to distinguish from a canceled dialog or an error, since returning
' *     zero (which won't cause an error) would be returning Black, which
' *     is a valid color. So, make sure you check for the return value before
' *     assigning it to any objects.
' *
' **********************************************************************

Option Explicit
DefLng A-Z

Private Type udtCHOOSECOLOR
   lStructSize          As Long
   hWndOwner            As Long
   hInstance            As Long
   rgbResult            As Long
   lpCustColors         As Long
   Flags                As Long
   lCustData            As Long
   lpfnHook             As Long
   lpTemplateName       As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As udtCHOOSECOLOR) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' *** ChooseColor() constants
Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SHOWHELP = &H8
'Private Const CC_ENABLEHOOK = &H10
'Private Const CC_ENABLETEMPLATE = &H20
'Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_ANYCOLOR = &H100

Public Enum ColorDialogFlags

   cdlgCCUseInitialColor = CC_RGBINIT
   cdlgCCHideCustomColors = CC_PREVENTFULLOPEN
   cdlgCCFullOpenOnLoad = CC_FULLOPEN
   cdlgCCShowHelp = CC_SHOWHELP
   cdlgAnyColor = CC_ANYCOLOR          ' *** any color is game, even dithered ones
   cdlgCCReturnOnlySolidColors = CC_SOLIDCOLOR
   cdlgCCDefault = CC_ANYCOLOR Or CC_RGBINIT

End Enum

Public Enum CommonDialogColorErrors
   CDERR_CANCELED = 0&
   CDERR_DIALOGFAILURE = &HFFFF
   CDERR_GENERALCODES = &H0
   CDERR_STRUCTSIZE = &H1
   CDERR_INITIALIZATION = &H2
   'CDERR_NOTEMPLATE = &H3
   'CDERR_NOHINSTANCE = &H4
   CDERR_LOADSTRFAILURE = &H5
   CDERR_FINDRESFAILURE = &H6
   CDERR_LOADRESFAILURE = &H7
   CDERR_LOCKRESFAILURE = &H8
   CDERR_MEMALLOCFAILURE = &H9
   CDERR_MEMLOCKFAILURE = &HA
   'CDERR_NOHOOK = &HB
   CDERR_REGISTERMSGFAIL = &HC
   CCERR_CHOOSECOLORCODES = &H5000     ' *** no color error codes...

End Enum

Private m_lpRGB(0 To 16) As Long        ' *** custom color array
Private m_lpcc          As udtCHOOSECOLOR        ' *** class-wide

' *** if the call to ChooseColor() fails, we save this here
Private m_enlastError   As CommonDialogColorErrors

Public Function TranslateColor(color As OLE_COLOR) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : TranslateColor
   ' * Parameters       :
   ' *                    Color As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' * Converts an OLE_COLOR value to a standard RGB value.
   ' *
   ' **********************************************************************

   Dim pColor           As Long

   Call OleTranslateColor(color, 0, pColor)
   TranslateColor = pColor

End Function

Public Property Let hWndOwner(ByVal vData As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : hWndOwner
   ' * Parameters       :
   ' *                    ByVal vData As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_lpcc.hWndOwner = vData

End Property

Public Property Get hWndOwner() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : hWndOwner
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Sets/returns the form or window that will act
   ' *                  as "parent" of the dialog so that it is shown
   ' * modally. If you leave this alone, the class will
   ' * attempt to use Screen.ActiveForm to obtain a valid
   ' * hWnd. If you don't want this, assign a -1 to this
   ' * property to force a non-modal display.
   ' *
   ' **********************************************************************

   hWndOwner = m_lpcc.hWndOwner

End Property

Public Property Get Flags() As ColorDialogFlags
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : Flags
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Flags = m_lpcc.Flags

End Property

Public Property Get color() As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : Color
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' * Standard entry point. After setting all your options,
   ' * make a variable assignment to this property and the
   ' * dialog will be invoked.
   ' * NOTE: If the user cancels and this property returns -1, and you
   ' * try to assign that to, say, Form.BackColor, you will get an
   ' * "Invalid property value" error. However, there's no other way
   ' * to distinguish from a canceled dialog or an error, since returning
   ' * zero (which won't cause an error) would be returning Black, which
   ' * is a valid color. So, make sure you check for the return value before
   ' * assigning it to any objects.
   ' *
   ' **********************************************************************

   If ChooseColorEx() = False Then
      m_lpcc.rgbResult = -1

   End If

   color = m_lpcc.rgbResult

End Property

Private Function ChooseColorEx() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : ChooseColorEx
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' * This is the actual wrapper to the API call that gets
   ' * invoked when you retreive the value of the Color[] property.
   ' *
   ' **********************************************************************

   Dim lngResult        As Long

   ' *** Initialize to this
   m_enlastError = CDERR_CANCELED

   With m_lpcc

      If .Flags = 0 Then .Flags = cdlgCCDefault
      If .hWndOwner = 0 Then .hWndOwner = GetOwnerWindow()    ' *** Try to get a valid owner window
      .lpCustColors = VarPtr(m_lpRGB(0))
      .lStructSize = Len(m_lpcc)

   End With

   lngResult = ChooseColor(m_lpcc)

   If lngResult = 1 Then
      ChooseColorEx = True

   Else
      ' *** Oops... clean everything up
      m_enlastError = CommDlgExtendedError()
      Call ZeroMemory(m_lpcc, Len(m_lpcc))
      ChooseColorEx = False

   End If

End Function

Public Property Let color(ByVal vData As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : Color
   ' * Parameters       :
   ' *                    ByVal vData As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_lpcc.rgbResult = TranslateColor(vData)

End Property

Public Property Get CustomColors(ByVal index As Integer) As OLE_COLOR
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : CustomColors
   ' * Parameters       :
   ' *                    ByVal Index As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Hello... :-)
   Debug.Assert (index >= 0 And index <= 16)

   On Error Resume Next

   CustomColors = m_lpRGB(index)

End Property

Public Property Let CustomColors(ByVal index As Integer, ByVal vData As OLE_COLOR)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : CustomColors
   ' * Parameters       :
   ' *                    ByVal Index As Integer
   ' *                    ByVal vData As OLE_COLOR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Hello... :-)
   Debug.Assert (index >= 0 And index <= 16)

   On Error Resume Next

   m_lpRGB(index) = TranslateColor(vData)

End Property

Public Property Let Flags(ByVal vData As ColorDialogFlags)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : Flags
   ' * Parameters       :
   ' *                    ByVal vData As ColorDialogFlags
   ' **********************************************************************
   ' * Comments         :
   ' *  Sets/returns the behavior flags for the dialog.
   ' *
   ' **********************************************************************

   m_lpcc.Flags = vData

End Property

Private Function GetOwnerWindow() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : GetOwnerWindow
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * This is an attempt to get the owner window so that
   ' * the dialog is shown modally. If you absolutely must
   ' * show the dialog non-modally, set the hWndOwner property
   ' * to -1. The class checks for this and leaves it alone.
   ' *
   ' **********************************************************************

   On Error Resume Next

   GetOwnerWindow = Screen.ActiveForm.hWnd

End Function

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim a                As Integer

   ' *** initialize all colors to white. Otherwise
   ' *** they show up as black (&H0), of course. And that's
   ' *** rather ugly.
   For a = 0 To 16
      m_lpRGB(a) = vbWhite
   Next

End Sub

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Klaus H. Probst
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/02/2000
   ' * Time             : 11:26
   ' * Module Name      : class_ColorDialog
   ' * Module Filename  :
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Erase m_lpRGB

End Sub
