Attribute VB_Name = "ListBox_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 04/11/1999
' * Time             : 16:53
' * Module Name      : ListBox_Module
' * Module Filename  : ListBox.bas
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private Const LB_GETHORIZONTALEXTENT = &H193
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const DT_CALCRECT = &H400
Private Const SM_CXVSCROLL = 2

Private Type RECT
   left                 As Long
   top                  As Long
   right                As Long
   bottom               As Long
End Type

Private Declare Function DrawText Lib "user32" _
   Alias "DrawTextA" _
   (ByVal hdc As Long, _
   ByVal lpStr As String, _
   ByVal nCount As Long, _
   lpRect As RECT, ByVal _
   wFormat As Long) As Long

Declare Function GetSystemMetrics Lib "user32" _
   (ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib _
   "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

' *** Declares for listbox
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
   x                    As Long
   y                    As Long
End Type

Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2

Private Const WM_KEYDOWN = &H100
Private Const WM_USER = &H400
Private Const LB_SETTABSTOPS = WM_USER + 19

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETITEMHEIGHT = &H154

Function CBLFind(CBL As Control, strToFind As String, nFirst As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : ListBox_Module
   ' * Module Filename  : ListBox.bas
   ' * Procedure Name   : CBLFind
   ' * Parameters       :
   ' *                    CBL As Control
   ' *                    strToFind As String
   ' *                    nFirst As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Recherche un string dans un combo ou une CBLox ***
   ' *** Renvoie l'index sélectionné ***

   Dim nIndex           As Long
   Dim nCount           As Long
   Dim nLength          As Long
   Dim szItem           As String

   nCount = CBL.ListCount
   nLength = Len(strToFind)

   CBLFind = -1
   strToFind = UCase$(strToFind)

   If (nCount = -1) Or (nLength = 0) Then
      Exit Function
   End If

   If (nFirst >= 0) Then
      For nIndex = nFirst To nCount
         szItem = UCase$(CBL.List(nIndex))
         If Len(szItem) = nLength Then
            If szItem = strToFind Then
               CBLFind = nIndex
               Exit For
            End If
         End If
      Next
   Else
      For nIndex = Abs(nFirst) To 0 Step -1
         szItem = UCase$(CBL.List(nIndex))
         If Len(szItem) = nLength Then
            If szItem = strToFind Then
               CBLFind = nIndex
               Exit For
            End If
         End If
      Next
   End If

End Function

Function CBLItemBackColor(CBL As Control, nColor As Long, nFirst As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : ListBox_Module
   ' * Module Filename  : ListBox.bas
   ' * Procedure Name   : CBLItemBackColor
   ' * Parameters       :
   ' *                    CBL As Control
   ' *                    nColor As Long
   ' *                    nFirst As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim nIndex           As Long
   Dim nCount           As Long

   nCount = CBL.ListCount

   CBLItemBackColor = -1

   If (nCount = -1) Then
      Exit Function
   End If

   For nIndex = 0 To nCount
      If CBL.ItemBackColor(nIndex) = nColor Then
         CBLItemBackColor = nIndex
         Exit For
      End If
   Next

End Function

Function CBLFindCode(CBL As Control, code As Long) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : ListBox_Module
   ' * Module Filename  : ListBox.bas
   ' * Procedure Name   : CBLFindCode
   ' * Parameters       :
   ' *                    CBL As Control
   ' *                    code As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** Recherche le ItemData dans un combo ***
   ' *** Renvoie l'index sélectionné ***

   Dim nIndex           As Integer, nCount As Integer
   Dim Item             As Long

   nCount = CBL.ListCount - 1

   CBLFindCode = -1
   If (nCount = -1) Then
      CBLFindCode = -1
      Exit Function
   End If

   For nIndex = 0 To nCount
      Item = CBL.ItemData(nIndex)
      If code = Item Then
         CBLFindCode = nIndex
         Exit For
      End If
   Next

End Function

Private Sub SetComboDropdownHeight(frm As Form, combo As ComboBox, nItems As Integer)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 04/11/1999
   ' * Time             : 16:53
   ' * Module Name      : ListBox_Module
   ' * Module Filename  : ListBox.bas
   ' * Procedure Name   : SetComboDropdownHeight
   ' * Parameters       :
   ' *                    frm As Form
   ' *                    combo As ComboBox
   ' *                    nItems As Integer
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' *** The Visual Basic combo box dropdown - unlike its C counterpart -
   ' *** is limited to displaying only eight items.
   ' *** This function changes the dropdown height to any number greater than eight,
   ' *** which is its apparent minimum under VB.

   Dim pt               As POINTAPI
   Dim rc               As RECT
   Dim cWidth           As Long
   Dim newHeight        As Long
   Dim oldScaleMode     As Long
   Dim itemHeight       As Long

   ' *** Save the current form scalemode, then switch to pixels
   oldScaleMode = frm.ScaleMode
   frm.ScaleMode = vbPixels

   ' *** The width of the combo, used below
   cWidth = combo.Width

   ' *** Get the system height of a single combo box list item
   itemHeight = SendMessageLong(combo.hWnd, CB_GETITEMHEIGHT, 0, 0)

   ' *** Calculate the new height of the combo box. This
   ' *** is the number of items times the item height
   ' *** plus two. The 'plus two' is required to allow
   ' *** the calculations to take into account the size
   ' *** of the edit portion of the combo as it relates
   ' *** to item height. In other words, even if the
   ' *** combo is only 21 px high (315 twips), if the
   ' *** item height is 13 px per item (as it is with
   ' *** small fonts), we need to use two items to  'achieve this height.
   newHeight = itemHeight * (nItems + 2)

   ' *** Get the co-ordinates of the combo box  'relative to the screen
   Call GetWindowRect(combo.hWnd, rc)
   pt.x = rc.left
   pt.y = rc.top

   ' *** Then translate into co-ordinates  relative to the form.
   Call ScreenToClient(frm.hWnd, pt)

   ' *** Using the values returned and set above, call MoveWindow to reposition the combo box
   Call MoveWindow(combo.hWnd, pt.x, pt.y, combo.Width, newHeight, True)

   ' *** Its done, so show the new combo height
   Call SendMessageLong(combo.hWnd, CB_SHOWDROPDOWN, True, 0)

   ' *** Restore the original form scalemode  'before leaving
   frm.ScaleMode = oldScaleMode

End Sub

Public Function AddItemToList(frm As Form, Ctl As Control, sNewItem As String, Optional dwNewItemData As Variant) As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 14/10/99
   ' * Time             : 14:50
   ' * Module Name      : ListBox_Module
   ' * Module Filename  : ListBox.bas
   ' * Procedure Name   : AddItemToList
   ' * Parameters       :
   ' *                    frm As Form
   ' *                    ctl As Control
   ' *                    sNewItem As String
   ' *                    Optional dwNewItemData As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Dim c                As Long
   Dim rcText           As RECT
   Dim NewWidth         As Long
   Dim currWidth        As Long
   Dim sysScrollWidth   As Long

   Dim tmpFontName      As String
   Dim tmpFontSize      As Long
   Dim tmpFontBold      As Boolean

   'get the current width used
   If Ctl.Tag <> "" Then
      currWidth = CLng(Ctl.Tag)
   End If

   'determine the needed width for the new item
   'save the font properties to tmp variables
   tmpFontName = frm.Font.Name
   tmpFontSize = frm.Font.Size
   tmpFontBold = frm.Font.Bold

   frm.Font.Name = Ctl.Font.Name
   frm.Font.Size = Ctl.Font.Size
   frm.Font.Bold = Ctl.Font.Bold

   'get the width of the system scrollbar
   sysScrollWidth = GetSystemMetrics(SM_CXVSCROLL)

   'use DrawText/DT_CALCRECT to determine item length
   Call DrawText(frm.hdc, sNewItem, -1&, rcText, DT_CALCRECT)
   NewWidth = rcText.right + sysScrollWidth

   'if this is wider than the current setting,
   'tweak the list and save the new horizontal
   'extent to the tag property
   If NewWidth > currWidth Then

      Call SendMessage(Ctl.hWnd, LB_SETHORIZONTALEXTENT, NewWidth, ByVal 0&)
      Ctl.Tag = NewWidth

   End If

   'restore the form font properties
   frm.Font.Name = tmpFontName
   frm.Font.Bold = tmpFontBold
   frm.Font.Size = tmpFontSize

   'add the items to the control, and
   'add the ItemData if supplied
   Ctl.AddItem sNewItem

   If Not IsMissing(dwNewItemData) Then
      If IsNumeric(dwNewItemData) Then
         Ctl.ItemData(Ctl.NewIndex) = dwNewItemData
      End If
   End If

   'return the new index as the function result
   AddItemToList = Ctl.NewIndex

End Function

