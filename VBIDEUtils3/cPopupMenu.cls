VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Creates unlimited new popup menus using the API
Implements ISubclass

Private m_cNCM          As New cNCMetrics

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
' The messages we will intercept:
Private Const WM_MENUSELECT = &H11F
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const WM_COMMAND = &H111
Private Const WM_MENUCHAR = &H120
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_WININICHANGE = &H1A
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212

' Array of menu items
Private m_tMI()         As tMenuItem
Private m_iMenuCount    As Long

' Stored menus:
Private m_cStoredMenu() As cStoreMenu
Private m_iStoreCount   As Long

' Display stuff, used to draw the control and also
' to evaluate menu font item sizes:
Private m_HDC           As Long
Private m_hBMPDither    As Long
Private m_bUseDither    As Boolean
Private m_hFntOld       As Long
Private m_bGotFont      As Boolean
' Handle to image list for drawing icons:
Private m_hIml          As Long
' Where to get a tick icon for checked stuff (or -1 to use Win default):
Private m_lTickIconIndex As Long
' Where to get a option button icon for checked stuff (or -1 to use Win default)
Private m_lOptionIconIndex As Long

' hWNd of owner:
Private m_hWndOwner     As Long
Private m_hWndAttached  As Long
' Height of a menu item:
Private m_lMenuItemHeight As Long

' Bitmap to tile into background of menu:
Private m_hDCBack       As Long
Private m_hBmpOld       As Long
Private m_hBmp          As Long
Private m_lBitmapW      As Long
Private m_lBitmapH      As Long

' Sub menus:
Private m_lSubMenuCount As Long
Private m_hSubMenus()   As Long
' Next id to choose for a menu item:
Private m_lLastMaxId    As Long

Private m_bGradientHighlight As Boolean

Private m_sTag          As String
Private m_bDrawHeadersAsSeparators As Boolean

Public Enum ECNMHeaderStyle
   ecnmHeaderCaptionBar = 0
   ecnmHeaderSeparator = 1
End Enum

' Events:
Public Event Click(ItemNumber As Long)
Public Event ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
Public Event MenuExit()
Public Event InitPopupMenu(ParentItemNumber As Long)
Public Event DrawItem(ByVal HDC As Long, ByVal lMenuIndex As Long, ByRef lLeft As Long, ByRef lTop As Long, ByRef lRight As Long, ByRef lBottom As Long, ByVal bSelected As Boolean, ByVal bChecked As Boolean, ByVal bDisabled As Boolean, bDoDefault As Boolean)
Public Event MeasureItem(ByVal lMenuIndex As Long, ByRef lWidth As Long, ByRef lHeight As Long)

Public Property Set BackgroundPicture( _
   ByRef sPic As StdPicture _
   )
   Dim tBm              As Bitmap
   Dim lHDC             As Long
   Dim lHDCTemp         As Long
   Dim lHBmpTemp        As Long
   Dim bBackOk          As Boolean

   ClearBackgroundPicture
   GetObjectAPI sPic.Handle, Len(tBm), tBm
   m_lBitmapW = tBm.bmWidth
   m_lBitmapH = tBm.bmHeight
   If m_lBitmapW > 0 And m_lBitmapH > 0 Then
      lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      If lHDC <> 0 Then
         lHDCTemp = CreateCompatibleDC(lHDC)
         If lHDCTemp <> 0 Then
            lHBmpTemp = SelectObject(lHDCTemp, sPic.Handle)
            If lHBmpTemp <> 0 Then
               m_hDCBack = CreateCompatibleDC(lHDC)
               If m_hDCBack <> 0 Then
                  m_hBmp = CreateCompatibleBitmap(lHDC, m_lBitmapW, m_lBitmapH)
                  If m_hBmp <> 0 Then
                     m_hBmpOld = SelectObject(m_hDCBack, m_hBmp)
                     If m_hBmpOld <> 0 Then
                        BitBlt m_hDCBack, 0, 0, m_lBitmapW, m_lBitmapH, lHDCTemp, 0, 0, vbSrcCopy
                        bBackOk = True
                     End If
                  End If
               End If
               SelectObject lHDCTemp, lHBmpTemp
            End If
            DeleteDC lHDCTemp
         End If
         DeleteDC lHDC
      End If
   End If
   If Not bBackOk Then
      ClearBackgroundPicture
   End If

End Property
Public Sub ClearBackgroundPicture()
   If m_hBmpOld <> 0 Then
      SelectObject m_HDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If m_hBmp <> 0 Then
      DeleteObject m_hBmp
   End If
   If m_hDCBack <> 0 Then
      DeleteDC m_hDCBack
   End If
   m_hDCBack = 0
   m_lBitmapW = 0
   m_lBitmapH = 0
End Sub

Friend Function AcceleratorPress(ByVal nKeyCode As KeyCodeConstants, ByVal wMask As ShiftConstants) As Boolean
   Dim i                As Long
   For i = 1 To m_iMenuCount
      If Not m_tMI(i).iShortCutShiftKey = 0 Then
         Debug.Print "Accel Press..."; nKeyCode, wMask, m_tMI(i).iShortCutShiftKey, m_tMI(i).iShortCutShiftMask
         If m_tMI(i).iShortCutShiftMask = wMask Then
            If m_tMI(i).iShortCutShiftKey = nKeyCode Then
               ' Yo!
               RaiseEvent Click(i)
               AcceleratorPress = True
               Exit For
            End If
         End If
      End If
   Next i
End Function
Public Property Get IDForItem(ByVal lIndex As Long) As Long
   If lIndex > 0 And lIndex <= m_iMenuCount Then
      IDForItem = m_tMI(lIndex).lActualID
   End If
End Property
Public Property Get ItemForID(ByVal wID As Long) As Long
   Dim lIndex           As Long
   For lIndex = 1 To m_iMenuCount
      If m_tMI(lIndex).lActualID = wID Then
         ItemForID = lIndex
         Exit For
      End If
   Next lIndex
End Property

Public Sub EmulateMenuClick(ByVal wID As Long)
   Dim lIndex           As Long
   For lIndex = 1 To m_iMenuCount
      If m_tMI(lIndex).lActualID = wID Then
         RaiseClickEvent wID
         Exit For
      End If
   Next lIndex
End Sub

Public Property Get GradientHighlight() As Boolean
   GradientHighlight = m_bGradientHighlight
End Property
Public Property Let GradientHighlight(ByVal bState As Boolean)
   m_bGradientHighlight = bState
End Property

Public Property Get HeaderStyle() As ECNMHeaderStyle
   If (m_bDrawHeadersAsSeparators) Then
      HeaderStyle = ecnmHeaderSeparator
   Else
      HeaderStyle = ecnmHeaderCaptionBar
   End If
End Property
Public Property Let HeaderStyle(ByVal eStyle As ECNMHeaderStyle)
   If (eStyle = ecnmHeaderCaptionBar) Then
      m_bDrawHeadersAsSeparators = False
   Else
      m_bDrawHeadersAsSeparators = True
   End If
End Property

Public Property Get Count() As Long
   Count = m_iMenuCount
End Property

Public Property Get HighlightCheckedItems() As Boolean
   HighlightCheckedItems = m_bUseDither
End Property
Public Property Let HighlightCheckedItems(ByVal bState As Boolean)
   m_bUseDither = bState
   If (bState) Then
      ' Get the dither bitmap from the resource file:
      m_hBMPDither = LoadImageByNum(App.hInstance, 49, IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS)
   Else
      If (m_hBMPDither <> 0) Then
         DeleteObject m_hBMPDither
         m_hBMPDither = 0
      End If
   End If
End Property
Public Property Get Tag() As String
   Tag = m_sTag
End Property
Public Property Let Tag(ByVal sTag As String)
   m_sTag = sTag
End Property
Public Sub Store(ByVal sKey As String)
   Dim lIndex           As Long
   ' Save the menu under the key sKey:
   lIndex = plStored(sKey)
   If (lIndex = 0) Then
      ' We need a new item
      m_iStoreCount = m_iStoreCount + 1
      ReDim Preserve m_cStoredMenu(1 To m_iStoreCount) As cStoreMenu
      Set m_cStoredMenu(m_iStoreCount) = New cStoreMenu
      lIndex = m_iStoreCount
   End If
   With m_cStoredMenu(lIndex)
      .Key = sKey
      .Store m_tMI(), m_iMenuCount
   End With
End Sub
Public Sub Restore(ByVal sKey As String)
   Dim lIndex           As Long
   ' Restore the menu from the key sKey:
   lIndex = plStored(sKey)
   If (lIndex > 0) Then
      ' Clear any menu:
      Clear
      ' Restore from storage:
      m_cStoredMenu(lIndex).Restore Me
   Else
      ' Error
      Debug.Print "Failed to restore..."
   End If
End Sub
Public Function StoreToFile( _
   Optional ByVal iFIle As Long = -1, _
   Optional ByVal sFile As String = "" _
   )
   Dim i                As Long
   Dim iUseFile         As Integer

   ' Really we should be de/serialising to a PStream via the IStream
   ' interface (i.e. PropertyBag)

   If iFIle < 1 And sFile = "" Then
      Err.Raise 9, App.EXEName & ".cPopupMenu", "Invalid call to RestoreFromFile; specify file name or handle."
      Exit Function
   End If

   If m_iStoreCount > 0 Then
      On Error Resume Next
      Kill sFile
      Err.Clear

      On Error GoTo ErrorHandler

      If iFIle = -1 Then
         iUseFile = FreeFile
         Open sFile For Binary Access Write Lock Read As #iUseFile
      Else
         iUseFile = iFIle
      End If
      Put #iUseFile, , "vbalNewMenu"
      Put #iUseFile, , m_iStoreCount
      For i = 1 To m_iStoreCount
         m_cStoredMenu(i).Serialise iUseFile
      Next i

      If iFIle = -1 Then
         Close #iUseFile
      End If
      iUseFile = 0

   Else
      Err.Raise 9, App.EXEName & ".cPopupMenu", "No stored menus to save"
   End If

   Exit Function

ErrorHandler:
   Dim lErr             As Long, sErr As String
   lErr = Err.Number: sErr = Err.Description
   If iUseFile > 0 And iFIle = -1 Then
      Close #iUseFile
      iUseFile = 0
   End If
   Err.Raise lErr, App.EXEName & ".cPopupMenu", sErr
   Exit Function

End Function
Public Function RestoreFromFile( _
   Optional ByVal iFIle As Long = -1, _
   Optional ByVal sFile As String = "" _
   )
   Dim i                As Long
   Dim iUseFile         As Long
   Dim sBuf             As String
   Dim lCount           As Long
   Dim bFail            As Boolean
   Dim sError           As String

   ' Really we should be de/serialising to a PStream via the IStream
   ' interface (i.e. PropertyBag)

   If iFIle < 1 And sFile = "" Then
      Err.Raise 9, App.EXEName & ".cPopupMenu", "Invalid call to RestoreFromFile; specify file name or handle."
      Exit Function
   End If

   On Error GoTo ErrorHandler

   If iFIle = -1 Then
      iUseFile = FreeFile
      Open sFile For Binary Access Read Lock Write As #iUseFile
   Else
      iUseFile = iFIle
   End If

   sBuf = Space$(11)
   Get #iUseFile, , sBuf
   If sBuf = "vbalNewMenu" Then
      Get #iUseFile, , lCount
      If lCount > 0 Then

         Clear
         m_iStoreCount = lCount
         ReDim m_cStoredMenu(1 To m_iStoreCount) As cStoreMenu

         For i = 1 To m_iStoreCount
            Set m_cStoredMenu(i) = New cStoreMenu
            If Not (m_cStoredMenu(i).Deserialise(iUseFile)) Then
               bFail = True
               sError = m_cStoredMenu(i).Error
               Exit For
            End If
         Next i
      End If
      If iFIle = -1 Then
         Close #iUseFile
      End If
      iUseFile = 0
      If bFail Then
         Err.Raise 9, App.EXEName & ".cPopupMenu", sError
      Else
         RestoreFromFile = True
      End If
   Else
      If iFIle = -1 Then
         Close #iUseFile
      End If
      iUseFile = 0
      Err.Raise 9, App.EXEName & ".cPopupMenu", "Not a cNewMenu file stream."
   End If

   Exit Function

ErrorHandler:
   Dim lErr             As Long, sErr As String
   lErr = Err.Number: sErr = Err.Description
   If iUseFile > 0 And iFIle = -1 Then
      Close #iUseFile
      iUseFile = 0
   End If
   Err.Raise lErr, App.EXEName & ".cPopupMenu", sErr
   Exit Function
   Resume 0
End Function
Private Property Get plStored(ByVal sKey As String) As Long
   Dim i                As Long
   For i = 1 To m_iStoreCount
      If (m_cStoredMenu(i).Key = sKey) Then
         plStored = i
         Exit For
      End If
   Next i
End Property
Public Property Get hWndOwner() As Long
   hWndOwner = m_hWndOwner
End Property
Public Property Let hWndOwner(ByVal hWndA As Long)
   Dim lHDC             As Long

   ' Clear up:
   Clear
   ' Clear DC:
   If (m_HDC <> 0) Then
      If (m_hFntOld <> 0) Then
         SelectObject m_HDC, m_hFntOld
      End If
      DeleteObject m_HDC
   End If

   ' Set for new owner:
   m_hWndOwner = hWndA
   lHDC = GetDC(hWndA)
   m_HDC = CreateCompatibleDC(lHDC)
   ReleaseDC m_hWndOwner, lHDC
   ' Select the menu font into it:
   pSelectMenuFont

End Property

Public Property Let ImageList( _
   ByRef vImageList As Variant _
   )
   If (VarType(vImageList) = vbLong) Then
      ' Assume a handle to an image list:
      m_hIml = vImageList
   ElseIf (VarType(vImageList) = vbObject) Then
      ' Assume a VB image list:
      On Error Resume Next
      ' Get the image list initialised..
      vImageList.ListImages(1).Draw 0, 0, 0, 1
      m_hIml = vImageList.hImageList
      If (Err.Number = 0) Then
         ' OK
      Else
         Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
      End If
      On Error GoTo 0
   End If
End Property

Public Function AddItem( _
   ByVal sCaption As String, _
   Optional ByVal sHelptext As String = "", _
   Optional ByVal lItemData As Long = 0, _
   Optional ByVal lParentIndex As Long = 0, _
   Optional ByVal lIconIndex As Long = -1, _
   Optional ByVal bChecked As Boolean = False, _
   Optional ByVal bEnabled As Boolean = True, _
   Optional ByVal sKey As String = "" _
   ) As Long
   Dim lID              As Long
   m_iMenuCount = m_iMenuCount + 1
   ReDim Preserve m_tMI(1 To m_iMenuCount) As tMenuItem
   lID = plGetNewID()
   With m_tMI(m_iMenuCount)
      .lID = lID
      .lActualID = lID
      pSetMenuCaption m_iMenuCount, sCaption, (sCaption = "-")
      .sAccelerator = psExtractAccelerator(sCaption)
      .sHelptext = sHelptext
      .lIconIndex = lIconIndex
      If (lParentIndex <> 0) Then
         .lParentId = m_tMI(lParentIndex).lActualID
      End If
      .lParentIndex = lParentIndex
      .lItemData = lItemData
      .bChecked = bChecked
      .bEnabled = bEnabled
      .bCreated = True
      .sKey = sKey
   End With
   pAddNewMenuItem m_tMI(m_iMenuCount)
   If (m_tMI(m_iMenuCount).bTitle) Then
      Header(m_iMenuCount) = True
      m_tMI(m_iMenuCount).lHeight = m_lMenuItemHeight * 3 \ 4
   Else
      If (m_tMI(m_iMenuCount).sCaption = "-") Then
         m_tMI(m_iMenuCount).lHeight = 6
      Else
         m_tMI(m_iMenuCount).lHeight = m_lMenuItemHeight
      End If
   End If
   ItemData(m_iMenuCount) = lItemData
   AddItem = m_iMenuCount

End Function
Public Function InsertItem( _
   ByVal sCaption As String, _
   ByVal vKeyBefore As Variant, _
   Optional ByVal sHelptext As String = "", _
   Optional ByVal lItemData As Long = 0, _
   Optional ByVal lIconIndex As Long = -1, _
   Optional ByVal bChecked As Boolean = False, _
   Optional ByVal bEnabled As Boolean = True, _
   Optional ByVal sKey As String = "" _
   ) As Long
   Dim lIndexBefore     As Long
   Dim lID              As Long
   'Inserts an item into a menu:
   lIndexBefore = plMenuIndex(vKeyBefore)
   If (lIndexBefore > 0) Then
      m_iMenuCount = m_iMenuCount + 1
      ReDim Preserve m_tMI(1 To m_iMenuCount) As tMenuItem
      lID = plGetNewID()
      With m_tMI(m_iMenuCount)
         .lID = lID
         .lActualID = lID
         pSetMenuCaption m_iMenuCount, sCaption, (sCaption = "-")
         .sAccelerator = psExtractAccelerator(sCaption)
         .sHelptext = sHelptext
         .lIconIndex = lIconIndex
         .lItemData = lItemData
         .bChecked = bChecked
         .bEnabled = bEnabled
         .bCreated = True
         .sKey = sKey
      End With
      pInsertNewMenuitem m_tMI(m_iMenuCount), lIndexBefore
      InsertItem = m_iMenuCount
   End If
End Function
Public Function ReplaceItem( _
   ByVal vKey As Variant, _
   Optional ByVal sCaption As Variant, _
   Optional ByVal sHelptext As Variant, _
   Optional ByVal lItemData As Variant, _
   Optional ByVal lIconIndex As Variant, _
   Optional ByVal bChecked As Variant, _
   Optional ByVal bEnabled As Variant _
   ) As Long
   Dim lIndex           As Long
   Dim sItems()         As String
   Dim lH()             As Long
   Dim lR               As Long
   Dim lFlags           As Long
   Dim lPosition        As Long
   Dim tMI              As MENUITEMINFO
   Dim hSubMenu         As Long

   ' Replaces a menu item with a new one.  Works
   ' around a bug with the caption property where if
   ' you changed the size of the caption the menu did
   ' not resize.  Also allows you to change the help
   ' text, item data, icon, check and enable at the
   ' same time.

   ' Check valid index:
   lIndex = plMenuIndex(vKey)
   If (lIndex > 0) Then
      If Not IsMissing(sCaption) Then
         pSetMenuCaption lIndex, sCaption, (sCaption = "-")
      End If
      If Not IsMissing(sHelptext) Then
         m_tMI(lIndex).sHelptext = sHelptext
      End If
      If Not IsMissing(lItemData) Then
         m_tMI(lIndex).lItemData = lItemData
      End If
      If Not IsMissing(lIconIndex) Then
         m_tMI(lIndex).lIconIndex = lIconIndex
      End If
      If Not IsMissing(bChecked) Then
         m_tMI(lIndex).bChecked = bChecked
      End If
      If Not IsMissing(bEnabled) Then
         m_tMI(lIndex).bEnabled = bEnabled
      End If

      pHierarchyForIndex lIndex, lH(), sItems()
      lPosition = lH(UBound(lH)) - 1
      ' Check if there is a sub menu:
      tMI.cbSize = Len(tMI)
      tMI.fMask = MIIM_SUBMENU
      GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, 0, tMI
      hSubMenu = tMI.hSubMenu
      ' Remove the menu item:
      lR = RemoveMenu(m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND)
      ' Insert it back again at the corect position with the same ID etc:
      lFlags = plMenuFlags(lIndex)
      lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
      lR = InsertMenuByLong(m_tMI(lIndex).hMenu, lPosition, lFlags, m_tMI(lIndex).lID, m_tMI(lIndex).lID)
      If (hSubMenu <> 0) Then
         ' If we had a submenu then put that back again:
         lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
         lFlags = lFlags Or MF_POPUP
         lR = ModifyMenuByLong(m_tMI(lIndex).hMenu, m_tMI(lIndex).lID, lFlags, hSubMenu, m_tMI(lIndex).lActualID)
      End If
      If (lR = 0) Then
         Debug.Print "Failed to insert new menu item."
      End If
   End If

End Function
Public Sub RemoveItem( _
   ByVal vKey As Variant _
   )
   Dim lIndex           As Long
   lIndex = IndexForKey(vKey)
   If (lIndex > 0) Then
      pRemoveItem lIndex
   End If
End Sub
Private Sub pRemoveItem( _
   ByVal lIndex As Long _
   )
   Dim hMenusToDestroy() As Long
   Dim lCount           As Long
   Dim lDestroy         As Long
   Dim lRealCount       As Long
   Dim lR               As Long
   Dim lMaxID           As Long

   ' Remove the Item:
   lR = RemoveMenu(m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND)
   m_tMI(lIndex).bMarkTODestroy = True
   ' Loop though all the children of the item at Index and determine
   ' what there is to remove:
   pRemoveSubMenus m_tMI(lIndex).lActualID, 1, hMenusToDestroy(), lCount

   ' Destroy the menus:
   For lDestroy = 1 To lCount
      DestroyMenu hMenusToDestroy(lDestroy)
      Debug.Print "Destroyed sub-menu:" & hMenusToDestroy(lDestroy)
   Next lDestroy

   ' Now repopulate the array & sort out the indexes to remove
   ' the indexes marked for deletion:
   If (lCount > 0) Or (lR <> 0) Then
      lRealCount = 0
      For lIndex = 1 To m_iMenuCount
         If Not (m_tMI(lIndex).bMarkTODestroy) Then
            If (GetMenuItemCount(m_tMI(lIndex).lActualID) = -1) Then
               If (m_tMI(lIndex).lActualID > lMaxID) Then
                  lMaxID = m_tMI(lIndex).lActualID
               End If
            End If
            lRealCount = lRealCount + 1
            If (lRealCount <> lIndex) Then
               ' A much neater way than previously (set all the items independently!
               ' what was I thinking of)
               LSet m_tMI(lRealCount) = m_tMI(lIndex)
            End If
         End If
      Next lIndex
      ReDim Preserve m_tMI(1 To lRealCount) As tMenuItem
      m_iMenuCount = lRealCount
      If (lMaxID > m_iMenuCount) Then
         m_lLastMaxId = lMaxID
      Else
         m_lLastMaxId = m_iMenuCount
      End If
   End If
   'Debug.Print m_iMenuCount, m_lLastMaxId

End Sub
Private Sub pRemoveSubMenus( _
   ByVal lParentId As Long, _
   ByVal lStartIndex As Long, _
   ByRef hMenusToDestroy() As Long, _
   ByRef lMenuToDestroyCount As Long _
   )
   Dim lIndex           As Long

   For lIndex = 1 To m_iMenuCount
      If (m_tMI(lIndex).lParentId = lParentId) Then
         m_tMI(lIndex).bMarkTODestroy = True
         pAddToDestroyArray m_tMI(lIndex).hMenu, hMenusToDestroy(), lMenuToDestroyCount
         pRemoveSubMenus m_tMI(lIndex).lActualID, lIndex, hMenusToDestroy(), lMenuToDestroyCount
      End If
   Next lIndex
End Sub
Private Sub pAddToDestroyArray( _
   ByVal hMenu As Long, _
   ByRef hMenusToDestroy() As Long, _
   ByRef lMenuToDestroyCount As Long _
   )
   Dim lIndex           As Long
   Dim bFound           As Boolean
   For lIndex = 1 To lMenuToDestroyCount
      If (hMenusToDestroy(lIndex) = hMenu) Then
         bFound = True
         Exit For
      End If
   Next lIndex
   If Not (bFound) Then
      lMenuToDestroyCount = lMenuToDestroyCount + 1
      ReDim Preserve hMenusToDestroy(1 To lMenuToDestroyCount) As Long
      hMenusToDestroy(lMenuToDestroyCount) = hMenu
   End If
End Sub
Public Function ClearSubMenusOfItem( _
   ByVal vKey As Variant _
   ) As Long
   Dim hMenu            As Long
   Dim iMenu            As Long
   Dim lIndex           As Long
   lIndex = plMenuIndex(vKey)
   If (lIndex > 0) Then
      ' The idea is to leave just the submenu
      ' but with nothing in it:

      ' The ActualID of a sub-menu will be the
      ' handle to the submenu:
      hMenu = m_tMI(lIndex).lActualID

      ' Now remove all the items in the sub-menu,
      ' mark them for destruction and also do
      ' any sub-menus they may have:
      For iMenu = m_iMenuCount To 1 Step -1
         If (iMenu <= m_iMenuCount) Then
            If (m_tMI(iMenu).hMenu = hMenu) Then
               pRemoveItem iMenu
            End If
         End If
      Next iMenu

      For iMenu = 1 To m_iMenuCount
         If (m_tMI(iMenu).lActualID = hMenu) Then
            ClearSubMenusOfItem = iMenu
            Exit For
         End If
      Next iMenu
   End If
End Function

Private Sub pInsertNewMenuitem( _
   ByRef tMI As tMenuItem, _
   ByVal lIndexBefore As Long _
   )
   Dim hMenu            As Long
   Dim lFlags           As Long
   Dim lPosition        As Long
   Dim lR               As Long
   Dim lH()             As Long
   Dim sItems()         As String

   ' Find out where we're inserting into existing sub menu:
   hMenu = m_tMI(lIndexBefore).hMenu

   If (hMenu <> 0) Then
      pHierarchyForIndex lIndexBefore, lH(), sItems()
      lPosition = lH(UBound(lH)) - 1

      lFlags = plMenuFlags(m_iMenuCount)
      lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
      lR = InsertMenuByLong(hMenu, lPosition, lFlags, tMI.lID, tMI.lID)
      If (lR = 0) Then
         Debug.Print "Failed to insert new Menu item"
      Else
         ' Store the hMenu for this item:
         tMI.hMenu = hMenu
      End If
   End If

End Sub
Private Sub pSetMenuCaption( _
   ByVal iItem As Long, _
   ByVal sCaption As String, _
   ByVal bSeparator As Boolean _
   )
   Dim sCap             As String
   Dim sShortCut        As String
   Dim iPos             As Long

   m_tMI(iItem).sInputCaption = sCaption

   If (bSeparator) Then
      m_tMI(iItem).sCaption = "-"
   Else
      ' Check if this is a title:
      If (left$(sCaption, 1) = "-") Then
         m_tMI(iItem).bTitle = True
         sCaption = Mid$(sCaption, 2)
      End If

      ' Check if this menu item will have a menu bar break:
      pParseCaption sCaption, "|", m_tMI(iItem).bMenuBarBreak
      ' Check if this menu item will be on the same line as
      ' the last one:
      pParseCaption sCaption, "^", m_tMI(iItem).bMenuBreak

      ' Check if we have a shortcut to the menu item:
      iPos = InStr(sCaption, vbTab)
      If (iPos <> 0) Then
         sCap = left$(sCaption, (iPos - 1))
         ' Extract the ctrl key item:
         sShortCut = Mid$(sCaption, (iPos + 1))
         pParseMenuShortcut iItem, sShortCut
      Else
         sCap = sCaption
      End If
      m_tMI(iItem).sAccelerator = psExtractAccelerator(sCap)
      m_tMI(iItem).sCaption = sCap

   End If

End Sub
Private Sub pParseCaption(ByRef sCaption As String, ByVal sToken As String, ByRef bFlag As Boolean)
   Dim iPos             As Long
   Dim iPos2            As Long
   Dim sCap             As String

   iPos = InStr(sCaption, sToken)
   If (iPos <> 0) Then
      ' Check for double token (i.e. interpret as untokenised character):
      iPos2 = InStr(sCaption, sToken & sToken)
      If (iPos2 <> 0) Then
         bFlag = False
         If (iPos2 > 1) Then
            sCap = left$(sCaption, iPos - 1)
         End If
         If (iPos2 + 1 < Len(sCaption)) Then
            sCap = sCap & Mid$(sCaption, iPos2 + 1)
         End If
      Else
         bFlag = True
         If (iPos > 1) Then
            sCap = left$(sCaption, iPos - 1)
         End If
         If (iPos < Len(sCaption)) Then
            sCap = sCap & Mid$(sCaption, iPos + 1)
         End If
         sCaption = sCap
      End If
   Else
      bFlag = False
   End If
End Sub
Private Sub pParseMenuShortcut( _
   ByVal iItem As Long, _
   ByVal sShortCut As String _
   )
   Dim iPos             As Long
   Dim iNextPos         As Long
   Dim iCount           As Long
   Dim sBits()          As String
   Dim sKeyNum          As String

   sShortCut = Trim$(sShortCut)

   m_tMI(iItem).iShortCutShiftMask = 0
   m_tMI(iItem).iShortCutShiftKey = 0
   m_tMI(iItem).sShortCutDisplay = sShortCut

   If Len(sShortCut) > 0 Then
      iPos = 1
      iNextPos = InStr(iPos, sShortCut, "+")
      Do While iNextPos <> 0
         iCount = iCount + 1
         ReDim Preserve sBits(1 To iCount) As String
         sBits(iCount) = Mid$(sShortCut, iPos, iNextPos - iPos)
         iPos = iNextPos + 1
         iNextPos = InStr(iPos, sShortCut, "+")
      Loop
      If iPos <= Len(sShortCut) Then
         iCount = iCount + 1
         ReDim Preserve sBits(1 To iCount) As String
         sBits(iCount) = Mid$(sShortCut, iPos)
      End If

      ' Parse the bits:
      For iPos = 1 To iCount
         If Len(sBits(iPos)) = 1 Then
            m_tMI(iItem).iShortCutShiftKey = Asc(UCase$(sBits(iPos)))
         Else
            Select Case sBits(iPos)
               Case "Ctrl"
                  m_tMI(iItem).iShortCutShiftMask = m_tMI(iItem).iShortCutShiftMask Or vbCtrlMask
               Case "Alt"
                  m_tMI(iItem).iShortCutShiftMask = m_tMI(iItem).iShortCutShiftMask Or vbAltMask
               Case "Shift"
                  m_tMI(iItem).iShortCutShiftMask = m_tMI(iItem).iShortCutShiftMask Or vbShiftMask
               Case "Home"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyHome
               Case "End"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyEnd
               Case "Left Arrow"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyLeft
               Case "Right Arrow"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyRight
               Case "Up Arrow"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyUp
               Case "Down Arrow"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyDown
               Case "Break"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyClear
               Case "Page Up"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyPageUp
               Case "Page Up"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyPageDown
               Case "Del"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyDelete
               Case "Esc"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyEscape
               Case "Tab"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyTab
               Case "Enter"
                  m_tMI(iItem).iShortCutShiftKey = vbKeyReturn
               Case Else
                  If left$(sShortCut, 1) = "F" Then
                     sKeyNum = Mid$(sShortCut, (iPos + 1))
                     m_tMI(iItem).iShortCutShiftKey = vbKeyF1 + Val(sKeyNum) - 1
                  End If
            End Select
         End If
      Next iPos
   End If

   '    If (sShortCut <> "") Then
   '        If (InStr(sShortCut, "Ctrl")) Then
   '            m_tMI(iItem).iShortCutShiftMask = vbCtrlMask
   '            bNotFKey = True
   '        End If
   '        If (InStr(sShortCut, "Shift")) Then
   '            m_tMI(iItem).iShortCutShiftMask = m_tMI(iItem).iShortCutShiftMask Or vbShiftMask
   '            bNotFKey = True
   '        End If
   '
   '        If (bNotFKey) Then
   '            ' Find the last + and get the key:
   '            iLen = Len(sShortCut)
   '            iPos = iLen
   '            Do While Mid$(sShortCut, iPos, 1) <> "+" And iPos > 1
   '                iPos = iPos - 1
   '            Loop
   '            sKey = Mid$(sShortCut, iPos)
   '            If (Len(sKey) = 1) Then
   '                m_tMI(iItem).iShortCutShiftKey = Asc(sKey)
   '            Else
   '               ' Check for Space, Backspace, Del, Escape, Fkey, Home, End, Page Up, Page Down,
   '               ' Left Arrow, Right Arrow, Tab, Enter
   '               Select Case sKey
   '               Case "Space"
   '                  m_tMI(iItem).iShortCutShiftKey = vbKeySpace
   '
   '
   '            End If
   '        Else
   '            ' Parse the Fkey:
   '            iPos = InStr(sShortCut, "F")
   '            If (iPos <> 0) Then
   '                SkeyNum = Mid$(sShortCut, (iPos + 1))
   '                m_tMI(iItem).iShortCutShiftKey = vbKeyF1 + Val(SkeyNum) - 1
   '            End If
   '        End If
   '    End If

End Sub
Private Function pHierarchyForIndex( _
   ByVal lIndex As Long, _
   ByRef lHierarchy() As Long, _
   ByRef sItems() As String _
   ) As String
   Dim lH()             As Long
   Dim sI()             As String
   Dim lItems           As Long
   Dim hMenuSeek        As Long
   Dim lPid             As Long
   Dim bComplete        As Boolean
   Dim l                As Long
   Dim lNewIndex        As Long

   Erase lHierarchy
   Erase sItems
   ' Now determine the hierarchy for this item:
   hMenuSeek = m_tMI(lIndex).hMenu
   Do
      lItems = lItems + 1
      ReDim Preserve lH(1 To lItems) As Long
      ReDim Preserve sI(1 To lItems) As String
      lH(lItems) = plMenuPositionForIndex(hMenuSeek, lIndex)
      sI(lItems) = m_tMI(lIndex).sCaption
      lPid = m_tMI(lIndex).lParentId
      If (lPid <> 0) Then
         lNewIndex = plGetIndexForId(m_tMI(lIndex).lParentId)
         ' Debug.Print lNewIndex
         lIndex = lNewIndex
         hMenuSeek = m_tMI(lIndex).hMenu
      Else
         bComplete = True
      End If
   Loop While Not (bComplete)

   ReDim lHierarchy(1 To lItems) As Long
   ReDim sItems(1 To lItems) As String
   For l = lItems To 1 Step -1
      lHierarchy(l) = lH(lItems - l + 1)
      sItems(l) = sI(lItems - l + 1)
   Next l

End Function
Private Function plMenuPositionForIndex( _
   ByVal hMenuSeek As Long, _
   ByVal lIndex As Long _
   ) As Long
   Dim l                As Long
   Dim tMII             As MENUITEMINFO
   Dim lCount           As Long
   ' fixed bug where this returned the wrong menu item...
   lCount = GetMenuItemCount(hMenuSeek)
   If (lCount > 0) Then
      For l = 0 To lCount - 1
         tMII.cbSize = Len(tMII)
         tMII.fMask = MIIM_ID
         GetMenuItemInfo hMenuSeek, l, True, tMII
         If (tMII.wID = m_tMI(lIndex).lActualID) And (m_tMI(lIndex).hMenu = hMenuSeek) Then
            plMenuPositionForIndex = l + 1
         End If
      Next l
   End If
End Function
Private Function plFindItemInMenu( _
   ByVal hMenuSeek As Long, _
   ByVal lPosition As Long _
   ) As Long
   Dim lPos             As Long
   Dim l                As Long, i As Long
   Dim lID              As Long
   Dim tMII             As MENUITEMINFO

   ' fixed bug where this returned the wrong menu item...
   tMII.cbSize = Len(tMII)
   tMII.fMask = MIIM_ID
   GetMenuItemInfo hMenuSeek, lPosition - 1, True, tMII

   For i = 1 To m_iMenuCount
      If m_tMI(i).lActualID = tMII.wID And m_tMI(i).hMenu = hMenuSeek Then
         plFindItemInMenu = i
         Exit Function
      End If
   Next i

End Function
Private Function plMenuIndex(ByVal vKey As Variant) As Long
   Dim i                As Long
   ' Signal default
   plMenuIndex = -1
   ' Check for numeric key (i.e. index):
   If (IsNumeric(vKey)) Then
      i = CLng(vKey)
      If (i > 0) And (i <= m_iMenuCount) Then
         plMenuIndex = i
      End If
   Else
      ' Check for string key:
      For i = 1 To m_iMenuCount
         If (m_tMI(i).sKey = vKey) Then
            plMenuIndex = i
            Exit Function
         End If
      Next i
   End If
End Function
Public Property Get IndexForKey( _
   ByVal sKey As String _
   ) As Long
   Dim i                As Long
   i = plMenuIndex(sKey)
   If i = -1 Then i = 0
   IndexForKey = i
End Property
Public Property Get ItemKey( _
   ByVal lIndex As Long _
   ) As String
   ItemKey = m_tMI(lIndex).sKey
End Property
Public Property Let ItemKey( _
   ByVal lIndex As Long, _
   ByVal sKey As String _
   )
   m_tMI(lIndex).sKey = sKey
End Property
Public Property Get ItemData( _
   ByVal lIndex As Long _
   ) As Long
   ItemData = m_tMI(lIndex).lItemData
End Property
Public Property Get ItemParentIndex( _
   ByVal lIndex As Long _
   ) As Long
   ItemParentIndex = m_tMI(lIndex).lParentIndex
End Property
Public Property Let ItemData( _
   ByVal lIndex As Long, _
   ByVal lItemData As Long _
   )
   Dim tMII             As MENUITEMINFO
   m_tMI(lIndex).lItemData = lItemData
   If (lIndex > 0) And (lIndex <= m_iMenuCount) Then
      tMII.cbSize = Len(tMII)
      tMII.fMask = MIIM_DATA
      GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
      tMII.dwItemData = lItemData
      SetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
   End If
End Property

Public Property Let ItemIcon( _
   ByVal lIndex As Long, _
   ByVal lIconIndex As Long _
   )
   m_tMI(lIndex).lIconIndex = lIconIndex
End Property
Public Property Get ItemIcon( _
   ByVal lIndex As Long _
   ) As Long
   ItemIcon = m_tMI(lIndex).lIconIndex
End Property
Public Property Get Checked( _
   ByVal lIndex As Long _
   ) As Boolean
   Dim tMII             As MENUITEMINFO
   tMII.fMask = MIIM_STATE
   tMII.cbSize = LenB(tMII)
   GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
   m_tMI(lIndex).bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED)
   Checked = m_tMI(lIndex).bChecked
End Property
Public Property Let Checked( _
   ByVal lIndex As Long, _
   ByVal bChecked As Boolean _
   )
   Dim lFlag            As Long
   Dim lFlagNot         As Long

   m_tMI(lIndex).bChecked = bChecked
   If (bChecked) Then
      lFlag = MF_CHECKED
      lFlagNot = 0
   Else
      lFlag = 0
      lFlagNot = MF_CHECKED
   End If
   pSetMenuFlag lIndex, lFlag, lFlagNot

End Property
Public Property Get RadioCheck(ByVal lIndex As Long) As Boolean
   ' BMS 19/9/99: Added Property Get RadioChecked, to allow client to
   ' determine if an item is of type RadioCheck or not.
   If (lIndex > 0) Then
      RadioCheck = m_tMI(lIndex).bRadioCheck
   End If
End Property
Public Property Let RadioCheck( _
   ByVal lIndex As Long, _
   ByVal bRadioCheck As Boolean _
   )
   ' BMS 20/9/99: Added Property Let RadioChecked, to allow client to
   ' set if an item is of type RadioCheck or not.
   Dim lFlag            As Long
   Dim lFlagNot         As Long
   If bRadioCheck Then
      If m_tMI(lIndex).bChecked Then
         Checked(lIndex) = False
      End If
   End If
   m_tMI(lIndex).bRadioCheck = bRadioCheck
   If (bRadioCheck) Then
      lFlag = MFT_RADIOCHECK
      lFlagNot = 0
   Else
      lFlag = 0
      lFlagNot = MFT_RADIOCHECK
   End If
   pSetMenuFlag lIndex, lFlag, lFlagNot
End Property

Public Sub GroupToggle(ByVal lIndex As Long, Optional ByVal bRadio As Boolean = True)
   Dim hMenuSeek        As Long
   Dim lPos             As Long
   Dim l                As Long
   Dim tMII             As MENUITEMINFO

   ' Check a radio item and toggle off any others within
   ' this menu space:
   If lIndex > 0 And lIndex <= m_iMenuCount Then
      RadioCheck(lIndex) = True
      hMenuSeek = m_tMI(lIndex).hMenu
      lPos = plMenuPositionForIndex(hMenuSeek, lIndex) - 1
      If lPos > -1 Then
         For l = lPos - 1 To 0 Step -1
            tMII.cbSize = Len(tMII)
            tMII.fMask = MIIM_ID Or MIIM_TYPE
            GetMenuItemInfo hMenuSeek, l, True, tMII
            If (tMII.fType And MF_SEPARATOR) <> MF_SEPARATOR Then
               lIndex = plGetIndexForId(tMII.wID)
               RadioCheck(lIndex) = False
            Else
               Exit For
            End If
         Next l

         For l = lPos + 1 To GetMenuItemCount(hMenuSeek) - 1
            tMII.cbSize = Len(tMII)
            tMII.fMask = MIIM_TYPE Or MIIM_ID
            GetMenuItemInfo hMenuSeek, l, True, tMII
            If (tMII.fType And MF_SEPARATOR) <> MF_SEPARATOR Then
               lIndex = plGetIndexForId(tMII.wID)
               RadioCheck(lIndex) = False
            Else
               Exit For
            End If
         Next l
      End If
   End If

End Sub

Public Property Get Enabled( _
   ByVal lIndex As Long _
   ) As Boolean
   Dim tMII             As MENUITEMINFO
   tMII.fMask = MIIM_STATE
   tMII.cbSize = LenB(tMII)
   GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
   m_tMI(lIndex).bEnabled = Not ((tMII.fState And MFS_DISABLED) = MFS_DISABLED)
   Enabled = m_tMI(lIndex).bEnabled
End Property
Public Property Let Enabled( _
   ByVal lIndex As Long, _
   ByVal bEnabled As Boolean _
   )
   Dim lFlag            As Long
   Dim lFlagNot         As Long

   m_tMI(lIndex).bEnabled = bEnabled
   If (bEnabled) Then
      lFlag = MF_ENABLED
      lFlagNot = MF_GRAYED
   Else
      lFlag = MF_DISABLED
      lFlagNot = MF_GRAYED
   End If
   pSetMenuFlag lIndex, lFlag, lFlagNot
End Property
Public Property Get Caption( _
   ByVal lIndex As Long _
   ) As String
   Caption = m_tMI(lIndex).sCaption
End Property
Public Property Let Caption( _
   ByVal lIndex As Long, _
   ByVal sCaption As String _
   )

   m_tMI(lIndex).sCaption = sCaption
   m_tMI(lIndex).sAccelerator = psExtractAccelerator(sCaption)

   ResetMenuForRecalc lIndex

End Property
Private Sub ResetMenuForRecalc(ByVal lIndex As Long)
   Dim tMI              As MENUITEMINFO
   Dim hMenu            As Long
   Dim hSubMenu         As Long
   Dim lFlags           As Long
   Dim lR               As Long
   Dim lPosition        As Long
   Dim iCount           As Long, i As Long

   ' Modify the menu item:
   hMenu = m_tMI(lIndex).hMenu
   tMI.cbSize = Len(tMI)
   tMI.fMask = MIIM_SUBMENU
   GetMenuItemInfo hMenu, m_tMI(lIndex).lActualID, 0, tMI
   hSubMenu = tMI.hSubMenu
   iCount = GetMenuItemCount(hMenu)
   tMI.fMask = MIIM_ID
   For i = 0 To iCount - 1
      GetMenuItemInfo hMenu, i, 1, tMI
      If (tMI.wID <> m_tMI(lIndex).lActualID) Then
         lPosition = lPosition + 1
      Else
         Exit For
      End If
   Next i
   ' remove it from the menu:
   RemoveMenu hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND

   ' Insert it back again at the corect position with the same ID etc:
   lFlags = plMenuFlags(lIndex)
   lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
   lR = InsertMenuByLong(m_tMI(lIndex).hMenu, lPosition, lFlags, m_tMI(lIndex).lID, m_tMI(lIndex).lItemData)
   If (hSubMenu <> 0) Then
      ' If we had a submenu then put that back again:
      lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
      lFlags = lFlags Or MF_POPUP
      lR = ModifyMenuByLong(m_tMI(lIndex).hMenu, m_tMI(lIndex).lID, lFlags, hSubMenu, m_tMI(lIndex).lItemData)
   End If

End Sub
Public Property Get Header( _
   ByVal lIndex As Long _
   ) As Boolean
   Header = m_tMI(lIndex).bTitle
End Property
Public Property Let Header( _
   ByVal lIndex As Long, _
   ByVal bState As Boolean _
   )
   m_tMI(lIndex).bTitle = bState
   ' Modify this menu item to reflect the new properties:
   If (bState) Then
      ' Headers are disabled
      pSetMenuFlag lIndex, MF_DISABLED, MF_GRAYED
   End If
End Property
Public Property Get OwnerDraw( _
   ByVal lIndex As Long _
   ) As Boolean
   OwnerDraw = m_tMI(lIndex).bOwnerDraw
End Property
Public Property Let OwnerDraw( _
   ByVal lIndex As Long, _
   ByVal bState As Boolean _
   )
   m_tMI(lIndex).bOwnerDraw = bState
End Property

Public Property Get Default( _
   ByVal lIndex As Long _
   ) As Boolean
   Default = m_tMI(lIndex).bDefault
End Property
Public Property Let Default( _
   ByVal lIndex As Long, _
   ByVal bState As Boolean _
   )
   m_tMI(lIndex).bDefault = bState
   ' Modify this menu item to reflect the new properties:
   ResetMenuForRecalc lIndex
End Property

Public Sub Clear()
   Dim lMenu            As Long

   m_lLastMaxId = &H800
   ' Clear up all submenus we have created:
   For lMenu = 1 To m_lSubMenuCount
      DestroyMenu m_hSubMenus(lMenu)
   Next lMenu
   m_lSubMenuCount = 0
   Erase m_hSubMenus

   ' Clear the main menu we have created:
   If (m_iMenuCount > 0) Then
      DestroyMenu m_tMI(1).hMenu
   End If

   ' Clear up the array:
   m_iMenuCount = 0
   Erase m_tMI

End Sub

Public Function ShowPopupMenu( _
   ByVal lLeft As Long, _
   ByVal lTop As Long, _
   Optional ByVal lExcludeLeft As Long = 0, _
   Optional ByVal lExcludeTop As Long = 0, _
   Optional ByVal lExcludeRight As Long = 0, _
   Optional ByVal lExcludeBottom As Long = 0, _
   Optional ByVal bTryToKeepLeft As Boolean = True _
   ) As Long
   Dim tP               As POINTAPI
   Dim tP2              As POINTAPI
   Dim tPM              As TPMPARAMS
   Dim lR               As Long
   Dim lIndex           As Long
   Dim i                As Long
   Dim lUN              As Long
   Dim bIsSubclass      As Boolean

   If Count = 0 Then
      Exit Function
   End If

   tP.x = lLeft \ Screen.TwipsPerPixelX
   tP.y = lTop \ Screen.TwipsPerPixelY
   ClientToScreen m_hWndOwner, tP

   If Abs(lExcludeLeft - lExcludeRight) > 0 Or Abs(lExcludeTop - lExcludeBottom) > 0 Then
      tP2.x = lExcludeLeft \ Screen.TwipsPerPixelX
      tP2.y = lExcludeTop \ Screen.TwipsPerPixelY
      ClientToScreen m_hWndOwner, tP2
      tPM.rcExclude.left = tP2.x
      tPM.rcExclude.tOp = tP2.y
      tP2.x = lExcludeRight \ Screen.TwipsPerPixelX
      tP2.y = lExcludeBottom \ Screen.TwipsPerPixelY
      ClientToScreen m_hWndOwner, tP2
      tPM.rcExclude.Right = tP2.x
      tPM.rcExclude.Bottom = tP2.y
   End If
   tPM.cbSize = Len(tPM)
   lUN = TPM_RETURNCMD
   If Not (bTryToKeepLeft) Then
      lUN = lUN Or TPM_VERTICAL
   End If
   SendMessageLong m_hWndOwner, WM_ENTERMENULOOP, 1, 0
   bIsSubclass = (m_hWndAttached <> 0)
   CreateSubClass m_hWndOwner
   lR = TrackPopupMenuEx(m_tMI(1).hMenu, lUN, tP.x, tP.y, m_hWndOwner, tPM)
   ' Find the index of the item with id lR within the menu:
   If lR > 0 Then
      lIndex = ItemForID(lR)
      ShowPopupMenu = lIndex
   End If
   If Not bIsSubclass Then
      If lR > 0 Then
         RaiseEvent Click(lIndex)
      End If
      DestroySubClass
   End If
   SendMessageLong m_hWndOwner, WM_EXITMENULOOP, 1, 0

   ' The WM_COMMAND message is sent after this sub exits.
End Function
Public Function ShowPopupMenuAtIndex( _
   ByVal lLeft As Long, _
   ByVal lTop As Long, _
   Optional ByVal lExcludeLeft As Long = 0, _
   Optional ByVal lExcludeTop As Long = 0, _
   Optional ByVal lExcludeRight As Long = 0, _
   Optional ByVal lExcludeBottom As Long = 0, _
   Optional ByVal bTryToKeepLeft As Boolean = True, _
   Optional ByVal lIndex As Long = 1 _
   ) As Long
   Dim tP               As POINTAPI
   Dim tP2              As POINTAPI
   Dim tPM              As TPMPARAMS
   Dim lR               As Long
   Dim i                As Long
   Dim lUN              As Long
   Dim bIsSubclass      As Boolean
   Dim hMenu            As Long

   If lIndex > 0 Then
      If lIndex <= 0 Or lIndex > Count Then
         Exit Function
      End If
      hMenu = m_tMI(lIndex).hMenu
   Else
      If Count = 0 Then
         Exit Function
      End If
      hMenu = m_tMI(1).hMenu
   End If

   tP.x = lLeft \ Screen.TwipsPerPixelX
   tP.y = lTop \ Screen.TwipsPerPixelY
   ClientToScreen m_hWndOwner, tP

   If Abs(lExcludeLeft - lExcludeRight) > 0 Or Abs(lExcludeTop - lExcludeBottom) > 0 Then
      tP2.x = lExcludeLeft \ Screen.TwipsPerPixelX
      tP2.y = lExcludeTop \ Screen.TwipsPerPixelY
      ClientToScreen m_hWndOwner, tP2
      tPM.rcExclude.left = tP2.x
      tPM.rcExclude.tOp = tP2.y
      tP2.x = lExcludeRight \ Screen.TwipsPerPixelX
      tP2.y = lExcludeBottom \ Screen.TwipsPerPixelY
      ClientToScreen m_hWndOwner, tP2
      tPM.rcExclude.Right = tP2.x
      tPM.rcExclude.Bottom = tP2.y
   End If
   tPM.cbSize = Len(tPM)
   lUN = TPM_RETURNCMD
   If Not (bTryToKeepLeft) Then
      lUN = lUN Or TPM_VERTICAL
   End If
   SendMessageLong m_hWndOwner, WM_ENTERMENULOOP, 1, 0
   bIsSubclass = (m_hWndAttached <> 0)
   CreateSubClass m_hWndOwner
   lR = TrackPopupMenuEx(hMenu, lUN, tP.x, tP.y, m_hWndOwner, tPM)
   ' Find the index of the item with id lR within the menu:
   If lR > 0 Then
      lIndex = ItemForID(lR)
      ShowPopupMenuAtIndex = lIndex
   End If
   If Not bIsSubclass Then
      If lR > 0 Then
         RaiseEvent Click(lIndex)
      End If
      DestroySubClass
   End If
   SendMessageLong m_hWndOwner, WM_EXITMENULOOP, 1, 0

   ' The WM_COMMAND message is sent after this sub exits.
End Function

Public Function ShowPopupAbsolute( _
   ByVal lLeftPixel As Long, _
   ByVal lTopPixel As Long, _
   Optional ByVal lIndex As Long = 0, _
   Optional ByVal bTryToKeepLeft As Boolean = True _
   ) As Long
   Dim tR               As RECT
   Dim lUN              As Long
   Dim hMenu            As Long
   Dim bIsSubclass      As Boolean
   Dim lR               As Long

   If lIndex > 0 Then
      If lIndex <= 0 Or lIndex > Count Then
         Exit Function
      End If
      hMenu = m_tMI(lIndex).hMenu
   Else
      If Count = 0 Then
         Exit Function
      End If
      hMenu = m_tMI(1).hMenu
   End If

   lUN = TPM_RETURNCMD
   If Not (bTryToKeepLeft) Then
      lUN = lUN Or TPM_VERTICAL
   End If
   SendMessageLong m_hWndOwner, WM_ENTERMENULOOP, 1, 0
   bIsSubclass = (m_hWndAttached <> 0)
   CreateSubClass m_hWndOwner
   lR = TrackPopupMenu(hMenu, lUN, lLeftPixel, lTopPixel, 0, m_hWndOwner, tR)
   ' Find the index of the item with id lR within the menu:
   lIndex = ItemForID(lR)
   ShowPopupAbsolute = lIndex
   If Not bIsSubclass Then
      If lR > 0 Then
         RaiseEvent Click(lIndex)
      End If
      DestroySubClass
   End If
   SendMessageLong m_hWndOwner, WM_EXITMENULOOP, 1, 0

   ' The WM_COMMAND message is sent after this sub exits.

End Function
Private Sub pSetMenuFlag( _
   ByVal lIndex As Long, _
   ByVal lFlag As Long, _
   ByVal lFlagNot As Long _
   )
   Dim tMII             As MENUITEMINFO
   Dim lFlags           As Long

   lFlags = plMenuFlags(lIndex)
   lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING
   tMII.fMask = MIIM_SUBMENU
   tMII.cbSize = LenB(tMII)
   GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
   If (tMII.hSubMenu <> 0) Then
      lFlags = lFlags Or MF_POPUP
   End If
   lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND

   lFlags = lFlags Or lFlag
   lFlags = lFlags And Not lFlagNot

   ModifyMenuByLong m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, lFlags, m_tMI(lIndex).lActualID, m_tMI(lIndex).lItemData

End Sub
Public Property Get hMenu(ByVal lIndex As Long) As Long
   hMenu = m_tMI(lIndex).hMenu
End Property

Private Sub pAddNewMenuItem( _
   ByRef tMI As tMenuItem _
   )
   Dim tMII             As MENUITEMINFO
   Dim hMenu            As Long
   Dim lPIndex          As Long
   Dim lFlags           As Long
   Dim lR               As Long
   Dim hMenuNew         As Long
   Dim bOk              As Boolean
   Dim i                As Long

   ' Find out where we're adding this item:
   With tMI
      If (.lParentId = 0) Then
         ' This is a new top level menu item:
         If (m_iMenuCount = 1) Then
            ' Create a new menu
            hMenu = CreatePopupMenu()
         Else
            ' Use the existing menu:
            hMenu = m_tMI(1).hMenu
         End If
      Else
         ' We are adding to an existing menu:
         ' First we need to determine if there is already a sub menu for the parent item:
         lPIndex = plGetIndexForId(tMI.lParentId)
         If (lPIndex = 0) Then
            Debug.Print " *** Couldn't find parent... *** "
         Else
            ' Determine if the parent menu has a sub-menu:
            tMII.fMask = MIIM_SUBMENU
            tMII.cbSize = LenB(tMII)
            GetMenuItemInfo m_tMI(lPIndex).hMenu, m_tMI(lPIndex).lActualID, False, tMII
            hMenu = tMII.hSubMenu
            If (hMenu = 0) Then
               ' We don't have a sub menu for this item so we're
               ' going to have to add one:
               ' Debug.Print "Adding new sub-menu:"

               ' Create the new menu item and store it's handle so we can clear up
               ' again later:
               hMenu = CreatePopupMenu()
               If (hMenu = 0) Then
                  Debug.Print " *** Failed to create sub menu *** "
               Else
                  ' Check if hMenu isn't an ID:
                  bOk = False
                  If Not (pbIDIsUnique(hMenu)) Then
                     For i = 1 To 100
                        hMenuNew = CreatePopupMenu()
                        If (pbIDIsUnique(hMenuNew)) Then
                           DestroyMenu hMenu
                           hMenu = hMenuNew
                           bOk = True
                           Exit For
                        Else
                           DestroyMenu hMenuNew
                        End If
                     Next i
                     If Not bOk Then
                        ' .. out of menu handles ...
                        DestroyMenu hMenu
                        Debug.Print "Out of Menu Handles"
                        Exit Sub
                     End If
                  End If

                  m_lSubMenuCount = m_lSubMenuCount + 1
                  ReDim Preserve m_hSubMenus(1 To m_lSubMenuCount) As Long
                  m_hSubMenus(m_lSubMenuCount) = hMenu

                  ' Now set the parent item so it has a popup menu:
                  lFlags = plMenuFlags(lPIndex)
                  lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING
                  lFlags = lFlags Or MF_POPUP
                  lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND

                  lR = ModifyMenuByLong(m_tMI(lPIndex).hMenu, m_tMI(lPIndex).lActualID, lFlags, hMenu, m_tMI(lPIndex).lItemData)
                  ItemData(lPIndex) = m_tMI(lPIndex).lItemData
                  If (lR = 0) Then
                     Debug.Print "Failed to modify menu to add the sub menu " & Err.LastDllError
                  End If

                  ' WHen you add a sub menu to an item, its id becomes the sub menu handle:
                  m_tMI(lPIndex).lActualID = hMenu
                  tMI.lParentId = hMenu
               End If
            End If
         End If
      End If

      If (hMenu <> 0) Then
         lFlags = plMenuFlags(m_iMenuCount)
         lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
         lR = AppendMenuBylong(hMenu, lFlags, tMI.lID, tMI.lItemData)
         If (lR = 0) Then
            Debug.Print "Failed to add new Menu item"
         End If
      End If

      ' Store the hMenu for this item:
      .hMenu = hMenu

   End With
End Sub
Private Function plMenuFlags( _
   ByVal lIndex As Long _
   )
   Dim lFlags           As Long
   With m_tMI(lIndex)
      If (.bChecked) Then
         lFlags = lFlags Or MF_CHECKED
      Else
         lFlags = lFlags Or MF_UNCHECKED
      End If
      If (.bEnabled) Then
         lFlags = lFlags Or MF_ENABLED
      Else
         lFlags = lFlags Or MF_GRAYED
      End If
      If left$(Trim$(.sInputCaption), 1) = "-" Or .bTitle Then
         ' Debug.Print .sInputCaption
         lFlags = lFlags Or MF_SEPARATOR
      End If
      If (m_tMI(lIndex).bMenuBarBreak) Then
         lFlags = lFlags Or MF_MENUBARBREAK
      End If
      If (m_tMI(lIndex).bMenuBreak) Then
         lFlags = lFlags Or MF_MENUBREAK
      End If

   End With
   plMenuFlags = lFlags

End Function
Private Function psExtractAccelerator( _
   ByVal sCaption As String _
   )
   Dim i                As Long

   For i = 1 To Len(sCaption)
      If (Mid$(sCaption, i, 1) = "&") Then
         If (i < Len(sCaption)) Then
            psExtractAccelerator = UCase$(Mid$(sCaption, (i + 1), 1))
         End If
         Exit For
      End If
   Next i
End Function
Private Function plGetNewID() As Long
   Dim lID              As Long

   If (m_lLastMaxId < m_iMenuCount) Then
      m_lLastMaxId = m_iMenuCount
   Else
      m_lLastMaxId = m_lLastMaxId + 1
   End If
   lID = m_lLastMaxId
   Do Until (pbIDIsUnique(lID))
      lID = lID + 1
      m_lLastMaxId = lID
   Loop
   plGetNewID = lID
End Function
Private Function pbIDIsUnique( _
   ByVal lID As Long _
   ) As Boolean
   Dim bFound           As Boolean
   Dim lMenu            As Long
   For lMenu = 1 To m_iMenuCount
      If (m_tMI(lMenu).lActualID = lID) Or (m_tMI(lMenu).lID = lID) Then
         bFound = True
         Exit For
      End If
   Next lMenu
   pbIDIsUnique = Not (bFound)
End Function
Property Let TickIconIndex( _
   ByVal lTickIconIndex As Long _
   )
   m_lTickIconIndex = lTickIconIndex
End Property
Property Get TickIconIndex() As Long
   TickIconIndex = m_lTickIconIndex
End Property
Property Let OptionIconIndex( _
   ByVal lOptionIconIndex As Long _
   )
   m_lOptionIconIndex = lOptionIconIndex
End Property
Property Get OptionIconIndex() As Long
   OptionIconIndex = m_lOptionIconIndex
End Property

Public Sub CreateSubClass(hWndA As Long)
   If m_hWndAttached = hWndA Then
   Else
      DestroySubClass
      Debug.Print "SUBCLASSING"
      AttachMessage Me, hWndA, WM_MENUSELECT
      AttachMessage Me, hWndA, WM_MEASUREITEM
      AttachMessage Me, hWndA, WM_DRAWITEM
      AttachMessage Me, hWndA, WM_COMMAND
      AttachMessage Me, hWndA, WM_MENUCHAR
      AttachMessage Me, hWndA, WM_INITMENUPOPUP
      AttachMessage Me, hWndA, WM_WININICHANGE
      m_hWndAttached = hWndA
   End If
End Sub

Public Sub DestroySubClass()
   If (m_hWndAttached <> 0) Then
      Debug.Print "NOT SUBCLASSING"
      DetachMessage Me, m_hWndAttached, WM_MENUSELECT
      DetachMessage Me, m_hWndAttached, WM_MEASUREITEM
      DetachMessage Me, m_hWndAttached, WM_DRAWITEM
      DetachMessage Me, m_hWndAttached, WM_COMMAND
      DetachMessage Me, m_hWndAttached, WM_MENUCHAR
      DetachMessage Me, m_hWndAttached, WM_INITMENUPOPUP
      DetachMessage Me, m_hWndAttached, WM_WININICHANGE
      m_hWndAttached = 0
   End If
End Sub
Friend Function plGetIndexForId( _
   ByVal lItemId As Long _
   ) As Long
   Dim l                As Long
   Dim lIndex           As Long
   'Debug.Print "Finding Index:"
   'Debug.Print lItemId
   lIndex = 0
   For l = 1 To m_iMenuCount
      'Debug.Print "    Index at l = " & m_tMI(l).lId
      If (m_tMI(l).lActualID = lItemId) Then
         lIndex = l
         Exit For
      End If
   Next l
   plGetIndexForId = lIndex
End Function
Private Function RaiseClickEvent(lID As Long) As Boolean
   ' Return true from this if we have completely handled the
   ' click on our own:
   Dim lIndex           As Long

   ' Find the Index of this menu id within our own array:
   lIndex = plGetIndexForId(lID)

   ' If we find it, then raise a click event for it:
   If (lIndex > 0) Then

      ' Send a click event with the index:
      RaiseEvent Click(lIndex)

      ' If this was one of the VB menu entries we have
      ' subclassed, we want to return false.  Then the
      ' click will filter through to the original Click
      ' event so your code should work as normal:
      If Not (m_tMI(lIndex).bIsAVBMenu) Then
         RaiseClickEvent = True
      End If

   Else

      ' This is a problem.  We've got a click on
      ' a menu id which doesn't seem to be any
      ' of the menu items of the form.  It shouldn't
      ' happen, but return false anyway so we don't eat
      ' the message.
      Debug.Print "Failed to find index"

      RaiseClickEvent = False
   End If
End Function
Private Sub RaiseHighlightEvent(lID As Long)
   Dim lIndex           As Long

   ' Debug.Print lItem
   lIndex = plGetIndexForId(lID)
   If (lIndex > 0) Then
      RaiseEvent ItemHighlight(lIndex, m_tMI(lIndex).bEnabled, (Trim$(m_tMI(lIndex).sCaption = "-")))
   Else
      Debug.Print "Failed to find Index for Highlight Id:", lID, lIndex
   End If
End Sub
Private Sub RaiseInitMenuEvent( _
   ByVal hMenu As Long _
   )
   Dim lIndex           As Long
   Dim lParentId        As Long
   Dim bFound           As Boolean

   ' Firstly, we need to find the index of an item
   ' in hMenu:
   For lIndex = m_iMenuCount To 1 Step -1
      If (m_tMI(lIndex).hMenu = hMenu) Then
         lParentId = m_tMI(lIndex).lParentId
         bFound = True
      End If
      If (bFound) Then
         If (m_tMI(lIndex).lActualID = lParentId) Then
            RaiseEvent InitPopupMenu(lIndex)
            Exit For
         End If
      End If
   Next lIndex

End Sub
Private Sub RaiseMenuExitEvent()
   RaiseEvent MenuExit
End Sub
Private Function pGetTextPosition( _
   ByVal lHDC As Long, _
   ByVal lIndex As Long, _
   ByRef rcItem As RECT _
   )
   Dim tC               As RECT
   Dim lDiff            As Long
   Dim lMenuHeight      As Long

   lMenuHeight = m_lMenuItemHeight

   ' Determine the size of the text to draw:
   DrawText lHDC, m_tMI(lIndex).sCaption, Len(m_tMI(lIndex).sCaption), tC, DT_CALCRECT

   ' We want to centre the text vertically:
   lDiff = lMenuHeight - (tC.Bottom - tC.tOp)
   If (lDiff > 0) Then
      rcItem.tOp = rcItem.tOp + lDiff \ 2
   End If

   ' All normal menu items are indented by to
   ' accomodate icon & checked surround for icon:
   rcItem.left = rcItem.left + lMenuHeight + 2

End Function

Private Function DrawItem( _
   ByRef lParam As Long, _
   ByRef wParam As Long _
   ) As Boolean
   Dim tDI              As DRAWITEMSTRUCT
   Dim lHDC             As Long
   Dim lIndex           As Long
   Dim lColour          As Long
   Dim lFillColour      As Long
   Dim bDisabled        As Boolean
   Dim bSelected        As Boolean
   Dim bChecked         As Boolean
   Dim bHighlight       As Boolean
   Dim bRadioChecked    As Boolean
   Dim bIsTopLevel      As Boolean
   Dim bHeader          As Boolean
   Dim bDoDefault       As Boolean
   Dim hBrush           As Long
   Dim tP               As POINTAPI
   Dim tB               As RECT, tC As RECT, tOB As RECT
   Dim tS               As RECT, tFR As RECT, tIR As RECT, tFC As RECT, tRF As RECT
   Dim sText            As String
   Dim x                As Long
   Dim hBr              As Long
   Dim bKeyHighlight    As Boolean
   Dim bKeyTextNo3d     As Boolean
   Static iLastSelDisabled As Long

   CopyMemory tDI, ByVal lParam, Len(tDI)
   ' Debug.Print tDI.CtlID, tDI.CtlType, tDI.hwndItem, tDI.itemID
   If (tDI.itemAction = 2) Then
      If (GetAsyncKeyState(vbKeyUp) <> 0) Or (GetAsyncKeyState(vbKeyDown) <> 0) Then
         bKeyHighlight = True
         bKeyTextNo3d = True
      End If
   End If

   If (tDI.CtlType = 1) Then ' Menu
      lIndex = (plGetIndexForId(tDI.itemID))
      If (lIndex <> 0) Then

         ' Debug.Print "Found item to draw."
         lHDC = tDI.HDC
         SetBkMode lHDC, TRANSPARENT

         bDisabled = ((tDI.itemState And ODS_DISABLED) = ODS_DISABLED) Or ((tDI.itemState And ODS_GRAYED) = ODS_GRAYED)
         bSelected = ((tDI.itemState And ODS_SELECTED) = ODS_SELECTED)
         If Not bSelected And bKeyHighlight Then
            bKeyTextNo3d = False
         End If
         bChecked = ((tDI.itemState And ODS_CHECKED) = ODS_CHECKED)
         bRadioChecked = RadioCheck(lIndex)
         bHeader = m_tMI(lIndex).bTitle

         If (m_tMI(lIndex).bOwnerDraw) Then
            RaiseEvent DrawItem(tDI.HDC, lIndex, tDI.rcItem.left, tDI.rcItem.tOp, tDI.rcItem.Right, tDI.rcItem.Bottom, bSelected, bChecked Or bRadioChecked, bDisabled, bDoDefault)
         Else
            bDoDefault = True
         End If

         If (bDoDefault) Then
            tP.x = tDI.rcItem.left
            tP.y = tDI.rcItem.tOp + 1

            ' Store the rectangle into tOB so we can draw an
            ' outer border if this is a top level item:
            CopyMemory tOB, tDI.rcItem, LenB(tB)

            If (m_hDCBack <> 0) Then
               TileArea lHDC, tDI.rcItem.left, tDI.rcItem.tOp, tDI.rcItem.Right - tDI.rcItem.left, tDI.rcItem.Bottom - tDI.rcItem.tOp, m_hDCBack, m_lBitmapW, m_lBitmapH
            End If

            If (bDisabled) Then
               If (Trim$(m_tMI(lIndex).sCaption = "-")) Then
                  ' If this is a separator, then draw the separator line:
                  tS.left = tP.x
                  tS.Right = tDI.rcItem.Right
                  tS.tOp = tP.y + 1
                  tS.Bottom = tS.tOp + 2
                  DrawEdge lHDC, tS, EDGE_ETCHED, BF_TOP
               Else
                  If (bHeader) Then
                     tS.left = tP.x
                     tS.Right = tDI.rcItem.Right
                     If (m_bDrawHeadersAsSeparators) Then
                        tS.tOp = tDI.rcItem.tOp + (tDI.rcItem.Bottom - tDI.rcItem.tOp - 2) \ 2
                        tS.Bottom = tS.tOp + 2
                        DrawEdge lHDC, tS, EDGE_ETCHED, BF_TOP
                        SetBkMode tDI.HDC, OPAQUE
                        SetTextColor tDI.HDC, GetSysColor(vbMenuText And &H1F&)
                     Else
                        tS.tOp = tDI.rcItem.tOp
                        tS.Bottom = tDI.rcItem.Bottom
                        InflateRect tS, -1, -2
                        hBr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        SetTextColor tDI.HDC, GetSysColor(vbInactiveCaptionText And &H1F&)
                        FillRect tDI.HDC, tS, hBr
                        DeleteObject hBr
                     End If

                     Dim hFont            As Long, hFontOld As Long

                     ' Draw a separator line and draw the text over
                     ' it, centred, opaque background, smaller font:
                     sText = Trim$(m_tMI(lIndex).sCaption)
                     If (sText <> "") Then
                        tS.tOp = tDI.rcItem.tOp
                        tS.Bottom = tDI.rcItem.Bottom
                        hFont = m_cNCM.FontHandle(SMCaptionFont)
                        'Alternative: hFont = m_cNCM.SmallerFontHandle(MenuFOnt)
                        hFontOld = SelectObject(tDI.HDC, hFont)
                        DrawText tDI.HDC, " " & sText & " ", 2 + Len(sText), tS, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
                        SelectObject tDI.HDC, hFontOld
                        'Alternative: DeleteObject hFont
                     End If

                     If (m_bDrawHeadersAsSeparators) Then
                        SetBkMode tDI.HDC, TRANSPARENT
                     End If

                  Else

                     If (bKeyHighlight) Or (lIndex = iLastSelDisabled And Not (bSelected)) Then
                        ' We draw the highlight:
                        iLastSelDisabled = lIndex
                        If (bSelected) Then
                           lFillColour = GetSysColor(COLOR_HIGHLIGHT)
                           lColour = GetSysColor(COLOR_HIGHLIGHTTEXT)
                           bHighlight = True
                        Else
                           lFillColour = GetSysColor(COLOR_MENU)
                           lColour = GetSysColor(COLOR_MENUTEXT)
                           bHighlight = False
                        End If

                        LSet tFR = tDI.rcItem
                        If (m_tMI(lIndex).lIconIndex > -1) Or (bChecked Or bRadioChecked) Then
                           ' Erase the icon background:
                           CopyMemory tIR, tDI.rcItem, LenB(tIR)
                           tIR.Right = tIR.left + m_lMenuItemHeight - 1
                           hBrush = GetSysColorBrush(COLOR_MENU) 'CreateSolidBrush(GetSysColor(COLOR_MENU))
                           FillRect lHDC, tIR, hBrush
                           DeleteObject hBrush
                           tFR.left = tFR.left + m_lMenuItemHeight
                        End If

                        pDrawHighlight lHDC, tFR, lFillColour, bHighlight
                        SetTextColor lHDC, lColour

                     End If

                     ' Get the position to output the text:
                     pGetTextPosition lHDC, lIndex, tDI.rcItem

                     ' Draw the text grayed:
                     If Not (bKeyTextNo3d) Then
                        With tDI.rcItem
                           .left = .left + 1
                           .tOp = .tOp + 1
                        End With
                        SetTextColor lHDC, GetSysColor(COLOR_BTNHIGHLIGHT)
                        pDrawMenuCaption lHDC, lIndex, tDI.rcItem
                        With tDI.rcItem
                           .left = .left - 1
                           .tOp = .tOp - 1
                        End With
                     End If
                     SetTextColor lHDC, GetSysColor(COLOR_BTNSHADOW)
                     pDrawMenuCaption lHDC, lIndex, tDI.rcItem

                     SetTextColor lHDC, GetSysColor(COLOR_MENUTEXT)
                  End If
               End If
            Else
               If (Trim$(m_tMI(lIndex).sCaption = "-")) Then
                  ' Draw a separator line:
                  If (bIsTopLevel) Then
                     ' We draw nothing - a separator at
                     ' the top level just leaves a space.
                     Debug.Print "Separator at top level"
                  Else
                     ' We draw the separator line:
                     tS.left = tP.x
                     tS.tOp = tP.y + 1
                     tS.Bottom = tS.tOp + 2
                     tS.Right = tDI.rcItem.Right
                     DrawEdge lHDC, tS, EDGE_ETCHED, BF_TOP
                  End If
               Else

                  ' Set the back colour and text colour for the menu item:
                  If bSelected And Not (bIsTopLevel) Then
                     lFillColour = GetSysColor(COLOR_HIGHLIGHT)
                     lColour = GetSysColor(COLOR_HIGHLIGHTTEXT)
                     bHighlight = True
                  Else
                     lFillColour = GetSysColor(COLOR_MENU)
                     lColour = GetSysColor(COLOR_MENUTEXT)
                     bHighlight = False
                  End If
                  CopyMemory tFR, tDI.rcItem, LenB(tFR)
                  If (m_tMI(lIndex).lIconIndex > -1) Or (bChecked Or bRadioChecked) Then
                     ' Erase the icon background:
                     CopyMemory tIR, tDI.rcItem, LenB(tIR)
                     tIR.Right = tIR.left + m_lMenuItemHeight - 1
                     hBrush = GetSysColorBrush(COLOR_MENU)
                     FillRect lHDC, tIR, hBrush
                     DeleteObject hBrush
                     tFR.left = tFR.left + m_lMenuItemHeight
                  End If
                  pDrawHighlight lHDC, tFR, lFillColour, bHighlight
                  SetTextColor lHDC, lColour

                  ' Get the position to output the text:
                  pGetTextPosition lHDC, lIndex, tDI.rcItem
                  pDrawMenuCaption lHDC, lIndex, tDI.rcItem

                  If (bIsTopLevel) And (bSelected) Then
                     ' We draw a sunken box around selected
                     ' top level menu items:
                     tOB.Right = tOB.Right - 1
                     tOB.Bottom = tOB.Bottom
                     DrawEdge lHDC, tOB, BDR_SUNKENOUTER, BF_RECT
                  End If
               End If
            End If

            If (bChecked) Or (bRadioChecked) Then
               ' SPM: 29/07/98 - note due to VB/res problems
               ' dithering only works in the compiled DLL...
               If Not (bDisabled) Then
                  If (bSelected) Then
                     hBrush = GetSysColorBrush(COLOR_BTNFACE)
                  Else
                     If m_bUseDither Then
                        If (NoPalette) Then
                           ' Much nicer: In true colour system go
                           ' for a lighter colour defined at RT
                           Debug.Print "Lighter"
                           hBrush = CreateSolidBrush(LighterColour(GetSysColor(COLOR_BTNFACE)))
                        ElseIf m_hBMPDither <> 0 Then
                           hBrush = CreatePatternBrush(m_hBMPDither)
                        End If
                     Else
                        If (NoPalette) Then
                           Debug.Print "Lighter"
                           hBrush = CreateSolidBrush(LighterColour(GetSysColor(COLOR_BTNFACE)))
                        Else
                           hBrush = GetSysColorBrush(COLOR_BTNFACE)
                        End If
                     End If
                  End If
                  tFC.left = tP.x
                  tFC.tOp = tP.y + 1
                  tFC.Right = tFC.left + m_lMenuItemHeight - 1
                  tFC.Bottom = tFC.tOp + m_lMenuItemHeight - 2

                  FillRect lHDC, tFC, hBrush
                  DeleteObject hBrush
               End If

               ' If we don't have an icon, we ought to draw the tick icon
               ' inside it:
               If (m_tMI(lIndex).lIconIndex = -1) Then
                  If (bIsTopLevel) Then
                     x = tP.x
                  Else
                     x = tP.x + (m_lMenuItemHeight - 18) \ 2
                  End If
                  If bRadioChecked Then
                     If m_lOptionIconIndex = -1 Then
                        tRF.left = x + 1
                        tRF.tOp = tP.y + 3
                        tRF.Right = tRF.left + GetSystemMetrics(SM_CYMENU) - 4
                        tRF.Bottom = tRF.tOp + GetSystemMetrics(SM_CYMENU) - 4
                        DrawMaskedFrameControl lHDC, tRF, _
                           DFC_MENU, _
                           (DFCS_MENUBULLET Or (DFCS_INACTIVE And Not bDisabled))
                     Else
                        If (bDisabled) Then
                           ImageListDrawIconDisabled lHDC, m_hIml, m_lOptionIconIndex, x + 1, tP.y + 3, 16
                        Else
                           ImageListDrawIcon lHDC, m_hIml, m_lOptionIconIndex, x + 1, tP.y + 3, , (bDisabled)
                        End If
                     End If
                  Else
                     If m_lTickIconIndex = -1 Then
                        tRF.left = x + 1
                        tRF.tOp = tP.y + 3
                        tRF.Right = tRF.left + GetSystemMetrics(SM_CYMENU) - 4
                        tRF.Bottom = tRF.tOp + GetSystemMetrics(SM_CYMENU) - 4
                        DrawMaskedFrameControl lHDC, tRF, _
                           DFC_MENU, _
                           (DFCS_MENUCHECK Or (DFCS_INACTIVE And Not bDisabled))
                     Else
                        If (bDisabled) Then
                           ImageListDrawIconDisabled lHDC, m_hIml, m_lTickIconIndex, x + 1, tP.y + 3, 16
                        Else
                           ImageListDrawIcon lHDC, m_hIml, m_lTickIconIndex, x + 1, tP.y + 3, , (bDisabled)
                        End If
                     End If
                  End If
               End If

               ' We draw a sunken box around the checked item:
               tB.left = tP.x
               tB.tOp = tP.y + 1
               tB.Right = tB.left + m_lMenuItemHeight - 1
               tB.Bottom = tB.tOp + m_lMenuItemHeight - 2
               DrawEdge lHDC, tB, BDR_SUNKENOUTER, BF_RECT

            ElseIf (bSelected) And Not (bIsTopLevel) And Not (bDisabled And Not bKeyHighlight) Then
               ' We should draw a raised box around the icon:
               If (m_tMI(lIndex).lIconIndex > -1) Then
                  tB.left = tP.x
                  tB.tOp = tP.y - 1
                  tB.Right = tB.left + m_lMenuItemHeight - 1
                  tB.Bottom = tB.tOp + m_lMenuItemHeight
                  DrawEdge lHDC, tB, BDR_RAISEDINNER, BF_RECT
               End If
            End If

            If (m_tMI(lIndex).lIconIndex > -1) Then
               If (bIsTopLevel) Then
                  x = tP.x
               Else
                  x = tP.x + (m_lMenuItemHeight - 18) \ 2
               End If
               If (bDisabled) Then
                  ImageListDrawIconDisabled lHDC, m_hIml, m_tMI(lIndex).lIconIndex, x + 1, tP.y + 3, 16
               Else
                  ImageListDrawIcon lHDC, m_hIml, m_tMI(lIndex).lIconIndex, x + 1, tP.y + 3, , (bDisabled)
               End If
            End If

         End If ' Owner draw
         DrawItem = True
      Else
         'Debug.Print "Failed to find item to draw.", tDI.itemID, lIndex
      End If
   End If
End Function
Private Sub pDrawHighlight( _
   ByVal lHDC As Long, _
   ByRef tFR As RECT, _
   ByVal lFillColour As Long, _
   ByVal bHighlight As Boolean _
   )
   Dim hBrush           As Long
   If Not (bHighlight) Then
      If m_hDCBack = 0 Then
         hBrush = CreateSolidBrush(lFillColour)
         FillRect lHDC, tFR, hBrush
         DeleteObject hBrush
      End If
   Else
      If m_bGradientHighlight Then
         DrawGradient lHDC, tFR, lFillColour, GetSysColor(COLOR_MENU), False
      Else
         hBrush = CreateSolidBrush(lFillColour)
         FillRect lHDC, tFR, hBrush
         DeleteObject hBrush
      End If
   End If
End Sub
Private Function MeasureItem( _
   ByVal lItemId As Long, _
   ByRef lWidth As Long, _
   ByRef lHeight As Long _
   ) As Boolean
   Dim lIndex           As Long
   Dim tR               As RECT
   Dim bDontEvalWidth   As Long
   Dim bIsTopLevel      As Long
   Dim sLongestCaption  As String
   Dim sLongestShortCut As String
   Dim l                As Long
   Dim lItemsOnMenu()   As Long
   Dim lItemCount       As Long
   Dim hMenu            As Long
   Dim hFont            As Long
   Dim hFontOld         As Long
   Dim lLongestItemIndex As Long

   lIndex = plGetIndexForId(lItemId)
   If (lIndex <> 0) Then
      lHeight = m_lMenuItemHeight

      ' Determine the width of the item:
      If Trim$(m_tMI(lIndex).sCaption = "-") Then
         lHeight = 6
         lWidth = 32
         bDontEvalWidth = True
      ElseIf (m_tMI(lIndex).bTitle) Then
         lWidth = 32
         lHeight = m_lMenuItemHeight * 3 \ 4
      Else
         lWidth = 32
      End If

      If Not (bDontEvalWidth) Then
         If bIsTopLevel Then
            ' For top level items we evaluate the width of
            ' the actual text item only:
            DrawText m_HDC, m_tMI(lIndex).sCaption, Len(m_tMI(lIndex).sCaption), tR, DT_CALCRECT
            lWidth = lWidth + tR.Right
         Else
            ' If we have shortcuts at this level (CTRL-C etc), we need
            ' to evaluate the widths of all text items in this menu,
            ' then set the width to be the width of the longest text +
            ' the longest shortcut item:
            hMenu = m_tMI(lIndex).hMenu
            For l = 1 To m_iMenuCount
               If (m_tMI(l).hMenu = hMenu) Then
                  If Len(m_tMI(l).sCaption) > Len(sLongestCaption) Then
                     sLongestCaption = m_tMI(l).sCaption
                     lLongestItemIndex = l
                  End If
                  If (Len(m_tMI(l).sShortCutDisplay) > Len(sLongestShortCut)) Then
                     sLongestShortCut = m_tMI(l).sShortCutDisplay
                  End If
                  lItemCount = lItemCount + 1
                  ReDim Preserve lItemsOnMenu(1 To lItemCount) As Long
                  lItemsOnMenu(lItemCount) = l
               End If
            Next l
            If (lLongestItemIndex > 0) Then
               If (m_tMI(lLongestItemIndex).bDefault) Then
                  hFont = m_cNCM.BoldenedFontHandle(MenuFOnt)
                  hFontOld = SelectObject(m_HDC, hFont)
               End If
            End If

            DrawText m_HDC, sLongestCaption, Len(sLongestCaption), tR, DT_CALCRECT
            lWidth = lWidth + tR.Right
            If (sLongestShortCut <> "") Then
               DrawText m_HDC, sLongestShortCut, Len(sLongestShortCut), tR, DT_CALCRECT
               lWidth = lWidth + 8
               For l = 1 To lItemCount
                  m_tMI(lItemsOnMenu(l)).lShortCutStartPos = lWidth
               Next l
               lWidth = lWidth + tR.Right
            End If

            If hFontOld <> 0 Then
               SelectObject m_HDC, hFontOld
               hFontOld = 0
            End If
            If hFont <> 0 Then
               DeleteObject hFont
               hFont = 0
            End If
         End If
      End If

      If (m_tMI(lIndex).bOwnerDraw) Then
         RaiseEvent MeasureItem(lIndex, lWidth, lHeight)
      End If

      m_tMI(lIndex).lWidth = lWidth
      m_tMI(lIndex).lHeight = lHeight

      MeasureItem = True
   End If
End Function

Private Sub pDrawMenuCaption( _
   ByVal lHDC As Long, _
   ByVal lIndex As Long, _
   ByRef tR As RECT _
   )
   Dim sText            As String
   Dim tSR              As RECT
   Dim hFont            As Long
   Dim hFontOld         As Long

   If m_tMI(lIndex).bDefault Then
      hFont = m_cNCM.BoldenedFontHandle(MenuFOnt)
      If hFont <> 0 Then
         hFontOld = SelectObject(lHDC, hFont)
      End If
   End If

   sText = Trim$(m_tMI(lIndex).sCaption)
   DrawText lHDC, sText, Len(sText), tR, DT_LEFT

   sText = Trim$(m_tMI(lIndex).sShortCutDisplay)
   If (sText <> "") Then
      CopyMemory tSR, tR, LenB(tR)
      tSR.left = m_tMI(lIndex).lShortCutStartPos
      DrawText lHDC, sText, Len(sText), tSR, DT_LEFT
   End If

   If hFontOld <> 0 Then
      SelectObject lHDC, hFontOld
   End If
   If hFont <> 0 Then
      DeleteObject hFont
   End If

End Sub
Private Function plParseMenuChar( _
   ByVal hMenu As Long, _
   ByVal iChar As Integer _
   ) As Long
   Dim sChar            As String
   Dim lPos             As Long
   Dim l                As Long

   sChar = UCase$(Chr$(iChar))
   For l = 1 To m_iMenuCount
      If (m_tMI(l).hMenu = hMenu) Then
         If (m_tMI(l).sAccelerator = sChar) Then
            plParseMenuChar = &H20000 + lPos
            ' Debug.Print "Found Menu Char"
            Exit Function
         End If
         lPos = lPos + 1
      End If
   Next l

End Function
Private Sub pSelectMenuFont()
   Dim tM               As RECT
   ' If we have already selected the font,
   ' then remove it from the DC:
   If (m_hFntOld <> 0) Then
      SelectObject m_HDC, m_hFntOld
   End If
   ' Get the metrics.  This will delete
   ' the hFont for menu:
   m_cNCM.GetMetrics
   ' Select the latest version of the menu font
   ' into the DC, storing what was there before:
   m_hFntOld = SelectObject(m_HDC, m_cNCM.FontHandle(MenuFOnt))

   ' Determine what height to make the menu items:
   DrawText m_HDC, "yY", 2, tM, DT_CALCRECT
   If (tM.Bottom - tM.tOp) > 18 Then
      m_lMenuItemHeight = tM.Bottom - tM.tOp + 6
   Else
      m_lMenuItemHeight = 22
   End If

End Sub

Public Property Get MenuItemHeight(ByVal lIndex As Long) As Long
   MenuItemHeight = m_tMI(lIndex).lHeight
End Property
Public Property Get MenuItemWidth(ByVal lIndex As Long) As Long
   MenuItemWidth = m_tMI(lIndex).lWidth
End Property

Private Sub Class_Initialize()
   Debug.Print "cPopupMenu:Initialize"
   m_lLastMaxId = &H800
   m_lTickIconIndex = -1
   m_lOptionIconIndex = -1
   HighlightCheckedItems = True
   AttachKeyboardHook Me
End Sub

Private Sub Class_Terminate()
   Dim i                As Long

   DetachKeyboardHook Me
   DestroySubClass
   Clear
   ' Remove the graphics:
   ClearBackgroundPicture
   If (m_HDC <> 0) Then
      If (m_hFntOld <> 0) Then
         SelectObject m_HDC, m_hFntOld
      End If
      DeleteObject m_HDC
   End If
   If (m_hBMPDither <> 0) Then
      DeleteObject m_hBMPDither
   End If
   ' Clear the non-client object, removing any fonts:
   Set m_cNCM = Nothing
   ClearUpWorkDC
   ' Clear up any stored menus:
   For i = 1 To m_iStoreCount
      Set m_cStoredMenu(i) = Nothing
   Next i
   Erase m_cStoredMenu
   m_iStoreCount = 0

   Debug.Print "cPopupMenu:Terminate"
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As VBIDEUtils0.EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As VBIDEUtils0.EMsgResponse
   Select Case CurrentMessage
      Case WM_COMMAND, WM_DRAWITEM, WM_MEASUREITEM
         ISubclass_MsgResponse = emrConsume
      Case Else
         ISubclass_MsgResponse = emrPreprocess
   End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim lMenuId          As Long, hMenu As Long, lItem As Long
   Dim lMenuCount       As Long
   Dim lHiWord          As Long
   Dim bEnabled         As Boolean, bSeparator As Boolean
   Dim bFound           As Boolean
   Dim iChar            As Integer
   Dim lFlag            As Long
   Dim bHandled         As Boolean
   Dim lR               As Long

   ' Debug.Print iMsg

   Select Case iMsg

      ' Handle Menu Select events:
      Case WM_MENUSELECT
         ' Extract the menu id and flags for the selected
         ' menu item:
         lHiWord = wParam \ &H10000
         lMenuId = wParam And &HFFFF&

         ' MenuId 0 corresponds to a separator on the system
         ' menu:
         If (lMenuId <> 0) Then

            ' Extract separator & enabled/disabled from the flags
            ' stored in the High Word of wParam:
            bSeparator = ((lHiWord And MF_SEPARATOR) = MF_SEPARATOR)
            bEnabled = ((lHiWord And MF_DISABLED) = MF_DISABLED) Or ((lHiWord And MF_GRAYED) = MF_GRAYED)

            ' Menu handle is passed in as lParam:
            hMenu = lParam

            ' Now check if the message is a menu item higlight,
            ' or whether it is indicating exit from the menu:
            lMenuCount = GetMenuItemCount(hMenu)
            For lItem = 0 To lMenuCount - 1
               If (lMenuId = GetMenuItemID(hMenu, lItem)) Then
                  bFound = True
                  Exit For
               End If
            Next lItem

            ' Raise a highlight or menu exit as required:
            If (bFound) Then
               RaiseHighlightEvent lMenuId
            Else
               RaiseMenuExitEvent
            End If

         End If

         ' Handle menu click events:
      Case WM_COMMAND
         Debug.Print "Got a WM_COMMAND", wParam, lParam

         ' Commands from menus are identified by an lParam of 0
         ' (otherwise it is set the hWnd of the control):
         bHandled = False
         If (lParam = 0) Or (lParam = m_hWndAttached) Then
            ' Low order word of the wParam item is the menu item id:
            lMenuId = (wParam And &HFFFF&)

            Debug.Print "ID: " & lMenuId
            If (RaiseClickEvent(lMenuId)) Then
               ' Don't send on the WM_COMMAND if the item
               ' wasn't a VB menu, it might interfere
               ' with some other control items!
            Else
               'Handled
               ISubclass_WindowProc = 1
               bHandled = True
            End If
         End If
         If Not bHandled Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         End If

         ' Draw Menu items:
      Case WM_DRAWITEM
         ' Debug.Print lParam, wParam
         If (DrawItem(lParam, wParam)) Then
            ISubclass_WindowProc = 1
         Else
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         End If

         ' Measure Menu items prior to drawing them:
      Case WM_MEASUREITEM
         Dim tMis             As MEASUREITEMSTRUCT
         bHandled = False
         CopyMemory tMis, ByVal lParam, Len(tMis)
         If (tMis.CtlType = 1) Then ' Menu
            ' Get the required width & height:
            If (MeasureItem(tMis.itemID, tMis.itemWidth, tMis.itemHeight)) Then
               ' Put the new items back into the structure:
               CopyMemory ByVal lParam, tMis, Len(tMis)
               bHandled = True
               ISubclass_WindowProc = 1
            End If
         End If
         If Not bHandled Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         End If

         ' Handle accelerator (&key) messages in the menu:
      Case WM_MENUCHAR
         ' Check that this is my menu:
         lFlag = wParam \ &H10000
         If ((lFlag And MF_SYSMENU) <> MF_SYSMENU) Then
            hMenu = lParam
            iChar = (wParam And &HFFFF&)
            ' Debug.Print hMenu, Chr$(iChar)
            ' See if this corresponds to an accelerator on the menu:
            lR = plParseMenuChar(hMenu, iChar)
            If lR > 0 Then
               bHandled = True
               ISubclass_WindowProc = lR
            End If
         End If
         If Not bHandled Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         End If

      Case WM_INITMENUPOPUP
         ' Check the sys menu flag:
         If (lParam \ &H10000) > 0 Then
            ' System menu.
         Else
            hMenu = wParam
            ' Find the item which is the parent
            ' of this popup menu:
            RaiseInitMenuEvent hMenu
         End If

      Case WM_WININICHANGE
         Debug.Print "WININICHANGE"
         If Not m_cNCM Is Nothing Then
            ' Clear any pre-created font handles:
            m_cNCM.ClearUp
         End If
         ' First ensure we have the correct font:
         pSelectMenuFont
         ' Now replace every menu item so the new sizes of the
         ' the menu items are correctly displayed...
         For lR = 1 To m_iMenuCount
            ReplaceItem lR
         Next lR

         ' Make sure we pass the message on for
         ' default processing!

      Case Else
         ' Debug.Print "Got an unrequested message:", iMsg
   End Select

End Function

