VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStoreMenu"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tMI()         As tMenuItem
Private m_iMenuCount    As Long
Private m_sKey          As String
Private m_sFailure      As String

Friend Property Get Error() As String
   Error = m_sFailure
End Property

Friend Function Serialise(ByVal iFIle As Integer) As Boolean
   Dim sHeader          As String
   Dim tMI              As tMenuItem
   Dim iItem            As Long

   On Error GoTo ErrorHandler
   Put #iFIle, , "cStoreMenu"
   sHeader = App.Major & ":" & App.Minor & "," & LenB(tMI)
   Put #iFIle, , Len(sHeader)
   Put #iFIle, , sHeader
   Put #iFIle, , Len(m_sKey)
   If Len(m_sKey) > 0 Then
      Put #iFIle, , m_sKey
   End If
   Put #iFIle, , m_iMenuCount
   For iItem = 1 To m_iMenuCount
      Put #iFIle, , m_tMI(iItem)
   Next iItem
   Serialise = True
   Exit Function

ErrorHandler:
   m_sFailure = Err.Description
   Serialise = False
   Exit Function
End Function
Friend Function Deserialise(ByVal iFIle As Integer) As Boolean
   Dim sInfo            As String
   Dim iLen             As Long
   Dim iCount           As Long
   Dim iItem            As Long

   On Error GoTo ErrorHandler
   sInfo = Space$(10)
   Get #iFIle, , sInfo
   If sInfo = "cStoreMenu" Then
      Get #iFIle, , iLen
      sInfo = Space$(iLen)
      Get #iFIle, , sInfo
      ' This returns the app version & length of the tMI structure.  In future
      ' versions this will inform us how to read old data streams, for now
      ' we can just read it as this is the first version which supports serialisation
      Get #iFIle, , iLen
      If iLen > 0 Then
         sInfo = Space$(iLen)
         Get #iFIle, , sInfo
      End If
      Get #iFIle, , iCount
      ' Check not stupid...
      If iCount > -1 And iCount < 32767 Then
         m_iMenuCount = iCount
         If iCount = 0 Then
            Erase m_tMI
         Else
            ReDim m_tMI(1 To m_iMenuCount) As tMenuItem
            For iItem = 1 To m_iMenuCount
               Get #iFIle, , m_tMI(iItem)
            Next iItem
            Deserialise = True
         End If
         m_sKey = sInfo
      Else
         m_sFailure = "Invalid number of Menu items"
      End If
   Else
      m_sFailure = "Not a cNewMenu data stream"
   End If
   Exit Function

ErrorHandler:
   m_sFailure = Err.Description
   Deserialise = False
   Exit Function
End Function

Friend Property Get Key() As String
   Key = m_sKey
End Property
Friend Property Let Key(ByVal sKey As String)
   m_sKey = sKey
End Property

Friend Sub Store(ByRef tMI() As tMenuItem, ByVal iMenuCount As Long)
   Dim iItem            As Long
   m_iMenuCount = iMenuCount
   If (m_iMenuCount > 0) Then
      ReDim m_tMI(1 To m_iMenuCount) As tMenuItem
      For iItem = 1 To m_iMenuCount
         LSet m_tMI(iItem) = tMI(iItem)
      Next iItem
   Else
      Erase m_tMI
   End If
End Sub
Friend Sub Restore(ByRef cMenu As cPopupMenu)
   Dim iItem            As Long
   With cMenu
      For iItem = 1 To m_iMenuCount
         .AddItem m_tMI(iItem).sInputCaption, m_tMI(iItem).sHelptext, m_tMI(iItem).lItemData, m_tMI(iItem).lParentIndex, m_tMI(iItem).lIconIndex, m_tMI(iItem).bChecked, m_tMI(iItem).bEnabled, m_tMI(iItem).sKey
         .Header(iItem) = m_tMI(iItem).bTitle
         .Default(iItem) = m_tMI(iItem).bDefault
         .OwnerDraw(iItem) = m_tMI(iItem).bOwnerDraw
         .RadioCheck(iItem) = m_tMI(iItem).bRadioCheck
      Next iItem
   End With
End Sub
