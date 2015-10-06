VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderedColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cMethod"
Attribute VB_Ext_KEY = "Member0" ,"cMethod"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : OrderedColl
' * Module Filename  : VBDocOrderedColl.cls
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Purpose: Store a Ordered alphabetically collection of objects.<BR>
'The core of object is the Add method.
'Author:   Marco Pipino
Option Explicit

Private mCol            As Collection

'Purpose: Add an object to the collection in the right position,
'ordered alphabelically<BR>
'If ofClass isTrue the key is NameObj_NameClass
Public Function Add(objObject As Object, sKey As String, _
   Optional OfClass As Boolean = False) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : OrderedColl
   ' * Module Filename  : VBDocOrderedColl.cls
   ' * Procedure Name   : Add
   ' * Purpose          :
   ' * Parameters       :
   ' *                    objObject As Object
   ' *                    sKey As String
   ' *                    Optional OfClass As Boolean = False
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim CurrObj          As Object
   Dim currKey          As String
   Dim sKeyBefore       As String
   Dim sKeyAfter        As String

   On Error GoTo err_Add

   sKeyBefore = ""
   sKeyAfter = ""

   If mCol.Count = 0 Then
      mCol.Add objObject, sKey
      Exit Function
   End If

   For Each CurrObj In mCol
      If OfClass Then
         currKey = CurrObj.Name & "_" & CurrObj.RefModule
      Else
         currKey = CurrObj.Name
      End If

      If UCase$(currKey) < UCase$(sKey) Then
         sKeyBefore = currKey
      Else
         sKeyAfter = currKey
         Exit For
      End If
   Next

   If sKeyBefore = "" Then
      mCol.Add objObject, sKey, Before:=sKeyAfter
   ElseIf sKeyAfter = "" Then
      mCol.Add objObject, sKey, After:=sKeyBefore
   Else
      mCol.Add objObject, sKey, After:=sKeyBefore
   End If
   Add = True
   Exit Function
err_Add:
   Add = False
End Function

'Purpose: count the number of items in the collection<BR>
'If the RefModule parameter is valorized then return only
'the count of members of this module
Public Function Count(Optional RefModule As String = "") As Long
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : OrderedColl
   ' * Module Filename  : VBDocOrderedColl.cls
   ' * Procedure Name   : Count
   ' * Purpose          :
   ' * Parameters       :
   ' *                    Optional RefModule As String = ""
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Dim lngCount         As Long
   Dim objTemp          As Object

   If Len(RefModule) = 0 Then
      Count = mCol.Count
   Else
      lngCount = 0
      For Each objTemp In mCol
         If objTemp.RefModule = RefModule Then
            lngCount = lngCount + 1
         End If
      Next
      Count = lngCount
   End If
End Function

'Purpose: Get the Item
Public Property Get Item(vntIndexKey As Variant) As Object
Attribute Item.VB_UserMemId = 0
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : OrderedColl
   ' * Module Filename  : VBDocOrderedColl.cls
   ' * Procedure Name   : Item
   ' * Purpose          :
   ' * Parameters       :
   ' *                    vntIndexKey As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set Item = mCol(vntIndexKey)
End Property

'Purpose: remove the Item
Public Sub Remove(vntIndexKey As Variant)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : OrderedColl
   ' * Module Filename  : VBDocOrderedColl.cls
   ' * Procedure Name   : Remove
   ' * Purpose          :
   ' * Parameters       :
   ' *                    vntIndexKey As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   mCol.Remove vntIndexKey
End Sub

'Purpose: This prpperty allows to enumerate this collection
'with the For...Each syntax
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : OrderedColl
   ' * Module Filename  : VBDocOrderedColl.cls
   ' * Procedure Name   : NewEnum
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set NewEnum = mCol.[_NewEnum]
End Property

'Purpose: Creates the collection when this class is created
Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : OrderedColl
   ' * Module Filename  : VBDocOrderedColl.cls
   ' * Procedure Name   : Class_Initialize
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set mCol = New Collection
End Sub

'Purpose: destroys collection when this class is terminated
Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Marco Pipino
   ' * Date             : 09/25/2002
   ' * Time             : 14:19
   ' * Module Name      : OrderedColl
   ' * Module Filename  : VBDocOrderedColl.cls
   ' * Procedure Name   : Class_Terminate
   ' * Purpose          :
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************
   Set mCol = Nothing
End Sub

