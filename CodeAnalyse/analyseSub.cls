VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 03/11/1999
' * Time             : 15:01
' * Module Name      : class_Sub
' * Module Filename  : class_Sub.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private mvarSubType     As String   'local copy
Private mvarSubName     As String   'local copy
Private mvarNumLines    As Long     'local copy

Public mcolVariable     As New Collection
Public mcolUnusedVar    As New Collection

Public Property Let NumLines(ByVal vData As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Sub
   ' * Module Filename  : class_Sub.cls
   ' * Procedure Name   : NumLines
   ' * Parameters       :
   ' *                    ByVal vData As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.NumLines = 5
   mvarNumLines = vData
End Property

Public Property Get NumLines() As Long
   'used when retrieving value of a property, on the right side of an assignment.
   NumLines = mvarNumLines
End Property

Public Property Let SubName(ByVal vData As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Sub
   ' * Module Filename  : class_Sub.cls
   ' * Procedure Name   : SubName
   ' * Parameters       :
   ' *                    ByVal vData As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.SubName = 5
   mvarSubName = vData
End Property

Public Property Get SubName() As String
   'used when retrieving value of a property, on the right side of an assignment.
   SubName = mvarSubName
End Property

Public Property Let SubType(ByVal vData As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 03/11/1999
   ' * Time             : 15:01
   ' * Module Name      : class_Sub
   ' * Module Filename  : class_Sub.cls
   ' * Procedure Name   : SubType
   ' * Parameters       :
   ' *                    ByVal vData As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.SubType = 5
   mvarSubType = vData
End Property

Public Property Get SubType() As String
   'used when retrieving value of a property, on the right side of an assignment.
   SubType = mvarSubType
End Property

Private Sub Class_Initialize()

   Set mcolVariable = New Collection
   Set mcolUnusedVar = New Collection

End Sub

Private Sub Class_Terminate()

   Set mcolVariable = Nothing
   Set mcolUnusedVar = Nothing

End Sub
