VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private DataStack()     As Variant
Private lStackSize      As Long
Private lArraySize      As Long

Public Type perLine
   lineID               As String
   lineContent          As String
   lcaseContent         As String
End Type

'alter this to balance resource vs processor usage
Private Const sizeChange = 150
Public Property Get Count() As Long
   Count = lStackSize
End Property

' Returns and removes the item at the top the stack
Public Function Pop() As Variant
   If lStackSize > 0 Then
      lStackSize = lStackSize - 1
      Pop = DataStack(lStackSize)
   Else
      Pop = Empty ' Stack empty
   End If
End Function

'Adds data to the top of the stack
Public Sub Push(data As Variant)
   If lArraySize < lStackSize Then
      lArraySize = lArraySize + sizeChange
      ReDim Preserve DataStack(lArraySize) ' Resize the array
   End If
   DataStack(lStackSize) = data
   lStackSize = lStackSize + 1
End Sub

' Returns the item at the top of the stack
Public Property Get top() As Variant
   If lStackSize > 0 Then
      top = DataStack(lStackSize - 1)
   Else
      top = Empty ' Stack empty
   End If
End Property

Private Sub Class_Initialize()
   lArraySize = sizeChange * 2
   ReDim DataStack(lArraySize)
End Sub
