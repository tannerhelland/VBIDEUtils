VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Memory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 08/11/1999
' * Time             : 10:57
' * Module Name      : class_Memory
' * Module Filename  : Memory.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

' ==========================================================================
' Class:    class_Memory
' Filename: class_Memory.cls
' Author:   Steve McMahon
' Date:     24 May 1998
'
' A class for manipulating API memory blocks.
' ==========================================================================

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Enum EMemoryFlags
   GMEM_DDESHARE = &H2000
   GMEM_DISCARDABLE = &H100
   GMEM_DISCARDED = &H4000
   GMEM_INVALID_HANDLE = &H8000
   GMEM_FIXED = &H0
   GMEM_LOCKCOUNT = &HFF
   GMEM_MODIFY = &H80
   GMEM_MOVEABLE = &H2
   GMEM_NODISCARD = &H20
   GMEM_NOCOMPACT = &H10
   GMEM_NOT_BANKED = &H1000
   GMEM_LOWER = GMEM_NOT_BANKED
   GMEM_NOTIFY = &H4000
   GMEM_SHARE = &H2000
   GMEM_VALID_FLAGS = &H7F72
   GMEM_ZEROINIT = &H40
   GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
End Enum

Private m_hMem          As Long
Private m_lPtr          As Long

Public Property Get Handle() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_Memory
   ' * Module Filename  : Memory.cls
   ' * Procedure Name   : Handle
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Handle = m_hMem
End Property

Public Property Let Handle(ByVal hMem As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_Memory
   ' * Module Filename  : Memory.cls
   ' * Procedure Name   : Handle
   ' * Parameters       :
   ' *                    ByVal hMem As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If (m_hMem <> 0) Then
      FreeMemory
   End If
   m_hMem = hMem
End Property

Public Property Get Pointer() As Long
   If (m_hMem <> 0) Then
      If (m_lPtr = 0) Then
         LockMemory
      End If
      Pointer = m_lPtr
   End If
End Property

Public Property Get Size() As Long
   If (m_hMem <> 0) Then
      Size = GlobalSize(m_hMem)
   End If
End Property

Public Function AllocateMemory( _
   ByVal lSize As Long, _
   Optional ByVal dwFlags As Long = GPTR _
   ) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_Memory
   ' * Module Filename  : Memory.cls
   ' * Procedure Name   : AllocateMemory
   ' * Parameters       :
   ' *                    ByVal lSize As Long
   ' *                    Optional ByVal dwFlags As Long = GPTR
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   FreeMemory
   m_hMem = GlobalAlloc(dwFlags, lSize)
   If (m_hMem <> 0) Then
      ' Success
      AllocateMemory = True
   Else
      ' Failed
   End If
End Function

Public Function LockMemory() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_Memory
   ' * Module Filename  : Memory.cls
   ' * Procedure Name   : LockMemory
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If (m_hMem <> 0) Then
      If (m_lPtr = 0) Then
         m_lPtr = GlobalLock(m_hMem)
         If (m_lPtr <> 0) Then
            ' Success
            LockMemory = True
         Else
            ' Failed
         End If
      End If
   End If
End Function

Public Sub UnlockMemory()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_Memory
   ' * Module Filename  : Memory.cls
   ' * Procedure Name   : UnlockMemory
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If (m_hMem <> 0) Then
      If (m_lPtr <> 0) Then
         GlobalUnlock m_hMem
         m_lPtr = 0
      End If
   End If
End Sub

Public Sub FreeMemory()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_Memory
   ' * Module Filename  : Memory.cls
   ' * Procedure Name   : FreeMemory
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   If (m_hMem <> 0) Then
      UnlockMemory
      GlobalFree m_hMem
   End If
   m_hMem = 0
End Sub

Public Sub ReleaseDontFreeMemory()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_Memory
   ' * Module Filename  : Memory.cls
   ' * Procedure Name   : ReleaseDontFreeMemory
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   ' For GMEM_DDESHARE operations...
   UnlockMemory
   m_hMem = 0
End Sub

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 08/11/1999
   ' * Time             : 10:57
   ' * Module Name      : class_Memory
   ' * Module Filename  : Memory.cls
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   FreeMemory
End Sub

