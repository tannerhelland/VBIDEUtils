VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_FunctionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Programmer Name  : removed
' * Web Site         : http://www.ppreview.net
' * E-Mail           : removed
' * Date             : 15/12/1999
' * Time             : 16:13
' * Module Name      : class_FunctionData
' * Module Filename  : CFuncData.cls
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit

Private m_Modulename    As String
Private m_FunctionName  As String
Private m_intLineNumber As Long

Public Property Get Module() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_FunctionData
   ' * Module Filename  : FunctionData.cls
   ' * Procedure Name   : Module
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   Module = m_Modulename
End Property

Public Property Let Module(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_FunctionData
   ' * Module Filename  : FunctionData.cls
   ' * Procedure Name   : Module
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_Modulename = vNewValue
End Property

Public Property Get FuncName() As String
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_FunctionData
   ' * Module Filename  : FunctionData.cls
   ' * Procedure Name   : FuncName
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   FuncName = m_FunctionName
End Property

Public Property Let FuncName(ByVal vNewValue As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_FunctionData
   ' * Module Filename  : FunctionData.cls
   ' * Procedure Name   : FuncName
   ' * Parameters       :
   ' *                    ByVal vNewValue As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_FunctionName = vNewValue
End Property

Public Property Get LineNumber() As Long
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_FunctionData
   ' * Module Filename  : FunctionData.cls
   ' * Procedure Name   : LineNumber
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   LineNumber = m_intLineNumber
End Property

Public Property Let LineNumber(ByVal vNewValue As Long)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : removed
   ' * Web Site         : http://www.ppreview.net
   ' * E-Mail           : removed
   ' * Date             : 15/12/1999
   ' * Time             : 16:12
   ' * Module Name      : class_FunctionData
   ' * Module Filename  : FunctionData.cls
   ' * Procedure Name   : LineNumber
   ' * Parameters       :
   ' *                    ByVal vNewValue As Long
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' **********************************************************************

   m_intLineNumber = vNewValue
End Property
