VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOSOutputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Author           : Marco Pipino
' * Date             : 09/25/2002
' * Time             : 14:19
' * Module Name      : DOSOutputs
' * Module Filename  : VBDocDOSOutputs.cls
' * Purpose          :
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

'Purpose: This object it's used for getting the output of a DOS Application
'   It don't work with 16-bit application on Win9x and WinME
'Author:    Marco Pipino

Option Explicit

'Purpose: The CreatePipe function creates an anonymous pipe,
'and returns handles to the read and write ends of the pipe.
Private Declare Function CreatePipe Lib "kernel32" ( _
   phReadPipe As Long, _
   phWritePipe As Long, _
   lpPipeAttributes As Any, _
   ByVal nSize As Long) As Long

'Purpose: Used to read the the pipe filled by the process create
'with the CretaProcessA function
Private Declare Function ReadFile Lib "kernel32" ( _
   ByVal hFile As Long, _
   ByVal lpBuffer As String, _
   ByVal nNumberOfBytesToRead As Long, _
   lpNumberOfBytesRead As Long, _
   ByVal lpOverlapped As Any) As Long

'Purpose: Structure used by the CreateProcessA function
Private Type SECURITY_ATTRIBUTES
   nLength              As Long
   lpSecurityDescriptor As Long
   bInheritHandle       As Long
End Type

'Purpose: Structure used by the CreateProcessA function
Private Type STARTUPINFO
   cb                   As Long
   lpReserved           As Long
   lpDesktop            As Long
   lpTitle              As Long
   dwX                  As Long
   dwY                  As Long
   dwXSize              As Long
   dwYSize              As Long
   dwXCountChars        As Long
   dwYCountChars        As Long
   dwFillAttribute      As Long
   dwFlags              As Long
   wShowWindow          As Integer
   cbReserved2          As Integer
   lpReserved2          As Long
   hStdInput            As Long
   hStdOutput           As Long
   hStdError            As Long
End Type

'Purpose: Structure used by the CreateProcessA function
Private Type PROCESS_INFORMATION
   hProcess             As Long
   hThread              As Long
   dwProcessID          As Long
   dwThreadID           As Long
End Type

'Purpose: This function launch the the commend and return the relative process
'into the PRECESS_INFORMATION structure
Private Declare Function CreateProcessA Lib "kernel32" ( _
   ByVal lpApplicationName As Long, _
   ByVal lpCommandLine As String, _
   lpProcessAttributes As SECURITY_ATTRIBUTES, _
   lpThreadAttributes As SECURITY_ATTRIBUTES, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long

'Purpose: Close opened handle
Private Declare Function CloseHandle Lib "kernel32" ( _
   ByVal hHandle As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
   ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long

'Purpose: Consts for the above functions
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const INFINITE = &HFFFF

Private mCommand        As String          'Private variable for the CommandLine property
Private mOutputs        As String          'Private variable for the ReadOnly Outputs property

'Purpose: Event that notify the temporary buffer to the object
Public Event ReceiveOutputs(CommandOutputs As String)

'Purpose: This property set and get the DOS command line
'It's possible to set this property directly from the
'parameter of the ExecuteCommand method
Public Property Get CommandLine() As String
   CommandLine = mCommand
End Property

Public Property Let CommandLine(DOSCommand As String)
   mCommand = DOSCommand
End Property

'Purpose: This property ReadOnly get the complete output after
'a command execution
Public Property Get Outputs()
   Outputs = mOutputs
End Property

Public Function ExecuteCommand(Optional CommandLine As String) As Boolean
   Dim proc             As PROCESS_INFORMATION     'Process info filled by CreateProcessA
   Dim ret              As Long                     'long variable for get the return value of the
   'API functions
   Dim start            As STARTUPINFO            'StartUp Info passed to the CreateProceeeA
   'function
   Dim sa               As SECURITY_ATTRIBUTES       'Security Attributes passeed to the
   'CreateProcessA function
   Dim hReadPipe        As Long               'Read Pipe handle created by CreatePipe
   Dim hWritePipe       As Long              'Write Pite handle created by CreatePipe
   Dim lngBytesread     As Long            'Amount of byte read from the Read Pipe handle
   Dim strBuff          As String * 256         'String buffer reading the Pipe

   'Return false if the command line is NOT termined correctly
   ExecuteCommand = False

   'if the parameter is not empty update the CommandLine property
   If Len(CommandLine) > 0 Then
      mCommand = CommandLine
   End If

   'if the command line is empty then exit whit a error message
   If Len(mCommand) = 0 Then
      MsgBox "Command Line empty", vbCritical
      Exit Function
   End If

   'Create the Pipe
   sa.nLength = Len(sa)
   sa.bInheritHandle = 1&
   sa.lpSecurityDescriptor = 0&
   ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)

   If ret = 0 Then
      'If an error occur during the Pipe creation exit
      MsgBox "CreatePipe failed. Error: " & err.LastDllError, vbCritical
      Exit Function
   End If

   'Launch the command line application
   start.cb = Len(start)
   start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
   'set the StdOutput and the StdError output to the same Write Pipe handle
   start.hStdOutput = hWritePipe
   start.hStdError = hWritePipe
   'Execute the command
   ret = CreateProcessA(0&, mCommand, sa, sa, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
   'WaitForSingleObject proc.hProcess, INFINITE

   If ret <> 1 Then
      'if the command is not found ....
      MsgBox "File or command not found", vbCritical
      Exit Function
   End If

   'Now We can ... must close the hWritePipe
   ret = CloseHandle(hWritePipe)
   mOutputs = ""

   'Read the ReadPipe handle
   Do
      ret = ReadFile(hReadPipe, strBuff, 256, lngBytesread, 0&)
      mOutputs = mOutputs & left$(strBuff, lngBytesread)
      'Send data to the object via ReceiveOutputs event
   Loop While ret <> 0
   RaiseEvent ReceiveOutputs(mOutputs)

   'Close the opened handles
   ret = CloseHandle(proc.hProcess)
   ret = CloseHandle(proc.hThread)
   ret = CloseHandle(hReadPipe)

   'Return true if the command line is termined correctly
   ExecuteCommand = True
End Function

