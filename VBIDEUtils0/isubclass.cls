VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EMsgResponse
    emrConsume      ' Process instead of original WindowProc
    emrPostProcess  ' Process after original WindowProc
    emrPreprocess   ' Process before original WindowProc
End Enum

Public MsgResponse As EMsgResponse
Attribute MsgResponse.VB_VarDescription = "Get/set how you will respond to a message.  Raised whenever a message is encountered.  Call the CurrentMessage property to determine what message you are currently responding to."

Function WindowProc(ByVal hwnd As Long, _
                    ByVal iMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long
Attribute WindowProc.VB_Description = "Raised whenever a message you have subclassed is sent."
End Function
'

