VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private iInterval As Long
Private id As Long

' User can attach any Variant data they want to the timer
Public Item As Variant
Attribute Item.VB_VarDescription = "Gets/sets a piece of additional data to store with the timer."

Public Event ThatTime()
Attribute ThatTime.VB_Description = "Raised when the timer fires."

' SubTimer is independent of VBCore, so it hard codes error handling

Public Enum EErrorTimer
    eeBaseTimer = 13650     ' CTimer
    eeTooManyTimers         ' No more than 10 timers allowed per class
    eeCantCreateTimer       ' Can't create system timer
End Enum

Friend Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
        Case eeTooManyTimers
            sText = "No more than 10 timers allowed per class"
        Case eeCantCreateTimer
            sText = "Can't create system timer"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
End Sub


Property Get Interval() As Long
Attribute Interval.VB_Description = "Gets/sets the interval at which the timer fires.  Set to zero to stop the timer."
    Interval = iInterval
End Property

' Can't just change interval--you must kill timer and start a new one
Property Let Interval(iIntervalA As Long)
    Dim f As Boolean
    If iIntervalA > 0 Then
        ' Don't mess with it if interval is the same
        If iInterval = iIntervalA Then Exit Property
        ' Must destroy any existing timer to change interval
        If iInterval Then
            f = TimerDestroy(Me)
            Debug.Assert f      ' Shouldn't fail
        End If
        ' Create new timer with new interval
        iInterval = iIntervalA
        If TimerCreate(Me) = False Then ErrRaise eeCantCreateTimer
    Else
        If (iInterval > 0) Then
            iInterval = 0
            f = TimerDestroy(Me)
            Debug.Assert f      ' Shouldn't fail
        End If
    End If
End Property

' Must be public so that Timer object can't terminate while client's ThatTime
' event is being processed--Friend wouldn't prevent this disaster
Public Sub PulseTimer()
Attribute PulseTimer.VB_MemberFlags = "40"
    RaiseEvent ThatTime
End Sub

Friend Property Get TimerID() As Long
    TimerID = id
End Property

Friend Property Let TimerID(idA As Long)
    id = idA
End Property

Private Sub Class_Terminate()
    Interval = 0
End Sub

