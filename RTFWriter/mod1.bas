Attribute VB_Name = "mod1"
Option Explicit

Const cTimerMax = 100

' Array of timers
Public aTimers(1 To cTimerMax) As cTimer
' Added SPM to prevent excessive searching through aTimers array:
Private m_cTimerCount As Integer

Function TimerCreate(timer As cTimer) As Boolean
    ' Create the timer
    timer.TimerID = SetTimer(0&, 0&, timer.Interval, AddressOf TimerProc)
    If timer.TimerID Then
        TimerCreate = True
        Dim i As Integer
        For i = 1 To cTimerMax
            If aTimers(i) Is Nothing Then
                Set aTimers(i) = timer
                If (i > m_cTimerCount) Then
                    m_cTimerCount = i
                End If
                TimerCreate = True
                Exit Function
            End If
        Next
        timer.ErrRaise eeTooManyTimers
    Else
        ' TimerCreate = False
        timer.TimerID = 0
        timer.Interval = 0
    End If
End Function

Public Function TimerDestroy(timer As cTimer) As Long
    ' TimerDestroy = False
    ' Find and remove this timer
    Dim i As Integer, f As Boolean
    For i = 1 To cTimerMax
        ' Find timer in array
        If Not aTimers(i) Is Nothing Then
            If timer.TimerID = aTimers(i).TimerID Then
                f = KillTimer(hNull, timer.TimerID)
                ' Remove timer and set reference to nothing
                Set aTimers(i) = Nothing
                TimerDestroy = True
                Exit Function
            End If
        
           ' SPM: aTimers(1) could well be nothing before
        ' aTimers(2) is.  This original [else] would leave
        ' timer 2 still running when the class terminates -
        ' not very nice!  Causes serious GPF in IE and VB design
        ' mode...
        'Else
        '    TimerDestroy = True
        '    Exit Function
        End If
    Next
End Function

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, _
                     ByVal idEvent As Long, ByVal dwTime As Long)
    Dim i As Integer
    ' Find the timer with this ID
    For i = 1 To m_cTimerCount
        ' SPM: Add a check to ensure aTimers(i) is not nothing!
        ' This would occur if we had two timers declared from
        ' the same thread and we terminated the first one before
        ' the second!  Causes serious GPF if we don't do this...
        If Not (aTimers(i) Is Nothing) Then
            If idEvent = aTimers(i).TimerID Then
                ' Generate the event
                aTimers(i).PulseTimer
                Exit Sub
            End If
        End If
    Next
End Sub

Private Function StoreTimer(timer As cTimer)
    Dim i As Integer
    For i = 1 To m_cTimerCount
        If aTimers(i) Is Nothing Then
            Set aTimers(i) = timer
            StoreTimer = True
            Exit Function
        End If
    Next
End Function



