Attribute VB_Name = "mTimer"
Option Explicit

'/~ core timer routines ~/

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal nIDEvent As Long) As Long


Private m_lTimer        As Long

Private Sub Timer_Proc(ByVal lngHwnd As Long, _
                       ByVal nIDEvent As Long, _
                       ByVal uElapse As Long, _
                       ByVal lpTimerFunc As Long)

'/* timer return

On Error GoTo Handler

    With mScheduler
        '/* track idle time
        .Get_Idle
        '/* check job status
        .Schedule_Check
    End With

Exit Sub

Handler:
    Kill_Timer

End Sub

Public Sub Start_Timer(ByVal lInterval As Long)
'/* start the timer

    m_lTimer = SetTimer(0&, 0&, lInterval, AddressOf Timer_Proc)

End Sub

Public Sub Kill_Timer()
'/* kill timer

    KillTimer 0&, m_lTimer

End Sub
