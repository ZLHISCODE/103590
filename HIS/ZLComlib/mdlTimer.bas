Attribute VB_Name = "mdlTimer"
Option Explicit

'---------------------------------------------------------------------------------
'模块名  ：mdlTimer
'模块说明：本模块作为全局类，可直接被其它部件声明使用。
'模块内容：一个API方式使用TIMER的方法模块
'模块整理：祝庆
'---------------------------------------------------------------------------------

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Const cTimerMax = 100

Public aTimers(1 To cTimerMax) As clsTimer
Private m_cTimerCount As Integer

Function TimerCreate(timer As clsTimer) As Boolean
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
        timer.TimerID = 0
        timer.Interval = 0
    End If
End Function

Public Function TimerDestroy(timer As clsTimer) As Long
    Dim i As Integer, f As Boolean

    For i = 1 To m_cTimerCount
        If Not aTimers(i) Is Nothing Then
            If timer.TimerID = aTimers(i).TimerID Then
                f = KillTimer(0, timer.TimerID)
                Set aTimers(i) = Nothing
                TimerDestroy = True
                Exit Function
            End If
        End If
    Next
End Function

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim i As Integer
    For i = 1 To m_cTimerCount
        If Not (aTimers(i) Is Nothing) Then
            If idEvent = aTimers(i).TimerID Then
                aTimers(i).PulseTimer
                Exit Sub
            End If
        End If
    Next
End Sub

Private Function StoreTimer(timer As clsTimer)
    Dim i As Integer
    For i = 1 To m_cTimerCount
        If aTimers(i) Is Nothing Then
            Set aTimers(i) = timer
            StoreTimer = True
            Exit Function
        End If
    Next
End Function


