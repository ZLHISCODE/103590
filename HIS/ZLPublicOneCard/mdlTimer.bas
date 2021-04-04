Attribute VB_Name = "mdlTimer"
Option Explicit

'---------------------------------------------------------------------------------
'ģ����  ��mdlTimer
'ģ��˵������ģ����Ϊȫ���࣬��ֱ�ӱ�������������ʹ�á�
'ģ�����ݣ�һ��API��ʽʹ��TIMER�ķ���ģ��
'ģ������ף��
'---------------------------------------------------------------------------------

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Const cTimerMax = 100

Public aTimers(1 To cTimerMax) As clsTimer
Private m_cTimerCount As Integer

Function TimerCreate(timer As clsTimer) As Boolean
    timer.TimerID = SetTimer(0&, 0&, timer.Interval, AddressOf TimerProc)
    If timer.TimerID Then
        TimerCreate = True
        Dim I As Integer
        For I = 1 To cTimerMax
            If aTimers(I) Is Nothing Then
                Set aTimers(I) = timer
                If (I > m_cTimerCount) Then
                    m_cTimerCount = I
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
    Dim I As Integer, f As Boolean

    For I = 1 To m_cTimerCount
        If Not aTimers(I) Is Nothing Then
            If timer.TimerID = aTimers(I).TimerID Then
                f = KillTimer(0, timer.TimerID)
                Set aTimers(I) = Nothing
                TimerDestroy = True
                Exit Function
            End If
        End If
    Next
End Function

Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim I As Integer
    For I = 1 To m_cTimerCount
        If Not (aTimers(I) Is Nothing) Then
            If idEvent = aTimers(I).TimerID Then
                aTimers(I).PulseTimer
                Exit Sub
            End If
        End If
    Next
End Sub

Private Function StoreTimer(timer As clsTimer)
    Dim I As Integer
    For I = 1 To m_cTimerCount
        If aTimers(I) Is Nothing Then
            Set aTimers(I) = timer
            StoreTimer = True
            Exit Function
        End If
    Next
End Function



