VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------
'模块名  ：clsTimer
'模块说明：本模块作为全局类，可直接被其它部件声明使用。
'模块内容：一个API方式使用TIMER的方法类
'模块整理：祝庆
'---------------------------------------------------------------------------------
Private lngInterval As Long
Private lngID As Long

Public Item As Variant

Public Event ThatTime()

Public Enum EErrorTimer
    eeBaseTimer = 13650 ' CTimer
    eeTooManyTimers ' No more than 10 timers allowed per class
    eeCantCreateTimer ' Can't create system timer
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
    Err.Raise e, sSource
    End If
End Sub

Property Get Interval() As Long
    Interval = lngInterval
End Property

Property Let Interval(lngIntervalA As Long)
Dim f As Boolean
    If lngIntervalA > 0 Then
    
        If lngInterval = lngIntervalA Then Exit Property
        
        If lngInterval Then
            f = TimerDestroy(Me)
        End If
        
        lngInterval = lngIntervalA
        If TimerCreate(Me) = False Then ErrRaise eeCantCreateTimer
    Else
        If (lngInterval > 0) Then
            lngInterval = 0
            f = TimerDestroy(Me)
        End If
    End If
End Property

Public Sub PulseTimer()
    RaiseEvent ThatTime
End Sub

Friend Property Get TimerID() As Long
    TimerID = lngID
End Property

Friend Property Let TimerID(idA As Long)
    lngID = idA
End Property

Private Sub Class_Terminate()
    Interval = 0
End Sub
