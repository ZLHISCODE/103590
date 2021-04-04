VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/2/13
'模块           clsServer
'说明           服务器类
'==================================================================================================
Private mstrServer                          As String                       '进程启动使用的服务器。IP:Port/SID
Private mjtLastJob                          As JobType                      '产生异常的任务类型
Private mstrExceptionInfo                   As String                       '产生异常的信息
Private mlngTryTimes                        As Long                         '产生异常的次数
Private mlngLastTry                         As Long                         '上一次重试时间
Private mlngFirstTry                        As Long                         '第一次重试时间
Private mblnDelay                           As Long                         '延迟执行的标志
Private mlngDelay                           As Long                         '延迟开始
'===========================================================================
'                2、属性定义
'===========================================================================
'Server(R/W):当前服务器
Public Property Let Server(ByVal strServer As String)
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsServer.Server", strServer)
    mstrServer = strServer
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsServer.Server")
    Exit Property
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsServer.Server") = 1 Then
        Resume
    End If
End Property

Public Property Get Server() As String
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsServer.Server")
    Server = mstrServer
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsServer.Server", Server)
    Exit Property
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsServer.Server") = 1 Then
        Resume
    End If
End Property
'IsDelay(R/W):是否延迟执行
Public Property Let IsDelay(ByVal blnDelay As Boolean)
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsServer.IsDelay", blnDelay)
    mblnDelay = blnDelay
    If mblnDelay Then
        mlngDelay = GetTickCount()
    Else
        mlngDelay = 0
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsServer.IsDelay")
    Exit Property
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsServer.IsDelay") = 1 Then
        Resume
    End If
End Property

Public Property Get IsDelay() As Boolean
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsServer.IsDelay")
    IsDelay = mblnDelay
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsServer.IsDelay", IsDelay)
    Exit Property
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsServer.IsDelay") = 1 Then
        Resume
    End If
End Property
'ExceptionJobType(R):产生异常的任务类型
Public Property Get ExceptionJobType() As JobType
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsServer.ExceptionJobType")
    ExceptionJobType = mjtLastJob
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsServer.ExceptionJobType", ExceptionJobType)
    Exit Property
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsServer.ExceptionJobType") = 1 Then
        Resume
    End If
End Property
'TryTimes(R):相同异常的重试次数
Public Property Get TryTimes() As Long
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsServer.TryTimes")
    TryTimes = mlngTryTimes
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsServer.TryTimes")
    Exit Property
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsServer.TryTimes") = 1 Then
        Resume
    End If
End Property
'LastTry(W):上次重试时间
Public Property Let LastTry(ByVal lngLastTry As Long)
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsException.LastTry")
    mlngLastTry = lngLastTry
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsException.LastTry")
    Exit Property
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsException.LastTry") = 1 Then
        Resume
    End If
End Property

'FirstTry(W):第一次重试时间
Public Property Let FirstTry(ByVal lngFirstTry As Long)
    mlngFirstTry = lngFirstTry
End Property
'IsCanTryAgain(R):是否可以再次重试
Public Property Get IsCanTryAgain() As Boolean
    Dim lngTimes        As Long
    'log(86400000/2/10000+3)=8
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsException.IsCanTryAgain")
    '延迟两分钟，存在ZLHIS进程时检查
    If mblnDelay Then
        If GetTickCountDiff(mlngDelay) > 120000 Then
            mblnDelay = False
            mlngDelay = 0
        End If
    End If
    If Not mblnDelay Then
        lngTimes = CLng(Log(GetTickCountDiff(mlngFirstTry) / 10000 + 3))
        If mlngTryTimes > G_LNG_MAX_JOBTRY Then
            IsCanTryAgain = False
        Else
            IsCanTryAgain = lngTimes > mlngTryTimes
        End If
    End If
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsException.IsCanTryAgain", IsCanTryAgain)
    Exit Property
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsException.IsCanTryAgain") = 1 Then
        Resume
    End If
End Property
'===========================================================================
'                3、公共方法
'===========================================================================

'--------------------------------------------------------------------------------------------------
'方法           RaiseSameException
'功能           生成一个相同异常
'返回值         Boolean                 是否和上一个异常相同
'入参列表:
'参数名         类型                    说明
'jtJobType      JobType                 任务类型
'strExceptionInfo   String              异常描述
'blnDelay       Boolean                 是否是延迟执行
'-------------------------------------------------------------------------------------------------
Public Function RaiseSameException(ByVal jtJobType As JobType, ByVal strExceptionInfo As String, Optional ByVal blnDetailCheck As Boolean) As Boolean
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsException.RaiseException", jtJobType, strExceptionInfo)

    If (jtJobType = mjtLastJob And Not blnDetailCheck) Or (jtJobType = mjtLastJob And mstrExceptionInfo = strExceptionInfo And blnDetailCheck) Then
        mlngTryTimes = mlngTryTimes + 1
        RaiseSameException = True
        mlngLastTry = GetTickCount()
    Else
        mlngTryTimes = 0
        mlngLastTry = GetTickCount()
        mlngFirstTry = GetTickCount()
        mjtLastJob = jtJobType
    End If
    mstrExceptionInfo = strExceptionInfo
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsException.RaiseException")
    Exit Function
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsException.RaiseException") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           Restart
'功能           清理服务器缓存
'返回值
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Public Sub Restart()
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsException.Restart")
    mlngTryTimes = 0
    mlngLastTry = 0
    mlngFirstTry = 0
    mjtLastJob = 0
    mblnDelay = False
    mlngDelay = 0
    mstrExceptionInfo = ""
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsException.Restart")
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsException.Restart") = 1 Then
        Resume
    End If
End Sub
'===========================================================================
'                4、私有方法
'===========================================================================
Private Sub Class_Initialize()
    On Error GoTo ErrH
    Call gobjLog.PushMethod(RLL_AllLog, "ZLHelperMain.clsServer.Class_Initialize")
    Call gobjLog.PopMethod(RLL_AllLog, "ZLHelperMain.clsServer.Class_Initialize")
    Exit Sub
ErrH:
    If gobjLog.ErrCenter(RLL_RunError, "ZLHelperMain.clsServer.Class_Initialize") = 1 Then
        Resume
    End If
End Sub


