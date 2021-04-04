VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@模块 clsServerInfoInfo-2019/7/2
'@编写 lshuo
'@功能
'   服务器类
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
Private mblnFirst                           As Boolean
'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------
Private mstrServer                          As String                       '进程启动使用的服务器。IP:Port/SID
Private mjtLastJob                          As JobType                      '产生异常的任务类型
Private mstrExceptionInfo                   As String                       '产生异常的信息
Private mlngTryTimes                        As Long                         '产生异常的次数
Private mlngLastTry                         As Long                         '上一次重试时间
Private mlngFirstTry                        As Long                         '第一次重试时间
Private mblnDelay                           As Boolean                      '延迟执行的标志
Private mlngDelay                           As Long                         '延迟开始
Private mblnDelete                          As Boolean                      '当前服务器是否可以删除
'Server(R/W):当前服务器
Public Property Let Server(ByVal strServer As String)
    mstrServer = strServer
End Property

Public Property Get Server() As String
    Server = mstrServer
End Property
'IsDelay(R/W):是否延迟执行
Public Property Let IsDelay(ByVal blnDelay As Boolean)
    mblnDelay = blnDelay
    If mblnDelay Then
        mlngDelay = GetTickCount()
    Else
        mlngDelay = 0
    End If
End Property

Public Property Get IsDelay() As Boolean
    IsDelay = mblnDelay
End Property
'ExceptionJobType(R):产生异常的任务类型
Public Property Get ExceptionJobType() As JobType
    ExceptionJobType = mjtLastJob
End Property
'TryTimes(R):相同异常的重试次数
Public Property Get TryTimes() As Long
    TryTimes = mlngTryTimes
End Property
'LastTry(W):上次重试时间
Public Property Let LastTry(ByVal lngLastTry As Long)
    mlngLastTry = lngLastTry
End Property

'FirstTry(W):第一次重试时间
Public Property Let FirstTry(ByVal lngFirstTry As Long)
    mlngFirstTry = lngFirstTry
End Property
'IsCanTryAgain(R):是否可以再次重试
Public Property Get IsCanTryAgain() As Boolean
    Dim lngTimes        As Long
    'log(86400000/2/100000+3)=8
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServerInfo.IsCanTryAgain")
    '延迟两分钟，存在ZLHIS进程时检查
    If mblnDelay Then
        If GetTickCountDiff(mlngDelay) > 120000 Then
            mblnDelay = False
            mlngDelay = 0
        End If
    End If
    If Not mblnDelay Then
        lngTimes = CLng(Log(GetTickCountDiff(mlngFirstTry) / 60000 + 3))
        If mlngTryTimes > G_LNG_MAX_JOBTRY Then
            IsCanTryAgain = False
            mblnDelete = True
        Else
            IsCanTryAgain = lngTimes > mlngTryTimes
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.IsCanTryAgain", mblnDelay, mlngDelay, mlngFirstTry, lngTimes, IsCanTryAgain, mblnFirst)
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServerInfo.IsCanTryAgain") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.IsCanTryAgain")
End Property
'IsCanDeleteServer,是否可以删除服务器
Public Property Get IsCanDeleteServer() As Boolean
    IsCanDeleteServer = mblnDelete
End Property
'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
'方法           RaiseSameException
'功能
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'jtJobType      JobType
'strExceptionInfo   String              异常描述
'blnDelay       Boolean                 是否是延迟执行
'-------------------------------------------------------------------------------------------------

'@方法    RaiseSameException
'   生成一个相同异常
'@返回值  Boolean
'   是否和上一个异常相同
'@参数:
'jtJobType JobType In
'   任务类型
'strExceptionInfo String In
'   异常描述
'blnDetailCheck Boolean In (Optional)
'   是否检查异常描述
'@备注
'
Public Function RaiseSameException(ByVal jtJobType As JobType, ByVal strExceptionInfo As String, Optional ByVal blnDetailCheck As Boolean) As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServerInfo.RaiseSameException", jtJobType, strExceptionInfo, mblnFirst)
    If Not mblnFirst Then
        mlngTryTimes = mlngTryTimes + 1
        RaiseSameException = True
        mlngLastTry = GetTickCount()
    Else
        mlngTryTimes = 0
        mlngLastTry = GetTickCount()
        mlngFirstTry = GetTickCount()
        mjtLastJob = jtJobType
        mblnFirst = False
    End If
    mstrExceptionInfo = strExceptionInfo
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.RaiseSameException", mblnFirst, mlngTryTimes, mlngLastTry, mjtLastJob)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServerInfo.RaiseSameException") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.RaiseSameException")
End Function
'@方法    Restart
'   清理服务器缓存
'@返回值
'
'@参数:
'@备注
'
Public Sub Restart()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServerInfo.Restart")
    mlngTryTimes = 0
    mlngLastTry = 0
    mlngFirstTry = 0
    mjtLastJob = 0
    mblnDelay = False
    mlngDelay = 0
    mstrExceptionInfo = ""
    mblnFirst = True
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.Restart")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServerInfo.Restart") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.Restart")
End Sub
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServerInfo.Class_Initialize")
    mblnFirst = True
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.Class_Initialize", mblnFirst)
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServerInfo.Class_Initialize") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.Class_Initialize")
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServerInfo.Class_Terminate")
    
    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.Class_Terminate")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServerInfo.Class_Terminate") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsServerInfo.Class_Terminate")
End Sub
