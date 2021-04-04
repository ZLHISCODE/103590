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
'��д           lshuo
'����           2019/2/13
'ģ��           clsServer
'˵��           ��������
'==================================================================================================
Private mstrServer                          As String                       '��������ʹ�õķ�������IP:Port/SID
Private mjtLastJob                          As JobType                      '�����쳣����������
Private mstrExceptionInfo                   As String                       '�����쳣����Ϣ
Private mlngTryTimes                        As Long                         '�����쳣�Ĵ���
Private mlngLastTry                         As Long                         '��һ������ʱ��
Private mlngFirstTry                        As Long                         '��һ������ʱ��
Private mblnDelay                           As Long                         '�ӳ�ִ�еı�־
Private mlngDelay                           As Long                         '�ӳٿ�ʼ
'===========================================================================
'                2�����Զ���
'===========================================================================
'Server(R/W):��ǰ������
Public Property Let Server(ByVal strServer As String)
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServer.Server", strServer)
    mstrServer = strServer
    Call Logger.PopMethod("ZLHelperMain.clsServer.Server")
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServer.Server") = 1 Then
        Resume
    End If
End Property

Public Property Get Server() As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServer.Server")
    Server = mstrServer
    Call Logger.PopMethod("ZLHelperMain.clsServer.Server", Server)
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServer.Server") = 1 Then
        Resume
    End If
End Property
'IsDelay(R/W):�Ƿ��ӳ�ִ��
Public Property Let IsDelay(ByVal blnDelay As Boolean)
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServer.IsDelay", blnDelay)
    mblnDelay = blnDelay
    If mblnDelay Then
        mlngDelay = GetTickCount()
    Else
        mlngDelay = 0
    End If
    Call Logger.PopMethod("ZLHelperMain.clsServer.IsDelay")
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServer.IsDelay") = 1 Then
        Resume
    End If
End Property

Public Property Get IsDelay() As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServer.IsDelay")
    IsDelay = mblnDelay
    Call Logger.PopMethod("ZLHelperMain.clsServer.IsDelay", IsDelay)
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServer.IsDelay") = 1 Then
        Resume
    End If
End Property
'ExceptionJobType(R):�����쳣����������
Public Property Get ExceptionJobType() As JobType
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServer.ExceptionJobType")
    ExceptionJobType = mjtLastJob
    Call Logger.PopMethod("ZLHelperMain.clsServer.ExceptionJobType", ExceptionJobType)
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServer.ExceptionJobType") = 1 Then
        Resume
    End If
End Property
'TryTimes(R):��ͬ�쳣�����Դ���
Public Property Get TryTimes() As Long
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServer.TryTimes")
    TryTimes = mlngTryTimes
    Call Logger.PopMethod("ZLHelperMain.clsServer.TryTimes")
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServer.TryTimes") = 1 Then
        Resume
    End If
End Property
'LastTry(W):�ϴ�����ʱ��
Public Property Let LastTry(ByVal lngLastTry As Long)
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsException.LastTry")
    mlngLastTry = lngLastTry
    Call Logger.PopMethod("ZLHelperMain.clsException.LastTry")
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsException.LastTry") = 1 Then
        Resume
    End If
End Property

'FirstTry(W):��һ������ʱ��
Public Property Let FirstTry(ByVal lngFirstTry As Long)
    mlngFirstTry = lngFirstTry
End Property
'IsCanTryAgain(R):�Ƿ�����ٴ�����
Public Property Get IsCanTryAgain() As Boolean
    Dim lngTimes        As Long
    'log(86400000/2/10000+3)=8
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsException.IsCanTryAgain")
    '�ӳ������ӣ�����ZLHIS����ʱ���
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
    Call Logger.PopMethod("ZLHelperMain.clsException.IsCanTryAgain", IsCanTryAgain)
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsException.IsCanTryAgain") = 1 Then
        Resume
    End If
End Property
'===========================================================================
'                3����������
'===========================================================================

'--------------------------------------------------------------------------------------------------
'����           RaiseSameException
'����           ����һ����ͬ�쳣
'����ֵ         Boolean                 �Ƿ����һ���쳣��ͬ
'����б�:
'������         ����                    ˵��
'jtJobType      JobType                 ��������
'strExceptionInfo   String              �쳣����
'blnDelay       Boolean                 �Ƿ����ӳ�ִ��
'-------------------------------------------------------------------------------------------------
Public Function RaiseSameException(ByVal jtJobType As JobType, ByVal strExceptionInfo As String, Optional ByVal blnDetailCheck As Boolean) As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsException.RaiseException", jtJobType, strExceptionInfo)
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
    Call Logger.PopMethod("ZLHelperMain.clsException.RaiseException")
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsException.RaiseException") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'����           Restart
'����           ��������������
'����ֵ
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Sub Restart()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsException.Restart")
    mlngTryTimes = 0
    mlngLastTry = 0
    mlngFirstTry = 0
    mjtLastJob = 0
    mblnDelay = False
    mlngDelay = 0
    mstrExceptionInfo = ""
    Call Logger.PopMethod("ZLHelperMain.clsException.Restart")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsException.Restart") = 1 Then
        Resume
    End If
End Sub
'===========================================================================
'                4��˽�з���
'===========================================================================
Private Sub Class_Initialize()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServer.Class_Initialize")
    Call Logger.PopMethod("ZLHelperMain.clsServer.Class_Initialize")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsServer.Class_Initialize") = 1 Then
        Resume
    End If
End Sub

