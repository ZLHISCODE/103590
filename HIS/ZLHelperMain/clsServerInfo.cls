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
'@ģ�� clsServerInfoInfo-2019/7/2
'@��д lshuo
'@����
'   ��������
'@����
'
'@��ע
'
Option Explicit
'---------------------------------------------------------------------------
'                0��API�ͳ�������
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                1���������
'---------------------------------------------------------------------------
Private mblnFirst                           As Boolean
'---------------------------------------------------------------------------
'                2�����Ա����붨��
'---------------------------------------------------------------------------
Private mstrServer                          As String                       '��������ʹ�õķ�������IP:Port/SID
Private mjtLastJob                          As JobType                      '�����쳣����������
Private mstrExceptionInfo                   As String                       '�����쳣����Ϣ
Private mlngTryTimes                        As Long                         '�����쳣�Ĵ���
Private mlngLastTry                         As Long                         '��һ������ʱ��
Private mlngFirstTry                        As Long                         '��һ������ʱ��
Private mblnDelay                           As Boolean                      '�ӳ�ִ�еı�־
Private mlngDelay                           As Long                         '�ӳٿ�ʼ
Private mblnDelete                          As Boolean                      '��ǰ�������Ƿ����ɾ��
'Server(R/W):��ǰ������
Public Property Let Server(ByVal strServer As String)
    mstrServer = strServer
End Property

Public Property Get Server() As String
    Server = mstrServer
End Property
'IsDelay(R/W):�Ƿ��ӳ�ִ��
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
'ExceptionJobType(R):�����쳣����������
Public Property Get ExceptionJobType() As JobType
    ExceptionJobType = mjtLastJob
End Property
'TryTimes(R):��ͬ�쳣�����Դ���
Public Property Get TryTimes() As Long
    TryTimes = mlngTryTimes
End Property
'LastTry(W):�ϴ�����ʱ��
Public Property Let LastTry(ByVal lngLastTry As Long)
    mlngLastTry = lngLastTry
End Property

'FirstTry(W):��һ������ʱ��
Public Property Let FirstTry(ByVal lngFirstTry As Long)
    mlngFirstTry = lngFirstTry
End Property
'IsCanTryAgain(R):�Ƿ�����ٴ�����
Public Property Get IsCanTryAgain() As Boolean
    Dim lngTimes        As Long
    'log(86400000/2/100000+3)=8
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsServerInfo.IsCanTryAgain")
    '�ӳ������ӣ�����ZLHIS����ʱ���
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
'IsCanDeleteServer,�Ƿ����ɾ��������
Public Property Get IsCanDeleteServer() As Boolean
    IsCanDeleteServer = mblnDelete
End Property
'---------------------------------------------------------------------------
'                3����������
'---------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
'����           RaiseSameException
'����
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'jtJobType      JobType
'strExceptionInfo   String              �쳣����
'blnDelay       Boolean                 �Ƿ����ӳ�ִ��
'-------------------------------------------------------------------------------------------------

'@����    RaiseSameException
'   ����һ����ͬ�쳣
'@����ֵ  Boolean
'   �Ƿ����һ���쳣��ͬ
'@����:
'jtJobType JobType In
'   ��������
'strExceptionInfo String In
'   �쳣����
'blnDetailCheck Boolean In (Optional)
'   �Ƿ����쳣����
'@��ע
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
'@����    Restart
'   ��������������
'@����ֵ
'
'@����:
'@��ע
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
'                4��˽�з���
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                5�����󷽷����¼�
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