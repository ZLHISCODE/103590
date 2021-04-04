VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mfrmWinsock As frmWinsock
Attribute mfrmWinsock.VB_VarHelpID = -1

Public Event DataArrival(ByVal lngNoticeCode As Long, ByVal intChangeType As Integer, ByVal strTableOwner As String, ByVal TableName As String, ByVal strRowid As String)
Private mlngPort As Long

Private Sub Class_Initialize()
    '��ʼ����ʱ, ����UDP�˿�
    Dim strErrMsg As String
    
    Set mfrmWinsock = New frmWinsock
    
    If StartUdp(strErrMsg) = False Then
        MsgBox """�������ݱ䶯֪ͨ""��������ʧ�ܡ�" & vbNewLine & strErrMsg, , "��ʾ"
    End If
End Sub

Private Sub Class_Terminate()
    Unload mfrmWinsock
    Set mfrmWinsock = Nothing
End Sub

Private Sub mfrmWinsock_DataArrival(ByVal lngNoticeCode As Long, ByVal intChangeType As Integer, ByVal strTableOwner As String, ByVal strTableName As String, ByVal strRowid As String)
    'winsock�յ���Ϣ��,���¼��׳�,������ҵ�񲿼�����
    '���ز��� lngNoticeCode = NoticeCode  intChangeType = �䶯���� strTableOwner = �������� strTableName = ����  strRowid = �䶯��Rowid
    RaiseEvent DataArrival(lngNoticeCode, intChangeType, strTableOwner, strTableName, strRowid)
End Sub

Public Function StartUdp(Optional ByRef strErr As String) As Boolean
    '����UDP����,�ɹ�����True
    '������������з�������,����Flase, strErr  = ��������
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '�����ж���Ϣ�շ����Ƿ��,���û�д�,����Ҫ�����˿�
    If Not IsServerStart Then
        StartUdp = True
        mlngPort = 0
        Exit Function
    End If
    
    strSql = "Select ��Ϣ�˿� From Zltools.Zlclientsession Where �Ự�� = [1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ����̨UDP�˿�", glngSessionID)
    
    If rsTmp.RecordCount = 0 Then
        strErr = "δ�ҵ���ǰ�Ự��Ӧ����Ϣ�˿ڡ�"
    Else
        If IsNull(rsTmp!��Ϣ�˿�) Then
            strErr = "δ�ҵ���ǰ�Ự��Ӧ����Ϣ�˿ڡ�"
        Else
            mlngPort = rsTmp!��Ϣ�˿�   '�˴���ȡ�˿�,���������� "�Ự���"
            
            '�жϲ������Ƿ�������̨�˿�
            Set rsTmp = GetZLOptions(9)
            If rsTmp.RecordCount <> 0 Then
                If rsTmp!����ֵ = "0" Then
                    mlngPort = 0
                    StartUdp = True
                    Exit Function
                End If
            End If
            
            '����˿ڿ���ʧ�� ,�ͽ��˿��޸�Ϊ0
            If mfrmWinsock.StartUdp(mlngPort, strErr) = False Then
                mlngPort = 0
            Else
                StartUdp = True
            End If
        End If
    End If
    Exit Function
errH:
    strErr = Err.Description
End Function

Public Function IsWinsockStart() As Boolean
    '���winsock�Ƿ�ɹ���
    
    IsWinsockStart = mlngPort <> 0
End Function

Private Function IsServerStart() As Boolean
    '�����Ϣ�շ����Ƿ��
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim intState As Integer, lngSid As Long
    
    On Error GoTo errH
    '�����Ϣ�շ���״̬
    Set rsTmp = GetZLOptions(27)
    If rsTmp.RecordCount = 0 Then Exit Function
    If rsTmp!����ֵ & "" = "" Then Exit Function
    
    'IP;�˿�;״̬;�Ự��
    intState = Split(rsTmp!����ֵ, ";")(2)
    lngSid = Split(rsTmp!����ֵ, ";")(3)
    
    If intState = 0 Then Exit Function
    
    strSql = "Select 1 From gv$Session Where AUDSID = [1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "�����Ϣ�շ���״̬", lngSid)
    
    IsServerStart = rsTmp.RecordCount > 0
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Sub ChangePosition(Optional ByVal lng����ID As Long, Optional ByVal str��ǰλ�� As String, Optional ByRef strErr As String)
'���ݲ��������������ʹ������ָ��÷���

    '����ǰ�Ự��Ϣ�������� zlClientSession
    Dim str����վ As String, lng�˿ں� As Long
    Dim strSql As String
    
    On Error GoTo errH
    str����վ = gobjComLib.OS.ComputerName
    lng�˿ں� = mlngPort
    
    strSql = "Zltools.Zl_Zlclientsession_Save('" & str����վ & "',0," & lng�˿ں� & ",'','',''," & lng����ID & ",'" & str��ǰλ�� & "',1)"
    gobjComLib.zlDatabase.ExecuteProcedure strSql, "����zlclientsession��Ϣ"
    
    Exit Sub
errH:
    strErr = Err.Description
End Sub

Public Sub ChangeSessionInfo(Optional ByVal lng����ID As Long, Optional ByVal str�������� As String, _
    Optional ByVal str��Ա���� As String, Optional ByVal str��ǰλ�� As String, _
    Optional ByRef strErr As String)
'���ݲ��������������ʹ������÷���2019.3��������ͬʱChangePosition������ɾ��

    '����ǰ�Ự��Ϣ�������� zlClientSession
    Dim str����վ As String, lng�˿ں� As Long
    Dim strSql As String
    
    On Error GoTo errH
    str����վ = gobjComLib.OS.ComputerName
    lng�˿ں� = mlngPort
    
    strSql = "Zltools.Zl_Zlclientsession_Save('" & str����վ & "', 0, " & lng�˿ں� & ", '', '', '', " _
           & lng����ID & ", '" & str��ǰλ�� & "', 1, '" & str�������� & "', '" & str��Ա���� & "')"
    gobjComLib.zlDatabase.ExecuteProcedure strSql, "����zlclientsession��Ϣ"
    
    Exit Sub
errH:
    strErr = Err.Description
End Sub

Public Function CheckDcnEnable(ByVal lngNoticeCode As Long) As Boolean
    '����: ���ݴ����NoticeCode���Dcn�Ƿ����
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 1 From zltools.zlNoticeLists Where NoticeCode = [1] And Status = 1"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "���DCN���ÿ�����", lngNoticeCode)
    
    CheckDcnEnable = rsTmp.RecordCount > 0
    Exit Function
errH:
    gobjComLib.ErrCenter
End Function

Public Function CheckDcnState() As Boolean
    '����:����ʵʱ���DCN��״̬
    CheckDcnState = mfrmWinsock.blnDcnState
End Function

Public Sub ChangeDepts(ByVal strDepts As String, Optional ByRef strErr As String)
    '����:�������վ���ڶ������,������������zlClientDepts����
    'strDepts = ������ŵ�ID, ��������֮���ö�����Ϊ���
    '���ʧ��,strErr����ʧ����Ϣ
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Zltools.Zl_Zlclientdepts_Save(" & glngSessionID & ",'" & strDepts & "')"
    gobjComLib.zlDatabase.ExecuteProcedure strSql, "����zlclientdepts��Ϣ"
    
    Exit Sub
errH:
    strErr = Err.Description
End Sub
