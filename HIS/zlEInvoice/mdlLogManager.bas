Attribute VB_Name = "mdlLogManager"
Option Explicit
'*********************************************************************************************************************************************
'����:��־����
'�ӿ�˵��:
'   1.zlWritLog:д��־
'����:���˺�
'����:2019*01*25 15:14:00
'*********************************************************************************************************************************************
Public gobjLogManager As Object
Public gblnCreateLogManager As Boolean
Public Sub zlWritLog(ByVal lngModule As Long, ByVal strFunName As String, ByVal strCallFunName As String, _
    ByVal strLogInfor As String, Optional ByVal intLogType As Integer = 0, Optional strLogName As String = "һ��ͨ�ӿڵ�����־", _
    Optional strGroupName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��־д��
    '���:lngModule-��ǰģ���
    '     strCallFunName-����������
    '     strFunName-��������
    '     intLogType-��־����:0-������־;1-����SQL;2-������Ϣ
    '     strLogInfor-д�����־����
    '     strLogName-��־����
    '     strGroupName-����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-01-15 15:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objLogManager As Object
    On Error GoTo errHandle
    If zlGetLogManagerObject(objLogManager) = False Then
        Call LogWrite(strLogName, lngModule, strFunName, "������:" & strCallFunName & IIf(strGroupName = "", "", "-" & strGroupName) & vbTab & strLogInfor)
        Exit Sub
    End If
    If objLogManager Is Nothing Then Exit Sub
    Call gobjLogManager.zlWritLog(lngModule, strFunName, strCallFunName, strLogInfor, intLogType, strLogName)
    Set objLogManager = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub zlWritLogEx(ByVal objCallMain As Object, ByVal lngModule As Long, ByVal strFunName As String, ByVal strLogClassify As String, _
    ByVal strLogInfor As String, Optional ByVal intLogType As Integer = 0, Optional strLogName As String = "����Ʊ�ݵ�����־", _
    Optional strGroupName As String, Optional strBusinessName As String = "����Ʊ��")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��־д��
    '���:objCallMain-�����߶��󣨿������࣬Ҳ�����Ǵ��壩
    '     lngModule-��ǰģ���
    '     strLogType-�������򷽷���
    '     strLogClassify-��־��𣬱��磺��ʼ��������
    '     intLogType-��־����:0-������־;1-����SQL;2-������Ϣ
    '     strLogInfor-д�����־����
    '     strLogName-��־����
    '     strGroupName-����
    '     strBusinessName-ҵ������
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-01-15 15:08:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objLogManager As Object
    Dim strCallFunName As String, strLogInforEx As String
    On Error GoTo errHandle
    If Not objCallMain Is Nothing Then
        strCallFunName = App.ProductName & "." & TypeName(objCallMain) & "." & strFunName
    ElseIf InStr(strFunName, ".") > 0 Then
        strCallFunName = App.ProductName & "." & strFunName
    Else
        strCallFunName = App.ProductName & ".�޷�ȷ��������." & strFunName
    End If
    
    '��־��Ϣ����:
    ' ������ +��( strLogClassify  ��+ strLogName
    strLogInforEx = strFunName & "(" & strLogClassify & ")" & strLogInfor
    If zlGetLogManagerObject(objLogManager) = False Then
        Call LogWrite(strLogName, lngModule, strBusinessName, "������:" & strCallFunName & IIf(strGroupName = "", "", "-" & strGroupName) & vbTab & strLogInforEx)
        Exit Sub
    End If
    If objLogManager Is Nothing Then Exit Sub
    Call gobjLogManager.zlWritLog(lngModule, strFunName, strCallFunName, strLogInforEx, intLogType, strLogName)
    Set objLogManager = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function zlGetLogManagerObject(ByRef objLogManager As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ù�����������
    '����:objLogManager-������־����������
    '����:��ȡ����true,���򷵻�False
    '����:���˺�
    '����:2019-01-25 09:57:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If Not gobjLogManager Is Nothing Then
        Set objLogManager = gobjLogManager: zlGetLogManagerObject = True
        Exit Function
    End If
    If gblnCreateLogManager Or gcnOracle Is Nothing Then Exit Function  'ֻ��ʼ��һ��,����ʱ�����ٳ�ʼ��
    
    
    Err = 0: On Error Resume Next
    If gobjLogManager Is Nothing Then
        Set gobjLogManager = CreateObject("zlLogManager.clsLogManager")
        gblnCreateLogManager = True
        If Err <> 0 Then Exit Function
    End If
    
    Err.Clear:  On Error GoTo errHandle
    If gobjLogManager Is Nothing Then Exit Function
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If gobjLogManager.InitCommon(gcnOracle, UserInfo.�û���) = False Then Exit Function
    Set objLogManager = gobjLogManager: zlGetLogManagerObject = True
    Exit Function
errHandle:
    Exit Function
End Function
 

