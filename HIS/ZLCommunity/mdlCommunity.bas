Attribute VB_Name = "mdlCommunity"
Option Explicit

Public gcnOracle As ADODB.Connection 'ȫ�����ݿ�����
Public gstrSysName As String '������Ϣ��ʾ��

Public grsCommunity As ADODB.Recordset '����Ŀ¼����
Public gcolCommunity As New Collection '������������
Public gobjCommunity As Object '��ǰʹ�õ���������

Public Function GetCommunity(ByVal int���� As Integer) As Object
'���ܣ���̬��ʼ��ָ�������������ͻ����������ض�Ӧ����������
'���أ������ʼ���ɹ��򷵻�������������
    Dim objTemp As Object
    
    'ȡ�������ѳ�ʼ���õ���������
    On Error Resume Next
    Set objTemp = gcolCommunity("_" & int����)
    Err.Clear: On Error GoTo 0
    
    '���û�б�ʾ��û�г�ʼ��
    If objTemp Is Nothing Then
        grsCommunity.Filter = "���=" & int����
        If grsCommunity.EOF Then Exit Function '��Ϊ�����������Ӧ�ò�������������
        If zlCommFun.Nvl(grsCommunity!����, 0) = 0 Then
            MsgBox grsCommunity!���� & "��ǰû�����á�", vbExclamation, gstrSysName
            Exit Function
        End If
        
        '��������������
        On Error Resume Next
        Set objTemp = CreateObject(grsCommunity!������ & ".clsCommunity")
        If Err.Number <> 0 Then
            MsgBox grsCommunity!���� & "����""" & grsCommunity!������ & ".dll""û����ȷ��װ��", vbExclamation, gstrSysName
            Err.Clear: Exit Function
        End If
        
        '��ʼ������������
        Err.Clear: On Error GoTo errH
        If Not objTemp.Initialize(gcnOracle) Then Exit Function
        
        '��ʼ���ɹ�֮����벿������
        gcolCommunity.Add objTemp, "_" & int����
    End If
    
    Set GetCommunity = objTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
