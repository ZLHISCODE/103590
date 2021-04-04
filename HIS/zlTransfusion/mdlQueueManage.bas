Attribute VB_Name = "mdlQueueManage"
Option Explicit

Private mobjQueueManage As Object    '���нӿڲ��� zlQueueManage.clsQueueManage
Private mobjLCDShow As Object        '��ʾ�ӿڲ���  zl9LCDShow.clsLCDShow

Public Sub QueueInit()
    '��ʼ�����ж���
    Dim strName(1) As String
    Dim strPrivs As String
    
    '�Ŷӽк�Ȩ��
    strPrivs = GetPrivFunc(glngSys, 1160)
    If Trim(strPrivs) = "" Then
        Exit Sub
    End If
    
    strName(1) = "��Һ��"
    On Error GoTo hErr
    Set mobjQueueManage = CreateObject("zlQueueManage.clsQueueManage")
    
    If mobjQueueManage Is Nothing Then
        Exit Sub
    Else
        Call mobjQueueManage.zlInitVar(gcnOracle, glngSys, 3, 0)
    End If
    
    If zlDatabase.GetPara("��ʾ�ŶӶ���", glngSys, 1160, "1") = "1" Then
        Set mobjLCDShow = CreateObject("zl9LCDShow.clsLCDShow")
        If Not mobjLCDShow Is Nothing Then
            Call mobjLCDShow.zlshow(gcnOracle, strName)     'LCDSHOW
        End If
    End If
    Exit Sub

hErr:
    Call QueueUnload
End Sub

Public Sub QueueUnload()
    If Not mobjQueueManage Is Nothing Then
        mobjQueueManage.CloseWindows
        Set mobjQueueManage = Nothing
    End If
    
    If Not mobjLCDShow Is Nothing Then
        mobjLCDShow.zlClose
        Set mobjLCDShow = Nothing
    End If
End Sub

Public Function QueueTimeCall() As Object
    On Error GoTo hErr
    If Not mobjQueueManage Is Nothing Then
        Set QueueTimeCall = mobjQueueManage.zlGetForm
    Else
        MsgBox "ʵ�����ŶӽкŲ���ʧ�ܣ�", vbInformation, gstrSysName
    End If
    Exit Function
    
hErr:
    Set QueueTimeCall = Nothing
    If ErrCenter = 1 Then Resume
End Function

Public Sub QueueOnePlay(ByVal strNO As String, ByVal strPlayInfo As String, ByVal lngNo As Long)
'���ܣ�����ָ������
    If mobjQueueManage Is Nothing Then Exit Sub
    
    Dim rsSQL As ADODB.Recordset
    Dim strSQL As String
    Dim strContent As String
    Dim strReserve As String
    Dim blnExcute As Boolean
    
    strSQL = "Select a.Id, a.No, a.ִ�в���id, a.����, a.����id, null ��� From ���˹Һż�¼ A Where a.No = [1]"
    On Error GoTo errHandle
    
    Set rsSQL = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Һŵ���Ϣ", strNO)
    If rsSQL.EOF = False Then
        '������������ŶӽкŶ���
        Call mobjQueueManage.zlDelQueue("��Һ��")
        '�����Ŷ�
        If mobjQueueManage.zlInQueue("��Һ��", 3, rsSQL!ID, zlCommFun.NVL(rsSQL!ִ�в���id, 0), zlCommFun.NVL(rsSQL!����), zlCommFun.NVL(rsSQL!����ID, 0), "", "") Then
            'ִ���Ŷ�
            Call mobjQueueManage.zlQueueExec("��Һ��", 3, rsSQL!ID, 1)
        End If
        'ˢ���Ŷӽк�LCD��ʾ
        Call mobjQueueManage.zlRefresh(Split("|��Һ��", "|"), "��Һ��", rsSQL!ID)
    End If
    
    If zlDatabase.GetPara("��ʾ�ŶӶ���", glngSys, 1160, "1") = "1" Then
        If Not mobjLCDShow Is Nothing Then
            Call mobjLCDShow.zlshow(gcnOracle, Split("|��Һ��", "|"))     'LCDSHOW
        End If
    End If
    
    Call PlugInFunc
    blnExcute = True
    If Not gobjPlugIn Is Nothing Then
        strContent = strPlayInfo
        On Error Resume Next
        blnExcute = gobjPlugIn.TransfusionCall(glngSys, glngModul, strNO, lngNo, strContent, strReserve)
        Call zlPlugInErrH(Err, "TransfusionShowPatiList")
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0
    End If
    If (blnExcute = False And strContent <> "") Or blnExcute = True Then
        If strContent <> "" Then strPlayInfo = strContent
        Call mobjQueueManage.zlQueueBroadcastCall(strPlayInfo)
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub QueueCall(ByVal strQueueName As String, ByVal lngDept As Long, _
                     ByVal objPati As cPatient)
'���ܣ�˳��
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lng�Һ�ID As Long
    
    '--- ˳��
    On Error GoTo hErr

    If mobjQueueManage Is Nothing Then Exit Sub     '���в�������
    If objPati Is Nothing Then Exit Sub
    
'    '�����ùҺ�ID��ͳһ��Long����
'    strSQL = "Select a.���, Decode(b.Id, Null, c.Id, b.Id) ID " & _
'             "From ���ﴩ��̨ A, ���˹Һż�¼ B, ���˹Һż�¼ C " & _
'             "Where a.�Һŵ�1 = b.No(+) And a.�Һŵ�2 = c.No(+) " & _
'             "    And b.��¼����(+) = 1 And b.��¼״̬(+) = 1 " & _
'             "    And c.��¼����(+) = 1 And c.��¼״̬(+) = 1 " & _
'             "    And ����ID = [1] And (a.�Һŵ�1 = [2] Or a.�Һŵ�2 = [2]) "
'
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�Ƿ�ɺ���", lngDept, strNo)
'    If Not rsTmp.EOF Then
'        lng�Һ�ID = zlCommFun.NVL(rsTmp!ID, 0)
'
'        '�����һ��
'        'Call mobjQueueManage.zlDelQueue(strQueueName, lng�Һ�ID)
'        '�������������Һ�ŶӽкŶ���
'        Call mobjQueueManage.zlDelQueue("��Һ��")
'        '�����Ŷ�
'        If mobjQueueManage.zlInQueue("��Һ��", 3, lng�Һ�ID, lngDept, strPatiName, lngPatiId, "", "") Then
'            'ִ���Ŷ�
'            Call mobjQueueManage.zlQueueExec("��Һ��", 3, lng�Һ�ID, 1)
'            SaveOperLog lngDept, strNo, CALLS, "��ʾ������"
'        End If
'        'ˢ���Ŷӽк�LCD��ʾ
'        Call mobjQueueManage.zlRefresh(Split("|��Һ��", "|"), "��Һ��", lng�Һ�ID)
'    End If
    
    '�������������Һ�ŶӽкŶ���
    Call mobjQueueManage.zlDelQueue("��Һ��")
    '�����Ŷ�
    If mobjQueueManage.zlInQueue("��Һ��", 3, objPati.����ID, lngDept, objPati.����, objPati.����ID, "", "") Then
        'ִ���Ŷ�
        Call mobjQueueManage.zlQueueExec("��Һ��", 3, objPati.����ID, 1)
        SaveOperLog lngDept, objPati, CALLS, "��ʾ������"
    End If
    'ˢ���Ŷӽк�LCD��ʾ
    Call mobjQueueManage.zlRefresh(Split("|��Һ��", "|"), "��Һ��", objPati.����ID)
    
    Exit Sub
hErr:
    SaveErrLog
End Sub

Public Sub QueueSetup(ByVal frmMe As Form)
    If Not mobjQueueManage Is Nothing Then
        Call mobjQueueManage.zlQueueParameterSetup(frmMe, glngSys)
    Else
        MsgBox "ȱ��10.30.40���ϰ汾�ĺ��нӿڲ�����zlQueueManage��,����", vbQuestion, "������Һ"
    End If
End Sub
