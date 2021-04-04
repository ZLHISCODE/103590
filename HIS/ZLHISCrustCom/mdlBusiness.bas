Attribute VB_Name = "mdlBusiness"
Option Explicit
'���øó���Ҫʵ�ֵĲ���
Public Enum OperateType
    OT_Repair = 0                                       '�����޸����൱������,���ж��Ƿ�Ԥ�������
    OT_PreUpgrade = 1                                   '��ǰ�������������ļ�������ʱĿ¼
    OT_OfficialUpgrade = 2                              '����ǰ����Ŀ¼�л��߷�����Ŀ¼����ȡ�ļ�����װ·��
    OT_CheckFile = 3                                        '��ʱֻ���ļ��ռ����ռ�APPSOFTĿ¼�µ�ָ�������ļ����������������������ۣ��������͵���Ϊ���ͻ��˲����Ƿ���Ҫ����
End Enum

Public Enum OperateStep
    OS_NotInProcessing = 0                              'δִ��
    OS_Completed = 1                                    'ִ�����,����OT_CheckFile,Ϊ�����ϣ���������
    OS_Failure = 2                                      'ִ��ʧ��,����OT_CheckFile,Ϊ�����ϣ�������
    OS_InProcessing = 3                                 'ִ����
End Enum

'��������
Public Enum MsgType
    MT_MsgHeader = 0                                    '��Ϣͷ
    MT_InitEnv = 1                                      '�ô�������δ��ʶ
    MT_SvrConn = 2                                      '���ӷ���������
    MT_ChcekUpdate = 3                                  '���¼��
    MT_DownAndDec = 4                                   '���ؽ�ѹ��������
    MT_SetUp = 5                                        '���������ڰ�װĿ¼����
    MT_RegCom = 6                                       '����ע�����
    MT_ExeBat = 7                                       'ִ�����������
    MT_MsgFoot = 8                                      '��Ϣβ��
End Enum

'�ļ�����
Public Enum FileType
    FT_Public = 0                   '��Ʒ��������
    FT_Apply = 1                    '��ƷӦ�ò���
    FT_Help = 2                     '��Ʒ�����ļ�
    FT_AdditionFile = 3             '��Ʒ�����ļ�
    FT_Other = 4                    '������Ʒ�ļ�
    FT_System = 5                   'ϵͳ�ļ�
End Enum
Public Function SetOperateProcess(ByVal otCurType As OperateType, ByVal osCurStep As OperateStep, Optional ByVal strMsg As String, Optional ByVal lngBeach As Long) As Boolean
'���ܣ����²������ȡ�
'������otCurType=��ǰ��������
'      osCurStep=��ǰ����
'      lngBeach=����������
'      strMsg=������Ϣ
'���أ��Ƿ�ִ�гɹ�
    Dim blnComplete As Boolean, strSQL As String
    Dim strBeach As String
    Dim objSend         As New clsMemoryShareFP
    Const SHARE_CLIENT_SEND           As String = "3892908F-5A80-484C-A031-FA95647E8EBE"              '����̨������Ϣ�������ڴ湲��
    gobjTrace.WriteSection "�����������", SL_LevelThree
    strMsg = MidB(strMsg, 1, glngNoteLength - 30)
    On Error Resume Next
    strSQL = "zlTOOLS.Zl_Zlclients_UpdateProcess('" & gstrComputerName & "'," & otCurType & "," & osCurStep & "," & SQLAdjust(strMsg) & "," & IIf(lngBeach <> 0 And osCurStep = OS_Completed, lngBeach, "Null") & ")"
    Call ExecuteProcedure(strSQL, "SetOperateProcess")
    If Err.Number <> 0 Then
        gobjTrace.WriteInfo "SetOperateProcess", "��ǽ��", "����", "���SQL", Replace(Replace(strSQL, Chr(10), ""), Chr(13), ""), "������Ϣ", Err.Description
        Err.Clear
        blnComplete = osCurStep = OS_Completed Or osCurStep = OS_Failure And otCurType = OT_CheckFile
        Select Case otCurType
            Case OT_OfficialUpgrade '��ʽ������������Ԥ���������Ϣ�������޸������Ϣ����ȡ��Ԥ������־��������־
                strSQL = "Update zlTOOLS.zlClients Set �������=" & osCurStep & " ,����˵��=" & SQLAdjust(strMsg) & "" & IIf(lngBeach <> 0 And osCurStep = OS_Completed, ",����=" & lngBeach, "") & IIf(blnComplete, ",������־=0,�Ƿ�Ԥ����=0,�޸�״̬=0,Ԥ�����=0,�Ƿ���������=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
            Case OT_PreUpgrade
                strSQL = "Update zlTOOLS.zlClients Set Ԥ�����=" & osCurStep & " ,Ԥ����˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",�Ƿ�Ԥ����=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
            Case OT_Repair '�����޸���������Ԥ���������Ϣ�������޸������Ϣ����ȡ��Ԥ������־��������־
                strSQL = "Update zlTOOLS.zlClients Set �޸�״̬=" & osCurStep & " ,�޸�˵��=" & SQLAdjust(strMsg) & "" & IIf(lngBeach <> 0 And osCurStep = OS_Completed, ",����=" & lngBeach, "") & IIf(blnComplete, ",������־=0,�Ƿ�Ԥ����=0,�������=0,Ԥ�����=0,�Ƿ���������=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
            Case OT_CheckFile
                strSQL = "Update zlTOOLS.zlClients Set �ռ�״̬=" & osCurStep & " ,�ռ�˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",�ռ���־=0", "") & " Where ����վ = '" & gstrComputerName & "'"
                
        End Select
        gcnOracle.Execute strSQL, , adCmdText
        If Err.Number <> 0 Then 'ִ��SQL����˵���ṹ��û������������ִ���Ͻṹ����
            gobjTrace.WriteInfo "SetOperateProcess", "��ǽ��", "����", "���SQL", Replace(Replace(strSQL, Chr(10), ""), Chr(13), ""), "������Ϣ", Err.Description
            Err.Clear
            Select Case otCurType
                Case OT_OfficialUpgrade '��ʽ������������Ԥ���������Ϣ�������޸������Ϣ����ȡ��Ԥ������־��������־
                    strSQL = "Update zlTOOLS.zlClients Set �������=" & osCurStep & " ,˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",������־=0,Ԥ�����=0,�Ƿ���������=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
                Case OT_PreUpgrade
                    strSQL = "Update zlTOOLS.zlClients Set Ԥ�����=" & osCurStep & " ,˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",������־=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
                Case OT_Repair '�����޸���������Ԥ���������Ϣ�������޸������Ϣ����ȡ��Ԥ������־��������־
                    strSQL = "Update zlTOOLS.zlClients Set �������=" & osCurStep & " ,˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",������־=0,Ԥ�����=0,�Ƿ���������=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
                Case OT_CheckFile
                    strSQL = "Update zlTOOLS.zlClients Set �ռ�״̬=" & osCurStep & ",˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",�ռ���־=0", "") & " Where ����վ = '" & gstrComputerName & "'"
            End Select
            gcnOracle.Execute strSQL, , adCmdText
            If Err.Number <> 0 Then
                gobjTrace.WriteInfo "SetOperateProcess", "��ǽ��", "ʧ��", "���SQL", Replace(Replace(strSQL, Chr(10), ""), Chr(13), ""), "������Ϣ", Err.Description
                Call RecordErrMsg(MT_InitEnv, "�������ִ�����", "��ȷ�Ϲ����߶�����Ȩ��������" & Err.Description)
                Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:����ʧ�� ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
                Err.Clear
                If Not gblnHelperMain Then MsgBox "�޷��������ִ�����������ϵ����Աȷ�Ϲ����߶���Ȩ��������", vbInformation, App.Title
                Exit Function
            ElseIf osCurStep = OS_InProcessing And otCurType = OT_CheckFile Then
                strSQL = "Delete Zlclientupdatelog A Where a.����վ ='" & gstrComputerName & "' And ���� = 1"
                gcnOracle.Execute strSQL, , adCmdText
                If Err.Number <> 0 Then Err.Clear
            End If
        ElseIf osCurStep = OS_InProcessing And otCurType = OT_CheckFile Then
            strSQL = "Delete Zlclientupdatelog A Where a.����վ ='" & gstrComputerName & "' And ���� = 1"
            gcnOracle.Execute strSQL, , adCmdText
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    gobjTrace.WriteInfo "SetOperateProcess", "��ǽ��", "�ɹ�", "���SQL", Replace(Replace(strSQL, Chr(10), ""), Chr(13), "")
    If (osCurStep = OS_Failure Or osCurStep = OS_Completed) And gblnHelperMain Then
        If objSend.OpenMemoryShare(SHARE_CLIENT_SEND) Then
            '0-�쳣,1-����|ϵͳ|ģ��|������Ϣ
            If objSend.WriteMemory(IIf(osCurStep = OS_Completed, 1, 0) & "|0|0|", GetCurrentProcessId, Decode(gotCurType, OT_Repair, 1, OT_OfficialUpgrade, 3, OT_PreUpgrade, 2, OT_CheckFile, 4), 2) Then
            End If
        End If
    End If
    SetOperateProcess = True
End Function

Public Function CheckJobs() As Boolean
'����:��鲢��ȡ�������������
    Dim rsTmp       As ADODB.Recordset, strSQL  As String
    Dim datCur      As Date, blnOnlyOfficialUp  As Boolean, blnOnlyPreUp    As Boolean
    Dim blnPreUp    As Boolean, blnOfficialUp   As Boolean, blnPreComplete  As Boolean, blnCollect  As Boolean
    Dim strMsg      As String
    
    On Error GoTo ErrH
    '���´���һ�㲻���ܳ���
    datCur = Currentdate
    '�ж������Ƿ������ȡ�Ƿ������˶�ʱ����
    strSQL = "Select Max(����) ���� From ZLTOOLS.zlRegInfo Where ��Ŀ='�ͻ�����������'"
    Set rsTmp = OpenSQLRecord(strSQL, "��鶨ʱ����")
    If rsTmp!���� & "" <> "" Then
        If CDate(Format(datCur, "YYYY-MM-DD hh:mm:ss")) >= CDate(Format(NVL(rsTmp!����), "YYYY-MM-DD hh:mm:ss")) Then
            blnOnlyOfficialUp = True 'ֻ����ʽ����
        Else
            blnOnlyPreUp = True 'ֻ��Ԥ����
        End If
    Else
        blnOnlyOfficialUp = True
    End If
    gobjTrace.WriteInfo "CheckJobs", "�Ƿ�ֻ����ʽ����", blnOnlyOfficialUp, "�Ƿ�ֻ��Ԥ����", blnOnlyPreUp
    On Error Resume Next
    Set rsTmp = Nothing
    '����û���Ƿ�Ԥ�����ֶ�(��ΪԤ����ʱ�����ݿ⻹û�������������Ҫ�������
    strSQL = "Select Nvl(�Ƿ�Ԥ����,0) �Ƿ�Ԥ����, Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־ From ZLTOOLS.Zlclients Where ����վ = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��鵱ǰ����", gstrComputerName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrH
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!�Ƿ�Ԥ���� = 1
            blnOfficialUp = rsTmp!������־ = 1
            blnPreComplete = rsTmp!Ԥ����� = 1
            blnCollect = rsTmp!�ռ���־ = 1
        End If
    Else
        '�����·�ʽ��ȡ��ʧ����ʹ���Ϸ�ʽ�����Ӽ�����
        strSQL = "Select Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־ From ZLTOOLS.Zlclients Where ����վ = [1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��鵱ǰ����", gstrComputerName)
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!������־ = 1
            blnOfficialUp = rsTmp!������־ = 1
            blnPreComplete = rsTmp!Ԥ����� = 1
            blnCollect = rsTmp!�ռ���־ = 1
        End If
    End If
    gobjTrace.WriteInfo "CheckJobs", "�Ƿ���ҪԤ����", blnPreUp, "�Ƿ���Ҫ��ʽ����", blnOnlyPreUp, "�ϴ�Ԥ�����Ƿ����", blnPreComplete, "�Ƿ�����ļ��ռ�", blnCollect
    If gotCurType = OT_Repair Then
        If blnOnlyPreUp Then
            gotCurType = OT_PreUpgrade
        End If
    ElseIf (blnOfficialUp Or blnPreUp) And blnOnlyPreUp Then
        gotCurType = OT_PreUpgrade
    ElseIf (blnOfficialUp Or blnPreUp) And blnOnlyOfficialUp Then
        gotCurType = OT_OfficialUpgrade
    ElseIf blnCollect Then
        gotCurType = OT_CheckFile
    Else
        gobjTrace.WriteInfo "CheckJobs", "�����", "��ǰû���κ�����ϵͳ���Զ��˳�"
        Call RecordErrMsg(MT_InitEnv, "������", "��ǰû���κ�����ϵͳ���Զ��˳�")
        CheckJobs = True
        Exit Function
    End If
    'Ԥ�����Ѿ����
    If blnPreComplete And gotCurType = OT_PreUpgrade Then
        gobjTrace.WriteInfo "CheckJobs", "�����", "��ǰֻ��Ԥ����������Ԥ�����Ѿ���ɣ�ϵͳ���Զ��˳���"
        Call RecordErrMsg(MT_InitEnv, "������", "��ǰֻ��Ԥ����������Ԥ�����Ѿ���ɣ�ϵͳ���Զ��˳���")
        CheckJobs = True
        Exit Function
    End If
    gblnSilence = gotCurType = OT_CheckFile Or gotCurType = OT_PreUpgrade
    gobjTrace.WriteInfo "CheckJobs", "�����", Decode(gotCurType, OT_OfficialUpgrade, "��ʽ����", OT_PreUpgrade, "Ԥ����", OT_Repair, "�޸���ǿ������", OT_CheckFile, "�ռ�������")
    If gotCurType <> OT_CheckFile Then
        Set gclsConnect = GetFileConnect(strMsg)
        If gclsConnect Is Nothing Then
            gobjTrace.WriteInfo "CheckJobs", "����ʧ��", strMsg
            Call RecordErrMsg(MT_InitEnv, "������", "�޷������ļ�������,���ܼ������в�������Ϣ��" & strMsg)
            If Not gblnHelperMain Then MsgBox "�޷������ļ�������������ϵ����Ա����Ϣ��" & vbNewLine & strMsg, vbInformation, App.Title
            Exit Function
        End If
    Else
        Set gclsConnect = New clsConnect
    End If
    CheckJobs = True
    Exit Function
ErrH:
    strMsg = Err.Description
    gobjTrace.WriteInfo "CheckJobs", "�����ⷢ����������", strMsg
    If gblnHelperMain Then MsgBox "�����ⷢ��������������ϵ����Ա����Ϣ��" & vbNewLine & strMsg, vbInformation, App.Title
    Err.Clear
End Function

Private Function GetFileConnect(ByRef strMsg As String) As clsConnect
'���ܣ���ȡ�������ļ�����
    Dim objConn As New clsConnect
    Dim sctConnType As ServerConnectType
    Dim strServerID As String, strServer As String, strUser As String, strPwd As String, strPort As String, strCollectType As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnDefalut As Boolean, blnConnOK As Boolean
    Dim blnOldStype As Boolean, blnActiveModeFTP As Boolean
    
    On Error Resume Next
    If gotCurType = OT_CheckFile Then
        strSQL = "Select ����, λ��, �û���, ����, �˿�, �ռ�����, FTP����ģʽ From Zltools.Zlupgradeserver Where Nvl(�Ƿ��ռ�, 0) = 1"
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�������������", gstrComputerName)
        If Err.Number = 0 Then
            If Not rsTmp.EOF Then
                strServerID = rsTmp!��� & ""
                sctConnType = IIf(rsTmp!���� = 0, SCT_Share, SCT_FTP)
                strServer = rsTmp!λ��
                strUser = rsTmp!�û���
                strPwd = DeCipher(rsTmp!���� & "")
                strPort = rsTmp!�˿� & ""
                blnActiveModeFTP = Val(rsTmp!FTP����ģʽ & "") = 1
                strCollectType = rsTmp!�ռ����� & ""
            End If
        Else
            Err.Clear
            blnOldStype = True
        End If
    Else
        strSQL = "Select �����ļ������� From ZLTools.zlClients Where ����վ=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�������������", gstrComputerName)
        If Err.Number = 0 Then
            If Not rsTmp.EOF Then strServerID = rsTmp!�����ļ������� & ""
        Else
            Err.Clear
            blnOldStype = True
        End If
        If strServerID <> "" Then
            strSQL = "Select ���,����, λ��, �û���, ����, �˿�,Nvl(�Ƿ�ȱʡ,0) �Ƿ�ȱʡ, ����, FTP����ģʽ From Zltools.Zlupgradeserver Where ��� = [1]"
            Set rsTmp = OpenSQLRecord(strSQL, "��ȡ����������", Val(strServerID))
            If Not rsTmp.EOF Then
                strServerID = rsTmp!��� & ""
                sctConnType = IIf(rsTmp!���� = 0, SCT_Share, SCT_FTP)
                strServer = rsTmp!λ��
                strUser = rsTmp!�û���
                strPwd = DeCipher(rsTmp!���� & "")
                strPort = rsTmp!�˿� & ""
                blnActiveModeFTP = Val(rsTmp!FTP����ģʽ & "") = 1
                glngFileBatch = Val(rsTmp!���� & "")
                blnDefalut = rsTmp!�Ƿ�ȱʡ = 1
            Else
                strServerID = ""
            End If
        End If
    End If
    If blnOldStype Then
        Set GetFileConnect = GetFileConnectOld(strMsg)
    Else
        If strServerID <> "" Then
            gobjTrace.WriteInfo "�ļ�������", "�������ļ�����", glngFileBatch, "���������", strServerID, "�Ƿ�Ĭ��", blnDefalut
            blnConnOK = objConn.ToConnect(sctConnType, strServer, strUser, strPwd, strPort, strCollectType, blnActiveModeFTP, strMsg)
        End If
        '���Ӳ��ɹ��������������Զ�����Ĭ�Ϸ�����
        If Not blnConnOK And gotCurType <> OT_CheckFile And Not blnDefalut Then
            strSQL = "Select ���,����, λ��, �û���, ����, �˿�, ����, FTP����ģʽ From Zltools.Zlupgradeserver Where Nvl(�Ƿ�ȱʡ,0) = 1"
            Set rsTmp = OpenSQLRecord(strSQL, "��ȡĬ������������")
            If Err.Number = 0 Then
                If Not rsTmp.EOF Then
                    strServerID = rsTmp!��� & ""
                    sctConnType = IIf(rsTmp!���� = 0, SCT_Share, SCT_FTP)
                    strServer = rsTmp!λ��
                    strUser = rsTmp!�û���
                    strPwd = DeCipher(rsTmp!���� & "")
                    strPort = rsTmp!�˿� & ""
                    blnActiveModeFTP = Val(rsTmp!FTP����ģʽ & "") = 1
                    glngFileBatch = Val(rsTmp!���� & "")
                    gobjTrace.WriteInfo "Ĭ�Ϸ�����", "�������ļ�����", glngFileBatch, "���������", strServerID
                    blnConnOK = objConn.ToConnect(sctConnType, strServer, strUser, strPwd, strPort, , blnActiveModeFTP, strMsg)
                End If
            Else
                Err.Clear
            End If
        End If
        If blnConnOK Then Set GetFileConnect = objConn
    End If
    Exit Function
ErrH:
    strMsg = Err.Description
End Function

Private Function GetFileConnectOld(ByRef strMsg As String) As clsConnect
'���ܣ���ȡ�ļ����������ӣ��Ϸ�ʽ
'������blnUpgrade=True-Ԥ���������������� ��false-�ļ��ռ�������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim sctConnType As ServerConnectType, strServerID As String
    Dim objConn As New clsConnect
    Dim arrParas() As Variant, arrValues(4) As String
    Dim strSQLPars As String, i As Integer
    Dim blnReadOk As Boolean, blnConnOK As Boolean, blnGo As Boolean
    
    On Error GoTo ErrH
    '��ȡ��������
    sctConnType = SCT_Share
    strSQL = "Select ��Ŀ,���� From ZLTools.zlregInfo where ��Ŀ=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��������", IIf(gotCurType <> OT_CheckFile, "��������", "�ռ���ʽ"))
    If Not rsTmp.EOF Then
        If NVL(rsTmp!����, 0) = 1 Then sctConnType = SCT_FTP
    End If
    If gotCurType = OT_CheckFile Then
        '�ļ��ռ�������ID
        strServerID = IIf(sctConnType = SCT_FTP, "F", "S")
    Else
        '��ȡ������ID
        strSQL = "Select ����������,FTP������ From ZLTools.zlClients Where ����վ=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�������������", gstrComputerName)
        If Not rsTmp.EOF Then strServerID = IIf(sctConnType = SCT_FTP, rsTmp!FTP������ & "", rsTmp!���������� & "")
    End If
    '��ȡ��������Ϣ
    If gotCurType <> OT_CheckFile Then
        If sctConnType = SCT_FTP Then
            arrParas = Array("FTP������", "FTP�û�", "FTP����", "FTP�˿�", "FTP����ģʽ")
        Else
            arrParas = Array("������Ŀ¼", "�����û�", "��������", "", "")
        End If
    Else
        arrParas = Array("�ռ�Ŀ¼", "�����û�", "��������", "���ʶ˿�", "�ռ�����")
    End If
ReGetParas:
    '�Ȼ�ȡSQL����
    strSQLPars = ""
    For i = LBound(arrParas) To UBound(arrParas)
        If arrParas(i) <> "" Then
            strSQLPars = strSQLPars & ",'" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
        End If
    Next
    strSQLPars = Mid(strSQLPars, 2)
    strSQL = "Select ��Ŀ,���� From ZLTools.zlregInfo where ��Ŀ in(" & strSQLPars & ")"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ������")
    If Not rsTmp.EOF Then
        For i = LBound(arrParas) To UBound(arrParas)
            If arrParas(i) <> "" Then
                rsTmp.Filter = "��Ŀ='" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
                If Not rsTmp.EOF Then arrValues(i) = rsTmp!���� & ""
            End If
        Next
    End If
    
    blnReadOk = True
    '���������û�������Ϊ�գ����ܽ����ռ�������
    If arrValues(0) = "" Or arrValues(1) = "" Or arrValues(2) = "" Then
        blnReadOk = False
    'FTP��ʽ��Ҫһ���˿�
    ElseIf sctConnType = SCT_FTP And arrValues(3) = "" Then
        blnReadOk = False
    '�ռ�ʱ���ռ����Ͳ���Ϊ��
    ElseIf gotCurType = OT_CheckFile And arrValues(4) = "" Then
        blnReadOk = False
    End If
    If blnReadOk Then
        gobjTrace.WriteInfo "GetFileConnectOld", "�ɷ�ʽ���������", strServerID
        If sctConnType = SCT_FTP Then
            blnConnOK = objConn.ToConnect(sctConnType, arrValues(0), arrValues(1), arrValues(2), arrValues(3) _
                , , arrValues(4), strMsg)
        Else
            blnConnOK = objConn.ToConnect(sctConnType, arrValues(0), arrValues(1), arrValues(2), arrValues(3) _
                , arrValues(4), False, strMsg)
        End If
    End If
    If (Not blnConnOK Or Not blnReadOk) And gotCurType <> OT_CheckFile Then
        If strServerID <> "" And strServerID <> "0" Then
            strServerID = "0"
            GoTo ReGetParas '���»�ȡ���ӷ������Ĳ���
        ElseIf (strServerID = "0" Or strServerID = "") And Not blnGo Then
            blnGo = True '��ֹѭ��
            strServerID = IIf(strServerID = "0", "", "0")
            GoTo ReGetParas '���»�ȡ���ӷ������Ĳ���
        End If
    End If
    If blnConnOK Then Set GetFileConnectOld = objConn
    Exit Function
ErrH:
    strMsg = Err.Description
End Function

Public Function CheckAndAdjustFolder() As Boolean
'���ܣ����а�װ·�����޸�
    Dim strSQL              As String, rsTmp        As ADODB.Recordset
    Dim strPath             As String, arrTmp       As Variant
    Dim i                   As Integer
    Dim strErrInfo          As String
    
    Err.Clear: On Error GoTo ErrH
    strSQL = "Select Distinct Upper(��װ·��) ��װ·�� From Zlfilesupgrade"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ·���ļ���")
    
    Do While Not rsTmp.EOF
        arrTmp = Split(rsTmp!��װ·�� & "", "\")
        strPath = ""
        If UBound(arrTmp) <> -1 Then
            arrTmp(0) = Trim(arrTmp(0))
            If arrTmp(0) = "[APPSOFT]" Then
                strPath = gstrSetupPath
            ElseIf arrTmp(0) = "[PUBLIC]" Then
                If Not gobjFSO.FolderExists(gstrSetupPath & "\PUBLIC") Then
                    gobjFSO.CreateFolder (gstrSetupPath & "\PUBLIC")
                End If
                strPath = gstrSetupPath & "\PUBLIC"
            ElseIf arrTmp(0) = "[APPLY]" Then
                strPath = gstrSetupPath & "\APPLY"
            ElseIf arrTmp(0) = "[OS:]" Then 'ϵͳ��
                strPath = Left(gstrSystemPath, 2)
            ElseIf arrTmp(0) = "[APP:]" Then  '��ǰ��װ��
                strPath = Left(gstrSetupPath, 2)
            End If
            If strPath <> "" Then
                For i = 1 To UBound(arrTmp)
                    If arrTmp(i) <> "" Then
                        strPath = strPath & "\" & arrTmp(i)
                        If Not gobjFSO.FolderExists(strPath) Then
                            gobjFSO.CreateFolder (strPath)
                        End If
                    End If
                Next
                '���氲װ·�����Ż�ת���ٶȡ�
                gcllSetPath.Add strPath, "K_" & rsTmp!��װ·��
            End If
        End If
        rsTmp.MoveNext
    Loop
    '���������װ·�����Ż�ת���ٶȡ�
    On Error Resume Next
    gcllSetPath.Add gstrSetupPath, "K_[APPSOFT]"
    gcllSetPath.Add gstrSetupPath & "\PUBLIC", "K_[PUBLIC]"
    gcllSetPath.Add gstrSetupPath & "\APPLY", "K_[APPLY]"
    gcllSetPath.Add Left(gstrSystemPath, 2), "K_[OS:]"
    gcllSetPath.Add Left(gstrSetupPath, 2), "K_[APP:]"
    gcllSetPath.Add gstrSystemPath, "K_[SYSTEM]"
    gcllSetPath.Add gobjFSO.GetParentFolderName(gstrSystemPath) & "\Help", "K_[HELP]"
    gcllSetPath.Add gstrSetupPath & "\APPLY", "K_[APPSOFT]\APPLY"
    If Err.Number Then Err.Clear
    On Error Resume Next
    '���������ļ�·��
    strSQL = "Select distinct upper(��װ·��) ��װ·�� From zlFilesExpired"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ·���ļ���")
    If Not rsTmp Is Nothing Then
        Err.Clear
        Do While Not rsTmp.EOF
            strPath = gcllSetPath("K_" & rsTmp!��װ·��)
            If Err.Number <> 0 Then
                Err.Clear
                arrTmp = Split(rsTmp!��װ·�� & "", "\")
                strPath = ""
                If UBound(arrTmp) <> -1 Then
                    arrTmp(0) = Trim(arrTmp(0))
                    If arrTmp(0) = "[APPSOFT]" Then
                        strPath = gstrSetupPath
                    ElseIf arrTmp(0) = "[PUBLIC]" Then
                        If Not gobjFSO.FolderExists(gstrSetupPath & "\PUBLIC") Then
                            gobjFSO.CreateFolder (gstrSetupPath & "\PUBLIC")
                        End If
                        strPath = gstrSetupPath & "\PUBLIC"
                    ElseIf arrTmp(0) = "[APPLY]" Then
                        strPath = gstrSetupPath & "\APPLY"
                    ElseIf arrTmp(0) = "[OS:]" Then 'ϵͳ��
                        strPath = Left(gstrSystemPath, 2)
                    ElseIf arrTmp(0) = "[APP:]" Then '��ǰ��װ��
                        strPath = Left(gstrSetupPath, 2)
                    End If
                    If strPath <> "" Then
                        For i = 1 To UBound(arrTmp)
                            If arrTmp(i) <> "" Then
                                strPath = strPath & "\" & arrTmp(i)
                                If Not gobjFSO.FolderExists(strPath) Then
                                    gobjFSO.CreateFolder (strPath)
                                End If
                            End If
                        Next
                        '���氲װ·�����Ż�ת���ٶȡ�
                        gcllSetPath.Add strPath, "K_" & rsTmp!��װ·��
                    End If
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    If Err.Number Then Err.Clear
    CheckAndAdjustFolder = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "CheckAndAdjustFolder", "����޸���װĿ¼ʧ��", strErrInfo
    Call RecordErrMsg(MT_InitEnv, "�޸���װĿ¼", strErrInfo)
    If Not gblnHelperMain Then MsgBox "����޸���װĿ¼����������������ϵ����Ա����Ϣ��" & vbNewLine & strErrInfo, vbInformation, App.Title
End Function

Public Function UpgradeBase(Optional ByVal blnUpgrade As Boolean = True) As Boolean
'���ܣ������Զ���������Ҫ�Ļ�������
    Dim strFile As String, blnAdmin As Boolean
    Dim strErr As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strReturn As String
    Dim strMsg As String
    Dim strCommand As String, strTmp As String
    Dim objText As TextStream, blnMust  As Boolean, blnErr  As Boolean
    
    If blnUpgrade Then
        gobjTrace.WriteSection "������������", SL_LevelTwo
        On Error GoTo ErrH
        strSQL = "Select ���, �ļ���, Upper(�ļ���) ��׼�ļ���," & IIf(gblnHaveVersion, "�ļ��汾��", " ") & " �汾��, �޸�����, �ļ�����, ҵ�񲿼�, Upper(��װ·��) ��װ·��, Md5, �Զ�ע��, ǿ�Ƹ���" & vbNewLine & _
                "From ZLTOOLS.Zlfilesupgrade" & vbNewLine & _
                "Where Upper(�ļ���) In ('ZLRUNAS.EXE','ZLHISCRUST.EXE','ZLHISCRUSTCOM.DLL','7Z.EXE','7Z.DLL','AAMD532.DLL','GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
        Set rsTmp = OpenSQLRecord(strSQL, App.Title)
        '1����������ZLRUNAS.EXE��ȡ����ԱȨ�ޣ��ɴ˿�������MD5���㲿��������ZlHISCrust������MD5
        On Error Resume Next
        strFile = gstrSetupPath & "\zlTestAdmin.txt"
        Call gobjFSO.CreateTextFile(strFile, True)
        Call gobjFSO.CopyFile(strFile, gstrSystemPath & "\zlTestAdmin.txt", True)
        If Err.Number = 75 Then
            blnAdmin = False
        ElseIf Dir(gstrSystemPath & "\zlTestAdmin.txt", vbNormal) <> "" Then
            blnAdmin = True
            Call gobjFSO.DeleteFile(gstrSystemPath & "\zlTestAdmin.txt", True)
        Else
            blnAdmin = False
        End If
        Call gobjFSO.DeleteFile(strFile, True)
        If Err.Number <> 0 Then Err.Clear
        gobjTrace.WriteInfo "UpgradeBase", "SystemĿ¼д��Ȩ��", blnAdmin
        If Not blnAdmin Then
            rsTmp.Filter = "��׼�ļ���='ZLRUNAS.EXE'"
            If Not rsTmp.EOF Then
                strFile = GetActualPath(rsTmp!��װ·��, Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
                If Not gobjFSO.FileExists(strFile) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                        If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                            strMsg = "�������ļ��ļ�����ʧ��(ZLRUNAS.EXE(USERȨ��ִ�й���))" & strErr
                        Else
                            gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                        End If
                    Else
                        strMsg = "�������ļ�ȱʧZLRUNAS.EXE(USERȨ��ִ�й���)"
                    End If
                End If
                If gobjFSO.FileExists(strFile) Then
                    '�ȱ��������У����´�����ʹ��
                    If gobjFSO.FileExists(gstrSetupPath & "\ZLRUNAS.ini") Then
                        gobjFSO.DeleteFile gstrSetupPath & "\ZLRUNAS.ini", True
                    End If
                    Set objText = gobjFSO.CreateTextFile(gstrSetupPath & "\ZLRUNAS.ini")
                    objText.WriteLine Cipher(gstrCommand)
                    objText.Close
                    Set objText = Nothing
                    strMsg = StartZLRunAs(strFile)
                End If
            Else
                strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧZLRUNAS.EXE(USERȨ��ִ�й���)"
            End If
            If strMsg <> "" Then
                gobjTrace.WriteInfo "UpgradeBase", "����Ա���й��߼��", strMsg
                Call RecordErrMsg(MT_InitEnv, "����Ա���й��߼��", strMsg)
                If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
                Exit Function
            End If
        End If
        '2������AAMD532.dll�ò�������������MD5,��������ZLHISCrust.exe�������޷����ZLHISCrust.exe�Ƿ���Ҫ������
        strMsg = ""
        rsTmp.Filter = "��׼�ļ���='AAMD532.DLL'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��AAMD532.DLL(MD5���㹤��)" & strErr
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧAAMD532.DLL(MD5���㹤��)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧAAMD532.DLL(MD5���㹤��)"
        End If
        
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "MD5���㹤�߼��", strMsg
            Call RecordErrMsg(MT_InitEnv, "MD5���㹤�߼��", strMsg)
            If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
        strMsg = ""
        '3������ZLHISCrust.exe���ò������Խ��м��������
        If Val(GetSetting("ZLSOFT", "����ģ��\�Զ�����", "���ߵ���", "0")) = 0 Then
            If gintCallTimes = 0 Then '�ڶ��ε����������߽�������������ZLRUNAS���õ���һ��
                rsTmp.Filter = "��׼�ļ���='ZLHISCRUST.EXE'"
                If Not rsTmp.EOF Then
                    strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
                    If IsFileUpgade(gstrAppPath & "\ZLHISCRUST.EXE", rsTmp!�汾�� & "", rsTmp!�޸����� & "", rsTmp!MD5 & "") Then
                        If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                            gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                            If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gstrTempPath, strErr) Then
                                strMsg = "�������ļ��ļ�����ʧ��:ZLHISCRUST.EXE(�Զ�����������)" & strErr
                            Else
                                gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", gstrTempPath & "\" & rsTmp!�ļ���
                                '�ļ��ֱ���ϲ��������ļ��ƶ���APPSOft\APPLY��
                                strTmp = UCase(GetVersionInfo(gstrTempPath & "\" & rsTmp!�ļ���, FVN_ProductName))
                                If strTmp = "" Then strTmp = "ZLHISINSTALLUPDATE"
                                If strTmp <> "ZLHISINSTALLUPDATE" Then 'zlHisInstallUpdate
                                    gobjTrace.WriteInfo "UpgradeBase", "ZLHISCRUST.EXE�����ع����ϵͰ汾", True
                                    strFile = gstrSetupPath & "\Apply\" & rsTmp!�ļ���
                                    If gobjFSO.FileExists(strFile) Then
                                        If FileSystem.GetAttr(strFile) <> vbNormal Then
                                             Call FileSystem.SetAttr(strFile, vbNormal)
                                        End If
                                        Call gobjFSO.DeleteFile(strFile)
                                    End If
                                    gobjFSO.CopyFile gstrTempPath & "\" & rsTmp!�ļ���, strFile, False
                                    strCommand = GetHisUpdateCommand(True)
                                Else
                                    gobjTrace.WriteInfo "UpgradeBase", "ZLHISCRUST.EXE�����ع����ϵͰ汾", False
                                    strFile = gstrTempPath & "\" & rsTmp!�ļ���
                                    strCommand = GetHisUpdateCommand()
                                End If
                                '���غ���Ҫʹ���µ�ZLHISCRUST.EXE����������
                                On Error Resume Next
                                Call gobjTrace.CloseLog
                                If Shell(strFile & " " & strCommand, vbNormalFocus) <> 0 Then
                                    Call gclsConnect.CloseConnect
                                    Call gobjMe.ExitApp
                                Else
                                End If
                            End If
                        Else
                            strMsg = "�������ļ�ȱʧZLHISCRUST.EXE(�Զ�����������)"
                        End If
                    End If
                Else
                    strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧZLHISCRUST.EXE(�Զ�����������)"
                End If
            End If
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "�Զ��������߼��", strMsg
            Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
            If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
        strMsg = ""
        '3.1 �Զ�����DLL
        rsTmp.Filter = "��׼�ļ���='ZLHISCRUSTCOM.DLL'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If IsFileUpgade(strFile, rsTmp!�汾�� & "", rsTmp!�޸����� & "", rsTmp!MD5 & "") Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gstrTempPath, strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(ZLHISCRUSTCOM.DLL(�Զ�����ҵ������))" & strErr
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "����(�쳣)", gstrTempPath & "\" & rsTmp!�ļ���
                    End If
                Else
                    strMsg = "�������ļ�ȱʧZLHISCRUSTCOM.DLL(�Զ�����ҵ������)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧZLHISCRUSTCOM.DLL(�Զ�����ҵ������)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "�Զ��������߼��", strMsg
            Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
            If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
        
        strMsg = ""
        '4������ѹ�����ߣ��Ա��������������Ľ�ѹ
        rsTmp.Filter = "��׼�ļ���='7Z.DLL'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(7Z.DLL(��ѹ������������))" & strErr
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧ7Z.DLL(��ѹ������������)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧ7Z.DLL(��ѹ������������)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "��ѹ���߼��", "��Ϣ", strMsg
            Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
            If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
        strMsg = ""
        '4������ѹ�����ߣ��Ա��������������Ľ�ѹ
        rsTmp.Filter = "��׼�ļ���='ZL7Z.DLL'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(ZL7Z.DLL(����ѹ������))" & strErr
                    Else
                        strMsg = ""
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                        If Not gclsRegCom.RegCom(strFile, strMsg, RFT_NormalReg) Then
                            gobjTrace.WriteInfo "UpgradeBase", "ZL7Zע��ʧ��", strMsg
                            Call RecordErrMsg(MT_InitEnv, "ZL7Zע��ʧ��", strMsg)
                        Else
                            gobjTrace.WriteInfo "UpgradeBase", "ZL7Zע��ɹ�", ""
                        End If
                        strMsg = ""
                    End If
                Else
                    strMsg = "�������ļ�ȱʧZL7Z.DLL(����ѹ������)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧZL7Z.DLL(����ѹ������)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "��ѹ���߼��", "��Ϣ", strMsg
            Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
            If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
    End If
    strMsg = ""
    rsTmp.Filter = "��׼�ļ���='7Z.EXE'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
        gobj7zZip.Path7z = strFile
        If blnUpgrade Then '������������
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(7Z.EXE(��ѹ����))" & strErr
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧ7Z.EXE(��ѹ����)"
                End If
            End If
        End If
    Else
        strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧ7Z.EXE(��ѹ����)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "��ѹ���߼��", strMsg
        Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
        If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
        Exit Function
    End If
    '5������
    strMsg = ""
    blnMust = IsMustGACUTIL(): blnErr = False
    rsTmp.Filter = "��׼�ļ���='GACUTIL.EXE'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
        gclsRegCom.GACUPath = strFile
        If blnUpgrade Then '������������
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(GACUTIL.EXE(ȫ�ֻ�����ӹ���))" & strErr
                        blnErr = True
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧGACUTIL.EXE(ȫ�ֻ�����ӹ���)"
                End If
            End If
        End If
    Else
        strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧGACUTIL.EXE(ȫ�ֻ�����ӹ���)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "ȫ�ֻ�����ӹ��߼��", strMsg
        If blnMust Or blnErr Then
            Call RecordErrMsg(MT_InitEnv, "ȫ�ֻ�����ӹ��߼��", strMsg)
            If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & ",����ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    If blnUpgrade Then '������������
        strMsg = ""
        blnErr = False
        rsTmp.Filter = "��׼�ļ���='GACUTIL.EXE.CONFIG'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(GACUTIL.EXE.CONFIG(ȫ�ֻ�����ӹ��������ļ�))" & strErr
                        blnErr = True
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧGACUTIL.EXE.CONFIG(ȫ�ֻ�����ӹ��������ļ�)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧGACUTIL.EXE.CONFIG(ȫ�ֻ�����ӹ��������ļ�)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "ȫ�ֻ�����ӹ��߼��", strMsg
            If blnMust Or blnErr Then
                Call RecordErrMsg(MT_InitEnv, "ȫ�ֻ�����ӹ��߼��", strMsg)
                If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & ",����ϵ����Ա��", vbInformation, App.Title
                Exit Function
            End If
        End If
    End If
    If Not gobj7zZip.Init7zZip Then
        gobjTrace.WriteInfo "UpgradeBase", "7zZip��ʼ��", "�޷�����ZL7z������û��7z.exe"
        Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", "�޷�����ZL7z������û��7z.exe")
        If Not gblnHelperMain Then MsgBox "�޷�����ZL7z������û��7z.exe" & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
        Exit Function
    End If
    '���������ַ�����������ֱ���˳��������������ٴ���������
    If UpdateZLHelper Then
        Call gobjMe.ExitApp
        Exit Function
    End If
    UpgradeBase = True
    Exit Function
ErrH:
    gobjTrace.WriteInfo "UpgradeBase", "������������������������", Err.Description
    Call RecordErrMsg(MT_InitEnv, "������������������������", Err.Description)
    If Not gblnHelperMain Then MsgBox "������������������������" & vbNewLine & "������ϵ����Ա����Ϣ��" & Err.Description, vbInformation, App.Title
    Err.Clear
End Function

'--------------------------------------------------------------------------------------------------
'����           UpdateZLHelper
'����           ������������
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Private Function UpdateZLHelper() As Boolean
    Dim strSQL          As String, rsTmp    As ADODB.Recordset
    Dim strFile         As String
    Dim objService      As New clsService
    Dim arrUpdate(2)    As Byte, arrFile(2) As String
    Dim cllProcess      As New Collection   '���̼�array(����,Exe�ļ���,ģ�����)
    Dim lngProcess      As Long
    Dim i               As Long
    Dim strMsg          As String
    Dim strServer       As String
    Dim strError        As String
    Dim blnHaveHelperMain       As Boolean
    Dim objMetux                As New clsMutex
    Dim objSendHelper           As New clsMemoryShareFP
    Dim strDB                   As String
    Dim blnOk                   As Boolean
    Dim blnRunning              As Boolean
    Dim lngHelperMainVersion    As String
    Dim strHelperMainSeting     As String
    
    Const M_SINGLE_INSTANCE             As String = "DEF43DC9-722D-48E0-9CBD-73E20E373E86"          '�������ֵ�ʵ������������֤��ʵ������
    Const G_HELPER_RECEIVE              As String = "67332524-C38A-4318-85C4-FA8151C85EDD"          '���������ַ���������Ϣ���ڴ湲��
    Const ERROR_INVALID_PARAMETER       As Long = &H57
    On Error GoTo ErrH
    strSQL = "Select ���, �ļ���, Upper(�ļ���) ��׼�ļ���," & IIf(gblnHaveVersion, "�ļ��汾��", " ") & " �汾��, �޸�����, �ļ�����, ҵ�񲿼�, Upper(��װ·��) ��װ·��, Md5, �Զ�ע��, ǿ�Ƹ���" & vbNewLine & _
            "From ZLTOOLS.Zlfilesupgrade" & vbNewLine & _
            "Where Upper(�ļ���) In ('ZLHELPERSERVICE.EXE','ZLHELPERMAIN.EXE','ZLSM4.DLL')"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ������������ļ�")
    rsTmp.Filter = "��׼�ļ���='ZLHELPERSERVICE.EXE'"
    If Not rsTmp.EOF Then
        arrFile(0) = gstrSetupPath & "\ZLHELPERSERVICE.EXE"
        If IsFileUpgade(arrFile(0), rsTmp!�汾�� & "", rsTmp!�޸����� & "", rsTmp!MD5 & "") Then
            arrUpdate(0) = 1
        End If
    End If
    rsTmp.Filter = "��׼�ļ���='ZLHELPERMAIN.EXE'"
    If Not rsTmp.EOF Then
        arrFile(1) = gstrSetupPath & "\ZLHELPERMAIN.EXE"
        If IsFileUpgade(arrFile(1), rsTmp!�汾�� & "", rsTmp!�޸����� & "", rsTmp!MD5 & "") Then
            arrUpdate(1) = 1
        End If
    End If
    
    rsTmp.Filter = "��׼�ļ���='ZLSM4.DLL'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
        arrFile(2) = strFile
        If IsFileUpgade(strFile, rsTmp!�汾�� & "", rsTmp!�޸����� & "", rsTmp!MD5 & "") Then
            arrUpdate(2) = 1
        End If
    End If
    
    
    '�������ռ��
    If arrUpdate(0) = 1 Or arrUpdate(1) = 1 Or arrUpdate(2) = 1 Then
        '����ֹͣ�Զ��ص����лỰ�ĺ�̨����
        If objService.IsInstalled("ZLHelperService") Then
            If Not objService.IsStopped("ZLHelperService") Then
                blnRunning = True
            End If
        Else
            blnRunning = True
        End If
        blnHaveHelperMain = objMetux.CheckMutex(M_SINGLE_INSTANCE)
        Set objMetux = Nothing
        gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "�������ĺ�̨��������ڣ�" & blnHaveHelperMain
        If blnHaveHelperMain Then
            gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-��ʼ��HELPERUPGRADE SAVEANDEXIT"
            '10.35.130ΪZLHelperMainSetup��10.35.130����SPΪZLHelperMainSetupV0001
            If gobjFSO.FileExists(arrFile(1)) Then
                lngHelperMainVersion = Val(Mid(GetVersionInfo(arrUpdate(1), FVN_ProductName), Len("ZLHELPERMAINSETUPV*")))
            End If
            gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-HelperMain�汾��" & lngHelperMainVersion
            If objSendHelper.OpenMemoryShare(G_HELPER_RECEIVE) Then
                If objSendHelper.WriteMemory("HELPERUPGRADE SAVEANDEXIT", GetCurrentProcessId) Then
                    If lngHelperMainVersion = 0 Then
                        For i = 1 To 50
                            If objSendHelper.ReadMemory Then
                                If objSendHelper.Writed = 0 Then
                                    gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-ReadMemory.Writed=0"
                                    blnOk = True
                                End If
                            Else
                                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-ReadMemory=False"
                                blnOk = True
                            End If
                        Next
                        Set objSendHelper = Nothing
                    Else
                        Set objSendHelper = Nothing
                        For i = 1 To 50
                            If FindExitsProcess("ZLHELPERMAIN.EXE") = 0 Then
                                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-FindExitsProcess ZLHELPERMAIN"
                                blnOk = True
                                Exit For
                            Else
                                Call Sleep(100)
                            End If
                        Next
                    End If
                Else
                    gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-WriteMemoryʧ��"
                    blnOk = True
                End If
            Else
                blnOk = True
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-OpenMemoryShareʧ��"
            End If
            
            If blnOk Then
                If Not objService.IsStopped("ZLHelperService") Then
                    If objService.Stopping("ZLHelperService") Then
                        gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-Stop ZLHelperService"
                    Else
                        gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-Stop ZLHelperService-ʧ��"
                    End If
                End If
            End If
        End If
        Set objSendHelper = Nothing
        If Not objService.IsStopped("ZLHelperService") Then
            If objService.Stopping("ZLHelperService") Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-Stop ZLHelperService"
            Else
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��ʼ֪ͨ���������������˳�-Stop ZLHelperService-ʧ��"
            End If
        End If
        lngProcess = FindExitsProcess("ZLHELPERSERVICE.EXE", , False)
        If lngProcess <> 0 Then
            Call TerminateProcess(lngProcess, 1&)
        End If
        lngProcess = FindExitsProcess("ZLHELPERMAIN.EXE", , False)
        If lngProcess <> 0 Then
            Call TerminateProcess(lngProcess, 1&)
        End If
        If arrUpdate(2) = 1 Then
            Call zlGetFileProcess(arrFile(2), cllProcess)
            For i = 1 To cllProcess.Count
                Call TerminatePID(cllProcess(i)(0))
            Next
        End If
        If arrUpdate(0) = 1 Then
            strMsg = ""
            If gclsConnect.IsServerFileExists("ZLHELPERSERVICE.EXE.7z") Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "ZLHELPERSERVICE.EXE"
                If Not gclsConnect.DownloadFile("ZLHELPERSERVICE.EXE.7z", gstrTempPath, strError) Then
                    strMsg = "�������ļ��ļ�����ʧ��(ZLHELPERSERVICE.EXE(�������ַ���))" & strError
                Else
                    If Not gobj7zZip.UnZipFile(gstrTempPath & "\ZLHELPERSERVICE.EXE.7z", gstrTempPath & "\ZLHELPERSERVICE.EXE", , strMsg) Then
                        If strMsg = "" Then
                            strMsg = "��ѹ���ļ�" & gstrTempPath & "\ZLHELPERSERVICE.EXE������,���ܱ�ɱ�����ɱ��"
                        Else
                            strMsg = "�ļ���ѹʧ�ܣ�" & strMsg
                        End If
                    Else
                        Call gobjFSO.CopyFile(gstrTempPath & "\ZLHELPERSERVICE.EXE", arrFile(0), True)
                    End If
                End If
            Else
'                strMsg = "�������ļ�ȱʧZLHELPERSERVICE.EXE(�������ַ���)"
            End If
            If strMsg <> "" Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "ZLHELPERSERVICE.EXE", "����", strMsg
                Call RecordErrMsg(MT_InitEnv, "������������", strMsg)
            End If
        End If
        
        If arrUpdate(1) = 1 Then
            strMsg = ""
            If gclsConnect.IsServerFileExists("ZLHELPERMAIN.EXE.7z") Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "ZLHELPERMAIN.EXE"
                If Not gclsConnect.DownloadFile("ZLHELPERMAIN.EXE.7z", gstrTempPath, strError) Then
                    strMsg = "�������ļ��ļ�����ʧ��(ZLHELPERMAIN.EXE(��������))" & strError
                Else
                    strMsg = ""
                    If Not gobj7zZip.UnZipFile(gstrTempPath & "\ZLHELPERMAIN.EXE.7z", gstrTempPath & "\ZLHELPERMAIN.EXE", , strMsg) Then
                        If strMsg = "" Then
                            strMsg = "��ѹ���ļ�" & gstrTempPath & "\ZLHELPERMAIN.EXE������,���ܱ�ɱ�����ɱ��"
                        Else
                            strMsg = "�ļ���ѹʧ�ܣ�" & strMsg
                        End If
                    Else
                        Call gobjFSO.CopyFile(gstrTempPath & "\ZLHELPERMAIN.EXE", arrFile(1), True)
                    End If
                End If
            Else
'                strMsg = "�������ļ�ȱʧZLHELPERMAIN.EXE(��������)"
            End If
            If strMsg <> "" Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "ZLHELPERMAIN.EXE", "����", strMsg
                Call RecordErrMsg(MT_InitEnv, "������������", strMsg)
            End If
        End If
        If arrUpdate(2) = 1 Then
            strMsg = ""
            If gclsConnect.IsServerFileExists("ZLSM4.DLL.7z") Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "ZLSM4.DLL"
                If Not gclsConnect.DownloadFile("ZLSM4.DLL.7z", gstrTempPath, strError) Then
                    strMsg = "�������ļ��ļ�����ʧ��(ZLSM4.DLL(�����㷨����))" & strError
                Else
                    If Not gobj7zZip.UnZipFile(gstrTempPath & "\ZLSM4.DLL.7z", gstrTempPath & "\ZLSM4.DLL", , strMsg) Then
                        If strMsg = "" Then
                            strMsg = "��ѹ���ļ�" & gstrTempPath & "\ZLSM4.DLL������,���ܱ�ɱ�����ɱ��"
                        Else
                            strMsg = "�ļ���ѹʧ�ܣ�" & strMsg
                        End If
                    Else
                        Call gobjFSO.CopyFile(gstrTempPath & "\ZLSM4.DLL", arrFile(2), True)
                    End If
                    
                End If
            Else
                strMsg = "�������ļ�ȱʧZLSM4.DLL(�����㷨����)"
            End If
            If strMsg <> "" Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "ZLSM4.DLL", "����", strMsg
                Call RecordErrMsg(MT_InitEnv, "������������", strMsg)
                If Not gblnHelperMain Then MsgBox strMsg & vbNewLine & ",����ϵ����Ա��", vbInformation, App.Title
            End If
        End If
    End If
    '�ж��������ֽ���
    If gobjFSO.FileExists(arrFile(1)) And gobjFSO.FileExists(arrFile(0)) Then
        '�������񣬲��˳���ǰ���̣��������ֺ�̨�����Զ�����
        If Not objService.IsInstalled("ZLHelperService") Then
            If gobjFSO.FileExists(arrFile(0)) Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��װ����ZLSOFT Upgrade Helper Service"
                Call objService.Install("ZLHelperService", "ZLSOFT Upgrade Helper Service", "�����������ַ���", arrFile(0))
            End If
        End If

        If blnRunning Then
            If objService.IsInstalled("ZLHelperService") Then
                If Not objService.IsRunning("ZLHelperService") Then
                    gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "��������ZLSOFT Upgrade Helper Service"
                    If objService.Start("ZLHelperService") Then
                        Sleep 1000
                    End If
                End If
            End If
        End If
        strDB = "EXCFUNC DB=" & GetServerInfo(gcnOracle)
        blnOk = False
        Set objSendHelper = New clsMemoryShareFP
        For i = 1 To 50
            If objSendHelper.OpenMemoryShare(G_HELPER_RECEIVE) Then
                blnOk = True
                Exit For
            Else
                Sleep 100
            End If
        Next
        gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "֪ͨ�������ֵ�ǰ��������Ϣ-��ʼ��" & strDB
        If blnOk Then
            If objSendHelper.WriteMemory(strDB, GetCurrentProcessId) Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "֪ͨ�������ֵ�ǰ��������Ϣ-�ɹ�"
            Else
                If Shell(arrFile(1) & " " & strDB, vbNormalNoFocus) = 0 Then
                    gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "֪ͨ�������ֵ�ǰ��������Ϣ-�ɹ�1"
                Else
                    gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "֪ͨ�������ֵ�ǰ��������Ϣ-ʧ��1"
                End If
            End If
        Else
            If Shell(arrFile(1) & " " & strDB, vbNormalNoFocus) = 0 Then
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "֪ͨ�������ֵ�ǰ��������Ϣ-�ɹ�2"
            Else
                gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "֪ͨ�������ֵ�ǰ��������Ϣ-ʧ��2"
            End If
        End If
        gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "֪ͨ�������ֵ�ǰ��������Ϣ-����"
    End If
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    gobjTrace.WriteInfo "UpgradeZLHelper", "������������", "�������������������̷�������" & Err.Description
    Call RecordErrMsg(MT_InitEnv, "������������", "�������������������̷�������" & Err.Description)
    '�������񣬲��˳���ǰ���̣��������ֺ�̨�����Զ�����
    If gobjFSO.FileExists(arrFile(1)) And gobjFSO.FileExists(arrFile(0)) Then
        If Not objService.IsInstalled("ZLHelperService") Then
            If gobjFSO.FileExists(arrFile(0)) Then
                Call objService.Install("ZLHelperService", "ZLSOFT Upgrade Helper Service", "�����������ַ���", arrFile(0))
            End If
        End If
        If blnRunning Then
            If objService.IsInstalled("ZLHelperService") Then
                If Not objService.IsRunning("ZLHelperService") Then
                    Call objService.Start("ZLHelperService")
'                    UpdateZLHelper = True
                End If
            End If
        End If
    End If
End Function

Private Function GetServerInfo(ByVal cnOracle As ADODB.Connection) As String
'���ܣ���ȡIP:Port/SID��Ϣ
    Dim strServerInfo       As String
    Dim strIp               As String, strPort      As String, strSID       As String
    If IsOLEDBConnection(cnOracle) Then
        '(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.60)(PORT=1522))(CONNECT_DATA=(SERVICE_NAME=qzyy)))
        'Testbase
        strServerInfo = UCase(Trim(Replace(cnOracle.Properties("Data Source Name"), " ", "")))
    Else
        'Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.60)(PORT=1522))(CONNECT_DATA=(SERVICE_NAME=qzyy)))
        'Driver={Microsoft ODBC for Oracle};Server=Testbase
        strServerInfo = Replace(cnOracle.Properties("Extended Properties"), " ", "")
        strServerInfo = UCase(Trim(Mid(strServerInfo, InStrRev(strServerInfo, "Server=") + Len("Server="))))
    End If
    If InStr(strServerInfo, "=") = 0 Then
        Call GetServerInfoByFile(strServerInfo, strSID, strIp, strPort)
        If strSID <> "" And strIp <> "" And strPort <> "" Then
            GetServerInfo = strIp & ":" & strPort & "/" & strSID
        Else
            GetServerInfo = strServerInfo
        End If
    Else
        If InStr(strServerInfo, "HOST=") > 0 Then
            strIp = Mid(strServerInfo, InStr(strServerInfo, "HOST=") + Len("HOST="))
            strIp = Trim(Mid(strIp, 1, InStr(strIp, ")") - 1))
        End If
        If InStr(strServerInfo, "PORT=") > 0 Then
            strPort = Mid(strServerInfo, InStr(strServerInfo, "PORT=") + Len("PORT="))
            strPort = Trim(Mid(strPort, 1, InStr(strPort, ")") - 1))
        End If
        If InStr(strServerInfo, "(SID=") > 0 Then
            strSID = Mid(strServerInfo, InStr(strServerInfo, "(SID=") + Len("(SID="))
            strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
        ElseIf InStr(strServerInfo, "(SERVICE_NAME=") > 0 Then
            strSID = Mid(strServerInfo, InStr(strServerInfo, "(SERVICE_NAME=") + Len("(SERVICE_NAME="))
            strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
        End If
        GetServerInfo = strIp & ":" & strPort & "/" & strSID
    End If
End Function

Public Sub GetServerInfoByFile(ByVal strServer As String, ByRef setServiceName As String, strServerIp As String, ByRef strServerPort As String)
    '����:����tnsname.ora�ļ���ȡ������IP���˿ڡ�ʵ����
    '�������: strServer=������
    '�������� setServiceName = ʵ����  strServerIp = ������IP   strServerPort = �������˿�
    Dim strTxt      As String, strFile As String
    Dim lngTmp      As Long, strTmp As String
    Dim lngIndex    As Long, lngPos As Long, i  As Long
    On Error Resume Next
    
    strFile = GetOracleHome()
    If strFile = "" Then Exit Sub
    strFile = strFile & "\network\ADMIN\tnsnames.ora"
    If Not gobjFSO.FileExists(strFile) Then Exit Sub
    
    strTxt = gobjFSO.OpenTextFile(strFile).ReadAll
    strServer = UCase(strServer): strTxt = ConvertStr(strTxt) '��ʽ���ַ�
    strTxt = Mid(strTxt, InStr(1, strTxt, strServer & "="))
    lngIndex = 0
    lngPos = 1
    lngPos = InStr(lngPos, strTxt, "(")
    If lngPos <> 0 Then
        For i = lngPos To Len(strTxt)
            Select Case Mid(strTxt, i, 1)
                Case "("
                    lngIndex = lngIndex + 1
                Case ")"
                    lngIndex = lngIndex - 1
            End Select
            If lngIndex = 0 Then
                Exit For
            End If
        Next
        If lngIndex = 0 Then
            strTxt = Mid(strTxt, 1, i)
        End If
        '��ȡIP
        lngTmp = InStr(1, strTxt, "HOST=")
        strTmp = Mid(strTxt, lngTmp + Len("HOST="))
        strServerIp = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
        
        '��ȡ�˿�
        lngTmp = InStr(1, strTxt, "PORT=")
        strTmp = Mid(strTxt, lngTmp + Len("PORT="))
        strServerPort = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
        
        '��ȡ������
        lngTmp = InStr(1, strTxt, "SERVICE_NAME=")
        If lngTmp > 0 Then
            strTmp = Mid(strTxt, lngTmp + Len("SERVICE_NAME="))
        Else
            lngTmp = InStr(1, strTxt, "SID=")
            strTmp = Mid(strTxt, lngTmp + Len("SID="))
        End If
        
        setServiceName = Mid(strTmp, 1, InStr(1, strTmp, ")") - 1)
    End If
End Sub

Public Function ConvertStr(ByVal strSource As String) As String
    '����:ȥ���ַ����Ŀո�\���з�,��ת��Ϊ��д
    
    strSource = UCase(strSource)
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, vbNewLine, "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    strSource = Replace(strSource, vbTab, "")
    strSource = Replace(strSource, vbBack, "")
    ConvertStr = strSource
End Function

Public Function IsOLEDBConnection(ByVal cnMain As ADODB.Connection) As Boolean
'���ܣ��жϵ�ǰ�����Ƿ���OraOLEDB����
'����Provider���жϣ��������ַ�ʽ
'��ʽһ��'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
'��ʽ����
'.Provider = "OraOLEDB.Oracle"
'.Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUserName, strPassWord
'�����ַ�ʽ�����Զ�����.Provider����
    'ʹ��Like����Ϊ���ܺ������Ӱ汾��OraOLEDB.Oracle.1
    If UCase(cnMain.Provider) Like "ORAOLEDB.ORACLE*" Then
        IsOLEDBConnection = True
    End If
End Function

Private Function StartZLRunAs(ByVal strPath As String) As String
'���ܣ�����ZLRunas
    Dim strSQL          As String, rsTmp    As ADODB.Recordset
    Dim strUser         As String, strPwd   As String
    Dim strCommandPara  As String, strMsg   As String, strReturn As String
    Dim blnOk           As Boolean
    Dim objShell        As New clsShell
    
    On Error Resume Next
    strSQL = "Select Max(����Ա�û�) ����Ա, Max(����Ա����)  ���� From ZLTOOLS.zlClients Where ����վ = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��ǰ�ͻ��˵�¼���")
    '����ģʽ���Ͱ汾û���������ֶ�
    If Err.Number = 0 Then
        strUser = NVL(rsTmp!����Ա, "Administrator")
        strPwd = Trim(rsTmp!���� & "")
    Else
        Err.Clear
    End If
    On Error GoTo ErrH
    '�������
    If strPwd <> "" And strUser <> "" Then
        strPwd = DeCipher(strPwd)
        strCommandPara = "-u " & strUser & " -p " & strPwd  '����ZLRunas.EXE������
        gobjTrace.WriteInfo "StartZLRunAs", "�ͻ��˹������", Cipher(strCommandPara)
        '���������������
        If objShell.Run(strPath & " " & strCommandPara & " -ex """ & gstrAppPath & "\ZLHISCRUST.EXE"" -lwp", strReturn, , 30000) Then
            If InStr(strReturn, (1326)) > 0 Then
                strMsg = "��¼ʧ��: δ֪���û�����������롣"
            ElseIf InStr(strReturn, (1058)) > 0 Then
                strMsg = "�޷���������ԭ�������SecLogon���񱻽��á�"
            ElseIf InStr(strReturn, (1717)) > 0 Then
                strMsg = "'·���в��������ģ�����ִ�в��ɹ�"
            Else
                blnOk = True
            End If
        End If
    Else
        gobjTrace.WriteInfo "StartZLRunAs", "�ͻ��˹������", "û��ͳһ��������"
    End If
    'ʹ��ÿ���ͻ��˵ĸ�������
    If Not blnOk Then
        strSQL = "Select Max(Decode(��Ŀ, '����Ա�˺�', ����, '')) As ����Ա, Max(Decode(��Ŀ, '����Ա����', ����, '')) As ����" & vbNewLine & _
                "From Zltools.Zlreginfo" & vbNewLine & _
                "Where ��Ŀ = '����Ա�˺�' Or ��Ŀ = '����Ա����'"
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡͳһ���")
        strUser = NVL(rsTmp!����Ա, "Administrator")
        strPwd = Trim(rsTmp!���� & "")
        If strPwd <> "" And strUser <> "" Then
            strPwd = DeCipher(strPwd)
            strCommandPara = "-u " & strUser & " -p " & strPwd  '����ZLRunas.EXE������
            gobjTrace.WriteInfo "StartZLRunAs", "��ǰ�ͻ��˵�¼���", Cipher(strCommandPara)
            '���������������
            If objShell.Run(strPath & " " & strCommandPara & " -ex """ & gstrAppPath & "\ZLHISCRUST.EXE"" -lwp", strReturn, , 30000) Then
                If InStr(strReturn, (1326)) > 0 Then
                    strMsg = "��¼ʧ��: δ֪���û�����������롣"
                ElseIf InStr(strReturn, (1058)) > 0 Then
                    strMsg = "�޷���������ԭ�������SecLogon���񱻽��á�"
                ElseIf InStr(strReturn, (1717)) > 0 Then
                    strMsg = "'·���в��������ģ�����ִ�в��ɹ�"
                Else
                    blnOk = True
                End If
            End If
        Else
            gobjTrace.WriteInfo "StartZLRunAs", "��ǰ�ͻ��˵�¼���", "û�е�¼�������"
        End If
    End If
    StartZLRunAs = strMsg
    Exit Function
ErrH:
    gobjTrace.WriteInfo "StartZLRunAs", "��ȡ�ͻ�����ɷ�����������", Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetUpgradeFileList() As Boolean
'���ܣ���ȡZLFIleUpgrade
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, strMsg As String
    
    On Error GoTo ErrH
    '���ͬ���ļ�
    strSQL = "Select Upper(a.�ļ���) �ļ��� From Zlfilesupgrade a Group By Upper(a.�ļ���) Having Count(1) > 1"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�ļ��嵥")
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!�ļ���
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then
        strMsg = "����ͬ��(��Сд����)������" & Mid(Mid(strTmp, 2), 1, 100)
        gobjTrace.WriteInfo "GetUpgradeFileList", "�����嵥�Ϸ��Լ��", strMsg
        Call RecordErrMsg(MT_InitEnv, "�����嵥�Ϸ��Լ��", strMsg)
        If Not gblnHelperMain Then MsgBox "�����嵥�������⣬����ϵ����Ա���д���" & vbNewLine & strMsg, vbInformation + vbDefaultButton1, App.Title
        Exit Function
    End If
    On Error Resume Next
    strSQL = "Select a.�ļ���, Upper(a.�ļ���) ��׼�ļ���," & IIf(gblnHaveVersion, "a.�ļ��汾�� ", " a.") & "�汾��, a.�޸�����, a.�ļ�����, a.ҵ�񲿼�, a.��װ·��, a.Md5, NVL(a.�Զ�ע��,0) �Զ�ע��, NVL(a.ǿ�Ƹ���,0) ǿ�Ƹ���,���Ӱ�װ·��" & vbNewLine & _
            "From Zltools.Zlfilesupgrade a" & vbNewLine & _
            "Where Upper(a.�ļ���) Not In ('ZLRUNAS.EXE', 'ZLHISCRUST.EXE','ZLHISCRUSTCOM.DLL', '7Z.EXE', '7Z.DLL', 'AAMD532.DLL', 'GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�ļ��嵥")
    If Err.Number <> 0 Then
        Err.Clear
        strSQL = "Select a.�ļ���, Upper(a.�ļ���) ��׼�ļ���, " & IIf(gblnHaveVersion, "a.�ļ��汾�� ", " a.") & "�汾��, a.�޸�����, a.�ļ�����, a.ҵ�񲿼�, a.��װ·��, a.Md5, NVL(a.�Զ�ע��,0) �Զ�ע��, NVL(a.ǿ�Ƹ���,0) ǿ�Ƹ���,Null ���Ӱ�װ·��" & vbNewLine & _
                "From Zltools.Zlfilesupgrade a" & vbNewLine & _
                "Where Upper(a.�ļ���) Not In ('ZLRUNAS.EXE', 'ZLHISCRUST.EXE','ZLHISCRUSTCOM.DLL', '7Z.EXE', '7Z.DLL', 'AAMD532.DLL', 'GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
        Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�ļ��嵥")
    End If
    'ʵ��·��-��װ·��ת��Ϊʵ��·��
    '�����ļ�·��-����·���ļ�
    Set grsFileUpgrade = CopyNewRec(rsTmp, , , Array("����", adInteger, 1, 0, "ʵ��·��", adVarChar, 500, Empty, "�����ļ�·��", adVarChar, 1000, Empty, "����ʵ��·��", adVarChar, 4000, Empty, _
                                                "�ж�����", adInteger, 3, 0, "Ԥ��������", adInteger, 1, 0, "������Ϣ", adVarChar, 1000, Empty, "�����Ϣ", adVarChar, 1000, Empty, _
                                                "�޺�׺�ļ���", adVarChar, 100, Empty, "��������", adInteger, 1, 0, "ע�����", adInteger, 1, 0))
    GetUpgradeFileList = True
    Exit Function
ErrH:
    gobjTrace.WriteInfo "GetUpgradeFileList", "�����嵥��ȡʧ��", Err.Description
    Call RecordErrMsg(MT_InitEnv, "�ļ��嵥��ȡ", Err.Description)
    If Not gblnHelperMain Then MsgBox "�����嵥��ȡʧ�ܣ�" & vbNewLine & "����ϵ����Ա����Ϣ��" & Err.Description, vbInformation, App.Title
End Function

Public Function GetKILLProcess() As Boolean
'���ܣ���ȡҪɱ���Ľ���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error Resume Next
    strSQL = "Select ���, ����,���� From Zltools.ZlkillProcess Order By ���"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�ļ��嵥")
    If rsTmp Is Nothing Then
        If Err.Number <> 0 Then Err.Clear
    Else
        Do While Not rsTmp.EOF
            strTmp = strTmp & ";" & UCase(rsTmp!����)
            rsTmp.MoveNext
        Loop
    End If
    
    If strTmp = "" Then
        strTmp = "zl9LabPrintSvr.exe;zl9LabReceiv.exe;zl9LabTcpSvr.exe;Zl9LISComm.exe;zl9PacsCapture.exe;zl9WizardMain.exe;zl9WizardStart.exe;ZL9Xls.exe;zlActMain.exe;ZLBAExport.exe;zlCDOpen.exe;zlCisAuditPrint.exe;zlDrugMachineManage.exe;zlGetImage.exe;zlGetImageEx.exe;zlHQMSDCollect.exe;zlLisReceiveSend.exe;zlMipClientManage.exe;zlMipClientPoll.exe;zlMipClientShell.exe;zlMsgBuilderStart.exe;zlMsgReceiver.exe;zlMsgSender.exe;ZLNewQuery.exe;zlOrclConfig.exe;ZLPacsBrowserStation.exe;ZlPacsSrv.exe;zlPeisAutoAnalyse.exe;zlQueueShow.exe;ZLRPTSQLAdjust.exe;ZLRUNAS.EXE;zlScreenKeyboard.exe;zlSoftShowArchive.exe;zlSvrNotice.exe;zlSvrStudio.exe;zlUpgradeReader.exe;zlWizardStart.exe;ZLPacsServerCenter.exe"
    Else
        strTmp = Mid(2, strTmp)
    End If
    gobjTrace.WriteInfo "GetKILLProcess", "�����嵥", strTmp
    garrKillProcess = Split(UCase(strTmp), ";")
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function IsMustGACUTIL() As Boolean
'���ܣ��Ƿ����ҪGACUTIL.EXE��GACUTIL.EXE.CONFIG
    Dim strSQL As String, rsTmp As ADODB.Recordset

    On Error GoTo ErrH
    strSQL = "Select Count(1) ���� From Zlfilesupgrade a Where a.�Զ�ע�� = [1] And a.Md5 Is Not Null"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�ļ��嵥", RFT_NETGAC)
    IsMustGACUTIL = rsTmp!���� > 0
    Exit Function
ErrH:
    gobjTrace.WriteInfo "IsMustGACUTIL", "��ȡGACUTILע�Ჿ��", Err.Description
    Err.Clear
End Function

