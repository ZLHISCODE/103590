Attribute VB_Name = "mdl����ũ��"
Option Explicit

Public gcn����ũ�� As New ADODB.Connection

Public Function openConn����ũ��() As Boolean
    Dim rsTemp As New ADODB.Recordset, str����ֵ As String, strUser As String, strServer As String, _
        strPass As String, strDatabase As String
    On Error GoTo errHandle
    If gcn����ũ��.State <> adStateOpen Then
        gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����ũ��)
        
        Do Until rsTemp.EOF
            str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Select Case rsTemp("������")
                Case "����ũ���û���"
                    strUser = str����ֵ
                Case "����ũ��������"
                    strServer = str����ֵ
                Case "����ũ���û�����"
                    strPass = str����ֵ
                Case "����ũ�����ݿ�"
                    strDatabase = str����ֵ
            End Select
            rsTemp.MoveNext
        Loop
        
        On Error Resume Next
        gcn����ũ��.ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=" & strDatabase & ";Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
        gcn����ũ��.CursorLocation = adUseClient
        gcn����ũ��.Open

        
        If Err <> 0 Then
            MsgBox "ҽ��ǰ�÷���������ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    openConn����ũ�� = True
    Exit Function

errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ʼ��_����ũ��() As Boolean
    Dim rsTemp As New ADODB.Recordset, str����ֵ As String, strUser As String, strServer As String, _
        strPass As String, strDatabase As String
    
    If openConn����ũ��() = False Then Exit Function
    
    gstrSQL = "Select * From ������� Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ����ʼ��", TYPE_����ũ��)
    gstrҽԺ���� = Trim(rsTemp!ҽԺ����)
    
    ҽ����ʼ��_����ũ�� = True
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��ݱ�ʶ_����ũ��(Optional bytType As Byte, Optional lng����ID As Long) As String
'����:ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'����:bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'����:�ջ���Ϣ��
'ע��:1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As Object, strPatiInfo As String
    
    WriteInfo vbCrLf & "��ʼ�����֤"
    
    If bytType = 0 Then
        Set frmIDentified = New frmIdentify����ũ��_����
    Else
        Set frmIDentified = New frmIdentify����ũ��
    End If
    
    strPatiInfo = frmIDentified.GetPatient(bytType)
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ
        lng����ID = BuildPatiInfo(bytType, strPatiInfo, lng����ID, TYPE_����ũ��)
        
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = frmIDentified.mstrPatient & lng����ID & ";" & frmIDentified.mstrOther
        Unload frmIDentified
    Else
        ��ݱ�ʶ_����ũ�� = ""
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    
    WriteInfo "���������֤"
    
    ��ݱ�ʶ_����ũ�� = strPatiInfo
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_����ũ�� = ""
End Function

Public Function �������_����ũ��(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_����ũ��)
    
    If rsTemp.EOF Then
        �������_����ũ�� = 0
    Else
        �������_����ũ�� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If
End Function

Public Function ��Ժ�Ǽ�_����ũ��(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
     '�����˵�״̬�����޸�
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select * From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����ũ��)
    If rsTemp.EOF Then
        MsgBox "�����ʻ���û��ҽ��������Ϣ�����ܰ�����Ժ�Ǽ�", vbInformation, gstrSysName
        Exit Function
    ElseIf IsNull(rsTemp!˳���) Then
        MsgBox "ҽ��������Ϣ�в���ȷ������ID�����ܰ�����Ժ�Ǽ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_����ũ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_����ũ�� = True
End Function

Public Function ��Ժ�Ǽǳ���_����ũ��(lng����ID As Long, lng��ҳID As Long) As Boolean
    '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_����ũ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽǳ���_����ũ�� = True
End Function

Public Function ��Ժ�Ǽ�_����ũ��(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, strSQL As String, datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select * From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����ũ��)
    
    strSQL = "Update inpatient Set outdate='" & Format(datCurr, "yyyy-mm-dd") & "',mark=1 Where id=" & rsTemp!˳���
    gcn����ũ��.Execute strSQL
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_����ũ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    ��Ժ�Ǽ�_����ũ�� = True
End Function

Public Function ��Ժ�Ǽǳ���_����ũ��(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, strSQL As String, datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select * From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����ũ��)
    
    strSQL = "Update inpatient Set outdate=NULL,mark=0 Where id=" & rsTemp!˳���
    gcn����ũ��.Execute strSQL
    '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_����ũ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    ��Ժ�Ǽǳ���_����ũ�� = True
End Function

Public Function סԺ�������_����ũ��(rs��ϸ As ADODB.Recordset, lng����ID As Long) As String
    Dim rsTemp As New ADODB.Recordset, strSQL As String, lngҽ��ID As Long, str��Ժ As String, lng��ҳID As Long, _
        rs��Ŀ As New ADODB.Recordset, str���� As String, datCurr As Date, cur�ܶ� As Currency, int���� As Integer
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    WriteInfo vbCrLf & "��ʼ�ϴ�������ϸ"
    
    gstrSQL = "Select * From �����ʻ� Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lngҽ��ID = rsTemp!˳���
    str���� = Nvl(rsTemp!����֤��, "")
    
    gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = rsTemp(0)
    
    gstrSQL = "Select ��Ժ���� From ������ҳ Where ����id=[1] And ��ҳid=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    str��Ժ = Format(rsTemp(0), "yyyy-mm-dd")
    int���� = CDate(Format(datCurr, "yyyy-mm-dd")) - CDate(str��Ժ)
    
    gstrSQL = "Select * From סԺ���ü�¼ Where �����־=2 And ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and nvl(ʵ�ս��,0)<>0 and" & _
        " ����id=[1] And ��ҳid=[2] order by ��ҳID,���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    
    gcn����ũ��.BeginTrans
    
    While Not rsTemp.EOF
        gstrSQL = "Select nvl(��ע,0) As ��ע From ����֧����Ŀ Where ����=[1] And �շ�ϸĿID=[2]"
        Set rs��Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����ũ��, CLng(rsTemp!�շ�ϸĿID))
        If rs��Ŀ.EOF Then          'ע��ѯ����Ŀ����ҽ����Ŀʱ��δ���
            gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
            Set rs��Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rsTemp!�շ�ϸĿID))
            MsgBox "��Ŀ[" & rs��Ŀ!���� & "]û�ж�Ӧ��ҽ�����룬�����ϴ�����", vbInformation, gstrSysName
            gcn����ũ��.RollbackTrans
            Exit Function
        Else
            strSQL = "Insert Into infee_mx (id,times,[Date],yp_id,sl,je) values (" & lngҽ��ID & ",1,'" & _
                Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:MM:SS") & "'," & rs��Ŀ!��ע & "," & rsTemp!���� * rsTemp!���� & _
                "," & rsTemp!ʵ�ս�� & ")"
            WriteInfo "д���ñ�:" & strSQL
            gcn����ũ��.Execute strSQL
        End If
        
        rsTemp.MoveNext
    Wend
    
    cur�ܶ� = 0
    While Not rs��ϸ.EOF
        cur�ܶ� = cur�ܶ� + rs��ϸ!���
        rs��ϸ.MoveNext
    Wend
    strSQL = "Delete From infee Where id=" & lngҽ��ID
    gcn����ũ��.Execute strSQL
    
    strSQL = "Insert Into infee (ID,times,[jzdate],Days,fee_sum) values (" & lngҽ��ID & ",1,'" & _
        Format(datCurr, "yyyy-mm-dd HH:MM:SS") & "'," & int���� & "," & cur�ܶ� & ")"
    WriteInfo "д�����:" & strSQL
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_����ũ�� & ",'����֤��','''" & Format(datCurr, "yyyy-mm-dd") & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��ID")
    
    gcn����ũ��.Execute strSQL
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rsTemp("ID") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            rsTemp.MoveNext
        Wend
    End If
    
    WriteInfo "��ɷ�����ϸ����"
    gcn����ũ��.CommitTrans
    
    Call UpdateClass(lng����ID, lng��ҳID)
    
    סԺ�������_����ũ�� = "ͳ�����;0;0"
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    gcn����ũ��.RollbackTrans
End Function

Public Function סԺ����_����ũ��(lng����ID As Long, lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, strSQL As String, datCurr As Date
On Error GoTo ErrH
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select * From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����ũ��)
    
    strSQL = "Update inpatient Set outdate='" & Format(datCurr, "yyyy-mm-dd") & "',mark=2 Where id=" & rsTemp!˳���
    gcn����ũ��.Execute strSQL
    
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_����ũ�� & "," & lng����ID & "," & _
        Year(datCurr) & ",0,0,0,0,0,NULL,NULL,NULL,0,0,0,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL,'" & rsTemp!˳��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    סԺ����_����ũ�� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_����ũ��(lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, strSQL As String, datCurr As Date
On Error GoTo ErrH
    
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select ��ע,����ID From ���ս����¼ Where ����=2 And ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If IsNull(rsTemp(0)) Then
        MsgBox "�����¼���Ҳ������˵�ҽ��ID�����ܽ����˽���", vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = "Update inpatient Set outdate='" & Format(datCurr, "yyyy-mm-dd") & "',mark=2 Where id=" & rsTemp(0)
    gcn����ũ��.Execute strSQL
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & rsTemp!����ID & "," & TYPE_����ũ�� & ",'˳���','''" & rsTemp(0) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��ID")
    
    סԺ�������_����ũ�� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function GetItemInfo_����(ByVal lngPatiID As Long, ByVal lngItemID As Long, Optional ByVal strժҪ As String, Optional intType As Integer = 0, Optional ByVal blnMsg As Boolean = False) As String
    Dim rsTemp As New ADODB.Recordset, strTemp As String, int���� As Integer
    
    Dim str���� As String
    
    '��ȡ��ǰ���˵�����
    gstrSQL = "Select * From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lngPatiID)
    int���� = Nvl(rsTemp!����, 0)
    If int���� = 0 Then Exit Function
    
    WriteInfo "��ʼȡ��Ŀ��Ϣ(" & int���� & ")"
    
    gstrSQL = "Select * From ����֧����Ŀ Where ����=[1] And �Ƿ�ҽ��=1 And ��Ŀ���� Is Not Null And �շ�ϸĿID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ŀ��Ϣ", int����, lngItemID)
    If rsTemp.EOF Then
        MsgBox "����Ŀû�ж���,�����Է���Ŀ����,��ע��ʹ��", vbInformation, gstrSysName
        str���� = "�Է�"
        GetItemInfo_���� = str����
        Exit Function
    End If
    
    
    Select Case int����
        Case TYPE_����
            strTemp = Nvl(rsTemp!��ע, "")
            If InStr(strTemp, "��") > 0 Or strTemp = "A������" Then Exit Function
            If strTemp = "" Then
                GetItemInfo_���� = "����ȷ������Ŀ��ҽ�������ע��ʹ��"
            Else
                GetItemInfo_���� = "����Ŀ��ҽ�����Ϊ��" & strTemp & "������ע��ʹ��"
                str���� = strTemp
            End If
        Case TYPE_����ũ��
            str���� = "�Է�"
            Call openConn����ũ��
            strTemp = Nvl(rsTemp!��ע, "")
            If strTemp = "" Then
                MsgBox "����ȷ������Ŀ��ҽ�������ע��ʹ��", vbInformation, gstrSysName
                str���� = "�Է�"
                GetItemInfo_���� = str����
                Exit Function
            End If
            WriteInfo "Select * From price_item Where id=" & strTemp
            Set rsTemp = gcn����ũ��.Execute("Select * From price_item Where id=" & strTemp)
            If rsTemp!yp_bz = True Then
                strTemp = rsTemp!GeneralID
                WriteInfo "Select * From GeneralDrug Where General_ID=" & strTemp
                Set rsTemp = gcn����ũ��.Execute("Select CompenSateMark From GeneralDrug Where General_ID=" & strTemp)
                If rsTemp.EOF Then
                    GetItemInfo_���� = "��ע�⣺ǰ�÷�������û���ҵ�����Ŀ������ȷ��ҽ�����"
                    str���� = "�Է�"
                Else
                    Select Case rsTemp(0)
                        Case 1
                            GetItemInfo_���� = "����Ŀ��ҽ�����Ϊ���Էѡ�����ע��ʹ��"
                            str���� = "�Է�"
                        Case 2
                            'GetItemInfo_���� = "����Ŀ�Ĳ�����ΧΪ���弶������ע��ʹ��"
                            str���� = "�弶"
                        Case 3
                            'GetItemInfo_���� = "����Ŀ�Ĳ�����ΧΪ���缶������ע��ʹ��"
                            str���� = "�缶"
                        Case 4
                            'GetItemInfo_���� = "����Ŀ�Ĳ�����ΧΪ���ؼ�������ע��ʹ��"
                            str���� = "�ؼ�"
                        Case 5
                            'GetItemInfo_���� = "����Ŀ�Ĳ�����ΧΪ���м�������ע��ʹ��"
                            str���� = "�м�"
                        Case 6
                            'GetItemInfo_���� = "����Ŀ�Ĳ�����ΧΪ��ʡ��������ע��ʹ��"
                            str���� = "ʡ��"
                    End Select
                End If
            Else
                strTemp = rsTemp!CenterID
                Set rsTemp = gcn����ũ��.Execute("Select CompenSationMark From FeeItemList Where ID=" & strTemp)
                If rsTemp.EOF Then
                    'GetItemInfo_���� = "��ע�⣺ǰ�÷�������û���ҵ�����Ŀ������ȷ��ҽ�����"
                    str���� = "�Է�"
                Else
                    Select Case rsTemp(0)
                        Case 1
                            GetItemInfo_���� = "����Ŀ��ҽ�����Ϊ���Էѡ�����ע��ʹ��"
                            str���� = "�Է�"
                        Case 2
                            str���� = "����"
                        Case 3
                            GetItemInfo_���� = "����Ŀ��ҽ�����Ϊ�����ࡱ����ע��ʹ��"
                            str���� = "����"
                    End Select
                End If
            End If
        Case TYPE_������
        
    End Select
    If blnMsg = True Then
        If GetItemInfo_���� <> "" Then MsgBox GetItemInfo_����, vbInformation, gstrSysName
    End If
    GetItemInfo_���� = str����
End Function

Public Function �����������_����ũ��(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, Optional ByRef strAdvance As String = "") As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '����Ƿ����δ�������Ŀ
    
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            gstrSQL = "Select nvl(��ע,0) As ��ע From ����֧����Ŀ Where ����=[1] And �շ�ϸĿID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����ĿID", TYPE_����ũ��, CLng(!�շ�ϸĿID))
            If rsTemp.RecordCount = 0 Then         'ע��ѯ����Ŀ����ҽ����Ŀʱ��δ���
                gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡHIS��Ŀ����������", CLng(!�շ�ϸĿID))
                MsgBox "��Ŀ[" & rsTemp!���� & "]û�ж�Ӧ��ҽ�����룬�����ϴ�����", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    'ʲô��������ֱ�ӷ���
    str���㷽ʽ = "�����ʻ�;0;0"
    �����������_����ũ�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_����ũ��(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, Optional ByRef strAdvance As String = "") As Boolean
    Dim lng����ID As Long
    Dim dbl�ܶ� As Double
    Dim strҽ����� As String
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    Call DebugTool("�������")
    '��ȡ���з�����ϸ,��������ϸ����
    gstrSQL = "Select * From ������ü�¼ Where ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID)
    lng����ID = rs��ϸ!����ID
    '2006-4-27 ����ǿ����,����Ҫ��Ʊ��,Ҫ��ȡ��
'    str��Ʊ�� = Nvl(rs��ϸ!ʵ��Ʊ��)
'    If str��Ʊ�� = "" Then
'        Err.Raise 9000,gstrSysName, "��Ʊ�Ų���Ϊ�գ�"
'        Exit Function
'    End If
    
    gcn����ũ��.BeginTrans
    Call DebugTool("׼���ϴ�...")
    '�ϴ�������ϸ
    With rs��ϸ
        Call DebugTool("�����ܶ�")
        dbl�ܶ� = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl�ܶ� = dbl�ܶ� + Val(Format(Nvl(!ʵ�ս��, 0), "#0.00;-#0.00;0;"))
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        Call DebugTool("���������ܱ�")
        '    �Ȳ��������ܱ�
        gstrSQL = "Insert into MZ11(id,xm,[date],tp_bz,je) " & _
            " Values (" & !����ID & ",'" & Nvl(!����) & "','" & _
            Format(!����ʱ��, "yyyy-MM-dd HH:MM:SS") & "',0" & _
            "," & dbl�ܶ� & ")"
        gcn����ũ��.Execute gstrSQL
    End With
    
    With rs��ϸ
        Call DebugTool("����������ϸ��")
        Do While Not .EOF
            gstrSQL = "Select ��ע From ����֧����Ŀ Where ����=[1] And �շ�ϸĿID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҽ����Ŀ���", TYPE_����ũ��, CLng(!�շ�ϸĿID))
            strҽ����� = Nvl(rsTemp!��ע)
            
            '    �ٲ���������ϸ��
            gstrSQL = "Insert Into MZ22(mz11_id,yp_id,sl,je) " & _
                " Values (" & lng����ID & "," & strҽ����� & "," & _
                !���� * Nvl(!����, 1) & "," & Format(Nvl(!ʵ�ս��, 0), "#0.00;-#0.00;0;") & ")"
            gcn����ũ��.Execute gstrSQL
            .MoveNext
        Loop
    End With
    
    Call DebugTool("���汣�ս����¼")
    '���汣�ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_����ũ�� & "," & lng����ID & "," & _
        Year(zlDatabase.Currentdate) & ",0,0,0,0,0,NULL,NULL,NULL,0,0,0,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL,'" & lng����ID & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gcn����ũ��.CommitTrans
    �������_����ũ�� = True
    Exit Function
errHand:
    gcn����ũ��.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_����ũ��(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, Optional ByRef strAdvance As String = "") As Boolean
    Dim lng����ID As Long
    Dim dbl�ܶ� As Double
    Dim strҽ����� As String
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
On Error GoTo errHand
    
    Call DebugTool("�����˷�")
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng����ID = rsTemp!����ID
    
    Call DebugTool("��ȡ���з�����ϸ")
    '��ȡ���з�����ϸ,��������ϸ����
    gstrSQL = "Select * From ������ü�¼ Where ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID)
'    str��Ʊ�� = Nvl(rs��ϸ!ʵ��Ʊ��)
    
    gcn����ũ��.BeginTrans
    Call DebugTool("׼���ϴ���ϸ...")
    '�ϴ�������ϸ
    With rs��ϸ
        Call DebugTool("�����ܶ�")
        dbl�ܶ� = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl�ܶ� = dbl�ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        Call DebugTool("���������ܱ�")
        '    �Ȳ��������ܱ�
        gstrSQL = "Insert into MZ11(id,xm,[date],tp_bz,je) " & _
            " Values (" & lng����ID & ",'" & Nvl(!����) & "','" & _
            Format(!����ʱ��, "yyyy-MM-dd HH:MM:SS") & "',0" & _
            "," & dbl�ܶ� & ")"
        gcn����ũ��.Execute gstrSQL
    End With
    
    Call DebugTool("����������ϸ��")
    With rs��ϸ
        Do While Not .EOF
            gstrSQL = "Select ��ע From ����֧����Ŀ Where ����=[1] And �շ�ϸĿID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҽ����Ŀ���", TYPE_����ũ��, CLng(!�շ�ϸĿID))
            strҽ����� = Nvl(rsTemp!��ע)
            
            '    �ٲ���������ϸ��
            gstrSQL = "Insert Into MZ22(mz11_id,yp_id,sl,je) " & _
                " Values (" & lng����ID & "," & strҽ����� & "," & _
                !���� * Nvl(!����, 1) & "," & !ʵ�ս�� & ")"
            gcn����ũ��.Execute gstrSQL
            .MoveNext
        Loop
    End With
    
    Call DebugTool("���汣�ս����¼")
    '���汣�ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_����ũ�� & "," & lng����ID & "," & _
        Year(zlDatabase.Currentdate) & ",0,0,0,0,0,NULL,NULL,NULL,0,0,0,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL,'" & lng����ID & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gcn����ũ��.CommitTrans
    ����������_����ũ�� = True
    Exit Function
errHand:
    gcn����ũ��.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Sub UpdateClass(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    Dim str�������� As String
    Dim rsTemp As New ADODB.Recordset
    'ѭ������������Ŀ�ķ�������
    gstrSQL = "Select ID,����ID,�շ�ϸĿID,�������� From סԺ���ü�¼" & _
        " Where ����ID=[1] And ��ҳID=[2]" & _
        " And Nvl(�Ƿ��ϴ�,0)=1 And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0 And Nvl(ʵ�ս��,0)<>0 And �������� is null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ѭ������������Ŀ�ķ�������", lng����ID, lng��ҳID)
    
    With rsTemp
        Do While Not .EOF
            str�������� = GetItemInfo_����(!����ID, !�շ�ϸĿID, "", 0, False)
            gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & !ID & ",NULL,NULL,'" & str�������� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���·�������")
            .MoveNext
        Loop
    End With
End Sub
