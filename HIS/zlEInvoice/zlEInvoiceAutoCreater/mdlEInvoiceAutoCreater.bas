Attribute VB_Name = "mdlEInvoiceAutoCreater"
Option Explicit '��ģ�����ڴ���漰���ݿ���ʵĹ�������
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��

Public glngSys As Long
Public glngModul As Long
Public gstrDBUser As String                 '��ǰ���ݿ��û�

Public gstrSysName As String                'ϵͳ����
Public gstrProductName As String            'OEM��Ʒ����

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gblnExecuting As Boolean
Public gfrmMain As Object
Public glngSplitTime As Long '���ʱ�䣬��
Private mobjPubEInvoice As Object 'zlPublicExpense.clsPubEInvoice

Private Function GetUserInfo() As Boolean
    '���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.��� = rsTmp!���
            UserInfo.����ID = zlCommFun.NVL(rsTmp!����ID, 0)
            UserInfo.���� = zlCommFun.NVL(rsTmp!����)
            UserInfo.���� = zlCommFun.NVL(rsTmp!����)
            GetUserInfo = True
        End If
    End If
End Function

Public Sub Main()
    Dim objRelogin As Object, strPrivs As String
    
    On Error GoTo ErrHandler
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "�������")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    On Error Resume Next
    Set objRelogin = CreateObject("ZLLogin.clsLogin")
    If objRelogin Is Nothing Then
        MsgBox "���� ZLLogin.clsLogin ����ʧ�ܡ������Ƿ���ȷע��  ZLLogin ������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set gcnOracle = objRelogin.Login(1, CStr(Command()))
    If gcnOracle Is Nothing Then
        Set objRelogin = Nothing
        Exit Sub
    End If
    
    glngSys = 100
    glngModul = 1145
    gstrDBUser = objRelogin.DBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ������ϵͳ����Ա��ϵ���ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strPrivs = GetPrivFunc(glngSys, glngModul)
    If zlStr.IsHavePrivs(strPrivs, "���ߵ���Ʊ��") = False Then MsgBox "�㲻�߱����ߵ���Ʊ�ݵ�Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub
    
    frmEInvoiceManager.ShowMe strPrivs
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Function GetPubEInvoiceObject(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
    objPubEInvoice As Object, Optional ByVal byt���� As Byte = 1, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ʊ�ݹ����ӿڲ���
    '���:
    '   frmMain�����õ�������
    '   lngModule����ǰ����ģ���
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '   blnDeviceSet���豸���õ��õĳ�ʼ��
    '����:
    '����:��ʼ���ɹ�����true,���򷵻�False
    '˵��:
    '   1.ʹ�ñ�����ǰ,�����ȵ��ñ��ӿڽ��г�ʼ��
    '   2.��ʼ���ӿ�,��HIS����ģ��ʱ����(���磺�����շѹ������)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExtend As String
    
    If objPubEInvoice Is Nothing Then
        On Error Resume Next
        Set objPubEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
        If Err <> 0 Then
            strErrMsg_Out = "�����ڿ��õĵ���Ʊ�ݽӿڲ���(zlPublicExpense.clsPubEInvoice)������ϵͳ����Ա��ϵ����ϸ�Ĵ�����ϢΪ:" & vbCrLf & Err.Description
            Exit Function
        End If
    End If
    If objPubEInvoice Is Nothing Then Exit Function
    
    GetPubEInvoiceObject = objPubEInvoice.zlInitialize(frmMain, byt����, gcnOracle, lngSys, lngModule, False, strExtend)
End Function

Private Function GetSwapCollectFromBalanceID(ByVal byt���� As Byte, ByVal lngԭ����ID As Long, _
    ByRef cllSwapData_Out As Collection, Optional ByVal bln������ As Boolean, _
    Optional ByVal lng����ID As Long, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ���׽�����Ϣ
    '���:
    '    byt����-1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-���￨
    '   lngԭ����ID byt����=2������Ԥ����¼.ID������������ID
    '����:
    '   cllSwapData_Out-���ؽ�����Ϣ
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����),key(_�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, cllBalanceInfo As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, strInsureSql As String
    
    On Error GoTo ErrHandler
    Select Case byt����
    Case 1, 4
        If bln������ Then
            strWhere = " And b.����id In(Select �շѽ���ID From ���ò����¼ Where ����ID=[1])"
        Else
            strWhere = " And b.����id = [1]"
        End If
    
        strSQL = _
            " Select Max(a.����id) As ����ID, Max(a.��ҳid) As ��ҳID, Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����," & _
            "        f_List2Str(Cast(Collect(a.No) As t_StrList)) As NO, Sum(a.���ʽ��) As ���ʽ��, Max(a.�Ǽ�ʱ��) As �շ�ʱ��" & _
            " From (Select a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.No, a.���, Sum(a.���ʽ��) As ���ʽ��, Max(b.�Ǽ�ʱ��) As �Ǽ�ʱ��" & _
            "        From ������ü�¼ A, ������ü�¼ B" & _
            "        Where Mod(a.��¼����, 10) = Mod(b.��¼����, 10) And a.No = b.No And a.��� = b.���" & strWhere & _
            "        Group By a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.No, a.���" & _
            "        Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0) A"
        
        strInsureSql = "Select Max(����) As ���� From ���ս����¼ Where ���� = 1 And ��¼id = [1]"
        
        strSQL = _
            " Select a.����id, a.��ҳid, a.����, a.�Ա�, a.����, m.�����, Nvl(n.סԺ��, m.סԺ��) As סԺ��," & _
            "           a.No, a.���ʽ��, a.�շ�ʱ��, b.����" & _
            " From (" & strSQL & ") A, (" & strInsureSql & ") B, ������Ϣ M, ������ҳ N" & _
            " Where a.����id = m.����id(+) And a.����id = n.����id(+) And a.��ҳid = n.��ҳid(+) And a.No Is Not Null"
    Case 2
        strSQL = _
            "   Select a.Id, a.No, a.����id, a.��ҳid, Sum(A.���) As ���ʽ��, Max(A.Ԥ������Ʊ��) As �Ƿ����Ʊ��, " & _
            "          Max(Nvl(d.����, c.����)) As ����, " & _
            "          Max(Nvl(d.�Ա�, c.�Ա�)) As �Ա�, Max(Nvl(d.����, c.����)) As ����, Max(Nvl(d.סԺ��, c.סԺ��)) As סԺ��, Max(c.�����) As �����, " & _
            "          max(M.����) as ����,to_char(max(A.�տ�ʱ��),'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ��,max(a.Ԥ�����) as Ԥ�����" & _
            "   From  ����Ԥ����¼ A, ������Ϣ C, ������ҳ D,(Select ��¼ID, ���� From ���ս����¼ where ����=3  and ��¼ID=[1] ) M" & _
            "   Where a.����id = c.����id(+) And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And a.Id=[1]  And A.ID=M.��¼ID(+)" & _
            "   Group By a.Id, a.No, a.����id, a.��ҳid"
    Case 3
        strSQL = _
            "   Select a.Id, a.No, a.����id, a.��ҳid, Sum(b.��Ԥ��) As ���ʽ��, Max(b.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, " & _
            "          Max(decode(nvl(A.����ID,0),0,A.ԭ��,Nvl(d.����, c.����))) As ����, " & _
            "          Max(Nvl(d.�Ա�, c.�Ա�)) As �Ա�, Max(Nvl(d.����, c.����)) As ����, Max(Nvl(d.סԺ��, c.סԺ��)) As סԺ��, Max(c.�����) As �����, " & _
            "          max(M.����) as ����,to_char(max(A.�շ�ʱ��),'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ��,max(A.��������) as ��������" & _
            "   From ���˽��ʼ�¼ A, ����Ԥ����¼ B, ������Ϣ C, ������ҳ D,(Select ��¼ID, ���� From ���ս����¼ where ����=2  and ��¼ID=[1] ) M" & _
            "   Where a.id=b.����ID and  a.����id = c.����id(+) And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And a.Id=[1]  And A.ID=M.��¼ID(+)" & _
            "   Group By a.Id, a.No, a.����id, a.��ҳid"
    Case 5
        strSQL = _
            "   Select a.����id As ID, b.No, a.����id, a.��ҳid, Sum(a.��Ԥ��) As ���ʽ��, Max(a.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, " & _
            "          Max(c.����) As ����, Max(c.סԺ��) As סԺ��, Max(c.�����) As �����, 0 As ����, " & _
            "          To_Char(Max(a.�տ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �շ�ʱ�� " & _
            "   From ����Ԥ����¼ A, (Select  ����id,No From סԺ���ü�¼ Where ����id = [1]) B, ������Ϣ C  " & _
            "   Where a.����id = b.����id And a.����id = c.����id(+)  And a.����id = [1] " & _
            "   Group By a.����id, b.No, a.����id, a.��ҳid"
    Case Else
        strErrMsg_Out = "���볡�ϡ�" & byt���� & "��������Ч��": Exit Function
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���ID��������Ʊ����Ϣ", IIf(byt���� = 2 And lng����ID <> 0, lng����ID, lngԭ����ID))
    If rsTemp.EOF Then strErrMsg_Out = "��ʣ��δ�˷������ݡ�": Exit Function
    
    '1.����������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����)
    Set cllPati = New Collection
    cllPati.Add Val(NVL(rsTemp!����ID)), "_����ID"
    cllPati.Add Val(NVL(rsTemp!��ҳid)), "_��ҳID"
    cllPati.Add NVL(rsTemp!����), "_����"
    cllPati.Add NVL(rsTemp!�Ա�), "_�Ա�"
    cllPati.Add NVL(rsTemp!����), "_����"
    cllPati.Add NVL(rsTemp!�����), "_�����"
    cllPati.Add NVL(rsTemp!סԺ��), "_סԺ��"
    cllPati.Add Val(NVL(rsTemp!����)), "_����"

    '2.����������Ϣ:(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add "", "_��Ʊ��"
    cllBalanceInfo.Add lngԭ����ID, "_����ID"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add NVL(rsTemp!No), "_���ݺ�"
    cllBalanceInfo.Add Format(NVL(rsTemp!�շ�ʱ��), "yyyy-mm-dd HH:MM:SS"), "_�Ǽ�ʱ��"
    If byt���� = 1 Or byt���� = 4 Then
        cllBalanceInfo.Add IIf(bln������, 1, 0), "_�Ƿ񲹽���"
    Else
        cllBalanceInfo.Add 0, "_�Ƿ񲹽���"
    End If
    cllBalanceInfo.Add 0, "_�Ƿ񲿷��˿�"
    cllBalanceInfo.Add UserInfo.���, "_����Ա���"
    cllBalanceInfo.Add UserInfo.����, "_����Ա����"
    cllBalanceInfo.Add Val(NVL(rsTemp!���ʽ��)), "_������"
    cllBalanceInfo.Add 0, "_����ID"
    Select Case byt����
    Case 2
        cllBalanceInfo.Add Decode(Val(NVL(rsTemp!Ԥ�����)) = 0, 3, Val(NVL(rsTemp!Ԥ�����))), "_��������" 'Ԥ�����:1-����;2-סԺ ;3-�����סԺ;
        cllBalanceInfo.Add 0, "_��Լ��λ����"
    Case 3
        cllBalanceInfo.Add Decode(Val(NVL(rsTemp!��������)) = 0, 3, Val(NVL(rsTemp!��������))), "_��������"  '��������:1-����;2-סԺ ;3-�����סԺ;
        cllBalanceInfo.Add IIf(Val(NVL(rsTemp!����ID)) = 0, 1, 0), "_��Լ��λ����"
    Case Else
        cllBalanceInfo.Add 1, "_��������"
        cllBalanceInfo.Add 0, "_��Լ��λ����"
    End Select
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    
    GetSwapCollectFromBalanceID = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Private Function GetExseData(ByVal bytҵ�񳡺� As Byte, ByVal str�շ�Ա As String, _
    ByVal dt��ʼʱ�� As Date, ByVal dt����ʱ�� As Date, ByRef rsExse As ADODB.Recordset, ByRef strErrMsg_Out As String) As Boolean
    '��ȡ����Ʊ�ݷ�������
    '��Σ�
    '   bytҵ�񳡺� 0-���У�1-�շѣ�2-Ԥ����3-���ʣ�4-�Һţ�5-���￨
    Dim strSQL As String, strWhere As String, strSqlSub As String
    
    On Error GoTo ErrHandler
    strWhere = " And a.�տ�ʱ�� Between [1] And [2]"
    If Trim(str�շ�Ա) <> "" Then strWhere = strWhere & " And a.����Ա����=[3]"
    
    '1)Ԥ����
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 2 Then
        strSQL = _
            " Select 2 As ҵ������, a.Id As ����ID, a.No, a.���, a.����Ա����, a.�տ�ʱ��," & _
            "           a.����id, a.��ҳid, Null As ����, Null As �Ա�, Null As ����, a.Ԥ�����, Null As ����ID, Null As ��������, Null As ������" & _
            " From ����Ԥ����¼ A" & _
            " Where a.��¼���� = 1 And a.��¼״̬ = 1 And a.Ԥ������Ʊ�� = 1" & strWhere & _
            "       And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.Id And Ʊ�� = 2 And ��¼״̬ = 1)"
        '����˿�
        strSQL = strSQL & " Union All" & _
            " Select 12 As ҵ������, b.ID As ����ID, a.No, a.���, a.����Ա����, a.�տ�ʱ��," & _
            "           a.����id, a.��ҳid, Null As ����, Null As �Ա�, Null As ����, a.Ԥ�����,a.Id As ����ID, Null As ��������, Null As ������" & _
            " From ����Ԥ����¼ A,����Ԥ����¼ B" & _
            " Where a.��¼���� = 11 And a.��¼״̬ = 1 And a.Ԥ������Ʊ�� = 1" & strWhere & _
            "       And Exists(Select 1 From ����Ԥ����¼ Where ��¼���� = 1 And ���ӱ�־ = 1 And ����id = a.����id)" & _
            "       And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.Id And Ʊ�� = 2 And ��¼״̬ = 1)" & _
            "       And a.No=b.No And b.��¼����=1 And b.��¼״̬ In(1,3)"
    End If
    '2)���￨
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 5 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select 5 As ҵ������, b.����id As ����ID, b.No, Sum(a.���ʽ��) As ���, b.����Ա����, b.�տ�ʱ��," & _
            "           a.����id, a.��ҳid, a.����, a.�Ա�, a.����, Null As Ԥ�����, Null As ����ID, Null As ��������, Null As ������" & _
            " From סԺ���ü�¼ A, סԺ���ü�¼ A1," & _
            "      (Select Distinct a.����id, a.����Ա����, a.�տ�ʱ��, b.No" & _
            "       From ����Ԥ����¼ A, סԺ���ü�¼ B" & _
            "       Where a.����id = b.����ID And b.��¼���� = 5 And b.��¼״̬ In(1,3) And a.�Ƿ����Ʊ�� = 1" & strWhere & _
            "             And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.����Id And Ʊ�� = 5 And ��¼״̬ = 1)) B" & _
            " Where a.No = a1.No And a.��� = a1.��� And a1.����id = b.����ID And a.��¼���� = 5" & _
            " Group By b.����id, b.No, b.����Ա����, b.�տ�ʱ��, a.����id, a.��ҳid, a.����, a.�Ա�, a.����" & _
            " Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0"
    End If
    '3)����
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 3 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select Distinct 3 As ҵ������, a.����id As ����ID, b.No, b.���ʽ�� As ���, a.����Ա����, a.�տ�ʱ��," & _
            "           b.����id, b.��ҳid, Null As ����, Null As �Ա�, Null As ����, Null As Ԥ�����, Null As ����ID, b.��������, Null As ������" & _
            " From ����Ԥ����¼ A, ���˽��ʼ�¼ B" & _
            " Where a.����id = b.ID And b.��¼״̬ = 1 And a.�Ƿ����Ʊ�� = 1" & strWhere & _
            "       And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.����Id And Ʊ�� = 3 And ��¼״̬ = 1)"
    End If
    '4)�Һš��շ�
    If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 1 Or bytҵ�񳡺� = 4 Then
        strSqlSub = _
            " Select a.����id, a.����Ա����, a.�տ�ʱ��, Mid(b.No) As No, 0 As ������, Null As ����ID" & _
            " From ����Ԥ����¼ A, ������ü�¼ B" & _
            " Where a.����id = b.����ID And b.��¼���� = [��¼����] And b.��¼״̬ In(1,3) And a.�Ƿ����Ʊ�� = 1" & strWhere & _
            "             And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.����id And Ʊ�� = [Ʊ��] And ��¼״̬ = 1)" & _
            "             And Not Exists(Select 1 From ���ò����¼ Where �շѽ���id = a.����id And ��¼���� = 1 And Nvl(���ӱ�־,0) = [���ӱ�־])" & _
            " Group By a.����id, a.����Ա����, a.�տ�ʱ��"
        '�������
        strSqlSub = strSqlSub & " Union All " & _
            " Select ����id, ����Ա����, �տ�ʱ��, No, 1 As ������, ����ID" & _
            " From (Select Distinct b.�շѽ���ID As ����id, a.����Ա����, a.�տ�ʱ��,b.No As No, b.����ID," & _
            "                    Row_Number() Over(Partition By b.��¼����, b.No Order By b.�Ǽ�ʱ��) As ���" & _
            "            From ����Ԥ����¼ A, ���ò����¼ B" & _
            "            Where a.����ID=b.����ID And b.��¼���� = 1 And Nvl(b.���ӱ�־,0) = [���ӱ�־] And b.��¼״̬ In(1, 3) And a.�Ƿ����Ʊ�� = 1" & strWhere & _
            "                       And Not Exists(Select 1 From ����Ʊ��ʹ�ü�¼ Where ����id = a.����id And Ʊ�� = [Ʊ��] And ��¼״̬ = 1))" & _
            " Where ��� = 1"
            
        strSqlSub = _
            " Select [ҵ������] As ҵ������, Nvl(b.����ID, b.����id) As ����ID, Min(b.No) As No, Sum(a.���ʽ��) As ���, b.����Ա����, b.�տ�ʱ��," & _
            "        a.����id, a.��ҳid, a.����, a.�Ա�, a.����, Null As Ԥ�����, Null As ����ID, Null As ��������, b.������" & _
            " From ������ü�¼ A, ������ü�¼ A1,(" & strSqlSub & ") B" & _
            " Where a.No = a1.No And a.��� = a1.��� And a1.����id = b.����ID And Mod(a.��¼����,10)=[��¼����]" & _
            " Group By Nvl(b.����ID, b.����id), b.����Ա����, b.�տ�ʱ��, a.����id, a.��ҳid, b.������, a.����, a.�Ա�, a.����" & _
            " Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0"
        
        If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 1 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(Replace(strSqlSub, "[ҵ������]", 1), "[��¼����]", 1), "[Ʊ��]", 1), "[���ӱ�־]", 0)
        End If
        
        If bytҵ�񳡺� = 0 Or bytҵ�񳡺� = 4 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(Replace(strSqlSub, "[ҵ������]", 4), "[��¼����]", 4), "[Ʊ��]", 4), "[���ӱ�־]", 1)
        End If
    End If
    
    strSQL = _
        " Select Nvl(n.����,Nvl(m.����,a.����)) As ����,Nvl(n.�Ա�,Nvl(m.�Ա�,a.�Ա�)) As �Ա�,Nvl(n.����,Nvl(m.����,a.����)) As ����," & _
        "           m.����� As �����, Nvl(n.סԺ��,m.סԺ��) As סԺ��, a.ҵ������, a.����id, a.No, a.���, a.����Ա����, a.�տ�ʱ��, " & _
        "           a.����id, a.��ҳid, a.Ԥ�����, a.����id, a.��������, a.������" & _
        " From (" & strSQL & ") A, ������Ϣ M, ������ҳ N" & _
        " Where a.����ID=m.����ID(+) And a.����ID=n.����ID(+) And a.��ҳID=n.��ҳID(+)" & _
        " Order By �տ�ʱ��"
    Set rsExse = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ʊ������", dt��ʼʱ��, dt����ʱ��, str�շ�Ա)
    GetExseData = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    '��ʱ���ص�����
    If gblnExecuting Then Exit Sub
    gblnExecuting = True
    Call AutoCreateEInvoice
    gblnExecuting = False
End Sub

Private Function AutoCreateEInvoice() As Boolean
    '�Զ����ߵ���Ʊ��
    Dim dtBegin  As Date, dtEnd As Date
    Dim rsExse As ADODB.Recordset
    Dim strErrMsg As String, byt���� As Byte, bytPre���� As Byte, blnInit As Boolean
    Dim lng����ID As Long, bln������ As Boolean, lng����ID As Long
    Dim cllSwapData As Collection
    
    On Error GoTo ErrHandler
    dtEnd = zlDatabase.Currentdate
    dtBegin = DateAdd("n", -1 * glngSplitTime, dtEnd)
    
    zlWritLog glngModul, "��ȡ�Զ����ߵ���Ʊ�ݷ�������", "AutoCreateEInvoice", "�շ�ʱ�䣺" & Format(dtBegin, "yyyy-MM-dd HH:mm:ss") & "��" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss")
    If GetExseData(0, "", dtBegin, dtEnd, rsExse, strErrMsg) = False Then
        zlWritLog glngModul, "��ȡ�Զ����ߵ���Ʊ�ݷ�������ʧ��", "AutoCreateEInvoice", "[������]" & strErrMsg
        Exit Function
    End If
    
    If rsExse.EOF Then
        zlWritLog glngModul, "��ȡ�Զ����ߵ���Ʊ�ݷ����������", "AutoCreateEInvoice", "��������Ҫ���ߵ���Ʊ�ݵķ������ݡ�"
        Exit Function
    End If
    
    rsExse.Sort = "ҵ������,�տ�ʱ��"
    Do While Not rsExse.EOF
        byt���� = Val(NVL(rsExse!ҵ������)) 'Array("1-�շ�", "2-Ԥ��", "3-����", "4-�Һ�", "5-���￨")
        lng����ID = Val(NVL(rsExse!����ID))
        If byt���� = 1 Or byt���� = 4 Then
            bln������ = Val(NVL(rsExse!������)) = 1
        ElseIf byt���� = 12 Then '����˿�
            lng����ID = Val(NVL(rsExse!����ID))
        End If
    
        byt���� = byt���� Mod 10
        If byt���� <> bytPre���� Then
            bytPre���� = byt����
            blnInit = True
            zlWritLog glngModul, "�������ù�������", "AutoCreateEInvoice", "ҵ�񳡺�=" & byt����
            If GetPubEInvoiceObject(gfrmMain, glngSys, glngModul, mobjPubEInvoice, byt����, strErrMsg) = False Then
                blnInit = False
                zlWritLog glngModul, "�������ù�������ʧ��", "AutoCreateEInvoice", strErrMsg
            End If
        End If
        
        If blnInit Then
            zlWritLog glngModul, "���ݽ���ID��ȡ���׽�����Ϣ", "AutoCreateEInvoice", "����ID=" & lng����ID
            If GetSwapCollectFromBalanceID(byt����, lng����ID, cllSwapData, bln������, lng����ID, strErrMsg) = False Then
                zlWritLog glngModul, "���ݽ���ID��ȡ���׽�����Ϣʧ��", "AutoCreateEInvoice", strErrMsg
            Else
                zlWritLog glngModul, "���ߵ���Ʊ��", "AutoCreateEInvoice", "����ID=" & lng����ID
                If mobjPubEInvoice.zlOnlyCreateEinvoice(gfrmMain, byt����, cllSwapData, Nothing, False, strErrMsg) = False Then
                    zlWritLog glngModul, "���ߵ���Ʊ��ʧ��", "AutoCreateEInvoice", strErrMsg
                End If
            End If
        End If
        
        rsExse.MoveNext
    Loop
    AutoCreateEInvoice = True
    Exit Function
ErrHandler:
    zlWritLog glngModul, "�Զ����ߵ���Ʊ��", "AutoCreateEInvoice", "[������]" & Err.Description
End Function

