Attribute VB_Name = "mdl����"
Option Explicit
'Modified By ���� 2005-07-21 10:19:16���޸�ԭ��1����λ�ѳ����޶�������ֶ�Ӧ�Է��룬���²���һ����¼�ϴ���2��������ϸ�����÷���ʱ�������ϴ�
'Modified By ���� 2005-07-25 �޸�ԭ��1����IC������Ϊ�����ļ�����סԺ����ΪסԺ�ļ�����2��סԺ����ʱ��ʵʱ�ϴ���ϸ��3�����˷��ò�ѯ��Ԥ����ʱ�����ض�ȡҽ�������ļ�

Private mcurͳ���� As Currency, mcur����֧�� As Currency
Public gcn���� As New ADODB.Connection, mstrSavePath As String

Public Const MAX_PATH = 260

Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        '�����������
        .hwndOwner = lWindowHwnd
        '����ѡ�е�Ŀ¼
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "��ѡ����ʼ�������ļ��У�"
        Else
            .lpszTitle = sTitle
        End If
    End With
    
    '�����������
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '��ȡ·��
        SHGetPathFromIDList lpIDList, sPath
        '�ͷ��ڴ�
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
End Function


Public Function ҽ����ʼ��_����() As Boolean
'���ܣ������Ƿ�������ӵ�ǰ�÷�������
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs���� As New ADODB.Recordset
    '��������Ѿ��򿪣��ǾͲ����ٲ���
    If gcn����.State = adStateOpen Then
        ҽ����ʼ��_���� = True
        Exit Function
    End If
     
    On Error GoTo ErrH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        If rsTemp!������ = "�ļ����λ��" Then mstrSavePath = rsTemp!����ֵ
        rsTemp.MoveNext
    Loop
    If Trim(mstrSavePath) = "" Then
        MsgBox "�뵽ҽ�����������������ļ����λ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error Resume Next
    gcn����.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=Visual FoxPro Tables;UID=;SourceDB=" & mstrSavePath & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"""
    gcn����.CursorLocation = adUseClient
    gcn����.Open
    
    If Err <> 0 Then
        MsgBox "�ļ����λ��ָ������", vbInformation, gstrSysName
        ҽ����ʼ��_���� = False
        Exit Function
    End If
    ҽ����ʼ��_���� = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    ҽ����ʼ��_���� = False
End Function

Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.ShowME(TYPE_����)
End Function

Public Function �������_����(lng����ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    �������_���� = Nvl(rsTemp!�ʻ����, 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte = 0, Optional lng����ID As Long = 0) As String
    '����ҽ��û�ṩר�ŵ������֤�ӿڣ�ͨ����ȡ�Һŵ�����ʵ����֤
    Dim strTemp As String
    strTemp = frmIdentify����.Identify(bytType, lng����ID)
    Unload frmIdentify����
    If strTemp = "" Then
        MsgBox "δ��ȡ������Ϣ", vbInformation, gstrSysName
    Else
        ��ݱ�ʶ_���� = strTemp
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'��Ϊ����δ�ṩԤ����ӿڣ�������õ��Ľ�������Ϊҽ���������ʽ���ݣ����õ�����ʱҽ������ʽ����
    Dim str��ˮ�� As String, lng����ID As Long, datCurr As Date, strSQL As String
    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng��� As Long
    Dim strCardNO As String
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ã����ܽ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    datCurr = zlDatabase.Currentdate
    lng����ID = rs��ϸ!����ID
    gstrSQL = "Select ���� From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        MsgBox "û���ҵ�������Ϣ��ҽ��ѡ�����", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!����
    '������ˮ��
    str��ˮ�� = strCardNO
    
    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.���� from (select * from ����֧����Ŀ where ����=[1]) A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, CLng(rs��ϸ!�շ�ϸĿID))
        If IsNull((rsTemp!��Ŀ����)) Then
            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
            Exit Function
        End If
        rs��ϸ.MoveNext
    Loop
    
    '����DBF�ļ�
    On Error Resume Next
    gcn����.Execute "Drop Table " & mstrSavePath & "\YM" & str��ˮ��
    
    On Error GoTo errHandle
    gcn����.Execute "Create Table " & mstrSavePath & "\YM" & str��ˮ�� & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
        "SelfJe N(18,4))"
    lng��� = 1
    rs��ϸ.MoveFirst
    While Not rs��ϸ.EOF
        gstrSQL = "Select A.��Ŀ����,B.ID,B.����,B.���,B.���㵥λ From ����֧����Ŀ A,�շ�ϸĿ B Where B.ID=A.�շ�ϸĿID And A.�շ�ϸĿid=[2] And A.����=[1]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, CLng(rs��ϸ!�շ�ϸĿID))             '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
        
        '���š���ˮ�š���š����롢���ơ���Ŀ���롢��񡢼�����λ�����������ۡ����Էѽ��
        gcn����.Execute "Insert Into " & mstrSavePath & "\YM" & str��ˮ�� & " values ('" & strCardNO & "','" & str��ˮ�� & "'," & _
            lng��� & ",'" & Trim(rsTemp!��Ŀ����) & "','" & Trim(rsTemp!����) & "','','" & Trim(rsTemp!���) & "','" & Trim(rsTemp!���㵥λ) & "'," & _
            Format(rs��ϸ!����, "0.####") & "," & Format(rs��ϸ!����, "0.####") & "," & Format(rs��ϸ!ʵ�ս��, "0.####") & "," & _
            "0)"
        lng��� = lng��� + 1
        rs��ϸ.MoveNext
    Wend
    On Error GoTo errHandle
    '�ȴ����ؽ�������
    If frm�ȴ����ػ���.waitReturn(mstrSavePath & "\SM" & str��ˮ��) = False Then
        MsgBox "Ԥ���㱻��ֹ", vbInformation, gstrSysName
        On Error Resume Next
        gcn����.Execute "Drop Table " & mstrSavePath & "\YM" & str��ˮ��
        Unload frm�ȴ����ػ���
        Exit Function
    End If
    Unload frm�ȴ����ػ���
    
    '���ؽ�����
    strSQL = "Select * From " & mstrSavePath & "\SM" & str��ˮ��
    Set rsTemp = gcn����.Execute(strSQL)
    mcur����֧�� = Val(rsTemp!JkAccR)
    mcurͳ���� = Val(rsTemp!JkSocialR)
    str���㷽ʽ = "�����ʻ�;" & Val(rsTemp!JkAccR) & ";0"
    str���㷽ʽ = str���㷽ʽ & "|ͳ�����;" & Val(rsTemp!JkSocialR) & ";0"
    On Error Resume Next
    gcn����.Execute "Drop Table " & mstrSavePath & "\YM" & str��ˮ��
    gcn����.Execute "Drop Table " & mstrSavePath & "\SM" & str��ˮ��
    �����������_���� = True
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency
    Dim datCurr As Date, lng����ID As Long
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� & ",0,0," & _
        "0," & mcurͳ���� & ",0,0," & mcur����֧�� & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, sngArrInfo(20) As Single
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        ����������_���� = False
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        Nvl(rsTemp("�����ʻ�֧��"), 0) * -1 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_����(rs��ϸ As ADODB.Recordset, lng����ID As Long, strҽ���� As String, Optional ByVal bln��ѯ As Boolean = False) As String
'��Ϊ����δ�ṩԤ����ӿڣ�������õ��Ľ�������Ϊҽ���������ʽ���ݣ����õ�����ʱҽ������ʽ����
    Dim str��ˮ�� As String, datCurr As Date, strSQL As String
    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng��� As Long
    Dim strCardNO As String
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ã����ܽ���", vbInformation, gstrSysName
        Exit Function
    End If

    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select A.����,B.סԺ�� From �����ʻ� A,������Ϣ B Where A.����ID=B.����ID And A.����id=[1] And A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        MsgBox "û���ҵ�������Ϣ��ҽ��ѡ�����", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!����
    str��ˮ�� = rsTemp!סԺ��

    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.���� from (select * from ����֧����Ŀ where ����=[1]) A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, CLng(rs��ϸ!�շ�ϸĿID))
        If IsNull((rsTemp!��Ŀ����)) Then
            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
            Exit Function
        End If
        rs��ϸ.MoveNext
    Loop
    ���ʴ���_���� "", 0, "", lng����ID
    
    If bln��ѯ Then Exit Function
    On Error GoTo errHandle
    
    '�ȴ����ؽ�������
    If frm�ȴ����ػ���.waitReturn(mstrSavePath & "\SZ" & str��ˮ��) = False Then
        MsgBox "Ԥ���㱻��ֹ", vbInformation, gstrSysName
        Unload frm�ȴ����ػ���
        Exit Function
    End If
    Unload frm�ȴ����ػ���
    
    '���ؽ�����
    strSQL = "Select Sum(JkaccR) As JkaccR,Sum(JkSocialR) As JkSocialR From " & mstrSavePath & "\SZ" & str��ˮ��
    Set rsTemp = gcn����.Execute(strSQL)
    mcur����֧�� = Val(rsTemp!JkAccR)
    mcurͳ���� = Val(rsTemp!JkSocialR)
    סԺ�������_���� = "�����ʻ�;" & Val(rsTemp!JkAccR) & ";0"
    סԺ�������_���� = סԺ�������_���� & "|ͳ�����;" & Val(rsTemp!JkSocialR) & ";0"
    On Error Resume Next
    gcn����.Execute "Drop Table " & mstrSavePath & "\SZ" & str��ˮ��
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency
    Dim datCurr As Date
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From סԺ���ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� & ",0,0," & _
        "0," & mcurͳ���� & ",0,0," & mcur����֧�� & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    סԺ����_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, sngArrInfo(20) As Single
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, lng����ID As Long
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From סԺ���ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        סԺ�������_���� = False
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        Nvl(rsTemp("�����ʻ�֧��"), 0) * -1 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    סԺ�������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHandle
    '��HIS֮�еĻ������ݽ����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo errHandle
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function

Public Function ���ʴ���_����(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
    Dim rs��ϸ As New ADODB.Recordset, lng��ҳID As Long, rsTemp As New ADODB.Recordset
    Dim str��ˮ�� As String, datCurr As Date, strSQL As String
    Dim strCardNO As String
    Dim int˳��� As Integer
    
    '���ºʹ�λ�����
    Dim dbl�޶� As Double, str�Է��� As String
    
    '�ȶ�ȡ���ղ����й��ڴ�λ�ѵ�����
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=[1] And ������='��λ���޶�'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��λ���޶�", TYPE_����)
    If rsTemp.RecordCount <> 0 Then dbl�޶� = Nvl(rsTemp!����ֵ, 0)
    If dbl�޶� <> 0 Then
        'ȡ��λ���Է���
        gstrSQL = "Select ����ֵ From ���ղ��� Where ����=[1] And ������='��λ���Է���'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��λ���Է���", TYPE_����)
        If rsTemp.RecordCount <> 0 Then str�Է��� = Nvl(rsTemp!����ֵ)
    End If
    
    If str���ݺ� <> "" Then
        gstrSQL = "Select ����id From סԺ���ü�¼ Where NO=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, str���ݺ�)
        lng����ID = rsTemp(0)
    End If
    gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = Nvl(rsTemp(0), 1)
    If str���ݺ� <> "" Then
        gstrSQL = " Select A.* From סԺ���ü�¼ A,�����ʻ� B " & _
                  " Where A.��¼״̬<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 And nvl(A.���ӱ�־,0)<>9 " & _
                  " and A.��¼����=" & int���� & " and A.NO='" & str���ݺ� & "'" & _
                  " and A.����ID=B.����ID And B.����=" & TYPE_���� & _
                  " order by A.����ID,A.����ʱ��,A.��¼����,A.NO,A.���"
    Else
        gstrSQL = "Select * From סԺ���ü�¼ Where Nvl(ʵ�ս��,0)<>0 And ��¼״̬<>0 And nvl(���ӱ�־,0)<>9 and ����id=[1] And ��ҳid=[2] And NVl(�Ƿ��ϴ�,0)=0 order by ����ʱ��,��¼����,NO,���"
    End If
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "", lng����ID, lng��ҳID)
    
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        ���ʴ���_���� = True
        Exit Function
    End If
    
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select A.����,A.����֤��,B.סԺ�� From �����ʻ� A,������Ϣ B Where A.����ID=B.����ID And A.����id=[1] And A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        MsgBox "û���ҵ�������Ϣ��ҽ��ѡ�����", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!����
    '������ˮ��
    str��ˮ�� = rsTemp!סԺ��
    int˳��� = Nvl(rsTemp!����֤��, 0)
    'ҽ���̵����⣬������ٴ�2��ʼ
    If int˳��� = 0 Then int˳��� = 1
    int˳��� = int˳��� + 1
    
    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.���� from (select * from ����֧����Ŀ where ����=[1]) A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", TYPE_����, CLng(rs��ϸ!�շ�ϸĿID))
        If IsNull((rsTemp!��Ŀ����)) Then
            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
            Exit Function
        End If
        rs��ϸ.MoveNext
    Loop
    
    '����DBF�ļ�
    On Error Resume Next
    gcn����.Execute "Drop Table " & mstrSavePath & "\YZ" & str��ˮ��
    
    On Error GoTo errHandle
    gcn����.Execute "Create Table " & mstrSavePath & "\YZ" & str��ˮ�� & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
        "SelfJe N(18,4),Bz1 C(25),Bz2 C(5),Bz3 C(5))"
    strSQL = "Select * From " & mstrSavePath & "\YZ" & str��ˮ��
    rs��ϸ.MoveFirst
    While Not rs��ϸ.EOF
        gstrSQL = "Select A.��Ŀ����,B.ID,B.����,B.���,B.���㵥λ From ����֧����Ŀ A,�շ�ϸĿ B Where B.ID=A.�շ�ϸĿID And A.�շ�ϸĿid=[1] And A.����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID), TYPE_����)            '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
        
        '���š���ˮ�š���š����롢���ơ���Ŀ���롢��񡢼�����λ�����������ۡ����Էѽ��
        If rs��ϸ!�շ���� <> "J" Or dbl�޶� = 0 Then
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','" & int˳��� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� (" & rs��ϸ!ID & ",0,'" & int˳��� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            gcn����.Execute "Insert Into " & mstrSavePath & "\YZ" & str��ˮ�� & " values ('" & strCardNO & "','" & str��ˮ�� & "'," & _
                int˳��� & ",'" & Trim(rsTemp!��Ŀ����) & "','" & Trim(rsTemp!����) & "','','" & Trim(rsTemp!���) & "','" & Trim(rsTemp!���㵥λ) & "'," & _
                Format(rs��ϸ!���� * rs��ϸ!����, "0.####") & "," & Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "0.####") & "," & Format(rs��ϸ!ʵ�ս��, "0.####") & "," & _
                "0,'" & Format(rs��ϸ!����ʱ��, "yyyy-MM-dd") & "','2','3')"
        Else
            '����Ǵ�λ�ѣ��ҳ����޶����Ҫ�������������Է����ϴ�(�ϴ�Ϊ2����ϸ)
            If Val(Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "#0.00")) > Val(Format(dbl�޶�, "#0.00")) Then
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','" & int˳��� & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                gstrSQL = "zl_���˼��ʼ�¼_�ϴ� (" & rs��ϸ!ID & ",0,'" & int˳��� & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
                gcn����.Execute "Insert Into " & mstrSavePath & "\YZ" & str��ˮ�� & " values ('" & strCardNO & "','" & str��ˮ�� & "'," & _
                    int˳��� & ",'" & Trim(rsTemp!��Ŀ����) & "','" & Trim(rsTemp!����) & "','','" & Trim(rsTemp!���) & "','" & Trim(rsTemp!���㵥λ) & "'," & _
                    Format(rs��ϸ!���� * rs��ϸ!����, "0.####") & "," & Format(dbl�޶�, "0.####") & "," & Format(rs��ϸ!���� * rs��ϸ!���� * dbl�޶�, "0.####") & "," & _
                    "0,'" & Format(rs��ϸ!����ʱ��, "yyyy-MM-dd") & "','2','3')"
                int˳��� = int˳��� + 1
                gcn����.Execute "Insert Into " & mstrSavePath & "\YZ" & str��ˮ�� & " values ('" & strCardNO & "','" & str��ˮ�� & "'," & _
                    int˳��� & ",'" & str�Է��� & "','" & Trim(rsTemp!����) & "','','" & Trim(rsTemp!���) & "','" & Trim(rsTemp!���㵥λ) & "'," & _
                    Format(rs��ϸ!���� * rs��ϸ!����, "0.####") & "," & Format((rs��ϸ!ʵ�ս�� - rs��ϸ!���� * rs��ϸ!���� * dbl�޶�) / (rs��ϸ!���� * rs��ϸ!����), "0.####") & "," & Format((rs��ϸ!ʵ�ս�� - rs��ϸ!���� * rs��ϸ!���� * dbl�޶�), "0.####") & "," & _
                    "0,'" & Format(rs��ϸ!����ʱ��, "yyyy-MM-dd") & "','2','3')"
            Else
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','" & int˳��� & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                gstrSQL = "zl_���˼��ʼ�¼_�ϴ� (" & rs��ϸ!ID & ",0,'" & int˳��� & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
                gcn����.Execute "Insert Into " & mstrSavePath & "\YZ" & str��ˮ�� & " values ('" & strCardNO & "','" & str��ˮ�� & "'," & _
                    int˳��� & ",'" & Trim(rsTemp!��Ŀ����) & "','" & Trim(rsTemp!����) & "','','" & Trim(rsTemp!���) & "','" & Trim(rsTemp!���㵥λ) & "'," & _
                    Format(rs��ϸ!���� * rs��ϸ!����, "0.####") & "," & Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "0.####") & "," & Format(rs��ϸ!ʵ�ս��, "0.####") & "," & _
                    "0,'" & Format(rs��ϸ!����ʱ��, "yyyy-MM-dd") & "','2','3')"
            End If
        End If
        
        int˳��� = int˳��� + 1
        rs��ϸ.MoveNext
    Wend
    
    ���ʴ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
