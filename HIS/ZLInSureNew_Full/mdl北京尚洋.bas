Attribute VB_Name = "mdl��������"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)
Public gcn���� As New ADODB.Connection, gint���õ���_���� As Integer
Public gint�Ƿ�ְ�� As Integer
Private mcurͳ���� As Currency, mcur����֧�� As Currency

Public Function ҽ����ʼ��_��������() As Boolean
'���ܣ������Ƿ�������ӵ�ǰ�÷�������
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs�������� As New ADODB.Recordset, str����ֵ As String
'    '��������Ѿ��򿪣��ǾͲ����ٲ���
'    If gcn����.State = adStateOpen Then
'        ҽ����ʼ��_�������� = True
'        Exit Function
'    End If
'
'    On Error GoTo ErrH
'
'    '���ȶ���������������
'    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & TYPE_��������
'    Call OpenRecordset(rsTemp, gstrSysName)
'    Do Until rsTemp.EOF
'        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
'        Select Case rsTemp("������")
'            Case "�û���"
'                strUser = str����ֵ
'            Case "������"
'                strServer = str����ֵ
'            Case "�û�����"
'                strPass = str����ֵ
'            Case "���õ���"
'                gint���õ���_���� = Val(str����ֵ)
'            Case "ͳ������"
'                gstrҽ���������� = str����ֵ
'        End Select
'        rsTemp.MoveNext
'    Loop
'    If strUser = "" Or strServer = "" Or strPass = "" Then
'        MsgBox "�������ò�����,�뵽ҽ��������������������", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    On Error Resume Next
'    If gint���õ���_���� = 1 Then
''        gcn����.ConnectionString = "Provider=Sybase.ASEOLEDBProvider.2;����=" & strPass & ";������ȫ����Ϣ=True;�û� ID=" & strUser & ";����Դ=" & strServer
'        gcn����.ConnectionString = "Provider=MSDASQL.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
'    Else
'        gcn����.ConnectionString = "Provider=MSDAORA.1;Password=" & strPass & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
'    End If
'    gcn����.CursorLocation = adUseClient
'    gcn����.Open
'
'    If Err <> 0 Then
'        MsgBox "����ǰ�÷�������������", vbInformation, gstrSysName
'        ҽ����ʼ��_�������� = False
'        Exit Function
'    End If
    ҽ����ʼ��_�������� = True
'    Exit Function
'ErrH:
'    If ErrCenter() = 1 Then Resume
'    ҽ����ʼ��_�������� = False
End Function


Public Function ����ҽ������() As Boolean
'���ܣ������Ƿ�������ӵ�ǰ�÷�������
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs�������� As New ADODB.Recordset, str����ֵ As String

     
    On Error GoTo ErrH
    
    
'    If MsgBox("�ò����Ƿ�ְ��ҽ��?", vbYesNo + vbQuestion + vbDefaultButton1, "ҽ���ӿ�") = vbYes Then
        
        gint�Ƿ�ְ�� = 1
        
        '���ȶ���������������
        gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������)
        Do Until rsTemp.EOF
            str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Select Case rsTemp("������")
                Case "�û���"
                    strUser = str����ֵ
                Case "������"
                    strServer = str����ֵ
                Case "�û�����"
                    strPass = str����ֵ
                Case "���õ���"
                    gint���õ���_���� = Val(str����ֵ)
                Case "ͳ������"
                    gstrҽ���������� = str����ֵ
            End Select
            rsTemp.MoveNext
        Loop
        If strUser = "" Or strServer = "" Or strPass = "" Then
            MsgBox "�������ò�����,�뵽ҽ��������������������", vbInformation, gstrSysName
            Exit Function
        End If
        
'    Else
'        gint�Ƿ�ְ�� = 0
'
'        strUser = "login_sa"
'        strServer = "sybase"
'        strPass = "passwd"
'        gint���õ���_���� = 1
'        gstrҽ���������� = "1403000000"
'    End If
'
    
    '��������Ѿ��򿪣��ǾͲ����ٲ���
    If gcn����.State = adStateOpen Then
        ����ҽ������ = True
        Exit Function
    End If
    
    
    On Error Resume Next
    If gint���õ���_���� = 1 Then
        gcn����.ConnectionString = "Provider=MSDASQL.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
    Else
        gcn����.ConnectionString = "Provider=MSDAORA.1;Password=" & strPass & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
    End If
    gcn����.CursorLocation = adUseClient
    gcn����.Open
    
    If Err <> 0 Then
        MsgBox "����ǰ�÷�������������", vbInformation, gstrSysName
        ����ҽ������ = False
        Exit Function
    End If
    ����ҽ������ = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    ����ҽ������ = False
End Function

Public Function ҽ������_��������() As Boolean
    ҽ������_�������� = frmSet��������.��������()
End Function

Public Function �������_��������(lng����ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��������)
    '��Ϊ����ȡ�������,��˸����㹻�����,����֧�������ҽ������
    �������_�������� = Nvl(rsTemp!�ʻ����, 1000000)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function ��ݱ�ʶ_��������(Optional bytType As Byte = 0, Optional lng����ID As Long = 0) As String
    '��������ҽ��û�ṩר�ŵ������֤�ӿ�
    Dim strTemp As String
    
    'Τ�����޸���2011-3-2
    If ����ҽ������ = False Then Exit Function
    
    
    strTemp = frmIdentify��������.Identify(bytType, lng����ID)
    Unload frmIdentify��������
    If strTemp = "" Then
        MsgBox "δ��ȡ������Ϣ", vbInformation, gstrSysName
    Else
        ��ݱ�ʶ_�������� = strTemp
    End If
End Function
'
'Public Function �����������_��������(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
''��Ϊ��������δ�ṩԤ����ӿڣ�������õ��Ľ�������Ϊҽ���������ʽ���ݣ����õ�����ʱҽ������ʽ����
'    Dim str��ˮ�� As String, lng����ID As Long, datCurr As Date, strSql As String, strTemp As String
'    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng��� As Long, str���� As String
'    Dim strCardNO As String, str�վ���Ŀ As String, str������Ŀ As String, str�����Ŀ As String
''    ����ID         adBigInt, 19, adFldIsNullable
''    �շ����       adVarChar, 2, adFldIsNullable
''    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
''    ���㵥λ       adVarChar, 6, adFldIsNullable
''    ������         adVarChar, 20, adFldIsNullable
''    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
''    ����           adSingle, 15, adFldIsNullable
''    ����           adSingle, 15, adFldIsNullable
''    ʵ�ս��       adSingle, 15, adFldIsNullable
''    ͳ����       adSingle, 15, adFldIsNullable
''    ����֧������ID adBigInt, 19, adFldIsNullable
''    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
''    ժҪ           adVarChar, 200, adFldIsNullable
''    �Ƿ���       adBigInt, 19, adFldIsNullable
''    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'    On Error GoTo errHandle
'    If rs��ϸ.RecordCount = 0 Then
'        MsgBox "û�в��˷��ã����ܽ���", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    datCurr = zlDatabase.Currentdate
'    lng����ID = rs��ϸ(0)
'    gstrSQL = "Select ���� From �����ʻ� Where ����id=" & lng����ID & " And ����=" & TYPE_��������
'    Call OpenRecordset(rsTemp, gstrSysName)
'    If rsTemp.EOF Then
'        MsgBox "û���ҵ�������Ϣ��ҽ��ѡ�����", vbInformation, gstrSysName
'        Exit Function
'    End If
'    strCardNO = rsTemp!����
'    '������ˮ��
'    str��ˮ�� = toHex(Format(datCurr, "YYMMDDHHMMSS") & Format(lng����ID, "0######"), 35)
'
'    '�ж��Ƿ���ҽ������δ��Ӧ
'    Do Until rs��ϸ.EOF
'        gstrSQL = "select A.��Ŀ����,B.����,B.˵�� from (select * from ����֧����Ŀ where ����=" & TYPE_�������� & ") A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = " & rs��ϸ!�շ�ϸĿID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        If IsNull(rsTemp!��Ŀ����) Then
'            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
'            Exit Function
'        End If
'        If IsNull(rsTemp!˵��) Then
'            MsgBox "����ȷ����Ŀ<" & rsTemp!���� & ">���վ���Ŀ���Ϳ��Һ������", vbInformation, gstrSysName
'            Exit Function
'        ElseIf Len(rsTemp!˵��) < 2 Then
'            MsgBox "����ȷ����Ŀ<" & rsTemp!���� & ">�Ŀ��Һ������", vbInformation, gstrSysName
'            Exit Function
'        End If
'        strTemp = rsTemp!��Ŀ����
'        strSql = "Select * From PARA_CAPTURE_ITEM Where Areaid='" & gstrҽ���������� & "' And Item_Code='" & UCase(rsTemp!��Ŀ����) & "'"
'        Set rsTemp = gcn����.Execute(strSql)
'        If rsTemp.EOF Then
'            MsgBox "���м�������δ�ҵ�����Ϊ[" & UCase(strTemp) & "]����Ŀ����˲�", vbInformation, gstrSysName
'            Exit Function
'        End If
'        rs��ϸ.MoveNext
'    Loop
'
'    '����DBF�ļ�
'    lng��� = 1
'    rs��ϸ.MoveFirst
'    While Not rs��ϸ.EOF
'        gstrSQL = "Select * From �շ�ϸĿ Where ID=" & rs��ϸ!�շ�ϸĿID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        strTemp = rsTemp!˵��      '˵������Ϊ�գ����е�һλ����վ���Ŀ��𣬵ڶ�λ��ſ��Һ������
'        str�վ���Ŀ = Left(strTemp, 1)
'        str������� = Mid(strTemp, 2, 1)
'        If rsTemp!��� = 5 Or rsTemp!��� = 6 Or rsTemp!��� = 7 Then
'            str������ = "A"       'ҩƷ
'        Else
'            str������ = "B"       'ҽ��
'        End If
'
'        gstrSQL = "Select ��Ŀ���� From ����֧����Ŀ Where ����=" & TYPE_�������� & " And �շ�ϸĿid=" & rs��ϸ!�շ�ϸĿID
'        Call OpenRecordset(rsTemp, gstrSysName)             '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
'
'        strSql = "Select * From PARA_CAPTURE_ITEM Where Areaid='" & gstrҽ���������� & "' And Item_Code='" & UCase(rsTemp!��Ŀ����) & "'"
'        Set rsTemp = gcn����.Execute(strSql)
'        '���š���ˮ�š���š����롢���ơ���Ŀ���롢��񡢼�����λ�����������ۡ����Էѽ��
''    VISIT_NUMBER                char(18)        not null,   //��������
''    ITEM_NO                     numeric(6, 0)   not null,   //ͬһ��������Ŀ���
''    ITEM_CLASS                  char(1)         not null,   //�շ���Ŀ���:��A��ҩ
''    ITEM_CODE                   char(12)        not null,   //��Ŀ����
''    ITEM_NAME                   char(40)        not null,   //��Ŀ����
''    SPEC                        varchar(50)     not null,   //���
''    PRICE_UNIT                  char(8)         not null,   //�Ƽ۵�λ
''    PRICE                       numeric(9, 4)   not null,   //����
''    QUANTITY                    numeric(6, 2)   not null,   //����
''    COST                        numeric(8, 2)   not null,   //���
''    RECEIPT_CLASS               char(1)         not null,   //�վ���Ŀ����
''    COLLATE_RELATION            char(12)        null,       //��ҽ�����Ķ�Ӧ��ϵ
''    OPERATOR                    char(15)        null,       //������
''    OPERATE_TIME                datetime        null,       //��������
''    CLINIC_FLAG                 numeric(1, 0)   not null,   //����/סԺ��־
''    EXE_DEPT                    char(20)        null,       //ִ�п���
''    APP_DOCTOR                  char(30)        null,       //����ҽ��
''    APP_DEPT                    char(20)        null,       //��������
''    TAKE_MEDICINE_FLAG          char(8)         not null,   //��Ժ��ҩ��־
''    ITEM_NO_DEPT_STAT           char(2)         null,       //���Һ�����Ŀ���
''    ITEM_NO_ACCOUNTANT_ITEM char(2)         null,       //��ƺ�����Ŀ���
''    constraint PK_SICK_PRICE_ITEM PRIMARY KEY CLUSTERED (VISIT_NUMBER, ITEM_NO)
''A ��ҩ��B ��ҩ��C ��ҩ��D ���ƣ�E ��飬F ���䣬G ���飬H �����ѣ�I ��Ѫ�ѣ�J �����ѣ�K CT��ECT��L ������M B����N �ĵ�ͼ��O �Ե�ͼ��P θ����Q ��
'
'        gcn����.Execute "Insert Into SICK_PRICE_ITEM values ('" & str��ˮ�� & "'," & lng��� & ",'" & _
'            Trim(rsTemp!ITEM_TYPE) & "','" & Trim(rsTemp!ITEM_CODE) & "','" & ToVarchar(Trim(rsTemp!ITEM_NAME), 40) & "','" & _
'            ToVarchar(Trim(rsTemp!ITEM_SPEC), 50) & "','" & ToVarchar(Trim(rsTemp!PRICE_UNIT), 8) & "','" & _
'            Trim(rsTemp!CUnit) & "'," & rs��ϸ!���� & "," & rs��ϸ!���� & "," & rs��ϸ!ʵ�ս�� & ",'" & _
'            str�վ���Ŀ & "','" & trim(rstemp!ITEM_CODE) & "','" & userinfo.���� & "',to_Date('" & _
'            format(zldatabase.Currentdate,"yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),0,'" & _
'
'        lng��� = lng��� + 1
'        rs��ϸ.MoveNext
'    Wend
'    On Error GoTo errHandle
'
'    '�ȴ����ؽ�������
'    If frm�ȴ����ر�������.waitReturn(mstrSavePath & "\SM" & str��ˮ��) = False Then
'        MsgBox "Ԥ���㱻��ֹ", vbInformation, gstrSysName
'        Unload frm�ȴ����ر�������
'        Exit Function
'    End If
'    Unload frm�ȴ����ر�������
'
'    '���ؽ�����
'    strSql = "Select * From " & mstrSavePath & "\SM" & str��ˮ��
'    Set rsTemp = gcn����.Execute(strSql)
'    mcur����֧�� = Val(rsTemp!JkAccR)
'    mcurͳ���� = Val(rsTemp!JkSocialR)
'    str���㷽ʽ = "�����ʻ�;" & Val(rsTemp!JkAccR) & ";0"
'    str���㷽ʽ = str���㷽ʽ & "|ͳ�����;" & Val(rsTemp!JkSocialR) & ";0"
'    �����������_�������� = True
'    Exit Function
'
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Public Function �������_��������(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency, Optional ByRef strAdvance As String) As Boolean
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency
    Dim str��ˮ�� As String, lng����ID As Long, datCurr As Date, strSQL As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset, lng��� As Long, strִ�в��� As String, str�������� As String
    Dim strCardNO As String, str�վ���Ŀ As String, str������� As String, str������ As String
    Dim str��Ժ��ҩ As String, cur����ͳ�� As Currency, cur��ͳ�� As Currency, rs��ϸ As New ADODB.Recordset
    Dim cur����Ա���� As Currency, cur����ҽ�� As Currency, str���㷽ʽ As String, lng���� As Integer
    Dim strTempID As String
    Dim blnOld As Boolean
    Dim strItemType As String, strItemCode As String, strItemName As String, strItemSpec As String, strPriceUnit As String
    Dim rsPati As New ADODB.Recordset
    Dim Guanwei As String
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select * From ���˷��ü�¼ Where ��¼״̬<>0 And Nvl(ʵ�ս��,0)<>0 and Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 And ����id=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ã����ܽ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ!����ID
    gstrSQL = "Select ���� From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��������)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "û���ҵ�������Ϣ��ҽ��ѡ�����"
        Exit Function
    End If
    strCardNO = rsTemp!����
    '������ˮ��
'    str��ˮ�� = toHex(Format(datCurr, "YYMMDDHHMMSS") & Format(lng����ID, "0######"), 35)
    str��ˮ�� = Format(zlDatabase.Currentdate(), "yyyy") & rs��ϸ!NO
    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.����,B.˵��,B.��� from (select * from ����֧����Ŀ where ����=[1]) A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������, CLng(rs��ϸ!�շ�ϸĿID))
'        If IsNull(rsTemp!��Ŀ����) Then
'            MsgBox "<" & rsTemp!���� & ">δ��Ӧҽ������,���Ƚ��ж���", vbInformation, gstrSysName
'            Exit Function
'        End If
        If rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7" Then
        
        Else
            If IsNull(rsTemp!˵��) Then
                Err.Raise 9000, gstrSysName, "����ȷ����Ŀ<" & rsTemp!���� & ">���վ���Ŀ���Ϳ��Һ������"
                Exit Function
            ElseIf Len(rsTemp!˵��) < 2 Then
                Err.Raise 9000, gstrSysName, "����ȷ����Ŀ<" & rsTemp!���� & ">�Ŀ��Һ������"
                Exit Function
            End If
        End If
'        strTemp = rsTemp!��Ŀ����
'        strSql = "Select * From PARA_CAPTURE_ITEM Where AREAID='" & gstrҽ���������� & "' And ITEM_CODE='" & UCase(rsTemp!��Ŀ����) & "'"
'        Set rsTemp = gcn����.Execute(strSql)
'        If rsTemp.EOF Then
'            MsgBox "���м�������δ�ҵ�����Ϊ[" & UCase(strTemp) & "]����Ŀ����˲�", vbInformation, gstrSysName
'            Exit Function
'        End If
        rs��ϸ.MoveNext
    Loop
    
    '��������ϸ
    lng��� = 1
    rs��ϸ.MoveFirst
    While Not rs��ϸ.EOF
        gstrSQL = "Select * From ���ű� Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!ִ�в���id))
        strִ�в��� = rsTemp!����
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!��������ID))
        str�������� = rsTemp!����
        
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        If IsNull(rsTemp!˵��) Then
            If rsTemp!��� = "5" Then
                strTemp = "AA"
            ElseIf rsTemp!��� = "6" Then
                strTemp = "BB"
            ElseIf rsTemp!��� = "7" Then
                strTemp = "CC"
            End If
        Else
            strTemp = rsTemp!˵��      '˵������Ϊ�գ����е�һλ����վ���Ŀ��𣬵ڶ�λ��ſ��Һ������
        End If
        str�վ���Ŀ = Left(strTemp, 1)
        str������� = Mid(strTemp, 2, 1)
        If rsTemp!��� = 5 Or rsTemp!��� = 6 Or rsTemp!��� = 7 Then
            str������ = "A"       'ҩƷ
        Else
            str������ = "B"       'ҽ��
        End If
        Select Case rsTemp!���
            Case "1"
                strItemType = "O"
            Case "4"
                strItemType = "U"
            Case "5"
                strItemType = "A"
            Case "6"
                strItemType = "C"
            Case "7"
                strItemType = "C"
            Case "C"
                strItemType = "D"
            Case "D"
                strItemType = "E"
            Case "E", "L"
                strItemType = "F"
            Case "F"
                strItemType = "G"
            Case "G"
                strItemType = "H"
            Case "H"
                strItemType = "I"
            Case "I", "Z"
                strItemType = "Z"
            Case "J"
                strItemType = "J"
            Case "K"
                strItemType = "L"
            Case "M"
                strItemType = "K"
        End Select
        str��Ժ��ҩ = "��Ժ��ҩ"           '��Ժ��ҩ��־��ȡֵ����������д�ѯ�ʣ�
        
        If gint�Ƿ�ְ�� = 0 Then
        gstrSQL = "Select ��Ŀ����,�շ�ϸĿID From ����֧����Ŀ Where ����=[1] And �շ�ϸĿid=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������, CLng(rs��ϸ!�շ�ϸĿID))           '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
        If rsTemp.EOF Then
            strTempID = ""
        ElseIf IsNull(rsTemp!��Ŀ����) Then
            strTempID = ""
        Else
            strTempID = rsTemp!��Ŀ����
        End If
         
        
        '�������޸ģ��շ������ǰȡ��C�����Ƹ�Ϊȡ������
        Else
        gstrSQL = "Select a.��Ŀ����,a.�շ�ϸĿID,substr(b.��ע,3,1) as ����,c.���� as �շ���� From ����֧������ C,����֧����Ŀ a ,������Ŀ b" & _
          " where a.����ID=c.id and a.����=b.���� and a.��Ŀ����=b.���� and a.����=[1] And a.�շ�ϸĿid=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������, CLng(rs��ϸ!�շ�ϸĿID))           '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
        If rsTemp.EOF Then
            strTempID = ""
            'str���� = ""
        Else
            strTempID = Nvl(rsTemp!��Ŀ����)
            'str���� = Nvl(rsTemp!����)
            str�վ���Ŀ = Nvl(rsTemp!�շ����)
        End If
        End If
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        strItemName = rsTemp!����
        strItemCode = rsTemp!����
        strPriceUnit = Nvl(rsTemp!���㵥λ)
        strItemSpec = Nvl(rsTemp!���)
        
        
        
        'Τ�����޸���2011-3-2
        '����ҽ��������ϸд��ʹ��ԭ��ʽ��ְ��ҽ��������ϸд���µ��м��
        If gint�Ƿ�ְ�� = 0 Then
            '���м��д������,סԺ/�����־�д�ѯ��
            gcn����.Execute "Insert Into SICK_PRICE_ITEM " & _
                " (VISIT_NUMBER,ITEM_NO,ITEM_CLASS,ITEM_CODE,ITEM_NAME,SPEC,PRICE_UNIT,PRICE,QUANTITY, " & _
                "  COST,RECEIPT_CLASS,COLLATE_RELATION,OPERATOR,OPERATE_TIME,CLINIC_FLAG,EXE_DEPT,APP_DOCTOR, " & _
                "  APP_DEPT,TAKE_MEDICINE_FLAG,ITEM_NO_DEPT_STAT,ITEM_NO_ACCOUNTANT_ITEM)" & _
                " values ('" & str��ˮ�� & "'," & lng��� & ",'" & _
                strItemType & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "','" & _
                ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "'," & _
                Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "0.####") & "," & Format(rs��ϸ!���� * rs��ϸ!����, "0.####") & "," & Format(rs��ϸ!ʵ�ս��, "0.####") & ",'" & _
                str�վ���Ŀ & "','" & strTempID & "','" & rs��ϸ!����Ա���� & "','" & _
                Format(rs��ϸ!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "',1,'" & _
                strִ�в��� & "','" & rs��ϸ!������ & "','" & str�������� & "','" & str��Ժ��ҩ & "','" & _
                str������� & "','" & str������ & "')"
                
        Else
            'ְ��ҽ����ϸд���µ��м��strItemType
        
           gcn����.Execute "Insert Into KC28 " & _
                        " (AKB020,AKC220,CKC158,AAE011,AAE036,AKA063,AKC222,AKC223,AKC227,CKC197,CKC198," & _
                        "  CKC159,CKC160,AKA070,CKC161,CKC169,CKC170,CKC171,CKE081,CKE085,CKE086,CKE090)" & _
                        " values ('" & gstrҽ���������� & "','" & str��ˮ�� & "','" & lng��� & "','" & rs��ϸ!����Ա���� & "'," & _
                        " to_date('" & rs��ϸ!�Ǽ�ʱ�� & "','yyyy-MM-DD hh24:MI:SS'),'" & str�վ���Ŀ & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "'," & _
                        Format(Format(rs��ϸ!��׼����, "0.####") * (rs��ϸ!���� * rs��ϸ!����), "0.####") & "," & Format(rs��ϸ!��׼����, "0.####") & "," & Format(rs��ϸ!���� * rs��ϸ!����, "0.####") & ",'" & _
                        ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "','/','" & strTempID & "','" & _
                        strִ�в��� & "','" & rs��ϸ!������ & "','" & str�������� & "','" & str�վ���Ŀ & "','" & str������� & "','" & str������ & "','" & str��Ժ��ҩ & "')"
        
         
        End If

        lng��� = lng��� + 1
        rs��ϸ.MoveNext
    Wend
    On Error GoTo errHandle
    
    '�ȴ����ؽ�������
    strTemp = frm�ȴ����ر�������.waitReturn(str��ˮ��, 0)
    If strTemp = "" Then
        Err.Raise 9000, gstrSysName, "������̱���ֹ"
        gcn����.Execute "Delete From SICK_PRICE_ITEM Where VISIT_NUMBER='" & str��ˮ�� & "'"
        Unload frm�ȴ����ر�������
        Exit Function
    End If
    Unload frm�ȴ����ر�������
    
    '���ؽ�����
    strSQL = "Select * From MED_RECEIPT_RECORD_MASTER Where CHARGE_NUMBER='" & strTemp & "'"
    Set rsTemp = gcn����.Execute(strSQL)
    If IsDate(rsTemp!BIRTH_DATE) Then
        lng���� = Int(zlDatabase.Currentdate() - CDate(rsTemp!BIRTH_DATE)) / 365
    End If
    
    '�ȸ��²�����Ϣ
    #If gverControl < 6 Then
        gstrSQL = "Select * From ������Ϣ A Where A.����ID =[1]"
    #Else
        gstrSQL = "Select A.����id, A.�����, A.סԺ��, A.���￨��, A.����֤��, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.����, A.�Ա�, A.����, A.��������, A.�����ص�, A.���֤��, A.����֤��, A.���, A.ְҵ, A.����, A.����, A.����, A.ѧ��, A.����״��, A.��ͥ��ַ," & vbNewLine & _
            "      A.��ͥ�绰, A.��ͥ��ַ�ʱ� As �����ʱ�, A.�໤��, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ, A.��ϵ�˵绰, A.��ͬ��λid, A.������λ, A.��λ�绰, A.��λ�ʱ�, A.��λ������, A.��λ�ʺ�, A.������, A.������, A.��������, A.����ʱ��, A.����״̬," & vbNewLine & _
            "      A.��������, A.סԺ����, A.��ǰ����id, A.��ǰ����id, A.��ǰ����, A.��Ժʱ��, A.��Ժʱ��, A.��Ժ, A.Ic����, A.������, A.ҽ����, A.����, A.��ѯ����, A.�Ǽ�ʱ��, A.ͣ��ʱ��, A.����" & vbNewLine & _
            "From ������Ϣ A Where A.����ID =[1]"
    #End If
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lng����ID)
    gstrSQL = "zl_������Ϣ_Update(" & _
        lng����ID & "," & IIf(IsNull(rsPati!�����), "NULL", rsPati!�����) & "," & _
        IIf(IsNull(rsPati!סԺ��), "NULL", rsPati!סԺ��) & ",'" & IIf(IsNull(rsPati!�ѱ�), "", rsPati!�ѱ�) & "'," & _
        "'" & IIf(IsNull(rsPati!ҽ�Ƹ��ʽ), "", rsPati!ҽ�Ƹ��ʽ) & "'," & _
        "'" & rsTemp!Name & "','" & IIf(Nvl(rsTemp!Sex, "0") = "1", "��", "Ů") & "'," & lng���� & "," & _
        "     To_Date('" & Format(rsTemp!BIRTH_DATE, "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
        "'" & IIf(IsNull(rsPati!�����ص�), "", rsPati!�����ص�) & "','" & rsTemp!PERSONAL_NUMBER & "'," & _
        "'" & IIf(IsNull(rsPati!���), "", rsPati!���) & "','" & IIf(IsNull(rsPati!ְҵ), "", rsPati!ְҵ) & "'," & _
        "'" & IIf(IsNull(rsPati!����), "", rsPati!����) & "','" & IIf(IsNull(rsPati!����), "", rsPati!����) & "'," & _
        "'" & IIf(IsNull(rsPati!ѧ��), "", rsPati!ѧ��) & "','" & IIf(IsNull(rsPati!����״��), "", rsPati!����״��) & "'," & _
        "'" & IIf(IsNull(rsPati!��ͥ��ַ), "", rsPati!��ͥ��ַ) & "','" & IIf(IsNull(rsPati!��ͥ�绰), "", rsPati!��ͥ�绰) & "'," & _
        "'" & IIf(IsNull(rsPati!�����ʱ�), "", rsPati!�����ʱ�) & "','" & IIf(IsNull(rsPati!��ϵ������), "", rsPati!��ϵ������) & "'," & _
        "'" & IIf(IsNull(rsPati!��ϵ�˹�ϵ), "", rsPati!��ϵ�˹�ϵ) & "','" & IIf(IsNull(rsPati!��ϵ�˵�ַ), "", rsPati!��ϵ�˵�ַ) & "'," & _
        "'" & IIf(IsNull(rsPati!��ϵ�˵绰), "", rsPati!��ϵ�˵绰) & "'," & IIf(IsNull(rsPati!��ͬ��λID), "NULL", rsPati!��ͬ��λID) & "," & _
        "'" & Nvl(rsPati!������λ) & "','" & IIf(IsNull(rsPati!��λ�绰), "", rsPati!��λ�绰) & "'," & _
        "'" & IIf(IsNull(rsPati!��λ�ʱ�), "", rsPati!��λ�ʱ�) & "','" & IIf(IsNull(rsPati!��λ������), "", rsPati!��λ������) & "'," & _
        "'" & IIf(IsNull(rsPati!��λ�ʺ�), "", rsPati!��λ�ʺ�) & "','" & IIf(IsNull(rsPati!������), "", rsPati!������) & "'," & _
        " " & IIf(IsNull(rsPati!������), "NULL", rsPati!������) & "," & TYPE_�������� & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
'        MsgBox "�޸Ĳ�����Ϣ��" & gstrSQL
    
    '��ȡ����ҽ�����㷽ʽ��֧�����
   
     cur�����ʻ� = rsTemp!PAY_SIDE2
    cur����ͳ�� = rsTemp!PAY_SIDE3
    cur��ͳ�� = rsTemp!PAY_SIDE4
    cur����ҽ�� = rsTemp!PAY_SIDE5
    cur����Ա���� = rsTemp!PAY_SIDE6
    'д������
  
    If cur�����ʻ� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    End If
    If cur����ͳ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||��������|" & cur����ͳ��
    End If
    If cur��ͳ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||�󲡻���|" & cur��ͳ��
    End If
    If cur����ҽ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||�������|" & cur����ҽ��
    End If
    If cur����Ա���� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
    End If
    
    '�������
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        #If gverControl < 2 Then
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
        #Else
            strAdvance = str���㷽ʽ
            gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���½�����")
    Else
        str���㷽ʽ = "�����ʻ�|0"
        strAdvance = str���㷽ʽ
        gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
    End If
    #If gverControl < 2 Then
        blnOld = True
        frm������Ϣ.ShowME (lng����ID)
    #End If
    
    gstrSQL = "Select ����ID,���ʽ�� From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�������� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�������� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� & ",0,0," & _
        "0," & mcurͳ���� & ",0,0," & mcur����֧�� & ",Null,Null,Null,Null" & IIf(blnOld, "", ",1") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    �������_�������� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_��������(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
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
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        ����������_�������� = False
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�������� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�������� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        Nvl(rsTemp("�����ʻ�֧��"), 0) * -1 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    ����������_�������� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_��������(rsDetail As ADODB.Recordset, lng����ID As Long, strҽ���� As String) As String
'��Ϊ��������δ�ṩԤ����ӿڣ�������õ��Ľ�������Ϊҽ���������ʽ���ݣ����õ�����ʱҽ������ʽ����
    Dim str��ˮ�� As String, datCurr As Date, strSQL As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset, lng��� As Long, strִ�в��� As String, str�������� As String
    Dim strCardNO As String, str�վ���Ŀ As String, str������� As String, str������ As String
    Dim str��Ժ��ҩ As String, cur����ͳ�� As Currency, cur��ͳ�� As Currency, rs��ϸ As New ADODB.Recordset
    Dim cur����Ա���� As Currency, cur����ҽ�� As Currency, str���㷽ʽ As String, cur�����ʻ� As Currency
    Dim strTempID As String
    Dim strItemType As String, strItemCode As String, strItemName As String, strItemSpec As String, strPriceUnit As String
    Dim strסԺ�� As String, str���� As String
    
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select Max(��ҳID) From סԺ���ü�¼ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    gstrSQL = " Select * From ���˷��ü�¼" & _
              " Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 And Nvl(ʵ�ս��,0)<>0 and nvl(����,0)*nvl(����,0)<>0 " & _
              " And ����id=[1] And ��ҳid=" & rsTemp(0)
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "", lng����ID)
    If rs��ϸ.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "û�в��˷��ã����ܽ���"
        Exit Function
    End If
    
    lng����ID = rs��ϸ!����ID
    gstrSQL = "Select ����,ҽ���� From �����ʻ� Where ����id=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��������)
    If rsTemp.EOF Then
        MsgBox "û���ҵ�������Ϣ��ҽ��ѡ�����", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!����
    strҽ���� = Nvl(rsTemp!ҽ����)
    
     '2011-06-23 ��ȡסԺ��С������
        gstrSQL = "Select * From ������Ϣ Where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
        strסԺ�� = Nvl(rsTemp!סԺ��)
    
    '������ˮ��
    str��ˮ�� = ToVarchar(Format(datCurr, "YYMMDDHHMMSS") & lng����ID, 18)
    
    '�ж��Ƿ���ҽ������δ��Ӧ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.��Ŀ����,B.����,B.˵��,B.����,B.��� from (select * from ����֧����Ŀ where ����=[1]) A, �շ�ϸĿ B where A.�շ�ϸĿid(+)=B.id and B.id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������, CLng(rs��ϸ!�շ�ϸĿID))
        If rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7" Then
        
        Else
            If IsNull(rsTemp!˵��) Then
                Err.Raise 9000, gstrSysName, "����ȷ����Ŀ<" & rsTemp!���� & ">���վ���Ŀ���Ϳ��Һ������"
                Exit Function
            ElseIf Len(rsTemp!˵��) < 2 Then
                Err.Raise 9000, gstrSysName, "����ȷ����Ŀ<" & rsTemp!���� & ">�Ŀ��Һ������"
                Exit Function
            End If
        End If
        rs��ϸ.MoveNext
    Loop
    
    '����Ƿ��ѽ�������
    If ����ҽ������ = False Then Exit Function
    
    '��������ϸ
    lng��� = 1
    rs��ϸ.MoveFirst
    While Not rs��ϸ.EOF
        
        gstrSQL = "Select * From ���ű� Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!ִ�в���id))
        strִ�в��� = rsTemp!����
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!��������ID))
        str�������� = rsTemp!����
        
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        If IsNull(rsTemp!˵��) Then
            If rsTemp!��� = "5" Then
                strTemp = "AA"
            ElseIf rsTemp!��� = "6" Then
                strTemp = "BB"
            ElseIf rsTemp!��� = "7" Then
                strTemp = "CC"
            End If
        Else
            strTemp = rsTemp!˵��      '˵������Ϊ�գ����е�һλ����վ���Ŀ��𣬵ڶ�λ��ſ��Һ������
        End If
        str�վ���Ŀ = Left(strTemp, 1)
        str������� = Mid(strTemp, 2, 1)
        
        If rsTemp!��� = 5 Or rsTemp!��� = 6 Or rsTemp!��� = 7 Then
            str������ = "A"       'ҩƷ
        Else
            str������ = "B"       'ҽ��
        End If
        Select Case rsTemp!���
            Case "1"
                strItemType = "O"
            Case "4"
                strItemType = "U"
            Case "5"
                strItemType = "A"
            Case "6"
                strItemType = "C"
            Case "7"
                strItemType = "C"
            Case "C"
                strItemType = "D"
            Case "D"
                strItemType = "E"
            Case "E", "L"
                strItemType = "F"
            Case "F"
                strItemType = "G"
            Case "G"
                strItemType = "H"
            Case "H"
                strItemType = "I"
            Case "J"
                strItemType = "J"
            Case "K"
                strItemType = "L"
            Case "M"
                strItemType = "K"
            Case Else
                strItemType = "Z"
        End Select
        gstrSQL = "Select ���� From ҩƷ�շ���¼ Where ����ID=[1] And NO=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!ID), CStr(rs��ϸ!NO))
        If rsTemp.EOF Then
            str��Ժ��ҩ = "��Ժ��ҩ"
        Else
            If Mid(CStr(Nvl(rsTemp(0), 0)), 2, 1) = "3" Then
                str��Ժ��ҩ = "��Ժ��ҩ"
            Else
                str��Ժ��ҩ = "��Ժ��ҩ"
            End If
        End If
        
        If gint�Ƿ�ְ�� = 0 Then
            gstrSQL = "Select ��Ŀ����,�շ�ϸĿID From ����֧����Ŀ Where ����=[1] And �շ�ϸĿid=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������, CLng(rs��ϸ!�շ�ϸĿID))          '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
            If rsTemp.EOF Then
                strTempID = ""
            ElseIf IsNull(rsTemp!��Ŀ����) Then
                strTempID = ""
            Else
                strTempID = rsTemp!��Ŀ����
            End If
            '�������޸ģ��շ������ǰȡ��C�����Ƹ�Ϊȡ������
        Else
            gstrSQL = "Select a.��Ŀ����,a.�շ�ϸĿID,substr(b.��ע,3,1) as ����,c.���� as �շ���� From ����֧������ C,����֧����Ŀ a ,������Ŀ b" & _
              " where a.����ID=c.id and a.����=b.���� and a.��Ŀ����=b.���� and a.����=[1] And a.�շ�ϸĿid=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������, CLng(rs��ϸ!�շ�ϸĿID))            '��Ϊ֮ǰ������Ƿ���ж��룬���Զ����ļ�¼һ�������
            If rsTemp.EOF Then
                strTempID = ""
                str���� = ""
            Else
                strTempID = Nvl(rsTemp!��Ŀ����)
                str���� = Nvl(rsTemp!����)
                str�վ���Ŀ = Nvl(rsTemp!�շ����)
            End If
        End If
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        strItemName = rsTemp!����
        strItemCode = rsTemp!����
        strPriceUnit = Nvl(rsTemp!���㵥λ)
        strItemSpec = Nvl(rsTemp!���)
        '---------------Τ�����޸���2011-3-2----------------------------
        '����ҽ��������ϸд��ʹ��ԭ��ʽ��ְ��ҽ��������ϸд���µ��м��
        '�������޸ģ�����BAL_DATE is null�жϣ���Ϊ���סԺ��סԺ�Ų�ͬ
        
        If gint�Ƿ�ְ�� = 0 Then
            
            '���м��д������,סԺ/�����־�д�ѯ��
            gcn����.Execute "Insert Into SICK_PRICE_ITEM " & _
                " (VISIT_NUMBER,ITEM_NO,ITEM_CLASS,ITEM_CODE,ITEM_NAME,SPEC,PRICE_UNIT,PRICE,QUANTITY, " & _
                "  COST,RECEIPT_CLASS,COLLATE_RELATION,OPERATOR,OPERATE_TIME,CLINIC_FLAG,EXE_DEPT,APP_DOCTOR, " & _
                "  APP_DEPT,TAKE_MEDICINE_FLAG,ITEM_NO_DEPT_STAT,ITEM_NO_ACCOUNTANT_ITEM)" & _
                " values ('" & str��ˮ�� & "'," & lng��� & ",'" & _
                strItemType & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "','" & _
                ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "'," & _
                Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "0.####") & "," & Format(rs��ϸ!���� * rs��ϸ!����, "0.####") & "," & Format(rs��ϸ!ʵ�ս��, "0.####") & ",'" & _
                str�վ���Ŀ & "','" & strTempID & "','" & rs��ϸ!����Ա���� & "','" & _
                Format(rs��ϸ!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "',1,'" & _
                strִ�в��� & "','" & rs��ϸ!������ & "','" & str�������� & "','" & str��Ժ��ҩ & "','" & _
                str������� & "','" & str������ & "')"
                
        Else
            'ְ��ҽ����ϸд���µ��м��,
            '�������޸ģ�1����AKAO63ȡ��strItemType��Ϊ��str�վ���,2������ҽ�����Ƽ��㵥�۳������Ƿ���ںϼ������޸���ͨ�����۱���4λС���������ý���ҽ�� & Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "0.####") &
           ' Set rsTemp = gcn����.Execute("Select * From SICK_VISIT_INFO Where PERSONAL_NUMBER='" & strҽ���� & "'  and BAL_DATE is null And HOSPITAL_NUMBER='" & gstrҽԺ���� & "'")
          'xiaofan �޸ĸ���סԺ����ȡҽ��������Ϣ
          Set rsTemp = gcn����.Execute("Select * From SICK_VISIT_INFO Where  HOSPITAL_NUMBER='" & gstrҽԺ���� & "' and RESIDENCE_NO='" & strסԺ�� & "'")
        
            If rsTemp.EOF Then
                MsgBox "��ȷ�ϸò����Ƿ�����ҽ��ϵͳ������Ժ�Ǽ�!", vbInformation, "ҽ���ӿ�"
                Exit Function
            End If
            
            strסԺ�� = Nvl(rsTemp!RESIDENCE_NO)
        
            gcn����.Execute "Insert Into KC27 " & _
                            " (AKB020,CKC179,AKC190,AKC220,CKC158,AAE011,AAE036,AKA063,AKC222,AKC223,AKC227,CKC197,CKC198," & _
                            "  CKC159,CKC160,AKA070,CKC161,CKC169,CKC170,CKC171,CKE081,CKE085,CKE086,CKE090)" & _
                            " values ('" & gstrҽ���������� & "','" & strסԺ�� & "','" & Null & "','" & str��ˮ�� & "'," & lng��� & ",'" & rs��ϸ!����Ա���� & "'," & _
                            " to_date('" & rs��ϸ!�Ǽ�ʱ�� & "','yyyy-MM-DD hh24:MI:SS'),'" & str�վ���Ŀ & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "'," & _
                            Format(Format(rs��ϸ!��׼����, "0.####") * (rs��ϸ!���� * rs��ϸ!����), "0.####") & "," & Format(rs��ϸ!��׼����, "0.####") & "," & Format(rs��ϸ!���� * rs��ϸ!����, "0.####") & ",'" & _
                            ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "','" & str���� & "','" & strTempID & "','" & _
                            strִ�в��� & "','" & rs��ϸ!������ & "','" & str�������� & "','" & str�վ���Ŀ & "','" & str������� & "','" & str������ & "','" & str��Ժ��ҩ & "')"
            
            
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            
        End If
        
        
        lng��� = lng��� + 1
        
        
        rs��ϸ.MoveNext
    Wend
    On Error GoTo errHandle
    
    
    
    If gint�Ƿ�ְ�� = 1 Then
        'ְ��ҽ����סԺ����ȡҽ�����㷵��ֵ
        str��ˮ�� = strסԺ��
    End If
    
    
    
    '---------------2011-3-2�޸Ĳ���---------------------------------VISIT_NUMBER
    
    '�ȴ����ؽ�������
    '�������޸ģ�����SICK_PRICE_ITEM��VISIT_NUMBER�ֶ�Ϊ�գ��ش�ȡ���������͸ĳ�RESIDENCE_NO
    Screen.MousePointer = 0
    strTemp = frm�ȴ����ر�������.waitReturn(str��ˮ��, 1)
    If gint�Ƿ�ְ�� = 0 Then
        If strTemp = "" Then
            Err.Raise 9000, gstrSysName, "������̱���ֹ", vbInformation, gstrSysName
            gcn����.Execute "Delete From SICK_PRICE_ITEM Where VISIT_NUMBER='" & str��ˮ�� & "'"
            Unload frm�ȴ����ر�������
            Exit Function
        End If
    Else
        If strTemp = "" Then
            Err.Raise 9000, gstrSysName, "������̱���ֹ", vbInformation, gstrSysName
            'XieRong ɾ��ְ����ϸ��ΪKC227
            gcn����.Execute "Delete From KC27 Where  AKB020='" & gstrҽ���������� & "' And AKC220='" & str��ˮ�� & "'"
            Unload frm�ȴ����ر�������
            Exit Function
        End If
    End If
    Unload frm�ȴ����ر�������
    
    '���ؽ�����
    strSQL = "Select * From MED_RECEIPT_RECORD_MASTER Where CHARGE_NUMBER='" & strTemp & "'"
    Set rsTemp = gcn����.Execute(strSQL)
    
    mcur����֧�� = rsTemp!PAY_SIDE2
    mcurͳ���� = rsTemp!PAY_SIDE3 + rsTemp!PAY_SIDE4 + rsTemp!PAY_SIDE5 + rsTemp!PAY_SIDE6

    cur�����ʻ� = rsTemp!PAY_SIDE2
    cur����ͳ�� = rsTemp!PAY_SIDE3
    cur��ͳ�� = rsTemp!PAY_SIDE4
    cur����ҽ�� = rsTemp!PAY_SIDE5
    cur����Ա���� = rsTemp!PAY_SIDE6
    'д������
    If cur�����ʻ� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|�����ʻ�;" & cur�����ʻ� & ";0"
    End If
    If cur����ͳ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|��������;" & cur����ͳ�� & ";0"
    End If
    If cur��ͳ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|�󲡻���;" & cur��ͳ�� & ";0"
    End If
    If cur����ҽ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|�������;" & cur����ҽ�� & ";0"
    End If
    If cur����Ա���� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & cur����Ա���� & ";0"
    End If
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 2)
    Else
        str���㷽ʽ = "�����ʻ�;" & cur�����ʻ� & ";0"
    End If
    סԺ�������_�������� = str���㷽ʽ
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    סԺ�������_�������� = ""
End Function

Public Function סԺ����_��������(lng����ID As Long, ByVal lng����ID As Long) As Boolean
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
    Call Get�ʻ���Ϣ(TYPE_��������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�������� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�������� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� + mcurͳ���� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� & ",0,0," & _
        "0," & mcurͳ���� & ",0,0," & mcur����֧�� & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    סԺ����_�������� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_��������(lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, sngArrInfo(20) As Single
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, lng����ID As Long
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ���˷��ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    lng����ID = rsTemp("ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��������, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        סԺ�������_�������� = False
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�������� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�������� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - Nvl(rsTemp("�����ʻ�֧��"), 0) & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        Nvl(rsTemp("�����ʻ�֧��"), 0) * -1 & ",Null,Null,Null,Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    סԺ�������_�������� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_��������(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHandle
    '��HIS֮�еĻ������ݽ����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_�������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�������� = False
End Function

Public Function ��Ժ�Ǽ�_��������(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo errHandle
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_�������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�������� = False
End Function

Private Function toHex(ByVal dblNum As Double, Optional ByVal dblKey As Double = 16) As String
    Dim dblTemp As Double, dblMod As Double, strTemp As String
    dblTemp = dblNum
    Do
        dblMod = dblTemp - Int(dblTemp / dblKey) * dblKey
        dblTemp = Int(dblTemp / dblKey)
        If dblMod >= 10 Then
            strTemp = Chr(dblMod + 55) & strTemp
        Else
            strTemp = dblMod & strTemp
        End If
    Loop While dblTemp >= dblKey
    dblMod = dblTemp
    If dblMod >= 10 Then
        strTemp = Chr(dblMod + 55) & strTemp
    Else
        strTemp = dblMod & strTemp
    End If
    toHex = strTemp
End Function

Public Sub WriteInfo(ByVal strInfo As String)
    Dim strFileName As String
    Dim objSystem As FileSystemObject
    Dim objStream As TextStream
    
    strFileName = "C:\��Ϣ" & Format(Date, "MMdd") & ".txt"
    Set objSystem = New FileSystemObject
    If Not objSystem.FileExists(strFileName) Then Call objSystem.CreateTextFile(strFileName, False)
    Set objStream = objSystem.OpenTextFile(strFileName, ForAppending, False, TristateMixed)
    objStream.WriteLine (strInfo)
    objStream.Close
End Sub


