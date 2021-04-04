Attribute VB_Name = "mdl����"
Option Explicit
'�޸ļ�¼:
    '2004-05-27 ZYB ��������ҽ����סԺҽ��
    '1�����÷ָ�()
    '2������ͳ��()
    '�㷨�����ԭ�ֶ��Ƿ�ҽ��=0��������ҽ����Ŀ
    '����Ƿ�ҽ��=1�����������ҽ���ҵ�ǰ�����ｻ�ף�����ҽ����Ŀ��סԺͬ��

'һ��IC����������ṹ����
'1�������ṹ:
'      1��������Ϣ�ṹ       TIC����
'      2��IC����ҽ��Ϣ�ṹ   TBlockPayInfo    �����֧����Ϣ��
Public Type TIC����
    CenterCode       As String * 4      ' ���Ĵ���
    Cardno           As String * 8      ' ����
    IDCardno         As String * 18     ' ���֤�� ���Ȳ����#0
    MediAccountNo    As String * 8      ' ҽ����
    Name             As String * 10     ' ����
    Sex              As String * 1      ' �Ա� 1-��  0-Ů
    Birthday         As String * 8      ' �������� YYYYMMDD
    UnitCode         As String * 5      ' ���˵�λ����
    ClassCode        As String * 2      ' ְ����ݣ�0x����ְ1x������, 05��11Ϊһ���Խɷ�
    DomainCode       As String * 1      ' ְ������ 0-���� 1-��פ��� 2-��ذ���
    MediYear         As String * 4      ' ҽ�����
    InNo             As Long            ' װǮ�ڴ�
    OutSerialNo      As Long            ' ֧��˳���
    InPerAcc         As Double          ' �����ʻ��ۼ�ע����
    OutPerAcc        As Double          ' �����ʻ��ۼ�֧�����
    PlanPaidFee      As Double          ' ͳ�����֧�������ۼƣ�����+���䣩
    PlanPaidAmt      As Double          ' ͳ�����֧������ۼƣ�����+���䣩
    ChronicPaidFee   As Double          ' ���Բ�֧�������ۼ�
    ChronicPaidAmt   As Double          ' ���Բ�֧������ۼ�
    InHosPaidAmt     As Double          ' סԺ�����ʻ�֧�����
    ClinicPaidAmt    As Double          ' ��������ʻ�֧�����
    Password         As String * 4      ' ��������
    InHosTimes       As Long            ' ������ЧסԺ����
    IsOffical        As String * 1      ' ����Ա 0-������-��
    IsAttend         As String * 1      ' ҽ���չ˶��� 0-��1-��
    InpatientFlag    As String * 1      ' סԺ��־ 0-��סԺ 1-סԺ
    Reserved         As String * 2      ' ������ʹ�á���ҪΪ������DLL������������
    QuotaPaidAmt     As Double          ' ���Բ������֧�����
    ChronicSillPaidAmt    As Double     ' ���Բ��𸶽���֧�����
End Type

Private Type TPayInfo
    OccurDate        As String * 8 '  ��ҽ����
    HospitalCode     As String * 4 '  ҽ�ƻ�������
    Tail             As String * 4
    Amount           As Double     '  ���η��úϼ�
    AccPay           As Double     '  �����ʻ�֧��
    CdFlag           As Long
End Type
Private Type TBlockPayInfo
    First            As TPayInfo   ' ��һ�ξ�ҽ��Ϣ
    Second           As TPayInfo   ' �ڶ��ξ�ҽ��Ϣ
    Third            As TPayInfo   ' �����ξ�ҽ��Ϣ
End Type
Private Type TInMoneyParameter
    CenterCode       As String * 4 ' ���Ĵ���
    Cardno           As String * 8 ' ����
    MediYear         As String * 4 ' ҽ�����
    InNo             As Long       ' װǮ�ڴ�
    InPerAcc         As Double     ' �����ʻ��ۼ�ע����
    InExAcc          As Double     ' �������ۼ�ע����
    InSubAcc         As Double     ' �����ʻ��ۼ�ע����
End Type
'����IC����д��������˵��

'2��������д
'      1����IC��������Ϣ
Private Declare Function ReadICCard Lib "ICREAD.DLL" (iIC���� As TIC����) As Long
'      2��дIC��������Ϣ
Private Declare Function WriteICCard Lib "ICWRITE.DLL" (iIC���� As TIC����) As Long

'��¼סԺ���
Private Declare Function ReadICCardPayInfo Lib "ICREAD.DLL" (BlockPayInfo As TBlockPayInfo) As Long
Private Declare Function WriteICCardPayInfo Lib "ICWRITE.DLL" (ByVal strCardNO As String, iIC���� As TPayInfo) As Long

'�������װǮ
'Modified By ���� 2003-12-10 ���������� ��������
Private Declare Function OnLineInMoney Lib "InMoneyOnLine.dll" (ByVal IC_CenterCode As String, ByVal IC_CardNo As String, ByVal IC_MediYear As String, ByVal HosCode As String, ByVal serverIP As String) As Long

Private Enum cardҽ���Ҷ�
    degֹ֧ͣ�� = 1
    deg�ϴ���ϸ = 2 'Ҳֹ֧ͣ��
    deg����֧�� = 3 '���ø����ʻ�֧����ͳ��ͣ----��Ϊ������ͳ��֧�����������������Ժ
    degҽ��֧�� = 4 '
    deg����֧�� = 5 '���·�
End Enum

Private Type ���ݽ�������  '���ṹ�еı��������뱾�ν����йأ�������Щ�ۼ�ֵ�������϶�Ҫ��ӿ���ȡ
    �Ҷ�         As cardҽ���Ҷ�
    ����ID       As Long
    ��ҳID         As Long
    �������     As Long
    ���         As Long
    ����סԺ     As Boolean
    �������     As Boolean
    סԺ����       As Long
    סԺ��������   As Long
    ��;����       As Long
    ����         As Currency
    �ⶥ��         As Currency
    ʵ������     As Currency  '����ʵ��֧�������߽��
    ��������     As Currency  '����Ԥ�ƻ�֧�������߽��
    ��������       As Currency
    ȫ�Է�         As Currency
    �����Ը�       As Currency
    ����ͳ��       As Currency
    ҽ����Ŀ���   As Currency
    ������Ŀ���   As Currency
    �����ʻ�֧��   As Currency
    סԺ����       As Long
    ͳ����֧����� As Currency   '��������������ʡ����ʱ���ܴӿ���ȥȡ������ֵ�����粻ʹ���ۼƵ���;���㣬��Ҫ�����ݿ�����ǰ�Ľ����¼����
    ͳ����֧������ As Currency
    ����ͳ��֧��   As Currency
    ����ͳ�����   As Currency
    ��������֧��   As Currency
    ������������   As Currency
    ͳ�����֧��   As Currency
    ͳ������Ը�   As Currency
    �μӲ��䱣��   As Long
    �������֧��   As Currency
    ��������Ը�   As Currency
    ��������֧��   As Currency
    ���������Ը�   As Currency
    �������ⶥ��   As Currency
    ������ⶥ��   As Currency
    ������֧��   As Currency
    ������������ʻ�֧��  As Currency
    ����סԺ�����ʻ�֧��  As Currency
    �������Բ��𸶽�  As Currency
    ���ִ��� As String
    �������� As String
    �������� As String
End Type

Private Type ����
    �����ʻ�֧��ȫ�Է� As Boolean
    �����ʻ�֧�������Ը� As Boolean
    �����ʻ�֧������ As Boolean
    ȫ��ͳ�� As Boolean
    ���÷ⶥ As Boolean
    ���ö�ֵ As Boolean
    ʹ���ۼ� As Boolean
    ���䱨�����𸶽� As Boolean
    �����ڶ��� As Boolean
    �����𸶽����� As Long          '0-��ԭ�𸶽�1�������ۣ�2���𸶽�
    ��������סԺ���� As Long        '1����һ�Σ�0������
End Type
'-------------��������
Public gIC���� As TIC����                 'ȫ�ֶ���Ĵ洢IC����Ϣ�Ľṹ
Public gIC����Temp As TIC����             '��Ҫ������Զ��������������
Public gcn���� As New ADODB.Connection        '���ӵ�ҽ��ǰ�÷�����
Private m���� As ���ݽ�������
Private m���� As ����

'���⼸���漰�����ĵı�����Ϊȫ��

'-------------��������

Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    '��Ϊ����Ҫ����ҽ��������������ǿ�Ƽ����������
    ҽ����ʼ��_���� = ���ҽ��������_����
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������strSelfNO-���˱�ţ�ˢ���õ���strSelfPwd-�������룻
'      bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ� �ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim strIdentify As String, strAddition As String
    Dim strBirthday As String, datToday As Date
    Dim str����ID As String, lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If frmIdentify����.GetPatient(bytType, False, lng����ID, str����ID) = True Then
        '���ʶ����ɣ����ز�����Ϣ
        With gIC����
            Call ҽ���Ҷ�(.CenterCode, .Cardno)
            If m����.�Ҷ� = degֹ֧ͣ�� Then
                MsgBox "�ò�����ʱֹͣҽ��֧�����뵽ҽ�����Ĵ���", vbInformation, gstrSysName
                Exit Function
            End If
            
            If bytType = 1 Then
                '�������ƵĲ��˽�������
                If m����.�Ҷ� = deg�ϴ���ϸ Then
                    MsgBox "�ò��˲���ʹ��ͳ�����֧��סԺ���á�", vbExclamation, gstrSysName
                End If
            End If
            
            '�������˵�����Ϣ�������ʽ��
            '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
            '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
            '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
            strIdentify = TrimStr(.Cardno)                              '0����
            strIdentify = strIdentify & ";" & TrimStr(.MediAccountNo)   '1ҽ����
            strIdentify = strIdentify & ";" & TrimStr(.Password)        '2����
            strIdentify = strIdentify & ";" & TrimStr(.Name) '3����
            strIdentify = strIdentify & ";" & IIf(.Sex = "1", "��", "Ů")   '4�Ա�
            
            strBirthday = TrimStr(.Birthday)
            datToday = zlDatabase.Currentdate
            If strBirthday = "" Then
                strBirthday = Format(datToday, "yyyy-MM-dd")
            Else
                strBirthday = Mid(strBirthday, 1, 4) & "-" & Mid(strBirthday, 5, 2) & "-" & Mid(strBirthday, 7, 2)
            End If
            strIdentify = strIdentify & ";" & strBirthday              '5��������
            strIdentify = strIdentify & ";" & TrimStr(.IDCardno)   '6���֤
            strIdentify = strIdentify & ";" & TrimStr(.UnitCode) & "(" & TrimStr(.UnitCode) & ")"  '7.��λ����(����)
            
            '�õ��������
            gstrSQL = "select ��� from ��������Ŀ¼ where ����=[1] and ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, .CenterCode)
            
            If rsTemp.RecordCount = 0 Then
                ��ݱ�ʶ_���� = ""
                MsgBox "�ò�������������δ����������ʹ�á�", vbInformation, gstrSysName
                Exit Function
            Else
                m����.������� = rsTemp("���")
            End If
            
            '�õ�ԭסԺ����
            If bytType <> 1 Then
                gstrSQL = "Select Nvl(����ID,0) ����ID From �����ʻ� Where ����=[1] And ҽ����=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�õ�ԭסԺ����", TYPE_������, CStr(TrimStr(.MediAccountNo)))
                If Not rsTemp.EOF Then
                    lng����ID = rsTemp!����ID
                End If
            End If
            
            strAddition = ";" & m����.�������                          '8.���Ĵ���
            strAddition = strAddition & ";"                             '9.˳���
            strAddition = strAddition & ";" & TrimStr(.ClassCode)       '10��Ա���
            strAddition = strAddition & ";" & (.InPerAcc - .OutPerAcc)  '11�ʻ����
            strAddition = strAddition & ";" & .InpatientFlag            '12��ǰ״̬
            strAddition = strAddition & ";" & IIf(lng����ID > 0, lng����ID, "") '13����ID

'            strAddition = strAddition & ";" & IIf(Left(TrimStr(.ClassCode), 1) = "0", 1, 0)    '14��ְ
            Select Case Left(TrimStr(.ClassCode), 1)                    '14��ְ(1,2,3)
            Case "0"
                strAddition = strAddition & ";1"
            Case "1"
                strAddition = strAddition & ";2"
            Case "5"
                strAddition = strAddition & ";3"
            End Select
            strAddition = strAddition & ";"                             '15����֤��
            strAddition = strAddition & ";" & DateDiff("yyyy", CDate(strBirthday), datToday) '16�����
            strAddition = strAddition & ";" & m����.�Ҷ�                   '17�Ҷȼ�
            strAddition = strAddition & ";" & .InPerAcc                 '18�ʻ������ۼ�
            strAddition = strAddition & ";" & .OutPerAcc                '19�ʻ�֧���ۼ�
            strAddition = strAddition & ";" & .PlanPaidFee              '20����ͳ���ۼ�
            strAddition = strAddition & ";" & .PlanPaidAmt              '21ͳ�ﱨ���ۼ�
            strAddition = strAddition & ";" & .InHosTimes               '22סԺ�����ۼ�
            strAddition = strAddition & ";"                             '23�������� (1����������)
            
            lng����ID = BuildPatiInfo(bytType, strIdentify & strAddition, lng����ID, TYPE_������)
            '���ظ�ʽ:�м���벡��ID
            ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & strAddition
            
            '���²�����Ϣ
            If bytType = 1 Then
                gstrSQL = "zlyb.zl_������Ϣ_INSERT(" & TYPE_������ & "," & lng����ID & ",0,0,'" & str����ID & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
            End If
        End With
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����:
'����: ���ظ����ʻ����Ľ��
    
    On Error GoTo errHandle
    
    'ִ��װǮ������˳��Ͷ�ȡ�����µĸ�������
    If װǮ����(lng����ID) = True Then
        '��������
        Call ҽ���Ҷ�(gIC����.CenterCode, gIC����.Cardno)
        If m����.�Ҷ� > deg�ϴ���ϸ Then
            '�������
            �������_���� = gIC����.InPerAcc - gIC����.OutPerAcc
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
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
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim clsҽ�� As New clsInsure, tmp���� As ���ݽ�������, tmp���� As ����
    Dim dblȫ�Է� As Currency, dbl�����Ը� As Currency, dbl����ͳ�� As Currency
    Dim lng����ID As Long, rsTemp As New ADODB.Recordset
    
    m���� = tmp����         '��ʼ������
    m���� = tmp����
    
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�з��ã����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    On Error GoTo errHandle
    
    'Modified By ���� 2003-12-10 ����������
    If Calc���÷ָ�(rs��ϸ, False, dblȫ�Է�, dbl�����Ը�, dbl����ͳ��, False, True) = False Then
        Exit Function
    End If
    With m����
        .�������� = dblȫ�Է� + dbl����ͳ�� + dbl�����Ը�
        .ȫ�Է� = dblȫ�Է�
        .�����Ը� = dbl�����Ը�
        .����ͳ�� = dbl����ͳ��
    End With
    
    gstrSQL = "Select B.���� From �����ʻ� A,���ղ��� B where A.����=[1] And A.����ID=[2] And A.����ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����������", TYPE_������, lng����ID)
    If rsTemp.EOF = False Then
        gstrSQL = "Select B.����,����,��� From ���ղ��� B where B.����=" & TYPE_������ & " And B.����='" & rsTemp("����") & "'"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
        If rsTemp.EOF = False Then
            m����.���ִ��� = rsTemp("����")
            m����.�������� = Nvl(rsTemp("����"))
            m����.�������� = Nvl(rsTemp("���"))
        Else
            m����.���ִ��� = ""
            m����.�������� = ""
            m����.�������� = ""
        End If
    End If
    
    '�����涨
    m����.�����ʻ�֧��ȫ�Է� = clsҽ��.GetCapability(support�շ��ʻ�ȫ�Է�, 0, TYPE_������)
    m����.�����ʻ�֧�������Ը� = clsҽ��.GetCapability(support�շ��ʻ������Ը�, 0, TYPE_������)
    
    gstrSQL = "SELECT B.ҽ����,A.�����ڶ���,A.��ֵ����,A.�ⶥ����,A.ʹ���ۼƱ���,A.�����˻���֧�������Ը� " & _
               " FROM ��������Ŀ¼ A,�������� B " & _
               " WHERE A.����=" & TYPE_������ & " AND A.����='" & gIC����.CenterCode & "' AND A.��������=B.���� AND A.����=B.���� "
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        m����.�����ʻ�֧�������Ը� = Nvl(rsTemp("�����˻���֧�������Ը�")) = 1
    End If
    
    With m����
        .�����ʻ�֧�� = 0
        If m����.�����ʻ�֧��ȫ�Է� = True Then
            .�����ʻ�֧�� = dblȫ�Է�
        End If
        
        If Is���ݲ���(lng����ID) = True Then
            '���ַ��ÿ�����ҽ��������
            .����ͳ����� = .����ͳ�� + .�����Ը�
            If Isȫ��ͳ��(lng����ID, TYPE_������) = True Then
                '�����Ը�Ҳ����ҽ������֧��
                .����ͳ��֧�� = .����ͳ�� + .�����Ը�
            Else
                .����ͳ��֧�� = .����ͳ��
                If m����.�����ʻ�֧�������Ը� = True Then
                    .�����ʻ�֧�� = .�����ʻ�֧�� + .�����Ը�
                End If
            End If
            .ͳ�����֧�� = .����ͳ��֧��
        Else
            'ֻ���ø����ʻ�֧��
            .�����ʻ�֧�� = .�����ʻ�֧�� + .����ͳ��
            If m����.�����ʻ�֧�������Ը� = True Then
                .�����ʻ�֧�� = .�����ʻ�֧�� + dbl�����Ը�
            End If
        End If
        
        '����ʻ�����Ƿ��㹻֧��
        If .�����ʻ�֧�� > gIC����.InPerAcc - gIC����.OutPerAcc Then
            .�����ʻ�֧�� = gIC����.InPerAcc - gIC����.OutPerAcc
            If .�����ʻ�֧�� < 0 Then .�����ʻ�֧�� = 0
        End If
    End With
    
    '����ҽ���Ҷ�
    Call ҽ���Ҷ�(gIC����.CenterCode, gIC����.Cardno)
    If m����.�Ҷ� < deg����֧�� Then m����.�����ʻ�֧�� = 0
    
    str���㷽ʽ = "�����ʻ�;" & m����.�����ʻ�֧�� & ";1"
    If m����.�Ҷ� >= degҽ��֧�� Then
        If m����.ͳ�����֧�� > 0 Then str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & m����.ͳ�����֧�� & ";0"
        If m����.�������֧�� > 0 Then str���㷽ʽ = str���㷽ʽ & "|�������;" & m����.�������֧�� & ";0"
    End If
    �����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, ByVal curȫ�Է� As Currency, ByVal cur�����Ը� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    Dim rsTemp As New ADODB.Recordset, rs��Ϣ As New ADODB.Recordset
    Dim ic���� As TIC����            '�����нṹ�����󷵻�ֵ�����⣨��Ҫ���漰���ļ�����Ա��
    Dim strҽԺ���� As String
    Dim lng���� As Long, lng����ID As Long
    Dim curͳ���� As Currency
    Dim dat��ǰ���� As Date
    Dim bln���� As Boolean
    
    On Error GoTo errHandle
        
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.��¼����,A.��¼״̬,A.�շ�ϸĿID,C.��Ŀ����,B.����,B.����,A.ʵ�ս�� " & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as ���� " & _
              "  From ������ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
              "  where A.����ID=[1] And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0" & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= " & TYPE_������ & _
              "  Order by A.����ID,A.����ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    'Modified By ���� 2003-12-10 ����������
    If Calc���÷ָ�(rsTemp, True, curȫ�Է�, cur�����Ը�, curͳ����, False, True) = False Then
        Exit Function
    End If
    With m����
        .ȫ�Է� = curȫ�Է�
        .�����Ը� = cur�����Ը�
        .����ͳ�� = curͳ����
        .�������� = curȫ�Է� + cur�����Ը� + curͳ����
    End With
    
    '����������д�����ս����¼�е���Ϣ
    gstrSQL = "SELECT A.����ID,A.NO,A.ʵ��Ʊ��,A.��¼����,substr(B.����,1,8) as ����,substr(B.�Ա�,1,2) as �Ա�,floor(MONTHS_BETWEEN(A.�Ǽ�ʱ��,B.��������)/12) AS ����" & _
              "         ,B.���֤��,C.����,C.ҽ����,a.�Ǽ�ʱ��,substr(A.����Ա����,1,8) as ҽ��,D.���� AS ����" & _
              "  FROM ������ü�¼ A,������Ϣ B,�����ʻ� C,���ű� D" & _
              "  Where A.����ID =[1] And A.����ID = B.����ID And B.����ID = C.����ID And C.���� =[2] And A.��������ID = D.ID(+) and rownum<2"
    Set rs��Ϣ = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID, TYPE_������)
    lng����ID = rs��Ϣ("����ID")
    
    If ReadIC(lng����ID, 0, False, "�շ�ʱ����ʧ�ܡ�", ic����, bln����) = False Then
        Exit Function
    End If
    
    Call ҽ���Ҷ�(ic����.CenterCode, ic����.Cardno)
    
    If m����.�Ҷ� = degֹ֧ͣ�� Then
        '�����ٴ����������
        �������_���� = True
        Exit Function
    End If
    
    dat��ǰ���� = zlDatabase.Currentdate
    
    '�жϸò��˵Ŀ��Ƿ������ȷ
    If ���IC��(lng����ID, TrimStr(ic����.Cardno), TrimStr(ic����.CenterCode)) = False Then Exit Function
    
    With ic����
        'Ϊ�˱�֤��ȫ���ۼ����ݻ��Ƕ���������

        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Format(dat��ǰ����, "yyyy") & "," & _
            .InPerAcc & "," & .OutPerAcc + cur�����ʻ� & "," & .PlanPaidFee & "," & _
            .PlanPaidAmt & "," & .InHosTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
            Format(dat��ǰ����, "yyyy") & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidFee & "," & _
            .PlanPaidAmt & "," & .InHosTimes & ",0,0,0," & _
            m����.�������� & "," & curȫ�Է� & "," & cur�����Ը� & "," & curͳ���� & ",0,0,0," & _
            cur�����ʻ� & ",'" & .OutSerialNo + 1 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
    
    'ҽ����������������Ȼ��������������һ�������У�������Ҫ��д����һ��
    With ic����
        gstrSQL = "INSERT INTO ���ս����¼ " & _
                  "   (����,��¼id,����,���,���Ĵ���,���,����id,��ҳid,����,�Ա�,����,ҽ����,����,���֤��,��ݴ���,��λҽ���� " & _
                  ",�Ƿ���Ա,�Ƿ�ҽ���չ˶���,�μӲ��䱣��,�ʻ��ۼ�����,�ʻ��ۼ�֧��,ͳ����֧�����,ͳ����֧������ " & _
                  ",������֧�����,������֧������,�����𸶽���֧��,��������,����סԺ���� " & _
                  ",��������ʻ�֧�����,סԺ�����ʻ�֧�����,�����֧�����,��������,ҽ������,���ִ���,��������,�������� " & _
                  ",�������ý��,�����ʻ�֧��,ȫ�Ը����,�����Ը����,ת�������Ը�,����,�ⶥ��,ʵ������ " & _
                  ",����ͳ��֧��,����ͳ�����,��������֧��,������������,ͳ����֧��,ͳ�����Ը�,ͳ�����֧��,ͳ������Ը� " & _
                  ",�������֧��,��������Ը�,��������֧��,���������Ը� " & _
                  ",�������ⶥ��,������ⶥ��,������֧��,������������ʻ�֧��,����סԺ�����ʻ�֧��,�������Բ��𸶽� " & _
                  ",���Ҷȼ�,��Ʊ��,Ʊ������,��Ʊ��־,����Ʊ�ݺ�,֧��˳���,�Ƿ��ϴ�) " & _
                  " Values "
         gstrSQL = gstrSQL & " (1," & lng����ID & "," & TYPE_������ & "," & .MediYear & ",'" & .CenterCode & "','" & rs��Ϣ("NO") & "1" & rs��Ϣ!��¼���� & "'," & lng����ID & ",0,'" & rs��Ϣ("����") & _
                  "','" & rs��Ϣ("�Ա�") & "'," & rs��Ϣ("����") & ",'" & rs��Ϣ("ҽ����") & "','" & rs��Ϣ("����") & "','" & rs��Ϣ("���֤��") & "','" & .ClassCode & "','" & .UnitCode & "' " & _
                  "," & .IsOffical & "," & .IsAttend & "," & m����.�μӲ��䱣�� & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidAmt & "," & .PlanPaidFee & _
                  "," & .ChronicPaidAmt & "," & .ChronicPaidFee & "," & .ChronicSillPaidAmt & ",null," & m����.סԺ�������� & _
                  "," & .ClinicPaidAmt & "," & .InHosPaidAmt & "," & .QuotaPaidAmt & ",'" & rs��Ϣ("����") & "','" & rs��Ϣ("ҽ��") & "','" & m����.���ִ��� & "','" & m����.�������� & "','" & m����.�������� & "' " & _
                  "," & m����.�������� & "," & cur�����ʻ� & "," & curȫ�Է� & "," & cur�����Ը� & ",0,0,0,0 " & _
                  "," & m����.����ͳ��֧�� & "," & m����.����ͳ����� & "," & m����.��������֧�� & "," & m����.������������ & "," & _
                  (m����.ͳ�����֧�� + m����.�������֧�� + m����.��������֧��) & "," & (m����.ͳ������Ը� + m����.��������Ը� + m����.���������Ը�) & "," & m����.ͳ�����֧�� & "," & m����.ͳ������Ը� & " " & _
                  "," & m����.�������֧�� & "," & m����.��������Ը� & ",0,0 " & _
                  "," & m����.�������ⶥ�� & "," & m����.������ⶥ�� & "," & m����.������֧�� & "," & cur�����ʻ� & ",0," & m����.�������Բ��𸶽� & " " & _
                  "," & m����.�Ҷ� & ",'" & Nvl(rs��Ϣ("ʵ��Ʊ��"), " ") & "'," & GetOracleFormat(rs��Ϣ("�Ǽ�ʱ��")) & ",1,'','" & .OutSerialNo + 1 & "',0)"
        '׼��д��Ŀ���Ϣ
        .OutPerAcc = .OutPerAcc + cur�����ʻ�                   '�����ʻ��ۼ�֧�����
        .ClinicPaidAmt = .ClinicPaidAmt + cur�����ʻ�           '��������ʻ�֧�����
        .InHosTimes = .InHosTimes + m����.סԺ��������          '��Щ���ز�������סԺ����
        .PlanPaidFee = .PlanPaidFee + m����.����ͳ�����        'ͳ�����֧�������ۼƣ�����+���䣩
        .PlanPaidAmt = .PlanPaidAmt + m����.����ͳ��֧��        ' ͳ�����֧������ۼƣ�����+���䣩
        .ChronicPaidFee = .ChronicPaidFee + m����.������������                 '���Բ�֧�������ۼ�
        .ChronicPaidAmt = .ChronicPaidAmt + m����.��������֧��                 '���Բ�֧������ۼ�
        .QuotaPaidAmt = .QuotaPaidAmt + m����.������֧��                     '���Բ������֧�����
        .ChronicSillPaidAmt = .ChronicSillPaidAmt + m����.�������Բ��𸶽�     '���Բ��𸶽���֧�����
        .OutSerialNo = .OutSerialNo + 1           ' ֧��˳���
    End With
        
    Dim payLog As TPayInfo
    With payLog
        .HospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
        .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' ����
        .AccPay = m����.�����ʻ�֧��
        .Amount = m����.��������
        .CdFlag = 1
    End With
    
    '��ɿ�д��
    Dim str������ As String
    With m����
        str������ = ic����.CenterCode & "|" & gstrҽԺ���� & "|0|" & rs��Ϣ("NO") & "1|" & _
                    TrimStr(ic����.MediAccountNo) & "|" & cur�����ʻ� & "|" & .ͳ�����֧�� & "|" & .�������֧�� & "|" & _
                    .����ͳ����� & "|" & .����ͳ��֧�� & "|" & .סԺ�������� & "|" & .�������ⶥ�� & "|1"
    End With
    
    If WriteIC(bln����, True, 0, gstrSQL, ic����, payLog, str������) = False Then
        Exit Function
    End If
    
    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    Dim rsTemp As New ADODB.Recordset, rs���� As New ADODB.Recordset
    Dim ic���� As TIC����
    Dim lng��� As Long, lng����ID As Long
    Dim dat��ǰ���� As Date
    Dim bln���� As Boolean
    
    On Error GoTo errHandle
    
    gstrSQL = "Select *  From ���ս����¼ Where ��¼ID=" & lng����ID
    rs����.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    
    lng����ID = rs����("����ID")
        
    If ReadIC(lng����ID, 0, True, "�˷�ʱ����ʧ�ܡ�", ic����, bln����) = False Then
        Exit Function
    End If
    
    'ȡ�������
    gstrSQL = "Select ��� From ��������Ŀ¼ Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", ic����.CenterCode)
    m����.������� = rsTemp!���
    
    If Val(ic����.MediYear) <> rs����("���") Then
        Err.Raise 9000, gstrSysName, "���겻�����ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Get��ǰҽ����) <> rs����("���") Then
        Err.Raise 9000, gstrSysName, "���겻�����ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ҽ���Ҷ�(ic����.CenterCode, ic����.Cardno)
    
    If m����.�Ҷ� = degֹ֧ͣ�� Then
        '�����ٴ����������
        '����������_���� = True
        Exit Function
    End If
    
    dat��ǰ���� = zlDatabase.Currentdate
        
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    lng��� = rsTemp("����ID")
    
    With ic����
        'Ϊ�˱�֤��ȫ���ۼ����ݻ��Ƕ���������

        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Format(dat��ǰ����, "yyyy") & "," & _
            .InPerAcc & "," & .OutPerAcc - cur�����ʻ� & "," & .PlanPaidFee - rs����("����ͳ�����") & "," & _
            .PlanPaidAmt - rs����("����ͳ��֧��") & "," & .InHosTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        gstrSQL = "zl_���ս����¼_insert(1," & lng��� & "," & TYPE_������ & "," & lng����ID & "," & _
            Format(dat��ǰ����, "yyyy") & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidFee & "," & _
            .PlanPaidAmt & "," & .InHosTimes & ",0,0,0," & _
            rs����("�������ý��") * -1 & "," & rs����("ȫ�Ը����") * -1 & "," & rs����("�����Ը����") * -1 & "," & rs����("����ͳ�����") * -1 & ",0,0,0," & cur�����ʻ� * -1 & ",'" & .OutSerialNo + 1 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
    
    'ҽ����������������Ȼ��������������һ�������У�������Ҫ��д����һ��
    With ic����
        gstrSQL = "INSERT INTO ���ս����¼ " & _
                  "   (����,��¼id,����,���,���Ĵ���,���,����id,��ҳid,����,�Ա�,����,ҽ����,����,���֤��,��ݴ���,��λҽ���� " & _
                  ",�Ƿ���Ա,�Ƿ�ҽ���չ˶���,�μӲ��䱣��,�ʻ��ۼ�����,�ʻ��ۼ�֧��,ͳ����֧�����,ͳ����֧������ " & _
                  ",������֧�����,������֧������,�����𸶽���֧��,��������,����סԺ���� " & _
                  ",��������ʻ�֧�����,סԺ�����ʻ�֧�����,�����֧�����,��������,ҽ������,���ִ���,��������,�������� " & _
                  ",�������ý��,�����ʻ�֧��,ȫ�Ը����,�����Ը����,ת�������Ը�,����,�ⶥ��,ʵ������ " & _
                  ",����ͳ��֧��,����ͳ�����,��������֧��,������������,ͳ����֧��,ͳ�����Ը�,ͳ�����֧��,ͳ������Ը� " & _
                  ",�������֧��,��������Ը�,��������֧��,���������Ը� " & _
                  ",�������ⶥ��,������ⶥ��,������֧��,������������ʻ�֧��,����סԺ�����ʻ�֧��,�������Բ��𸶽� " & _
                  ",���Ҷȼ�,��Ʊ��,Ʊ������,��Ʊ��־,����Ʊ�ݺ�,֧��˳���,�Ƿ��ϴ�) " & _
                  " Values "
         gstrSQL = gstrSQL & " (1," & lng��� & "," & TYPE_������ & "," & .MediYear & ",'" & .CenterCode & "','" & Mid(rs����("���"), 1, Len(rs����("���")) - 2) & "2" & Right(rs����!���, 1) & "'," & lng����ID & ",0,'" & rs����("����") & _
                  "','" & rs����("�Ա�") & "'," & rs����("����") & ",'" & rs����("ҽ����") & "','" & rs����("����") & "','" & rs����("���֤��") & "','" & .ClassCode & "','" & .UnitCode & "' " & _
                  "," & .IsOffical & "," & .IsAttend & "," & rs����("�μӲ��䱣��") & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidAmt & "," & .PlanPaidFee & _
                  "," & .ChronicPaidAmt & "," & .ChronicPaidFee & "," & .ChronicSillPaidAmt & ",null," & rs����("����סԺ����") & _
                  "," & .ClinicPaidAmt & "," & .InHosPaidAmt & "," & .QuotaPaidAmt & ",'" & rs����("��������") & "','" & rs����("ҽ������") & "','" & rs����("���ִ���") & "','" & rs����("��������") & "','" & rs����("��������") & "' " & _
                  "," & rs����("�������ý��") & "," & cur�����ʻ� & "," & rs����("ȫ�Ը����") & "," & rs����("�����Ը����") & ",0,0,0,0 " & _
                  "," & rs����("����ͳ��֧��") & "," & rs����("����ͳ�����") & "," & rs����("��������֧��") & "," & rs����("������������") & "," & rs����("ͳ����֧��") & "," & rs����("ͳ�����Ը�") & "," & rs����("ͳ�����֧��") & "," & rs����("ͳ������Ը�") & " " & _
                  "," & rs����("�������֧��") & "," & rs����("��������Ը�") & ",0,0 " & _
                  "," & rs����("�������ⶥ��") & "," & rs����("������ⶥ��") & "," & rs����("������֧��") & "," & rs����("������������ʻ�֧��") & "," & rs����("����סԺ�����ʻ�֧��") & "," & rs����("�������Բ��𸶽�") & " " & _
                  "," & m����.�Ҷ� & ",'" & Nvl(rs����("��Ʊ��")) & "',sysdate,-1,'" & rs����("���") & "','" & .OutSerialNo + 1 & "',0)"
        '׼��д��
        .OutPerAcc = .OutPerAcc - cur�����ʻ�                  '�����ʻ��ۼ�֧�����
        .ClinicPaidAmt = .ClinicPaidAmt - cur�����ʻ�           '��������ʻ�֧�����
        .InHosTimes = .InHosTimes - rs����("����סԺ����")      '��Щ���ز�������סԺ����
        .PlanPaidFee = .PlanPaidFee - rs����("����ͳ�����")      'ͳ�����֧�������ۼƣ�����+���䣩
        .PlanPaidAmt = .PlanPaidAmt - rs����("����ͳ��֧��")        ' ͳ�����֧������ۼƣ�����+���䣩
        .ChronicPaidFee = .ChronicPaidFee - rs����("������������")                '���Բ�֧�������ۼ�
        .ChronicPaidAmt = .ChronicPaidAmt - rs����("��������֧��")                '���Բ�֧������ۼ�
        .QuotaPaidAmt = .QuotaPaidAmt - rs����("������֧��")                     '���Բ������֧�����
        .ChronicSillPaidAmt = .ChronicSillPaidAmt - rs����("�������Բ��𸶽�")      '���Բ��𸶽���֧�����
        .OutSerialNo = .OutSerialNo + 1           ' ֧��˳���
    End With
        
    Dim payLog As TPayInfo
    With payLog
        .HospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
        .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' ����
        .AccPay = cur�����ʻ�
        .Amount = rs����("�������ý��")
        .CdFlag = 0
    End With
    
    '��ɿ�д��
    Dim str������ As String
        
    str������ = ic����.CenterCode & "|" & gstrҽԺ���� & "|0|" & Mid(rs����("���"), 1, Len(rs����("���")) - 1) & "2|" & _
                TrimStr(ic����.MediAccountNo) & "|" & cur�����ʻ� & "|" & rs����("ͳ�����֧��") & "|" & rs����("�������֧��") & "|" & _
                rs����("����ͳ�����") & "|" & rs����("����ͳ��֧��") & "|" & rs����("����סԺ����") & "|" & rs����("�������ⶥ��") & "|-1"
    
    
    If WriteIC(bln����, True, 0, gstrSQL, ic����, payLog, str������) = False Then
        Exit Function
    End If
    
    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �����ʻ�תԤ��_����(lngԤ��ID As Long, curMoney As Currency) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
           
    '��������ҽ����֧�ָ�ҵ������ǿ�з���ʧ��
    �����ʻ�תԤ��_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim ic��Ժ As TIC����       '��Ժ�ǼǶ����ṹ
    Dim dat��ǰ���� As Date
    Dim bln���� As Boolean
    
    On Error GoTo errHandle
    
    If ReadIC(lng����ID, 1, True, "��Ժ�Ǽ�ʱ����ʧ�ܡ�", ic��Ժ, bln����) = False Then
        Exit Function
    End If
        
    dat��ǰ���� = zlDatabase.Currentdate
    
    Call ҽ���Ҷ�(ic��Ժ.CenterCode, ic��Ժ.Cardno)
    
    If m����.�Ҷ� = degֹ֧ͣ�� Then
        '�����ٴ����������
        ��Ժ�Ǽ�_���� = False
        MsgBox "�ò����Ѿ�ֹͣҽ��֧����������Ϊҽ��������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    If m����.�Ҷ� = deg����֧�� Then
        '�����ٴ����������
        ��Ժ�Ǽ�_���� = False
        MsgBox "�ò����Ѿ�ֹͣҽ��֧�����Ҷ�Ϊ3����������Ϊҽ��������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If

    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    'Modified by ���� 2004-01-07 ����ǰҽ����д�뱣���ʻ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'ҽ����','''" & Get��ǰҽ���� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "zlyb.zl_������Ϣ_��Ժ(" & TYPE_������ & "," & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������֤��¼��Ĳ�����Ϣ����Ϊ���ε���Ժ����")
    
    Dim payLog As TPayInfo
    '��ɿ�д��
    With ic��Ժ
        .InpatientFlag = 1
    End With
    If WriteIC(bln����, False, 1, "", ic��Ժ, payLog, "") = False Then
        Exit Function
    End If
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(ByVal lng����ID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'���أ����׳ɹ�����true�����򣬷���false
    Dim ic��Ժ As TIC����
    Dim bln���� As Boolean
    
    On Error GoTo errHandle
    
    If ReadIC(lng����ID, 1, True, "��Ժ����ʱ����ʧ�ܡ�", ic��Ժ, bln����) = False Then
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    Dim payLog As TPayInfo
    '��ɿ�д��
    With ic��Ժ
        .InpatientFlag = 0
    End With
    If WriteIC(bln����, False, 1, "", ic��Ժ, payLog, "") = False Then
        Exit Function
    End If
        
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rsExse As Recordset) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼����
'      NO����š�ҽ����Ŀ���롢�շ����ơ��������š���񡢲��ء��������۸񡢽�ҽ��,�Ǽ�ʱ��(����ʱ��),Ӥ����,�շ����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim ic���� As TIC����, tmp���� As ���ݽ�������, tmp���� As ����
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency, curͳ�� As Currency
    Dim bln���� As Boolean
    Dim str��Ժ��� As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If ���ҽ��������_���� = False Then
        '�������ӵ�ǰ�÷�����������Ϊ����ʹ��
        Exit Function
    End If
    
    gIC���� = ic���� '��˿��Խ��������ڲ������ĳ�ʼ��
    m���� = tmp����
    m���� = tmp����
    
    If ReadIC(rsExse("����ID"), 1, True, "������Ϣʧ�ܡ�", gIC����, bln����) = False Then
        Exit Function
    End If
        
    '���һЩ���ݵĳ�ʼ������������ԱҲҪʹ�õ�����
    With m����
        .����ID = rsExse("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", CLng(rsExse("����ID")))
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
        g��������.��ҳID = rsTemp("��ҳID")
    
        '�����ڳ�Ժ���ʺ��ٴν��н���
        gstrSQL = "SELECT ����ID FROM ���ս����¼ WHERE ��;����=0 AND ����ID=[1] AND ��ҳID=[2] AND ����=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", .����ID, .��ҳID, TYPE_������)
        
        If rsTemp.RecordCount > 0 Then
            MsgBox "�����Ѿ����й�סԺ���㣬�����ٽ��н��ʲ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    '��鲡�˵ķ����Ƿ��Ѿ����¼�����������
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,C.��Ŀ����,B.����,B.����,A.ʵ�ս�� " & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as ���� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
              "  where A.����ID=[1] and A.��ҳID=[2] and A.���ʷ���=1 And A.����Ա���� is not null AND A.ʵ�ս�� IS NOT NULL " & _
              "  And A.����ID Is NULL And nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= [3]" & _
              "  Order by A.����ʱ��,A.NO,A.��¼����,Decode(A.��¼״̬,3,1,1,1,2),A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", m����.����ID, m����.��ҳID, TYPE_������)
    If rsTemp.EOF = False Then
        '������û�зָ���õ���ϸ
        If Calc���÷ָ�(rsTemp, True, curȫ�Է�, cur�����Ը�, curͳ��) = False Then
            Exit Function
        End If
    End If
    
    'Ŀǰֻ������ҽ��ʹ�øò���
    'Modified by ���� 2004-01-07
    gstrSQL = "select A.����ID,A.ҽ����,B.����,B.��� " & _
            " from �����ʻ� A,��������Ŀ¼ B " & _
            " where A.����ID=[1] and A.����=[2]" & _
            "  and A.����=B.���� and A.����=B.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", m����.����ID, TYPE_������)
    If rsTemp.EOF = True Then
        MsgBox "��ϵͳ����Ա���ҽ�����ĵ����á�", vbInformation, gstrSysName
        Exit Function
    End If
    If Nvl(rsTemp!����ID, 0) = 0 Then
        MsgBox "û��ѡ���֣���������ʣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    m����.������� = rsTemp("���")
    'Modified by ���� 2004-01-07
    str��Ժ��� = Nvl(rsTemp!ҽ����)
    
    gstrSQL = "Select B.���� From �����ʻ� A,���ղ��� B where A.����=[1] And A.����ID=[2] And A.����ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺ�������", TYPE_������, m����.����ID)
    If rsTemp.EOF = False Then
        gstrSQL = "Select B.����,����,��� From ���ղ��� B where B.����=" & TYPE_������ & " And B.����='" & rsTemp("����") & "'"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
        m����.���ִ��� = rsTemp("����")
        m����.�������� = Nvl(rsTemp("����"))
        m����.�������� = Nvl(rsTemp("���"))
    End If
    
    '1.2 �������˵���Ժʱ��
    gstrSQL = "select ��Ժ����,nvl(��Ժ����,to_date('3000-01-01','yyyy-MM-dd')) as ��Ժ����,sysdate ��ǰ���� " & _
              "from ������ҳ where ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", m����.����ID, m����.��ҳID)
    
    With m����
        If rsTemp("��Ժ����") = CDate("3000-01-01") Then
            .��;���� = 1
        Else
            '��ʾ�ò����Ѿ���Ժ
            .��;���� = 0
        End If
        'Modified By ���� 2003-12-10 ����������
        .��� = Get��ǰҽ����
        'Modified by ���� 2004-01-07
        If str��Ժ��� = "" Then str��Ժ��� = Format(rsTemp!��Ժ����, "yyyy")
        If str��Ժ��� <> .��� Then
            .����סԺ = True '��Ӱ�����ߵ�ֵ���Լ��Ƿ�����סԺ����
            
            '�����ǿ���ĵ�һ�ν���
            gstrSQL = "Select ��� From ���ս����¼ Where ����=2 and ����=" & TYPE_������ & _
                " And ����ID=" & m����.����ID & " And ��ҳID=" & m����.��ҳID & " And ���=" & m����.���
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn����
            
            If rsTemp.EOF = True Then
                .������� = True  '����Ҫ���ۼƽ��ȫ�����
            End If
        End If
    End With
        
    '�˴�ʹ��װǮ��������ҪĿ���ǳ�ʼ�����˵Ŀ��ϵ����Լ��ۼƽ���ͳ���ͳ���ۼƱ���
    If װǮ����(m����.����ID) = False Then
        MsgBox "����װǮ����ʧ�ܣ��޷�׼ȷ�õ����˵�������ۼƱ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    With gIC����
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & m����.����ID & "," & TYPE_������ & "," & .MediYear & "," & _
            .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidFee & "," & _
            .PlanPaidAmt & "," & .InHosTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
    
    Call ҽ���Ҷ�(gIC����.CenterCode, gIC����.Cardno)
    
    If Calc����ͳ��() = False Then
        Exit Function
    End If
    
    If m����.�Ҷ� >= deg����֧�� Then
        With m����
            '������������Ҫ�ۼƵ���ֵ������
            .����ͳ��֧�� = .ͳ�����֧��
            '����ͳ����ò��ܴ��ڷⶥ��
            'takecare
            If m����.���÷ⶥ = True Then
                .����ͳ����� = .�������� - .ȫ�Է� - .�������ⶥ��
                If .����ͳ����� > (.�ⶥ�� - .ͳ����֧������) Then .����ͳ����� = (.�ⶥ�� - .ͳ����֧������)
                If .ʵ������ > .����ͳ����� Then .ʵ������ = .����ͳ�����
            Else
                .����ͳ����� = .ͳ�����֧�� + .ͳ������Ը�
            End If
        
            If Calc���䱨��() = False Then
                Exit Function
            End If
            
            If gIC����.IsOffical = 1 Then '����Ա�Ž��в�������
                If Calc��������() = False Then
                    Exit Function
                End If
            End If
            
            If m����.ȫ��ͳ�� = True Then
                סԺ�������_���� = "ҽ������;" & .����ͳ�� + .�����Ը� & ";0"
            Else
                סԺ�������_���� = "ҽ������;" & .ͳ�����֧�� & ";0"
                If .�������֧�� > 0 Then
                    סԺ�������_���� = סԺ�������_���� & "|�������;" & .�������֧�� & ";0"
                End If
                If .��������֧�� > 0 Then
                    סԺ�������_���� = סԺ�������_���� & "|��������;" & .��������֧�� & ";0"
                End If
            End If
        End With
    End If
'
    '����Ҫ���Ǹ����ʻ���֧����Χ
    '�����ⶥ��,����Calc����ͳ��()�м���������Ը���ʵ�ʵ������Ը�������,��Ҫ���¼���
    With m����
        If .�Ҷ� >= deg����֧�� Then
            Dim dbl�����Ը� As Double, dbl���Ը� As Double, dbl�����Ը� As Double '���Ը�=�����Ը�+�����Ը�+ȫ�Է�
            dbl���Ը� = .�������� - .ͳ�����֧�� - .�������֧��
            dbl�����Ը� = .����ͳ����� - .ͳ�����֧�� - .�������֧��
            '������Ը��к������Ը�����Ҫ���¼���
            'takecare
            If m����.���÷ⶥ Then
                If .����ͳ����� > .ҽ����Ŀ��� Then
                    dbl�����Ը� = (.����ͳ����� - .ҽ����Ŀ���) * 0.2
                Else
                    dbl�����Ը� = 0
                End If
                .�����Ը� = dbl�����Ը�
            End If
            
            If m����.���÷ⶥ = True Then
                dbl�����Ը� = dbl�����Ը� - .�����Ը�
            Else
                '֧���ⶥģʽ�£�����.����ͳ�����δ�������ߣ���ˣ�������Ҫ����ʵ������
                dbl�����Ը� = dbl�����Ը� + .ʵ������
            End If
            If dbl�����Ը� <= 0 Then dbl�����Ը� = 0
            
            .�����ʻ�֧�� = dbl�����Ը�
            If m����.�����ʻ�֧�������Ը� = True Then
                .�����ʻ�֧�� = .�����ʻ�֧�� + .�����Ը�
            End If
    
            If m����.�����ʻ�֧��ȫ�Է� = True Then
                .�����ʻ�֧�� = .�����ʻ�֧�� + .ȫ�Է�
            End If
     
            '����ʻ�����Ƿ��㹻֧��
            If m����.�����ʻ�֧�� > gIC����.InPerAcc - gIC����.OutPerAcc Then
                m����.�����ʻ�֧�� = gIC����.InPerAcc - gIC����.OutPerAcc
            End If
            If m����.�����ʻ�֧�� < 0 Then m����.�����ʻ�֧�� = 0
   
            סԺ�������_���� = סԺ�������_���� & IIf(סԺ�������_���� = "", "", "|") & "�����ʻ�;" & .�����ʻ�֧�� & ";1"
        End If
    End With
    
    If סԺ�������_���� = "" Then סԺ�������_���� = "�����ʻ�;0;1"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID     ���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim ic���� As TIC����               'סԺ��������ṹ
    Dim bln���� As Boolean
    Dim rs��Ϣ As New ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    Dim var������� As Variant, lng���� As Long, str�ֶ� As String
    
    On Error GoTo errHandle
    
    '����������д�����ս����¼�е���Ϣ
    gstrSQL = "SELECT A.����ID,A.NO,A.ʵ��Ʊ��,substr(B.����,1,8) as ����,substr(B.�Ա�,1,2) as �Ա�,floor(MONTHS_BETWEEN(A.�շ�ʱ��,B.��������)/12) AS ����" & _
              "         ,B.���֤��,C.����,C.ҽ����,A.�շ�ʱ��,substr(A.����Ա����,1,8) as ҽ��" & _
              "," & IIf(m����.��;���� = 1, "A.��ʼ����", "D.��Ժ����") & " AS ��Ժ����," & IIf(m����.��;���� = 1, "A.��������", "D.��Ժ����") & " AS ��Ժ���� " & _
              "  FROM ���˽��ʼ�¼ A,������Ϣ B,�����ʻ� C,������ҳ D" & _
              "  Where A.ID =[1] And A.����ID = B.����ID And B.����ID = C.����ID And C.���� =[2]" & _
             "         And B.����ID=D.����ID And D.��ҳID=[3]"
    Set rs��Ϣ = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID, TYPE_������, m����.��ҳID)
    'ֻҪ����;���㣬��Ҫ���1��
    m����.סԺ���� = Fix(CDate(Format(rs��Ϣ("��Ժ����"), "yyyy-MM-dd")) - _
                         CDate(Format(rs��Ϣ("��Ժ����"), "yyyy-MM-dd"))) + m����.��;����
    If m����.סԺ���� <= 0 Then m����.סԺ���� = 1
    
    If ReadIC(rs��Ϣ("����ID"), 1, True, "����ʱ����ʧ�ܡ�", ic����, bln����) = False Then
        Exit Function
    End If
    
    Call ҽ���Ҷ�(ic����.CenterCode, ic����.Cardno)
    
'    If m����.�Ҷ� = degֹ֧ͣ�� Then
'        '�����ٴ����������
'        סԺ����_���� = True
'        Exit Function
'    End If
        
    '������ʻ�֧�����
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    If Not rsTemp.EOF Then
        m����.�����ʻ�֧�� = rsTemp!���
    Else
        m����.�����ʻ�֧�� = 0
    End If
    
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    With m����
        'Ϊ�˱�֤��ȫ���ۼ����ݻ��Ƕ���������

        gstrSQL = "zl_�ʻ������Ϣ_insert(" & .����ID & "," & TYPE_������ & "," & .��� & "," & _
            ic����.InPerAcc & "," & ic����.OutPerAcc + .�����ʻ�֧�� & "," & ic����.PlanPaidFee + .����ͳ����� & "," & _
            ic����.PlanPaidAmt + .����ͳ��֧�� & "," & ic����.InHosTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & .����ID & "," & _
            .��� & "," & ic����.InPerAcc & "," & ic����.OutPerAcc & "," & ic����.PlanPaidFee & "," & _
            ic����.PlanPaidAmt & "," & ic����.InHosTimes & "," & .���� & "," & .�ⶥ�� & "," & .ʵ������ & "," & _
            .�������� & "," & .ȫ�Է� & "," & .�����Ը� & "," & .����ͳ����� & "," & .����ͳ��֧�� & ",0," & _
            .�������ⶥ�� & "," & .�����ʻ�֧�� & ",'" & ic����.OutSerialNo + 1 & "'," & .��ҳID & "," & .��;���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        For Each var������� In gcol�������
            '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
            gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
                var�������(0) & "," & var�������(1) & "," & var�������(2) & "," & var�������(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
            
            lng���� = lng���� + 1
            If lng���� <= 5 Then
                str�ֶ� = str�ֶ� & "," & var�������(2) & "," & IIf(m����.�Ҷ� < deg����֧��, 0, (var�������(1) - var�������(2)))
            End If
        Next
        '�������
        For lng���� = lng���� + 1 To 5
            str�ֶ� = str�ֶ� & ",0,0"
        Next
    End With
    
    'ҽ����������������Ȼ��������������һ�������У�������Ҫ��д����һ��
    With ic����
        gstrSQL = "INSERT INTO ���ս����¼ " & _
                "   (����,��¼id,����,���,���Ĵ���,���,����id,��ҳid,����,�Ա�,����,ҽ����,����,���֤��,��ݴ���,��λҽ���� " & _
                ",�Ƿ���Ա,�Ƿ�ҽ���չ˶���,�μӲ��䱣��,�ʻ��ۼ�����,�ʻ��ۼ�֧��,ͳ����֧�����,ͳ����֧������ " & _
                ",������֧�����,������֧������,�����𸶽���֧��,�������� " & _
                ",��������ʻ�֧�����,סԺ�����ʻ�֧�����,�����֧�����,��������,ҽ������,���ִ���,��������,�������� " & _
                ",סԺ����,����סԺ����,�������,��Ժ����,��Ժ����,סԺ���� " & _
                ",�������ý��,�����ʻ�֧��,ȫ�Ը����,�����Ը����,ת�������Ը�,����,�ⶥ��,ʵ������ " & _
                ",����ͳ��֧��,����ͳ�����,��������֧��,������������,ͳ����֧��,ͳ�����Ը�,ͳ�����֧��,ͳ������Ը� " & _
                ",�������֧��,��������Ը�,��������֧��,���������Ը� " & _
                ",��һ��֧��,��һ���Ը�,�ڶ���֧��,�ڶ����Ը�,������֧��,�������Ը�,���Ķ�֧��,���Ķ��Ը�,�����֧��,������Ը� " & _
                ",�������ⶥ��,������ⶥ��,������֧��,������������ʻ�֧��,����סԺ�����ʻ�֧��,�������Բ��𸶽� " & _
                ",���Ҷȼ�,��Ʊ��,Ʊ������,��Ʊ��־,����Ʊ�ݺ�,֧��˳���,��;����,�Ƿ��ϴ�) " & _
                  " Values "
         gstrSQL = gstrSQL & " (2," & lng����ID & "," & TYPE_������ & "," & .MediYear & ",'" & .CenterCode & "','" & rs��Ϣ("NO") & "1'," & m����.����ID & "," & m����.��ҳID & ",'" & rs��Ϣ("����") & _
                  "','" & rs��Ϣ("�Ա�") & "'," & rs��Ϣ("����") & ",'" & rs��Ϣ("ҽ����") & "','" & rs��Ϣ("����") & "','" & rs��Ϣ("���֤��") & "','" & .ClassCode & "','" & .UnitCode & "' " & _
                  "," & .IsOffical & "," & .IsAttend & "," & m����.�μӲ��䱣�� & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidAmt & "," & .PlanPaidFee & _
                  "," & .ChronicPaidAmt & "," & .ChronicPaidFee & "," & .ChronicSillPaidAmt & ",null" & _
                  "," & .ClinicPaidAmt & "," & .InHosPaidAmt & "," & .QuotaPaidAmt & ",'" & ToVarchar(UserInfo.����, 20) & "','" & rs��Ϣ("ҽ��") & "','" & m����.���ִ��� & "','" & m����.�������� & "','" & m����.�������� & "' " & _
                  "," & ic����.InHosTimes & "," & m����.סԺ�������� & ",'0'," & GetOracleFormat(rs��Ϣ("��Ժ����")) & "," & GetOracleFormat(rs��Ϣ("��Ժ����")) & "," & m����.סԺ���� & _
                  "," & m����.�������� & "," & m����.�����ʻ�֧�� & "," & m����.ȫ�Է� & "," & m����.�����Ը� & ",0," & m����.���� & "," & m����.�ⶥ�� & "," & m����.ʵ������ & " " & _
                  "," & m����.����ͳ��֧�� & "," & m����.����ͳ����� & "," & m����.��������֧�� & "," & m����.������������ & "," & _
                  (m����.ͳ�����֧�� + m����.�������֧�� + m����.��������֧��) & "," & (m����.ͳ������Ը� + m����.��������Ը� + m����.���������Ը�) & "," & m����.ͳ�����֧�� & "," & m����.ͳ������Ը� & " " & _
                  "," & m����.�������֧�� & "," & m����.��������Ը� & "," & m����.��������֧�� & "," & m����.���������Ը� & str�ֶ� & _
                  "," & m����.�������ⶥ�� & "," & m����.������ⶥ�� & "," & m����.������֧�� & ",0," & m����.�����ʻ�֧�� & "," & m����.�������Բ��𸶽� & " " & _
                  "," & m����.�Ҷ� & ",'" & Nvl(rs��Ϣ("ʵ��Ʊ��"), " ") & "'," & GetOracleFormat(rs��Ϣ("�շ�ʱ��")) & ",1,'','" & .OutSerialNo + 1 & "'," & m����.��;���� & ",0)"
        '׼��д��
        .OutPerAcc = .OutPerAcc + m����.�����ʻ�֧��                   '�����ʻ��ۼ�֧�����
        .InHosPaidAmt = .InHosPaidAmt + m����.�����ʻ�֧��             'סԺ�����ʻ�֧�����
        .InHosTimes = .InHosTimes + m����.סԺ��������                 'ֻ�г�Ժ���������סԺ����
        .PlanPaidFee = .PlanPaidFee + m����.����ͳ�����        'ͳ�����֧�������ۼƣ�����+���䣩
        .PlanPaidAmt = .PlanPaidAmt + m����.����ͳ��֧��        ' ͳ�����֧������ۼƣ�����+���䣩
        .ChronicPaidFee = .ChronicPaidFee + m����.������������                 '���Բ�֧�������ۼ�
        .ChronicPaidAmt = .ChronicPaidAmt + m����.��������֧��                 '���Բ�֧������ۼ�
        .QuotaPaidAmt = .QuotaPaidAmt + m����.������֧��                     '���Բ������֧�����
        .ChronicSillPaidAmt = .ChronicSillPaidAmt + m����.�������Բ��𸶽�     '���Բ��𸶽���֧�����
        .OutSerialNo = .OutSerialNo + 1           ' ֧��˳���
    End With
    '��¼סԺ�������һ������Ϣ����̫��Ҫ����ʹ����Ҳ���Ժ��ԣ������ܻع�ǰһ��д��
    Dim payLog As TPayInfo
    With payLog
        .HospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
        .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' ����
        .AccPay = m����.�����ʻ�֧��
        .Amount = m����.��������
        .CdFlag = 1
    End With
        
    '��ɿ�д��
    Dim str������ As String
    With m����
        str������ = ic����.CenterCode & "|" & gstrҽԺ���� & "|1|" & rs��Ϣ("NO") & "1|" & _
                    TrimStr(ic����.MediAccountNo) & "|" & m����.�����ʻ�֧�� & "|" & .ͳ�����֧�� & "|" & .�������֧�� & "|" & _
                    .����ͳ����� & "|" & .����ͳ��֧�� & "|" & .סԺ�������� & "|" & IIf(.�μӲ��䱣�� = 1, .������ⶥ��, .�������ⶥ��) & "|1"
    End With
    If WriteIC(bln����, True, 1, gstrSQL, ic����, payLog, str������) = False Then
        Exit Function
    End If
    
    סԺ����_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
'----------------------------------------------------------------
'���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
'������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
'      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
'      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
'----------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset, rs���� As New ADODB.Recordset, rs������� As New ADODB.Recordset
    Dim icסԺ As TIC����                'סԺ��������ṹ
    Dim lng����ID As Long
    Dim bln���� As Boolean
    Dim cur�����ʻ� As Currency
    
    On Error GoTo errHandle
    
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rs����("ID") '�������ݵ�ID
    rs����.Close
    
    gstrSQL = "Select *  From ���ս����¼ Where ��¼ID=" & lng����ID
    rs����.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    
    If rs����.RecordCount = 0 Then
        MsgBox "�ò��˵�ҽ���������ݶ�ʧ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If CanסԺ�������(rs����("����ID"), rs����("��ҳID")) = False Then Exit Function
    
    If ReadIC(rs����("����ID"), 1, True, "���Ͻ���ʱ����ʧ�ܡ�", icסԺ, bln����) = False Then
        Exit Function
    End If
    
    'ȡ�������
    gstrSQL = "Select ��� From ��������Ŀ¼ Where ����= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", icסԺ.CenterCode)
    m����.������� = rsTemp!���
    
    If Val(icסԺ.MediYear) <> rs����("���") Then
        Err.Raise 9000, gstrSysName, "���겻�����ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Get��ǰҽ����) <> rs����("���") Then
        Err.Raise 9000, gstrSysName, "���겻�����ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ҽ���Ҷ�(icסԺ.CenterCode, icסԺ.Cardno)
    
    If m����.�Ҷ� = degֹ֧ͣ�� Then
        '�����ٴ����������
        סԺ�������_���� = False
        Err.Raise 9000, gstrSysName, "�ò����Ѿ�ֹͣҽ��֧�������ܽ��г���������", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    '�жϸò��˵Ŀ��Ƿ������ȷ
    If ���IC��(rs����("����ID"), TrimStr(icסԺ.Cardno), TrimStr(icסԺ.CenterCode)) = False Then Exit Function
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rs����("����ID") & "," & TYPE_������ & "," & rs����("���") & "," & _
        icסԺ.InPerAcc & "," & icסԺ.OutPerAcc - rs����("�����ʻ�֧��") & "," & icסԺ.PlanPaidFee - rs����("����ͳ�����") & "," & _
        icסԺ.PlanPaidAmt - rs����("����ͳ��֧��") & "," & icסԺ.InHosTimes & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '�������ݻ������Ǹ���ԭ����
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & rs����("����ID") & "," & _
        rs����("���") & "," & icסԺ.InPerAcc & "," & icסԺ.OutPerAcc & "," & icסԺ.PlanPaidFee & "," & _
        icסԺ.PlanPaidAmt & "," & icסԺ.InHosTimes & "," & rs����("����") * -1 & "," & rs����("�ⶥ��") & "," & rs����("ʵ������") * -1 & "," & _
        rs����("�������ý��") * -1 & "," & rs����("ȫ�Ը����") * -1 & "," & rs����("�����Ը����") * -1 & "," & rs����("����ͳ�����") * -1 & "," & _
        rs����("����ͳ��֧��") * -1 & ",0," & rs����("�������ⶥ��") * -1 & "," & rs����("�����ʻ�֧��") * -1 & ",'" & icסԺ.OutSerialNo + 1 & "'," & _
        IIf(IsNull(rs����("��ҳID")), "null", rs����("��ҳID")) & "," & rs����("��;����") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    cur�����ʻ� = rs����("�����ʻ�֧��")
    
    gstrSQL = "select ����,����ͳ����,ͳ�ﱨ�����,���� from ���ս������ where ����ID=[1]"
    Set rs������� = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    Do Until rs�������.EOF
        '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
        gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
            rs�������("����") & "," & rs�������("����ͳ����") * -1 & "," & rs�������("ͳ�ﱨ�����") * -1 & "," & rs�������("����") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        rs�������.MoveNext
    Loop
    
    'ҽ����������������Ȼ��������������һ�������У�������Ҫ��д����һ��
    With icסԺ
        gstrSQL = "INSERT INTO ���ս����¼ " & _
                "   (����,��¼id,����,���,���Ĵ���,���,����id,��ҳid,����,�Ա�,����,ҽ����,����,���֤��,��ݴ���,��λҽ���� " & _
                ",�Ƿ���Ա,�Ƿ�ҽ���չ˶���,�μӲ��䱣��,�ʻ��ۼ�����,�ʻ��ۼ�֧��,ͳ����֧�����,ͳ����֧������ " & _
                ",������֧�����,������֧������,�����𸶽���֧��,�������� " & _
                ",��������ʻ�֧�����,סԺ�����ʻ�֧�����,�����֧�����,��������,ҽ������,���ִ���,��������,�������� " & _
                ",סԺ����,����סԺ����,�������,��Ժ����,��Ժ����,סԺ���� " & _
                ",�������ý��,�����ʻ�֧��,ȫ�Ը����,�����Ը����,ת�������Ը�,����,�ⶥ��,ʵ������ " & _
                ",����ͳ��֧��,����ͳ�����,��������֧��,������������,ͳ����֧��,ͳ�����Ը�,ͳ�����֧��,ͳ������Ը� " & _
                ",�������֧��,��������Ը�,��������֧��,���������Ը� " & _
                ",��һ��֧��,��һ���Ը�,�ڶ���֧��,�ڶ����Ը�,������֧��,�������Ը�,���Ķ�֧��,���Ķ��Ը�,�����֧��,������Ը� " & _
                ",�������ⶥ��,������ⶥ��,������֧��,������������ʻ�֧��,����סԺ�����ʻ�֧��,�������Բ��𸶽� " & _
                ",���Ҷȼ�,��Ʊ��,Ʊ������,��Ʊ��־,����Ʊ�ݺ�,֧��˳���,��;����,�Ƿ��ϴ�) " & _
                  " Values "
         gstrSQL = gstrSQL & " (2," & lng����ID & "," & TYPE_������ & "," & .MediYear & ",'" & .CenterCode & "','" & Mid(rs����("���"), 1, Len(rs����("���")) - 1) & "2'," & rs����("����ID") & "," & rs����("��ҳID") & ",'" & rs����("����") & _
                  "','" & rs����("�Ա�") & "'," & rs����("����") & ",'" & rs����("ҽ����") & "','" & rs����("����") & "','" & rs����("���֤��") & "','" & .ClassCode & "','" & .UnitCode & "' " & _
                  "," & .IsOffical & "," & .IsAttend & "," & rs����("�μӲ��䱣��") & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidAmt & "," & .PlanPaidFee & _
                  "," & .ChronicPaidAmt & "," & .ChronicPaidFee & "," & .ChronicSillPaidAmt & ",null" & _
                  "," & .ClinicPaidAmt & "," & .InHosPaidAmt & "," & .QuotaPaidAmt & ",'" & rs����("��������") & "','" & rs����("ҽ������") & "','" & rs����("���ִ���") & "','" & rs����("��������") & "','" & rs����("��������") & "' " & _
                  "," & .InHosTimes & "," & rs����("����סԺ����") & ",'0'," & GetOracleFormat(rs����("��Ժ����")) & "," & GetOracleFormat(rs����("��Ժ����")) & "," & rs����("סԺ����") & _
                  "," & rs����("�������ý��") & "," & rs����("�����ʻ�֧��") & "," & rs����("ȫ�Ը����") & "," & rs����("�����Ը����") & ",0," & rs����("����") & "," & rs����("�ⶥ��") & "," & rs����("ʵ������") & " " & _
                  "," & rs����("����ͳ��֧��") & "," & rs����("����ͳ�����") & "," & rs����("��������֧��") & "," & rs����("������������") & "," & rs����("ͳ����֧��") & "," & rs����("ͳ�����Ը�") & "," & rs����("ͳ�����֧��") & "," & rs����("ͳ������Ը�") & " " & _
                  "," & rs����("�������֧��") & "," & rs����("��������Ը�") & "," & rs����("��������֧��") & "," & rs����("���������Ը�") & _
                  "," & rs����("��һ��֧��") & "," & rs����("��һ���Ը�") & "," & rs����("�ڶ���֧��") & "," & rs����("�ڶ����Ը�") & "," & rs����("������֧��") & _
                  "," & rs����("�������Ը�") & "," & rs����("���Ķ�֧��") & "," & rs����("���Ķ��Ը�") & "," & rs����("�����֧��") & "," & rs����("������Ը�") & " " & _
                  "," & rs����("�������ⶥ��") & "," & rs����("������ⶥ��") & "," & rs����("������֧��") & "," & rs����("������������ʻ�֧��") & "," & rs����("����סԺ�����ʻ�֧��") & "," & rs����("�������Բ��𸶽�") & " " & _
                  "," & m����.�Ҷ� & ",'" & Nvl(rs����("��Ʊ��"), " ") & "',sysdate,-1,'" & rs����("���") & "','" & .OutSerialNo + 1 & "'," & rs����("��;����") & ",0)"
        '׼��д��
        .OutPerAcc = .OutPerAcc - cur�����ʻ�                  '�����ʻ��ۼ�֧�����
        .InHosPaidAmt = .InHosPaidAmt - cur�����ʻ�            '��������ʻ�֧�����
        .InHosTimes = .InHosTimes - rs����("����סԺ����")      '��Щ���ز�������סԺ����
        .PlanPaidFee = .PlanPaidFee - rs����("����ͳ�����")      'ͳ�����֧�������ۼƣ�����+���䣩
        .PlanPaidAmt = .PlanPaidAmt - rs����("����ͳ��֧��")        ' ͳ�����֧������ۼƣ�����+���䣩
        .ChronicPaidFee = .ChronicPaidFee - rs����("������������")                '���Բ�֧�������ۼ�
        .ChronicPaidAmt = .ChronicPaidAmt - rs����("��������֧��")                '���Բ�֧������ۼ�
        .QuotaPaidAmt = .QuotaPaidAmt - rs����("������֧��")                     '���Բ������֧�����
        .ChronicSillPaidAmt = .ChronicSillPaidAmt - rs����("�������Բ��𸶽�")      '���Բ��𸶽���֧�����
        .OutSerialNo = .OutSerialNo + 1           ' ֧��˳���
    End With
        
    '��¼סԺ���
    Dim payLog As TPayInfo
    With payLog
        .HospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
        .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' ����
        .AccPay = cur�����ʻ�
        .Amount = rs����("�������ý��")
        .CdFlag = 0
    End With
        
    '��ɿ�д��
    Dim str������ As String
        
    str������ = icסԺ.CenterCode & "|" & gstrҽԺ���� & "|1|" & Mid(rs����("���"), 1, Len(rs����("���")) - 1) & "2|" & _
                TrimStr(icסԺ.MediAccountNo) & "|" & cur�����ʻ� & "|" & rs����("ͳ�����֧��") & "|" & rs����("�������֧��") & "|" & _
                rs����("����ͳ�����") & "|" & rs����("����ͳ��֧��") & "|" & rs����("����סԺ����") & "|" & IIf(rs����("�μӲ��䱣��") = 1, rs����("������ⶥ��"), rs����("�������ⶥ��")) & "|-1"
    
    If WriteIC(bln����, True, 1, gstrSQL, icסԺ, payLog, str������) = False Then
        Exit Function
    End If
            
    סԺ�������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ���ʴ���_����(ByVal str���ݺ� As String, ByVal int���� As Integer, Optional ByVal lng����ID As Long = 0) As Boolean
'����:�ϴ��²����ļ�����ϸ��ҽ������
'����:  str���ݺ�   NO
'       int����     ��¼����
'       str��Ϣ    �����������������ѣ�����ǰ̨������ɣ����ⳤʱ���������
'       lng����ID  Ĭ��Ϊ0����ʾ�������ŵ��ݣ�����Ϊ������ָ�����˵ġ�����Ҫ����Ϊҽ���ڱ�����ʵ�ʱ���Ƿֲ������ύ���ݶ�����һ���ύ��
'����:
    Dim rsTemp As New ADODB.Recordset
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency, curͳ���� As Currency
    
    '��ע�⣺����ҽ�����ڼ��ʵ�������ٵ��ô�����̵ġ�
    
    On Error GoTo errHandle
    
    '�������ŵ��ݵķ�����ϸ
    
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.��¼״̬,A.�շ�ϸĿID,C.��Ŀ����,B.����,B.����,A.ʵ�ս�� " & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as ���� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,������ҳ D,������Ϣ E " & _
              "  where A.NO=[1] and A.��¼����=[2] and A.��¼״̬=1 And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= [3]" & _
              "        and A.����ID=D.����ID And A.����ID=E.����ID And D.��ҳID=E.סԺ���� and D.����=[3]" & _
              "  Order by A.����ID,A.����ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ʴ���", str���ݺ�, int����, TYPE_������)
    
    If Calc���÷ָ�(rsTemp, True, curȫ�Է�, cur�����Ը�, curͳ����) = False Then
        Exit Function
    End If
        
    ���ʴ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��������_����(ByVal str���ݺ� As String, ByVal int���� As Integer, ����ID As Long) As Boolean
'����:�����Ѿ��ϴ���ҽ�����ĵļ�����ϸ
'����:  str���ݺ�   NO
'       int����     ��¼����
'       str��Ϣ    �����������������ѣ�����ǰ̨������ɣ����ⳤʱ���������
'����:
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, arrOutput As Variant
    Dim lng�ϴ���־ As Long
    
    On Error GoTo errHandle
    
    '�������ŵ��ݵķ�����ϸ����δ�ϴ��ļ�¼��ȡԭʼ���ݣ�
'    gstrSQL = "Select distinct nvl(A.�Ƿ��ϴ�,0) �ϴ���־ " & _
'              "  From ���˷��ü�¼ A" & _
'              "  where A.NO='" & str���ݺ� & "' and A.��¼����=" & int���� & " and A.��¼״̬<>2 and nvl(A.ʵ�ս��,0)<>0 "
'    Call OpenRecordset(rsTemp, "��������")
'
'    If rsTemp.RecordCount > 1 Then
'        MsgBox "�õ�����ķ�����ϸ��δȫ����ɷ��÷ָ", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    If rsTemp("�ϴ���־") <> 0 Then
'        '�Ѿ���ɷ��÷ָ�����ϴ������ݣ����ϵ�����Ҫ��ԭʼ���ݵķָ�����ͬ
'        lng�ϴ���־ = rsTemp("�ϴ���־")
'        gstrSQL = "Select ID " & _
'                  "  From ���˷��ü�¼ A" & _
'                  "  where A.NO='" & str���ݺ� & "' and A.��¼����=" & int���� & " and A.��¼״̬=2 and nvl(A.ʵ�ս��,0)<>0 "
'        Call OpenRecordset(rsTemp, "��������")
'
'        Do Until rsTemp.EOF
'            '�������˵ĵ��ݸ�Ϊ�Ѿ�����˷��÷ָ��״̬
'            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rsTemp("ID") & ",null,null,null,null,2)"
'            gcnOracle.Execute gstrSQL, , adCmdStoredProc
'
'            rsTemp.MoveNext
'        Loop
'    End If
    
    ��������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ������Ϣ_����(ByVal lngErrCode As Long) As String
'���ܣ����ݴ���ŷ��ش�����Ϣ
    Select Case lngErrCode
        Case -2
            ������Ϣ_���� = "������������"
        Case -3
            ������Ϣ_���� = "�����˿�ʧ�ܡ�"
        Case -4
            ������Ϣ_���� = "�򿪶�����ʧ��,������������Ӻ͵�Դ��"
        Case -5
            ������Ϣ_���� = "�޿���"
        Case 0
            ������Ϣ_���� = "��ȷ��"
        Case 2
            ������Ϣ_���� = "������"
        Case 3
            ������Ϣ_���� = "�ļ�������"
        Case 4
            ������Ϣ_���� = "����PIN��"
'        Case 5
'            ������Ϣ_���� = "��"
        Case 6
            ������Ϣ_���� = "��λʧ�ܡ�"
        Case 7
            ������Ϣ_���� = "�������"
        Case 8
            ������Ϣ_���� = "�޸�����ʧ�ܡ�"
        Case 9
            ������Ϣ_���� = "����ȴ���"
        Case 10
            ������Ϣ_���� = "״̬����"
        Case 11
            ������Ϣ_���� = "�ļ�������"
        Case 12
            ������Ϣ_���� = "�ļ�δѡ��"
        Case 13
            ������Ϣ_���� = "�������á�"
        Case 14
            ������Ϣ_���� = "�ļ��Ѿ����ڡ�"
        Case 15
            ������Ϣ_���� = "�����P1/P2��"
        Case 16
            ������Ϣ_���� = "��������"
        Case 17
            ������Ϣ_���� = "�����P2��"
        Case 18
            ������Ϣ_���� = "�ļ�û���ҵ���"
        Case 19
            ������Ϣ_���� = "�ļ����㹻�ռ䡣"
        Case 20
            ������Ϣ_���� = "��������"
        Case 21
            ������Ϣ_���� = "ƫ��������"
        Case 22
            ������Ϣ_���� = "ָ�������Ч��"
        Case 23
            ������Ϣ_���� = "��Ч��CLA��"
        Case 24
            ������Ϣ_���� = "��������"
        Case 25
            ������Ϣ_���� = "д������ת������"
        Case 26
            ������Ϣ_���� = "�����ʻ����ָ���,��ҽ�����Ĵ���"
        Case 33
            ������Ϣ_���� = "IC���Ѿ����Ƿ�����,д��ʧ�ܡ�"
        Case 100
            ������Ϣ_���� = "һ�ڿ�����Ҫ��ʽת����"
        Case 101
            ������Ϣ_���� = "�Ǳ�ϵͳ����"
        Case 210
            ������Ϣ_���� = "д��ʧ�ܡ�"
        Case 211
            ������Ϣ_���� = "д��ʧ��,�ۿ���ҽ�����Ĵ���"
        Case 300
            ������Ϣ_���� = "CRCУ�����"
        Case 301
            ������Ϣ_���� = "IC���Ѿ����Ƿ�����,д��ʧ��.��"
        Case 600
            ������Ϣ_���� = "����ֵת������"
        Case Else
            ������Ϣ_���� = "����ʶ��Ĵ���"
    End Select
End Function

Private Function װǮ����(ByVal lng����ID As Long) As Boolean
'���ܣ����ȶ϶��Ƿ�ҪװǮ��Ȼ�������Ӧ����
    Dim rsTemp As New ADODB.Recordset
    
    Dim strװǮģʽ As String, blnǿ��װǮ As Boolean, blnԶ����֤ As Boolean, strԶ�̵�ַ As String
    Dim strҽ����  As String, lngװǮ�ڴ� As Long
    Dim dbl�ۼ�ע�� As Double
    Dim ic�� As TIC����
    Dim strҽ����_IC  As String, lngװǮ�ڴ�_IC As Long
    Dim dbl�ۼ�ע��_IC As Double
    Dim lngTemp As Long, bln���� As Boolean
    
    Dim str����ֵ As String
    
    On Error GoTo errHandle
    
    If Get���ղ���_����(blnԶ����֤, strԶ�̵�ַ) = False Then
        Exit Function
    End If
    
    If blnԶ����֤ = True Then
        װǮ���� = True
        Exit Function
    End If
    
    '�õ����µ�IC����Ϣ
    'ʹ�ñ��صģ���Ϊ���ܽ��и��ĵ��ֲ��ɹ�
    If ReadIC(lng����ID, 1, True, "װǮʱ����ʧ�ܡ�", gIC����, bln����) = False Then
        Exit Function
    End If
    If bln���� = True Then
        '������Ա��װǮ
        װǮ���� = True
        Exit Function
    End If
    
    ic�� = gIC����
    
    With ic��
        strҽ����_IC = .MediYear
        lngװǮ�ڴ�_IC = .InNo
        dbl�ۼ�ע��_IC = .InPerAcc
    End With
    
    '���װǮģʽ
    '���кϷ�����֤
    gstrSQL = "SELECT B.ҽ����,B.װǮ���,B.װǮģʽ " & _
               " FROM ��������Ŀ¼ A,�������� B " & _
               " WHERE A.����=" & TYPE_������ & " AND A.����='" & ic��.CenterCode & "' AND A.��������=B.���� AND A.����=B.���� "
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        strװǮģʽ = Nvl(rsTemp("װǮģʽ"))
        strҽ���� = Nvl(rsTemp("ҽ����"))
        lngװǮ�ڴ� = Nvl(rsTemp("װǮ���"), 0)
    End If
    If strװǮģʽ = "" Or strҽ���� = "" Then
        MsgBox "���������Ա���ҽ�����ݵ����ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If strװǮģʽ = "1" Then
'        If strҽ���� > strҽ����_IC Then
'            Call ����ҽ����װǮ(ic��, strҽ����, lngװǮ�ڴ�, ic��.InPerAcc - ic��.OutPerAcc)
'
'            '����Ϣд�ؿ���
'            If ��¼װǮ��־(ic��, strҽ����_IC, lngװǮ�ڴ�_IC, dbl�ۼ�ע��_IC) = True Then
'                '����ȫ�ֱ�������������
'                gIC���� = ic��
'                װǮ���� = True
'                Exit Function
'            Else
'                'װǮʧ��
'                Exit Function
'            End If
'        Else
'            lngTemp = OnLineInMoney(ic��.CenterCode, ic��.Cardno, strҽ����_IC, Trim(gstrҽԺ����), serverIP)
'            If lngTemp <> 0 Then
'                Exit Function
'            Else
'                'װǮ�ɹ����ӿ��ж����µ�ֵ
'                If ReadICCard(gIC����) <> 0 Then
'                    װǮ���� = True
'                    Exit Function
'                End If
'            End If
'        End If
        '����װǮ
        Dim serverIP As String
        serverIP = Get����IP
        lngTemp = OnLineInMoney(ic��.CenterCode, ic��.Cardno, strҽ����_IC, Trim(gstrҽԺ����), serverIP)
        If lngTemp <> 0 Then
            'װǮ���ɹ�
            '�ж��Ƿ��и���ҽ����
            If strҽ���� > strҽ����_IC Then
                MsgBox "װǮ�嵥��û�д˿�����Ϣ���뵽���Ĵ���", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            'װǮ�ɹ����ӿ��ж����µ�ֵ
            If ReadIC(lng����ID, 1, False, "װǮʱ����ʧ�ܡ�", gIC����, bln����) = True Then
                װǮ���� = True
                Exit Function
            End If
        End If
    End If
    
    If strװǮģʽ = "0" Then
        '��װǮ
        If ic��.MediYear = "2001" And ic��.InNo = 0 Then
            'ǿ������װǮģʽ
            blnǿ��װǮ = True
        Else
            '�ж��Ƿ��и���ҽ����
            If strҽ���� > ic��.MediYear Then
                Call ����ҽ����װǮ(ic��, strҽ����, lngװǮ�ڴ�, ic��.InPerAcc - ic��.OutPerAcc)
                If ��¼װǮ��־(ic��, strҽ����_IC, lngװǮ�ڴ�_IC, dbl�ۼ�ע��_IC) = True Then
                    '����ȫ�ֱ�������������
                    gIC���� = ic��
                Else
                    'װǮʧ��
                    Exit Function
                End If
            End If
        End If
        
    End If
    
    If (strװǮģʽ = "2" Or blnǿ��װǮ = True) And lngװǮ�ڴ� > ic��.InNo Then
        '����װǮ
        If ���ҽ��������_���� = False Then
            '�������ӵ�ǰ�÷�����������Ϊ����ʹ��
            Exit Function
        End If
        
        '�õ�װǮ�嵥
        With ic��
            gstrSQL = "select �ʻ�ע�� from װǮ�嵥 " & _
                     "where ���Ĵ���='" & .CenterCode & "' and ����='" & .Cardno & "' and װǮ�ڴ�=" & lngװǮ�ڴ�
                     '"where ���Ĵ���='" & .CenterCode & "' and ����='" & .Cardno & "' and ҽ����='" & strҽ���� & "' and װǮ�ڴ�=" & lngװǮ�ڴ�
        End With
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn����, adOpenStatic
        If rsTemp.RecordCount = 0 Then
            '�ж��Ƿ��и���ҽ����
            If strҽ���� > ic��.MediYear Then
                MsgBox "װǮ�嵥��û�д˿�����Ϣ���뵽���Ĵ���", vbInformation, gstrSysName
                Exit Function
'                Call ����ҽ����װǮ(ic��, strҽ����, lngװǮ�ڴ�, ic��.InPerAcc - ic��.OutPerAcc)
'                If ��¼װǮ��־(ic��, strҽ����_IC, lngװǮ�ڴ�_IC, dbl�ۼ�ע��_IC) = True Then
'                    '����ȫ�ֱ�������������
'                    gIC���� = ic��
'                    װǮ���� = True
'                End If
            Else
                MsgBox "װǮ�嵥��û�д˿�����Ϣ���뵽���Ĵ���", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        
        'ע�⣺�˴�Ӧ�ø�Ϊ���ܺ�õ����
        dbl�ۼ�ע�� = Val(EncryptStr(IIf(IsNull(rsTemp("�ʻ�ע��")), "", rsTemp("�ʻ�ע��")), "256", False))
        If strҽ���� > ic��.MediYear Then
            '����ҽ����װǮ
            Call ����ҽ����װǮ(ic��, strҽ����, lngװǮ�ڴ�, dbl�ۼ�ע��)
        Else
            '����ҽ����װǮ
            With ic��
                .InNo = lngװǮ�ڴ�
                .InPerAcc = dbl�ۼ�ע��
                .OutSerialNo = .OutSerialNo + 1
            End With
        End If
        If ��¼װǮ��־(ic��, strҽ����_IC, lngװǮ�ڴ�_IC, dbl�ۼ�ע��_IC) = True Then
            '����ȫ�ֱ�������������
            gIC���� = ic��
        Else
            'װǮʧ��
            Exit Function
        End If
    End If
    
    װǮ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ����ҽ����װǮ(ic���� As TIC����, ByVal strҽ���� As String, ByVal lngװǮ�ڴ� As Long, ByVal dbl�ۼ�ע�� As Double)
    With ic����
        .MediYear = strҽ����
        .InNo = lngװǮ�ڴ�
        .InPerAcc = dbl�ۼ�ע��
        .OutPerAcc = 0
        .PlanPaidAmt = 0
        .PlanPaidFee = 0
        .ChronicPaidAmt = 0
        .ChronicPaidFee = 0
        .InHosTimes = 0
        .QuotaPaidAmt = 0
        .InHosPaidAmt = 0
        .ClinicPaidAmt = 0
        .ChronicSillPaidAmt = 0
        .OutSerialNo = .OutSerialNo + 1
    End With
End Sub

Private Function ��¼װǮ��־(ic���� As TIC����, ByVal IC_MediYear As String, ByVal IC_InNo As Long, ByVal IC_InPerAcc As Double) As Boolean
    
    If ���ҽ��������_���� = False Then
        '�������ӵ�ǰ�÷�����������Ϊ����ʹ��
        Exit Function
    End If
    
    gcn����.BeginTrans
    On Error Resume Next
    
    '���ȱ���װǮ��־
    With ic����
        gstrSQL = "insert into װǮ��־ (���Ĵ���,����,����ҽ����,����װǮ�ڴ�,�����˻�ע��" & _
            ",����ҽ����,����װǮ�ڴ�,�����˻�ע��,��������) values ('" & _
            .CenterCode & "','" & .Cardno & "','" & IC_MediYear & "'," & IC_InNo & "," & Format(IC_InPerAcc, "#####0.00") & ",'" & _
            .MediYear & "'," & .InNo & "," & Format(.InPerAcc, "#####0.00") & ",sysdate)"
        
    End With
    gcn����.Execute gstrSQL
    If Err <> 0 Then
        gcn����.RollbackTrans
        Err.Clear
        Exit Function
    End If
    
    '���д������
    If WriteICCard(ic����) <> 0 Then
        gcn����.RollbackTrans
        MsgBox "IC��װǮ����ʧ�ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Err <> 0 Then '�п���д��ʱ����ʵʱ����
        gcn����.RollbackTrans
        Err.Clear
        Exit Function
    End If
    
    gcn����.CommitTrans
    ��¼װǮ��־ = True
End Function

Private Sub ҽ���Ҷ�(ByVal str���� As String, ByVal str���� As String)
'����ָ���û���ҽ���Ҷȼ�
    Dim rsTemp As New ADODB.Recordset
    
    If ���ҽ��������_���� = False Then
        '�������ӵ�ǰ�÷�����������Ϊ����ʹ��
        m����.�Ҷ� = degֹ֧ͣ��
        Exit Sub
    End If
    
    gstrSQL = "select �Ҷ� from ������ where ���Ĵ���='" & str���� & "' and ����='" & str���� & "'"
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    
    If rsTemp.RecordCount > 0 Then
        '���ûҶ�ֵ
        m����.�Ҷ� = Val(rsTemp("�Ҷ�"))
    Else
        '�����Ĳ��·�
        m����.�Ҷ� = deg����֧��
    End If
    
End Sub

Private Function ���IC��(ByVal lng����ID As Long, ByVal str���� As String, ByVal str���� As String) As Boolean
'���ܣ��жϸò��˵Ŀ��Ƿ������ȷ
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.����,A.ҽ����,B.���� from �����ʻ� A,��������Ŀ¼ B " & _
              " where A.����=[1] and A.����ID=[2] and a.����=B.���� and A.����=B.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, lng����ID)
    
    If rsTemp("����") <> str���� Or rsTemp("����") <> str���� Then
        MsgBox "ˢ�����еĿ����ǵ�ǰ���˵ģ��������ȷ��IC����", vbInformation, gstrSysName
        Exit Function
    End If
    
    ���IC�� = True
End Function

Public Function ���ҽ��������_����() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If gcn����.State = adStateOpen Then
        ���ҽ��������_���� = True
        Exit Function
    End If
    
    '��������ҽ��������������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                '����
                If strPass <> "" Then strPass = EncryptStr(strPass, 256, False)
        End Select
        rsTemp.MoveNext
    Loop
    
    If OraDataOpen(gcn����, strServer, strUser, strPass, False) = True Then
        ���ҽ��������_���� = True
        Exit Function
    End If
        
    MsgBox "ҽ��ǰ�÷���������ʧ�ܡ�", vbInformation, gstrSysName
End Function

Public Function Get���ݲ���_����(ByVal strIdentify As String, ic���� As TIC����, Optional ByVal blnҽ���� As Boolean = True) As Boolean
'���ܣ��������嵥�ж�ȡ�������������IC���ṹ��
'������strIdentify     ���������֤��blnҽ����=False Ϊ���֤ ��blnҽ����=True ��ҽ���ţ�
'      IC����        ���ݶ�������Ϣ��дIC���ṹ
    Dim rsTemp As New ADODB.Recordset

    If ���ҽ��������_���� = False Then
        Exit Function
    End If
    
    gstrSQL = "select * from ������Ա where " & IIf(blnҽ���� = True, "ҽ����", "���֤��") & _
                "='" & strIdentify & "'"
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        'û�ҵ������ݲ��˵ļ�¼
        Exit Function
    End If
    
    With ic����
        .CenterCode = rsTemp("���Ĵ���")     'As String * 4      ' ���Ĵ���
        .Cardno = rsTemp("ҽ����")           'As String * 8      ' ����
        .IDCardno = rsTemp("���֤��")       'As String * 18     ' ���֤�� ���Ȳ����#0
        .MediAccountNo = rsTemp("ҽ����")    'As String * 8      ' ҽ����
        .Name = rsTemp("����")               'As String * 10     ' ����
        .Sex = IIf(IsNull(rsTemp("�Ա�")), "1", rsTemp("�Ա�"))       'As String * 1      ' �Ա� 1-��  0-Ů
        .Birthday = rsTemp("����")           'As String * 8      ' �������� YYYYMMDD
        .UnitCode = rsTemp("��λҽ����")     'As String * 5      ' ���˵�λ����
        .ClassCode = rsTemp("��ݴ���")      'As String * 2      ' ְ����ݣ�0x����ְ1x������, 05��11Ϊһ���Խɷ�
        .DomainCode = 0     'As String * 1      ' ְ������ 0-���� 1-��פ��� 2-��ذ���
        .MediYear = Year(zlDatabase.Currentdate)          'As String * 4      ' ҽ�����
        .InNo = 0           'As Long            ' װǮ�ڴ�
        .OutSerialNo = 0    'As Long            ' ֧��˳���
        .InPerAcc = 0       'As Double          ' �����ʻ��ۼ�ע����
        .OutPerAcc = 0      'As Double          ' �����ʻ��ۼ�֧�����
        .PlanPaidAmt = 0     'As Double          ' ����ͳ��֧������ۼ�
        .PlanPaidFee = 0 'As Double          ' �������ͳ�����ۼ�
        .ChronicPaidFee = 0 '   As Double          ' ���Բ�֧�������ۼ�
        .ChronicPaidAmt = 0 '   As Double          ' ���Բ�֧������ۼ�
        .InHosPaidAmt = 0 '     As Double          ' סԺ�����ʻ�֧�����
        .ClinicPaidAmt = 0 '    As Double          ' ��������ʻ�֧�����
        .QuotaPaidAmt = 0 '     As Double          ' ���Բ������֧�����
        .ChronicSillPaidAmt = 0 '    As Double     ' ���Բ��𸶽���֧�����
        .IsOffical = "0" '        As String * 1      ' ����Ա 0-������-��
        .IsAttend = "0" '       As String * 1      ' ҽ���չ˶��� 0-��1-��
        .Password = "9000"       'As String * 4      ' ��������
        .InHosTimes = 0 'As Long           ' ������ЧסԺ����
        .InpatientFlag = 0  'As String * 1      ' סԺ��־ 0-��סԺ 1-סԺ
    End With
    
    Get���ݲ���_���� = True
End Function


Private Function Is���ݲ���(ByVal lng����ID As Long) As Boolean
'���ܣ������ʻ���Ϣ�жϲ����Ƿ����ݲ���
'���������ز��˵�ҽ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.��ְ from �����ʻ� A where A.����=[1] and A.����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, lng����ID)
    
    If rsTemp.EOF = True Then
        '�ò���û����
        Is���ݲ��� = False
    Else
        Is���ݲ��� = IIf(rsTemp("��ְ") = 3, True, False)
    End If
End Function

Private Function Get�ʻ���Ϣ(ByVal lng����ID As Long, strҽ���� As String, str���֤�� As String, str���� As String) As Boolean
'���ܣ������ʻ���Ϣ�жϲ����Ƿ����ݲ���
'���������ز��˵�ҽ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.ҽ����,A.����,B.���֤�� from �����ʻ� A,������Ϣ B where A.����=[1]" & _
        " and A.����ID=[2] And A.����ID=B.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, lng����ID)
    
    If rsTemp.EOF = False Then
        '�ò��˷���
        strҽ���� = Nvl(rsTemp("ҽ����"))
        str���֤�� = Nvl(rsTemp("���֤��"))
        str���� = Nvl(rsTemp("����"))
        Get�ʻ���Ϣ = True
    Else
        MsgBox "�޷���ȡ�ʻ���Ϣ��", vbInformation, gstrSysName
    End If
End Function

'Modified By ���� 2003-12-10 ���������� ���Ӳ���
Private Function Calc���÷ָ�(rs������ϸ As ADODB.Recordset, ByVal �Ƿ���� As Boolean _
                , curȫ�Է� As Currency, cur�����Ը� As Currency, curͳ�� As Currency, _
                Optional ByVal ���÷ָ� As Boolean = False, Optional ByVal bln���� As Boolean = False) As Boolean
'���ܣ����ݷ�����ϸ�����¼�����ϸ�з��õı���������õĽ�����ֱ���ϴ�
'������rs������ϸ  ������ϸ���������õ�ϸĿID�����ۡ����������
'      �Ƿ����     �Ƿ���Ҫ�����ݿ��в��˷��ü�¼��ҽ�����ݽ��и��¡�����Ԥ��ʱ������
'      curȫ�Է�    ���������������ȫ�ԷѲ��ֵĽ��
'      cur�����Ը�  ��������������������Ը����ֵĽ��
'      curͳ��      ���������������ͳ�ﲿ�ֵĽ��
'      ���÷ָ�     ���������Ϊ���ʾ�޼۴Ӳ��˷��ü�¼�ж�ȡ�������㵱ǰ�Ǳʼ�¼
'���أ��������ɹ�������й��ܣ�ΪTrue
'����λ�ã�����Ԥ�㡢������㡢סԺ���ʡ�סԺԤ�㡢סԺ���㡢������ϸ�ϴ�

    Dim str���ı��� As String, str���ֱ��� As String, lng����ID As Long
    Dim rs���մ��� As New ADODB.Recordset
    Dim rs������׼ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset, str��Ŀ���� As String, strϸĿ���� As String
    Dim cur��� As Currency, curʵ�ʵ��� As Currency, cur���۸� As Currency, cur���� As Currency, cur�Ը����� As Currency, cur��λ�� As Currency, cur������Ŀ As Currency
    Dim curͳ���� As Currency, cur�Ը� As Currency, lng���մ���ID As Long, lng������Ŀ�� As Long
    Dim blnҽ������ As Boolean, blnҽ����Ŀ As Boolean, bln���� As Boolean, bln���� As Boolean
    
    If ���ҽ��������_���� = False Then
        Exit Function
    End If
    curȫ�Է� = 0
    cur�����Ը� = 0
    curͳ�� = 0
    
    On Error GoTo errHandle
    '�õ�����ҽ������
    gstrSQL = "SELECT A.ID,A.���� FROM ����֧������ A Where A.���� =" & TYPE_������
    rs���մ���.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    'Modified by zyb ##2003-08-31
    If Not ���÷ָ� Then If rs������ϸ.RecordCount > 0 Then rs������ϸ.MoveFirst
    Do Until rs������ϸ.EOF
        bln���� = True
        If Nvl(rs������ϸ!����, 0) = 0 Then
            curʵ�ʵ��� = 0
        Else
            curʵ�ʵ��� = rs������ϸ!ʵ�ս�� / Nvl(rs������ϸ!����, 0)
        End If
        
        If lng����ID <> rs������ϸ("����ID") Then
            '���ж��ǲ���ҽ������
            blnҽ������ = False
            If Not bln���� Then
                gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1] And A.��ҳID=B.סԺ���� And A.����=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ҽ������", CLng(rs������ϸ!����ID), TYPE_������)
                blnҽ������ = (rsTemp!Records = 1)
            Else
                blnҽ������ = True
            End If
            
            If blnҽ������ Then
                lng����ID = rs������ϸ("����ID")
                '��ͬ�Ĳ��ˣ��������ڲ�ͬ�����ģ��䴲λ�޼�Ҳ���ܲ�ͬ������Ҫ��������
                gstrSQL = "SELECT B.���� ����,C.���� AS ���ֱ��� " & _
                    "FROM �����ʻ� A,��������Ŀ¼ B,���ղ��� C " & _
                    "WHERE A.����ID=" & lng����ID & " AND A.����=" & TYPE_������ & " AND A.����=B.���� AND nvl(A.����,0)=nvl(B.���,0) AND A.����ID=C.ID(+)"
                If rsTemp.State = adStateOpen Then rsTemp.Close
                rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                
                '�õ���ҽ�����˵Ĳ�����׼��Ŀ
                gstrSQL = "SELECT A.��Ŀ���,A.�����Ը����� FROM ���ղ�����Ŀ A Where A.������� ='" & rsTemp("���ֱ���") & "'"
                If rs������׼.State = adStateOpen Then rs������׼.Close
                rs������׼.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
                
                '�õ������Ĺ涨�Ĵ�λ���޼�
                str���ı��� = rsTemp("����")
                gstrSQL = "Select ÿ�촲λ���޼�,������Ŀ�۸� From ��������Ŀ¼ Where ����=" & TYPE_������ & " And ����='" & rsTemp("����") & "'"
                If rsTemp.State = adStateOpen Then rsTemp.Close
                rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
                cur��λ�� = rsTemp("ÿ�촲λ���޼�")
                cur������Ŀ = Nvl(rsTemp("������Ŀ�۸�"), 0)
            End If
        End If
        
        If blnҽ������ Then
            If �Ƿ���� = False Then
                If Getҽ������(rs������ϸ("�շ�ϸĿID"), str��Ŀ����, strϸĿ����) = False Then
                    MsgBox strϸĿ���� & "��û����ɱ��ձ���Ķ�Ӧ��������ɽ��㡣", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If IsNull(rs������ϸ("��Ŀ����")) = True Then
                    MsgBox "��Ϊ" & rs������ϸ("����") & "����ҽ�����롣", vbInformation, gstrSysName
                    Exit Function
                End If
                str��Ŀ���� = rs������ϸ("��Ŀ����")
                strϸĿ���� = rs������ϸ("����")
            End If
            
            '��ñ�����Ŀ����ϸ��Ϣ���������
            blnҽ����Ŀ = False
            gstrSQL = "Select a.����ҽ��,a.סԺҽ��,a.���۸�����,a.�۸�,a.�Ƿ�ҽ��,a.�������,b.�����Ը����� from ������Ŀ a,����֧��������� b Where a.����=" & TYPE_������ & " And a.����='" & str��Ŀ���� & "' and a.����=b.���� and a.�������=b.������� and b.���Ĵ���='" & str���ı��� & "'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
            If rsTemp.EOF Then
                MsgBox strϸĿ���� & "�ı��ձ������󣬲�����ɽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            
            If bln���� Then
                blnҽ����Ŀ = (Nvl(rsTemp!����ҽ��, 0) = 1)
            Else
                blnҽ����Ŀ = (Nvl(rsTemp!סԺҽ��, 0) = 1)
            End If
            If rs������ϸ("�շ����") = "J" Then
                '��λ��
                lng������Ŀ�� = 1
                If curʵ�ʵ��� <= cur��λ�� Then
                    curͳ���� = rs������ϸ("ʵ�ս��")
                Else
                    curͳ���� = cur��λ�� * rs������ϸ("����")
                End If
                curͳ�� = curͳ�� + curͳ����
                curȫ�Է� = curȫ�Է� + (rs������ϸ("ʵ�ս��") - curͳ����)
                cur���۸� = cur��λ��
            Else
                'Modified by zyb 20050429
                '�������Ŀ�������Ա����ļ۸�
                '�����ǰ��¼�ǳ�����¼����ȡԭʼ��¼���޼�
                If bln���� = False Then
                    If rs������ϸ!��¼״̬ = 2 Then
                        Dim rsGet As New ADODB.Recordset
                        Set rsGet = New ADODB.Recordset
                        If rsGet.State = 1 Then rsGet.Close
                        rsGet.Open "Select ���մ���ID,������Ŀ��,���ձ���,ͳ����,nvl(�޼�,0) AS �޼� From סԺ���ü�¼ " & _
                            " Where (NO,��¼����,��¼״̬,���) IN " & _
                            "     (Select NO,��¼����,3,��� From סԺ���ü�¼ Where ID=" & rs������ϸ!ID & ") And (Nvl(�Ƿ��ϴ�,0)=1 or nvl(����id,0)>0)", gcnOracle
                        If rsGet.RecordCount <> 0 Then
                            cur���۸� = rsGet!�޼�
                            curͳ���� = -1 * Nvl(rsGet!ͳ����, 0)
                            lng���մ���ID = Nvl(rsGet!���մ���id, 0)
                            lng������Ŀ�� = Nvl(rsGet!������Ŀ��, 0)
                            str��Ŀ���� = Nvl(rsGet!���ձ���)
                            bln���� = False
                        Else
                            cur���۸� = IIf(Nvl(rsTemp("���۸�����"), 0) = 0, Nvl(rsTemp("�۸�"), 0), rsTemp("���۸�����"))
                        End If
                    Else
                        cur���۸� = IIf(Nvl(rsTemp("���۸�����"), 0) = 0, Nvl(rsTemp("�۸�"), 0), rsTemp("���۸�����"))
                    End If
                Else
                    cur���۸� = IIf(Nvl(rsTemp("���۸�����"), 0) = 0, Nvl(rsTemp("�۸�"), 0), rsTemp("���۸�����"))
                End If
                
                If bln���� Then
                    'Modified by zyb ##2003-08-31
                    If ���÷ָ� Then
                        If Nvl(rs������ϸ("�޼�"), 0) = 0 And Nvl(rs������ϸ("ͳ����"), 0) = 0 Then
                            '������ü�¼�б�����޼�Ϊ����ͳ����ҲΪ�㣬��˵����ǰ�Ƿ�ҽ�����ˣ��Ե�ǰ���޼�Ϊ׼
                            'ҽ�������������ʣ�δ�����޼ۻ������޼�ǰ�ǵ��ʣ������ܲ������˷��ü�¼�е��޼�Ϊ����������ͳ�������Ϊ��
                            '��ҽ����Ŀ�����ܴ����޼۵����
                        Else
                            cur���۸� = Nvl(rs������ϸ("�޼�"), 0)
                        End If
                    End If
                    'Modified end
                    If cur���۸� > 0 And cur���۸� < curʵ�ʵ��� Then
                        '����Ŀ��������޼ۣ����ұ�ҽԺ�۸�Ҫ��
                        cur���� = cur���۸�
                    Else
                        cur���� = curʵ�ʵ���
                    End If
                    
                    rs������׼.Filter = "��Ŀ���='" & str��Ŀ���� & "'"
                    If rs������׼.EOF = False Then
                        '�Ƿ�ҽ����Ŀ�����˴���׼
                        lng������Ŀ�� = IIf(rs������׼("�����Ը�����") = 1, 0, 1)
                        cur�Ը����� = rs������׼("�����Ը�����")
                    Else
                        '�Ա�����Ŀ�е�ֵΪ׼
                        lng������Ŀ�� = rsTemp("�Ƿ�ҽ��")
                        cur�Ը����� = rsTemp("�����Ը�����")
                        
                        If lng������Ŀ�� = 1 And cur������Ŀ > 0 And _
                            (rs������ϸ("�շ����") <> "5" And rs������ϸ("�շ����") <> "6" And rs������ϸ("�շ����") <> "7") Then
                            
                            '���ڰ��۸����ּ����������Ŀ������
                            If curʵ�ʵ��� >= cur������Ŀ Then
                                cur�Ը����� = 0.2
                            Else
                                cur�Ը����� = 0
                            End If
                        End If
                        
                        '��Ȼ����Ϊ������Ŀ���������Ը��������Ը�Ϊȫ�Է�
                        If lng������Ŀ�� = 1 And rsTemp("�����Ը�����") = 1 Then lng������Ŀ�� = 0
                    End If
                    
                    If lng������Ŀ�� = 0 Or Not blnҽ����Ŀ Then
                        'ȫ�Է���Ŀ
                        '2005-09-12 by gzy lng������Ŀ��=rstemp("�Ƿ�ҽ��")*iff(rstemp("�����Ը�����"),1,0,1)*bln(ҽ����Ŀ)
                        lng������Ŀ�� = 0
                        curͳ���� = 0
                        curȫ�Է� = curȫ�Է� + rs������ϸ("ʵ�ս��")
                    Else
                        If cur���۸� = 0 Or curʵ�ʵ��� <= cur���۸� Then
                            'û�м۸����ƣ��������Ƶļ۸�û�г���
                            curͳ���� = rs������ϸ("ʵ�ս��") * (1 - cur�Ը�����)
                        Else
                            '�м۸����ƣ���ֻ��ȡ���۸�
                            curͳ���� = cur���۸� * rs������ϸ("����") * (1 - cur�Ը�����)
                        End If
                        curͳ�� = curͳ�� + curͳ����
                        
                        'Modified by zyb ##2003-08-31
                        '���������۸�����ʱ,�������Ը��ļ������Ӧ����(ȫ�Ը�=���޲���+��ҽ����Ŀ�ķ���;ʵ�ս��=ͳ����+�����Ը�+ȫ�Ը�)
                        If cur���۸� > 0 And cur���۸� < curʵ�ʵ��� Then
                            cur�Ը� = (cur���۸� * rs������ϸ("����") - curͳ����)
                        Else
                            cur�Ը� = (rs������ϸ("ʵ�ս��") - curͳ����)
                        End If
                        cur�����Ը� = cur�����Ը� + cur�Ը�
                        curȫ�Է� = curȫ�Է� + (rs������ϸ("ʵ�ս��") - curͳ���� - cur�Ը�)
                        'Modified end
                    End If
                End If
            End If
            
            If bln���� Then
                rs���մ���.Filter = "����='" & rsTemp("�������") & "'"
                If rs���մ���.EOF = False Then
                    lng���մ���ID = rs���մ���("ID")
                Else
                    lng���մ���ID = 0
                End If
            End If
            
            'ֻ������Ԥ���㲻����
            If �Ƿ���� = True Then
                '����������ƣ����������������շѷ���һ�������С�Ȼ��סԺ���ݶ����Ѿ�������˵ģ������ô���㶼����ν
                'Modified by zyb ##2003-09-01(��Ϊͳһ��ΪԤ����ʱȫ������,���Բ������Ƿ��ϴ���־)
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rs������ϸ("ID") & "," & curͳ���� & "," & _
                    lng���մ���ID & "," & lng������Ŀ�� & ",'" & str��Ŀ���� & "',NULL," & cur���۸� & ")"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
            
            'Modified by zyb ##2003-08-31
            If ���÷ָ� Then Exit Do
        End If
        rs������ϸ.MoveNext
    Loop
    
    Calc���÷ָ� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Getҽ������(ByVal ��ϸID As Long, ҽ������ As String, ϸĿ���� As String) As Boolean
'���ܣ����ݷ�����ϸID���õ���ҽ������
'��������ϸID     �շ�ϸĿ��ID
'      ҽ������   ���ֵ���շ�ϸĿ��Ӧ��ҽ������
'���أ��������ɹ�������й��ܣ�ΪTrue
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select A.��Ŀ����,B.���� From ����֧����Ŀ A,�շ�ϸĿ B Where B.ID=" & ��ϸID & " And B.ID=A.�շ�ϸĿID(+) And A.����(+)=" & TYPE_������
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    If rsTemp.EOF = False Then
        ҽ������ = Nvl(rsTemp("��Ŀ����"))
        ϸĿ���� = Nvl(rsTemp("����"))
    Else
        ҽ������ = ""
        ϸĿ���� = "IDΪ" & ��ϸID & "����Ŀ"
    End If
    
    Getҽ������ = (ҽ������ <> "")
End Function

Private Function Calc����ͳ��() As Boolean
'���ܣ������סԺ���˵���ͨ����ͳ����
'���������
'���������
'���أ��ɹ����㣬�򷵻�True
    
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim lng��ְ As Long, lng����� As Long, lng���� As Long
    
    Dim clsҽ�� As New clsInsure
    Dim dbl������ߺ� As Currency, dblԭ���� As Currency, dbl������ As Currency
    Dim dbl��ν���ͳ��� As Currency, dbl��������Ը��� As Currency     '�����ָ�ò��˱���סԺ��ǰ���ʵ��ۼ�
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency, curͳ�� As Currency
    Dim str��Ŀ���� As String, str��Ŀ���� As String
    '�������
    Dim bln������ As Boolean, bln�޷ⶥ�� As Boolean, blnҽ����Ŀ As Boolean
    
    On Error GoTo errHandle
    '������������������������������������������������������������������������������������
    '1����ʼ��һЩ�������Լ��������
    Set gcol������� = New Collection
    
    m����.�����ʻ�֧��ȫ�Է� = clsҽ��.GetCapability(support�����ʻ�ȫ�Է�, 0, TYPE_������)
    m����.�����ʻ�֧�������Ը� = clsҽ��.GetCapability(support�����ʻ������Ը�, 0, TYPE_������)
    m����.�����ʻ�֧������ = clsҽ��.GetCapability(support�����ʻ�����, 0, TYPE_������)
    
    gstrSQL = "SELECT B.ҽ����,A.�����ڶ���,A.��ֵ����,A.�ⶥ����,A.���䱨�����𸶽�,A.ʹ���ۼƱ���,A.�����˻���֧�������Ը� " & _
               " ,A.�����𸶽�����,A.��������סԺ���� " & _
               " FROM ��������Ŀ¼ A,�������� B " & _
               " WHERE A.����=" & TYPE_������ & " AND A.����='" & gIC����.CenterCode & "' AND A.��������=B.���� AND A.����=B.���� "
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        m����.��� = Val(Nvl(rsTemp("ҽ����")))
        m����.���ö�ֵ = Nvl(rsTemp("��ֵ����")) = 1
        m����.���÷ⶥ = Nvl(rsTemp("��ֵ����")) = 1
        m����.�����ڶ��� = Nvl(rsTemp("�����ڶ���")) = 1
        m����.ʹ���ۼ� = Nvl(rsTemp("ʹ���ۼƱ���")) = 1
        m����.���䱨�����𸶽� = Nvl(rsTemp("ʹ���ۼƱ���")) = 1
        m����.�����ʻ�֧�������Ը� = Nvl(rsTemp("�����˻���֧�������Ը�")) = 1
        m����.�����𸶽����� = Nvl(rsTemp("�����𸶽�����"), 0)
        m����.��������סԺ���� = Nvl(rsTemp("��������סԺ����"), 0)
    End If
    If m����.��� = 0 Then
        MsgBox "��ϵͳ����Ա���ҽ�����ݵ����ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '1.1��������˱���סԺ�ĸ��ַ���
    'Modified by zyb ##2003-08-31(��׼���۸�Ϊ����,����Ϊʵ�ս��)
    '���㹫ʽ:ȫ�Ը�=���޲���+��ҽ����Ŀ�ķ���;ʵ�ս��=ͳ����+�����Ը�+ȫ�Ը�
    gstrSQL = _
        "Select Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.NO,Nvl(A.�۸񸸺�,���) as ���,A.����ID,A.��ҳID," & _
        "   A.�շ����,A.�շ�ϸĿID,Nvl(A.���մ���ID,0) as ���մ���ID,Avg(Nvl(A.����,1)*A.����) as ����,NVL(A.ͳ����,0) as ͳ����," & _
        "   Sum(A.��׼����) as ����,Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as ʵ�ս��,A.����ʱ��,Nvl(A.������Ŀ��,0) as ������Ŀ��,Nvl(Sum(�޼�),0) �޼�" & _
        "   From סԺ���ü�¼ A" & _
        "   Where A.���ʷ���=1 And Nvl(A.��¼״̬,0)<>0 And A.����ID=[1] and A.��ҳID=[2] And A.����Ա���� is not null" & _
        "   Group by Mod(A.��¼����,10),A.��¼״̬,A.NO,Nvl(A.�۸񸸺�,���),A.����ID,A.��ҳID," & _
        "       A.�շ����,A.�շ�ϸĿID,Nvl(A.���մ���ID,0),A.����ʱ��,Nvl(A.������Ŀ��,0),NVL(A.ͳ����,0)" & _
        "   Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", m����.����ID, m����.��ҳID)
    
    With m����
        Do Until rsTemp.EOF
            '2004-05-27
            '�ж��Ƿ�����ҽ����Ŀ�����ڴ�λ��δ���÷��÷ָ����ֻ�е����жϣ�
            If Getҽ������(rsTemp("�շ�ϸĿID"), str��Ŀ����, str��Ŀ����) = False Then
                MsgBox str��Ŀ���� & "��û����ɱ��ձ���Ķ�Ӧ��������ɽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            '2005-09-12 by gzy �ж�סԺҽ����Ŀ��cale���÷ָ�����ɣ���"�Ƿ�ҽ��"��ʶ
            'gstrSQL = "Select * from ������Ŀ Where ����=" & TYPE_������ & " And ����='" & str��Ŀ���� & "'"
            'If rsCheck.State = adStateOpen Then rsCheck.Close
            'rsCheck.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
            'blnҽ����Ŀ = (Nvl(rsCheck!סԺҽ��, 0) = 1)
            
            If rsTemp("������Ŀ��") = 0 Then  'Or Not blnҽ����Ŀ
                .ȫ�Է� = .ȫ�Է� + rsTemp("ʵ�ս��")
            Else
                If rsTemp("�շ����") = "J" Then
                    .����ͳ�� = .����ͳ�� + rsTemp("ͳ����")
                    .ҽ����Ŀ��� = .ҽ����Ŀ��� + rsTemp("ͳ����")
                    If rsTemp("ʵ�ս��") <> rsTemp("ͳ����") Then
                        .ȫ�Է� = .ȫ�Է� + rsTemp("ʵ�ս��") - rsTemp("ͳ����")
                    End If
                Else
                    .����ͳ�� = .����ͳ�� + rsTemp("ͳ����")
                    If rsTemp("ʵ�ս��") <> rsTemp("ͳ����") Then
                        '����������Ŀ�Ľ��
                        'Modified by zyb ##2004-11-13   ���ܰ�����ʱҽ����Ŀ��״̬�ٴ����㣬��Ϊ����������ͳ����ֱ��ȡ��ԭ����ֵ
'                        Call Calc���÷ָ�(rsTemp, False, curȫ�Է�, cur�����Ը�, curͳ��, True)
'                        If cur�����Ը� = 0 Then 'ֻ��������������Ը��������Ƕ��޼۵Ĵ���
'                            .ҽ����Ŀ��� = .ҽ����Ŀ��� + curͳ��
'                        Else
'                            .������Ŀ��� = .������Ŀ��� + curͳ�� + cur�����Ը�
'                        End If
                        If rsTemp!�޼� <> 0 Then
                            '�޼��ǵ����������޼�
                            If rsTemp!�޼� * rsTemp!���� = rsTemp!ͳ���� Then
                                '����
                                .ҽ����Ŀ��� = .ҽ����Ŀ��� + rsTemp!ͳ����
                                curȫ�Է� = rsTemp!ʵ�ս�� - rsTemp!ͳ����
                                cur�����Ը� = 0
                            Else
                                'Modified by zyb 20050429
                                '������Ŀ
                                If rsTemp!�޼� >= (rsTemp!ʵ�ս�� / Nvl(rsTemp!����, 1)) Then
                                    '�������˵���޼۴��ڵ��ۣ��Ե���Ϊ׼�����ͳ�����ʱȫ�Է�Ϊ��
                                    curȫ�Է� = 0
                                    cur�����Ը� = rsTemp!ʵ�ս�� - rsTemp!ͳ����
                                    .������Ŀ��� = .������Ŀ��� + (rsTemp!ͳ���� + cur�����Ը�)
                                Else
                                    curȫ�Է� = rsTemp!ʵ�ս�� - (rsTemp!�޼� * rsTemp!����)
                                    cur�����Ը� = (rsTemp!�޼� * rsTemp!����) - rsTemp!ͳ����
                                    .������Ŀ��� = .������Ŀ��� + (rsTemp!ͳ���� + cur�����Ը�)
                                End If
                            End If
                        Else
                            curȫ�Է� = 0
                            cur�����Ը� = rsTemp!ʵ�ս�� - rsTemp!ͳ����
                            .������Ŀ��� = .������Ŀ��� + (rsTemp!ͳ���� + cur�����Ը�)
                        End If

                        .ȫ�Է� = .ȫ�Է� + curȫ�Է�
                        .�����Ը� = .�����Ը� + cur�����Ը�
                        'Modified end
                    Else
                        .ҽ����Ŀ��� = .ҽ����Ŀ��� + rsTemp("ʵ�ս��")
                    End If
                End If
            End If
            
            .�������� = .�������� + rsTemp("ʵ�ս��")
            rsTemp.MoveNext
        Loop
    End With
    
    '1.2���õ��ʻ��������Ϣ
    With m����
        gstrSQL = "select A.��Ա���,A.��ְ,A.�����," & _
                  "      B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ�" & _
                  " from �����ʻ� A,�ʻ������Ϣ B" & _
                  " where A.����ID=B.����ID(+) and A.����=B.����(+) " & _
                  "     and B.���(+)=[1] and A.����ID=[2] and A.����=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", .���, .����ID, TYPE_������)
        
        lng��ְ = IIf(IsNull(rsTemp("��ְ")), 1, rsTemp("��ְ"))
        lng���� = IIf(IsNull(rsTemp("�����")), 0, rsTemp("�����"))
        .סԺ���� = IIf(IsNull(rsTemp("סԺ�����ۼ�")), 0, rsTemp("סԺ�����ۼ�"))
        
        gstrSQL = "select �����,nvl(ȫ��ͳ��,0) as ȫ��ͳ�� ,nvl(������,0) as ������ ,nvl(�޷ⶥ��,0) as �޷ⶥ�� " & _
                " from ���������" & _
                " where ����=" & TYPE_������ & " and nvl(����,0)=" & .������� & _
                "       and ��ְ=" & lng��ְ & " and ����<=" & lng���� & " and (" & lng���� & "<=���� or ����=0)"
        Call OpenRecordset_OtherBase(rsTemp, "", gstrSQL, gcn����)
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڡ��������������������������õ���", vbInformation, gstrSysName
            Exit Function
        End If
        lng����� = rsTemp("�����")
        bln������ = (rsTemp("������") = 1)
        bln�޷ⶥ�� = (rsTemp("�޷ⶥ��") = 1)
        
        m����.ȫ��ͳ�� = (rsTemp("ȫ��ͳ��") = 1)
    End With
    
    '1.3 ��������סԺ�ڼ��ۼƽ������
    gstrSQL = "select nvl(max(A.����),0) as ԭ����,nvl(sum(A.ʵ������*��Ʊ��־),0) as ����,nvl(sum((A.�������ý��-A.ȫ�Ը����-A.�����Ը����)*��Ʊ��־),0) as ����ͳ����,nvl(sum(A.�����Ը����*��Ʊ��־),0) as �����Ը���� " & _
              "  from ���ս����¼ A " & _
              "  Where A.����ID = " & m����.����ID & " And A.��ҳID = " & m����.��ҳID & _
              " And A.���� = " & TYPE_������
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    dblԭ���� = rsTemp("ԭ����")
    dbl������ߺ� = rsTemp("����")
    dbl��ν���ͳ��� = rsTemp("����ͳ����")
    dbl��������Ը��� = rsTemp("�����Ը����")
    
    '������������������������������������������������������������������������������������
    '3��������ߡ��ⶥ�ߡ�֧������������
    '3.1��������ߡ��ⶥ��
    'Modified By ���� 2004-05-08 ԭ�򣺽���������
    With m����
        gstrSQL = "select max(decode(A.����,'A',A.���,0)) as ������ ,max(decode(A.����,'1',A.���,0)) as ���� " & _
                  "         ,max(decode(A.����,'" & (.סԺ���� + 1) & "',A.���,0)) as ʵ������,min(A.���) as ������� " & _
                  "  from ����֧���޶� A " & _
                  "  where A.����=" & TYPE_������ & " and A.����=" & .������� & " and A.���=" & .��� & " And A.��ְ=" & lng��ְ & " And A.��Ա���=" & IIf(gIC����.DomainCode = 0, 1, 2)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
                
        If bln������ Then
            dbl������ = 0
        Else
            dbl������ = IIf(IsNull(rsTemp("ʵ������")), 0, rsTemp("ʵ������"))
            If dbl������ = 0 Then
                'һ�㶼���У����ʵ�ڳ�����סԺ��������ȡ���һ�Σ�Ҳ���ǽ����С��һ�Σ�
                dbl������ = IIf(IsNull(rsTemp("�������")), 0, rsTemp("�������"))
            End If
            If dbl������ = 0 Then
                MsgBox "���ڡ���Ƚ�����������ñ���ȵ����ߡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If bln�޷ⶥ�� Then
            .�ⶥ�� = 0
        Else
            .�ⶥ�� = IIf(IsNull(rsTemp("������")), 0, rsTemp("������"))
            If .�ⶥ�� = 0 Then
                MsgBox "���ڡ���Ƚ�����������ñ���ȵķⶥ�ߡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    Dim bln������ As Boolean
    If m����.����סԺ = False Then
        m����.���� = dbl������
        bln������ = True
    Else
        Select Case m����.�����𸶽�����
            Case 0
                m����.���� = dblԭ����
                bln������ = True
            Case 1
                m����.���� = dbl������
                bln������ = True
            Case Else
                m����.���� = dbl������
                bln������ = False
        End Select
    End If
    
    '���������Ҫ�۳�������
    If bln������ = True Then
        If m����.���� > dbl������ߺ� Then
            '�õ�Ԥ��֧�������ߣ����������յ�
            m����.�������� = m����.���� - dbl������ߺ�
        Else
            'û��Ҫ֧��������
            m����.�������� = 0
        End If
    End If
    
    '�Ƿ�����סԺ����
    If m����.��;���� = 0 Then
        '��Ժ
        If m����.����סԺ = True Then
            '�����סԺ
            m����.סԺ�������� = m����.��������סԺ����
        Else
            m����.סԺ�������� = IIf(m����.�Ҷ� = degֹ֧ͣ��, 0, 1)
        End If
    End If
    
    If m����.�Ҷ� < deg����֧�� Then
        '����Ҫ�ټ����뱨����ص�ֵ
        Calc����ͳ�� = True
        Exit Function
    End If
    
    '������������������������������������������������������������������������������������
    '4������ôν���ɱ����Ľ�Ϊ�˱Ƚ����Ե��˽������ʹ�ã��ʰѱ�������д������
    With m����
        If m����.ʹ���ۼ� = True Then
            '�ۼƽ��ʹӿ���ȡ
            .ͳ����֧������ = gIC����.PlanPaidFee
            .ͳ����֧����� = gIC����.PlanPaidAmt
        Else
            '������סԺ��Ҫ�ۼ�
            gstrSQL = "SELECT nvl(sum(����ͳ��֧��*��Ʊ��־),0) �ۼ�֧��,nvl(sum(����ͳ�����*��Ʊ��־),0) �ۼƷ��� " & _
                      "FROM ���ս����¼ WHERE ����ID=" & .����ID & " AND ��ҳID=" & .��ҳID & " AND ����=2 AND ����=" & TYPE_������
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn����
            .ͳ����֧������ = rsTemp("�ۼƷ���")
            .ͳ����֧����� = rsTemp("�ۼ�֧��")
        End If
    
        '������㣬��Щ��Ӧ����0
        If .������� = True Then
            '�������Ͳ��ÿ�����ǰ�Ľ�����
            dbl������ߺ� = 0
            dbl��ν���ͳ��� = 0
            .ͳ����֧������ = 0
            .ͳ����֧����� = 0
        End If
        
    
        '����Ѿ������ⶥ��ֱ���˳�������Ҫ�ٿ�������
        If m����.���÷ⶥ = True Then
            '���÷ⶥ�ĳ��ⶥ�߿��ܺ��������Ը�����
            If .ͳ����֧������ >= .�ⶥ�� And .�ⶥ�� > 0 Then
                .�������ⶥ�� = .�������� - .ȫ�Է�
                Calc����ͳ�� = True
                Exit Function
            End If
        Else
            '֧���ⶥ�ĳ��ⶥ��ֻ�ܺ��н���ͳ�ﲿ��
            If .ͳ����֧����� >= .�ⶥ�� And .�ⶥ�� > 0 Then
                .�������ⶥ�� = .����ͳ��
                Calc����ͳ�� = True
                Exit Function
            End If
        End If
    
        '3.3��ȡ�÷��õ���
        If rsTemp.State = adStateOpen Then rsTemp.Close
        'Modified By ���� 2004-05-08 ԭ�򣺽���������
        gstrSQL = "select B.����,B.����,B.����,A.���� " & _
                  "  from ����֧������ A,���շ��õ� B " & _
                  "  Where A.���� =" & TYPE_������ & " And A.���� =" & m����.������� & " And A.��� =" & m����.��� & " And A.��ְ =" & lng��ְ & " And A.����� =" & lng����� & _
                  "       and A.����=B.���� and A.����=b.���� and A.����=B.���� And A.��ְ=B.��ְ and A.��Ա���=" & IIf(gIC����.DomainCode = 0, 1, 2) & _
                  "  order by B.����"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڡ���Ƚ�����������ñ���ȵ�ͳ��֧����������", vbInformation, gstrSysName
            Exit Function
        End If
        
        'Ȼ�����ֶμ���
        '�����ʵ�����ߡ��ֶα������ֶν������
        If m����.���ö�ֵ = True Then
            '���ö�ֵ
            If m����.���÷ⶥ = False Then
                '֧���ⶥ���������Թ�����ģʽ
                If Calc�����ֶ�1(rsTemp, m����.�����ڶ��� = False, dbl������ߺ�, dbl��ν���ͳ���) = False Then Exit Function
            Else
                '���÷ⶥ������������ģʽ
                If Calc�����ֶ�2(rsTemp, m����.�����ڶ��� = False, dbl������ߺ�, dbl��ν���ͳ���, dbl��������Ը���) = False Then Exit Function
            End If
        Else
            '֧����ֵ
            If m����.���÷ⶥ = False Then
                '֧���ⶥ
                If Calc�����ֶ�3(rsTemp) = False Then Exit Function
            Else
                '���÷ⶥ
                If Calc�����ֶ�4(rsTemp) = False Then Exit Function
            End If
        End If
        
        'takecare
        '���㳬���Ը�����
        If .�ⶥ�� > 0 Then
            '�зⶥ��
            If m����.���÷ⶥ = True Then
                '.����ͳ�� �� .�����Ը� �����ڷ��ã�������ؿ��𸶽�Ļ������ⶥ����Ҳ���ܰ����ǲ��ֽ�
                .�������ⶥ�� = (.�������� + .ͳ����֧������) - .ȫ�Է� - .�ⶥ�� '- IIf(m����.���䱨�����𸶽� = True, .ʵ������, 0)
            Else
                '֧���ⶥ��ֻ��ͳ�ﲿ��
                .�������ⶥ�� = .����ͳ�� - .ͳ�����֧�� - .ͳ������Ը� - .ʵ������
            End If
            If .�������ⶥ�� < 0 Then .�������ⶥ�� = 0                   '�������ͳ����������ߣ�Ϊ����
        End If
        
        '����ں������У��Ҷ�<ҽ��֧��������ʵ������=0�����ͳ���Ը�ֵ=0
        If m����.�Ҷ� = degֹ֧ͣ�� Then
            .ʵ������ = 0
        End If
        If m����.�Ҷ� < deg����֧�� Then
            .��������Ը� = 0
            .���������Ը� = 0
            .ͳ������Ը� = 0
        End If
    End With
        
    Calc����ͳ�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Calc�����ֶ�1(rs���ö� As ADODB.Recordset, bln�ȿ����� As Boolean, dbl������� As Currency, dbl��ν���ͳ�� As Currency) As Boolean
'���ܣ����㰴���÷ֶΣ�֧���ⶥ�����
    Dim dbl��֧����� As Currency  '���ݲ����õ��Ѿ�ʹ�õĽ������
    Dim dbl���� As Currency       'ÿһ�ε����ֵ�������Ƿ��ã�Ҳ������֧�����
    Dim dbl���� As Currency       'ÿһ�ε����ֵ�������Ƿ��ã�Ҳ������֧�����
    Dim dbl�ֶν��� As Currency   '����ĳһ�ε�ͳ����
    Dim dbl�ֶα��� As Currency   '����ĳһ�ε�ͳ�ﱨ�����
    Dim dbl���ν��� As Currency   '�����ܵĽ���ͳ����
    Dim dbl���α��� As Currency   '�����ܵĽ��뱨�����
    
    Dim dbl��� As Currency  '���ڼ�������ֵ
    Dim dblʣ�� As Currency  '���������õ�ͳ����
    
    Dim dblTemp As Currency, lng���� As Long
    Dim dbl���� As Currency
    
    dbl���� = m����.��������
    dbl��֧����� = m����.ͳ����֧�����
    
    If bln�ȿ����� = True Then
        '���Ȱ����߽��۳�
        If m����.����ͳ�� > dbl���� Then
            '���۳�
            m����.ʵ������ = dbl����
            dbl���� = 0
            '��Ϊ�����Ѿ���ɿ۳����������������εĽ���ͳ�����ȥ����
            If dbl��ν���ͳ�� > dbl������� Then
                dbl��� = dbl��ν���ͳ�� - dbl�������
            Else
                dbl��� = 0
            End If
            dblʣ�� = m����.����ͳ�� - m����.ʵ������
        Else
            '�����߶�������֧����ֱ���˳�
            m����.ʵ������ = m����.����ͳ��
            
            Do Until rs���ö�.EOF
                lng���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
                dblTemp = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
                    
                gcol�������.Add Array(lng����, 0, 0, dblTemp)
                rs���ö�.MoveNext
            Loop
            Calc�����ֶ�1 = True
            Exit Function
        End If
    Else
        dbl��� = dbl��ν���ͳ��
        dblʣ�� = m����.����ͳ��
    End If
    
    Do Until rs���ö�.EOF
        dbl�ֶν��� = 0
        dbl�ֶα��� = 0
        dbl���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        dbl���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        
        '֧���ⶥ���������Թ�����ģʽ
        If dbl��֧����� < m����.�ⶥ�� Or m����.�ⶥ�� = 0 Then    'δ�����ⶥ�߻��޷ⶥ��
            '�����Լ�������
            If dbl���� = 0 Then
                '��һ����Ҫ�������������ȷ�Լ�飬������޹�
                If m����.���� > dbl���� And dbl���� > 0 Then
                    MsgBox "�ò��˵�ʵ�����߱ȵ�һ�����õ����޻��࣬���鱣�շ��õ���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If dbl��� >= dbl���� And (dbl��� < dbl���� Or dbl���� = 0) And dblʣ�� > 0 Then
                '�ö���ǰ��δ������ȫ�����������Ҫ����۳��Ľ��Ѿ�������ĶΣ����߽����Ķβ�����룩
                If dbl���� = 0 Then
                    dbl�ֶν��� = dblʣ�� '��ȫ������
                Else
                    '��ʣ��ֵ�뱾�οռ�֮��ѡ��Сֵ
                    dbl�ֶν��� = dbl���� - dbl���
                    If dbl�ֶν��� > dblʣ�� Then dbl�ֶν��� = dblʣ��
                End If
                '�����ƣ��ɱ������仯
                dbl��� = dbl��� + dbl�ֶν���
                dblʣ�� = dblʣ�� - dbl�ֶν���
                If dbl���� > 0 Then
                    '����Ҫ�����߾ͽ�����������
                    If dbl�ֶν��� > dbl���� Then
                        '������������ߣ�'���۳������⻹��һ�������ڱ���
                        m����.ʵ������ = m����.ʵ������ + dbl����
                        dbl�ֶν��� = dbl�ֶν��� - dbl����
                        dbl���� = 0
                    Else
                        'ȫ��������������ߣ�ʣ������߻�Ҫ������һ��
                        m����.ʵ������ = m����.ʵ������ + dbl�ֶν���
                        dbl���� = dbl���� - dbl�ֶν���
                        dbl�ֶν��� = 0
                    End If
                End If
                
                '����������öεı������
                dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
                dbl�ֶα��� = Val(Format(dbl�ֶν��� * rs���ö�("����") / 100, "0.00")) '���Ǹö������Ա����Ľ��
                
                If dbl��֧����� + dbl�ֶα��� > m����.�ⶥ�� And m����.�ⶥ�� <> 0 Then
                    '���������˷ⶥ�ߣ����Ҵ��ڷⶥ������
                    dbl�ֶα��� = m����.�ⶥ�� - dbl��֧�����
                    
                    '���ƽ���ͳ����
                    If rs���ö�("����") <> 0 Then
                        dbl�ֶν��� = dbl�ֶα��� * 100 / rs���ö�("����")
                    Else
                        dbl�ֶν��� = 0
                    End If
                End If
                
            End If
        End If
        
        dbl��֧����� = dbl��֧����� + dbl�ֶα���
        
        '���Ρ�����ͳ���ͳ�ﱨ��������
        '���и�ʽ��
        dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
        dbl�ֶα��� = Val(Format(dbl�ֶα���, "0.00"))
        lng���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        dblTemp = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        gcol�������.Add Array(lng����, dbl�ֶν���, dbl�ֶα���, dblTemp)
        
        dbl���ν��� = dbl�ֶν��� + dbl���ν���
        dbl���α��� = dbl���α��� + dbl�ֶα���
        rs���ö�.MoveNext
    Loop
    
    m����.ͳ�����֧�� = dbl���α���
    m����.ͳ������Ը� = dbl���ν��� - dbl���α���
    
    Calc�����ֶ�1 = True
End Function

Private Function Calc�����ֶ�2(rs���ö� As ADODB.Recordset, bln�ȿ����� As Boolean, dbl������� As Currency, dbl��ν���ͳ�� As Currency, dbl��������Ը� As Currency) As Boolean
'���ܣ����㰴���÷ֶΣ�֧���ⶥ�����
    Dim dbl���� As Currency       'ÿһ�ε����ֵ�������Ƿ��ã�Ҳ������֧�����
    Dim dbl���� As Currency       'ÿһ�ε����ֵ�������Ƿ��ã�Ҳ������֧�����
    Dim dbl�ֶν��� As Currency   '����ĳһ�ε�ͳ����
    Dim dbl�ֶα��� As Currency   '����ĳһ�ε�ͳ�ﱨ�����
    Dim dbl���ν��� As Currency   '�����ܵĽ���ͳ����
    Dim dbl���α��� As Currency   '�����ܵĽ��뱨�����
    
    Dim dbl��� As Currency  '���ڼ�������ֵ
    Dim dblʣ����� As Currency  '���������õķ���
    Dim dblʣ��ͳ�� As Currency  '���������õ�ͳ����
    
    Dim dblTemp As Currency, lng���� As Long
    Dim dbl���� As Currency
    
    dbl���� = m����.��������
    If m����.�ⶥ�� > 0 Then
        '�������������ʹ�õķ���
        dblʣ����� = m����.�ⶥ�� - m����.ͳ����֧������
        If dblʣ����� < 0 Then dblʣ����� = 0
        
        '������ⲿ�ַ����е�ͳ����
        If dblʣ����� > m����.ҽ����Ŀ��� Then
            dblʣ��ͳ�� = m����.ҽ����Ŀ���
            dblʣ����� = dblʣ����� - m����.ҽ����Ŀ���
            
            If dblʣ����� > m����.������Ŀ��� Then
                dblʣ��ͳ�� = dblʣ��ͳ�� + m����.������Ŀ��� * 0.8
            Else
                'Modified by zyb ##2003-08-31
                '�����ⶥ��,����Calc����ͳ��()�м���������Ը���ʵ�ʵ������Ը�������,��Ҫ���¼���
                dblʣ��ͳ�� = dblʣ��ͳ�� + dblʣ����� * 0.8 '����ʹ��һ����ֵ
                m����.�����Ը� = dblʣ����� * 0.2
            End If
        Else
            dblʣ��ͳ�� = dblʣ�����
            m����.�����Ը� = 0
        End If
    Else
        dblʣ��ͳ�� = m����.����ͳ��
    End If
    
    If bln�ȿ����� = True Then
        '���Ȱ����߽��۳�
        If dblʣ��ͳ�� > dbl���� Then
            '���۳�
            m����.ʵ������ = dbl����
            dbl���� = 0
            '��Ϊ�����Ѿ���ɿ۳����������������εĽ���ͳ�����ȥ����
            If dbl��ν���ͳ�� > dbl������� Then
                dbl��� = m����.ͳ����֧������ - dbl������� '��֧�������а�����dbl��ν���ͳ����Ա�dbl������ߴ�
            Else
                dbl��� = m����.ͳ����֧������ - dbl��ν���ͳ�� '���뻹����ǰ�������Ը�Ҫ��
            End If
            
            'Modified By ���� 2003-12-10 ����������
            If dbl��� < 0 Then dbl��� = 0
            dblʣ��ͳ�� = dblʣ��ͳ�� - m����.ʵ������
        Else
            '�����߶�������֧����ֱ���˳�
            m����.ʵ������ = dblʣ��ͳ��
            
            Do Until rs���ö�.EOF
                lng���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
                dblTemp = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
                    
                gcol�������.Add Array(lng����, 0, 0, dblTemp)
                rs���ö�.MoveNext
            Loop
            Calc�����ֶ�2 = True
            Exit Function
        End If
    Else
        dbl��� = m����.ͳ����֧������
    End If
    
    Do Until rs���ö�.EOF
        dbl�ֶν��� = 0
        dbl�ֶα��� = 0
        dbl���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        dbl���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        If dbl���� = 0 Then dbl���� = m����.�ⶥ�� '�����Ƿ��÷ⶥ��Ҳ�Ϳ�����Ϊ��ֵ
        
        '�����Լ�������
        If dbl���� = 0 Then
            '��һ����Ҫ�������������ȷ�Լ�飬������޹�
            If m����.���� > dbl���� And dbl���� > 0 Then
                MsgBox "�ò��˵�ʵ�����߱ȵ�һ�����õ����޻��࣬���鱣�շ��õ���", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If dbl��� >= dbl���� And (dbl��� < dbl���� Or dbl���� = 0) And dblʣ��ͳ�� > 0 Then
            '�ö���ǰ��δ������ȫ�����������Ҫ����۳��Ľ��Ѿ�������ĶΣ����߽����Ķβ�����룩
            If dbl���� = 0 Then
                dbl�ֶν��� = dblʣ��ͳ�� '��ȫ������
            Else
                '��ʣ��ֵ�뱾�οռ�֮��ѡ��Сֵ
                dbl�ֶν��� = dbl���� - dbl���
                If dbl�ֶν��� > dblʣ��ͳ�� Then dbl�ֶν��� = dblʣ��ͳ��
            End If
            
            '�����ƣ���ʹ�÷��ñ仯
            dbl��� = dbl��� + dbl�ֶν���
            dblʣ��ͳ�� = dblʣ��ͳ�� - dbl�ֶν���
            
            If dbl���� > 0 Then
                '����Ҫ�����߾ͽ�����������
                If dbl�ֶν��� > dbl���� Then
                    '������������ߣ�'���۳������⻹��һ�������ڱ���
                    m����.ʵ������ = m����.ʵ������ + dbl����
                    dbl�ֶν��� = dbl�ֶν��� - dbl����
                    dbl���� = 0
                Else
                    'ȫ��������������ߣ�ʣ������߻�Ҫ������һ��
                    m����.ʵ������ = m����.ʵ������ + dbl�ֶν���
                    dbl���� = dbl���� - dbl�ֶν���
                    dbl�ֶν��� = 0
                End If
            End If
            
            '����������öεı���������������Ĳα���Ա����ֵ�����Ӧ��Ϊ�㣩
            dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
            If m����.�Ҷ� < deg����֧�� Then
                dbl�ֶα��� = 0
            Else
                dbl�ֶα��� = Val(Format(dbl�ֶν��� * rs���ö�("����") / 100, "0.00")) '���Ǹö������Ա����Ľ��
            End If
        End If
        
        '���Ρ�����ͳ���ͳ�ﱨ��������
        '���и�ʽ��
        dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
        dbl�ֶα��� = Val(Format(dbl�ֶα���, "0.00"))
        lng���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        dblTemp = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        gcol�������.Add Array(lng����, dbl�ֶν���, dbl�ֶα���, dblTemp)
        
        dbl���ν��� = dbl�ֶν��� + dbl���ν���
        dbl���α��� = dbl���α��� + dbl�ֶα���
        rs���ö�.MoveNext
    Loop
    
    m����.ͳ�����֧�� = dbl���α���
    m����.ͳ������Ը� = dbl���ν��� - dbl���α���
    
    Calc�����ֶ�2 = True
End Function

Private Function Calc�����ֶ�3(rs���ö� As ADODB.Recordset) As Boolean
'���ܣ����㰴���÷ֶΣ�֧���ⶥ�����
    Dim dbl��֧����� As Currency  '���ݲ����õ��Ѿ�ʹ�õĽ������
    Dim dbl���� As Currency       'ÿһ�ε����ֵ�������Ƿ��ã�Ҳ������֧�����
    Dim dbl���� As Currency       'ÿһ�ε����ֵ�������Ƿ��ã�Ҳ������֧�����
    Dim dbl�ֶν��� As Currency   '����ĳһ�ε�ͳ����
    Dim dbl�ֶα��� As Currency   '����ĳһ�ε�ͳ�ﱨ�����
    Dim dbl���ν��� As Currency   '�����ܵĽ���ͳ����
    Dim dbl���α��� As Currency   '�����ܵĽ��뱨�����
    
    Dim dbl��� As Currency  '���ڼ�������ֵ
    Dim dblʣ�� As Currency  '���������õ�ͳ����
    
    Dim dblTemp As Currency, lng���� As Long
    Dim dbl���� As Currency
    
    dbl���� = m����.��������
    dbl��֧����� = m����.ͳ����֧�����
    
    '���Ȱ����߽��۳�����Ϊ������Զ���ܱ����ģ����Բ��ܷ�����һ��ȥ�жϡ��о��У����оͲ��С�
    If m����.����ͳ�� > dbl���� Then
        '���۳�
        m����.ʵ������ = dbl����
        dbl��� = m����.ͳ����֧�����   '˵���Ѿ�֧������
        dblʣ�� = m����.����ͳ�� - m����.ʵ������
    Else
        '�����߶�������֧����ֱ���˳�
        m����.ʵ������ = m����.����ͳ��
        
        Do Until rs���ö�.EOF
            lng���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
            dblTemp = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
                
            gcol�������.Add Array(lng����, 0, 0, dblTemp)
            rs���ö�.MoveNext
        Loop
        Calc�����ֶ�3 = True
        Exit Function
    End If
    
    Do Until rs���ö�.EOF
        dbl�ֶν��� = 0
        dbl�ֶα��� = 0
        dbl���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        dbl���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        
        '֧���ⶥ
        If dbl��֧����� < m����.�ⶥ�� Or m����.�ⶥ�� = 0 Then    'δ�����ⶥ�߻��޷ⶥ��
            '�����Լ�������
            If dbl��� >= dbl���� And (dbl��� < dbl���� Or dbl���� = 0) And dblʣ�� > 0 Then
                '�ö���ǰ��δ������ȫ�����������Ҫ����۳��Ľ��Ѿ�������ĶΣ����߽����Ķβ�����룩
                If dbl���� = 0 Then
                    'dbl�ֶα��� = dblʣ�� * rs���ö�("���ö�") '��ȫ������
                    dbl�ֶα��� = dblʣ�� * rs���ö�("����") / 100
                Else
                    '��ʣ��ֵ�뱾�οռ�֮��ѡ��Сֵ
                    dbl�ֶα��� = dbl���� - dbl���
                    'If dbl�ֶα��� > dblʣ�� * rs���ö�("���ö�") Then dbl�ֶα��� = dblʣ�� * rs���ö�("���ö�")
                    If dbl�ֶα��� > dblʣ�� * rs���ö�("����") / 100 Then
                        dbl�ֶα��� = dblʣ�� * rs���ö�("����") / 100
                    End If
                End If
                '��������öο��Խ�������ͳ�����
                'dbl�ֶν��� = dbl�ֶα��� / rs���ö�("���ö�")
                dbl�ֶν��� = dbl�ֶα��� / (rs���ö�("����") / 100)
                
                dbl��� = dbl��� + dbl�ֶα���
                dblʣ�� = dblʣ�� - dbl�ֶν���
                
                '����������öεı������
                dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
                If m����.�Ҷ� < deg����֧�� Then
                    dbl�ֶα��� = 0
                Else
                    dbl�ֶα��� = Val(Format(dbl�ֶν��� * rs���ö�("����") / 100, "0.00")) '���Ǹö������Ա����Ľ��
                End If
                
                If dbl��֧����� + dbl�ֶα��� > m����.�ⶥ�� And m����.�ⶥ�� <> 0 Then
                    '���������˷ⶥ�ߣ����Ҵ��ڷⶥ������
                    dbl�ֶα��� = m����.�ⶥ�� - dbl��֧�����
                    
                    '���ƽ���ͳ����
                    If rs���ö�("����") <> 0 Then
                        dbl�ֶν��� = dbl�ֶα��� * 100 / rs���ö�("����")
                    Else
                        dbl�ֶν��� = 0
                    End If
                End If
                
            End If
        End If
        
        dbl��֧����� = dbl��֧����� + dbl�ֶα���
        
        '���Ρ�����ͳ���ͳ�ﱨ��������
        '���и�ʽ��
        dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
        dbl�ֶα��� = Val(Format(dbl�ֶα���, "0.00"))
        lng���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        dblTemp = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        gcol�������.Add Array(lng����, dbl�ֶν���, dbl�ֶα���, dblTemp)
        
        dbl���ν��� = dbl�ֶν��� + dbl���ν���
        dbl���α��� = dbl���α��� + dbl�ֶα���
        rs���ö�.MoveNext
    Loop
    
    m����.ͳ�����֧�� = dbl���α���
    m����.ͳ������Ը� = dbl���ν��� - dbl���α���
    
    Calc�����ֶ�3 = True
End Function

Private Function Calc�����ֶ�4(rs���ö� As ADODB.Recordset) As Boolean
'���ܣ����㰴���÷ֶΣ�֧���ⶥ�����
    Dim dbl���� As Currency       'ÿһ�ε����ֵ�������Ƿ��ã�Ҳ������֧�����
    Dim dbl���� As Currency       'ÿһ�ε����ֵ�������Ƿ��ã�Ҳ������֧�����
    Dim dbl�ֶν��� As Currency   '����ĳһ�ε�ͳ����
    Dim dbl�ֶα��� As Currency   '����ĳһ�ε�ͳ�ﱨ�����
    Dim dbl���ν��� As Currency   '�����ܵĽ���ͳ����
    Dim dbl���α��� As Currency   '�����ܵĽ��뱨�����
    
    Dim dbl��� As Currency  '���ڼ�������ֵ
    Dim dblʣ����� As Currency  '���������õķ���
    Dim dblʣ��ͳ�� As Currency  '���������õ�ͳ����
    
    Dim dblTemp As Currency, lng���� As Long
    Dim dbl���� As Currency
    
    dbl���� = m����.��������
    If m����.�ⶥ�� > 0 Then
        '�������������ʹ�õķ���
        dblʣ����� = m����.�ⶥ�� - m����.ͳ����֧������
        If dblʣ����� < 0 Then dblʣ����� = 0
        
        '������ⲿ�ַ����е�ͳ����
        If dblʣ����� > m����.ҽ����Ŀ��� Then
            dblʣ��ͳ�� = m����.ҽ����Ŀ���
            dblʣ����� = dblʣ����� - m����.ҽ����Ŀ���
            
            If dblʣ����� > m����.������Ŀ��� Then
                dblʣ��ͳ�� = dblʣ��ͳ�� + m����.������Ŀ��� * 0.8
            Else
                'Modified by zyb ##2003-08-31
                '�����ⶥ��,����Calc����ͳ��()�м���������Ը���ʵ�ʵ������Ը�������,��Ҫ���¼���
                dblʣ��ͳ�� = dblʣ��ͳ�� + dblʣ����� * 0.8 '����ʹ��һ����ֵ
                m����.�����Ը� = dblʣ����� * 0.2
            End If
        Else
            dblʣ��ͳ�� = dblʣ�����
        End If
    Else
        dblʣ��ͳ�� = m����.����ͳ��
    End If
    
    '���Ȱ����߽��۳�����Ϊ������Զ���ܱ����ģ����Բ��ܷ�����һ��ȥ�жϡ��о��У����оͲ��С�
    If dblʣ��ͳ�� > dbl���� Then
        '���۳�
        m����.ʵ������ = dbl����
        dbl��� = m����.ͳ����֧�����   '˵���Ѿ�֧������
        dblʣ��ͳ�� = dblʣ��ͳ�� - m����.ʵ������
    Else
        '�����߶�������֧����ֱ���˳�
        m����.ʵ������ = dblʣ��ͳ��
        
        Do Until rs���ö�.EOF
            lng���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
            dblTemp = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
                
            gcol�������.Add Array(lng����, 0, 0, dblTemp)
            rs���ö�.MoveNext
        Loop
        Calc�����ֶ�4 = True
        Exit Function
    End If
    
    Do Until rs���ö�.EOF
        dbl�ֶν��� = 0
        dbl�ֶα��� = 0
        dbl���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        dbl���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        
        '�����Լ�������
        If dbl��� >= dbl���� And (dbl��� < dbl���� Or dbl���� = 0) And dblʣ��ͳ�� > 0 Then
            '�ö���ǰ��δ������ȫ�����������Ҫ����۳��Ľ��Ѿ�������ĶΣ����߽����Ķβ�����룩
            If dbl���� = 0 Then
                'dbl�ֶα��� = dblʣ��ͳ�� * rs���ö�("���ö�") '��ȫ������
                dbl�ֶα��� = dblʣ��ͳ�� * rs���ö�("����") / 100 '��ȫ������
            Else
                '��ʣ��ֵ�뱾�οռ�֮��ѡ��Сֵ
                dbl�ֶα��� = dbl���� - dbl���
                'If dbl�ֶα��� > dblʣ��ͳ�� * rs���ö�("���ö�") Then dbl�ֶα��� = dblʣ��ͳ�� * rs���ö�("���ö�")
                If dbl�ֶα��� > dblʣ��ͳ�� * rs���ö�("����") / 100 Then
                    dbl�ֶα��� = dblʣ��ͳ�� * rs���ö�("����") / 100
                End If
            End If
            '��������öο��Խ�������ͳ�����
            'dbl�ֶν��� = dbl�ֶα��� / rs���ö�("���ö�")
            dbl�ֶν��� = dbl�ֶα��� / (rs���ö�("����") / 100)
            
            dbl��� = dbl��� + dbl�ֶα���
            dblʣ��ͳ�� = dblʣ��ͳ�� - dbl�ֶν���
        End If
        
        '���Ρ�����ͳ���ͳ�ﱨ��������
        '���и�ʽ��
        dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
        dbl�ֶα��� = Val(Format(dbl�ֶα���, "0.00"))
        lng���� = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        dblTemp = IIf(IsNull(rs���ö�("����")), 0, rs���ö�("����"))
        gcol�������.Add Array(lng����, dbl�ֶν���, dbl�ֶα���, dblTemp)
        
        dbl���ν��� = dbl�ֶν��� + dbl���ν���
        dbl���α��� = dbl���α��� + dbl�ֶα���
        rs���ö�.MoveNext
    Loop
    
    m����.ͳ�����֧�� = dbl���α���
    m����.ͳ������Ը� = dbl���ν��� - dbl���α���
    
    Calc�����ֶ�4 = True
End Function

Private Function Calc���ز�() As Boolean
'���ܣ�����������������󲡲��˵���ͨ����ͳ����
'���������
'���������
'���أ��ɹ����㣬�򷵻�True
    
    On Error GoTo errHandle
    
    Calc���ز� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Calc���䱨��() As Boolean
'���ܣ������סԺ���ˡ�����������󲡲��˵Ĳ��䱨�����
'���������
'���������
'���أ��ɹ����㣬�򷵻�True
    Dim rsTemp As New ADODB.Recordset
    Dim bln���÷ⶥ As Boolean, dbl���� As Currency, dbl�޶� As Currency
    Dim dblʣ��ҽ�� As Currency, dblʣ������ As Currency
    Dim dbl������� As Currency, dbl����֧�� As Currency, dblʣ��ͳ�� As Currency, dbl������ As Currency
    
    m����.�μӲ��䱣�� = 0
    On Error GoTo errHandle
    gstrSQL = "SELECT A.��չ���䱣�ձ���,A.���䱨������,A.���䱨���޶�,A.���䱨���޶����� " & _
               " FROM ��������Ŀ¼ A " & _
               " WHERE A.����=" & TYPE_������ & " AND A.����='" & gIC����.CenterCode & "'"
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp("��չ���䱣�ձ���") = 0 Then
        '����չ���䱣�ձ���ҵ��
        Calc���䱨�� = True
        Exit Function
    End If
    
    bln���÷ⶥ = Nvl(rsTemp("���䱨���޶�����")) = 1
    dbl���� = rsTemp("���䱨������")
    dbl�޶� = rsTemp("���䱨���޶�")
    
    gstrSQL = "Select * From ������Ա Where ���Ĵ���='" & gIC����.CenterCode & "' and to_number(ְ������)=" & Val(TrimStr(gIC����.MediAccountNo))
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        '����û�в������䱣��
        Calc���䱨�� = True
        Exit Function
    End If
    
    m����.�μӲ��䱣�� = 1
    
    
    '�õ����뱻��εķ���
    With m����
        If bln���÷ⶥ = True Then
            'ͳ����֧������Ӧ�ôӿ��ϵõ�����Ϊ���䱣�մ�������Ҫ�ۼƵ�
            '���÷ⶥ�����������ȵõ�������ⶥ�ߵķ���
            If .ͳ����֧������ + .����ͳ�� + .�����Ը� > dbl�޶� Then
                '���ⶥ��
                .������ⶥ�� = .ͳ����֧������ + .����ͳ�� + .�����Ը� - dbl�޶�
            End If
            '���벹��ķ�����������ʹ�õ�ͳ����
            dbl������� = .�������� - .ȫ�Է� - .����ͳ����� - .������ⶥ��
            
            If m����.ҽ����Ŀ��� > .����ͳ����� Then
                'ҽ����Ŀ�Ľ���Ѿ������˻���ͳ������Ҫ�Ľ��
                dblʣ��ҽ�� = m����.ҽ����Ŀ��� - .����ͳ�����
                dblʣ������ = .������Ŀ���
            Else
                dblʣ��ҽ�� = 0
                dblʣ������ = .������Ŀ��� - (.����ͳ����� - .ҽ����Ŀ���)
            End If
            
            If dbl������� > dblʣ��ҽ�� Then
                dblʣ��ͳ�� = dblʣ��ҽ��
                dbl������� = dbl������� - dblʣ��ҽ��
                
                If dbl������� > dblʣ������ Then
                    dblʣ��ͳ�� = dblʣ��ͳ�� + dblʣ������ * 0.8
                Else
                    dblʣ��ͳ�� = dblʣ��ͳ�� + dbl������� * 0.8 '����ʹ��һ����ֵ
                End If
            Else
                dblʣ��ͳ�� = dbl�������
            End If
            
            If m����.���䱨�����𸶽� = True Then
                If dblʣ��ͳ�� > .�������� - .ʵ������ Then
                    '��ȡ��������
                    dblʣ��ͳ�� = dblʣ��ͳ�� - (.�������� - .ʵ������)
                    .ʵ������ = .��������
                Else
                    'ֻ��֧����������
                    .ʵ������ = .ʵ������ + dblʣ��ͳ��
                    dblʣ��ͳ�� = 0
                End If
            End If
            
            .�������֧�� = dblʣ��ͳ�� * dbl����
            .��������Ը� = dblʣ��ͳ�� - dblʣ��ͳ�� * dbl����
            
        Else
            '֧���ⶥ
            dblʣ��ͳ�� = .�������ⶥ�� '����ͳ����
            If m����.���䱨�����𸶽� = True Then
                If dblʣ��ͳ�� > .�������� - .ʵ������ Then
                    '��ȡ��������
                    dbl������ = (.�������� - .ʵ������)
                    dblʣ��ͳ�� = dblʣ��ͳ�� - dbl������
                    .ʵ������ = .��������
                Else
                    'ֻ��֧����������
                    dbl������ = dblʣ��ͳ��
                    dblʣ��ͳ�� = 0
                    .ʵ������ = .ʵ������ + dbl������
                End If
            End If
            
            dbl����֧�� = dblʣ��ͳ�� * dbl����     '֧���ⶥ�г��ⶥ��ȫ�ǽ���ͳ��Ľ��
            If dbl����֧�� > dbl�޶� - .ͳ����֧����� - .ͳ�����֧�� Then
                '�Ѿ������ܱ�������
                dbl����֧�� = dbl�޶� - .ͳ����֧����� - .ͳ�����֧��
                If dbl����֧�� < 0 Then dbl����֧�� = 0 '�������Ѿ������޶���
                dbl������� = dbl����֧�� / dbl����
            Else
                dbl������� = dblʣ��ͳ��
            End If
            
            .�������֧�� = dbl����֧��
            .��������Ը� = dbl������� - dbl����֧��
            .������ⶥ�� = .�������ⶥ�� - dbl������� - dbl������
        End If
        
        '��Ҫ��ס����Ҫ��֧�ֲ���Ǽǵ�����²Ÿı�
        .����ͳ��֧�� = .����ͳ��֧�� + .�������֧��
        .����ͳ����� = .����ͳ����� + dbl������� '��һ����ҲҪ�ۼ�
    End With
    Calc���䱨�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Calc��������() As Boolean
'���ܣ������סԺ���˵Ĳ����������������ǹ���Ա��
'���������
'���������
'���أ��ɹ����㣬�򷵻�True
    Dim dbl�ܷ��� As Currency, dbl���Ը� As Currency
    Dim dbl��ʼֵ As Currency, dbl��ֵֹ As Currency, dbl���� As Currency
    Dim dbl�����Ը� As Currency, dbl����֧�� As Currency
    Dim dbl�ֶα��� As Currency, dbl�ֶη��� As Currency
    Dim rs���� As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "SELECT A.��չ�������� " & _
               " FROM ��������Ŀ¼ A " & _
               " WHERE A.����=" & TYPE_������ & " AND A.����='" & gIC����.CenterCode & "'"
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp("��չ��������") = 0 Then
        '����չ���䱣�ձ���ҵ��
        Calc�������� = True
        Exit Function
    End If
    
    gstrSQL = "Select ��ֵ,���� From ���ղ������� Where ����=" & TYPE_������ & _
            " And ����=" & m����.������� & " and ���=" & m����.��� & " Order by ��ֵ"
    rs����.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    
    With m����
        If .�μӲ��䱣�� = 1 Then
            dbl�ܷ��� = .�����Ը� + .����ͳ�� - .������ⶥ��
        Else
            dbl�ܷ��� = .�����Ը� + .����ͳ�� - .�������ⶥ��
        End If
        dbl���Ը� = dbl�ܷ��� - .ͳ�����֧�� - .�������֧��
        
        '�ֶμ��㡣��ֵ��һ������������ dbl���Ը�/dbl�ܷ���
        Do Until rs����.EOF
            If rs����.AbsolutePosition = 1 Then
                '��һ��ֻ��Ϊ��ʼֵ
                dbl��ʼֵ = dbl�ܷ��� * rs����("��ֵ")
                dbl���� = rs����("����")
            Else
                dbl��ֵֹ = dbl�ܷ��� * rs����("��ֵ")
                If dbl���Ը� > dbl��ʼֵ Then
                    If dbl���Ը� <= dbl��ֵֹ Then
                        dbl�ֶη��� = dbl���Ը� - dbl��ʼֵ
                    Else
                        dbl�ֶη��� = dbl��ֵֹ - dbl��ʼֵ
                    End If
                    
                    dbl�ֶα��� = dbl�ֶη��� * dbl����
                    m����.��������֧�� = m����.��������֧�� + dbl�ֶα���
                    m����.���������Ը� = m����.���������Ը� + dbl�ֶη��� - dbl�ֶα���
                End If
                '��Ϊ��һ�ε���ʼֵ
                dbl��ʼֵ = dbl��ֵֹ
                dbl���� = rs����("����")
            End If
            
            If rs����.AbsolutePosition = rs����.RecordCount Then
                '���һ��
                If dbl���Ը� > dbl��ʼֵ Then
                    dbl�ֶη��� = dbl���Ը� - dbl��ʼֵ
                    dbl�ֶα��� = dbl�ֶη��� * dbl����
                    m����.��������֧�� = m����.��������֧�� + dbl�ֶα���
                    m����.���������Ը� = m����.���������Ը� + dbl�ֶη��� - dbl�ֶα���
                End If
            
            End If
            rs����.MoveNext
        Loop
    End With
    
    Calc�������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ReadIC(ByVal ����ID As Long, ByVal ���� As Integer, ByVal ��鿨��ȷ As Boolean, ByVal ������Ϣ As String _
                       , ic�� As TIC����, ���ݲ��� As Boolean) As Boolean
'���ܣ��Ӷ����������ݿ⡢Զ�̵õ����˵���Ϣ
'�������������ID           �����жϲ����Ƿ������ݲ���
'          ����             1-�����շѡ�2-סԺ
'          ��鿨��ȷ       ����Ҫ���д��������ҵ���������շѣ�����Ҫ�ж��Ƿ��Ǹò��˵Ŀ�
'          ������Ϣ         Ϊ�˸�׼ȷ����ʾ������Ϣ
'���������ic��             �������IC����Ϣ
'          ���ݲ���         ��ǰ�����Ƿ�������Ա
'���أ��ɹ���ȡ������True
    Dim strҽ���� As String, str���֤�� As String, str���� As String
    Dim lngReturn As Long
    Dim blnԶ����֤ As Boolean, strԶ�̵�ַ As String
    
    On Error GoTo errHandle
    
    If Get���ղ���_����(blnԶ����֤, strԶ�̵�ַ) = False Then
        Exit Function
    End If
    
    If Get�ʻ���Ϣ(����ID, strҽ����, str���֤��, str����) = False Then Exit Function
    ���ݲ��� = Is���ݲ���(����ID)
    
    If ���ݲ��� = False Then
        If blnԶ����֤ = False Then
            If ReadICCard(ic��) <> 0 Then
                MsgBox ������Ϣ, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gIC����Temp.IDCardno = str���֤��
            If frmSock����.CommIC(strԶ�̵�ַ, True, ����, str���֤�� & "|" & str����) = False Then
                Exit Function
            End If
            ic�� = gIC����Temp
        End If
        If ic��.InpatientFlag = "1" And ���� = 0 Then
            MsgBox "�ò�����Ȼ��Ժ�����ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
        
        If ��鿨��ȷ = True Then
            '�жϸò��˵Ŀ��Ƿ������ȷ
            If ���IC��(����ID, TrimStr(ic��.Cardno), TrimStr(ic��.CenterCode)) = False Then Exit Function
        End If
    Else
        If Get���ݲ���_����(strҽ����, ic��) = False Then
            Exit Function
        End If
    End If
    
    ReadIC = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function WriteIC(ByVal ���ݲ��� As Boolean, ByVal �շ���־ As Boolean, ByVal ���� As Integer, ByVal Insert���� As String, ic�� As TIC���� _
    , payLog As TPayInfo, ByVal ������ As String) As Boolean
'���ܣ��Ӷ����������ݿ⡢Զ�̵õ����˵���Ϣ
'������������ݲ���         ��������ݲ��ˣ��򲻽���д��
'          �շ���־         ������Ժ��Ժ��д�����Ͳ���Ҫд��־
'          ����             0-����;1-סԺ
'���������ic��             ׼��д���IC����Ϣ
'          payLog           ׼��д�����־��Ϣ
'���أ��ɹ���ȡ������True
    Dim lngReturn As Long
    Dim blnԶ����֤ As Boolean, strԶ�̵�ַ As String
    
    If Get���ղ���_����(blnԶ����֤, strԶ�̵�ַ) = False Then
        Exit Function
    End If
    
    gcn����.BeginTrans
    On Error GoTo errHandle
    '����������ݿ�Ĳ���
    If Insert���� <> "" Then gcn����.Execute Insert����
    
    If ���ݲ��� = False Then
        '����д��
        If blnԶ����֤ = False Then
            lngReturn = WriteICCard(ic��)
            If lngReturn <> 0 Then
                gcn����.RollbackTrans
                MsgBox "д�뿨ʧ�ܡ�" & ������Ϣ_����(lngReturn), vbInformation, gstrSysName
                Exit Function
            End If
            If �շ���־ = True Then
                '��¼������־�������һ������Ϣ����̫��Ҫ����ʹ����Ҳ���Ժ��ԣ������ܻع�ǰһ��д��
                On Error Resume Next
                lngReturn = WriteICCardPayInfo(ic��.Cardno, payLog)
            End If
        ElseIf ������ <> "" Then
            '������Զ�̿����⣬��Ҫ���շѲ���
            If frmSock����.CommIC(strԶ�̵�ַ, False, ����, ������) = False Then
                gcn����.RollbackTrans
                Exit Function
            End If
        End If
    End If
    
    gcn����.CommitTrans
    
    WriteIC = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcn����.RollbackTrans
End Function

Public Function Get���ղ���_����(�Ƿ�Զ�� As Boolean, Զ�̵�ַ As String) As Boolean
'���ܣ���ñ��ղ���
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "select A.������,A.����ֵ from ���ղ��� A " & _
              " where A.����=[1] and A.���� is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������)

    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "���������֤"
                �Ƿ�Զ�� = Nvl(rsTemp("����ֵ")) = "��"
            Case "ҽ�����ĵ�ַ"
                Զ�̵�ַ = Nvl(rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    
    Get���ղ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'Modified by ���� 2004-01-07
Public Function �ҺŽ���_����(ByVal lng����ID As Long) As Boolean
    Dim lng����ID As Long
    Dim cur�ܶ� As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    'ȡ����ID
    gstrSQL = "Select ����ID From ������ü�¼ Where ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID)
    lng����ID = rsTemp!����ID
    
    'ȡ�����ܶ�
    gstrSQL = "Select Sum(ʵ�ս�� ) as ��� From ������ü�¼ Where ����ID= [1] And Nvl(��¼״̬,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�����ܶ�", lng����ID)
    cur�ܶ� = rsTemp!���
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        cur�ܶ� & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ",0," & _
        0 & "," & 0 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Һ�����")
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
    
    �ҺŽ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

'Modified by ���� 2004-01-07
Public Function �ҺŽ������_����(ByVal lng����ID As Long) As Boolean
    Dim lngԭ����ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    lngԭ����ID = lng����ID
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
    lng����ID = rsTemp("����ID")
    
    '��ȡԭ�����¼��Ϣ
    gstrSQL = "Select ����ID,�������ý��,ȫ�Ը���� from ���ս����¼ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭ�����¼��Ϣ", lngԭ����ID)
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & rsTemp!����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & 0 & "," & 0 & "," & 0 & ",0," & _
        0 & "," & 0 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Һ�����")
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
    
    �ҺŽ������_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

'Modified By ���� 2003-12-10 ����������
Private Function Get��ǰҽ����() As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ҽ���� From �������� A,��������Ŀ¼ B" & _
              " Where A.����=" & TYPE_������ & " And A.����=B.�������� And B.���=" & m����.�������
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn����
    End With
    
    Get��ǰҽ���� = Nvl(rsTemp!ҽ����)
End Function

'Modified By ���� 2003-12-10 ����������
Private Function Get����IP() As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select װǮIP��ַ1 IP From ������������ A,��������Ŀ¼ B " & _
             " Where A.����=" & TYPE_������ & " And A.����=B.�������� And B.���=" & m����.�������
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn����
    End With
    
    Get����IP = Nvl(rsTemp!IP)
End Function

Public Function ��ݱ�ʶ_����2(ByVal strCard As String, ByVal strPass As String, Optional lng����ID As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    Dim mIC���� As TIC����
    Dim blnԶ����֤ As Boolean, strԶ�̵�ַ As String
    Dim rsTemp As New ADODB.Recordset
    
    If Get���ղ���_����(blnԶ����֤, strԶ�̵�ַ) = False Then
        Exit Function
    End If
    
    If strCard <> "1" Then
        If blnԶ����֤ = False Then
            lngReturn = ReadICCard(mIC����)
        Else
            'Զ������
            If Trim(strPass) = "" Then
                Exit Function
            End If
            If frmSock����.CommIC(strԶ�̵�ַ, True, 0, strPass & "|" & strNewPass) = False Then
                Exit Function
            End If
            mIC���� = gIC����Temp
        End If
    Else
        '�������嵥�ж�ȡ�������������IC���ṹ��
        If Get���ݲ���_����(strPass, mIC����, False) = False Then
            MsgBox "û���ҵ���ҽ�����˵Ļ�����Ϣ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If lngReturn <> 0 Then Exit Function
    
    'H�� ����У���Ƿ���ȷ����֤IC����Password�������PasswordΪ9000�������������֤����
    If TruncZero(mIC����.Password) <> "9000" Then
        If blnԶ����֤ = False Then
            If TruncZero(mIC����.Password) <> strPass Then
                MsgBox "�����������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '��ȡҽ�����˵�ID
    gstrSQL = "Select ����ID From �����ʻ� Where ����=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˵�ID", TYPE_������, CStr(mIC����.Cardno))
    If rsTemp.EOF Then
        MsgBox "�ò������κη��ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rsTemp!����ID
    ��ݱ�ʶ_����2 = lng����ID
End Function


