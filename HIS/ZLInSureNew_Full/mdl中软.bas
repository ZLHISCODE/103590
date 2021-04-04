Attribute VB_Name = "mdl����"
Option Explicit
'һ��IC����������ṹ����
'1�������ṹ:
'      1��������Ϣ�ṹ       TIC����
'      2��IC����ҽ��Ϣ�ṹ   TBlockPayInfo    �����֧����Ϣ��
'2��ҵ��ṹ
'      1������ҺŶ����ṹ   TRegisterResult
'      2�������շѶ����ṹ   TChargeResult
'      3�������շ�д��ṹ   TChargeParameter
'      4��סԺ�ǼǶ������ṹ TInpatientRegResult
'      5��סԺ�Ǽ�д���ṹ   TInpatientRegParameter
'      6����Ժ����д���ṹ   TInpatientPayParameter
'      7��װǮд���ṹ       TInMoneyParameter
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
    InExAcc          As Double          ' �������ۼ�ע����
    OutExAcc         As Double          ' �����ʻ��ۼ�֧�����
    InSubAcc         As Double          ' �����ʻ��ۼ�ע����
    OutSubAcc        As Double          ' �����ʻ��ۼ�֧�����
    OutAnnPlan       As Double          ' ����ͳ��֧������ۼ�
    OutAnnOverLine   As Double          ' �������ͳ�����ۼ�
    Password         As String * 4      ' ��������
    AnnInpatientTimes As Long           ' ������ЧסԺ����
    InpatientFlag    As String * 1      ' סԺ��־ 0-��סԺ 1-סԺ
    HasSubInsurance  As String * 1      ' �Ƿ�μӹ���Ա��������  0-��  ����-��
    HasExInsurance   As String * 1      ' �Ƿ�μӲ��䱣��0����1��
    HasBigIllness    As String * 1      ' �Ƿ�μӴ�ҽ��
End Type
Private Type TPayInfo
    OccurDate        As String * 8 '  ��ҽ����
    HospitalCode     As String * 4 '  ҽ�ƻ�������
    Amount           As Double     '  ���η��úϼ�
    AccPay           As Double     '  �����ʻ�֧��
End Type
Private Type TBlockPayInfo
    First            As TPayInfo   ' ��һ�ξ�ҽ��Ϣ
    Second           As TPayInfo   ' �ڶ��ξ�ҽ��Ϣ
    Third            As TPayInfo   ' �����ξ�ҽ��Ϣ
End Type
Private Type TRegisterResult
    CenterCode       As String * 4 ' ���Ĵ���
    Cardno           As String * 8 ' ����
    Name             As String * 10 ' ����
    Sex              As String * 1 ' �Ա� 1-��  0-Ů
    Birthday         As String * 8 ' �������� YYYYMMDD
    MediAccountNo    As String * 8 ' ҽ����
    UnitCode         As String * 5 ' ���˵�λ����
    ClassCode        As String * 2 ' ְ����� 0X-��ְ 1X-����
    DomainCode       As String * 1 ' ְ������ 0-���� 1-��פ��� 2-��ذ���
    Password         As String * 4 ' ��������
    MediYear         As String * 4 ' ҽ�����
    InNo             As Long       ' װǮ�ڴ�
    InPerAcc         As Double     ' �����ʻ��ۼ�ע����
    InExAcc          As Double     ' �������ۼ�ע����
    InSubAcc         As Double     ' �����ʻ��ۼ�ע����
    OutPerAcc        As Double     ' �����ʻ��ۼ�֧�����
    OutExAcc         As Double     ' �����ʻ��ۼ�֧�����
    OutSubAcc        As Double     ' �����ʻ��ۼ�֧�����
    InpatientFlag    As String * 1 ' סԺ��־ 0-��סԺ 1-סԺ
End Type
Private Type TChargeResult
    CenterCode       As String * 4 ' ���Ĵ���
    Cardno           As String * 8 ' ����
    Name             As String * 10 ' ����
    Sex              As String * 1 ' �Ա� 1-��  0-Ů
    Birthday         As String * 8 ' �������� YYYYMMDD
    MediAccountNo    As String * 8 ' ҽ����
    UnitCode         As String * 5 ' ���˵�λ����
    ClassCode        As String * 2 ' ְ����� 0X-��ְ 1X-����
    DomainCode       As String * 1 ' ְ��״̬ 0-���� 1-��פ���
    Password         As String * 4 ' ��������
    MediYear         As String * 4 ' ҽ�����
    InNo             As Long       ' װǮ�ڴ�
    InPerAcc         As Double     ' �����ʻ��ۼ�ע����
    InExAcc          As Double     ' �������ۼ�ע����
    InSubAcc         As Double     ' �����ʻ��ۼ�ע����
    OutPerAcc        As Double     ' �����ʻ��ۼ�֧�����
    OutExAcc         As Double     ' �����ʻ��ۼ�֧�����
    OutSubAcc        As Double     ' �����ʻ��ۼ�ע����
    OutSerialNo      As Long       ' ֧��˳���
    InpatientFlag    As String * 1 ' סԺ��־
End Type
Private Type TChargeParameter
    Cardno           As String * 8 ' ����
    OutPerAcc        As Double     ' �����ʻ��ۼ�֧�����
    OutExAcc         As Double     ' �����ʻ��ۼ�֧�����
    OutSubAcc        As Double     ' �����ʻ��ۼ�֧�����
    OutSerialNo      As Long       ' ֧��˳���
    PayOccurDate     As String * 8 ' ����
    PayHospitalCode  As String * 4 ' ҽԺ����
    PayAccPay        As Double     ' �����ʻ�֧��
    PayAmount        As Double     ' �ܶ�
End Type
Private Type TInpatientRegResult
    CenterCode       As String * 4 ' ���Ĵ���
    Cardno           As String * 8 ' ����
    IDCardno         As String * 18 ' ���֤�� ���Ȳ����#0
    MediAccountNo    As String * 8 ' ҽ����
    Name             As String * 10 ' ����
    Sex              As String * 1 ' �Ա� 1-��  0-Ů
    Birthday         As String * 8 ' �������� YYYYMMDD
    UnitCode         As String * 5 ' ���˵�λ����
    ClassCode        As String * 2 ' ְ����� 0X-��ְ 1X-����
    DomainCode       As String * 1 ' ְ��״̬ 0-���� 1-��פ���
    MediYear         As String * 4 ' ҽ�����
    InNo             As Long       ' װǮ�ڴ�
    OutSerialNo      As Long       ' ֧��˳���
    InPerAcc         As Double     ' �����ʻ��ۼ�ע����
    OutPerAcc        As Double     ' �����ʻ��ۼ�֧�����
    InExAcc          As Double     ' �������ۼ�ע����
    OutExAcc         As Double     ' �����ʻ��ۼ�֧�����
    InSubAcc         As Double     ' �����ʻ��ۼ�ע����
    OutSubAcc        As Double     ' �����ʻ��ۼ�֧�����
    OutAnnPlan       As Double     ' ����ͳ��֧������ۼ�
    OutAnnOverLine   As Double     ' �������ͳ�����ۼ�
    Password         As String * 4 ' ��������
    AnnInpatientTimes As Long       ' ������ЧסԺ����
    InpatientFlag    As String * 1 ' סԺ��־ 0-��סԺ 1-סԺ
    HasSubInsurance  As String * 1 ' ����Ա��־  0-��  ����-��
    HasExInsurance   As String * 1 ' �Ƿ�μӲ��䱣��
    HasBigIllness    As String * 1 ' �Ƿ�μӴ�ҽ��
End Type
Private Type TInpatientRegParameter
    Cardno           As String * 8 ' ����
    InpatientFlag    As String * 1 ' סԺ��־ 0-��סԺ 1-סԺ
End Type
Private Type TInpatientPayParameter
    Cardno           As String * 8 ' ����
    OutPerAcc        As Double     ' �����ʻ��ۼ�֧�����
    OutExAcc         As Double     ' �����ʻ��ۼ�֧�����
    OutSubAcc        As Double     ' �����ʻ��ۼ�֧�����
    OutSerialNo      As Long       ' ֧��˳���
    OutAnnOverLine   As Double     ' �����𸶶����ϻ���ҽ�Ʒ�
    OutAnnPlan       As Double     ' ����ͳ��֧������ۼ�
    InpatientFlag    As String * 1 ' סԺ��־ 0-��סԺ 1-סԺ
    AnnInpatientTimes As Long       ' ������ЧסԺ����
    PayOccurDate     As String * 8 ' ����
    PayHospitalCode  As String * 4 ' ҽԺ����
    PayAccPay        As Double     ' �����ʻ�֧��
    PayAmount        As Double     ' �ܶ�
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
Private Type TPayLog
    OccurDate        As String * 8   '  ��ҽ����
    HospitalCode     As String * 4   '  ҽ�ƻ�������
    Amount           As String * 8   '  ���η��úϼ�
    AccPay           As String * 8   '  �����ʻ�֧��
End Type
'��¼סԺ���
Private Declare Function ChargeLog Lib "ICAPI.DLL" (payLog As TPayLog) As Long
'
''����IC����д��������˵��
''1����ʼ��
''      1������ԭ(��IC����PIN��ԭ�ɳ�ʼֵ)
'Private Declare Function ReturnICCard Lib "ICWRITE.DLL" () As Long
''      2����IC��
'Private Declare Function MakeICCard Lib "ICWRITE.DLL" (iIC���� As TIC����) As Long
'
''2��������д
''      1����IC��������Ϣ
Private Declare Function ReadICCard Lib "ICREAD.DLL" (iIC���� As TIC����) As Long
''      2��дIC��������Ϣ
Private Declare Function WriteICCard Lib "ICWRITE.DLL" (iIC���� As TIC����) As Long
''      3����IC����ҽ��Ϣ
'Private Declare Function ReadICCardPayInfo Lib "ICREAD.DLL" (BlockPayInfo As TBlockPayInfo) As Long
'
''3��ҵ���д
''      1���ҺŶ���
'Private Declare Function RegisterRead Lib "ICAPI.DLL" (RegisterResult As TRegisterResult) As Long
''      2�������շѶ���
'Private Declare Function ChargeRead Lib "ICAPI.DLL" (ChargeResult As TChargeResult) As Long
''      3�������շ�д��
Private Declare Function ChargeWrite Lib "ICAPI.DLL" (ChargeParameter As TChargeParameter) As Long
''      4��סԺ�ǼǶ���
Private Declare Function InpatientRegRead Lib "ICAPI.DLL" (InpatientRegResult As TInpatientRegResult) As Long
''      5��סԺ�Ǽ�д��
Private Declare Function InpatientRegWrite Lib "ICAPI.DLL" (InpatientRegParameter As TInpatientRegParameter) As Long
''      6��סԺ�нᡢ����д��
'Private Declare Function InpatientPayWrite Lib "ICAPI.DLL" (InpatientPayParameter As TInpatientPayParameter) As Long
'
''4���޸������װǮ
''      1���޸�����
'Private Declare Function ChangePassword Lib "ICAPI.DLL" (Cardno As Variant, Password As Variant) As Long
''      2�����겻װǮ��ʼ��:��֧������,�ۼ�����,ע��Ϊ�����,֧���ż�1
''         ֻ��CardNo , MediYear�ֶ�
'Private Declare Function YearInitICCard Lib "ICAPI.DLL" (InMoneyParameter As TInMoneyParameter) As Long
''      3������װǮ��ʼ��:��֧������,�ۼ�����,ע��Ϊ����ע����, ֧���ż�1
''         ��CardNo, MediYear, InNo, InPerAcc, InSubAcc, InExAcc�ֶ�.
'Private Declare Function YearInitICCardWithInMoney Lib "ICAPI.DLL" (InMoneyParameter As TInMoneyParameter) As Long
''      4������ҽ����װǮ
''         ��CardNo, InNo, InPerAcc, InSubAcc, InExAcc�ֶ�.
'Private Declare Function InMoney Lib "ICAPI.DLL" (InMoneyParameter As TInMoneyParameter) As Long
'
''�������װǮ
Private Declare Function OnLineInMoney Lib "InMoneyOnLine.dll" (ByVal IC_CenterCode As String, ByVal IC_CardNo As String, ByVal IC_MediYear As String, ByVal HosCode As String) As Long


Private Enum cardҽ���Ҷ�
    degֹ֧ͣ�� = 1
    deg�ϴ���ϸ = 2 'Ҳֹ֧ͣ��
    deg����֧�� = 3 '���ø����ʻ�֧����ͳ��ͣ
    degҽ��֧�� = 4 '
    deg����֧�� = 5 '���·�
End Enum

'-------------��������

Public gIC���� As TIC����                 'ȫ�ֶ���Ĵ洢IC����Ϣ�Ľṹ
Public gcn���� As New ADODB.Connection        '���ӵ�ҽ��ǰ�÷�����

'-------------��������

Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false

    ҽ����ʼ��_���� = True
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������strSelfNO-���˱�ţ�ˢ���õ���strSelfPwd-�������룻bytType-ʶ�����ͣ�0-���1-סԺ
'���أ� �ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim strIdentify As String, strAddition As String
    Dim strBirthday As String, datToday As Date
    Dim lng���� As Long, lng���� As Long, str���� As String
    Dim rsTemp As New ADODB.Recordset, rs���� As ADODB.Recordset
    Dim lng�Ҷ� As cardҽ���Ҷ�
    
    On Error GoTo errHandle
    
    If frmIdentify����.GetPatient(bytType <> 2) = True Then
        '���ʶ����ɣ����ز�����Ϣ
        With gIC����
            lng�Ҷ� = ҽ���Ҷ�(.CenterCode, .Cardno)
            If lng�Ҷ� = degֹ֧ͣ�� Then
                MsgBox "�ò�����ʱֹͣҽ��֧�����뵽ҽ�����Ĵ���", vbInformation, gstrSysName
                Exit Function
            End If
            
            If bytType = 1 Then
                '�������ƵĲ��˽�������
                If lng�Ҷ� = deg����֧�� Or lng�Ҷ� = deg�ϴ���ϸ Then
                    MsgBox "�ò��˲���ʹ��ͳ�����֧��סԺ���á�", vbExclamation, gstrSysName
                End If
            End If
            
            If bytType = 1 Then
                Dim rsSelected As New ADODB.Recordset
                gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                        " From ���ղ��� A where 1=2 And A.����=[1]"
                Set rsSelected = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ѡ��Ĳ���", TYPE_�Թ���)
                
                'סԺҪѡ���֣���ȷ��һЩ�����շ���Ŀ
                gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                        " From ���ղ��� A where A.����=[1]"
                Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", TYPE_�Թ���)
                If rs����.RecordCount > 0 Then
VirusSelect:
                    If frm�ಡ��ѡ��.ShowSelect(rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�", rsSelected, False) = True Then
                        lng���� = 0
                        str���� = ""
                        With rs����
                            If .RecordCount <> 0 Then .MoveFirst
                            lng���� = rs����("ID")
                            Do While Not .EOF
                                str���� = str���� & "|" & rs����!ID
                                .MoveNext
                            Loop
                            If str���� <> "" Then str���� = Mid(str����, 2)
                        End With
                    Else
                        MsgBox "����Ҫѡ���֣�", vbInformation, gstrSysName
                        GoTo VirusSelect
                    End If
                End If
            End If
            
            '�������˵�����Ϣ�������ʽ��
            '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
            '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
            '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
            strIdentify = TrimStr(.Cardno)                              '0����
            strIdentify = strIdentify & ";" & TrimStr(.MediAccountNo)   '1ҽ����
            strIdentify = strIdentify & ";" & .Password        '2����
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
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_�Թ���, .CenterCode)
            
            If rsTemp.RecordCount = 0 Then
                ��ݱ�ʶ_���� = ""
                MsgBox "�ò�������������δ����������ʹ�á�", vbInformation, gstrSysName
                Exit Function
            Else
                lng���� = rsTemp("���")
            End If
            
            '�õ�ԭסԺ����
            If bytType <> 1 Then
                gstrSQL = "Select Nvl(����ID,0) ����ID From �����ʻ� Where ����=[1] And ҽ����=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�õ�ԭסԺ����", TYPE_�Թ���, TrimStr(.MediAccountNo))
                If Not rsTemp.EOF Then
                    lng���� = rsTemp!����ID
                End If
            End If

            strAddition = ";" & lng����                                 '8.���Ĵ���
            strAddition = strAddition & ";"                             '9.˳���
            strAddition = strAddition & ";" & TrimStr(.ClassCode)       '10��Ա���
            strAddition = strAddition & ";" & (.InPerAcc - .OutPerAcc)  '11�ʻ����
            strAddition = strAddition & ";" & .InpatientFlag            '12��ǰ״̬
            strAddition = strAddition & ";" & IIf(lng���� > 0, lng����, "") '13����ID

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
            strAddition = strAddition & ";" & lng�Ҷ�                   '17�Ҷȼ�
            strAddition = strAddition & ";" & .InPerAcc                 '18�ʻ������ۼ�
            strAddition = strAddition & ";" & .OutPerAcc                '19�ʻ�֧���ۼ�
            strAddition = strAddition & ";" & .OutAnnOverLine           '20����ͳ���ۼ�
            strAddition = strAddition & ";" & .OutAnnPlan               '21ͳ�ﱨ���ۼ�
            strAddition = strAddition & ";" & .AnnInpatientTimes        '22סԺ�����ۼ�
            strAddition = strAddition & ";"                             '23�������� (1����������)
            
            lng����ID = BuildPatiInfo(bytType, strIdentify & strAddition, lng����ID, TYPE_�Թ���)
            '���ظ�ʽ:�м���벡��ID
            ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & strAddition
            
            If bytType = 1 Then
                gstrSQL = "zl_������Ϣ_INSERT(" & TYPE_�Թ��� & "," & lng����ID & ",'" & str���� & "')"
                gcn����.Execute gstrSQL, , adCmdStoredProc
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
    Dim lngReturn As Long
    
    On Error GoTo errHandle
    
    'ִ��װǮ������˳��Ͷ�ȡ�����µĸ�������
    If װǮ����(lng����ID) = True Then
        '��������
        If ҽ���Ҷ�(gIC����.CenterCode, gIC����.Cardno) > deg�ϴ���ϸ Then
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

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, ByVal curȫ�Է� As Currency, ByVal cur�����Ը� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    Dim rsTemp As New ADODB.Recordset
    'Dim ic������� As TChargeResult       '�����շѶ����ṹ
    Dim ic������� As TIC����            '�����нṹ�����󷵻�ֵ�����⣨��Ҫ���漰���ļ�����Ա��
    Dim ic����д�� As TChargeParameter    '�����շ�д��ṹ
    Dim card�Ҷ� As cardҽ���Ҷ�
    Dim strҽԺ���� As String
    Dim lng���� As Long, lngReturn As Long, lng����ID As Long
    Dim curƱ���ܽ�� As Currency
    Dim dat��ǰ���� As Date
    Dim bln���� As Boolean, strҽ���� As String
    
    On Error GoTo errHandle
        
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    gstrSQL = "Select ����ID,���ʽ��  From ������ü�¼ Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    lng����ID = rsTemp("����ID")
    Do Until rsTemp.EOF
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    bln���� = Is���ݲ���(lng����ID, strҽ����)
    
    If bln���� = False Then
        If ReadICCard(ic�������) <> 0 Then
            Err.Raise 9000, gstrSysName, "�շ�ʱ����ʧ�ܡ�"
            Exit Function
        End If
        If ic�������.InpatientFlag = "1" Then
            Err.Raise 9000, gstrSysName, "�ò�����Ȼ��Ժ�����ܼ����� "
            Exit Function
        End If
    Else
        If Get���ݲ���_����(strҽ����, ic�������) = False Then
            Exit Function
        End If
    End If
    
    card�Ҷ� = ҽ���Ҷ�(ic�������.CenterCode, ic�������.Cardno)
    
    If card�Ҷ� = degֹ֧ͣ�� Then
        '�����ٴ����������
        �������_���� = True
        Exit Function
    End If
    
    dat��ǰ���� = zlDatabase.Currentdate
    
    '�жϸò��˵Ŀ��Ƿ������ȷ
    If ���IC��(lng����ID, TrimStr(ic�������.Cardno), TrimStr(ic�������.CenterCode)) = False Then Exit Function
    
    With ic�������
        'Ϊ�˱�֤��ȫ���ۼ����ݻ��Ƕ���������

        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�Թ��� & "," & Format(dat��ǰ����, "yyyy") & "," & _
            .InPerAcc & "," & .OutPerAcc + cur�����ʻ� & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Թ��� & "," & lng����ID & "," & _
            Format(dat��ǰ����, "yyyy") & "," & .InPerAcc & "," & .OutPerAcc & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ",0,0,0," & _
            curƱ���ܽ�� & "," & curȫ�Է� & "," & cur�����Ը� & "," & curƱ���ܽ�� - curȫ�Է� - cur�����Ը� & ",0,0,0," & _
            cur�����ʻ� & ",'" & .OutSerialNo + 1 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    If bln���� = False Then
        With ic����д��
            .Cardno = ic�������.Cardno         ' ����
            .OutPerAcc = ic�������.OutPerAcc + cur�����ʻ�  ' �����ʻ��ۼ�֧�����
            .OutExAcc = ic�������.OutExAcc                  ' �����ʻ��ۼ�֧�����
            .OutSubAcc = ic�������.OutSubAcc                ' �����ʻ��ۼ�֧�����
            .OutSerialNo = ic�������.OutSerialNo + 1  ' ֧��˳���
            .PayOccurDate = Format(dat��ǰ����, "yyyyMMdd")  ' ����
            .PayHospitalCode = Trim(Mid(gstrҽԺ����, 1, 4)) ' ҽԺ����
            .PayAccPay = cur�����ʻ�      ' �����ʻ�֧��
            .PayAmount = curƱ���ܽ��    ' �ܶ�
        End With
        
        lngReturn = ChargeWrite(ic����д��)
        If lngReturn Then
            Err.Raise 9000, gstrSysName, "����д�뿨ʧ�ܡ�" & ������Ϣ_����(lngReturn)
            Exit Function
        End If
    End If
        
    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    Dim rsTemp As New ADODB.Recordset
    'Dim ic������� As TChargeResult       '�����շѶ����ṹ
    Dim ic������� As TIC����            '�����нṹ�����󷵻�ֵ�����⣨��Ҫ���漰���ļ�����Ա��
    Dim ic����д�� As TChargeParameter    '�����շ�д��ṹ
    Dim card�Ҷ� As cardҽ���Ҷ�
    Dim lngReturn As Long, lng��� As Long, lng����ID As Long
    Dim curƱ���ܽ�� As Currency, curȫ�Է� As Currency, cur�����Ը� As Currency, cur����ͳ�� As Currency
    Dim dat��ǰ���� As Date
    Dim bln���� As Boolean, strҽ���� As String
    Dim lngԭҽ���� As Long
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ����ID,�������ý��,ȫ�Ը����,�����Ը����,����ͳ����,���  From ���ս����¼ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
        
    lng����ID = rsTemp("����ID")
    lngԭҽ���� = rsTemp("���")
    curƱ���ܽ�� = IIf(IsNull(rsTemp("�������ý��")), 0, rsTemp("�������ý��"))
    curȫ�Է� = IIf(IsNull(rsTemp("ȫ�Ը����")), 0, rsTemp("ȫ�Ը����")) * -1
    cur�����Ը� = IIf(IsNull(rsTemp("�����Ը����")), 0, rsTemp("�����Ը����")) * -1
    cur����ͳ�� = IIf(IsNull(rsTemp("����ͳ����")), 0, rsTemp("����ͳ����")) * -1
    
    bln���� = Is���ݲ���(lng����ID, strҽ����)
    
    If bln���� = False Then
        If ReadICCard(ic�������) <> 0 Then
            Err.Raise 9000, gstrSysName, "�˷�ʱ����ʧ�ܡ�"
            Exit Function
        End If
        If ic�������.InpatientFlag = "1" Then
            Err.Raise 9000, gstrSysName, "�ò�����Ȼ��Ժ�����ܼ�����"
            Exit Function
        End If
    Else
        If Get���ݲ���_����(strҽ����, ic�������) = False Then
            Exit Function
        End If
    End If
    
    If Not Check��Ч��(ic�������.CenterCode) Then Exit Function
    
    card�Ҷ� = ҽ���Ҷ�(ic�������.CenterCode, ic�������.Cardno)
    If card�Ҷ� = degֹ֧ͣ�� Then
        '�����ٴ����������
        '����������_���� = True
        Exit Function
    End If
    
    gstrSQL = "select B.����,B.��� " & _
            " from �����ʻ� A,��������Ŀ¼ B " & _
            " where A.����ID=[1] and A.����=[2]" & _
            "  and A.����=B.���� and A.����=B.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID, TYPE_�Թ���)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "��ϵͳ����Ա���ҽ�����ĵ����á�"
        Exit Function
    End If
    'ȡ��ǰҽ����
    g��������.��� = Val(Get���ղ���_����(rsTemp("����"), "ҽ����", True))
    If g��������.��� = 0 Then
        Err.Raise 9000, gstrSysName, "��ϵͳ����Ա���ҽ�����ݵ����ء�"
        Exit Function
    End If
    'ֻ�ܳ�����ҽ����ȵ��շѼ�¼
    If lngԭҽ���� < g��������.��� Then
       Err.Raise 9000, gstrSysName, "���ܳ����Ǳ�ҽ����ȵ������շѼ�¼��"
       Exit Function
    End If
    dat��ǰ���� = zlDatabase.Currentdate
        
    '�жϸò��˵Ŀ��Ƿ������ȷ
    If ���IC��(lng����ID, TrimStr(ic�������.Cardno), TrimStr(ic�������.CenterCode)) = False Then Exit Function
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    lng��� = rsTemp("����ID")
    
    With ic�������
        'Ϊ�˱�֤��ȫ���ۼ����ݻ��Ƕ���������

        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�Թ��� & "," & g��������.��� & "," & _
            .InPerAcc & "," & .OutPerAcc - cur�����ʻ� & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        gstrSQL = "zl_���ս����¼_insert(1," & lng��� & "," & TYPE_�Թ��� & "," & lng����ID & "," & _
            g��������.��� & "," & .InPerAcc & "," & .OutPerAcc - cur�����ʻ� & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ",0,0,0," & _
            curƱ���ܽ�� * -1 & "," & curȫ�Է� & "," & cur�����Ը� & "," & cur����ͳ�� & ",0,0,0," & cur�����ʻ� * -1 & ",'" & .OutSerialNo + 1 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    If bln���� = False Then
        With ic����д��
            .Cardno = ic�������.Cardno         ' ����
            .OutPerAcc = ic�������.OutPerAcc - cur�����ʻ� ' �����ʻ��ۼ�֧�����
            .OutExAcc = ic�������.OutExAcc                  ' �����ʻ��ۼ�֧�����
            .OutSubAcc = ic�������.OutSubAcc                ' �����ʻ��ۼ�֧�����
            .OutSerialNo = ic�������.OutSerialNo + 1  ' ֧��˳���
            .PayOccurDate = Format(dat��ǰ����, "yyyyMMdd")  ' ����
            .PayHospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
            .PayAccPay = cur�����ʻ�      ' �����ʻ�֧��
            .PayAmount = curƱ���ܽ��    ' �ܶ�
        End With
        
        lngReturn = ChargeWrite(ic����д��)
        If lngReturn Then
            Err.Raise 9000, gstrSysName, "�˷�ʱд�뿨ʧ�ܡ�" & ������Ϣ_����(lngReturn)
            Exit Function
        End If
    End If
    
    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
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
    Dim ic���� As TIC����
    Dim ic��Ժ���� As TInpatientRegResult       '��Ժ�ǼǶ����ṹ
    Dim ic��Ժд�� As TInpatientRegParameter    '��Ժ�Ǽ�д��ṹ
    Dim lngReturn As Long
    Dim dat��ǰ���� As Date, card�Ҷ� As cardҽ���Ҷ�
    Dim bln���� As Boolean
    
    On Error GoTo errHandle
    
    bln���� = Is���ݲ���(lng����ID, strҽ����)
    
    If bln���� = False Then
        If InpatientRegRead(ic��Ժ����) <> 0 Then
            MsgBox "��Ժ�Ǽ�ʱ����ʧ�ܡ�", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Get���ݲ���_����(strҽ����, ic����) = False Then
            Exit Function
        End If
        '�����ݴ��ݹ�����ֻ����Ҫ�õ��ļ����ֶ�
        With ic��Ժ����
            .Cardno = ic����.Cardno
            .CenterCode = ic����.CenterCode
        End With
    End If
        
        
    dat��ǰ���� = zlDatabase.Currentdate
    
    '���ˢ�����ĵĿ��Ƿ�ǰ���˵�
    If ���IC��(lng����ID, TrimStr(ic��Ժ����.Cardno), TrimStr(ic��Ժ����.CenterCode)) = False Then Exit Function

    card�Ҷ� = ҽ���Ҷ�(ic��Ժ����.CenterCode, ic��Ժ����.Cardno)
    
    If card�Ҷ� = degֹ֧ͣ�� Then
        '�����ٴ����������
        ��Ժ�Ǽ�_���� = False
        MsgBox "�ò����Ѿ�ֹͣҽ��֧����������Ϊҽ��������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If

    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Թ��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    If bln���� = False Then
        With ic��Ժд��
            .Cardno = ic��Ժ����.Cardno         ' ����
            .InpatientFlag = 1
        End With
        
        lngReturn = InpatientRegWrite(ic��Ժд��)
        If lngReturn Then
            MsgBox "��Ժ�Ǽ�д�뿨ʧ�ܡ�" & ������Ϣ_����(lngReturn), vbInformation, gstrSysName
            Exit Function
        End If
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
    Dim ic���� As TIC����
    Dim ic��Ժ���� As TInpatientRegResult       '��Ժ�ǼǶ����ṹ
    Dim ic��Ժд�� As TInpatientRegParameter    '��Ժ�Ǽ�д��ṹ
    Dim lngReturn As Long
    Dim bln���� As Boolean, strҽ���� As String
    
    On Error GoTo errHandle
    
    bln���� = Is���ݲ���(lng����ID, strҽ����)
    
    If bln���� = False Then
        If InpatientRegRead(ic��Ժ����) <> 0 Then
            MsgBox "��Ժ����ʱ����ʧ�ܡ�", vbInformation, gstrSysName
            Exit Function
        End If
        '���ˢ�����Ŀ��Ƿ�ǰ���˵�
        If ���IC��(lng����ID, TrimStr(ic��Ժ����.Cardno), TrimStr(ic��Ժ����.CenterCode)) = False Then Exit Function
    Else
        If Get���ݲ���_����(strҽ����, ic����) = False Then
            Exit Function
        End If
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Թ��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    If bln���� = False Then
        With ic��Ժд��
            .Cardno = ic��Ժ����.Cardno         '����
            .InpatientFlag = 0                  '��ʾ��Ժ
        End With
        
        lngReturn = InpatientRegWrite(ic��Ժд��)
        If lngReturn <> 0 Then
            MsgBox "��Ժ����д�뿨ʧ�ܡ�" & ������Ϣ_����(lngReturn), vbInformation, gstrSysName
            Exit Function
        End If
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
    Dim lngReturn As Long, ic���� As TIC����
    Dim bln���� As Boolean, strҽ���� As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If ���ҽ��������_���� = False Then
        '�������ӵ�ǰ�÷�����������Ϊ����ʹ��
        Exit Function
    End If
    
    bln���� = Is���ݲ���(rsExse("����ID"), strҽ����)
    
    If bln���� = False Then
        lngReturn = ReadICCard(ic����)
        If lngReturn <> 0 Then
            MsgBox "������Ϣʧ�ܡ�" & ������Ϣ_����(lngReturn), vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Get���ݲ���_����(strҽ����, ic����) = False Then Exit Function
    End If
    
    '���һЩ���ݵĳ�ʼ������������ԱҲҪʹ�õ�����
    With g��������
        .����ID = rsExse("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", CLng(rsExse("����ID")))
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
    
        '�����ڳ�Ժ���ʺ��ٴν��н���
        gstrSQL = "SELECT ����ID FROM ���ս����¼ WHERE ��;����=0 AND ����ID=[1] AND ��ҳID=[2] AND ����=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", .����ID, .��ҳID, TYPE_�Թ���)
        
        If rsTemp.RecordCount > 0 Then
            MsgBox "�����Ѿ����й�סԺ���㣬�����ٽ��н��ʲ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    
    'Ŀǰֻ���Թ�ҽ��ʹ�øò���
    '���ʹ�ñ��ղ����ж���ģ����ֻҪû�����أ�ҽԺ�ͻ�����ǰ������ϴ���
    gstrSQL = "select A.����ID,B.����,B.��� " & _
            " from �����ʻ� A,��������Ŀ¼ B " & _
            " where A.����ID=[1] and A.����=[2]" & _
            "  and A.����=B.���� and A.����=B.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, TYPE_�Թ���)
    If rsTemp.EOF = True Then
        MsgBox "��ϵͳ����Ա���ҽ�����ĵ����á�", vbInformation, gstrSysName
        Exit Function
    End If
    If Nvl(rsTemp!����ID) = 0 Then
        MsgBox "û��ѡ���֣���������ʣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    g��������.��� = Val(Get���ղ���_����(rsTemp("����"), "ҽ����", True))
    If g��������.��� = 0 Then
        MsgBox "��ϵͳ����Ա���ҽ�����ݵ����ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '1.2 �������˵���Ժʱ��
    gstrSQL = "select ��Ժ����,nvl(��Ժ����,to_date('3000-01-01','yyyy-MM-dd')) as ��Ժ���� " & _
              "from ������ҳ where ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, g��������.��ҳID)
    If rsTemp("��Ժ����") = CDate("3000-01-01") Then
        g��������.��;���� = 1
    Else
        '��ʾ�ò����Ѿ���Ժ
        g��������.��;���� = 0
    End If

    '�˴�ʹ��װǮ��������ҪĿ���ǳ�ʼ�����˵Ŀ��ϵ����Լ��ۼƽ���ͳ���ͳ���ۼƱ���
    If װǮ����(rsExse("����ID")) = False Then
        MsgBox "����װǮ����ʧ�ܣ��޷�׼ȷ�õ����˵�������ۼƱ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    With gIC����
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsExse("����ID") & "," & TYPE_�Թ��� & "," & .MediYear & "," & _
            .InPerAcc & "," & .OutPerAcc & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
    '������㷨������ҽ����ͬ
    סԺ�������_���� = סԺ�������(rsExse, ҽ���Ҷ�(ic����.CenterCode, ic����.Cardno))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function סԺ�������(rs������ϸ As Recordset, ByVal deg�Ҷ� As cardҽ���Ҷ�) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����Ҫ��NO����š�����ID��ҽ����Ŀ���롢�շ�����շ����ơ��������š���񡢲��ء��������۸񡢽�ҽ��,�Ǽ�ʱ��(����ʱ��),Ӥ����,���մ���ID
    Dim rs������� As Recordset     '��ҽ��֧��������ܵõ�
    Dim rs��׼��Ŀ���� As New ADODB.Recordset
    Dim rs�㷨 As New ADODB.Recordset, rs���� As New ADODB.Recordset         '����
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng����ID As Long
    Dim lng���� As Long, str���� As String
    Dim lng��ְ As Long, lng����� As Long, lng���� As Long
    Dim dblTemp As Double, lng���� As Long
    
    Dim dbl�����  As Double ''��һ����סԺ�ռ������Ŀ������ܵõ��Ľ��
    Dim dbl�ѱ������ As Double, dbl�ۼƽ��� As Double
    Dim dbl���� As Double, dbl���� As Double, dbl�ֶν��� As Double, dbl�ֶα��� As Double
    
    Dim clsҽ�� As New clsInsure
    Dim bln�����ʻ�֧��ȫ�Է� As Boolean, bln�����ʻ�֧�������Ը� As Boolean, bln�����ʻ�֧������ As Boolean
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency
    Dim blnȫ��ͳ�� As Boolean, bln������ As Boolean, bln�޷ⶥ�� As Boolean
    
    Dim bln������� As Boolean   '�����Թ�ҽ��������ǿ�����㣬��ʹ�ò����ǵڶ��ν��ʡ����ֶμ���Ҳ�Ǵ�ͷ��ʼ
    Dim dbl������ߺ� As Double, dbl��ν���ͳ��� As Double   '�����ָ�ò�����ǰ���ʵ��ۼ�
    Dim dbl�������� As Double, dbl�������� As Double
    Dim lngԭҽ���� As Long
    
    On Error GoTo errHandle
    '������������������������������������������������������������������������������������
    '1����ʼ��һЩ����
    Set gcol������� = New Collection
    
'    gstrSQL = "select D.ID ����ID,A.�շ�ϸĿID " & _
'             " from  " & _
'             " (select C.�շ�ϸĿID " & _
'             " from ���ղ��� A,ZLYB.������Ϣ B,������׼��Ŀ C " & _
'             " Where A.����=" & TYPE_�Թ��� & " And A.����=B.���� And B.����ID=" & g��������.����ID & " And Nvl(C.����,0)=0 And Nvl(C.����,0)<>2 And C.����ID=B.����ID And B.����ID=A.ID) A, " & _
'             " ������Ŀ B,����֧����Ŀ  C,����֧������ D " & _
'             " Where A.�շ�ϸĿID=C.�շ�ϸĿID And C.����=B.���� And B.����=C.��Ŀ���� And B.�������=D.���� And B.����=" & TYPE_�Թ��� & _
'             " And D.����=B.����"
'    Call OpenRecordset(rs��׼��Ŀ����, "��ȡ�ò������в��ֵ���׼��Ŀ����")
    
    bln�����ʻ�֧��ȫ�Է� = clsҽ��.GetCapability(support�����ʻ�ȫ�Է�, 0, TYPE_�Թ���)
    bln�����ʻ�֧�������Ը� = clsҽ��.GetCapability(support�����ʻ������Ը�, 0, TYPE_�Թ���)
    bln�����ʻ�֧������ = clsҽ��.GetCapability(support�����ʻ�����, 0, TYPE_�Թ���)
    
    gstrSQL = "select B.����,B.��� " & _
            " from �����ʻ� A,��������Ŀ¼ B " & _
            " where A.����ID=[1] and A.����=[2]" & _
            "  and A.����=B.���� and A.����=B.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, TYPE_�Թ���)
    If rsTemp.EOF = True Then
        MsgBox "��ϵͳ����Ա���ҽ�����ĵ����á�", vbInformation, gstrSysName
        Exit Function
    End If
    lng���� = rsTemp("���")
    str���� = rsTemp("����")
    
    gstrSQL = "select max(���) as ��� from ���ս����¼ where ����id=[1] and ��ҳid=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, g��������.��ҳID)
    If rsTemp.EOF = False Then
       lngԭҽ���� = IIf(IsNull(rsTemp("���")) = True, g��������.���, rsTemp("���"))
    Else
       lngԭҽ���� = g��������.���
    End If
    
    'If g��������.��� > Val(Format(rs������ϸ("����ʱ��"), "yyyy")) Then
    If g��������.��� > lngԭҽ���� Then
        bln������� = True
    End If
        
    '1.3 ��������סԺ�ڼ��ۼƽ������
    gstrSQL = "select nvl(sum(A.����),0) as ����,nvl(sum(A.����ͳ����),0) as ����ͳ���� " & _
              "  from ���ս����¼ A,���˽��ʼ�¼ B " & _
              "  Where A.����ID = [1] And A.��ҳID = [2]" & _
              " And A.���� = [3] And A.��¼ID = B.ID "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, g��������.��ҳID, TYPE_�Թ���)
    dbl������ߺ� = rsTemp("����")
    dbl��ν���ͳ��� = rsTemp("����ͳ����")
    
    With g��������
        gstrSQL = "select A.����,A.��Ա���,A.��ְ,A.�����," & _
                  "      B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ�" & _
                  " from �����ʻ� A,�ʻ������Ϣ B" & _
                  " where A.����ID=B.����ID(+) and A.����=B.����(+) " & _
                  "     and B.���(+)=[1] and A.����ID=[2] and A.����=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", .���, .����ID, TYPE_�Թ���)
        
        lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        lng��ְ = IIf(IsNull(rsTemp("��ְ")), 1, rsTemp("��ְ"))
        lng���� = IIf(IsNull(rsTemp("�����")), 0, rsTemp("�����"))
        .סԺ���� = IIf(IsNull(rsTemp("סԺ�����ۼ�")), 0, rsTemp("סԺ�����ۼ�"))
        .�ʻ��ۼ����� = IIf(IsNull(rsTemp("�ʻ������ۼ�")), 0, rsTemp("�ʻ������ۼ�"))
        .�ʻ��ۼ�֧�� = IIf(IsNull(rsTemp("�ʻ�֧���ۼ�")), 0, rsTemp("�ʻ�֧���ۼ�"))
        .�ۼƽ���ͳ�� = IIf(IsNull(rsTemp("����ͳ���ۼ�")), 0, rsTemp("����ͳ���ۼ�"))
        .�ۼ�ͳ�ﱨ�� = IIf(IsNull(rsTemp("ͳ�ﱨ���ۼ�")), 0, rsTemp("ͳ�ﱨ���ۼ�"))
    
        
        gstrSQL = "select �����,nvl(ȫ��ͳ��,0) as ȫ��ͳ�� ,nvl(������,0) as ������ ,nvl(�޷ⶥ��,0) as �޷ⶥ�� " & _
                " from ���������" & _
                " where ����=" & TYPE_�Թ��� & " and nvl(����,0)=" & lng���� & _
                "       and ��ְ=" & lng��ְ & " and ����<=" & lng���� & " and (" & lng���� & "<=���� or ����=0)"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڡ��������������������������õ���", vbInformation, gstrSysName
            Exit Function
        End If
        lng����� = rsTemp("�����")
        blnȫ��ͳ�� = (rsTemp("ȫ��ͳ��") = 1)
        bln������ = (rsTemp("������") = 1)
        bln�޷ⶥ�� = (rsTemp("�޷ⶥ��") = 1)
    End With
    
    '������������������������������������������������������������������������������������
    '2����ͳ��֧����Ŀ�ϼƷ�����������
    '2.1����ʼ����¼��
    Set rs������� = New ADODB.Recordset
    With rs�������
        If .State = adStateOpen Then .Close
        .Fields.Append "���մ���ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 8, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ͳ����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With

    Do Until rs������ϸ.EOF
    'װ����д���¼��������������ʹ��
        If rs������ϸ("������Ŀ��") = 1 Then
'            rs��׼��Ŀ����.Filter = "�շ�ϸĿID=" & rs������ϸ!�շ�ϸĿID
'            If rs��׼��Ŀ����.EOF Then
'                lng����ID = rs������ϸ!���մ���ID
'            Else
'                lng����ID = rs��׼��Ŀ����!����ID
'            End If
'            '���·�����ϸ
'            gstrSQL = ""
'            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ������ID")
            
            lng����ID = rs������ϸ!���մ���id
            If rs�������.RecordCount = 0 Then
                rs�������.AddNew
                rs�������("���մ���ID") = lng����ID
                rs�������("����") = rs������ϸ("����")
                rs�������("���") = rs������ϸ("���")
            Else
                rs�������.MoveFirst
                rs�������.Find "���մ���ID=" & lng����ID
                If rs�������.EOF Then
                    rs�������.AddNew
                    rs�������("���մ���ID") = lng����ID
                    rs�������("����") = rs������ϸ("����")
                    rs�������("���") = rs������ϸ("���")
                Else
                    rs�������("����") = rs�������("����") + rs������ϸ("����")
                    rs�������("���") = rs�������("���") + rs������ϸ("���")
                End If
            End If
            rs�������.Update
        Else
            curȫ�Է� = curȫ�Է� + rs������ϸ("���")
        End If
            
        dblTemp = dblTemp + rs������ϸ("���")
        rs������ϸ.MoveNext
    Loop
    g��������.�������ý�� = dblTemp
    
    '2.2���������ͳ����
    gstrSQL = "select ID,����,�㷨,ͳ��ȶ�,��׼����,��׼����,�Ƿ�ҽ�� FROM ����֧������  where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_�Թ���)
    rs�㷨.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    
    dblTemp = 0
    If rs�������.RecordCount > 0 Then rs�������.MoveFirst
    Do Until rs�������.EOF
        
        rs����.Filter = "ID=" & rs�������("���մ���ID")
        If rs����.EOF = False Then
            rs�㷨.Filter = "����='" & rs����("����") & "'"
        Else
            rs�㷨.Filter = "����='90009'"
        End If
        If rs�㷨.RecordCount > 0 Then
            If rs�㷨("�Ƿ�ҽ��") = 1 Then
                '�㷨:1-�ܶ������Ŀ��2-סԺ�պ˶���Ŀ
                If rs�㷨("�㷨") = 1 Then
                    If rs�㷨("ͳ��ȶ�") = 0 Then
                        curȫ�Է� = curȫ�Է� + rs�������("���")
                    Else
                        dblTemp = dblTemp + rs�������("���") * rs�㷨("ͳ��ȶ�") / 100
                    End If
                Else
                    If Val(rs�������("����")) > Val(rs�㷨("��׼����")) Then
                        '���סԺ�ճ�����׼��������ô�������� ��׼����*��׼���� +  (����-��׼����)*ͳ��ȶ�
                        '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                        dbl����� = rs�㷨("��׼����") * rs�㷨("��׼����") + _
                            (rs�������("����") - IIf(rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0, 0, rs�㷨("��׼����"))) * rs�㷨("ͳ��ȶ�")
                    Else
                        '���סԺ�յ�����׼��������ô�������� ����*��׼���� ���� ����*ͳ��ȶ�
                        '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                        If rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0 Then
                            dbl����� = rs�������("����") * rs�㷨("ͳ��ȶ�")
                        Else
                            dbl����� = rs�������("����") * rs�㷨("��׼����")
                        End If
                    End If
                    
                    '�ܽ��������С����ȡȫ��������ֻ�����
                    dblTemp = dblTemp + IIf(rs�������("���") < dbl�����, rs�������("���"), dbl�����)
                    
                    If rs�������("���") > dbl����� Then
                        'ȫ������ȫ�Է�
                        curȫ�Է� = curȫ�Է� + rs�������("���") - dbl�����
                    End If
                End If
            Else
                curȫ�Է� = curȫ�Է� + rs�������("���")
            End If
        Else
            curȫ�Է� = curȫ�Է� + rs�������("���")
        End If
        rs�������.MoveNext
    Loop
    g��������.����ͳ���� = dblTemp
    g��������.ȫ�Էѽ�� = curȫ�Է�
    g��������.�����Ը���� = g��������.�������ý�� - curȫ�Է� - dblTemp
    
    '������������������������������������������������������������������������������������
    '3��������ߡ��ⶥ�ߡ�֧������������
    '3.1��������ߡ��ⶥ��
    With g��������
        
        gstrSQL = "select max(decode(A.����,'A',A.���,0)) as ������ ,max(decode(A.����,'1',A.���,0)) as ���� " & _
                  "         ,max(decode(A.����,'" & (.סԺ���� + 1) & "',A.���,0)) as ʵ������,min(A.���) as ������� " & _
                  "  from ����֧���޶� A " & _
                  "  where A.����=" & TYPE_�Թ��� & " and A.����=" & lng���� & " and A.���=" & .���
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
                
        If bln������ Then
            .ʵ������ = 0
            .���� = 0
        Else
            .���� = IIf(IsNull(rsTemp("ʵ������")), 0, rsTemp("ʵ������"))
            If .���� = 0 Then
                'һ�㶼���У����ʵ�ڳ�����סԺ��������ȡ���һ�Σ�Ҳ���ǽ����С��һ�Σ�
                .���� = IIf(IsNull(rsTemp("�������")), 0, rsTemp("�������"))
            End If
            If .���� = 0 Then
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
    
        '3.2��������ǰ�۳������߽��ó����ε�ʵ������
        If dbl������ߺ� > 0 Then
            '�����ò��˿϶��ж�ν���
            
            If dbl������ߺ� > dbl��ν���ͳ��� Then
                '�ò��˵ı��ν��㻹Ҫ�۳�һ�������߽��
                dbl�������� = dbl������ߺ� - dbl��ν���ͳ���
            Else
                '�����Ѿ�����
                dbl�������� = 0
            End If
            
            If .���� > dbl������ߺ� Then
                '���������ߣ�Ҫ����β�ֵ
                .���� = .���� - dbl������ߺ�
            Else
                '��ǰ�����߽���Ѿ�ȫ��棬���β����ٱ�����
                .���� = 0
            End If
            
            dbl�������� = dbl�������� + .����
        Else
            dbl�������� = .����
        End If
        dbl�������� = dbl��������
    End With
    
    '3.3��ȡ�÷��õ���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.����,B.����,B.����,A.���� " & _
              "  from ����֧������ A,���շ��õ� B " & _
              "  Where A.���� =" & TYPE_�Թ��� & " And A.���� =" & lng���� & " And A.��� =" & g��������.��� & " And A.��ְ =" & lng��ְ & " And A.����� =" & lng����� & _
              "       and A.����=B.���� and A.����=b.���� and A.����=B.���� " & _
              "  order by B.����"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        MsgBox "���ڡ���Ƚ�����������ñ���ȵ�ͳ��֧����������", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������������������������������������������������������������������������������������
    '4������ôν���ɱ����Ľ��
    dbl�ۼƽ��� = 0   '����ֶ��ۼƽ���ͳ��
    dbl�ѱ������ = g��������.�ۼ�ͳ�ﱨ��
    g��������.ͳ�ﱨ����� = 0
    
    If bln������� = True Then
        '�������Ͳ��ÿ�����ǰ�Ľ�����
        dbl��ν���ͳ��� = 0
    End If
    Do Until rsTemp.EOF
        dbl�ֶν��� = 0
        dbl�ֶα��� = 0
        
        If dbl�ѱ������ < g��������.�ⶥ�� Or g��������.�ⶥ�� = 0 Then    'δ�����ⶥ�߻��޷ⶥ��
            '�����Լ�������
            dbl���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
            dbl���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
            If dbl���� = 0 Then
                If g��������.���� > dbl���� Then
                    MsgBox "�ò��˵�ʵ�����߱ȵ�һ�����õ����޻��࣬���鱣�շ��õ���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If g��������.����ͳ���� + dbl��ν���ͳ��� > dbl���� And (dbl��ν���ͳ��� < dbl���� Or dbl���� = 0) Then
                '�ö���ǰ��δ������ȫ�����������Ҫ����۳��Ľ��
                dblTemp = 0
                If dbl��ν���ͳ��� > dbl���� Then
                    '��ǰ�Ѿ��������
                    dblTemp = dbl��ν���ͳ��� - dbl����
                End If
                
                '����Ҫ�۳�һ�������ߺ��ѽ���������޽����б仯
                If dbl���� + dblTemp + dbl�������� > dbl���� And dbl���� > 0 Then
                    dbl���� = dbl����
                    dbl�������� = dbl�������� - (dbl���� - dbl���� - dblTemp) '�����Ѿ����꣬�����¶ο�
                Else
                    dbl���� = dbl���� + dbl�������� + dblTemp
                    dbl�������� = 0
                End If
                
                If g��������.����ͳ���� + dbl��ν���ͳ��� <= dbl���� Or dbl���� = 0 Then
                    '��ʵ��ֵ����
                    dbl�ֶν��� = g��������.����ͳ���� + dbl��ν���ͳ��� - dbl����
                    
                    '������ڼ������ߡ�����ǰ�Ľ��ʽ����½���ͳ��Ľ����ܴﵽ���ޣ���ֻ��ȡ0
                    If dbl�ֶν��� < 0 Then dbl�ֶν��� = 0
                Else
                    'ȫ�����
                    dbl�ֶν��� = dbl���� - dbl����
                End If
                '����������öεı������
                dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
                dbl�ֶα��� = Val(Format(dbl�ֶν��� * rsTemp("����") / 100, "0.00"))
                
                If dbl�ѱ������ + dbl�ֶα��� > g��������.�ⶥ�� And g��������.�ⶥ�� <> 0 Then
                    '���������˷ⶥ�ߣ����Ҵ��ڷⶥ������
                    dbl�ֶα��� = g��������.�ⶥ�� - dbl�ѱ������
                    
                    '���ƽ���ͳ����
                    If rsTemp("����") <> 0 Then
                        dbl�ֶν��� = dbl�ֶα��� * 100 / rsTemp("����")
                    Else
                        dbl�ֶν��� = 0
                    End If
                End If
                
                '���и�ʽ��
                dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
                dbl�ֶα��� = Val(Format(dbl�ֶα���, "0.00"))
                
                dbl�ѱ������ = dbl�ѱ������ + dbl�ֶα���
                g��������.ͳ�ﱨ����� = g��������.ͳ�ﱨ����� + dbl�ֶα���
            End If
        End If
        
        '���Ρ�����ͳ���ͳ�ﱨ��������
        lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        dblTemp = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        dbl�ۼƽ��� = dbl�ֶν��� + dbl�ۼƽ���
            
        gcol�������.Add Array(lng����, dbl�ֶν���, dbl�ֶα���, dblTemp)
        rsTemp.MoveNext
    Loop
    
    g��������.ʵ������ = dbl�������� - dbl��������
    
    With g��������
        '���㳬���Ը�����
        .�����Ը���� = .����ͳ���� - dbl�������� - dbl�ۼƽ���
        If .�����Ը���� < 0 Then .�����Ը���� = 0                   '�������ͳ����������ߣ�Ϊ����
    End With
    
    If deg�Ҷ� < degҽ��֧�� Then
        '������ҽ������֧��
        g��������.ͳ�ﱨ����� = 0
        g��������.�����Ը���� = 0
        
        סԺ������� = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
    Else
        If blnȫ��ͳ�� = True Then
            סԺ������� = "ҽ������;" & g��������.ͳ�ﱨ����� + g��������.�����Ը���� & ";0"
        Else
            סԺ������� = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
        End If
    End If
    
    '����Ҫ���Ǹ����ʻ���֧����Χ
    With g��������
        dblTemp = 0   '��ʱ�����ʹ�õĸ����ʻ����
        
        If bln�����ʻ�֧��ȫ�Է� = True Then
            dblTemp = dblTemp + .ȫ�Էѽ��
        End If
        
        If bln�����ʻ�֧�������Ը� = True And blnȫ��ͳ�� = False Then
            dblTemp = dblTemp + .�����Ը����
        End If
        
        If bln�����ʻ�֧������ = True Then
            'ֻ��֧������ͳ���δ�����Ĳ���
            dblTemp = dblTemp + .����ͳ���� - .ͳ�ﱨ�����
        Else
            dblTemp = dblTemp + .����ͳ���� - .ͳ�ﱨ����� - .�����Ը����
        End If
        
        If deg�Ҷ� >= deg����֧�� Then
            If .�ʻ��ۼ����� - .�ʻ��ۼ�֧�� - dblTemp > 0 Then
               סԺ������� = סԺ������� & "|�����ʻ�;" & dblTemp & ";1"
            Else
               סԺ������� = סԺ������� & "|�����ʻ�;" & IIf(.�ʻ��ۼ����� - .�ʻ��ۼ�֧�� > 0, .�ʻ��ۼ����� - .�ʻ��ۼ�֧��, 0) & ";1"
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    סԺ������� = ""
End Function

Public Function סԺ����_����(lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID     ���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim icסԺ���� As TIC����               'סԺ��������ṹ
    Dim icסԺд�� As TIC����               'סԺ��������ṹ
    Dim card�Ҷ� As cardҽ���Ҷ�
    Dim lngReturn As Long
    Dim bln���� As Boolean, strҽ���� As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim cur�����ʻ� As Currency, var������� As Variant
    
    On Error GoTo errHandle
    
    bln���� = Is���ݲ���(g��������.����ID, strҽ����)
    If bln���� = False Then
        If ReadICCard(icסԺ����) <> 0 Then
            Err.Raise 9000, gstrSysName, "����ʱ����ʧ�ܡ�"
            Exit Function
        End If
        
        '�жϸò��˵Ŀ��Ƿ������ȷ
        If ���IC��(g��������.����ID, TrimStr(icסԺ����.Cardno), TrimStr(icסԺ����.CenterCode)) = False Then Exit Function
    Else
        If Get���ݲ���_����(strҽ����, icסԺ����) = False Then Exit Function
    End If
    
    If Not Check��Ч��(icסԺ����.CenterCode) Then Exit Function
    card�Ҷ� = ҽ���Ҷ�(icסԺ����.CenterCode, icסԺ����.Cardno)
    
'    If card�Ҷ� = degֹ֧ͣ�� Then
'        '�����ٴ����������
'        סԺ����_���� = True
'        Exit Function
'    End If
        
    '������ʻ�֧�����
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    If Not rsTemp.EOF Then cur�����ʻ� = rsTemp!���
    
    
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    If g��������.��;���� = 0 Then
        '��ʾ�ò����Ѿ���Ժ
        icסԺ����.AnnInpatientTimes = icסԺ����.AnnInpatientTimes + 1
    End If
        
    With g��������
        'Ϊ�˱�֤��ȫ���ۼ����ݻ��Ƕ���������

        gstrSQL = "zl_�ʻ������Ϣ_insert(" & .����ID & "," & TYPE_�Թ��� & "," & .��� & "," & _
            icסԺ����.InPerAcc & "," & icסԺ����.OutPerAcc + cur�����ʻ� & "," & icסԺ����.OutAnnOverLine + .����ͳ���� & "," & _
            icסԺ����.OutAnnPlan + .ͳ�ﱨ����� & "," & icסԺ����.AnnInpatientTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Թ��� & "," & .����ID & "," & _
            .��� & "," & icסԺ����.InPerAcc & "," & icסԺ����.OutPerAcc & "," & icסԺ����.OutAnnOverLine & "," & _
            icסԺ����.OutAnnPlan & "," & icסԺ����.AnnInpatientTimes & "," & .���� & "," & .�ⶥ�� & "," & .ʵ������ & "," & _
            .�������ý�� & "," & .ȫ�Էѽ�� & "," & .�����Ը���� & "," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",0," & _
            .�����Ը���� & "," & cur�����ʻ� & ",'" & icסԺ����.OutSerialNo + 1 & "'," & .��ҳID & "," & .��;���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        For Each var������� In gcol�������
            '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
            gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
                var�������(0) & "," & var�������(1) & "," & var�������(2) & "," & var�������(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        Next
    End With
    
    If bln���� = False Then
        With icסԺ����
            .Cardno = icסԺ����.Cardno         ' ����
            .OutPerAcc = icסԺ����.OutPerAcc + cur�����ʻ�         ' �����ʻ��ۼ�֧�����
            .OutExAcc = icסԺ����.OutExAcc                         ' �����ʻ��ۼ�֧�����
            .OutSubAcc = icסԺ����.OutSubAcc                       ' �����ʻ��ۼ�֧�����
            .OutSerialNo = icסԺ����.OutSerialNo + 1               ' ֧��˳���
            .OutAnnOverLine = icסԺ����.OutAnnOverLine + g��������.����ͳ����  ' �������ͳ�����ۼ�
            .OutAnnPlan = icסԺ����.OutAnnPlan + g��������.ͳ�ﱨ�����          ' ����ͳ��֧������ۼ�
            .InpatientFlag = icסԺ����.InpatientFlag                             ' סԺ��־ 0-��סԺ 1-סԺ
            .AnnInpatientTimes = icסԺ����.AnnInpatientTimes                     ' ������ЧסԺ����
'            .PayOccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' ����
'            .PayHospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
'            .PayAccPay = cur�����ʻ�      ' �����ʻ�֧��
'            .PayAmount = g��������.�������ý��    ' �ܶ�
        End With
        
        lngReturn = WriteICCard(icסԺ����)
        If lngReturn Then
            Err.Raise 9000, gstrSysName, "����д�뿨ʧ�ܡ�" & ������Ϣ_����(lngReturn)
            Exit Function
        End If
        
        '��¼סԺ���
        Dim payLog As TPayLog
        With payLog
            .HospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
            .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' ����
            .AccPay = Space(8 - Len(CStr(cur�����ʻ� * 100))) & CStr(cur�����ʻ� * 100)
            .Amount = Space(8 - Len(CStr(g��������.�������ý�� * 100))) & CStr(g��������.�������ý�� * 100)
        End With
        ChargeLog payLog
    End If
        
    סԺ����_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
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

    Dim rsTemp As New ADODB.Recordset, rs������� As New ADODB.Recordset
    Dim icסԺ���� As TIC����                'סԺ��������ṹ
    Dim icסԺд�� As TIC����                'סԺ����д��ṹ
    Dim card�Ҷ� As cardҽ���Ҷ�
    Dim lng����ID As Long, lngReturn As Long
    Dim bln���� As Boolean, strҽ���� As String
    Dim cur�����ʻ� As Currency
    Dim lng����ID As Long, lngԭҽ���� As Long
    
    On Error GoTo errHandle
    
    gstrSQL = "select distinct A.ID,A.����id from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    lng����ID = rsTemp("ID") '�������ݵ�ID
    lng����ID = rsTemp("����id")
    gstrSQL = "select B.����,B.��� " & _
            " from �����ʻ� A,��������Ŀ¼ B " & _
            " where A.����ID=[1] and A.����=[2]" & _
            "  and A.����=B.���� and A.����=B.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID, TYPE_�Թ���)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "��ϵͳ����Ա���ҽ�����ĵ����á�"
        Exit Function
    End If
    'ȡ��ǰҽ����
    g��������.��� = Val(Get���ղ���_����(rsTemp("����"), "ҽ����", True))
    If g��������.��� = 0 Then
        Err.Raise 9000, gstrSysName, "��ϵͳ����Ա���ҽ�����ݵ����ء�"
        Exit Function
    End If
    
    'ֻ�������;���ʽ�������
    gstrSQL = "Select * " & _
              "  From ���ս����¼ Where ����=2 and ��;����=1 and ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "�ò��˵�ҽ�����㲻����;���ʣ��������ϡ�"
        Exit Function
    End If
    lngԭҽ���� = rsTemp("���")
    'ֻ�ܳ�����ҽ����ȵ�ҽ�������¼
    If lngԭҽ���� < g��������.��� Then
       Err.Raise 9000, gstrSysName, "���ܳ����Ǳ�ҽ����ȵ�ҽ�������¼��"
       Exit Function
    End If
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "Select * " & _
              "  From ���ս����¼ Where ����=2 and ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "�ò��˵�ҽ���������ݶ�ʧ���������ϡ�"
        Exit Function
    End If
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    bln���� = Is���ݲ���(rsTemp("����ID"), strҽ����)
    If bln���� = False Then
        If ReadICCard(icסԺ����) <> 0 Then
            Err.Raise 9000, gstrSysName, "����ʱ����ʧ�ܡ�"
            Exit Function
        End If
    Else
        If Get���ݲ���_����(strҽ����, icסԺ����) = False Then Exit Function
    End If
    
    If Not Check��Ч��(icסԺ����.CenterCode) Then Exit Function
    
    card�Ҷ� = ҽ���Ҷ�(icסԺ����.CenterCode, icסԺ����.Cardno)
    If card�Ҷ� = degֹ֧ͣ�� Then
        '�����ٴ����������
        סԺ�������_���� = False
        Err.Raise 9000, gstrSysName, "�ò����Ѿ�ֹͣҽ��֧�������ܽ��г���������"
        Exit Function
    End If
    
    
    '�жϸò��˵Ŀ��Ƿ������ȷ
    If ���IC��(rsTemp("����ID"), TrimStr(icסԺ����.Cardno), TrimStr(icסԺ����.CenterCode)) = False Then Exit Function
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & TYPE_�Թ��� & "," & rsTemp("���") & "," & _
        icסԺ����.InPerAcc & "," & icסԺ����.OutPerAcc - rsTemp("�����ʻ�֧��") & "," & icסԺ����.OutAnnOverLine - rsTemp("����ͳ����") & "," & _
        icסԺ����.OutAnnPlan - rsTemp("ͳ�ﱨ�����") & "," & icסԺ����.AnnInpatientTimes & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '�������ݻ������Ǹ���ԭ����
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Թ��� & "," & rsTemp("����ID") & "," & _
        rsTemp("���") & "," & rsTemp("�ʻ��ۼ�����") & "," & rsTemp("�ʻ��ۼ�֧��") & "," & rsTemp("�ۼƽ���ͳ��") & "," & _
        rsTemp("�ۼ�ͳ�ﱨ��") & "," & rsTemp("סԺ����") & "," & rsTemp("����") * -1 & "," & rsTemp("�ⶥ��") & "," & rsTemp("ʵ������") * -1 & "," & _
        rsTemp("�������ý��") * -1 & "," & rsTemp("ȫ�Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & rsTemp("����ͳ����") * -1 & "," & _
        rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") * -1 & "," & rsTemp("�����ʻ�֧��") * -1 & ",'" & icסԺ����.OutSerialNo + 1 & "'," & _
        IIf(IsNull(rsTemp("��ҳID")), "null", rsTemp("��ҳID")) & "," & rsTemp("��;����") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    cur�����ʻ� = rsTemp("�����ʻ�֧��")
    
    gstrSQL = "select ����,����ͳ����,ͳ�ﱨ�����,���� from ���ս������ where ����ID=[1]"
    Set rs������� = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    Do Until rs�������.EOF
        '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
        gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
            rs�������("����") & "," & rs�������("����ͳ����") * -1 & "," & rs�������("ͳ�ﱨ�����") * -1 & "," & rs�������("����") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        rs�������.MoveNext
    Loop
    
    If bln���� = False Then
        With icסԺ����
            .Cardno = icסԺ����.Cardno         ' ����
            .OutPerAcc = icסԺ����.OutPerAcc - rsTemp("�����ʻ�֧��") ' �����ʻ��ۼ�֧�����
            .OutExAcc = icסԺ����.OutExAcc                            ' �����ʻ��ۼ�֧�����
            .OutSubAcc = icסԺ����.OutSubAcc                          ' �����ʻ��ۼ�֧�����
            .OutSerialNo = icסԺ����.OutSerialNo + 1                  ' ֧��˳���
            .OutAnnOverLine = icסԺ����.OutAnnOverLine - rsTemp("����ͳ����")  ' �������ͳ�����ۼ�
            .OutAnnPlan = icסԺ����.OutAnnPlan - rsTemp("ͳ�ﱨ�����")          ' ����ͳ��֧������ۼ�
            .InpatientFlag = icסԺ����.InpatientFlag                  ' סԺ��־ 0-��סԺ 1-סԺ
            .AnnInpatientTimes = icסԺ����.AnnInpatientTimes          ' ������ЧסԺ����
'            .PayOccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")           ' ����
'            .PayHospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
'            .PayAccPay = cur�����ʻ� * -1    ' �����ʻ�֧��
'            .PayAmount = g��������.�������ý��    ' �ܶ�
        End With
        
        lngReturn = WriteICCard(icסԺ����)
        If lngReturn Then
            Err.Raise 9000, gstrSysName, "����д�뿨ʧ�ܡ�" & ������Ϣ_����(lngReturn)
            Exit Function
        End If
        '��¼סԺ���
        Dim payLog As TPayLog
        With payLog
            .HospitalCode = Mid(gstrҽԺ����, 1, 4) ' ҽԺ����
            .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' ����
            .AccPay = Space(8 - Len(CStr(cur�����ʻ� * -100))) & CStr(cur�����ʻ� * -100)
            .Amount = Space(8 - Len(CStr(g��������.�������ý�� * 100))) & CStr(g��������.�������ý�� * 100)
        End With
        ChargeLog payLog
    End If
        
    סԺ�������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
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
    Dim strҽ���� As String
    
    Dim strװǮģʽ As String, blnǿ��װǮ As Boolean
    Dim strҽ����  As String, lngװǮ�ڴ� As Long
    Dim dbl�ۼ�ע�� As Double
    Dim ic�� As TIC����
    Dim strҽ����_IC  As String, lngװǮ�ڴ�_IC As Long
    Dim dbl�ۼ�ע��_IC As Double
    Dim lngTemp As Long
    
    Dim str����ֵ As String
    
    On Error GoTo errHandle
    
    '�õ����µ�IC����Ϣ
    'ʹ�ñ��صģ���Ϊ���ܽ��и��ĵ��ֲ��ɹ�
    If Is���ݲ���(lng����ID, strҽ����) = False Then
        If ReadICCard(gIC����) <> 0 Then
            Exit Function
        End If
    Else
        'ҽ�����˲���ҪװǮ
        If Get���ݲ���_����(strҽ����, gIC����) = False Then Exit Function
        װǮ���� = True
        Exit Function
    End If
    '�жϿ��Ƿ�ǰ���˵�
    If lng����ID > 0 Then
        If ���IC��(lng����ID, TrimStr(gIC����.Cardno), TrimStr(gIC����.CenterCode)) = False Then
            Exit Function
        End If
    End If
    ic�� = gIC����
    
    With ic��
        strҽ����_IC = .MediYear
        lngװǮ�ڴ�_IC = .InNo
        dbl�ۼ�ע��_IC = .InPerAcc
    End With
    
    '���װǮģʽ
    '���кϷ�����֤
    strװǮģʽ = Left(Get���ղ���_����(ic��.CenterCode, "װǮģʽ", False), 1)
    strҽ���� = Get���ղ���_����(ic��.CenterCode, "ҽ����", True)
    lngװǮ�ڴ� = Val(Get���ղ���_����(ic��.CenterCode, "װǮ���", True))
    
    If strװǮģʽ = "" Or strҽ���� = "" Then
        MsgBox "���������Ա���ҽ�����ݵ����ء�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If strװǮģʽ = "1" Then
        '����װǮ
        'Modified By ���� 2003-12-10 ���������� ����ǰ��ģʽû�д����ԸĻ�����
        lngTemp = OnLineInMoney(ic��.CenterCode, ic��.Cardno, strҽ����_IC, Trim(gstrҽԺ����))
        If lngTemp <> 0 Then
            'װǮ���ɹ�
            '�ж��Ƿ��и���ҽ����
            If strҽ���� > strҽ����_IC Then
                MsgBox "װǮ�嵥��û�д˿�����Ϣ���뵽���Ĵ���", vbInformation, gstrSysName
                Exit Function
'                Call ����ҽ����װǮ(ic��, strҽ����, lngװǮ�ڴ�, ic��.InPerAcc - ic��.OutPerAcc)
'                '����Ϣд�ؿ���
'                If ��¼װǮ��־(ic��, strҽ����_IC, lngװǮ�ڴ�_IC, dbl�ۼ�ע��_IC) = True Then
'                    '����ȫ�ֱ�������������
'                    gIC���� = ic��
'                Else
'                    'װǮʧ��
'                    Exit Function
'                End If
            End If
        Else
            'װǮ�ɹ����ӿ��ж����µ�ֵ
            If ReadICCard(gIC����) <> 0 Then
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
        .OutAnnPlan = 0
        .OutAnnOverLine = 0
        .AnnInpatientTimes = 0
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
        Err.Clear
        gcn����.RollbackTrans
        Exit Function
    End If
    
    '���д������
    If WriteICCard(ic����) <> 0 Then
        gcn����.RollbackTrans
        MsgBox "IC��װǮ����ʧ�ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    gcn����.CommitTrans
    ��¼װǮ��־ = True
End Function

Private Function ҽ���Ҷ�(ByVal str���� As String, ByVal str���� As String) As cardҽ���Ҷ�
'����ָ���û���ҽ���Ҷȼ�
    Dim rsTemp As New ADODB.Recordset
    
    If ���ҽ��������_���� = False Then
        '�������ӵ�ǰ�÷�����������Ϊ����ʹ��
        ҽ���Ҷ� = degֹ֧ͣ��
        Exit Function
    End If
    
    gstrSQL = "select �Ҷ� from ������ where ���Ĵ���='" & str���� & "' and ����='" & str���� & "'"
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    
    If rsTemp.RecordCount > 0 Then
        '���ûҶ�ֵ
        ҽ���Ҷ� = Val(rsTemp("�Ҷ�"))
    Else
        '�����Ĳ��·�
        ҽ���Ҷ� = deg����֧��
    End If
    
End Function

Private Function ���IC��(ByVal lng����ID As Long, ByVal str���� As String, ByVal str���� As String) As Boolean
'���ܣ��жϸò��˵Ŀ��Ƿ������ȷ
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.����,A.ҽ����,B.���� from �����ʻ� A,��������Ŀ¼ B " & _
              " where A.����=[1] and A.����ID=[2] and a.����=B.���� and A.����=B.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_�Թ���, lng����ID)
    
    If rsTemp("����") <> str���� Or rsTemp("����") <> str���� Then
        MsgBox "ˢ�����еĿ����ǵ�ǰ���˵ģ��������ȷ��IC����", vbInformation, gstrSysName
        Exit Function
    End If
    
    ���IC�� = True
End Function

Private Function ���ҽ��������_����() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If gcn����.State = adStateOpen Then
        ���ҽ��������_���� = True
        Exit Function
    End If
    
    '��������ҽ��������������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_�Թ���)
    
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
        .InExAcc = 0        'As Double          ' �������ۼ�ע����
        .OutExAcc = 0       'As Double          ' �����ʻ��ۼ�֧�����
        .InSubAcc = 0       'As Double          ' �����ʻ��ۼ�ע����
        .OutSubAcc = 0      'As Double          ' �����ʻ��ۼ�֧�����
        .OutAnnPlan = 0     'As Double          ' ����ͳ��֧������ۼ�
        .OutAnnOverLine = 0 'As Double          ' �������ͳ�����ۼ�
        .Password = "9000"       'As String * 4      ' ��������
        .AnnInpatientTimes = 0 'As Long           ' ������ЧסԺ����
        .InpatientFlag = 0  'As String * 1      ' סԺ��־ 0-��סԺ 1-סԺ
        .HasSubInsurance = 0 'As String * 1      ' �Ƿ�μӹ���Ա��������  0-��  ����-��
        .HasExInsurance = 0 'As String * 1      ' �Ƿ�μӲ��䱣��0����1��
        .HasBigIllness = 0  'As String * 1      ' �Ƿ�μӴ�ҽ��
    End With
    
    Get���ݲ���_���� = True
End Function


Private Function Is���ݲ���(ByVal lng����ID As Long, strҽ���� As String) As Boolean
'���ܣ������ʻ���Ϣ�жϲ����Ƿ����ݲ���
'���������ز��˵�ҽ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select ��ְ,ҽ���� from �����ʻ� where ����=[1] and ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_�Թ���, lng����ID)
    
    If rsTemp.EOF = True Then
        '�ò���û����
        Is���ݲ��� = False
    Else
        Is���ݲ��� = IIf(rsTemp("��ְ") = 3, True, False)
        strҽ���� = rsTemp("ҽ����")
    End If
End Function

Public Function Get���ղ���_����(ByVal str���Ĵ��� As String, ByVal str������ As String, blnҽ�������� As Boolean) As String
'���ܣ���ñ��ղ���
    Dim rsTemp As New ADODB.Recordset
    
    If ���ҽ��������_���� = False Then
        Exit Function
    End If
    
    gstrSQL = "select A.������,A.����ֵ from ���ղ��� A " & _
              " where A.������='" & str������ & "' and A.����=" & TYPE_�Թ��� & " and (A.���� is null or A.���� in (select B.��� from ��������Ŀ¼ B where B.����=" & TYPE_�Թ��� & " and B.����='" & str���Ĵ��� & "'))"
    If blnҽ�������� = True Then
        Call OpenRecordset_OtherBase(rsTemp, "", gstrSQL, gcn����)
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��")
    End If
    
    If rsTemp.EOF = False Then
        Get���ղ���_���� = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
    End If
End Function

Public Function Check��Ч��(ByVal strCenterCode As String) As Boolean
    '������ĵ���Ч��
    Dim str��Ч��  As String
    
    str��Ч�� = Get���ղ���_����(strCenterCode, "��Ч��", True)
    
    If IsDate(str��Ч��) = False Then
        MsgBox "���ȴ�ҽ�������������ݺ���ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If CDate(str��Ч��) < CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")) Then
        MsgBox "��������ҽ�������Ѿ�������Ч�ڡ�", vbInformation, gstrSysName
        Exit Function
    End If
    Check��Ч�� = True
End Function
