Attribute VB_Name = "mdl����"
Option Explicit
'�㽭����ҽ���ӿڲ���˵��
'���ṩʵʱ�ϴ��ӿڣ�Ҫ��ҽԺÿ���°�ǰִ�в��˷���Ԥ������ɣ���Ϊ�Է��ӿڴ���ȱ��
'���Ҳ����ڶ���

'���������$$������$$
'�û������������ʱ��ͨ���ⲿ�������ݰ��ṩ������һ�㽻�״��뽻����ǰ4����̶�Ϊ���Ƿ����籣����0�ޣ�1�У�~ IC�����ݣ��п�ʱһ�������壬�޿�ʱ��ҽ���ţ�~�ֽ�֧����ʽ��1�ֽ�2����Ǯ����3���н�ǿ���~���п���Ϣ��ָ������û��ĳ������ʱ��ӦֵΪ�գ�û���籣��ʱ����Ҫ��д���Ľ��ײ���ִ�С�
'
'����������$$����״̬~������Ϣ~���׽����Ϣ$$
'���н���״̬��0�ɹ���>0�ɹ������о�����ʾ��<0ʧ�ܡ�
'����û����״���ɹ������ظ��û�ʱ״̬Ϊ"���׳ɹ�"��������Ϣһ��Ϊ�գ�������ʵ�ʸ�ʽΪ"$$0~~���׽����Ϣ$$"������о�����Ϣʱ���侯�������ڴ�����Ϣ�У�" �������ʽһ��Ϊ""$$x~������Ϣ~���׽����Ϣ$$""�����⣬���׽����Ϣ��ǰ3����̶���дҽ�������(дҽ���������0��ʾ��д��д���ɹ���������ʾд��������Ϣ)~�����п����(0��ʾ���ۻ�۳ɹ�)~д����IC�����ݡ�û����������ӦֵΪ�ա�
'�������ʧ�ܣ�һ��ֻ�г�����Ϣ��û�н��׽����Ϣ����ʽΪ"$$-1~������Ϣ~"����Ȼ��������ҪҲ���Լ��д�����Ϣ���н��׽����Ϣ��
'������Ϣ�ⲿ��ʽΪ��"��Ʒ��%%������.�����%%�����%%����ԭ��"������"����ԭ��"Ϊ���������������ṹ����ͼ��ʾ��
'
'��1��$$-1~3333%%f_UserBargaingApply.3%%-3%%�Ƿ�����~$$
'��2��$$0~~���׽����Ϣ$$
'һ�㽻��ʧ��ʱ����ҪҽԺϵͳ������ԭ����ʾ�������û���������Ҵ���ԭ�������Ҫ��������Ϣ�ṩ���ӿڿ����߲���ԭ������Ҫ�ṩ�����ط��ز�����

'----------�����ַ�˵��----------
'   �ַ�    �ַ�˵��           ����˵��
'   $$         ˫��Ԫ����      �ָ������ָ��������ݰ�
'   ~          ��������        �ָ����װ��в�ͬ��
'   %%         ˫�ٷֱȷ���    �ָ����װ���ͬ���Ԫ��
'   '          ������          ϵͳ�ַ����ָ���������ʹ��

'----------IC�����ݲ���----------
'3.4.    IC�����ݸ�ʽ
'����IC���ṹ���Է���ӿ�ʹ��Ϊǰ�ᣬ��������IC�������������۶�д����������ģ����ӿ�ģ�������ģ�鶼ʹ�ô˿����ݸ�ʽ�������淶Ϊ�����ڳ���10λ����ʽyyyy-mm-dd��������12λ����ʽ000000000.00�������Ȳ���ʱǰ�油0���������ͳ��Ȳ���ʱҲǰ�油0���ַ����ͳ��Ȳ���ʱ����ո��������ȡ��������ʽ�����ڱ�����ȫ�����ַ�����ʵ�ʿ��������֣���
'������ͳһ��һ���ļ����ļ������������ɸ��ֽڱ����ֶΣ�
'���    �ֽ�    ����Ԫ��    ��ʽ    ����
'���ֽڣ�
'1   1-14    ҽ��֤�ţ��ڲ����ţ�    �ַ�    14
'2   15-24   ����ҽ�Ʊ��ո����ʺ�    �ַ�    10
'3   25-28   ҽ����Ա��𣨲α�״̬��    ����    4
'4   29-32   ҽ�����������ߣ���𣬽�������  ����    4
'5   33  ���ⲡ��־  �ַ�    1
'6   34  ����Ա�α���־  �ַ�    1
'7   35-52   ������ݺ�  GB11643 18
'8   53-62   ����    �ַ�    10
'9   63-66   �Ա�    GB2261  4
'10  67-68   ����    GB3304  2
'11  69-74   ������  GB/T 2260   6
'12  75-84   ��������    ���ڡ�  10
'13  85-124  ��λ����    �ַ�    40
'14  125-134 ҽ����Ƚ�תʱ��    ���ڡ�  10
'15  135-146 �����ת���    ��  12
'16  147-156 ҽ����Ϣ����ʱ��    ���ڡ�  10
'17  157-168 ���ʵ��겦����� /����ʵ�ʲ���  ���    12
'18  169-170 �Ը������ݼ�����    ����    2
'19  171-180 �����ֶ�1   �ַ�    10
'20  181-194 ����״̬��ȫΪ0��ʾδ��������Ϊָ������д��ʧ�� ����    14
'21  195-206 ����ҽ�Ʊ��ո����˻����    ���    12
'22  207-218 ���ʵ���ʹ���ۼƽ��    ���    12
'23  219-230 ��������ʹ���ۼƽ��    ���    12
'24  231-242 ��ȸ����Ը��ۼƽ���ͨ���    ���    12
'25  243-254 ����ҽ�����ۼƽ��  ���    12
'26  255-266 ������ۼƽ��    ���    12
'27  267-278 ��������ͳ��֧���ۼƽ��    ���    12
'28  279-290 ����סԺͳ��֧���ۼƽ��    ���    12
'29  291-302 סԺ���������ۼƽ��  ���    12
'30  303-314 ���ߺ��Ը��ۼƽ��    ���    12
'31  315-326 �����������ⲡ�ۼƽ��  ���    12
'32  327-338 ���깫��Ա�ۼƽ��  ��  12
'33  339-341 ��������ͳ�������  ���֡�  3
'34  342-344 ����סԺͳ�������  ���֡�  3
'35  345-347 ����ҽ���ۼӴ�����  ���֡�  3
'36  348 ��ǰסԺ��־    �ַ�    1
'37  349-352 �������    ���֡�  4
'38  353-368 �����ֶ�2   �ַ�    16
'
'����:
'�����˻���� = ���ʵ��겦����� / ʵ�� - ���ʵ���ʹ���ۼƽ��
'�����˻���� = �����ת��� - ��������ʹ���ۼƽ��
'�����ֶ�Ϊ�����ַ����������ݿ��ܲ��Ǳ�׼����ASCII�룬���浽���ݿ�ʱ��Ҫע�⡣
'
'������ʽ:
'1����ȡ��һ���ļ�������ֻ��һ���ļ�����˶�ȡһ���ļ��Ͷ�ȡ�����ļ�������ͬ
'2����ȡ�ڶ����ļ���ҽԺ���޷�ʹ��
'10: ��ȡ�����ļ�
'
'
'IC����ʽ����������
'"1111111111111122222222220033001111555555555555555555������    ��  11�㽭  1977-05-25�����ߵ�λ                              2002-02-05000001000.002002-02-05000002000.0005          00000000000000000003000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.0000000000000000000000000.00000000000.00000000000.00000000000.00"


'����ֵ���壺0-����;>0-���ھ���;<0-ʧ�ܣ�������Ϣ��strReturnMsg��
Private Declare Function LHYB_Init Lib "BargaingApply" Alias "f_UserBargaingInit" _
    (ByVal StrInput As String, ByVal strReturnMsg As String, ByVal strOutput As String) As Integer
Private Declare Function LHYB_Close Lib "BargaingApply" Alias "f_UserBargaingClose" _
    (ByVal StrInput As String, ByVal strReturnMsg As String, ByVal strOutput As String) As Integer
Private Declare Function LHYB_Business Lib "BargaingApply" Alias "f_UserBargaingApply" _
    (ByVal intCode As Integer, ByVal dblSequence As Double, ByVal StrInput As String, _
    ByVal strReturnMsg As String, ByVal strOutput As String) As Integer

Type IC_Struct
    IC������                As String           '���没��IC������������
    ҽ��֤��                As String
    �ʺ�                    As String
    ��Ա���                As String
    �������                As String
    ���ⲡ                  As Byte
    ����Ա                  As Byte
    ��ݺ�                  As String
    ����                    As String
    �Ա�                    As String
    ����                    As String
    ������                  As String
    ��������                As String
    ��λ����                As String
    ��תʱ��                As String
    ��ת���                As Double
    ����ʱ��                As String
    ����ʵ�ʲ���            As Double
    �Ը�����                As Double
    ��״̬                  As String
    �����˻����            As Double
    ���ʵ���ʹ���ۼ�        As Double
    ��������ʹ���ۼ�        As Double
    �����Ը��ۼƽ��        As Double
    ���ۼƽ��              As Double
    ���ۼƽ��            As Double
    ����ͳ���ۼƽ��        As Double
    סԺͳ���ۼƽ��        As Double
    סԺ���������ۼƽ��  As Double
    ���ߺ��Ը��ۼƽ��    As Double
    �������ⲡ�ۼƽ��      As Double
    ����Ա�ۼƽ��          As Double
    ����ͳ�������          As Double
    סԺͳ�������          As Double
    ҽ���ۼӴ�����          As Double
    סԺ��־                As Byte
    �������                As Integer

'��������Ϊ��������
    mstrҽԺ���� As String
    mstrҽԺ�ȼ� As String
    mstrҵ������ As String
    mlng����ID As Long
    mstr������ˮ�� As String
    mstr���ﵥ�ݺ� As String
    mstr������ As String
    mdbl��ҽ����� As Double
    mstr������ڲ����� As String
End Type
Public IC_Data_���� As IC_Struct

Private mintFunc As Long   '���ܺ�
Private mstrFunc As String  '������
Private mstrInput As String '���
Private mstrOutput As String '�����
Private mstrMsg As String   '���ص���Ϣ

Public gcn���� As New ADODB.Connection

Private mblnInit As Boolean

Public Enum Function_����
    InitInsure = 0                '��
    EndInsure = 1               '�ر�
    ReadIC = 22             '����
    GetSequence = 23        '��ȡ������ˮ��
    PreRegist = 27          '�Һ�Ԥ����
    Regist = 28             '�ҺŽ���
    RegistDel = 31          '�Һ�����/��������
    PreClinic = 29          '����Ԥ����
    clinic = 30             '�������
    ClinicDel = 31          '�����������
    Comein = 32             '��Ժ�Ǽ�
    ComeIndel = 40          'ȡ����Ժ�Ǽ�
    ChargeDetail = 33       'סԺ����/ҽ��ִ��
    PreSettle = 34          'סԺԤ����
    Settle = 36             'סԺ����
    SettleDel = 37          '��֧����;���㣬��ˣ�ȡ����Ժ��ͬʱ�����Ͻ���
    ModifyPatient = 38      '�޸���Ժ������Ϣ
    ��Ժ���� = 35           '��д��Ժ����(�൱�ڳ�Ժ)
    ҽ��ת�ԷѲ��� = 39     'ҽ������תΪ�ԷѲ���
    ת��תԺ���� = 41
    ת��תԺ��ѯ = 42
    RequestBusiness = 43    '��ѯ���׽��
    Decide = 49             '����ȷ��
    ConfigPara = 52         '�ӿڲ�������
End Enum

Private Const strSplit As String = "%%"
Private Const strField As String = "~"

Public Function ReadIC_����(Optional ByVal strҽ���� As String = "") As Boolean
    Dim arrReturn
    Dim intReturn As Integer
    Dim strTest As String, strBit As String
    
    '���IC���Ķ�����������������д��IC_Data_����
    Call Interface_Prepare_����(ReadIC, IIf(strҽ���� <> "", "0", "1") & "~" & strҽ���� & "~~~10", "")
    intReturn = Interface_Exec_����()
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    '���ӿ�Ҫ����м��
    '���׽��~������Ϣ + ���¿���� + �� + ������IC������ + �����֤�����15λ��ÿλ����һ�����⣩+ ��
    '�����֤������壨λ������߿�ʼ���ж����ȼ���987126534����λ����Ϊ0��ʾ��Ӧ�����֤��������
    '��1λ: ��Ա���������� ��2λ: �������������
    '��3λ: �����˻������� ��4λ: �����ʻ�������
    '��5λ: סԺ���� ��7λ: �ڲ�����λ
    '��6λ����ҪȦ����ת(0�ɹ�����ҪȦ�棬1��ҪȦ�浫û��Ȧ�棬2ʧ��)
    '��7λ����������Ҫ����(0�ɹ�����Ҫ���£�1�������һ�������ݵ�û�и��£�2�����б�ҽԺȡ���Ľ������ݵ���û�и���3������ҽԺ��������ʧ�ܣ�4��������ҽԺ���������޷��������Ļ�����û�����ݸ���ʧ�ܣ�5����ԭ�������Ҫ���µ�����ʧ��)
    '��8λ���ϴν���д��ʧ�ܱ��������ɽ����κδ��������Ƚ�����0���� 1���ǣ�
    '��9λ���α���Ա�Ƿ���Ч��0�������α� 1��û�α���α�״̬��Ч��
    '����λ: �ڲ�����
    arrReturn = Split(mstrMsg, "~")
    strTest = arrReturn(5)
    If Mid(strTest, 9, 1) <> 0 Then
        MsgBox "�ò���û�вα���α�״̬��Ч��", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 8, 1) <> 0 Then
        MsgBox "�ϴν���д��ʧ�ܱ��������ɽ����κδ��������Ƚ�����", vbInformation, gstrSysName
        Exit Function
    End If
    strBit = Mid(strTest, 7, 1)
    If strBit <> 0 Then
        If strBit = 1 Then
            MsgBox "�������һ�������ݵ�û�и��£�", vbInformation, gstrSysName
            Exit Function
        ElseIf strBit = 2 Then
            MsgBox "�����б�ҽԺȡ���Ľ������ݵ���û�и��£�", vbInformation, gstrSysName
            Exit Function
        ElseIf strBit = 3 Then
            MsgBox "������ҽԺ��������ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        ElseIf strBit = 4 Then
            MsgBox "��������ҽԺ���������޷��������Ļ�����û�����ݸ���ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBox "����ԭ�������Ҫ���µ�����ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Mid(strTest, 1, 1) <> 0 Then
        MsgBox "��Ա���������ᣡ", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 2, 1) <> 0 Then
        MsgBox "������������ϣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 6, 1) <> 0 Then
        MsgBox "��ҪȦ����ת��", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 5, 1) <> 0 Then
        MsgBox "סԺ���ᣡ", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 3, 1) <> 0 Then
        MsgBox "�����˻������ᣡ", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 4, 1) <> 0 Then
        MsgBox "�����˻������ᣡ", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������д���ṹ����
    IC_Data_����.IC������ = arrReturn(4)
    If Not ExchangeICData(arrReturn(4), strҽ����) Then Exit Function
    
    ReadIC_���� = True
End Function

Private Function ExchangeICData(ByVal strBuffer As String, Optional ByVal strҽ���� As String) As Boolean
    On Error GoTo errHand
    With IC_Data_����
        If strҽ���� = "" Then
            .ҽ��֤�� = Getsubstr(strBuffer, 1, 1, 14)
        Else
            .ҽ��֤�� = strҽ����
        End If
        .�ʺ� = Getsubstr(strBuffer, 1, 171, 10)        'ȡ�ı����ֶ�1���˴��ͽӿ��ĵ�����
        .��Ա��� = Getsubstr(strBuffer, 1, 25, 4)
        .������� = Getsubstr(strBuffer, 1, 29, 4)
        .���ⲡ = Getsubstr(strBuffer, 1, 33, 1)
        .����Ա = Getsubstr(strBuffer, 1, 34, 1)
        .��ݺ� = Getsubstr(strBuffer, 1, 35, 18)
        .���� = Getsubstr(strBuffer, 1, 53, 10)
        .�Ա� = Getsubstr(strBuffer, 1, 63, 4)
        .���� = Getsubstr(strBuffer, 1, 67, 2)
        .������ = Getsubstr(strBuffer, 1, 69, 6)
        .�������� = Getsubstr(strBuffer, 1, 75, 10)
        .��λ���� = Getsubstr(strBuffer, 1, 85, 40)
        .��תʱ�� = Getsubstr(strBuffer, 1, 125, 10)
        .��ת��� = Val(Getsubstr(strBuffer, 1, 135, 12))
        .����ʱ�� = Getsubstr(strBuffer, 1, 147, 10)
        .����ʵ�ʲ��� = Val(Getsubstr(strBuffer, 1, 157, 12))
        .�Ը����� = Val(Getsubstr(strBuffer, 1, 169, 2))
        .��״̬ = Getsubstr(strBuffer, 1, 181, 14)
        .�����˻���� = Val(Getsubstr(strBuffer, 1, 195, 12))
        .���ʵ���ʹ���ۼ� = Val(Getsubstr(strBuffer, 1, 207, 12))
        .��������ʹ���ۼ� = Val(Getsubstr(strBuffer, 1, 219, 12))
        .�����Ը��ۼƽ�� = Val(Getsubstr(strBuffer, 1, 231, 12))
        .���ۼƽ�� = Val(Getsubstr(strBuffer, 1, 243, 12))
        .���ۼƽ�� = Val(Getsubstr(strBuffer, 1, 255, 12))
        .����ͳ���ۼƽ�� = Val(Getsubstr(strBuffer, 1, 267, 12))
        .סԺͳ���ۼƽ�� = Val(Getsubstr(strBuffer, 1, 279, 12))
        .סԺ���������ۼƽ�� = Val(Getsubstr(strBuffer, 1, 291, 12))
        .���ߺ��Ը��ۼƽ�� = Val(Getsubstr(strBuffer, 1, 303, 12))
        .�������ⲡ�ۼƽ�� = Val(Getsubstr(strBuffer, 1, 315, 12))
        .����Ա�ۼƽ�� = Val(Getsubstr(strBuffer, 1, 327, 12))
        .����ͳ������� = Val(Getsubstr(strBuffer, 1, 339, 3))
        .סԺͳ������� = Val(Getsubstr(strBuffer, 1, 342, 3))
        .ҽ���ۼӴ����� = Val(Getsubstr(strBuffer, 1, 345, 3))
        .סԺ��־ = Val(Getsubstr(strBuffer, 1, 348, 1))
        .������� = Val(Getsubstr(strBuffer, 1, 349, 4))
    End With
    
    ExchangeICData = True
    Exit Function
errHand:
End Function

Public Function Getsubstr(ByVal strBuffer As String, ByVal intBase As Integer, ByVal intStart As Integer, ByVal intLen As Integer) As String
    Dim intMAX As Integer   '��������ʵ�ʳ���
    '��ȡ�Ӵ����ô���ʼλ�ü�ȥ�����͵õ���ʵ����ʼλ��
    intStart = intStart - intBase + 1
    Getsubstr = Trim(StrConv(MidB(StrConv(strBuffer, vbFromUnicode), intStart, intLen), vbUnicode))
End Function

Public Function Init_����() As Boolean
    Dim intReturn As Integer
    Call Interface_Prepare_����(InitInsure, "", "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    Init_���� = True
End Function

Public Sub Interface_Prepare_����(ByVal intFunc As Integer, ByVal StrInput As String, ByVal strOutput As String, Optional ByVal str������ˮ�� As String = "")
    '���ӿڵ���ǰ��׼������
    mintFunc = intFunc
    mstrInput = StrInput
    mstrOutput = strOutput
    IC_Data_����.mstr������ˮ�� = Trim(TruncZero(str������ˮ��))
    Call DebugTool("����:" & mintFunc & ";���:" & mstrInput)
    
    Select Case mintFunc
    Case InitInsure
        mstrFunc = "IntInsure"
    Case EndInsure
        mstrFunc = "EndInsure"
    Case ReadIC
        mstrFunc = "ReadIC"
    Case GetSequence
        mstrFunc = "GetSequence"
    Case PreRegist
        mstrFunc = "PreRegist"
    Case Regist
        mstrFunc = "Regist"
    Case RegistDel
        mstrFunc = "RegistDel"
    Case PreClinic
        mstrFunc = "PreClinic"
    Case clinic
        mstrFunc = "Clinic"
    Case ClinicDel
        mstrFunc = "ClinicDel"
    Case Comein
        mstrFunc = "ComeIn"
    Case ComeIndel
        mstrFunc = "ComeInDel"
    Case ChargeDetail
        mstrFunc = "ChargeDetail"
    Case PreSettle
        mstrFunc = "PreSettle"
    Case Settle
        mstrFunc = "Settle"
    Case SettleDel
        mstrFunc = "SettleDel"
    Case ModifyPatient
        mstrFunc = "ModifyPatient"
    Case RequestBusiness
        mstrFunc = "RequestBusiness"
    Case Decide
        mstrFunc = "Decide"
    Case ConfigPara
        mstrFunc = "ConfigPara"
    Case ת��תԺ����
        mstrFunc = "ת��תԺ����"
    Case ת��תԺ��ѯ
        mstrFunc = "ת��תԺ��ѯ"
    End Select
    
End Sub

Public Function Interface_Exec_����() As Integer
    'ִ�нӿ�ָ������
    Dim intReturn  As Integer
    Dim dbl������ˮ�� As Double
    On Error GoTo errHand
    
    mstrInput = "$$" & mstrInput & "$$"
    mstrMsg = "$$" & String(3000, " ") & "$$"
    dbl������ˮ�� = CDbl(Val(IC_Data_����.mstr������ˮ��))
    Select Case mintFunc
    Case InitInsure
        intReturn = LHYB_Init(mstrInput, mstrMsg, mstrOutput)
    Case EndInsure
        intReturn = LHYB_Close(mstrInput, mstrMsg, mstrOutput)
        Interface_Exec_���� = (intReturn >= 0)
        Exit Function
    Case Else
        intReturn = LHYB_Business(mintFunc, dbl������ˮ��, mstrInput, mstrMsg, mstrOutput)
    End Select
    
    Call DebugTool("������Ϣ:" & mstrMsg)
    mstrMsg = Replace(mstrMsg, "$$", "")
    Interface_Exec_���� = intReturn
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Interface_Exec_���� = -1
    mstrMsg = "-1~δ֪����~~~~~~~~~~~~"
End Function

Public Function Interface_Analyse_����() As Boolean
    '�����ӿڷ��ص�����
    
End Function

Public Function ҽ����ʼ��_����(Optional ByVal blnTest As Boolean = False) As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim strServer As String, strUser As String, strPass As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    If mblnInit = False Then
        If Not blnTest Then '����ǲ��ԣ���˵���Ǳ��ղ������ô�����
            '��������ҽ��������������
            gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_����)
            
            Do Until rsTemp.EOF
                Select Case rsTemp("������")
                    Case "ҽ���û���"
                        strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Case "ҽ��������"
                        strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                    Case "ҽ���û�����"
                        strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                End Select
                rsTemp.MoveNext
            Loop
            
            If OraDataOpen(gcn����, strServer, strUser, strPass, False) = False Then
                MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If Not Init_����() Then Exit Function
        
        'ȡҽԺ����
        gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_����)
        IC_Data_����.mstrҽԺ���� = Nvl(rsTemp!ҽԺ����)
        'ȡҽԺ�ȼ�
        If IC_Data_����.mstrҽԺ���� <> "" Then
            gstrSQL = "Select YYDJ From SIM_YLJG Where YYBH='" & IC_Data_����.mstrҽԺ���� & "'"
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn����
            IC_Data_����.mstrҽԺ�ȼ� = Nvl(rsTemp!YYDJ)
        End If
        
        If Not blnTest Then mblnInit = True
    End If
    
    ҽ����ʼ��_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.��������
End Function

Public Function ҽ����ֹ_����() As Boolean
    Call Interface_Prepare_����(EndInsure, "", "")
    Call Interface_Exec_����
    
    mblnInit = False
    ҽ����ֹ_���� = True
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ҽ��������ݵ�ʶ��
    ��ݱ�ʶ_���� = frmIdentify����.GetPatient(bytType, lng����ID)
End Function

Private Function GetסԺ��(ByVal lng����ID As Long, Optional ByVal blnNew As Boolean = False) As String
    Dim strText As String
    Dim str����ʱ�� As String
    Dim str��ǰʱ�� As String
    Dim strSequence As String
    Dim rsTemp As New ADODB.Recordset
    Dim intDO As Integer, intCOUNT As Integer, intPos As Integer
    '����ǰʱ�䡢����ʱ����д���ת��ΪΨһ����ˮ�ű�ʶ
    '���˼·�����ꡢ�¡��ա�ʱ���֡��붼ת��Ϊһ����ĸ����ʽ��ʾ����Ϊһ��ֻ��12λ
    intCOUNT = 6
    intPos = 1
    str��ǰʱ�� = Format(zlDatabase.Currentdate, "yyMMddHHmmss")
    
    '��ȡ�ò��˵ľ���ʱ��
    gstrSQL = "Select ����֤�� From �����ʻ�" & _
        " Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˵ľ���ʱ��", TYPE_����, lng����ID)
    '�϶�����Ϊ��
    If Not blnNew Then
        GetסԺ�� = Nvl(rsTemp!����֤��)
    Else
        GetסԺ�� = CStr(zlDatabase.GetNextID("���ű�"))
    End If
End Function

Private Function Get�Ը�����(ByVal int��Ŀ���� As Integer, ByVal strҽ������ As String, Optional ByVal str�������� As String = "11") As Double
    Dim rsTemp As New ADODB.Recordset
    '����ָ����Ŀ���Ը������뵥���޶�
    '�������˵����
'    (sLbbz in hi_zymx.lbbz%type,   -- 1ҩƷ2����
'     sXmbh in hi_zymx.xmbh%type,   -- ��Ŀ���
'     iDylb in sio_ybdyzb.dylb%type,-- �������
'     sYydj in sio_ybfdjs.yydj%type,-- ҽԺ�ȼ�
'     sJzlx in sio_jzlx.jzlx%type,  -- ��������
'     iDfff in number,              -- 1����2����
'     nJbbm in sim_jbda.jbbm%type   -- ��������
'     ) return number is
    gstrSQL = "Select orafGetzfbl(" & int��Ŀ���� & ",'" & strҽ������ & "','" & IC_Data_����.������� & "'," & _
        "'" & IC_Data_����.mstrҽԺ�ȼ� & "','" & str�������� & "',1," & IC_Data_����.mlng����ID & ") from dual"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcn����
    Get�Ը����� = Nvl(rsTemp.Fields(0).Value, 0)
End Function

Private Function Get������ˮ��() As String
    '����������ˮ��
    '10λ���֣�����ȡ���ű�����к�ʮλ
    Get������ˮ�� = Right(CStr(zlDatabase.GetNextID("���ű�")), 10)
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strBill As String                   '����ͷ
    Dim strDetail As String                 '��ϸ��
    Dim strErrInfo As String
    Dim lngPatient As Long                  '����ID
    Dim intReturn As Integer
    Dim strBalance As String                '�ӿڷ��صĽ�����Ϣ
    Dim arrBalance
    Dim strҵ������ As String               '�����ʻ��е�ҵ�����ͣ�Ϊ1��ʾ����
    Dim strDepart As String, strDoctor As String
    Dim lngDisease As Long, strDiseaseCode As String, strDiseaseName As String
    Dim dbl�����ܶ� As Double, dbl��ҽ���ܶ� As Double, dbl�Ը����� As Double
    Dim int��ϸ�� As Integer, int������ As Integer
    Dim StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    
    Const int�����ܶ� As Integer = 0
    Const int�Է��ܶ� As Integer = 1
    Const int������� As Integer = 2
    Const intͳ����� As Integer = 3
    Const int�����ʻ� As Integer = 4
    Const int�����ʻ� As Integer = 5
    Const int�󲡾��� As Integer = 6
    Const int����Ա���� As Integer = 7
    Const int��λ֧�� As Integer = 8
    Const int�����Ը� As Integer = 9
    
    On Error GoTo errHand
    
    IC_Data_����.mstr���ﵥ�ݺ� = Get������ˮ��
    IC_Data_����.mstr������ = Format(zlDatabase.Currentdate(), "yyyyMMddHHmmss")
'    ��ڲ��� (Data)
'    �Ƿ���ҽ���� + IC��Ϣ + ��~�� + ���ν��㵥������ + ҽ���շ���Ŀ�б�Clinic
'    Clinic�ṹ�壨[]��ʾ�����ظ�����Clinic = [Bill(����)] + [Prescription����ϸ��]��
'    ����Һź��շ�ʱ����Ҫͨ�����ַ����ṹ���������ݴ��ݵ����㺯���С����ݺ�BillID�ǵ���Ψһ�ı�־����HIS�����ظ������м���������δ֪ʱ��д"0"����ҽ���ܶ�ָ"����ҽ��Ŀ¼��Χ����Ŀ"���ܶ������ҽ��Ŀ¼��Χ����Ŀ���Ը��������֡�������ַ����ṹ�����¸�ʽ��װ���м���%%�ָ�����������ת��Ϊyyyy.mm.dd��ʽ����
'    Bill = ���ݺ�(N10) + �����(N10) +��������(VC15) +��������(Dt) +�շ�����(N1:0����Һţ�1�����շѣ�2�����շ�) +��������(VC20) +ҽ������(VC10)+�������(VC12)+��������(VC50)+��������(VC255) + �˵����з�ҽ����Ŀ�ܶ�N(12,2) + �˵������շ���ϸ����(count)Integer����������ҽ����¼����ϸ��¼����Ӧ����Countֵ����
'    Prescription = ���ݺ���(N10)+ҩƷ��������(N1:1ҩƷ��2����)+��Ŀ���(N10)+��ĿҽԺ������(VC80)+ ҽԺ�˹��(VC20) + ��������־(N1:0�ǲ�ҩ(�����ڵ�����)��1��ҩ������2��ҩ����) + ����N(14,4) + ����N(14,4)+�Ը�����N(5,4)��
    With rs��ϸ
        'ȡ�����ܶ�
        Do While Not .EOF
            lngPatient = !����ID
            dbl�����ܶ� = dbl�����ܶ� + Nvl(!ʵ�ս��, 0)
            
            '�ж��Ƿ�ҽ����Ŀ
            gstrSQL = "Select B.����,B.���,A.��Ŀ���� AS ҽ������ " & _
                    "From ����֧����Ŀ A,�շ�ϸĿ B " & _
                    "Where B.ID=[1] And A.�շ�ϸĿID=B.ID And A.����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����Ŀ��š���Ŀ���Ƽ����", CLng(!�շ�ϸĿID), TYPE_����)
            If rsTemp.RecordCount = 0 Then
                dbl��ҽ���ܶ� = dbl��ҽ���ܶ� + Nvl(!ʵ�ս��, 0)
            Else
                If Nvl(!ʵ�ս��, 0) <> 0 Then int��ϸ�� = int��ϸ�� + 1
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        'ֻ����ҽ����Ŀ��ϸ
        Do While Not .EOF
            If Nvl(!�Ƿ�ҽ��, 0) <> 0 And Nvl(!ʵ�ս��, 0) <> 0 Then
                '������������־
                int������ = IIf(!�շ���� = "7", 1, 0)
                
                '��ȡҽ����Ŀ��š���Ŀ���Ƽ����
                        
                ''''�¶� 20041228
               'gstrSQL = "Select B.����,B.���,A.��Ŀ���� AS ҽ������ " & _
               '         "From ����֧����Ŀ A,�շ�ϸĿ B " & _
               '         "Where B.ID=" & !�շ�ϸĿID & " And A.�շ�ϸĿID=B.ID And A.����=" & TYPE_����
               
                gstrSQL = "Select C.���� as ����,B.����,B.���,A.��Ŀ���� AS ҽ������ " & _
                        "From ����֧����Ŀ A,�շ�ϸĿ B,����֧������ C " & _
                        "Where B.ID=[1] And A.�շ�ϸĿID=B.ID  And A.����ID=C.ID And A.����=[2]"
                
                ''''
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����Ŀ��š���Ŀ���Ƽ����", CLng(!�շ�ϸĿID), TYPE_����)
                
                If rsTemp.RecordCount <> 0 Then
                    '��ȡ�Ը�����
                    
                    ' �¶�  20041228
                    'dbl�Ը����� = Get�Ը�����(IIf(InStr(1, "5,6,7", !�շ����) > 0, 1, 2), rsTemp!ҽ������, "11")
                    dbl�Ը����� = Get�Ը�����(IIf(rsTemp!���� = "ҩƷ", 1, 2), rsTemp!ҽ������, IC_Data_����.mstrҵ������)
                    
                    
                    If dbl�Ը����� < 0 Then
                        MsgBox "��Ŀ[" & rsTemp!���� & "]���Ը�������ȡ����", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    
                    ' �¶�  20041228
                    'strDetail = strDetail & strSplit & IC_Data_����.mstr���ﵥ�ݺ� & strSplit & IIf(InStr(1, "5,6,7", !�շ����) <> 0, 1, 2) & strSplit & _
                    '    rsTemp!ҽ������ & strSplit & ToVarchar(rsTemp!����, 80) & strSplit & ToVarchar(Nvl(rsTemp!���), 20) & strSplit & _
                   '    int������ & strSplit & Format(!����, "#0.0000") & strSplit & _
                   '     Format(!����, "#0.0000") & strSplit & dbl�Ը�����
                        
                    strDetail = strDetail & strSplit & IC_Data_����.mstr���ﵥ�ݺ� & strSplit & IIf(rsTemp!���� = "ҩƷ", 1, 2) & strSplit & _
                        rsTemp!ҽ������ & strSplit & ToVarchar(rsTemp!����, 80) & strSplit & ToVarchar(Nvl(rsTemp!���), 20) & strSplit & _
                       int������ & strSplit & Format(!����, "#0.0000") & strSplit & _
                        Format(!����, "#0.0000") & strSplit & dbl�Ը�����
                        
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        If strDetail <> "" Then strDetail = Mid(strDetail, 3)
    End With
    
    '��ȡ�ò��˵ļ�����Ϣ
    lngDisease = IC_Data_����.mlng����ID
    
    strDiseaseCode = "0"
    strDiseaseName = "δ֪"
    If lngDisease <> 0 Then
        gstrSQL = "Select JBBZDM,JBMC From SIM_JBDA Where JBBM=" & lngDisease
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn����
        If rsTemp.RecordCount <> 0 Then
            strDiseaseCode = lngDisease
            strDiseaseName = rsTemp!JBMC
        End If
    End If
    
    '��ȡ�õ��ݵĿ�������
    strDoctor = Trim(Nvl(rs��ϸ!������))
    If strDoctor <> "" Then
        gstrSQL = "SELECT C.���� AS �������� " & _
                 " FROM ������Ա A,��������˵�� B,���ű� C " & _
                 " WHERE A.��ԱID= " & _
                 "     (SELECT ID FROM ��Ա�� WHERE ����=[1]) " & _
                 " AND A.����ID=B.����ID AND A.����ID=C.ID AND B.��������='�ٴ�' AND ������� IN (1,3) " & _
                 " AND ROWNUM<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", CStr(rs��ϸ!������))
        strDepart = "0"
        If rsTemp.RecordCount <> 0 Then strDepart = Nvl(rsTemp!��������)
    End If
    
'    strBill = IC_Data_����.mstr���ﵥ�ݺ� & strSplit & IC_Data_����.mstr���ﵥ�ݺ� & strSplit & _
'        IC_Data_����.mstr������ & strSplit & Format(zlDatabase.Currentdate, "yyyy.MM.dd") & strSplit & _
'        IIf(strҵ������ = "1", "2", "1") & strSplit & strDepart & strSplit & strDoctor & strSplit & _
'        strDiseaseCode & strSplit & strDiseaseName & strSplit & strSplit & _
'        Format(dbl��ҽ���ܶ�, "#0.00") & strSplit & int��ϸ��
    'Modified by ZYB 2006-04-12���̶�����1����ʾ�����շѣ��� IIf(strҵ������ = "1", "2", "1") �滻Ϊ "1"
    strBill = IC_Data_����.mstr���ﵥ�ݺ� & strSplit & IC_Data_����.mstr���ﵥ�ݺ� & strSplit & _
        IC_Data_����.mstr������ & strSplit & Format(zlDatabase.Currentdate, "yyyy.MM.dd") & strSplit & _
        "1" & strSplit & strDepart & strSplit & strDoctor & strSplit & _
        strDiseaseCode & strSplit & strDiseaseName & strSplit & strSplit & _
        Format(dbl��ҽ���ܶ�, "#0.00") & strSplit & int��ϸ��
    
    'Modified by ZYB 2006-04-12������������ҽ������2006-04-06�·����ļ�Ҫ���޸ģ����������Ӿ�������
    StrInput = "1" & strField & IC_Data_����.IC������ & strField & strField & strField & "1" & strField & strBill & strSplit & strDetail & strField & IC_Data_����.mstrҵ������
    Call Interface_Prepare_����(PreClinic, StrInput, strOutput)
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
'        If (Trim(Split(mstrMsg, "~")(7)) <> "") Then
'            strErrInfo = strErrInfo & vbCrLf & "��ϸ��Ϣ��" & vbCrLf & Trim(Split(mstrMsg, "~")(7))
'        End If
        MsgBox strErrInfo, vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    
    'ȡ������Ϣ
    strBalance = Split(mstrMsg, "~")(6)
    '������(�м���%%�ָ�)���ٷ����ܶ�+���Է��ܶ�(��ҽ����ֻ���ֽ�)+�������ܶ�(Ŀ¼���Ը���������)+
    '��ͳ�����֧��+�������ʻ�֧��+�޵����ʻ�֧��+�ߴ󲡾���֧��+�๫��Ա����֧��+�ᵥλ֧�� + '
    '������Ը� (�˻������ֽ�) + �����ֽ�֧�� + ���������ز�����ǰ�����ۼ�
    '�����ܶ��=��+��+��+��+��+��+��+��+�⣬�ֽ�֧��=��+��+��+��
    arrBalance = Split(strBalance, "%%")
    str���㷽ʽ = "�����ʻ�;" & Val(arrBalance(int�����ʻ�)) + Val(arrBalance(int�����ʻ�)) & ";0"
    If Val(arrBalance(intͳ�����)) <> 0 Then str���㷽ʽ = str���㷽ʽ & "|ͳ�����;" & Val(arrBalance(intͳ�����)) & ";0"
    If Val(arrBalance(int����Ա����)) <> 0 Then str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & Val(arrBalance(int����Ա����)) & ";0"
    If Val(arrBalance(int�󲡾���)) <> 0 Then str���㷽ʽ = str���㷽ʽ & "|�󲡾���;" & Val(arrBalance(int�󲡾���)) & ";0"
    
    IC_Data_����.mstr������ڲ����� = "1" & strField & strBill & strSplit & strDetail
    �����������_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    Dim intReturn As Integer
    Dim lng����ID As Long
    Dim blnTrans As Boolean
    Dim str������ˮ�� As String
    Dim strErrInfo As String
    Dim strBalance As String
    Dim arrBalance
    Dim StrInput As String, strOutput As String
    Dim dblͳ����� As Double, dbl�󲡲��� As Double, dbl����Ա���� As Double, dbl��λ֧�� As Double
    Dim dbl�����ʻ� As Double, dbl�����ʻ� As Double, dbl�����ʻ�_��� As Double, dbl�����ʻ�_��� As Double
    Dim dbl�����ܶ� As Double, dbl�ֽ�֧�� As Double    '�ֽ�֧��������λ֧��
    Dim rsTemp As New ADODB.Recordset
    
    Const int�����ܶ� As Integer = 0
    Const int�Է��ܶ� As Integer = 1
    Const int������� As Integer = 2
    Const intͳ����� As Integer = 3
    Const int�����ʻ� As Integer = 4
    Const int�����ʻ� As Integer = 5
    Const int�󲡾��� As Integer = 6
    Const int����Ա���� As Integer = 7
    Const int��λ֧�� As Integer = 8
    Const int�����Ը� As Integer = 9
    
    On Error GoTo errHand
    'ȡ����ID
    gstrSQL = "Select ����ID From ������ü�¼ Where ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
    lng����ID = rsTemp!����ID
    
    'ȡ���
    gstrSQL = "Select �����ʻ����,�����ʻ���� From �����ʻ� Where ����=[2] And ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���", lng����ID, TYPE_����)
    dbl�����ʻ�_��� = Nvl(rsTemp!�����ʻ����, 0)
    dbl�����ʻ�_��� = Nvl(rsTemp!�����ʻ����, 0)
    
    'ȡ��ˮ��
    Call Interface_Prepare_����(GetSequence, "~~~~" & clinic, "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
        If (intReturn < 0) Then Exit Function
    End If
    str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    '�������
    'ȱʡ�ֽ�֧��
    'Modified by ZYB 2006-04-12������������ҽ������2006-04-06�·����ļ�Ҫ���޸ģ����������Ӿ�������
    StrInput = "1~~1~~" & IC_Data_����.mstr������ڲ����� & strField & UserInfo.���� & strField & IC_Data_����.mstrҵ������
    Call Interface_Prepare_����(clinic, StrInput, strOutput, str������ˮ��)
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
'        If Val(Split(mstrMsg, "~")(2)) <> 0 Then
'            strErrInfo = strErrInfo & vbCrLf & "��ϸ��Ϣ��" & vbCrLf & "д��ʧ�ܣ������д������ԭ�����¶����򻻻������¶�һ�鿨���Զ�ͬ�������ݣ�"
'        End If
        Err.Raise 9000, gstrSysName, strErrInfo
        If (intReturn < 0) Then Exit Function
    End If
    blnTrans = True
    
    'ȡ������Ϣ
    strBalance = Split(mstrMsg, "~")(6)
    '������(�м���%%�ָ�)���ٷ����ܶ�+���Է��ܶ�(��ҽ����ֻ���ֽ�)+�������ܶ�(Ŀ¼���Ը���������)+
    '��ͳ�����֧��+�������ʻ�֧��+�޵����ʻ�֧��+�ߴ󲡾���֧��+�๫��Ա����֧��+�ᵥλ֧�� + '
    '������Ը� (�˻������ֽ�) + �����ֽ�֧�� + ���������ز�����ǰ�����ۼ�
    '�����ܶ��=��+��+��+��+��+��+��+��+�⣬�ֽ�֧��=��+��+��+��
    arrBalance = Split(strBalance, "%%")
    dbl�����ܶ� = Val(arrBalance(int�����ܶ�))
    dbl�����ʻ� = Val(arrBalance(int�����ʻ�))
    dbl�����ʻ� = Val(arrBalance(int�����ʻ�))
    dblͳ����� = Val(arrBalance(intͳ�����))
    dbl����Ա���� = Val(arrBalance(int����Ա����))
    dbl�󲡲��� = Val(arrBalance(int�󲡾���))
    dbl��λ֧�� = Val(arrBalance(int��λ֧��))
    dbl�ֽ�֧�� = dbl�����ܶ� - dblͳ����� - dbl�󲡲��� - dbl����Ա���� - dbl�����ʻ� - dbl�����ʻ�
    
    '���Ը�=�󲡲���;�����Ը�=����Ա����
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(zlDatabase.Currentdate()) & ",0,0,0,0," & dbl�����ʻ�_��� & "," & dbl�����ʻ�_��� & "," & dbl�����ʻ� & "," & dbl�����ʻ� & "," & dbl�����ܶ� & "," & dbl�ֽ�֧�� & ",0," & _
        dblͳ����� & "," & dblͳ����� & "," & dbl�󲡲��� & "," & dbl����Ա���� & "," & dbl�����ʻ� + dbl�����ʻ� & ",'" & IC_Data_����.mstr������ˮ�� & "|" & IC_Data_����.������� & "|" & IC_Data_����.mstrҵ������ & "',NULL,NULL,'" & Replace(Split(mstrMsg, "~")(6), "'", "") & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '����ȷ��
    '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
    'ȱʡ�ֽ�֧��
    StrInput = "~~~~" & clinic & strField & IC_Data_����.mstr������ˮ�� & strField & "0" & strField & "HIS�ɹ���"
    Call Interface_Prepare_����(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_����
    If intReturn < 0 Then
        Err.Raise 9000, gstrSysName, "���棺���ν���ȷ��ʧ�ܣ����¼�±��ν�����ˮ�ţ���֪ͨϵͳ����Աʹ�ù��߰��ٴ�ȷ�ϸý���" & _
        vbCrLf & "������ˮ�ţ�" & str������ˮ��
    End If
    �������_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then
        '����ȷ��
        '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
        'ȱʡ�ֽ�֧��
        StrInput = "~~~~" & clinic & strField & str������ˮ�� & strField & "-1" & strField & "ҽ���ɹ�����HISʧ�ܣ�"
        Call Interface_Prepare_����(Decide, StrInput, strOutput)
    End If
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    Dim intReturn As Integer
    Dim lng����ID As Long
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim blnTrans As Boolean
    Dim str������ˮ�� As String, strԭ������ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'ȡ��ˮ��
    Call Interface_Prepare_����(GetSequence, "~~~~" & ClinicDel, "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
        If (intReturn < 0) Then Exit Function
    End If
    str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    'ȡ������¼�Ľ���ID
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng����ID = rsTemp!����ID
    
    'ȡԭ������ˮ��
    gstrSQL = "Select ֧��˳��� From ���ս����¼ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡԭ������ˮ��", lng����ID)
    strԭ������ˮ�� = Split(rsTemp!֧��˳���, "|")(0)
    
    '�������Ͻ���
    '�Ƿ���ҽ���� + IC��Ϣ(���Դ���)+ ~��~�� + Ҫ���ϵ�����/�ҺŽ��㽻�׺�
    Call Interface_Prepare_����(ClinicDel, "1~~~~" & strԭ������ˮ��, "", str������ˮ��)
    intReturn = Interface_Exec_����()
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
        Err.Raise 9000, gstrSysName, strErrInfo
        If (intReturn < 0) Then Exit Function
    End If
    blnTrans = True
    
    '��ȡԭ�����¼����Ϊ�������ν����¼������
    gstrSQL = "Select * From ���ս����¼ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭ�����¼����Ϊ�������ν����¼������", lng����ID)
    
    '��������¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(zlDatabase.Currentdate()) & ",0,0,0,0,0,0,0,0," & -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & ",0," & _
        -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & "," & -1 * Nvl(rsTemp!���Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & _
        -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & str������ˮ�� & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '����ȷ��
    '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
    'ȱʡ�ֽ�֧��
    StrInput = "~~~~" & ClinicDel & strField & str������ˮ�� & strField & "0" & strField & "HIS�ɹ���"
    Call Interface_Prepare_����(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_����
    If intReturn < 0 Then
        Err.Raise 9000, gstrSysName, "���棺���ν���ȷ��ʧ�ܣ����¼�±��ν�����ˮ�ţ���֪ͨϵͳ����Աʹ�ù��߰��ٴ�ȷ�ϸý���" & _
            vbCrLf & "������ˮ�ţ�" & str������ˮ��
    End If
    ����������_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then
        '����ȷ��
        '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
        'ȱʡ�ֽ�֧��
        StrInput = "~~~~" & ClinicDel & strField & str������ˮ�� & strField & "-1" & strField & "ҽ���ɹ�����HISʧ�ܣ�"
        Call Interface_Prepare_����(Decide, StrInput, strOutput)
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim intReturn As Integer
    Dim strErrInfo As String
    Dim strסԺ�� As String, str������ˮ�� As String
    Dim str��Ժ���� As String, strҽ�� As String, str��Ժ��� As String, str������� As String
    Dim lng����ID As Long, lng����ID As Long, str�������� As String, str���ұ�� As String, str���� As String
    Dim StrInput As String, strOutput As String
    Dim blnTrans As Boolean
    Dim str�޿����� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'ȡ��ˮ��
    Call Interface_Prepare_����(GetSequence, "~~~~" & Comein, "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    ',D.��������  ',�ٴ����� D  ' And B.ID=D.����ID(+)
    '��ȡ��Ժ���ڡ�ҽ������Ժ��ϡ���Ժ��������Ժ�������ơ�ҽ�����ұ�š�����
    gstrSQL = " Select to_char(A.��Ժ����,'yyyy-MM-dd') ��Ժ����,B.ID AS ��Ժ����ID,B.���� as ���ұ���," & _
              " B.���� ����,A.סԺҽʦ ҽ��,A.��Ժ����,C.����ID " & _
              " From ������ҳ A,���ű� B,�����ʻ� C " & _
              " Where A.����ID=[1] And A.��ҳID=[2]" & _
              " And A.��Ժ����ID=B.ID And A.����ID=C.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ������Ϣ", lng����ID, lng��ҳID)
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy.MM.dd")
    strҽ�� = Nvl(rsTemp!ҽ��)
    str�������� = Nvl(rsTemp!����)
    lng����ID = Nvl(rsTemp!��Ժ����ID, 0)
    str���ұ�� = Nvl(rsTemp!���ұ���)
    str���� = Nvl(rsTemp!��Ժ����)
    lng����ID = Nvl(rsTemp!����ID, 0)
    'ȡ��Ժ���
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True)
'    'ȡҽ���˼�����׼����
'    gstrSQL = "Select JBBZDM From SIM_JBDA Where JBBM=" & lng����ID
'    If rsTemp.State = 1 Then rsTemp.Close
'    rsTemp.Open gstrSQL, gcn����
'    If rsTemp.RecordCount <> 0 Then
'        str������� = Nvl(rsTemp!JBBZDM, "0")
'    Else
'        str������� = "0"
'    End If
    
    'סԺȷ�ϳɹ���סԺ�ű�ʹ�ã�סԺȷ��ʧ�ܺ�סԺ�ű����ϣ�סԺ�Ŷ�������ʹ��
    strסԺ�� = GetסԺ��(lng����ID, True)
    str�޿����� = IS�޿�����(lng����ID)
    '�Ƿ���ҽ���� + IC��Ϣ(���Դ���)+ ~��~�� + סԺ�� + ��Ժ���� + ��Ժ���ҽ������ + ��Ժ�������� + ��Ժ������� + ��Ժ�������� + ��Ժ����ID + ���š�
    StrInput = Split(str�޿�����, "|")(0) & strField & Split(str�޿�����, "|")(1) & strField & strField & strField & strסԺ�� & strField & _
        str��Ժ���� & strField & strҽ�� & strField & str��Ժ��� & strField & lng����ID & strField & _
        str�������� & strField & lng����ID & strField & str����
    Call Interface_Prepare_����(Comein, StrInput, strOutput, str������ˮ��)
    intReturn = Interface_Exec_����()
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
        If Val(Split(mstrMsg, "~")(2)) <> 0 Then
            strErrInfo = strErrInfo & vbCrLf & "��ϸ��Ϣ��" & vbCrLf & "д��ʧ�ܣ������д������ԭ�����¶����򻻻������¶�һ�鿨���Զ�ͬ�������ݣ�"
        End If
        MsgBox strErrInfo, vbInformation, gstrSysName
        If intReturn < 0 Then Exit Function
    End If
    blnTrans = True
    
    '����ȷ��
    '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
    'ȱʡ�ֽ�֧��
    StrInput = "~~~~" & Comein & strField & str������ˮ�� & strField & "0" & strField & "HIS�ɹ���"
    Call Interface_Prepare_����(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_����
    If intReturn < 0 Then MsgBox "���棺���ν���ȷ��ʧ�ܣ����¼�±��ν�����ˮ�ţ���֪ͨϵͳ����Աʹ�ù��߰��ٴ�ȷ�ϸý���" & _
        vbCrLf & "������ˮ�ţ�" & str������ˮ��, vbInformation, gstrSysName
        
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'˳���','''" & str������ˮ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����˳���")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','''" & strסԺ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��")
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then
        '����ȷ��
        '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
        'ȱʡ�ֽ�֧��
        StrInput = "~~~~" & Comein & strField & str������ˮ�� & strField & "-1" & strField & "ҽ���ɹ�����HISʧ�ܣ�"
        Call Interface_Prepare_����(Decide, StrInput, strOutput)
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim intReturn As Integer
    Dim strErrInfo As String
    Dim str������ˮ�� As String, strסԺ�� As String, str�޿����� As String
    Dim StrInput As String, strOutput As String
    Dim blnAllow As Boolean '�Ƿ���������Ժ
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'δ�����Ϊ�㣬��δ���н������������ȡ����Ժ�Ǽ�
    
    blnAllow = True
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        '�жϸò����Ƿ�������û�н�����Ĳ��˷���Ϊ�㣬˵����Ҫ���þ���Ǽǳ���
        gstrSQL = "Select 1 From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] And Nvl(����ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���й����ý���", lng����ID, lng��ҳID)
        If Not rsTemp.EOF Then
            blnAllow = False
        End If
    Else
        blnAllow = False
    End If
    
    If Not blnAllow Then
        MsgBox "�ò��˴���δ����û��ѽ��й�סԺ���㣬����������Ժ��" & vbCrLf & _
        "��ֻ������ã���δ���й�����Ĳ��ˣ���������ҽ����Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ��ˮ��
    Call Interface_Prepare_����(GetSequence, "~~~~" & ComeIndel, "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    strסԺ�� = GetסԺ��(lng����ID)
    str�޿����� = IS�޿�����(lng����ID)
    
    '�Ƿ���ҽ���� + IC��Ϣ(���Դ���)+ �� + �� + Ҫע����סԺ��
    StrInput = Split(str�޿�����, "|")(0) & strField & Split(str�޿�����, "|")(1) & strField & strField & strField & strסԺ��
    Call Interface_Prepare_����(ComeIndel, StrInput, strOutput, str������ˮ��)
    intReturn = Interface_Exec_����()
    '���׽��~������Ϣ+дҽ������� + �� + д����IC������
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
        If Val(Split(mstrMsg, "~")(2)) <> 0 Then
            strErrInfo = strErrInfo & vbCrLf & "��ϸ��Ϣ��" & vbCrLf & "д��ʧ�ܣ������д������ԭ�����¶����򻻻������¶�һ�鿨���Զ�ͬ�������ݣ�"
        End If
        MsgBox strErrInfo, vbInformation, gstrSysName
        If intReturn < 0 Then Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ��ת��ͨ����_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim intReturn As Integer
    Dim strErrInfo As String
    Dim str������ˮ�� As String, strסԺ�� As String, str�޿����� As String
    Dim StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'δ�����Ϊ�㣬��δ���н������������ȡ����Ժ�Ǽ�
    
    '��Ӧ���ܵ�Ҫ�����ε�
    '�жϸò����Ƿ�������û�н�����Ĳ��˷���Ϊ�㣬˵����Ҫ���þ���Ǽǳ���
'    gstrSQL = "Select 1 From ���˷��ü�¼ Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID & " And Nvl(����ID,0)<>0 and Rownum<2"
'    Call OpenRecordset(rsTemp, "�ж��Ƿ���й����ý���")
'    If Not rsTemp.EOF Then
'        MsgBox "�ò����ѽ��й�סԺ���㣬������תΪ��ͨ���ˣ�", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    'ȡ��ˮ��
    Call Interface_Prepare_����(GetSequence, "~~~~" & ҽ��ת�ԷѲ���, "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    strסԺ�� = GetסԺ��(lng����ID)
    str�޿����� = IS�޿�����(lng����ID)
    
    '�Ƿ���ҽ���� + IC��Ϣ(���Դ���)+ �� + �� + Ҫע����סԺ��
    StrInput = Split(str�޿�����, "|")(0) & strField & Split(str�޿�����, "|")(1) & strField & strField & strField & strסԺ��
    Call Interface_Prepare_����(ҽ��ת�ԷѲ���, StrInput, strOutput, str������ˮ��)
    intReturn = Interface_Exec_����()
    '���׽��~������Ϣ+дҽ������� + �� + д����IC������
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
        If Val(Split(mstrMsg, "~")(2)) <> 0 Then
            strErrInfo = strErrInfo & vbCrLf & "��ϸ��Ϣ��" & vbCrLf & "д��ʧ�ܣ������д������ԭ�����¶����򻻻������¶�һ�鿨���Զ�ͬ�������ݣ�"
        End If
        MsgBox strErrInfo, vbInformation, gstrSysName
        If intReturn < 0 Then Exit Function
    End If
    
    gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
    MsgBox "�����ɹ�����ҽ�������Ѿ�תΪ��ͨ���ˣ�", vbInformation, gstrSysName
    
    ҽ��ת��ͨ����_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim intReturn As Integer
    Dim lng��ϸID As Long
    Dim lng��ҳID As Long
    Dim dbl�Ը����� As Double
    Dim blnUpload As Boolean
    Dim strBalance As String
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim strҽ������ As String, strҽ������ As String, strҽ����� As String, strҽ����λ As String
    Dim strסԺ��ˮ�� As String, strסԺ�� As String, strҽ��֤�� As String, str������ˮ�� As String
    Dim int��¼�� As Integer, dbl��ҽ����� As Double
    Dim str�޿����� As String
    Dim dbl�����ܶ� As Double, dbl�����ܶ�_YB As Double, dbl�Է��ܶ� As Double, dbl�����ܶ� As Double, dblͳ����� As Double
    Dim dbl�ʻ�֧�� As Double, dbl�󲡲��� As Double, dbl����Ա���� As Double, dbl��λ֧�� As Double, dbl���� As Double
    
    Const int�����ܶ� As Integer = 0
    Const intͳ����� As Integer = 3
    Const int�����ʻ� As Integer = 4
    Const int�����ʻ� As Integer = 5
    Const int�󲡾��� As Integer = 6
    Const int����Ա���� As Integer = 7
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim gcn�ϴ� As New ADODB.Connection
    
    On Error GoTo errHand
    
    '������
    Set gcn�ϴ� = GetNewConnection
    
    '����ȡ�ò��˵�סԺ��ˮ��
    gstrSQL = "Select A.ҵ������,A.ҽ����,A.˳���,A.IC,A.����ID,B.סԺ���� ��ҳID From �����ʻ� A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ȡ�ò��˵�סԺ��ˮ��", lng����ID)
    strҽ��֤�� = rsTemp!ҽ����
    strסԺ��ˮ�� = rsTemp!˳���
    lng��ҳID = rsTemp!��ҳID
    IC_Data_����.IC������ = Nvl(rsTemp!ic)
    IC_Data_����.mlng����ID = Nvl(rsTemp!����ID, 0)
    strסԺ�� = GetסԺ��(lng����ID)
    'Modified by ZYB 2006-04-12������������ҽ������2006-04-06�·����ļ�Ҫ���޸ģ����������Ӿ�������
    IC_Data_����.mstrҵ������ = Nvl(rsTemp!ҵ������, "21")
    
    Call ExchangeICData(IC_Data_����.IC������)
    
    '��ȡ���η�����ϸ
    '�¶� 20041228
    'gstrSQL = "Select A.ID,A.NO,A.����ID,A.�շ����,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,C.��Ŀ���� AS ҽ����Ŀ����,B.����,B.����,A.ʵ�ս�� AS ���" & _
    '          "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸�,A.������ AS ҽ��,A.�Ǽ�ʱ�� " & _
    '          "  From ���˷��ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
    '          "  where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & " and A.���ʷ���=1 And A.����Ա���� is not null AND A.ʵ�ս�� IS NOT NULL " & _
    '          "        and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= " & TYPE_���� & _
    '          "  Order by A.����ID,A.����ʱ��"
    gstrSQL = "Select D.���� as ����,A.ID,A.NO,A.����ID,A.�շ����,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,C.��Ŀ���� AS ҽ����Ŀ����,B.����,B.����,A.ʵ�ս�� AS ���" & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸�,A.������ AS ҽ��,A.�Ǽ�ʱ�� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,����֧������ D " & _
              "  where A.����ID=[1] and A.��ҳID=[2] and A.���ʷ���=1 And A.����Ա���� is not null AND Nvl(A.ʵ�ս��,0)<>0 " & _
              "        and nvl(A.�Ƿ��ϴ�,0)=0 And C.����ID=D.ID And Nvl(A.��¼״̬,0)<>0 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= [3]" & _
              "  Order by A.����ID,A.����ʱ��"
    
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���η�����ϸ", lng����ID, lng��ҳID, TYPE_����)
    
    '��ɾ��������ϸ
    gstrSQL = "Delete Hi_zymx_temp Where JYH='" & strסԺ��ˮ�� & "'"
    gcn����.Execute gstrSQL
    
    With rs��ϸ
        Do While Not .EOF
            strҽ������ = Nvl(!ҽ����Ŀ����)
            
            If strҽ������ <> "" Then
                '��ȡ�Ը�����
                lng��ϸID = !ID
                ''�¶� 20041228
                'dbl�Ը����� = Get�Ը�����(IIf(InStr(1, "5,6,7", !�շ����) > 0, 1, 2), strҽ������, "21")
                dbl�Ը����� = Get�Ը�����(IIf(rs��ϸ!���� = "ҩƷ", 1, 2), strҽ������, IC_Data_����.mstrҵ������)
                
                '����ȡҽԺ��Ŀ�Ĺ�񼰵�λ
                gstrSQL = "Select ����,���,���㵥λ From �շ�ϸĿ Where ID=[1]"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "����ȡҽԺ��Ŀ�Ĺ�񼰵�λ", CLng(!�շ�ϸĿID))
                strҽ������ = ToVarchar(Nvl(rsItem!����), 40)
                strҽ����� = ToVarchar(Nvl(rsItem!���), 30)
                strҽ����λ = ToVarchar(Nvl(rsItem!���㵥λ), 8)
                
            '    ϵͳ��ʶ: hi_zymx_temp
            '    ���    �ֶα�ʶ    �ֶ�����    ����    ����    С��    �����  ȱʡֵ  ��ע
            '    YYBH    ҽԺ���    VARCHAR2    6       N
            '    JYH     ���׺�  NUMBER  20      N       ҽԺ��дסԺ�Ǽǽ��׺ţ�����ʱ��ʽ�����Զ���Ϊ���㽻�׺�
            '    JYH2    �ύ���׺�  NUMBER  20      N   0   ҽԺ������д���ύʱ�Զ���д
            '    MXXH    ��ϸ���    NUMBER  20      N       ÿ��סԺ�ŷ�����ϸ��Ŷ����ظ�
            '    JZLX    ��������    CHAR    2       N       11���21סԺ
            '    YNBH    Ժ�ڱ��    VARCHAR2    12      N
            '    TYBZ    ��ҩ��־    CHAR    1       N       �˷�ʱΪ1��������Ϊ����
            '    GRNM    �����籣���    VARCHAR2    18      N
            '    LBBZ    ����־    CHAR    1       N       1:ҩƷ  2:����
            '    XMBH    ��Ŀ���    VARCHAR2    10      N
            '    XMMC    ��Ŀ����    VARCHAR2    40      N
            '    XMGG    ��Ŀ���    VARCHAR2    30
            '    XMDW    ��Ŀ��λ    VARCHAR2    8
            '    YZRQ    ҽ������    DATE
            '    SSXM    ҽ������    VARCHAR2    20
            '    XMDJ    ��Ŀ����    NUMBER  12  2   N   0
            '    XMSL    ��Ŀ����    NUMBER  12  4   N   0   �˷�ʱΪ����
            '    XMTS    ��Ŀ����    NUMBER  6   2   N   0   ʼ��Ϊ1
            '    XMJE    ��Ŀ���    NUMBER  10  4   N   0
            '    ZFBL    �Ը�����    NUMBER  5   4   N   0
            '    ZFJE    �Ը����    NUMBER  12  2   N   0   û�����壬���ܽ���ʱ�Ż��
                
                gstrSQL = "Insert Into Hi_zymx_temp(" & _
                          "YYBH,JYH,MXXH,JZLX,YNBH,TYBZ,GRNM,LBBZ,XMBH,XMMC," & _
                          "XMGG,XMDW,YZRQ,SSXM,XMDJ,XMSL,XMTS,XMJE,ZFBL,ZFJE)" & _
                          "Values (" & _
                          "'" & IC_Data_����.mstrҽԺ���� & "'," & strסԺ��ˮ�� & "," & lng��ϸID & "," & _
                          "'21','" & strסԺ�� & "'," & IIf(!���� < 0, 1, 0) & ",'" & strҽ��֤�� & "'," & _
                          "" & IIf(rs��ϸ!���� = "ҩƷ", 1, 2) & ",'" & strҽ������ & "'," & _
                          "'" & strҽ������ & "','" & strҽ����� & "','" & strҽ����λ & "'," & _
                          "To_Date('" & Format(!�Ǽ�ʱ��, "yyyy-MM-dd") & "','yyyy-MM-dd')," & _
                          "'" & Nvl(!ҽ��) & "'," & Format(!�۸�, "#0.00") & "," & Format(!����, "#0.0000") & ",1," & _
                          "" & Format(!���, "#0.00") & "," & dbl�Ը����� & "," & Round(Format(!���, "#0.00") * dbl�Ը�����, 2) & ")"
                gcn����.Execute gstrSQL
                
                'ȡ��ˮ��
                Call Interface_Prepare_����(GetSequence, "~~~~" & ChargeDetail, "")
                intReturn = Interface_Exec_����
                If intReturn <> 0 Then
                    MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
                    If (intReturn < 0) Then Exit Function
                End If
                str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
                
                '�ϴ���ϸ
                '���
                '��~��~��~�� + סԺ�Ÿ��� + סԺ���б���%%�ָ���
'                    ���ز���(����"���׽��~������Ϣ~���׽����Ϣ")
'                    ���׽��~������Ϣ+��~��~�� + �޷������סԺ���б� + �޷������ҽ����ˮ���б����ֶβ��ã�Ϊ�գ� + �޷�����ķ�����ˮ���б��б��%%�ָ���
'                    �޷������סԺ���б� = סԺ�Ÿ���%%[סԺ��]
'                    �޷�����ķ�����ˮ���б�= ���ܱ����ԭ��1�����ظ���2�Ը���������%%��¼����%%[��ϸ���%%��ȷ���Ը������������ܱ����ԭ��Ϊ1�����ظ��������Ϊ�գ�]
'                    ע: סԺ���ص��޷�������ϸ�б�����ﷵ�صĲ��ɱ�ԭ���б�ͬ?
'                    ��������ֵ
'                    0�������óɹ����ұ��η���ȫ��ͨ��У���ύ
'                    -1��������ʧ��
'                    -2������һ��סԺ�����ڲ���סԺ�ж�У��ʧ�ܣ���ʱֻ��ȡ�޷������סԺ���б��޷�����ķ�����ϸ�б���Ϊ�գ������ڲ���Ϊ����~��~��+�޷������סԺ���б�+�ա�
'                    -3������һ��������ϸ���ܱ��棬��ʱֻҪ��ȡ�޷�����ķ�����ϸ�б��޷������סԺ���б���Ϊ�գ������ڲ���Ϊ����~��~��+��+�޷�����ķ�����ϸ�б�
                StrInput = "~~~~1~" & strסԺ��
                Call Interface_Prepare_����(ChargeDetail, StrInput, strOutput, str������ˮ��)
                intReturn = Interface_Exec_����()
                If intReturn <> 0 Then
                    Select Case intReturn
                    Case -1
                        strErrInfo = "��������ʧ�ܣ�"
                    Case -2
                        strErrInfo = "���ڵ�ǰ���˲���סԺ�У�У��ʧ�ܣ�"
                    Case -3
                        If Val(Split(Split(mstrMsg, strField)(7), "%%")(0)) = 1 Then
                            strErrInfo = "�����ظ���"
                        Else
                            strErrInfo = "�Ը���������"
                        End If
                    End Select
                    MsgBox strErrInfo, vbInformation, gstrSysName
                    If (intReturn < 0) Then Exit Function
                End If
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
                gcn�ϴ�.Execute gstrSQL, , adCmdStoredProc
            End If
            .MoveNext
        Loop
    End With
    
    'ͳ��ҽ����ϸ������ҽ����Ŀ�ܶ�
    With rsExse
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl�����ܶ� = dbl�����ܶ� + Nvl(!���, 0)
            If Nvl(!ҽ����Ŀ����) = "" Then
                dbl��ҽ����� = dbl��ҽ����� + Nvl(!���, 0)
            Else
                int��¼�� = int��¼�� + 1
            End If
            .MoveNext
        Loop
    End With
    
    '׼������Ԥ����
    str�޿����� = IS�޿�����(lng����ID)
    '�Ƿ���ҽ���� + IC��Ϣ + ��~��  +סԺ��+���ν�����ϸ����+ ��ҽ����Ŀ�ܶ�
    StrInput = Split(str�޿�����, "|")(0) & strField & IIf(Split(str�޿�����, "|")(0) = 1, IC_Data_����.IC������, Split(str�޿�����, "|")(1)) & strField & strField & strField & strסԺ�� & strField & int��¼�� & strField & dbl��ҽ����� & strField & IC_Data_����.mstrҵ������
    IC_Data_����.mstr������ڲ����� = int��¼�� & strField & dbl��ҽ�����
    Call Interface_Prepare_����(PreSettle, StrInput, strOutput)
    intReturn = Interface_Exec_����()
    If intReturn < 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ��ش�����Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ֽ������Ϣ
    '���׽��~������Ϣ+��~��~�� +
    '�������������ܶ�%%�Է��ܶ�%%�����ܶ�%%ͳ�����֧��%%�����ʻ�֧��%%�����ʻ�֧��
    '%%��֧��%%����Ա����֧��%%��λ֧��%%�����ֽ�֧��%%�𸶱�׼%%�ֶ���Ϣ�� + IC��д��ǰ����Ϣ(His��ʹ��)
    strBalance = Split(mstrMsg, strField)(5)
    dbl�����ܶ�_YB = Val(Split(strBalance, "%%")(int�����ܶ�))
    dblͳ����� = Val(Split(strBalance, "%%")(intͳ�����))
    dbl�ʻ�֧�� = Val(Split(strBalance, "%%")(int�����ʻ�)) + Val(Split(strBalance, "%%")(int�����ʻ�))
    dbl�󲡲��� = Val(Split(strBalance, "%%")(int�󲡾���))
    dbl����Ա���� = Val(Split(strBalance, "%%")(int����Ա����))
    IC_Data_����.mdbl��ҽ����� = dbl��ҽ�����
    
    If Format(dbl�����ܶ� - dbl��ҽ�����, "#0.00") <> Format(dbl�����ܶ�_YB, "#0.00") Then
        MsgBox "HIS�����ܶ����ҽ�������ܶ" & vbCrLf & _
        "ҽԺ��" & Format(dbl�����ܶ� - dbl��ҽ�����, "#0.00") & Space(10) & "ҽ����" & Format(dbl�����ܶ�_YB, "#0.00"), vbInformation, gstrSysName
    End If
    
    סԺ�������_���� = "�����ʻ�;" & dbl�ʻ�֧�� & ";0"
    סԺ�������_���� = סԺ�������_���� & "|ͳ�����;" & dblͳ����� & ";0"
    סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & dbl����Ա���� & ";0"
    סԺ�������_���� = סԺ�������_���� & "|�󲡾���;" & dbl�󲡲��� & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim intReturn As Integer
    Dim lng��ҳID As Long
    Dim blnTrans As Boolean
    Dim str������ˮ�� As String, strסԺ�� As String, strBalance As String
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim str�޿����� As String
    Dim dbl�����ܶ� As Double, dbl�Է��ܶ� As Double, dbl�����ܶ� As Double, dblͳ����� As Double, dbl�ֽ�֧�� As Double
    Dim dbl�ʻ�֧�� As Double, dbl�󲡲��� As Double, dbl����Ա���� As Double, dbl��λ֧�� As Double, dbl���� As Double
    Dim dbl�����ʻ� As Double, dbl�����ʻ� As Double, dbl�����ʻ�_��� As Double, dbl�����ʻ�_��� As Double
    
    Const int�����ܶ� As Integer = 0
    Const intͳ����� As Integer = 3
    Const int�����ʻ� As Integer = 4
    Const int�����ʻ� As Integer = 5
    Const int�󲡾��� As Integer = 6
    Const int����Ա���� As Integer = 7
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    'If MsgBox("ҽ���ӿڲ�֧����;���㣨һ��סԺֻ�ܽ���һ�ν��㣩����ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    'ȡ��ҳID
    gstrSQL = "Select סԺ���� From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
    lng��ҳID = rsTemp!סԺ����
    strסԺ�� = GetסԺ��(lng����ID)
    
    'ȡ���
    gstrSQL = "Select �����ʻ����,�����ʻ���� From �����ʻ� Where ����=[2] And ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���", lng����ID, TYPE_����)
    dbl�����ʻ�_��� = Nvl(rsTemp!�����ʻ����, 0)
    dbl�����ʻ�_��� = Nvl(rsTemp!�����ʻ����, 0)
    
    'ȡ��ˮ��
    Call Interface_Prepare_����(GetSequence, "~~~~" & Settle, "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
        If (intReturn < 0) Then Exit Function
    End If
    str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    '����סԺ����
'    ��ڲ��� (Data)
'    �Ƿ���ҽ���� IC������(���Դ���) + �ֽ�֧����ʽ + �� + סԺ�� + ���ν�����ϸ���� + ��ҽ����Ŀ�ܶ� + ����Ա����
'    ���ز���(����"���׽��~������Ϣ~���׽����Ϣ")
'    ���׽��~������Ϣ+дҽ���������+ �������˻���� + д����IC������ + ������(�ο�סԺԤ����)
    str�޿����� = IS�޿�����(lng����ID)
    StrInput = Split(str�޿�����, "|")(0) & strField & IIf(Split(str�޿�����, "|")(0) = 1, IC_Data_����.IC������, Split(str�޿�����, "|")(1)) & strField & "1" & strField & strField & _
    strסԺ�� & strField & IC_Data_����.mstr������ڲ����� & strField & UserInfo.���� & strField & IC_Data_����.mstrҵ������
    Call Interface_Prepare_����(Settle, StrInput, strOutput, str������ˮ��)
    intReturn = Interface_Exec_����
    If intReturn < 0 Then
        Err.Raise 9000, gstrSysName, "[" & mstrFunc & "]�ӿڷ��ش�����Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1)
        Exit Function
    End If
    blnTrans = True
    
    '�ֽ������Ϣ
    '���׽��~������Ϣ+��~��~�� +
    '�������������ܶ�%%�Է��ܶ�%%�����ܶ�%%ͳ�����֧��%%�����ʻ�֧��%%�����ʻ�֧��
    '%%��֧��%%����Ա����֧��%%��λ֧��%%�����ֽ�֧��%%�𸶱�׼%%�ֶ���Ϣ�� + IC��д��ǰ����Ϣ(His��ʹ��)
    strBalance = Split(mstrMsg, strField)(5)
    dbl�����ܶ� = Val(Split(strBalance, "%%")(int�����ܶ�))
    dblͳ����� = Val(Split(strBalance, "%%")(intͳ�����))
    dbl�ʻ�֧�� = Val(Split(strBalance, "%%")(int�����ʻ�)) + Val(Split(strBalance, "%%")(int�����ʻ�))
    dbl�󲡲��� = Val(Split(strBalance, "%%")(int�󲡾���))
    dbl����Ա���� = Val(Split(strBalance, "%%")(int����Ա����))
    dbl�����ʻ� = Val(Split(strBalance, "%%")(int�����ʻ�))
    dbl�����ʻ� = Val(Split(strBalance, "%%")(int�����ʻ�))
    dbl�ֽ�֧�� = dbl�����ܶ� - dblͳ����� - dbl�ʻ�֧�� - dbl�󲡲��� - dbl����Ա���� + IC_Data_����.mdbl��ҽ�����
    
    'д���ս����¼
    '���Ը�=�󲡲���;�����Ը�=����Ա����
    strBalance = "'" & TruncZero(strBalance) & "'"
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(zlDatabase.Currentdate()) & ",0,0,0,0," & dbl�����ʻ�_��� & "," & dbl�����ʻ�_��� & "," & dbl�����ʻ� & "," & dbl�����ʻ� & "," & dbl�����ܶ� & "," & dbl�ֽ�֧�� & ",0," & _
        dblͳ����� & "," & dblͳ����� & "," & dbl�󲡲��� & "," & dbl����Ա���� & "," & dbl�ʻ�֧�� & ",'" & str������ˮ�� & "|" & IC_Data_����.������� & "|" & IC_Data_����.mstrҵ������ & "'," & lng��ҳID & ",NULL," & strBalance & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "д���ս����¼")
    
    '����ȷ��
    '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
    'ȱʡ�ֽ�֧��
    StrInput = "~~~~" & Settle & strField & str������ˮ�� & strField & "0" & strField & "HIS�ɹ���"
    Call Interface_Prepare_����(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_����
    If intReturn < 0 Then
        Err.Raise 9000, gstrSysName, "���棺���ν���ȷ��ʧ�ܣ����¼�±��ν�����ˮ�ţ���֪ͨϵͳ����Աʹ�ù��߰��ٴ�ȷ�ϸý���" & _
            vbCrLf & "������ˮ�ţ�" & str������ˮ��
    End If
    סԺ����_���� = True
    
    Call ��Ժ�Ǽ�_����(lng����ID, lng��ҳID)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then
        '����ȷ��
        '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
        'ȱʡ�ֽ�֧��
        StrInput = "~~~~" & Settle & strField & str������ˮ�� & strField & "-1" & strField & "ҽ���ɹ�����HISʧ�ܣ�"
        Call Interface_Prepare_����(Decide, StrInput, strOutput)
    End If
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    '������Ժ�Ǽǵ�ͬʱ��ҽ�����Զ������סԺ�����������ˣ����ӿڲ����κδ���
    MsgBox "ҽ�����˳�����Ժʱ���ò�����ҽ�����ĵĳ�Ժ���㵥ͬʱ���ϣ����ν������Ͻ�����HIS�˵ķ��ã�", vbInformation, gstrSysName
    If MsgBox("��ȷ��Ҫ���������𣿣��粻�����������ѯϵͳ����Ա��" & vbCrLf & "�����������̣��Ȱ�������Ժ�Ǽǣ��ٽ���סԺ�������ϣ�Ȼ���ٴν��㣬�ٴΰ����Ժ�Ǽǣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    סԺ�������_���� = True
End Function

Public Function סԺ��Ϣ�䶯_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '��Ժ��Ϣ�䶯�򼲲�ѡ����ô˽ӿ�
    Dim intReturn As Integer
    Dim StrInput As String, strOutput As String, strErrInfo As String
    
    Dim str������ˮ�� As String, strסԺ�� As String
    Dim str���� As String, str�������� As String, lng����ID As Long
    Dim str��Ժ��� As String, str��Ժ���� As String, strҽ�� As String, str������� As String, lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '��~��~��~��+ סԺ�� + �䶯ʱ��(yyyy.mm.dd) + ���˴���+���ҽ������+�������+������ţ�����������ţ�+��������+����ID(��ȷ��ʱ��0)"���м���~�ָ����������䶯���ݣ�δ�䶯�Ŀ���Ϊ�մ������磺����Ϊ"�ڿƣ�201��"���䶯��ϢΪ"~~~~�ڿ�~201"
    '���ز���(����"���׽��~������Ϣ~���׽����Ϣ")
    '���׽��~������Ϣ+��~��~��
    
    gstrSQL = "Select A.��ǰ����ID,B.����,A.��ǰ���� From ������Ϣ A,���ű� B" & _
        " Where A.����ID=[1] And A.��ǰ����ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Ƽ�����", lng����ID)
    lng����ID = Nvl(rsTemp!��ǰ����ID, 0)
    str�������� = Nvl(rsTemp!����)
    str���� = Nvl(rsTemp!��ǰ����)
    strסԺ�� = GetסԺ��(lng����ID)
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True)
    'û�г�Ժ���ʱ������Ժ���Ϊ׼
    If Trim(str��Ժ���) = "" Then str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True)
    gstrSQL = "Select סԺҽʦ,��Ժ���� From ������ҳ Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����������Ժ����", lng����ID, lng��ҳID)
    If Nvl(rsTemp!��Ժ����) <> "" Then
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy.MM.dd")
    End If
    strҽ�� = Nvl(rsTemp!סԺҽʦ)
    gstrSQL = "Select Nvl(����ID,0) ����ID From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
    lng����ID = rsTemp!����ID
'    If lng����ID <> 0 Then
'        '��ȡǰ�û��еļ�����Ϣ
'        gstrSQL = "Select JBBZDM From SIM_JBDA Where JBBM=" & lng����ID
'        If rsTemp.State = 1 Then rsTemp.Close
'        rsTemp.Open gstrSQL, gcn����
'        If rsTemp.RecordCount <> 0 Then str������� = rsTemp!JBBZDM
'    End If
    
    'ȡ��ˮ��
    Call Interface_Prepare_����(GetSequence, "~~~~" & ModifyPatient, "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    '����סԺ��Ϣ�䶯
    StrInput = "~~~~" & strסԺ�� & strField & Format(zlDatabase.Currentdate(), "yyyy.MM.dd") & strField & _
        str���� & strField & strҽ�� & strField & str��Ժ��� & strField & lng����ID & strField & str�������� & strField & lng����ID
    Call Interface_Prepare_����(ModifyPatient, StrInput, strOutput, str������ˮ��)
    intReturn = Interface_Exec_����()
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    
    סԺ��Ϣ�䶯_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim intReturn As Integer
    Dim lng����ID As Long
    Dim str������ˮ�� As String
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim strҽ�� As String, strסԺ�� As String, str��Ժ��� As String, str������� As String, str��Ժ���� As String
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '��ڲ��� (Data)
    '��~��~��~�� + סԺ�� + ��Ժ���ҽ��+��Ժ���˵��+��Ժȷ�Ｒ�����+��Ժ����(yyyy.mm.dd)
    '���ز���(����"���׽��~������Ϣ~���׽����Ϣ")
    '���׽��~������Ϣ+��~��~��
    
    '�����Ƚ��㣬���Ժ
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        'ȡ��ˮ��
        Call Interface_Prepare_����(GetSequence, "~~~~" & ��Ժ����, "")
        intReturn = Interface_Exec_����
        If intReturn <> 0 Then
            MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
            If (intReturn < 0) Then Exit Function
        End If
        str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
        
        '׼�����ó�Ժ������д�ӿ�
        strסԺ�� = GetסԺ��(lng����ID)
        str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True)
        gstrSQL = "Select סԺҽʦ,��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����������Ժ����", lng����ID, lng��ҳID)
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy.MM.dd")
        strҽ�� = Nvl(rsTemp!סԺҽʦ)
        gstrSQL = "Select Nvl(����ID,0) ����ID From �����ʻ� Where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
        lng����ID = rsTemp!����ID
         If lng����ID <> 0 Then
            '��ȡǰ�û��еļ�����Ϣ
            gstrSQL = "Select JBMC From SIM_JBDA Where JBBM=" & lng����ID
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn����
             If rsTemp.RecordCount <> 0 Then str������� = rsTemp!JBMC '��������
        End If
        
        StrInput = "~~~~" & strסԺ�� & strField & str������� & strField & lng����ID & strField & str��Ժ����
        Call Interface_Prepare_����(��Ժ����, StrInput, strOutput, str������ˮ��)
        If intReturn <> 0 Then
            MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
            If (intReturn < 0) Then Exit Function
        End If
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '��ҽ���ĳ�Ժ���㳷������ʵ�֣�ͬʱ������Ժ��ҽ���ָ�����Ժ״̬
    Dim intReturn As Integer
    Dim blnTrans As Boolean
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim str�޿����� As String
    Dim str������ˮ�� As String, str������ˮ�� As String, strסԺ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '��ڲ��� (Data)
    '�Ƿ���ҽ���� IC��Ϣ(���Դ���) + �� + �� + Ҫ���ϵ�סԺ���㽻�׺� + Ҫ���ϵ�סԺ��
    '˵����ֻ��"Ҫ���ϵ�סԺ���㽻�׺�"��Ӧ�Ľ��㵥����סԺ����"Ҫ���ϵ�סԺ��"ʱ�ſ����˷ѣ���������˷ѡ�
    '���ز���(����"���׽��~������Ϣ~���׽����Ϣ")
    '���׽��~������Ϣ+дҽ���������+ �� + д����IC������+�Ƿ�Ϊ�ظ��˷ѣ�0�����˷ѣ�1���㵥��ҽ���ӿ��Ѿ����˷ѹ���+ ���������μ�סԺԤ���㣩
    '��ֵΪ����ʾ�˷ѣ����ظ��˷����ò������ظ��˷�Ϊ1���ҷ��سɹ���
    
    'ȡ��ˮ��
    Call Interface_Prepare_����(GetSequence, "~~~~" & SettleDel, "")
    intReturn = Interface_Exec_����
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str������ˮ�� = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    strסԺ�� = GetסԺ��(lng����ID)
    '��ȡ������ˮ��
    gstrSQL = "Select ֧��˳��� From ���ս����¼" & _
        " Where ��¼ID=(Select Max(��¼ID) From ���ս����¼ Where ����=2 And ��ҳID=[2] And ����ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ˮ��", lng����ID, lng��ҳID)
    str������ˮ�� = Split(rsTemp!֧��˳���, "|")(0)
    
    str�޿����� = IS�޿�����(lng����ID)
'    ��ڲ��� (Data)
'    �Ƿ���ҽ���� IC��Ϣ(���Դ���) + �� + �� + Ҫ���ϵ�סԺ���㽻�׺� + Ҫ���ϵ�סԺ��
    StrInput = Split(str�޿�����, "|")(0) & strField & Split(str�޿�����, "|")(1) & strField & strField & strField & str������ˮ�� & strField & strסԺ��
    Call Interface_Prepare_����(SettleDel, StrInput, strOutput, str������ˮ��)
    intReturn = Interface_Exec_����()
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]�ӿڷ���" & IIf(intReturn > 0, "����", "����") & "��Ϣ��" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    blnTrans = True
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    '����ȷ��
    '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
    'ȱʡ�ֽ�֧��
    StrInput = "~~~~" & SettleDel & strField & str������ˮ�� & strField & "0" & strField & "HIS�ɹ���"
    Call Interface_Prepare_����(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_����
    If intReturn < 0 Then MsgBox "���棺���ν���ȷ��ʧ�ܣ����¼�±��ν�����ˮ�ţ���֪ͨϵͳ����Աʹ�ù��߰��ٴ�ȷ�ϸý���" & _
        vbCrLf & "������ˮ�ţ�" & str������ˮ��, vbInformation, gstrSysName
    
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then
        '����ȷ��
        '��~��~��~��~��������~ҽ��������ˮ��~HIS������~������Ϣ
        'ȱʡ�ֽ�֧��
        StrInput = "~~~~" & SettleDel & strField & str������ˮ�� & strField & "-1" & strField & "ҽ���ɹ�����HISʧ�ܣ�"
        Call Interface_Prepare_����(Decide, StrInput, strOutput)
    End If
End Function

Public Function �������_����(ByVal lng����ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Nvl(�ʻ����,0) From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ����", lng����ID, TYPE_����)
    �������_���� = rsTemp.Fields(0).Value
End Function

Public Function IS�޿�����(ByVal lng����ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '�Ƿ��޿����ˣ��޿�����0���п�����1�����ظ�ʽ������־|ҽ����
    gstrSQL = "Select Nvl(�Ҷȼ�,0) AS ����־,ҽ���� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ��޿�����", lng����ID, TYPE_����)
    If rsTemp!����־ = 0 Then
        '�п����ˣ�ҽ���ŷ��ؿ�
        IS�޿����� = "1|"
    Else
        IS�޿����� = "0|" & rsTemp!ҽ����
    End If
End Function


