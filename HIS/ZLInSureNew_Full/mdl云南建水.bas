Attribute VB_Name = "mdl���Ͻ�ˮ"
Option Explicit
'����������ҽ�����ڲ��������
Private mstr˳��� As String        '���˳���,����������,סԺ����ڱ����ʻ���
Private mstrҽ���� As String        '���ҽ����,����������
Private mcur�ʻ���� As Double      '��Ÿ����ʻ����,���Ҫ��,����������(�����֤����)

Private mlng����ID As Long          '��Ų���ID����������������
Private mstr��������� As String    '���������ƺţ������ڴ������������ϸ����
Private mstr��Ա״̬ As String
Private mstr֧����� As String

Private mstrErr As String * 4

'###ҽ���ӿں���ԭ�ͣ���Ҫ��дΪAPI��ʽ
'���¼�����ע�⣺
'��1���ַ����������۴��뻹�Ǵ�����������ByVal�ؼ��֣�
'��2���������ַ��������ڵ���ǰ�����ʼ����
'��3����ֵ�������ڴ���������Ҫ����ByVal�ؼ��ֵģ���������һ�����ܼ�
'��4�����ڸ����������Ӧ������Double
'��5��ǧ�����ṹ����

'====================================================================================
'1 ������ϸ����
'���룺˳��ţ�����ǼǺţ����������š��շѴ�����롢�շ���Ŀ���롢��Ŀ���ơ��������۸񣨵��ۣ������ء�����÷������������ˡ��������ơ�������ƺš�ҽ��������
'������Ը��������Ը�����������������룻

Private Declare Sub yh_feedetailtrans Lib "Hisint" Alias "int_feedetailtrans" _
    (ByVal Serial_No As String, ByVal data_number As String, ByVal Charge_Category As String, _
    ByVal Charge_Item As String, ByVal Charge_Name As String, ByVal Count As Double, ByVal Price As Double, ByVal Pr_Area As String, _
    ByVal Standard As String, ByVal Usage_Dosage As String, ByVal Arranger As String, ByVal Section_Name As String, ByVal Transaction_No As String, _
    ByVal Doctor_Name As String, ByVal strDosage As String, ByVal strUnits As String, ByVal strChargeTime As String, _
    ByRef Pay_Proportion As Double, ByRef Pay_amount As Double, ByRef Wipe_Amount As Double, ByRef self_Amount As Double, ByRef error_code As String)

'2 ���ý���
'������

'3��������ϸ���ģ���ע������������˷Ѳ�����
'������

'4 ��Ժ�Ǽ�
'���룺���������͡�ҽԺ���롢�����ˡ��������ơ������š�סԺ�š��Ƿ����ֲ������ֲ����롢��Ժʱ�䡢��Ժ��ϡ�������ƺţ�
'�����˳��š����š����˱��롢�������Ա𡢳������ڡ����֤�š���ʼ���������ơ�������룻
'ע�����ֲ��������Ϊ��
Private Declare Sub yh_admit Lib "Hisint" Alias "int_admit" _
    (ByVal card_mode As String, ByVal Hospial_No As String, ByVal Arranger As String, ByVal Section_Name As String, _
    ByVal anamnesis_No As String, ByVal Admit_No As String, ByVal Ifspecialsick As String, ByVal specialsick_no As String, _
    ByVal admit_time As String, ByVal admit_diagnose As String, ByVal Transaction_No As String, ByRef Serial_No As String, ByRef CARD_NO As String, _
    ByRef Personal_No As String, ByRef Name As String, ByRef Sex As String, ByRef birthdate As String, ByRef strryzt As String, _
    ByRef Identify As String, ByRef initinstitution As String, ByRef strzflb As String, ByRef strDeptNO As String, ByRef strDeptName As String, ByRef error_code As String)
    
'5 IC��֧��
'���룺���������͡�˳��ţ�����ǼǺţ��������ˡ�֧��ԭ��,֧����
'�������ʼ���������ơ�������룻
Private Declare Sub yh_cardpay Lib "Hisint" Alias "int_cardpay" _
    (ByVal card_mode As String, ByVal Serial_No As String, ByVal Arranger As String, ByVal Pay_reason As String, ByVal Pay_amount As Double, _
     ByRef initinstitution As String, ByRef error_code As String)

'6 �������
'������

'7 �������ʶ��
'���룺���������͡�ҽԺ���롢�����ˡ��������ơ������š�����š���ҽʱ�䣻
'�����˳��š����š����˱��롢�������Ա𡢳������ڡ����֤�š���ʼ���������ơ�����������룻
Private Declare Sub yh_outpatientidentify Lib "Hisint" Alias "int_outpatientidentify" _
    (ByVal card_mode As String, ByVal Hospital_No As String, ByVal Arranger As String, ByVal Section_No As String, _
    ByVal anamnesis_No As String, ByVal outpatient_No As String, ByVal hospitalize_time As String, ByRef Serial_No As String, _
    ByRef CARD_NO As String, ByRef Personal_No As String, ByRef Name As String, ByRef Sex As String, ByRef birthdate As String, _
    ByRef Identify As String, ByRef initinstitution As String, ByRef accountremain As Double, ByRef strryzt As String, ByRef strzflb As String, ByRef error_code As String)

'8 IC��������Ϣ��ѯ
'���룺���������ͣ�
'���: �����š�ҽ���š��������Ա����֤�š����䡢�������
Private Declare Sub yh_cardinfo Lib "Hisint" Alias "int_cardinfo" _
    (ByVal Code_Mode As String, ByRef Amount As Double, ByRef CARD_NO As String, ByRef Personal_No As String, _
    ByRef Name As String, ByRef Sex As String, ByRef Identify As String, ByRef age As Double, ByRef error_code As String)

'9 �������
'����: ����������
'���: �������
Private Declare Sub yh_changepassword Lib "Hisint" Alias "int_changepassword" _
    (ByVal Code_Mode As String, ByRef error_code As String)

'10    �����ʻ�֧����ѯ
'���룺˳��ţ�
'�������֧���ܶ�������
Private Declare Sub yh_accountpay Lib "Hisint" Alias "int_accountpay" _
    (ByVal Serial_No As String, Amount As Double, ByVal error_code As String)

'11    �����ʻ�֧��
'���룺���������͡�ҽԺ���롢�������ơ������ˡ�֧��ԭ�򡢷����ܶ�ʻ�֧���
'�������ʼ���������ơ�˳��š�������룻
Private Declare Sub yh_outpay Lib "Hisint" Alias "outpay" _
    (ByVal card_mode As String, ByVal Hospital_No As String, ByVal Section_No As String, ByVal Arranger As String, ByVal payreason As String, _
    ByVal Amount As Double, ByVal accountpay As Double, ByVal initinstitution As String, ByVal Serial_No As String, ByVal error_code As String)

'12    ��ʼ��
'����: ��
'���: �������
Private Declare Sub yh_init Lib "Hisint" Alias "init" _
    (ByRef Errcode As String)

'13    �Ͽ�����
'���룺��
'���: ��
'Public Declare Sub yh_quit Lib "Hisint" Alias "quit" ()    '������ҽ�����Ѿ�����

'14 IC��Ȧ��
'���룺��
'���: �������
Private Declare Sub yh_loadcard Lib "Hisint" Alias "int_loadcard" (ByRef error_code As String)
    
'15 ���ݴ���
'���룺��
'���: �������
Private Declare Sub yh_datatrans Lib "Hisint" Alias "int_datatrans" (ByRef error_code As String)

'��Ժ�Ǽ��޸ģ����к�,��������,��������־0-��;1-��,��Ժʱ��yyyy-MM-dd hh24:mi:ss,��Ժ��ϣ�
Public Declare Sub yh_Recedeadmit Lib "hisint.dll" Alias "int_recedeadmit" (ByVal Serial_No As String, _
ByVal Section_Name As String, ByVal Ifspecialsick As String, ByVal specialsick_no As String, ByVal admit_time As String, ByVal admit_diagnose As String, _
ByRef error_code As String)
'specialsick_no��ֵ���£�
'0001   ��������
'0002   ������˥��
'0003   ������ֲ����������
'0004   ϵͳ�Ժ���Ǵ�
'0005   �����ϰ���ƶѪ
'0006   ����
'0007   ���
'0008   ������(����ɭ�ϲ�)
'0009   ���Ĳ�
'0010   ֧��������
'0011   ���Ĳ�
'0012   ����˥��
'0013   ��Ѫ������
'0014   ����
'0015   ��Ӳ��
'0016   ������ǰ��������II,III
'0017   ������С������
'0018   ��˲�
'0019   ���Ի�Ը���
'0020   ԭ����̷��Ը�ѪѹII~III��
'0021   ���������ʪ�ؽ���
'0022   ��״�ٻ��ܿ���(����)

'16 �������
'���룺������𣬾���˳��ţ�������ƺţ�����������ͣ�
'���: �������
Private Declare Sub yh_transaction Lib "Hisint" Alias "int_transaction" _
    (ByVal Trade_Sort As String, ByVal Serial_No As String, ByVal Transaction_No As String, ByVal Affirm_Mode As String, ByRef error_code As String)

'17 ��ȡ������ƺ�
'���룺�ޣ�
'���: ������ƺ�
Private Declare Sub yh_gettranssequence Lib "Hisint" Alias "int_gettranssequence" (ByRef Transaction_No As String)

'18    ��������ֶη��ò�ѯ
'���������˳��ţ�
'����������ֶα�׼���ֶ���š��ҹ��Ը���ͳ��֧����ͳ���Ը��������Ը�������Ը����ͳ��֧������Ը���ר�����֧���������룻
Private Declare Sub yh_SubsecFee Lib "Hisint" Alias "int_SubsecFee" _
    (ByVal Serial_No As String, ByVal Standard_Subsec As String, ByVal Subsec_No As String, _
      Selfpay As Double, Hookpay As Double, Tcpay As Double, Tcselfpay As Double, _
      Basepay As Double, outpay As Double, Preqpay As Double, Preqselfpay As Double, _
      SubsidyPay As Double, ByVal error_code As String)

'19 �˷Ѵ���
'���������˳��ţ����˱�־�������ţ�������ƺţ�
'�������: ������
Private Declare Sub yh_recedefeebalance Lib "Hisint" Alias "int_recedefeebalance" _
    (ByVal Serial_No As String, ByVal return_flag As String, ByVal balance_no As String, ByVal Transaction_No As String, _
        ByVal error_code As String)

'ɾ������δִ�н����Ԥ����ǰ�ķ�����ϸ���������ֻ������������㣬�Իᱻɾ��
Private Declare Sub yh_rollbackdetail Lib "Hisint" Alias "int_rollbackdetail" _
    (ByVal Serial_No As String, ByVal error_code As String)

'��ѯĳ�ν������ͳ���ۼ�,����ͳ��֧���޶��ͳ��֧���޶����Ϣ
'���������˳��ţ�
'�������: ���ߣ�ͳ���ۼƣ�����ͳ��֧���޶��ͳ��֧���޶�������
Private Declare Sub yh_RyspInfo Lib "Hisint" Alias "int_RyspInfo" _
   (ByVal series_no As String, qfx As Double, tclj As Double, dczfxe As Double, _
    dbxe As Double, ByVal error_code As String)


'======================================nt==============================================
'����ҽ����2.0�汾������ֻ����������ҽ����ͬ�ĺ���
'2 ���ý���
'���룺˳��ţ�����ǼǺţ��������ˡ��������ơ�������ƺţ�
'�����ȫ�Ը����ҹ��Ը���ͳ��֧����ͳ���Ը��������Ը�������Ը����ͳ��֧������Ը���
'       ҽ���չ���Ա���ԷѲ��֡�ҽ���չ���Ա��ͳ�ﲿ�֡���ʼ���������ơ�������룻
Private Declare Sub yh2_feebalance Lib "Hisint" Alias "int_feebalance" _
    (ByVal Serial_No As String, ByVal Arranger As String, ByVal Section_Name As String, ByVal Transaction_No As String, _
    ByRef Selfpay As Double, ByRef Hookpay As Double, ByRef Tcpay As Double, ByRef Tcselfpay As Double, ByRef Basepay As Double, _
    ByRef outpay As Double, ByRef Preqpay As Double, ByRef Preqselfpay As Double, ByRef ActualselfPay As Double, ByRef SubsidyPay As Double, _
    ByRef initinstitution As String, ByRef error_code As String)
    
'3��������ϸ���ģ���ע������������˷Ѳ�����
'���룺˳��ţ�����ǼǺţ����������š�������롢�µ��������µļ۸�
'������Ը��������Ը�����������������룻
Private Declare Sub yh2_recedefeedetail Lib "Hisint" Alias "int_recedefeedetail" _
    (ByVal Serial_No As String, ByVal data_number As String, ByVal Charge_Category As String, ByVal Count As Double, ByVal Price As Double, _
    ByRef Pay_Proportion As Double, ByRef Pay_amount As Double, ByRef Wipe_Amount As Double, ByRef self_Amount As Double, ByRef error_code As String)

'6 �������
'���롢���������ʹ�ó��Ϻ�ʱ������ý�����ͬ��
'���룺˳��ţ�����ǼǺţ�
'�����ȫ�Ը����ҹ��Ը���ͳ��֧����ͳ���Ը��������Ը�������Ը����ͳ��֧������Ը���
'       ҽ���չ���Ա���ԷѲ��֡�ҽ���չ���Ա��ͳ�ﲿ�֡���ʼ���������ơ�������룻

Private Declare Sub yh2_virtualbalance Lib "Hisint" Alias "int_virtualbalance" _
    (ByVal Serial_No As String, _
    ByRef Selfpay As Double, ByRef Hookpay As Double, ByRef Tcpay As Double, ByRef Tcselfpay As Double, ByRef Basepay As Double, _
    ByRef outpay As Double, ByRef Preqpay As Double, ByRef Preqselfpay As Double, ByRef ActualselfPay As Double, ByRef SubsidyPay As Double, _
    ByRef initinstitution As String, ByRef error_code As String)
'====================================================================================

Public Function ҽ����ʼ��_���Ͻ�ˮ() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    On Error GoTo errHandle

    mstrErr = Space(4)
    Call yh_init(mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbExclamation, gstrSysName
    Else
        ҽ����ʼ��_���Ͻ�ˮ = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function ��ݱ�ʶ_���Ͻ�ˮ(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������strSelfNO-���˱�ţ�ˢ���õ���strSelfPwd-�������룻bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim str���� As String, STR���� As String, str�Ա� As String
    Dim str���֤�� As String, str�������� As String, lng���� As Double
    Dim str��ʼ������ As String
    
    Dim strArranger As String
    Dim strSection As String
    Dim strPatiNo As String
    
    Dim str������ As String, lng����ID As Long, str�������� As String
    Dim rsTemp As New ADODB.Recordset
    Dim dat��ǰ As Date
    Dim strIdentify As String, str���� As String
    
    On Error GoTo errHandle
    '��ʼ������ȫ�ֵı���
    mstrҽ���� = Space(20)
    mstr˳��� = Space(19)
    mcur�ʻ���� = 0
    
    str���� = Space(18)
    STR���� = Space(60)
    str�Ա� = Space(3)
    str���֤�� = Space(20)
    str�������� = Space(10)
    str��ʼ������ = Space(4)
    mstr��Ա״̬ = Space(3)
    mstr֧����� = Space(3)
    dat��ǰ = zlDatabase.Currentdate
    
    If frmIdentify����.GetIdentifyMode(TYPE_���Ͻ�ˮ, bytType, str������, lng����ID, str��������) = False Then
        Exit Function
    End If
    DoEvents
        
    '�������֤��
    '���صı��ν��׵�˳��ŷ���:mstr˳���,�ڽ���ʱʹ��
    '���ص��������mcur�ʻ�����У���ȡ���ʱʹ��
    
    '��ȡIC����Ϣ
    strArranger = LeftDB(UserInfo.����, 8)
    strSection = LeftDB(UserInfo.����, 24)
    strPatiNo = LeftDB(UserInfo.���, 12)
    
    Screen.MousePointer = vbHourglass
    mstrErr = Space(4)
    If bytType <> 0 Then
        Call yh_cardinfo(str������, mcur�ʻ����, str����, mstrҽ����, STR����, str�Ա�, str���֤��, lng����, mstrErr)
    Else
        Call yh_outpatientidentify(str������, Trim(gstrҽԺ����), strArranger, strSection, strPatiNo, _
            strPatiNo, Format(dat��ǰ, "yyyy-MM-dd"), mstr˳���, str����, _
            mstrҽ����, STR����, str�Ա�, str��������, str���֤��, str��ʼ������, mcur�ʻ����, mstr��Ա״̬, mstr֧�����, mstrErr)
    End If
    If mstrErr <> "0000" Then
        Screen.MousePointer = vbDefault
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr˳��� = TrimStr(mstr˳���)
    mstrҽ���� = TrimStr(mstrҽ����)
    str���� = TrimStr(str����)
    If bytType = 0 Then
        If mstr˳��� = "" Then
            MsgBox "δ�ܴ�ǰ�÷��������˳���,�����Ի��鿨��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mstrҽ���� = "" Then
        MsgBox "δ�ܴӿ��ж�ȡҽ����,�����Ի��鿨��", vbInformation, gstrSysName
        Exit Function
    End If
    If str���� = "" Then
        MsgBox "δ�ܴӿ��ж�ȡ����,�����Ի��鿨��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    mstrҽ���� = str������ & Left(mstrҽ����, 19)
    strIdentify = str���� & ";" & mstrҽ���� & ";;" & TrimStr(STR����) & ";" & TrimStr(str�Ա�) & ";" & TrimStr(str��������) & ";" & TrimStr(str���֤��) & ";"
    strIdentify = Replace(strIdentify, " ", "")
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    ';8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�;23�������� (1����������)
    If bytType = 0 Then
        '���������,�ҵ�ǰסԺ,���������˳��Ų��˳�
        gstrSQL = "Select Count(����ID) Records From �����ʻ� Where nvl(��ǰ״̬,0)=1 And ҽ����='" & mstrҽ���� & "' And ����=" & TYPE_���Ͻ�ˮ
        Call OpenRecordset(rsTemp, "�ж��Ƿ���Ժ")
        If rsTemp!Records <> 0 Then
            MsgBox "��ǰҽ�������Ѿ���Ժ,������������Ǽ�!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If bytType = 2 Then
        '������������סԺ���ǾͲ���ʹ���µ�˳��š���ʹ����ǰ��
        gstrSQL = "select ˳��� from �����ʻ� where ����=" & TYPE_���Ͻ�ˮ & " and ����='" & str���� & "'"
        Call OpenRecordset(rsTemp, "��ˮҽ��")
        
        If rsTemp.RecordCount > 0 Then
            mstr˳��� = IIf(IsNull(rsTemp("˳���")), mstr˳���, rsTemp("˳���"))
        End If
    End If
    
    If IsDate(str��������) = True Then
        lng���� = DateDiff("yyyy", CDate(str��������), dat��ǰ)
    End If
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";" & mstr˳���                '9.˳���
    str���� = str���� & ";"                             '10��Ա���
    str���� = str���� & ";" & mcur�ʻ����              '11�ʻ����
    str���� = str���� & ";0"                            '12��ǰ״̬
    str���� = str���� & ";" & IIf(lng����ID <> 0, lng����ID, "") '13����ID
    str���� = str���� & ";1"                            '14��ְ(1,2)
    str���� = str���� & ";"                             '15����֤��
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";" & mcur�ʻ����              '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    str���� = str���� & ";"                             '23�������� (1����������)
    
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID, TYPE_���Ͻ�ˮ)
    If lng����ID = 0 Then Exit Function 'δ������ȷ�ı����ʻ�
    
    mlng����ID = lng����ID '4����������Ԥ��
    '�õ�������ƺ�,����������������
    mstr��������� = Get�����(False)
    If mstr��������� = "" Then
        Exit Function
    End If
    
    '���±����ʻ��е������
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���Ͻ�ˮ & ",'�����','''" & mstr��������� & "''')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '���ظ�ʽ:�м���벡��ID
    ��ݱ�ʶ_���Ͻ�ˮ = strIdentify & ";" & lng����ID & str����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_���Ͻ�ˮ(strSelfNo As String, ByVal bytPlace As Byte) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��
'      ��ʾ����λ�ã�10-����,20-��Ժ,30-Ԥ��,40-����
'����: ���ظ����ʻ����Ľ��
    
    On Error GoTo errHandle
    
    If strSelfNo = mstrҽ���� And (bytPlace = 10 Or bytPlace = 20) Then
        'ֱ�������ϴ����ʶ��ʱ�õ������ݷ���
        �������_���Ͻ�ˮ = mcur�ʻ����
    Else
        '��IC���ϵ����
        Call Get�����(strSelfNo, �������_���Ͻ�ˮ)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_���Ͻ�ˮ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim str�������� As String, strTemp As String
    
    Dim cur�Ը����� As Double, cur�Ը���� As Double, cur������� As Double, curȫ�Է� As Double
    Dim strҽ�� As String, str���� As String, str��� As String, str���� As String, str���� As String, str��λ As String, str�������� As String
    Dim cur�������� As Currency, dbl��� As Double, dbl���� As Double
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    If rs��ϸ.EOF = True Then
        MsgBox "�����������ϸ�ٽ���ҽ��Ԥ�㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rs��ϸ("����ID") <> mlng����ID Then
        MsgBox "�ò���δͨ�������֤�����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ֻ�����������ʹ�ñ�����
    On Error GoTo errHandle
    
    'ɾ��ǰ�÷�����������δ����ϸ
    mstrErr = Space(4)
    Call yh_transaction("1", mstr˳���, mstr���������, "0", mstrErr)
            
    '������ϸ����
    strTemp = rs��ϸ("����ID") & "_" & Format(zlDatabase.Currentdate, "ddHHmmss")
    Do Until rs��ϸ.EOF
        'ȡҩƷ����
        str���� = ""
        If InStr(1, ",5,6,7,", "," & rs��ϸ!�շ���� & ",") <> 0 Then
            gstrSQL = "Select A.���� From ҩƷ���� A,ҩƷ��Ϣ B,ҩƷĿ¼ C" & _
                " Where A.����=B.���� And B.ҩ��ID=C.ҩ��ID And C.ҩƷID=" & rs��ϸ!�շ�ϸĿID
            Call OpenRecordset(rsTemp, "ȡҩƷ����")
            str���� = ToVarchar(Nvl(rsTemp!����), 50)
        End If
        
        gstrSQL = "select A.����,A.����,A.���,A.���㵥λ,B.��Ŀ����,B.��ע" & _
                    " ,Decode(Sign(Instr(A.���,'��')),0,A.���,Substr(A.���,1,Instr(A.���,'��')-1)) as ���" & _
                    " ,Decode(Sign(Instr(A.���,'��')),0,A.���,Substr(A.���,Instr(A.���,'��')+1)) as ����" & _
                    " from �շ�ϸĿ A,����֧����Ŀ B where A.ID=" & rs��ϸ("�շ�ϸĿID") & " and A.ID=B.�շ�ϸĿID and B.����=" & TYPE_���Ͻ�ˮ
        Call OpenRecordset(rsTemp, "����Ԥ��")
        If rsTemp.EOF = True Then
            MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
            Exit Function
        End If
        
        strҽ�� = LeftDB(UserInfo.����, 8)
        str��λ = ToVarchar(Nvl(rsTemp!���㵥λ), 30)
        str��� = LeftDB(IIf(IsNull(rsTemp("���")), "�޹��", rsTemp("���")), 30)
        str���� = LeftDB(IIf(IsNull(rsTemp("����")), "", rsTemp("����")), 30)
        str���� = LeftDB(UserInfo.����, 24)
        '���ܴ��ݸ�������0��Ŀ����Ϊ��ɾ���Ѿ��ϴ����������ķ��ü�¼
        dbl���� = Val(IIf(rs��ϸ("����") > 0, rs��ϸ("����"), 0))
        dbl��� = Val(IIf(rs��ϸ("����") > 0, rs��ϸ("����"), 0))
        str���� = Get�������(Nvl(rsTemp!��Ŀ����), TYPE_���Ͻ�ˮ)
        str�������� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        str�������� = ToVarchar(strTemp & "_" & rs��ϸ.AbsolutePosition, 18)
        
        mstrErr = Space(4)
        Call yh_feedetailtrans(mstr˳���, str��������, str����, rsTemp("��Ŀ����"), _
            rsTemp("����"), dbl����, dbl���, str����, str���, "", strҽ��, str����, mstr���������, strҽ��, _
            str����, str��λ, str��������, cur�Ը�����, cur�Ը����, cur�������, curȫ�Է�, mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
            'ҽ�����ݿ�ع�
            Call yh_transaction("1", mstr˳���, mstr���������, "0", mstrErr)
            Exit Function
        End If
        
        cur�������� = cur�������� + rs��ϸ("ʵ�ս��")
        rs��ϸ.MoveNext
    Loop
        
    '�������
    Dim str�����־ As String, cur�����Է� As Double, cur��� As Currency
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, str��ʼ������ As String
    Dim cur������Ա�Ը� As Double, cur������Աͳ�� As Double, cur����Աͳ�� As Double
    Dim str��������� As String
    
    '��������Ԥ��
    str��������� = Get�����(True, mlng����ID)
    If str��������� = "" Then
        Exit Function
    End If
    
    str��ʼ������ = Space(4)
    mstrErr = Space(4)
    Call yh2_virtualbalance(mstr˳���, curȫ�Ը�, cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, _
        cur�����Ը�, cur��ͳ��, cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, str��ʼ������, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
        Exit Function
    End If
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .�������ý�� = cur��������
    End With
    
    cur��� = �������_���Ͻ�ˮ(mstrҽ����, 10)
    If cur������Աͳ�� > 0 Then
        cur�����Է� = cur������Ա�Ը�
    Else
        cur�����Է� = curȫ�Ը� + cur�ҹ��Ը� + cur�����Ը� + curͳ���Ը� + cur���Ը� + cur�����Ը� - cur����Աͳ��
    End If
    cur��� = IIf(cur��� > cur�����Է�, cur�����Է�, cur���) 'ȡ���ߵ�Сֵ
        
    str���㷽ʽ = "�����ʻ�;" & cur��� & ";1" '�����޸�
    
    If curͳ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & curͳ��֧�� & ";0"
    End If
    If cur��ͳ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|��ͳ��;" & cur��ͳ�� & ";0"
    End If
    If cur����Աͳ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & cur����Աͳ�� & ";0"
    End If
    If cur������Աͳ�� > 0 Then
        str���㷽ʽ = str���㷽ʽ & "|���ⲹ��;" & cur������Աͳ�� & ";0"
    End If
    
    �����������_���Ͻ�ˮ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_���Ͻ�ˮ(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset, lng����ID As Long
    Dim i As Long, curDate As Date, cur�������� As Currency, lng����ID As Long
    Dim str������ As String
    Dim str��������� As String   '������ƺ�
    Dim str��ʼ������ As String
    
    Dim cur�Ը����� As Double, cur�Ը���� As Double, cur������� As Double
    Dim strҽ�� As String, str���� As String
    Dim str��� As String, str���� As String
    
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double
    Dim cur������Աͳ�� As Double, cur������Ա�Ը� As Double, cur����Աͳ�� As Double
    
    
    On Error GoTo errHandle
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    gstrSQL = "Select A.ID,A.����ID,A.NO,A.�Ǽ�ʱ��,A.������ as ҽ��," & _
            "   A.����*A.���� as ����,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ��," & _
            "   A.�շ����,D.��Ŀ���� as �շ���Ŀ,B.���� as ��Ŀ����," & _
            "   decode(Instr(B.���,'��'),0,B.���,substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
            "   decode(Instr(B.���,'��'),0,'',substr(B.���,Instr(B.���,'��')+1)) as ����," & _
            "   C.���� as ��������" & _
            " From (Select * From ������ü�¼ Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=" & TYPE_���Ͻ�ˮ & _
            " Order by A.ID"
    Call OpenRecordset(rs��ϸ, "��ˮҽ��")
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If
    lng����ID = rs��ϸ("����ID")
    
    '�жϸò����Ƿ�������������
    gstrSQL = "select nvl(����ID,0) ����ID from �����ʻ� where ����ID=" & lng����ID & " and ����=" & TYPE_���Ͻ�ˮ
    Call OpenRecordset(rsTemp, "ҽ���ӿ�")
    If rsTemp.EOF = False Then
        '�����ⲡ�Ĳ�����ҪԤ��
        lng����ID = rsTemp("����ID")
    End If
    
    'һ��������ϸ����
    '˳��Ų��������֤ʱ���ص�ֵ:mstr˳���
    strҽ�� = LeftDB(IIf(IsNull(rs��ϸ("ҽ��")), UserInfo.����, rs��ϸ("ҽ��")), 8)
    str���� = LeftDB(IIf(IsNull(rs��ϸ("��������")), UserInfo.����, rs��ϸ("��������")), 24)
    cur�������� = g��������.�������ý�� '�ô���Ӧ�ս���Ԥ�㱣��һ��
        
    '����дIC��
    str������ = Left(strSelfNo, 1)
    str��ʼ������ = Space(4)
    mstrErr = Space(4)
    Call yh_cardpay(str������, mstr˳���, strҽ��, "�����շ�", CDbl(cur�����ʻ�), str��ʼ������, mstrErr)
    
    If mstrErr <> "0000" Then
        'ҽ�����ݿ�ع�
        Call yh_transaction("1", mstr˳���, mstr���������, "0", mstrErr)
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ)
        Exit Function
    End If
    
    '�������ý���
    str��������� = Get�����(True, lng����ID)
    If str��������� = "" Then
        Exit Function
    End If
    
    str��ʼ������ = Space(4)
    mstrErr = Space(4)
    Call yh2_feebalance(mstr˳���, strҽ��, str����, str���������, _
        curȫ�Ը�, cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, cur�����Ը�, cur��ͳ��, _
        cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, str��ʼ������, mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ)
        'ҽ�����ݿ�ع�
        Call yh_transaction("2", mstr˳���, str���������, "0", mstrErr)
        Exit Function
    End If
    Call yh_transaction("2", mstr˳���, str���������, "1", mstrErr)
    
    '�ġ���������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    '���� curͳ���ۼ� ������Ŀ����Ϊ�˵���API�����ͼ���
    Dim cur���� As Double, curͳ���ۼ� As Double, cur����ͳ���޶� As Double, cur���ͳ���޶� As Double
    Dim intסԺ�����ۼ� As Integer
    curDate = zlDatabase.Currentdate
            
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_���Ͻ�ˮ, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���Ͻ�ˮ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���Ͻ�ˮ & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�����Ը� & "," & Get���ֱ���(lng����ID) & "," & cur������Ա�Ը� & "," & _
        cur�������� & "," & curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ� & ",'" & mstr˳��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    '---------------------------------------------------------------------------------------------
    
    �������_���Ͻ�ˮ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_���Ͻ�ˮ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��

    ����������_���Ͻ�ˮ = False
End Function

Public Function �����ʻ�תԤ��_���Ͻ�ˮ(lngԤ��ID As Long, cur�����ʻ� As Currency, strSelfNo As String, str˳��� As String, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim str������ As String
    Dim str��ʼ������ As String
    Dim strҽ�� As String
    
    On Error GoTo errHandle
    
    If Is����ȷ(lng����ID) = False Then Exit Function
    
    str��ʼ������ = Space(4)
    str������ = "3"
    
    mstrErr = Space(4)
    strҽ�� = LeftDB(UserInfo.����, 8)
    Call yh_cardpay(str������, str˳���, CStr(UserInfo.����), "Ԥ����", cur�����ʻ�, str��ʼ������, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim rsTemp As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date
    
    '---------------------------------------------------------------------------------------------
    '��д�����
    curDate = zlDatabase.Currentdate
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_���Ͻ�ˮ, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���Ͻ�ˮ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(3," & lngԤ��ID & "," & TYPE_���Ͻ�ˮ & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & cur�����ʻ� & ",0,0,0,0,0,0," & _
        cur�����ʻ� & ",'" & str˳��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    �����ʻ�תԤ��_���Ͻ�ˮ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_���Ͻ�ˮ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false

    Dim rsTemp As New ADODB.Recordset
    Dim blnRollback As Boolean
    Dim str������ As String
    Dim str���� As String
    Dim STR���� As String
    Dim str�Ա� As String
    Dim str�������� As String
    Dim str���֤�� As String
    Dim str��ʼ������ As String
    Dim str����� As String   '������ƺ�
    Dim lng����ID As Long, str�������� As String, str��λ���� As String, str��λ���� As String
    
    On Error GoTo errHandle
    mstr˳��� = Space(19)
    strҽ���� = Space(20)
    str����� = Space(18)
    str���� = Space(18)
    STR���� = Space(60)
    str�Ա� = Space(3)
    str�������� = Space(10)
    str���֤�� = Space(20)
    str��ʼ������ = Space(4)
    mstr��Ա״̬ = Space(3)
    mstr֧����� = Space(3)
    str��λ���� = Space(10)
    str��λ���� = Space(56)
    
    gstrSQL = "Select A.��Ժ����,A.��Ժ����,B.���� as ��Ժ����,C.סԺ��,A.�Ǽ�ʱ��,D.ҽ����,E.���� as ���ֱ���,E.��� as ������� " & _
            " From ������ҳ A,���ű� B,������Ϣ C,�����ʻ� D,���ղ��� E " & _
            " Where A.��Ժ����ID=B.ID And A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
            " And A.����ID=C.����ID And A.����ID=D.����ID and D.����=" & TYPE_���Ͻ�ˮ & " and D.����ID=E.ID(+)"
    Call OpenRecordset(rsTemp, "��ˮҽ��")
    
    If rsTemp.EOF = True Then
        MsgBox "û�з��ִ˲��˵���Ϣ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    str�������� = Nvl(rsTemp!���ֱ���)
    str������ = Left(rsTemp("ҽ����"), 1)
    str����� = Get�����(True, lng����ID)
    If str����� = "" Then
        Exit Function
    End If
    
    mstrErr = Space(4)
    Call yh_admit(str������, gstrҽԺ����, LeftDB(UserInfo.����, 8), LeftDB(rsTemp("��Ժ����"), 24), _
        LeftDB(lng����ID, 12), LeftDB(rsTemp("סԺ��"), 12), IIf(str�������� = "", "0", "1"), LeftDB(str��������, 8), _
        Format(rsTemp!��Ժ����, "yyyy-MM-dd"), LeftDB(��ȡ���Ժ���(lng����ID, lng��ҳID, True, False), 50), str�����, mstr˳���, str����, _
        strҽ����, STR����, str�Ա�, str��������, mstr��Ա״̬, str���֤��, str��ʼ������, mstr֧�����, str��λ����, str��λ����, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
        'ҽ�����ݿ�ع�
        Call yh_transaction("0", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
    blnRollback = True
    
    '���±����ʻ��е������
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���Ͻ�ˮ & ",'����֤��','''" & str��λ���� & "|" & Trim(str��λ����) & "''')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    mstr˳��� = TrimStr(mstr˳���)
    If mstr˳��� = "" Then
        MsgBox "���ܵõ���ȷ����Ժ�Ǽ�˳��š�", vbInformation, gstrSysName
        Call yh_transaction("0", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
    strҽ���� = str������ & Left(TrimStr(strҽ����), 19)
    str���� = TrimStr(str����)
    
    'ǿ�ưѵǼ�˳��š����µ�ҽ��������
    gstrSQL = "ZL_�����ʻ�_�޸�ҽ����(" & lng����ID & "," & TYPE_���Ͻ�ˮ & _
                ",'" & str���� & "','" & strҽ���� & "','" & mstr˳��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���Ͻ�ˮ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    ��Ժ�Ǽ�_���Ͻ�ˮ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnRollback Then
        'ҽ�����ݿ�ع�
        Call yh_transaction("0", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
End Function

Public Function ��Ժ�Ǽǳ���_���Ͻ�ˮ(lng����ID As Long, lng��ҳID As Long, str˳��� As String) As Boolean
    Dim str����� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '������Ժ�Ǽ�,����š�˳����ɱ����ʻ�����ȡ
    gstrSQL = "Select �����,˳��� From �����ʻ� Where ����=" & TYPE_���Ͻ�ˮ & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����š�˳����ɱ����ʻ�����ȡ")
    str����� = Nvl(rsTemp!�����)
    str˳��� = Nvl(rsTemp!˳���)
    If str����� = "" Or str˳��� = "" Then
        MsgBox "����Ż�˳���Ϊ�գ��޷������Ժ�Ǽǳ������ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��Ժ
    mstrErr = Space(4)
    Call yh_transaction("0", str˳���, str�����, "0", mstrErr)
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���Ͻ�ˮ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    ��Ժ�Ǽǳ���_���Ͻ�ˮ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_���Ͻ�ˮ(lng����ID As Long, lng��ҳID As Long, str˳��� As String) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '��ҽ���ӿڲ����ڳ�Ժ�ӿ�,ֻҪ��������������ѳ�Ժ,����������Ҫ������Ժ��,��Ҫϵͳ����Ա����
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���Ͻ�ˮ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    ��Ժ�Ǽ�_���Ͻ�ˮ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_���Ͻ�ˮ(rsExse As Recordset, ByVal lng����ID As Long) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim str����� As String   '������ƺ�
    Dim cn�ϴ� As New ADODB.Connection, str�������� As String
    
    Dim cur�Ը����� As Double, cur�Ը���� As Double, cur������� As Double, curȫ�Է� As Double
    Dim strҽ�� As String, str���� As String, str��� As String, str���� As String, str���� As String, str��λ As String, str�������� As String
    Dim cur�������� As Currency, dbl��� As Double, dbl���� As Double
    Dim lng��ҳID As Long
    Dim str���� As String
    Dim strSickNO As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    With g��������
        .����ID = rsExse("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & rsExse("����ID")
        Call OpenRecordset(rsTemp, "�������")
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
        lng��ҳID = rsTemp!��ҳID
    End With
    '������һ�����Ӵ����Դﵽ���ܵ�ǰ��������Ŀ���
    Set cn�ϴ� = GetNewConnection
    
    '˳���ȡ��Ժ�Ǽ���֤���ص�
    gstrSQL = "Select ҽ����,˳��� From �����ʻ� " & _
              "Where ˳��� is Not NULL And ����ID=" & lng����ID & " And ����=" & TYPE_���Ͻ�ˮ
    Call OpenRecordset(rsTemp, "�������")
    
    If rsTemp.EOF Then
        MsgBox "δ���ָò��˵�סԺ����˳���,����ִ��ҽ�����ף�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mstr˳��� = rsTemp("˳���")
    mstrҽ���� = rsTemp!ҽ����
    str����� = Get�����(True, lng����ID)
    If str����� = "" Then
        Exit Function
    End If
    
    '�ȸ���������Ժ����
    '��Ժ�Ǽ��޸�yh_Recedeadmit�����к�,��������,��������־0-��;1-��,��Ժʱ��yyyy-MM-dd hh24:mi:ss,��Ժ��ϣ�
    gstrSQL = "Select A.��Ժ����,A.��Ժ����,B.���� as ��Ժ����,C.סԺ��,A.�Ǽ�ʱ��,D.ҽ����,E.���� as ���ֱ���,E.��� as ������� " & _
            " From ������ҳ A,���ű� B,������Ϣ C,�����ʻ� D,���ղ��� E " & _
            " Where A.��Ժ����ID=B.ID And A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
            " And A.����ID=C.����ID And A.����ID=D.����ID and D.����=" & TYPE_���Ͻ�ˮ & " and D.����ID=E.ID(+)"
    Call OpenRecordset(rsTemp, "��ˮҽ��")
    mstrErr = Space(4)
    strSickNO = Nvl(rsTemp!���ֱ���)
    Call yh_Recedeadmit(mstr˳���, LeftDB(rsTemp!��Ժ����, 24), IIf(strSickNO = "", "0", "1"), _
        strSickNO, Format(rsTemp!��Ժ����, "yyyy-MM-dd"), LeftDB(��ȡ���Ժ���(lng����ID, lng��ҳID, True, False), 50), mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
        Exit Function
    End If
    
    '�Ȼع�������ϸ
    Call yh_transaction("1", mstr˳���, str�����, "0", mstrErr)
    
    '������ϸ����
    '��ˮҽ��ֻ������δ�ϴ��ķ��ü�¼
    Do Until rsExse.EOF
        'ȡҩƷ����
        str���� = ""
        If InStr(1, ",5,6,7,", "," & rsExse!�շ���� & ",") <> 0 Then
            gstrSQL = "Select A.���� From ҩƷ���� A,ҩƷ��Ϣ B,ҩƷĿ¼ C" & _
                " Where A.����=B.���� And B.ҩ��ID=C.ҩ��ID And C.ҩƷID=" & rsExse!�շ�ϸĿID
            Call OpenRecordset(rsTemp, "ȡҩƷ����")
            str���� = ToVarchar(Nvl(rsTemp!����), 50)
        End If
        
        str��λ = ToVarchar(Nvl(rsExse!���㵥λ), 30)
        strҽ�� = LeftDB(IIf(IsNull(rsExse("ҽ��")), UserInfo.����, rsExse("ҽ��")), 8)
        str��� = LeftDB(IIf(IsNull(rsExse("���")), "�޹��", rsExse("���")), 30)
        str���� = LeftDB(IIf(IsNull(rsExse("����")), "", rsExse("����")), 30)
        str���� = LeftDB(IIf(IsNull(rsExse("��������")), UserInfo.����, rsExse("��������")), 24)
        dbl���� = CDbl(Nvl(rsExse("����"), 0))
        dbl��� = CDbl(Nvl(rsExse("�۸�"), 0))
        str���� = Get�������(Nvl(rsExse!ҽ����Ŀ����), TYPE_���Ͻ�ˮ)
        str�������� = Format(rsExse!����ʱ��, "yyyy-MM-dd HH:mm:ss")
        str�������� = rsExse("NO") & "_" & rsExse("���") & "_" & rsExse("��¼����") & "_" & rsExse("��¼״̬")
        
        mstrErr = Space(4)
        Call yh_feedetailtrans(mstr˳���, str��������, str����, rsExse("ҽ����Ŀ����"), _
            rsExse("�շ�����"), dbl����, dbl���, str����, str���, "", strҽ��, str����, str�����, strҽ��, _
            str����, str��λ, str��������, cur�Ը�����, cur�Ը����, cur�������, curȫ�Է�, mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
            'ҽ�����ݿ�ع�
            Call yh_transaction("1", mstr˳���, str�����, "0", mstrErr)
            Exit Function
        End If
        
        'Ϊ�������ü�¼�����ϴ���־���ϴ�һ������һ��
        gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
        cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
        
        cur�������� = cur�������� + rsExse("���")
        rsExse.MoveNext
    Loop
        
    '�������
    Dim str�����־ As String
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double, cur�ʻ�֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double, cur�ʻ���� As Currency
    Dim cur��ͳ�� As Double, cur���Ը� As Double, str��ʼ������ As String
    Dim cur������Ա�Ը� As Double, cur������Աͳ�� As Double, cur����Աͳ�� As Double
    
    str��ʼ������ = Space(4)
    mstrErr = Space(4)
    str�����־ = "0" '�������
    Call yh2_virtualbalance(mstr˳���, curȫ�Ը�, cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, _
        cur�����Ը�, cur��ͳ��, cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, str��ʼ������, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
        Exit Function
    End If
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .����ID = lng����ID
        .�������ý�� = cur��������
    End With
    
    סԺ�������_���Ͻ�ˮ = "ҽ������;" & curͳ��֧�� & ";0"
    If cur��ͳ�� <> 0 Then
        סԺ�������_���Ͻ�ˮ = סԺ�������_���Ͻ�ˮ & "|��ͳ��;" & cur��ͳ�� & ";0"
    End If
    If cur����Աͳ�� <> 0 Then
        סԺ�������_���Ͻ�ˮ = סԺ�������_���Ͻ�ˮ & "|����Ա����;" & cur����Աͳ�� & ";0"
    End If
    If cur������Աͳ�� > 0 Then
        סԺ�������_���Ͻ�ˮ = סԺ�������_���Ͻ�ˮ & "|���ⲹ��;" & cur������Աͳ�� & ";0"
    End If
    '�ֽ�֧�����ֿ���ȫ���ʻ�֧��
    cur�ʻ�֧�� = cur�������� - curͳ��֧�� - cur��ͳ�� - cur����Աͳ�� - cur������Աͳ��
    If cur�ʻ�֧�� > 0 Then
        Call Get�����(mstrҽ����, cur�ʻ����)
        cur�ʻ�֧�� = IIf(cur�ʻ�֧�� < cur�ʻ����, cur�ʻ�֧��, cur�ʻ����)
        סԺ�������_���Ͻ�ˮ = סԺ�������_���Ͻ�ˮ & "|�����ʻ�;" & cur�ʻ�֧�� & ";1"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_���Ͻ�ˮ(lng����ID As Long, str˳��� As String, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    Dim str����� As String   '������ƺ�
    Dim str������ As String
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double, cur�����ʻ� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, str��ʼ������ As String
    Dim cur������Ա�Ը� As Double, cur������Աͳ�� As Double, cur����Աͳ�� As Double
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date, lng����ID As Long, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    str������ = Left(mstrҽ����, 1)
    
    '��ȡ���θ����ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & TYPE_���Ͻ�ˮ & _
        " And A.���㷽ʽ in ('�����ʻ�') And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���θ����ʻ�֧����")
    cur�����ʻ� = 0
    If Not rsTemp.EOF Then
        cur�����ʻ� = rsTemp!�����ʻ�
    End If
    
    'ȡ��Ժ�Ǽ���֤�����ص�˳���
    mstr˳��� = str˳���
    str����� = Get�����(True, lng����ID)
    If str����� = "" Then
        Exit Function
    End If
    
    '����дIC��
    If CDbl(cur�����ʻ�) > 0 Then
        str��ʼ������ = Space(4)
        mstrErr = Space(4)
        Call yh_cardpay(str������, mstr˳���, UserInfo.����, "סԺ����", CDbl(cur�����ʻ�), str��ʼ������, mstrErr)
        If mstrErr <> "0000" Then
            'ҽ�����ݿ�ع�
            Call yh_transaction("2", mstr˳���, str�����, "0", mstrErr)
            Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ)
            Exit Function
        End If
    End If
    
    mstrErr = Space(4)
    str��ʼ������ = Space(4)
    Call yh2_feebalance(mstr˳���, LeftDB(UserInfo.����, 8), LeftDB(UserInfo.����, 24), str�����, curȫ�Ը�, _
        cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, cur�����Ը�, _
        cur��ͳ��, cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, _
        str��ʼ������, mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ)
        'ҽ�����ݿ�ع�
        Call yh_transaction("2", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
    
    '��д�����
    curDate = zlDatabase.Currentdate
    '�����ò��˱��ν���Ĳ�����Ϣ
    gstrSQL = "Select nvl(����ID,0) ����ID From �����ʻ� A Where A.����=" & TYPE_���Ͻ�ˮ & " and A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���ս���")
    If rsTemp.EOF = False Then
        lng����ID = rsTemp("����ID")
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_���Ͻ�ˮ, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '���� curͳ���ۼ� ������Ŀ����Ϊ�˵���API�����ͼ���
    Dim cur���� As Double, curͳ���ۼ� As Double, cur����ͳ���޶� As Double, cur���ͳ���޶� As Double
    curͳ�ﱨ���ۼ� = curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ��
            
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���Ͻ�ˮ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���Ͻ�ˮ & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�����Ը� & "," & Get���ֱ���(lng����ID) & "," & cur������Ա�Ը� & "," & _
        g��������.�������ý�� & "," & curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & ",0,'" & mstr˳��� & "'," & g��������.��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ˮҽ��")
    
    סԺ����_���Ͻ�ˮ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_���Ͻ�ˮ(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curDate As Date
    
    If TYPE_���Ͻ�ˮ = TYPE_���Ͻ�ˮ Then Exit Function '��ˮҽ����֧��
    
End Function

Private Function LeftDB(ByVal strText As String, ByVal lngLength As Long)
'���ܣ������ݿ�ĳ��ȼ��㷽ʽ�õ��ַ�����ʵ�ʿ����Ӵ�
    LeftDB = StrConv(MidB(StrConv(strText, vbFromUnicode), 1, lngLength), vbUnicode)
End Function

Private Function Get�����(Optional ByVal blnRead As Boolean = False, Optional ByVal lng����ID As Long = 0) As String
    '���������,��֤���������,ʹ��Ψһ�������
    Dim str����� As String
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    blnExist = False
    If blnRead Then
        '�ӱ����ʻ�����ȡ�����
        If lng����ID = 0 Then Exit Function
        gstrSQL = "Select ����� From �����ʻ� Where ����=" & TYPE_���Ͻ�ˮ & " And ����ID=" & lng����ID
        Call OpenRecordset(rsTemp, "�ӱ����ʻ�����ȡ�����")
        str����� = Nvl(rsTemp!�����)
    End If
    
    blnExist = (Trim(str�����) <> "")
    If Not blnExist Then
        str����� = Space(18)
        Call yh_gettranssequence(str�����) '������ô��ݺͽ��������������
        str����� = TrimStr(str�����)
        If str����� = "" Then
            MsgBox "��ȡ������ƺ�ʧ�ܡ�", vbInformation, gstrSysName
        End If
    End If
    
    '���ǵ���ǰ�����ݣ������ʻ�����������ֶΣ�������������ͽ��²����ı���
    If blnRead And Not blnExist Then
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���Ͻ�ˮ & ",'�����','''" & str����� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���������")
    End If
    
    Get����� = str�����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Is����ȷ(ByVal lng����ID As Long) As Boolean
'���ܣ��ж϶������Ŀ��Ƿ����Ҫ�����Ĳ��˵�
    Dim rsTemp As New ADODB.Recordset
    Dim str����_�� As String, str���� As String, strҽ���� As String, str������ As String
    
    Dim cur��� As Double, STR���� As String, str�Ա� As String
    Dim str���֤�� As String, lng���� As Double
    
    On Error GoTo errHandle
    
    gstrSQL = "select ����,ҽ���� from �����ʻ� where ����=" & TYPE_���Ͻ�ˮ & " and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ˮҽ��")
    
    str����_�� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    str������ = "3"
    
    str���� = Space(20)
    STR���� = Space(60)
    str�Ա� = Space(3)
    str���֤�� = Space(20)
    
    mstrErr = Space(4)
    Call yh_cardinfo(str������, cur���, str����, strҽ����, STR����, str�Ա�, str���֤��, lng����, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
        Exit Function
    End If
    str���� = TrimStr(str����)
    
    If str���� <> str����_�� Then
        MsgBox "ˢ�����еĿ����ǵ�ǰ���˵ģ��������ȷ��IC����", vbInformation, gstrSysName
        Exit Function
    End If
    
    Is����ȷ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get�����(ByVal strҽ���� As String, ����� As Currency) As Boolean
'���ܣ��õ������
    Dim cur��� As Double, STR���� As String, str�Ա� As String, str���� As String
    Dim str���֤�� As String, lng���� As Double, str������ As String
    
    str������ = "3"
    
    str���� = Space(20)
    STR���� = Space(60)
    str�Ա� = Space(3)
    str���֤�� = Space(20)
    
    mstrErr = Space(4)
    Call yh_cardinfo(str������, cur���, str����, strҽ����, STR����, str�Ա�, str���֤��, lng����, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, TYPE_���Ͻ�ˮ), vbInformation, gstrSysName
        Exit Function
    End If
    
    ����� = cur���
    Get����� = True
End Function

Private Function Get���ֱ���(ByVal lng����ID As Long) As String
'���ܣ��ж϶������Ŀ��Ƿ����Ҫ�����Ĳ��˵�
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "select ���� from ���ղ��� where ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ˮҽ��")
    
    If rsTemp.EOF = False Then
        Get���ֱ��� = Val(rsTemp("����")) 'Ϊ�˱����ڷⶥ���ֶΣ����Ա���������
        If Val(Get���ֱ���) = 0 Then Get���ֱ��� = "9000" '�������ֲ�ҲΪ0000������ǿ�Ƹ�Ϊ9000
    Else
        Get���ֱ��� = 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
