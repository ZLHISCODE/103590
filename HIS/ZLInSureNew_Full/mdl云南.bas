Attribute VB_Name = "mdl����"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

'����������ҽ�����ڲ��������
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ��
Private mstr˳��� As String        '���˳���,����������,סԺ����ڱ����ʻ���
Private mstrҽ���� As String        '���ҽ����,����������
Private mcur�ʻ���� As Double      '��Ÿ����ʻ����,���Ҫ��,����������(�����֤����)
Public mbln����Ա As Boolean       '��Ź���Ա��־
Private mlng����ID As Long          '��Ų���ID����������������
Private mstr��ϸ����� As String    '���������ƺţ������ڴ������������ϸ����

Private mstrAverageFeeType As String
Private mstrTsyybz As String    '���ղ����е�ƽ���������������ҽԺ��־��ÿ�γ�ʼ��ʱ����

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
    ByVal Doctor_Name As String, ByVal Charge_Time As String, Pay_Proportion As Double, Pay_amount As Double, Wipe_Amount As Double, ByVal error_code As String)

'2 ���ý���(ʡ��ҽ��������һ������)
'���룺˳��ţ�����ǼǺţ��������ˡ��������ơ�������ƺţ�
'�����ȫ�Ը����ҹ��Ը���ͳ��֧����ͳ���Ը��������Ը�������Ը����ͳ��֧������Ը���
'       ҽ���չ���Ա���ԷѲ��֡�ҽ���չ���Ա��ͳ�ﲿ�֡�����Աͳ��֧�����֡���Ա״̬����ʼ���������ơ�����ҹ�֧�����֡���������,Ѫ͸����,������룻
Private Declare Sub yh_feebalance Lib "Hisint" Alias "int_feebalance" _
    (ByVal Serial_No As String, ByVal Arranger As String, ByVal Section_Name As String, ByVal SickSortCode As String, ByVal Transaction_No As String, _
    Selfpay As Double, Hookpay As Double, Tcpay As Double, Tcselfpay As Double, Basepay As Double, _
    outpay As Double, Preqpay As Double, Preqselfpay As Double, ActualselfPay As Double, SubsidyPay As Double, _
    OfficialPay As Double, ByVal ryzt As String, ByVal initinstitution As String, tsggzfbf As Double, _
    ByVal feebalancetype As String, xtcs As Double, ByVal error_code As String)
    
'3��������ϸ���ģ���ע������������˷Ѳ�����
'���룺˳��ţ�����ǼǺţ����������š��µ��������µļ۸�������ƺţ�
'������Ը��������Ը�����������������룻
Private Declare Sub yh_recedefeedetail Lib "Hisint" Alias "int_recedefeedetail" _
    (ByVal Serial_No As String, ByVal data_number As String, ByVal Count As Double, ByVal Price As Double, _
     ByVal Transaction_No As String, Pay_Proportion As Double, Pay_amount As Double, Wipe_Amount As Double, ByVal error_code As String)

'4 ��Ժ�Ǽ�
'���룺���������͡�ҽ��������ҽԺ���롢�����ˡ��������ơ������š�סԺ�š��Ƿ����ֲ������ֲ����롢��Ժʱ�䡢��Ժ��ϡ�������ƺţ�
'�����˳��š����š����˱��롢�������Ա𡢳������ڡ����֤�š���ʼ���������ơ���λ���롢��λ���ơ�������룻
'ע�����ֲ��������Ϊ��
'2007-10-15�޸�:
'int_admit�����������������������Ա״̬��akc021����Ⱥ���akc300
'����α���Ա�������ⲡ�������ڽ������ⲡ���סԺҵ�����ʱ�����봫�룺���ֲ���־�����ֲ����룬���ֲ���־�����ֲ�����ְͬ������������
'��ϸ���ݼ�: ����˵��
'����������HIS�ڰ����˾���Ǽ�ʱ�������ӿڳ�����������ݿ��л�ȡ��Ա״̬��סԺ������סԺ�޶�Ƚ������ڷ��÷ָ���ý�������ݣ�ͬʱ�ӽӿڵõ������������Ա����������HIS������Ժ�Ǽ�ʱҪʹ�õ�IC���еĻ�����Ϣ��
'������������������͡�ҽԺ���롢�����ˡ��������ơ������š�סԺ�š��Ƿ����ֲ������ֲ����롢��Ժʱ�䡢��Ժ��ϡ�������ƺţ�
'���������˳��š����š����˱��롢�������Ա𡢳������ڡ����֤�š���Ա״̬����Ⱥ��𡢳�ʼ���������ơ�������룻
'2008-3-10�޸�:
'int_admit�����������������yck002��������Ⱥ��־ 0Ϊ��1Ϊ�ǣ�
'������������������͡�ҽԺ���롢�����ˡ��������ơ������š�סԺ�š��Ƿ����ֲ������ֲ����롢��Ժʱ�䡢��Ժ��ϡ�������ƺţ�
'���������˳��š����š����˱��롢�������Ա𡢳������ڡ����֤�š���Ա״̬����Ⱥ���������Ⱥ��־����ʼ���������ơ�������룻
Private Declare Sub yh_admit Lib "Hisint" Alias "int_admit" _
    (ByVal card_mode As String, ByVal doctorname As String, ByVal Hospital_No As String, ByVal Arranger As String, ByVal Section_Name As String, _
    ByVal anamnesis_No As String, ByVal Admit_No As String, ByVal Ifspecialsick As String, ByVal specialsick_no As String, _
    ByVal admit_time As String, ByVal admit_diagnose As String, ByVal Transaction_No As String, ByVal Serial_No As String, ByVal CARD_NO As String, _
    ByVal Personal_No As String, ByVal Name As String, ByVal Sex As String, ByVal birthdate As String, _
    ByVal Identify As String, ByVal akc021 As String, ByVal akc300 As String, ByVal yck002 As String, ByVal initinstitution As String, ByVal dwbm As String, ByVal dwmc As String, ByVal error_code As String)

'���ս�ת��Ժ
'2007-10-15�޸�:
'int_kndadmit()�����������������������Ա״̬��akc021����Ⱥ���akc300
'��ϸ���ݼ�: ����˵��
'������������ͨ������Ƚ����Ժ�Ĳ��ˣ�����Ȱ�����Ժʱʹ�á�
'���������ҽ�����������˱�ţ�ҽԺ���룬�����ˣ��������ƣ������ţ�סԺ�ţ��Ƿ����ֲ������ֲ����룬��Ժʱ�䣬��Ժ��ϣ�������ƺ�
'���������˳��ţ�IC���ţ��������Ա𣬳������ڣ���Ա״̬����Ⱥ��𣬳�ʼ���������룬��λ���룬�������
'2008-3-10�޸�:
'int_admit�����������������ykc002��������Ⱥ��־ 0Ϊ��1Ϊ�ǣ�
Private Declare Sub yh_kndadmit Lib "Hisint" Alias "int_kndadmit" _
    (ByVal doctorname As String, ByVal Personal_No As String, ByVal Hospital_No As String, ByVal Arranger As String, _
    ByVal Section_Name As String, ByVal anamnesis_No As String, ByVal Admit_No As String, ByVal Ifspecialsick As String, _
    ByVal specialsick_no As String, ByVal admit_time As String, ByVal admit_diagnose As String, ByVal Transaction_No As String, _
    ByVal Serial_No As String, ByVal CARD_NO As String, ByVal Name As String, ByVal Sex As String, ByVal birthdate As String, _
    ByVal akc021 As String, ByVal akc300 As String, ByVal ykc002 As String, ByVal initinstitution As String, ByVal dwbm As String, ByVal dwmc As String, ByVal error_code As String)

'5 IC��֧��
'���룺���������͡�˳��ţ�����ǼǺţ��������ˡ�֧��ԭ��,֧����
'�������ʼ���������ơ�������룻
Private Declare Sub yh_cardpay Lib "Hisint" Alias "int_cardpay" _
    (ByVal card_mode As String, ByVal Serial_No As String, ByVal Arranger As String, ByVal Pay_reason As String, ByVal Pay_amount As Double, _
     ByVal initinstitution As String, ByVal error_code As String)


'6 �������
'���롢���������ʹ�ó��Ϻ�ʱ������ý�����ͬ��
'���룺˳��ţ�����ǼǺţ���Ԥ�����־�������š�������ƺţ�
'�����ȫ�Ը����ҹ��Ը���ͳ��֧����ͳ���Ը��������Ը�������Ը����ͳ��֧������Ը���
'       ҽ���չ���Ա���ԷѲ��֡�ҽ���չ���Ա��ͳ�ﲿ�֡�����Աͳ��֧������Ա״̬����ʼ���������ơ�����ҹ�֧�����֡�������룻
'ע�⣺Ԥ�����־          0 ��ʾ������㣬��ҽ������û���κμ�¼��1  ��ʾԤ���㣬������Ϊ��;����ʹ��
'      ҽ���չ���Ա���    �����Ϊ�գ���ֻ���������ֶ���Ч��

Private Declare Sub yh_virtualbalance Lib "Hisint" Alias "int_virtualbalance" _
    (ByVal Serial_No As String, ByVal ForeBalance_Flag As String, ByVal balance_no As String, ByVal Transaction_No As String, _
    Selfpay As Double, Hookpay As Double, Tcpay As Double, Tcselfpay As Double, Basepay As Double, _
    outpay As Double, Preqpay As Double, Preqselfpay As Double, ActualselfPay As Double, SubsidyPay As Double, _
    OfficialPay As Double, ByVal ryzt As String, ByVal initinstitution As String, tsggzfbf As Double, ByVal error_code As String)

'7 �������ʶ��
'���룺���������͡�ҽ��������ҽԺ���롢�����ˡ��������ơ������š�����š���ҽʱ�䣻
'�����˳��š����š����˱��롢�������Ա𡢳������ڡ����֤�š���ʼ���������ơ�����������룻
'2007-10-15�޸�:
'int_outpatientidentify()�����������������������Ա״̬��akc021����Ⱥ���akc300
'��ϸ���ݼ�: ����˵��
'�����׵�Ŀ���ǲ������������ǰ��IC���ж���������Ϣ��HIS���ߴ���ҽ���������ݿ��л�ȡ���˵Ļ�����Ϣ����Ҫʱ���������ݿ�ȡ����Ա״̬��������������Ϣ����HIS��
'���룺���������͡�ҽԺ���롢�����ˡ����ҡ������š�����š���ҽʱ�䡢������ϡ�������ƺţ�
'�����˳��š����š����˱��롢�������Ա𡢳������ڡ����֤�š���Ա״̬����Ⱥ��𡢳�ʼ���������ơ�����������룻
Private Declare Sub yh_outpatientidentify Lib "Hisint" Alias "int_outpatientidentify" _
    (ByVal card_mode As String, ByVal doctorname As String, ByVal Hospital_No As String, ByVal Arranger As String, ByVal Section_No As String, _
    ByVal anamnesis_No As String, ByVal outpatient_No As String, ByVal hospitalize_time As String, _
    ByVal admit_diagnose As String, ByVal Transaction_No As String, ByVal Serial_No As String, _
    ByVal CARD_NO As String, ByVal Personal_No As String, ByVal Name As String, ByVal Sex As String, ByVal birthdate As String, _
    ByVal Identify As String, ByVal akc021 As String, ByVal akc300 As String, ByVal initinstitution As String, accountremain As Double, ByVal officesign As String, ByVal error_code As String)

'8 IC��������Ϣ��ѯ
'���룺���������ͣ�
'���: �����š��������Ա����֤�š����䡢�������
'2007-10-15�޸�:
'int_cardinfo()����������AKC300��Ⱥ������������1��ְ����2������
'����������HIS��ѯ���˻�����Ϣ?
'���룺���������ͣ�
'����������š��������Ա����֤�š����䡢��Ⱥ��𣬴������
Private Declare Sub yh_cardinfo Lib "Hisint" Alias "int_cardinfo" _
    (ByVal Code_Mode As String, Amount As Double, ByVal CARD_NO As String, ByVal Name As String, _
    ByVal Sex As String, ByVal Identify As String, age As Double, ByVal akc300 As String, ByVal error_code As String)

'9 �������
'����: ����������
'���: �������
Private Declare Sub yh_changepassword Lib "Hisint" Alias "int_changepassword" _
    (ByVal Code_Mode As String, ByVal error_code As String)

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
Private Declare Sub yh_init_yns Lib "Hisint" Alias "init" _
    (ByVal Errcode As String)

'13    �Ͽ�����
'���룺��
'���: ��
Public Declare Sub yh_quit Lib "Hisint" Alias "quit" ()

'14 IC��Ȧ��
'���룺��
'���: �������
Private Declare Sub yh_loadcard Lib "Hisint" Alias "int_loadcard" (ByVal error_code As String)
    
'15 ���ݴ���
'���룺��
'���: �������
Private Declare Sub yh_datatrans Lib "Hisint" Alias "int_datatrans" (ByVal error_code As String)


'16 �������
'���룺������𣬾���˳��ţ�������ƺţ�����������ͣ�
'���: �������
Private Declare Sub yh_transaction Lib "Hisint" Alias "int_transaction" _
    (ByVal Trade_Sort As String, ByVal Serial_No As String, ByVal Transaction_No As String, ByVal Affirm_Mode As String, ByVal error_code As String)

'17 ��ȡ������ƺ�
'���룺�ޣ�
'���: ������ƺ�
Private Declare Sub yh_gettranssequence Lib "Hisint" Alias "int_gettranssequence" (ByVal Transaction_No As String)

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
'�������: ���ߣ�ͳ���ۼƣ�����ͳ��֧���޶��ͳ��֧���޶�����ۼƣ�������Ϣ��������־���룬��ҩ���ƣ��������
Private Declare Sub yh_RyspInfo Lib "Hisint" Alias "int_RyspInfo" _
   (ByVal series_no As String, qfx As Double, tclj As Double, dczfxe As Double, _
    dbxe As Double, jslj As Double, ByVal qfxinfo As String, ByVal spbzbm As String, ByVal yyxz As String, ByVal error_code As String)

'�����������ǳ�Ժ����ʱ���޸ĳ�Ժ��ϡ���Ժʱ��ʱ���á�
'���룺˳��š���Ժԭ�򡢳�Ժʱ�䡢��Ժ��ϡ���Ժ�����ˡ���Ժ���ҡ���Ժ��λ��
'�����������룻
Private Declare Sub yh_ReLeaveHosInfo Lib "Hisint" Alias "int_ReLeaveHosInfo" _
   (ByVal series_no As String, ByVal Cyyy As String, ByVal Cysj As String, ByVal Cyzd As String, _
   ByVal Cyjbr As String, ByVal Cyks As String, ByVal Cycw As String, ByVal error_code As String)

'��Ӧ������ҽ��
'���ַ�Χ��������int_sicksortchk�������ú���Ϊ����������HIS�ӿ��ڵ���int_feebalance��������������ʽ����ǰ������øú���
'����int_sicksortchk����������ŵ���int_feebalance�������������ú���ʵ�ֹ�������ɷ�����ϸ��Ӧ�ĵ����ַ�Χ�Ľ�����
'��������ϸ��Ӧ�ĵ����ַ���ǰ̨�����ֱ���֮��ͨ��'$'�ָ�����0101$0102$0103$0104��HISǰֻ̨���ڸú������صĲ��ַ�Χ�ڽ��в���ѡ��
'���Ϊ��������û��߲��ܽ��е����ֽ��㡣Ŀǰֻ�����ⲡ�����סԺ������á����HIS���ϸ���Ʋ���ѡ��Χ��
'����������ĵĲ����벡��������ϸ��ƥ�䣬������˷����շ���ϸ�벡�ֽ����շ���Ŀ��������������HIS�����̺�ҽԺ�е���
Private Declare Sub yh_sicksortchk Lib "Hisint" Alias "int_sicksortchk" _
    (ByVal Serial_No As String, ByRef sicksorts As String, ByRef error_code As String)

Private Declare Sub yh_init_kms Lib "Hisint" Alias "init" _
    (ByVal HospNO As String, ByVal AverageFeeType As String, ByVal Tsyybz As String, ByVal Errcode As String)

'������ҽ�����У�Ԥ������
Private Declare Sub yh_AlertInfo_kms Lib "Hisint" Alias "int_alertinfo" _
    (ByVal SerialNO As String, ByRef ErrorCode As String, ByRef ErrMsg As String)

Public Const gint������ As Integer = 31

'���½ṹ�����ڼ�¼������������Ա��ڽ���ʱ�˶�
Private Type typBalance
    cur�����ʻ� As Double
    curҽ������ As Double
    cur��ͳ�� As Double
    cur����Ա���� As Double
    cur���ⲹ�� As Double
End Type
Private pre_Balance As typBalance

Public Function ҽ����ʼ��_����(ByVal intinsure As Integer) As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim strAverageFeeType As String, strTsyybz As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mblnInit Then
        ҽ����ʼ��_���� = True
        Exit Function
    End If
    
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & intinsure
    Call OpenRecordset(rsTemp, "��ȡҽԺ����")
    gstrҽԺ���� = Nvl(rsTemp!ҽԺ����, "")
    
    mstrErr = Space(4)
    If intinsure <> gint������ Then
        Call yh_init_yns(mstrErr)
    Else
        Call yh_init_kms(gstrҽԺ����, mstrAverageFeeType, mstrTsyybz, mstrErr)
    End If
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbExclamation, gstrSysName
    Else
        mblnInit = True
        ҽ����ʼ��_���� = True
    End If
    
    '�����صķ������������ҽԺ��־���浽�����ʻ��У���������־�����ڽ���ʱԭ�����Ƶ����ս����¼��
    gstrSQL = "zl_���ղ���_Insert(" & intinsure & ",0,'ƽ���������','''" & mstrAverageFeeType & "''',10)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ƽ���������")
    gstrSQL = "zl_���ղ���_Insert(" & intinsure & ",0,'����ҽԺ��־','''" & mstrTsyybz & "''',11)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������ҽԺ��־")
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ֹ_����() As Boolean
    Call yh_quit
    mblnInit = False
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long, Optional ByVal intinsure As Integer) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������strSelfNO-���˱�ţ�ˢ���õ���strSelfPwd-�������룻bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim str���� As String, STR���� As String, str�Ա� As String
    Dim str���֤�� As String, str�������� As String, lng���� As Double, str��λ���� As String, str��λ���� As String
    Dim str��ʼ������ As String, str����� As String, str��Ⱥ��� As String, str��Ա״̬ As String, str������Ⱥ��־ As String
    Dim str������� As String, str������ƺ� As String, str����Ա As String, strҽ���� As String
    Dim str��ʷ˳��� As String
    
    Dim strArranger As String
    Dim strSection As String
    Dim strPatiNo As String
    
    Dim str������ As String, lng����ID As Long, str�������� As String
    Dim rsTemp As New ADODB.Recordset
    Dim ybhcf As New ADODB.Recordset    '���ڼ�¼���¿���
    Dim dat��ǰ As Date
    Dim strIdentify As String, str���� As String
    '---------��������ʹ��--------
    Dim str��ϱ��� As String, str������� As String, int������� As Integer
    '-----------------------------
    
    On Error GoTo errHandle
    '��ʼ������ȫ�ֵı���
    mstrҽ���� = Space(20)
    mstr˳��� = Space(19)
    mcur�ʻ���� = 0
    
    str���� = Space(18)
    STR���� = Space(60)
    str�Ա� = Space(3)
    str���֤�� = Space(20)
    str��Ⱥ��� = Space(3)
    str��Ա״̬ = Space(3)
    str�������� = Space(10)
    str��ʼ������ = Space(4)
    str������� = Space(56)
    str������ƺ� = Space(18)
    str����Ա = Space(4)
    dat��ǰ = zlDatabase.Currentdate
    
    If frmIdentify����.GetIdentifyMode(intinsure, bytType, str������, lng����ID, str��������) = False Then
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
    '��ȡ������ƺ� gzh
    str������ƺ� = Get�����()
    If str������ƺ� = "" Then Exit Function
    If bytType = 0 Then
        '���ã������С�����ʡ����ͨ����ŵ�OutPatientidentifhy�����������CardInfo
        If lng����ID = 0 Then
            Call yh_outpatientidentify(str������, strArranger, gstrҽԺ����, strArranger, strSection, strPatiNo, _
                strPatiNo, Format(dat��ǰ, "yyyy-MM-dd"), str�������, str������ƺ�, mstr˳���, str����, _
                mstrҽ����, STR����, str�Ա�, str��������, str���֤��, str��Ա״̬, str��Ⱥ���, str��ʼ������, mcur�ʻ����, str����Ա, mstrErr)
        Else
            Call yh_cardinfo(str������, mcur�ʻ����, str����, STR����, str�Ա�, str���֤��, lng����, str��Ⱥ���, mstrErr)
        End If
    Else
        Call yh_cardinfo(str������, mcur�ʻ����, str����, STR����, str�Ա�, str���֤��, lng����, str��Ⱥ���, mstrErr)
    End If
    If mstrErr <> "0000" Then
        Screen.MousePointer = vbDefault
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr˳��� = TrimStr(mstr˳���)
    str���� = TrimStr(str����)
    STR���� = TrimStr(STR����)
    str���֤�� = TrimStr(str���֤��)
    str��Ⱥ��� = TrimStr(str��Ⱥ���)
   ' str��Ա״̬ = TrimStr(str��Ա״̬)

    If bytType = 0 And lng����ID = 0 Then
        'ֻ����ͨ������ܵõ�ҽ���ţ��������׵���CardInfo�������޷��õ�ҽ����
        mstrҽ���� = TrimStr(mstrҽ����)
    Else
        '��ΪסԺδ����ҽ���ţ�ֻ�д����ݿ���ȡ�����ûȡ�����򽫿�����Ϊҽ���ű��棬����Ժʱ�ٸ���
        gstrSQL = "Select ҽ���� From �����ʻ� Where ����=" & intinsure & " And ����='" & str���� & "'"
        Call OpenRecordset(rsTemp, "��ȡԭҽ����")
        If Not rsTemp.EOF Then
            mstrҽ���� = Nvl(rsTemp!ҽ����)
        End If
        If Trim(mstrҽ����) = "" Then
            mstrҽ���� = str����
        Else
            mstrҽ���� = Mid(mstrҽ����, 2)
        End If
    End If
    strҽ���� = TrimStr(mstrҽ����)
    mbln����Ա = (TrimStr(str����Ա) = "1")
    
    If bytType = 0 And lng����ID = 0 Then
        'ֻ����ͨ����ͨ������outpatientidentify�ӿڵõ�˳���
        If mstr˳��� = "" Then
            MsgBox "δ�ܴ�ǰ�÷��������˳���,�����Ի��鿨��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If str���� = "" Then
        MsgBox "δ�ܴӿ��ж�ȡ����,�����Ի��鿨��", vbInformation, gstrSysName
        Exit Function
    End If
    If mstrҽ���� = "" Then
        MsgBox "δ�ܴӿ��ж�ȡҽ����,�����Ի��鿨��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mbln����Ա Then
        '����ǹ���Ա����Ҫ����yh_RyspInfo��ȡ������Ϣ
        Dim cur���� As Double, curͳ���ۼ� As Double, cur����ͳ���޶� As Double, cur���ͳ���޶� As Double, cur�����ۼ� As Double
        Dim str������Ϣ As String, str������־���� As String, str��ҩ���� As String
        Call yh_RyspInfo(mstr˳���, cur����, curͳ���ۼ�, cur����ͳ���޶�, cur���ͳ���޶�, cur�����ۼ�, str������Ϣ, str������־����, str��ҩ����, mstrErr)
        cur���� = strVal(cur����)
        curͳ���ۼ� = strVal(curͳ���ۼ�)
        cur����ͳ���޶� = strVal(cur����ͳ���޶�)
        cur���ͳ���޶� = strVal(cur���ͳ���޶�)
        cur�����ۼ� = strVal(cur�����ۼ�)
        str������Ϣ = strVal(str������Ϣ)
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'����','''" & cur���� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���¹���Ա������")
        gstrSQL = "zl_����������Ϣ_insert(1," & lng����ID & "," & intinsure & "," & Year(dat��ǰ) & ",'" & _
        mstr˳��� & "'," & cur���� & "," & curͳ���ۼ� & "," & _
        cur����ͳ���޶� & "," & cur���ͳ���޶� & "," & cur�����ۼ� & ",'" & str������Ϣ & "','" & str������־���� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����������Ϣ")
    End If
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    mstrҽ���� = str������ & Left(mstrҽ����, 19)
    strIdentify = str���� & ";" & mstrҽ���� & ";;" & TrimStr(STR����) & ";" & TrimStr(str�Ա�) & ";" & TrimStr(str��������) & ";" & TrimStr(str���֤��) & ";"
    strIdentify = Replace(strIdentify, " ", "")
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    ';8����;9.˳���;10��Ա���(��ְ�����ݡ�ѧ����ͯ����ѧ������ѧ����������);11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1,2);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�;23�������� (1����������)
        
    '������������סԺ���ǾͲ���ʹ���µ�˳��š�����������ǰ��
    gstrSQL = "select ˳��� from �����ʻ� where ����=" & intinsure & " and ����='" & str���� & "'"
    Call OpenRecordset(rsTemp, "����ҽ��")
    If rsTemp.RecordCount > 0 Then
        str��ʷ˳��� = Nvl(rsTemp("˳���"))
    End If
    If bytType = 2 Then mstr˳��� = str��ʷ˳���
    
    If IsDate(str��������) = True Then
        lng���� = DateDiff("yyyy", CDate(str��������), dat��ǰ)
    End If
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";" & str��ʷ˳���             '9.˳���
    str���� = str���� & ";"                             '10��Ա���
    str���� = str���� & ";" & mcur�ʻ����              '11�ʻ����
    str���� = str���� & ";0"                            '12��ǰ״̬
    str���� = str���� & ";" & IIf(lng����ID <> 0, lng����ID, "") '13����ID
    str���� = str���� & ";1"                            '14��ְ(1,2,3)
    str���� = str���� & ";"                             '15����֤��
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";" & mcur�ʻ����              '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    str���� = str���� & ";"                             '23�������� (1����������)
    
    '-------------------------------------------------------------------------------
    '������ҽ�����ظ����(2007-6-8)
    On Error Resume Next
    Err = 0
    gstrSQL = "update �����ʻ� set ����='" & str���� & "' where ����ID in(select ����ID from �����ʻ� where substr(ҽ����,2)='" & strҽ���� & "') "
    Call OpenRecordset(rsTemp, "���¿���,ҽ����")
    If Err <> 0 Then
        gstrSQL = "update ҽ�����˵��� set ����='" & str���� & "' where ����ID in(select ����ID from �����ʻ� where substr(ҽ����,2)='" & strҽ���� & "') "
        Call OpenRecordset(rsTemp, "�ύ")
    End If
    On Error GoTo errHandle
    
    '----------------------------------------
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID, intinsure)
    If lng����ID = 0 Then Exit Function 'δ������ȷ�ı����ʻ�
    
    If bytType = 0 And lng����ID > 0 Then
        '��������ⲡ�����Բ����ͬʱ���о���Ǽ�
        
        '�ٴγ�ʼ������
        mstrҽ���� = Space(20)
        str���� = Space(18)
        STR���� = Space(60)
        str�Ա� = Space(3)
        str���֤�� = Space(20)
        str��Ա״̬ = Space(3)
        str��Ⱥ��� = Space(3)
        str������Ⱥ��־ = Space(3)
        str�������� = Space(10)
        str��ʼ������ = Space(4)
        mstr˳��� = Space(19)
        
        str����� = Get�����
        If str����� = "" Then
            Exit Function
        End If
        
        'ȡ�ò��ֵ������������ز��ʹ�1
        gstrSQL = "Select Nvl(���,0) ��� From ���ղ��� Where ID=" & lng����ID
        Call OpenRecordset(rsTemp, "ȡ�������")
        int������� = rsTemp!���
        
        'ֻ���������ز���Ҫ��ȡ���˵��������
        If int������� <> 0 Then
            Call frm�����Ϣ.ShowME(lng����ID, str��ϱ���, str�������, True)
        End If
        If str������� = "" Then str������� = "��ͨ"
        'str������� = "����������" '����ʱ��
        '0092-����Ⱥ�������0094-סԺ���������ó���ͨ��
        mstrErr = Space(4)
        Call yh_admit(str������, LeftDB(UserInfo.����, 8), gstrҽԺ����, LeftDB(UserInfo.����, 8), "����", _
            LeftDB(lng����ID, 12), LeftDB(lng����ID, 12), IIf(Val(rsTemp!���) = 0, "0", "1"), LeftDB(str��������, 8), _
            Format(dat��ǰ, "yyyy-MM-dd HH:mm:ss"), str�������, str�����, mstr˳���, str����, _
            mstrҽ����, STR����, str�Ա�, str��������, str���֤��, str��Ա״̬, str��Ⱥ���, str������Ⱥ��־, str��ʼ������, str��λ����, str��λ����, mstrErr)
        
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
            'ҽ�����ݿ�ع�
            Call yh_transaction("0", mstr˳���, str�����, "0", mstrErr)
            Exit Function
        End If
        mstr˳��� = TrimStr(mstr˳���) '1����������Ԥ��
        If mstr˳��� = "" Then
            MsgBox "���ܵõ���ȷ����Ժ�Ǽ�˳��š�", vbInformation, gstrSysName
            Call yh_transaction("0", mstr˳���, str�����, "0", mstrErr)
            Exit Function
        End If
        Call yh_transaction("0", mstr˳���, str�����, "1", mstrErr)
        
        str������Ⱥ��־ = TrimStr(str������Ⱥ��־)
        If str������Ⱥ��־ = 1 Then
            On Error Resume Next
            Err = 0
            gstrSQL = "update �����ʻ� set ������Ⱥ��־=" & str������Ⱥ��־ & "  where ����ID=" & lng����ID
            Call OpenRecordset(rsTemp, "������Ա���")
            If Err <> 0 Then
                gstrSQL = "update ҽ�����˵��� set ������Ⱥ��־=" & str������Ⱥ��־ & "  where ����ID=" & lng����ID
                Call OpenRecordset(rsTemp, "������Ա���")
            End If
            On Error GoTo errHandle
        End If
        
        '����Ⱥ������,סԺ����Ҫ����yh_RyspInfo��ȡ������Ϣ
        Call yh_RyspInfo(mstr˳���, cur����, curͳ���ۼ�, cur����ͳ���޶�, cur���ͳ���޶�, cur�����ۼ�, str������Ϣ, str������־����, str��ҩ����, mstrErr)
        cur���� = strVal(cur����)
        curͳ���ۼ� = strVal(curͳ���ۼ�)
        cur����ͳ���޶� = strVal(cur����ͳ���޶�)
        cur���ͳ���޶� = strVal(cur���ͳ���޶�)
        cur�����ۼ� = strVal(cur�����ۼ�)
        str������Ϣ = strVal(str������Ϣ)
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'����','''" & cur���� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
        gstrSQL = "zl_����������Ϣ_insert(1," & lng����ID & "," & intinsure & "," & Year(dat��ǰ) & ",'" & _
        mstr˳��� & "'," & cur���� & "," & curͳ���ۼ� & "," & _
        cur����ͳ���޶� & "," & cur���ͳ���޶� & "," & cur�����ۼ� & "," & str������Ϣ & ",'" & str������־���� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����������Ϣ")
        mstrҽ���� = str������ & Left(TrimStr(mstrҽ����), 19) '2����������Ԥ��
        str���� = TrimStr(str����)
        
        '�����Ѵ��ڼ�¼
        gstrSQL = " Select ����ID,˳��� from �����ʻ� Where ����=" & intinsure & " And ҽ����='" & mstrҽ���� & "'"
        Call OpenRecordset(rsTemp, "�ж��Ƿ��Ѵ���")
        If rsTemp.RecordCount = 0 Then
            'ǿ�ưѵǼ�˳��š����µ�ҽ�������루�����ʻ��е�˳���ֻ����סԺ������ʹ��mstr˳��ţ�
            gstrSQL = "ZL_�����ʻ�_�޸�ҽ����(" & lng����ID & "," & intinsure & _
                        ",'" & str���� & "','" & mstrҽ���� & "','" & str��ʷ˳��� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        Else
            '����ǰû�м�¼������û���ҵ�˳��ţ����ֻ���¿��ż���
            gstrSQL = "ZL_�����ʻ�_�޸�ҽ����(" & rsTemp!����ID & "," & intinsure & ",'" & str���� & "','" & mstrҽ���� & "','" & Nvl(rsTemp!˳���) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
            lng����ID = rsTemp!����ID
        End If
    End If
    '�õ�������ϸ���ݵ�������ƺţ��Ա��ڶ������
    If bytType = 0 Then
        mstr��ϸ����� = Get����� '3�������������
        If mstr��ϸ����� = "" Then
            Exit Function
        End If
    End If
    
    On Error Resume Next
    Err = 0
    If str��Ա״̬ <> "" Then
        str��Ա״̬ = TrimStr(str��Ա״̬)
        '���±����ʻ��е���Ա���:
        'str��Ա״̬ = Decode(str��Ա״̬, "11", "��ְ", "21", "����", "61", "ѧ����ͯ", "62", "��ѧ��", "63", "��ѧ��", "64", "������")
        gstrSQL = "update �����ʻ� set ��Ա���=" & str��Ա״̬ & "  where ����ID=" & lng����ID
        Call OpenRecordset(rsTemp, "������Ա���")
        If Err <> 0 Then
            gstrSQL = "update ҽ�����˵��� set ��Ա���=" & str��Ա״̬ & "  where ����ID=" & lng����ID
            Call OpenRecordset(rsTemp, "������Ա���")
        End If
    End If
    
    '���±����ʻ��е���ְ:
    'str��Ⱥ��� = Decode(str��Ⱥ���, 1, "����ְ��", 2, "�������", 3, "����")
    Err = 0
    gstrSQL = "update �����ʻ� set ��ְ=" & str��Ⱥ��� & " where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "������ְ")
    If Err <> 0 Then
        gstrSQL = "update ҽ�����˵��� set ��ְ=" & str��Ⱥ��� & "  where ����ID=" & lng����ID
        Call OpenRecordset(rsTemp, "������ְ")
    End If
    
    On Error GoTo errHandle
    '���ظ�ʽ:�м���벡��ID
    mlng����ID = lng����ID '4����������Ԥ��
    ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & str����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, strSelfNo As String, ByVal bytPlace As Byte, ByVal intinsure As Integer) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��
'      ��ʾ����λ�ã�10-����,20-��Ժ,30-Ԥ��,40-����
'����: ���ظ����ʻ����Ľ��
    Dim cur��� As Currency
    On Error GoTo errHandle
    
    If Not (strSelfNo = mstrҽ���� And (bytPlace = 10 Or bytPlace = 20)) Then
        Call Get�����(strSelfNo, cur���, intinsure)
        mcur�ʻ���� = cur���
    End If
    'ֱ�������ϴ����ʶ��ʱ�õ������ݷ���
    �������_���� = mcur�ʻ����
    
    '���±����ʻ��е��ʻ����
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'�ʻ����','" & mcur�ʻ���� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, ByVal intinsure As Integer) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim str�������� As String, strTemp As String
    
    Dim cur�Ը����� As Double, cur�Ը���� As Double, cur������� As Double
    Dim strҽ�� As String, str���� As String, str��� As String, str���� As String
    Dim cur�������� As Currency, dbl��� As Double, dbl���� As Double, str�������� As String
    Dim str����ʱ�� As String, str������� As String, str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    If rs��ϸ.EOF = True Then
        MsgBox "�����������ϸ�ٽ���ҽ��Ԥ�㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rs��ϸ("����ID") <> mlng����ID Then
        MsgBox "�ò���δͨ�������֤�����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    With pre_Balance
        .cur��ͳ�� = 0
        .cur�����ʻ� = 0
        .cur����Ա���� = 0
        .cur���ⲹ�� = 0
        .curҽ������ = 0
    End With
    
    'ֻ�����������ʹ�ñ�����
    On Error GoTo errHandle
    '�жϸò����Ƿ�������������
    gstrSQL = "select nvl(A.����ID,0) ����ID,Nvl(B.���,0) ��� from �����ʻ� A,���ղ��� B where A.����ID=" & mlng����ID & " And A.����ID=B.ID(+) and A.����=" & intinsure
    Call OpenRecordset(rsTemp, "ҽ���ӿ�")
    If rsTemp.EOF Then
        '�ǹ���Ա��ʾ gzh
        If mbln����Ա = False Then
        '�����ⲡ�Ĳ�����ҪԤ��
            MsgBox "�ò��˲���Ҫ����Ԥ�㡣", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        str������� = rsTemp!���
    End If
    
    'ɾ��ǰ�÷�����������δ����ϸ
    mstrErr = Space(4)
    Call yh_transaction("1", mstr˳���, mstr��ϸ�����, "0", mstrErr)
            
    '������ϸ����
    strTemp = rs��ϸ("����ID") & "_" & Format(zlDatabase.Currentdate, "ddHHmmss")
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.����,A.����,A.���,A.���㵥λ,B.��Ŀ����,B.��ע" & _
                    " ,Decode(Sign(Instr(A.���,'��')),0,A.���,Substr(A.���,1,Instr(A.���,'��')-1)) as ���" & _
                    " ,Decode(Sign(Instr(A.���,'��')),0,A.���,Substr(A.���,Instr(A.���,'��')+1)) as ����" & _
                    " from �շ�ϸĿ A,����֧����Ŀ B where A.ID=" & rs��ϸ("�շ�ϸĿID") & " and A.ID=B.�շ�ϸĿID and B.����=" & intinsure
        Call OpenRecordset(rsTemp, "����Ԥ��")
        If rsTemp.EOF = True Then
            MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
            Exit Function
        End If
        str���� = Nvl(rsTemp!��ע)
        If str������� = "1" Then
            If str���� <> "01" And str���� <> "02" And str���� <> "90" Then
                MsgBox "����ҽ������ֻ��ʹ��ҩƷ��", vbInformation, gstrSysName
                mstrErr = Space(4)
                Call yh_transaction("1", mstr˳���, mstr��ϸ�����, "0", mstrErr)
                Exit Function
            End If
        End If
        
        strҽ�� = LeftDB(UserInfo.����, 8)
        str��� = LeftDB(IIf(IsNull(rsTemp("���")), "�޹��", rsTemp("���")), 30)
        str���� = LeftDB(IIf(IsNull(rsTemp("����")), " ", rsTemp("����")), 30)
        str���� = LeftDB(UserInfo.����, 24)
        '���ܴ��ݸ�������0��Ŀ����Ϊ��ɾ���Ѿ��ϴ����������ķ��ü�¼
        dbl���� = Val(IIf(rs��ϸ("����") > 0, rs��ϸ("����"), 0))
        If Nvl(rs��ϸ!ʵ�ս��, 0) > 0 Then
            dbl��� = Round(rs��ϸ!ʵ�ս�� / rs��ϸ!����, 3)
        Else
            dbl��� = 0
        End If
        str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        str�������� = ToVarchar(strTemp & "_" & rs��ϸ.AbsolutePosition, 18)
        
        mstrErr = Space(4)
        Call yh_feedetailtrans(mstr˳���, str��������, str����, rsTemp("��Ŀ����"), _
            rsTemp("����"), dbl����, dbl���, str����, str���, " ", strҽ��, str����, mstr��ϸ�����, strҽ��, str����ʱ��, _
            cur�Ը�����, cur�Ը����, cur�������, mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
            'ҽ�����ݿ�ع�
            Call yh_transaction("1", mstr˳���, mstr��ϸ�����, "0", mstrErr)
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
    Dim str��������� As String, str��Ա״̬ As String, cur����ҹ�֧������ As Double
    
    '��������Ԥ��
    str��������� = Get�����
    If str��������� = "" Then
        Exit Function
    End If
    
    str��ʼ������ = Space(4)
    mstrErr = Space(4)
    str�����־ = "0" '�������
    Call yh_virtualbalance(mstr˳���, str�����־, "", str���������, curȫ�Ը�, cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, _
        cur�����Ը�, cur��ͳ��, cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, cur����Աͳ��, str��Ա״̬, str��ʼ������, cur����ҹ�֧������, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .�������ý�� = cur��������
    End With
    
    cur��� = �������_����(mlng����ID, mstrҽ����, 10, intinsure)
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
    
    With pre_Balance
        .cur��ͳ�� = cur��ͳ��
        .curҽ������ = curͳ��֧��
        .cur����Ա���� = cur����Աͳ��
        .cur���ⲹ�� = cur������Աͳ��
        .cur�����ʻ� = cur���
    End With
    
    �����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, ByVal intinsure As Integer, ByRef strAdvance As String) As Boolean
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
    
    '���������
    Dim strSicks As String          '�������صĲ����б�ע���������ⲡĿǰֻ��ѡ��Ѫ͸
    Dim strFeeBalanceType As String, dblXTCS As Double      '���㺯���ķ��س��Σ�֧���������Ѫ͸����
    Dim strSickSel As String        '����Աѡ��Ĳ��ֱ���
    Dim bln������ As Boolean
    
    Dim cur�Ը����� As Double, cur�Ը���� As Double, cur������� As Double
    Dim strҽ�� As String, str���� As String
    Dim str��� As String, str���� As String
    Dim str���� As String, str˳��� As String
    
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double
    Dim cur������Աͳ�� As Double, cur������Ա�Ը� As Double, cur����Աͳ�� As Double, str��Ա״̬ As String, cur����ҹ�֧������ As Double
    Dim str����ʱ�� As String, str��Ժ��� As String, str��Ժԭ�� As String, str��Ժ������ As String, str��Ժ���� As String, str��Ժ���� As String
    Dim strErrMsg As String, str������Ⱥ��־ As String
    Dim blnReverse As Boolean   'У�����ݱ�־
    Dim str���㷽ʽ As String
    On Error GoTo errHandle
    
    Call DebugTool("�����շ�")
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    gstrSQL = "Select A.ID,A.����ID,A.NO,A.�Ǽ�ʱ��,A.������ as ҽ��," & _
            "   A.����*A.���� as ����,Round(A.���ʽ��/(A.����*A.����),3) as ʵ�ʼ۸�,A.���ʽ��," & _
            "   A.�շ����,D.��Ŀ���� as �շ���Ŀ,D.��ע,B.���� as ��Ŀ����," & _
            "   decode(Instr(B.���,'��'),0,B.���,substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
            "   decode(Instr(B.���,'��'),0,'',substr(B.���,Instr(B.���,'��')+1)) as ����," & _
            "   C.���� as ��������" & _
            " From (Select * From ������ü�¼ Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=" & intinsure & _
            " Order by A.ID"
    Call OpenRecordset(rs��ϸ, "����ҽ��")
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000, gstrSysName, "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    lng����ID = rs��ϸ("����ID")
    
    '�жϱ����Ƿ�Ϊ�������ⲡ�������������ID<>0
    gstrSQL = " Select Nvl(����ID,0) AS ����ID From �����ʻ� " & _
              " Where ����=" & intinsure & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�жϱ����Ƿ�Ϊ�������ⲡ����")
    lng����ID = Val(rsTemp("����ID"))
    bln������ = (lng����ID <> 0)
    '�ж��Ƿ�Ϊ������Ⱥ,����������Ⱥ��־=1
    gstrSQL = "select nvl(������Ⱥ��־,0) ������Ⱥ��־ from �����ʻ� where ����ID=" & lng����ID & " and ����=" & intinsure
    Call OpenRecordset(rsTemp, "�жϱ����Ƿ�Ϊ����������Ⱥ��־����")
    str������Ⱥ��־ = Nvl(rsTemp!������Ⱥ��־, 0)
    'һ��������ϸ����
    '˳��Ų��������֤ʱ���ص�ֵ:mstr˳���
    Call DebugTool("������ϸ����")
    strҽ�� = LeftDB(IIf(IsNull(rs��ϸ("ҽ��")), UserInfo.����, rs��ϸ("ҽ��")), 8)
    str���� = LeftDB(IIf(IsNull(rs��ϸ("��������")), UserInfo.����, rs��ϸ("��������")), 24)
    
    '��ͨ��������û��Ԥ�㣬���Ի���Ҫ���������ϸ
    'ɾ��ǰ�÷�����������δ����ϸ������ǰһ��ȷ��ʱ��ϸ����ɹ���������ʧ��ʱ��
    Call DebugTool("ɾ��������ϸ")
    mstrErr = Space(4)
    Call yh_transaction("1", mstr˳���, mstr��ϸ�����, "0", mstrErr)
    
    Do Until rs��ϸ.EOF
        str��� = LeftDB(IIf(IsNull(rs��ϸ("���")), "�޹��", rs��ϸ("���")), 30)
        str���� = LeftDB(IIf(IsNull(rs��ϸ("����")), " ", rs��ϸ("����")), 30)
        str���� = LeftDB(IIf(IsNull(rs��ϸ("��������")), UserInfo.����, rs��ϸ("��������")), 24)
        cur�������� = cur�������� + rs��ϸ("���ʽ��")
        str����ʱ�� = Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
        str���� = Nvl(rs��ϸ!��ע)
        
        Call DebugTool("��ϸ�ϴ�")
        mstrErr = Space(4)
        Call yh_feedetailtrans(mstr˳���, rs��ϸ("ID"), str����, rs��ϸ("�շ���Ŀ"), LeftDB(rs��ϸ("��Ŀ����"), 24), _
            rs��ϸ("����"), Round(rs��ϸ("ʵ�ʼ۸�"), 3), str����, str���, " ", strҽ��, str����, mstr��ϸ�����, strҽ��, str����ʱ��, _
            cur�Ը�����, cur�Ը����, cur�������, mstrErr)
           mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            Call DebugTool("�ϴ���������")
            MsgBox GetErrInfo(mstrErr, intinsure)
            'ҽ�����ݿ�ع�
            Call yh_transaction("1", mstr˳���, mstr��ϸ�����, "0", mstrErr)
            Exit Function
        End If
        gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rs��ϸ("ID") & "," & cur������� & ",'" & cur�Ը����� & "|" & cur�Ը���� & "|" & cur������� & "')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        rs��ϸ.MoveNext
    Loop
    If lng����ID <> 0 Then cur�������� = g��������.�������ý��        '�ô���Ӧ�ս���Ԥ�㱣��һ��
        
    '��Ԥ��������Ŀǰ��ҽ��ֻ������ſ���Ԥ����ʡҽ��������ô˺���
    If intinsure = gint������ Then
        mstrErr = Space(4)
        strErrMsg = Space(255)
        Call yh_AlertInfo_kms(mstr˳���, mstrErr, strErrMsg)
           mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            If MsgBox("ҽ������������ϢԤ��������������㡰�ǡ���ȡ��������㡰��" & vbCrLf & _
            strErrMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    '����дIC��
    Call DebugTool("дIC��")
    str������ = Left(strSelfNo, 1)
    str��ʼ������ = Space(4)
    If CDbl(cur�����ʻ�) <> 0 Then
        mstrErr = Space(4)
        Call yh_cardpay(str������, mstr˳���, strҽ��, "�����շ�", CDbl(cur�����ʻ�), str��ʼ������, mstrErr)
    End If
    mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        Call DebugTool("д������")
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        'ҽ�����ݿ�ع�
        Call yh_transaction("1", mstr˳���, mstr��ϸ�����, "0", mstrErr)
        Exit Function
    End If
    
    '����������ҽ��
    '˵�������ﵥ�����շ���ֻ��������Ѫ͸�����HISǰ̨��������ϸint_sicksortchk������������������֣�
    'HISǰ̨Ҳֻ��ѡ�����е�Ѫ͸�����֣�����¼�����������֡�������صĲ�����Ѫ͸�������ֱ���1301���򣬲����������ֽ��㡣
    If intinsure = gint������ And bln������ Then
        mstrErr = Space(4)
        strSicks = Space(100)
        
        Call yh_sicksortchk(mstr˳���, strSicks, mstrErr)
        If mstrErr <> "0000" Then
            Err.Raise 9000, gstrSysName, "����ʧ�ܵ������ʻ��ѿۿ" & cur�����ʻ� & "Ԫ������HIS����ϵ��" & vbCrLf & "��ϸ����" & GetErrInfo(mstrErr, intinsure)
            Exit Function
        End If
        
        '�ж�������֣���Ҫ����Աѡ��
        strSicks = Trim(strSicks)
        '�������ֹ�����Աѡ��
        strSickSel = frm���������ⲡ��ѡ��.ShowSelect(strSicks)
    End If
    
    '�������ý���
    Call DebugTool("���ý���")
    str��������� = Get�����
    If str��������� = "" Then
        Exit Function
    End If
    
    str��ʼ������ = Space(4)
    mstrErr = Space(4)
    Call yh_feebalance(mstr˳���, strҽ��, str����, strSickSel, str���������, _
        curȫ�Ը�, cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, cur�����Ը�, cur��ͳ��, _
        cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, cur����Աͳ��, str��Ա״̬, str��ʼ������, cur����ҹ�֧������, _
        strFeeBalanceType, dblXTCS, mstrErr)
    mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        'ҽ�����ݿ�ع�
        Call yh_transaction("2", mstr˳���, str���������, "0", mstrErr)
        Exit Function
    End If
    Call yh_transaction("2", mstr˳���, str���������, "1", mstrErr)
    
    '���������У���������
    If Not (pre_Balance.cur��ͳ�� = cur��ͳ�� And pre_Balance.cur���ⲹ�� = cur������Աͳ�� And pre_Balance.curҽ������ = curͳ��֧�� _
            And pre_Balance.cur�����ʻ� = cur�����ʻ� And pre_Balance.cur����Ա���� = cur����Աͳ��) Then
        blnReverse = True
        str���㷽ʽ = "�����ʻ�|" & cur�����ʻ�
        str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curͳ��֧��
        str���㷽ʽ = str���㷽ʽ & "||��ͳ��|" & cur��ͳ��
        str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Աͳ��
        str���㷽ʽ = str���㷽ʽ & "||���ⲹ��|" & cur������Աͳ��
        
        #If gverControl < 2 Then
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
        #Else
            strAdvance = str���㷽ʽ
            gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    
    '�ġ���������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    '���� curͳ���ۼ� ������Ŀ����Ϊ�˵���API�����ͼ���
    Dim cur���� As Double, curͳ���ۼ� As Double, cur����ͳ���޶� As Double, cur���ͳ���޶� As Double
    Dim intסԺ�����ۼ� As Integer, cur�����ۼ� As Double, str������Ϣ As String, str������־���� As String, str��ҩ���� As String
    curDate = zlDatabase.Currentdate
            
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    
    mstrErr = Space(4)
    Call yh_RyspInfo(mstr˳���, cur����, curͳ���ۼ�, cur����ͳ���޶�, cur���ͳ���޶�, cur�����ۼ�, str������Ϣ, str������־����, str��ҩ����, mstrErr)
    curͳ�ﱨ���ۼ� = strVal(curͳ���ۼ�)
    cur���� = strVal(cur����)
    cur����ͳ���޶� = strVal(cur����ͳ���޶�)
    cur���ͳ���޶� = strVal(cur���ͳ���޶�)
    cur�����ۼ� = strVal(cur�����ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�����Ը� & "," & cur�����ۼ� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")

    '��cur������Ա�Ը���Ϊ����
    '��ע�б������ݣ��ز��б�|���ν���ѡ��Ĳ��ֱ���|���㺯�����صķ��ý������|Ѫ͸����|���γ�ʼ���������ص�ƽ���������|����ҽԺ��־
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & Get���ֱ���(lng����ID) & "," & cur�����Ը� & "," & _
        cur�������� & "," & curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ� & ",'" & mstr˳��� & "'" & _
        ",NULL,NULL,'" & strSicks & "|" & strSickSel & "|" & strFeeBalanceType & "|" & dblXTCS & "|" & mstrAverageFeeType & "|" & mstrTsyybz & "','" & str������Ⱥ��־ & "'," & IIf(blnReverse, "1", "0") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")

    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    '�塢��������Ҫ���г�Ժ��Ϣ�޸�
    If lng����ID <> 0 Then
        If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
        str����ʱ�� = Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
        str��Ժ��� = lng����ID
        str��Ժ������ = LeftDB(IIf(IsNull(rs��ϸ("ҽ��")), UserInfo.����, rs��ϸ("ҽ��")), 8)
        str��Ժ���� = "1"
        str��Ժԭ�� = "9"
        str��Ժ���� = LeftDB(IIf(IsNull(rs��ϸ("��������")), UserInfo.����, rs��ϸ("��������")), 24)
        mstrErr = Space(4)
        Call yh_ReLeaveHosInfo(mstr˳���, str��Ժԭ��, str����ʱ��, str��Ժ���, str��Ժ������, str��Ժ����, str��Ժ����, mstrErr)
        mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        'Exit Function ��������������֣���Ϊ�Ѿ�����ɹ�����ִ��
        End If
    End If
    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, str�˷������ As String
    Dim lng����ID As Long, str˳��� As String, lng�������� As Double
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curƱ���ܽ�� As Currency
    Dim curDate As Date
    Dim str������Ⱥ��־ As String
    
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ��  From ������ü�¼ Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    lng����ID = rsTemp("����ID")
    rsTemp.Close
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & intinsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    str˳��� = rsTemp("֧��˳���")
    str������Ⱥ��־ = Trim(Nvl(rsTemp("������Ⱥ��־"), 0))
    lng�������� = IIf(IsNull(rsTemp("�ⶥ��")), 0, rsTemp("�ⶥ��"))
    
    If Is����ȷ(lng����ID, intinsure) = False Then
        Exit Function
    End If
    
    str�˷������ = Get�����
    If str�˷������ = "" Then
        Exit Function
    End If
    
    '3-��ʾ��ͨ����ĸ����˻��˷Ѵ���2-��ʾ��������ĸ����˻�Ԥͳ�������˷�
    mstrErr = Space(4)
    Call yh_recedefeebalance(str˳���, IIf(lng�������� > 0, 2, 3), "", str�˷������, mstrErr)
       mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        'ҽ�����ݿ�ع�
        Call yh_transaction("2", mstr˳���, str�˷������, "0", mstrErr)
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & rsTemp("����") * -1 & "," & rsTemp("�ⶥ��") & "," & _
        rsTemp("ʵ������") * -1 & "," & curƱ���ܽ�� * -1 & "," & rsTemp("ȫ�Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & "," & rsTemp("���Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & _
        cur�����ʻ� * -1 & ",'" & str˳��� & "',null,null,null,'" & str������Ⱥ��־ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")

    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function �����ʻ�תԤ��_����(lngԤ��ID As Long, cur�����ʻ� As Currency, strSelfNo As String, str˳��� As String, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim str������ As String
    Dim str��ʼ������ As String
    Dim strҽ�� As String
    
    On Error GoTo errHandle
    
    If Is����ȷ(lng����ID, intinsure) = False Then Exit Function
    
    str��ʼ������ = Space(4)
    str������ = Left(strSelfNo, 1)
    
    mstrErr = Space(4)
    strҽ�� = LeftDB(UserInfo.����, 8)
    If cur�����ʻ� <> 0 Then Call yh_cardpay(str������, str˳���, LeftDB(UserInfo.����, 8), "Ԥ����", cur�����ʻ�, str��ʼ������, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
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
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(3," & lngԤ��ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & cur�����ʻ� & ",0,0,0,0,0,0," & _
        cur�����ʻ� & ",'" & str˳��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    �����ʻ�תԤ��_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false

    Dim rsTemp As New ADODB.Recordset
    Dim rsTemp1 As New ADODB.Recordset
    Dim curDate As Date
    Dim str������ As String
    Dim str���� As String
    Dim STR���� As String
    Dim str�Ա� As String
    Dim str�������� As String
    Dim str���֤�� As String
    Dim str��Ա״̬ As String
    Dim str��Ⱥ��� As String
    Dim str������Ⱥ��־ As String
    Dim str��ʼ������ As String, str��λ���� As String, str��λ���� As String
    Dim str����� As String   '������ƺ�
    Dim blnTrans As Boolean
    Dim bln���� As Boolean
    Dim str����ID As String
    Dim lng����ID As Long, str�������� As String
    '-----------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    '���� curͳ���ۼ� ������Ŀ����Ϊ�˵���API�����ͼ���
    Dim cur���� As Double, curͳ���ۼ� As Double, cur����ͳ���޶� As Double, cur���ͳ���޶� As Double
    Dim intסԺ�����ۼ� As Integer, cur�����ۼ� As Double, str������Ϣ As String, str������־���� As String, str��ҩ���� As String
    On Error GoTo errHandle
    mstr˳��� = Space(19)
    strҽ���� = Space(20)
    str����� = Space(18)
    str���� = Space(18)
    STR���� = Space(60)
    str�Ա� = Space(3)
    str�������� = Space(10)
    str���֤�� = Space(20)
    str��Ա״̬ = Space(3)
    str��Ⱥ��� = Space(3)
    str������Ⱥ��־ = Space(3)
    str��ʼ������ = Space(4)
    curDate = zlDatabase.Currentdate
    
    'ע�⣺��ʱ���ܶ������ʻ�����Ϊ��δȡ��ҽ���ţ�������Ҫ����ҽ����
    gstrSQL = "Select A.��Ժ����,A.��Ժ����,B.���� as ��Ժ����,C.סԺ��,A.�Ǽ�ʱ��,D.ҽ����,E.ID AS ����ID,E.���� as ���ֱ���,E.��� as ������� " & _
            " From ������ҳ A,���ű� B,������Ϣ C,�����ʻ� D,���ղ��� E " & _
            " Where A.��Ժ����ID=B.ID And A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
            " And A.����ID=C.����ID And A.����ID=D.����ID and D.����=" & intinsure & " and D.����ID=E.ID(+)"
    Call OpenRecordset(rsTemp, "����ҽ��")
    If rsTemp.EOF = True Then
        MsgBox "û�з��ִ˲��˵���Ϣ��", vbExclamation, gstrSysName
        Exit Function
    End If
    lng����ID = Nvl(rsTemp!����ID, 0)
    str�������� = Nvl(rsTemp!���ֱ���)
    
    If IsNull(rsTemp("ҽ����")) = False Then
        str������ = Left(rsTemp("ҽ����"), 1)
    Else
        If frmIdentify����.GetIdentifyMode(intinsure, 1, str������, lng����ID, str��������) = False Then Exit Function
    End If
    
    '��Ժ�Ǽ�
    str����� = Get�����
    If str����� = "" Then
        Exit Function
    End If
'    '�жϲ����Ƿ�Ϊ0
'    gstrSQL = "select nvl(����ID,0) ����ID,ҽ���� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & intInsure
'    Call OpenRecordset(rsTemp, "ҽ���ӿ�")
'    'bln���� = Nvl(rsTemp!����ID, 0)
'    str����ID = Nvl(rsTemp!����ID, 0)
'    strҽ���� = rsTemp!ҽ����
'    strҽ���� = Right(strҽ����, 10)
'    If str����ID <> "0" Then
'       gstrSQL = "select ����,���� from ���ղ��� where ID=" & str����ID & " and ����=" & intInsure
'       Call OpenRecordset(rsTemp, "ҽ���ӿ�")
'       str�������� = Trim(rsTemp!����)
'    End If
    '0092-����Ⱥ�������0094-סԺ���������ó���ͨ��
    mstrErr = Space(4)
    If str�������� = "0093" Then    'ֻ����ҽ����֧�����ս�תסԺ��Ҫ���ֱ��봫Ϊ��
        Call yh_kndadmit(LeftDB(UserInfo.����, 8), Mid(mstrҽ����, 2), gstrҽԺ����, LeftDB(UserInfo.����, 8), LeftDB(rsTemp("��Ժ����"), 8), _
            LeftDB(lng����ID, 12), LeftDB(rsTemp("סԺ��"), 12), IIf(rsTemp("�������") <> "0", "1", "0"), "", _
            Format(rsTemp!��Ժ����, "yyyy-01-01 01:01:01"), LeftDB(��ȡ���Ժ���(lng����ID, lng��ҳID, True, False), 50), str�����, mstr˳���, str����, _
            STR����, str�Ա�, str��������, str��Ա״̬, str��Ⱥ���, str������Ⱥ��־, str��ʼ������, str��λ����, str��λ����, mstrErr)
    Else
        Call yh_admit(str������, LeftDB(UserInfo.����, 8), gstrҽԺ����, LeftDB(UserInfo.����, 8), LeftDB(rsTemp("��Ժ����"), 8), _
            LeftDB(lng����ID, 12), LeftDB(rsTemp("סԺ��"), 12), IIf(rsTemp("�������") <> "0", "1", "0"), LeftDB(IIf(IsNull(rsTemp("���ֱ���")), "0", rsTemp("���ֱ���")), 8), _
            Format(rsTemp!��Ժ����, "yyyy-MM-dd hh:mm:ss"), LeftDB(��ȡ���Ժ���(lng����ID, lng��ҳID, True, False), 50), str�����, mstr˳���, str����, _
            strҽ����, STR����, str�Ա�, str��������, str���֤��, str��Ա״̬, str��Ⱥ���, str������Ⱥ��־, str��ʼ������, str��λ����, str��λ����, mstrErr)
    End If
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        'ҽ�����ݿ�ع�
        Call yh_transaction("0", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
    blnTrans = True
    str������Ⱥ��־ = TrimStr(str������Ⱥ��־)
    str��Ա״̬ = TrimStr(str��Ա״̬)
    str���� = TrimStr(str����)
    If str�������� = "0093" Then
        strҽ���� = mstrҽ����
    Else
        strҽ���� = str������ & Left(TrimStr(strҽ����), 19)
    End If
    
    On Error Resume Next
    Err = 0
    gstrSQL = "update �����ʻ� set ������Ⱥ��־='" & str������Ⱥ��־ & "',��Ա���='" & str��Ա״̬ & "',����='" & str���� & "'  where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "������Ա���")
    If Err <> 0 Then
        gstrSQL = "update ҽ�����˵��� set ������Ⱥ��־='" & str������Ⱥ��־ & "',��Ա���='" & str��Ա״̬ & "',����='" & str���� & "'  where ����ID=" & lng����ID
        Call OpenRecordset(rsTemp, "������Ա���")
    End If
    On Error GoTo errHandle
    
    mstr˳��� = TrimStr(mstr˳���)
    If mstr˳��� = "" Then
        MsgBox "���ܵõ���ȷ����Ժ�Ǽ�˳��š�", vbInformation, gstrSysName
        Call yh_transaction("0", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
    '�ʻ������Ϣ��ȡ���
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
        'סԺ����Ҫ����yh_RyspInfo��ȡ������Ϣ
    Call yh_RyspInfo(mstr˳���, cur����, curͳ���ۼ�, cur����ͳ���޶�, cur���ͳ���޶�, cur�����ۼ�, str������Ϣ, str������־����, str��ҩ����, mstrErr)
        cur���� = strVal(cur����)
        curͳ���ۼ� = strVal(curͳ���ۼ�)
        cur����ͳ���޶� = strVal(cur����ͳ���޶�)
        cur���ͳ���޶� = strVal(cur���ͳ���޶�)
        cur�����ۼ� = strVal(cur�����ۼ�)
        str������Ϣ = strVal(str������Ϣ)
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'����','''" & cur���� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��һ��סԺ���˸�������")
    gstrSQL = "zl_����������Ϣ_insert(2," & lng����ID & "," & intinsure & "," & Year(curDate) & ",'" & _
        mstr˳��� & "'," & cur���� & "," & curͳ���ۼ� & "," & _
        cur����ͳ���޶� & "," & cur���ͳ���޶� & "," & cur�����ۼ� & "," & str������Ϣ & ",'" & str������־���� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����������Ϣ")
    
    'ǿ�ưѵǼ�˳��š����µ�ҽ��������
    gstrSQL = "ZL_�����ʻ�_�޸�ҽ����(" & lng����ID & "," & intinsure & _
                ",'" & str���� & "','" & strҽ���� & "','" & mstr˳��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    Call yh_transaction("0", mstr˳���, str�����, "1", mstrErr)
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        Call yh_transaction("0", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, str˳��� As String, ByVal intinsure As Integer, _
    Optional ByVal bln���ʳ�Ժ As Boolean = False, Optional ByVal bln������Ժ As Boolean = False) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
    Dim str����� As String   '������ƺ�
    Dim strMsg As String
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, str��ʼ������ As String
    Dim cur������Ա�Ը� As Double, cur������Աͳ�� As Double, cur����Աͳ�� As Double, str��Ա״̬ As String, cur����ҹ�֧������ As Double
    
    '���������
    Dim strSicks As String          '�������صĲ����б�ע���������ⲡĿǰֻ��ѡ��Ѫ͸
    Dim strFeeBalanceType As String, dblXTCS As Double      '���㺯���ķ��س��Σ�֧���������Ѫ͸����
    Dim strSickSel As String        '����Աѡ��Ĳ��ֱ���
    Dim str��ע As String
    
    Dim blnTrans As Boolean
    Dim rsInfo As New ADODB.Recordset
    Dim str��Ժԭ�� As String, str��Ժʱ�� As String, str��Ժ��� As String
    Dim str��Ժ������ As String, str��Ժ���� As String, str��Ժ���� As String
    Dim str������Ⱥ��־ As String
    '��Ժ��ʽ:1-����;2-תԺ;3-��������Ӧҽ���ĳ�Ժԭ��0��������Ժ��1��������2��תԺ��3������δסԺ����;ȡ������9������
    Dim rsTemp As New ADODB.Recordset
    Dim rstemp���� As New ADODB.Recordset
    
    On Error GoTo errHandle
    '�������δ����ã��������HIS��Ժ������ͬʱ����ҽ����HIS��Ժ
    Call DebugTool("�����Ժ�Ǽǽӿ�")
    If bln������Ժ Or Not ����δ�����(lng����ID, lng��ҳID) Then
        str��ʼ������ = Space(4)
        
        str����� = Get�����
        If str����� = "" Then
            
        End If
        mstr˳��� = str˳���
        
        '����������ҽ��
        '˵�������ﵥ�����շ���ֻ��������Ѫ͸�����HISǰ̨��������ϸint_sicksortchk������������������֣�
        'HISǰ̨Ҳֻ��ѡ�����е�Ѫ͸�����֣�����¼�����������֡�������صĲ�����Ѫ͸�������ֱ���1301���򣬲����������ֽ��㡣
        'סԺ��������������֣�����Աֻ��ѡ������һ�����н���.
        If intinsure = gint������ Then
            mstrErr = Space(4)
            strSicks = Space(100)
            Call yh_sicksortchk(mstr˳���, strSicks, mstrErr)
            If mstrErr <> "0000" Then
                MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
                Exit Function
            End If
            
            '�������ֹ�����Աѡ��
            strSickSel = frm���������ⲡ��ѡ��.ShowSelect(strSicks)
        End If
        '��Ժ�Ǽ���ͨ�����ý��㽻����ɡ���ʱ���財�˵ķ����Ѿ�ȫ������
        Call DebugTool("����ҽ����Ժ�ӿ�")
        mstrErr = Space(4)
        Call yh_feebalance(mstr˳���, LeftDB(UserInfo.����, 8), LeftDB(UserInfo.����, 24), strSickSel, str�����, curȫ�Ը�, _
            cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, cur�����Ը�, _
            cur��ͳ��, cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, cur����Աͳ��, _
            str��Ա״̬, str��ʼ������, cur����ҹ�֧������, strFeeBalanceType, dblXTCS, mstrErr)
           mstrErr = TrimStr(mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
            'ҽ�����ݿ�ع�
            Call yh_transaction("2", mstr˳���, str�����, "0", mstrErr)
            Exit Function
        End If
        '�ύҽ��ǰ�û����ݿ�
        Call yh_transaction("2", mstr˳���, str�����, "1", mstrErr)
    
        '���±��ս����¼�еı�ע�ֶ�
        '���㺯���Զ����ñ���������˿϶���Ȩ��ִ��zl_���ս����¼_Insert()��
        gstrSQL = "Select nvl(������Ⱥ��־,0) as ������Ⱥ��־ From �����ʻ� Where ����=" & intinsure & " And ����ID=" & lng����ID
        Call OpenRecordset(rstemp����, "��ȡҽ�����˵��������")
        str������Ⱥ��־ = Nvl(rstemp����!������Ⱥ��־, 0)
        str��ע = strSicks & "|" & strSickSel & "|" & strFeeBalanceType & "|" & dblXTCS & "|" & mstrAverageFeeType & "|" & mstrTsyybz
        gstrSQL = " Select ����,��¼ID,����,����ID,���,�ʻ��ۼ�����,�ʻ��ۼ�֧��,�ۼƽ���ͳ��,�ۼ�ͳ�ﱨ��,סԺ����," & _
                  " ����,�ⶥ��,ʵ������,�������ý��,ȫ�Ը����,�����Ը����,����ͳ����,ͳ�ﱨ�����," & _
                  " ���Ը����,�����Ը����,�����ʻ�֧��,֧��˳���,��ҳID From ���ս����¼" & _
                  " Where ��¼ID=(Select Max(��¼ID) From ���ս����¼ Where ����=2 And ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID & ")"
        Call OpenRecordset(rsTemp, "ȡ���һ�ν����¼����")
        If rsTemp.RecordCount <> 0 Then '��������ó�Ժ(2007-08-30)
        gstrSQL = "zl_���ս����¼_insert(2," & rsTemp!��¼ID & "," & rsTemp!���� & "," & rsTemp!����ID & "," & _
            rsTemp!��� & "," & rsTemp!�ʻ��ۼ����� & "," & rsTemp!�ʻ��ۼ�֧�� & "," & rsTemp!�ۼƽ���ͳ�� & "," & _
            rsTemp!�ۼ�ͳ�ﱨ�� & "," & rsTemp!סԺ���� & "," & rsTemp!���� & "," & rsTemp!�ⶥ�� & "," & rsTemp!ʵ������ & "," & _
            rsTemp!�������ý�� & "," & rsTemp!ȫ�Ը���� & "," & rsTemp!�����Ը���� & "," & _
            rsTemp!����ͳ���� & "," & rsTemp!ͳ�ﱨ����� & "," & rsTemp!���Ը���� & "," & rsTemp!�����Ը���� & "," & _
            rsTemp!�����ʻ�֧�� & ",'" & rsTemp!֧��˳��� & "'," & rsTemp!��ҳID & ",NULL,'" & str��ע & "','" & str������Ⱥ��־ & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        End If
        blnTrans = True
        '���³�Ժ��ϣ�ʡ��ҽ����֧�����ս�ת����4��
        mstrErr = Space(4)
        gstrSQL = "select decode(��Ժ��ʽ,'����',0,'תԺ',2,'����',1,'���ս�ת',4,9) ��Ժ��ʽ From ������ҳ " & _
                " Where ����ID = " & lng����ID & " And ��ҳID = " & lng��ҳID
        Call OpenRecordset(rsInfo, "��Ժ��ʽ")
        str��Ժԭ�� = rsInfo!��Ժ��ʽ
        
        gstrSQL = "select b.���� ��Ժ����,��ֹʱ��,����Ա����  " & _
                 " from ���˱䶯��¼ A,���ű� B  " & _
                 " where ����ID=" & lng����ID & " and ��ҳID=" & lng��ҳID & " and ��ֹԭ��=1 " & _
                 " and A.����ID=B.ID"
        Call DebugTool("��ȡ���˳�Ժʱ���SQL��" & gstrSQL)
        Call OpenRecordset(rsInfo, "��Ժ���")
        If rsInfo.RecordCount <> 0 Then
            str��Ժʱ�� = Format(rsInfo!��ֹʱ��, "yyyy-MM-dd HH:mm:ss")
            str��Ժ���� = LeftDB(rsInfo!��Ժ����, 20)
            str��Ժ���� = "10"
            str��Ժ������ = LeftDB(rsInfo!����Ա����, 20)
            str��Ժ��� = LeftDB(��ȡ���Ժ���(lng����ID, lng��ҳID, False, False), 100)
        Else
            '������ԺҲ��������������ܲ���û�г�Ժ��Ϣ
            str��Ժʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            str��Ժ���� = LeftDB("������Ժ", 20)
            str��Ժ���� = "10"
            str��Ժ������ = LeftDB(gstrUserName, 20)
            str��Ժ��� = "������Ժ"
        End If
        mstrErr = Space(4)
        Call yh_ReLeaveHosInfo(mstr˳���, str��Ժԭ��, str��Ժʱ��, str��Ժ���, str��Ժ������, str��Ժ����, str��Ժ����, mstrErr)
        Call DebugTool("����ID=" & lng����ID & "|��ҳID=" & lng��ҳID & "|��Ժʱ��=" & str��Ժʱ��)
    Else
        strMsg = "������δ�����,���ܰ���ҽ����Ժ��"
        If Not bln���ʳ�Ժ Then
            strMsg = strMsg & "���ν�����HIS��Ժ"
        Else
            strMsg = strMsg & "���ڱ����ʻ���Ϊ�ò��˰������Ժ�Ǽ�"
        End If
        MsgBox strMsg, vbInformation, gstrSysName
        If bln���ʳ�Ժ Then Exit Function
    End If
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        'ҽ�����ݿ�ع�
        Call yh_transaction("2", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
    Dim str����� As String   '������ƺ�
    Dim str˳��� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    '�������δ����ã��������HIS��Ժ������ͬʱ����ҽ����HIS��Ժ
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        '��Ժ�Ǽ���ͨ�����ý��㽻����ɡ���ʱ���財�˵ķ����Ѿ�ȫ������
        gstrSQL = "Select ֧��˳��� From ���ս����¼ Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
        Call OpenRecordset(rsTemp, "������Ժ")
        If rsTemp.EOF = True Then
            MsgBox "�ò���δ����ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        
        str˳��� = Nvl(rsTemp("֧��˳���"), "")
        mstrErr = Space(4)
        Call yh_recedefeebalance(str˳���, "1", "", String(18, "1"), mstrErr) 'Ŀǰ������Ԥ�����ڴ���
        mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, ByVal intinsure As Integer) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim str����� As String   '������ƺ�
    Dim cn�ϴ� As New ADODB.Connection, str�������� As String
    Dim cur�����ʻ� As Currency, cur�Ը��ܶ� As Currency
    Dim cur�Ը����� As Double, cur�Ը���� As Double, cur������� As Double
    Dim strҽ�� As String, str���� As String, str��� As String, str���� As String
    Dim cur�������� As Currency, dbl��� As Double, dbl���� As Double
    Dim str����ʱ�� As String, str������Ժʱ�� As String, str���� As String, str�Ǽ�ʱ�� As String
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
    End With
    
    'ȡ������Ժʱ��
    gstrSQL = " Select To_Char(��Ժ����,'yyyy-MM-dd hh24:mi:ss') ��Ժʱ�� From ������ҳ" & _
              " Where ����ID=" & lng����ID & " And ��ҳID=" & g��������.��ҳID
    Call OpenRecordset(rsTemp, "��ȡ������Ժʱ��")
    str������Ժʱ�� = rsTemp!��Ժʱ��
    
    '������һ�����Ӵ����Դﵽ���ܵ�ǰ��������Ŀ���
    Set cn�ϴ� = GetNewConnection
    
    '˳���ȡ��Ժ�Ǽ���֤���ص�
    gstrSQL = "Select ҽ����,˳��� From �����ʻ� " & _
              "Where ˳��� is Not NULL And ����ID=" & lng����ID & " And ����=" & intinsure
    Call OpenRecordset(rsTemp, "�������")
    
    If rsTemp.EOF Then
        MsgBox "δ���ָò��˵�סԺ����˳���,����ִ��ҽ�����ף�", vbExclamation, gstrSysName
        Exit Function
    End If
    mstrҽ���� = rsTemp("ҽ����")
    mstr˳��� = rsTemp("˳���")
    
    'ɾ��ǰ�÷�����������δ����ϸ
    mstrErr = Space(4)
    Call yh_rollbackdetail(mstr˳���, mstrErr)
       mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
            
    str����� = Get�����
    If str����� = "" Then
        Exit Function
    End If
    
    'Ϊ�˱��⸺��¼��ǰ��������¼�ں󡣲�����Ч����
    rsExse.Sort = "NO,��� asc,���� Desc"
    
    '������ϸ����
    Do Until rsExse.EOF
        '����ҽ��ȫ�����´�
        strҽ�� = LeftDB(IIf(IsNull(rsExse("ҽ��")), UserInfo.����, rsExse("ҽ��")), 8)
        str��� = LeftDB(IIf(IsNull(rsExse("���")), "�޹��", rsExse("���")), 30)
        str���� = LeftDB(IIf(IsNull(rsExse("����")), "", rsExse("����")), 30)
        str���� = LeftDB(IIf(IsNull(rsExse("��������")), UserInfo.����, rsExse("��������")), 24)
        '���ܴ��ݸ���
        If rsExse("��¼״̬") = 1 And rsExse("����") < 0 Then
            MsgBox "ҽ����֧��ֱ��¼�븺����ֻ��ѡ��ԭ�е��ݽ��г�����", vbInformation, gstrSysName
            Exit Function
        End If
        '��0��Ŀ����Ϊ��ɾ���Ѿ��ϴ����������ķ��ü�¼
        dbl���� = Val(IIf(rsExse("����") > 0, rsExse("����"), 0))
        dbl��� = Val(IIf(rsExse("�۸�") > 0, rsExse("�۸�"), 0))
        str����ʱ�� = Format(rsExse("����ʱ��"), "yyyy-MM-dd HH:mm:ss")
        str�Ǽ�ʱ�� = Format(rsExse("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
        str���� = Get�������(rsExse!�շ�ϸĿID, intinsure)
        mstrErr = Space(4)
        
        'Ϊ���ø���¼����ȷ�ҵ�����¼���������������в�������¼״̬
        str�������� = rsExse("NO") & "_" & rsExse("���") & "_" & rsExse("��¼����") '& "_" & rsExse("��¼״̬")
        
        '����Ǽ�ʱ��С�ڱ���סԺʱ�����ϴ�
        If str����ʱ�� >= str������Ժʱ�� Then
            Call DebugTool("�ϴ���ϸ����" & mstr˳��� & "," & str�������� & "," & str���� & "," & rsExse("ҽ����Ŀ����") & "," & _
                rsExse("�շ�����") & "," & dbl���� & "," & dbl��� & "," & str���� & "," & str��� & "," & "" & "," & strҽ�� & "," & str���� & "," & str����� & "," & strҽ�� & "," & str�Ǽ�ʱ�� & "," & _
                cur�Ը����� & "," & cur�Ը���� & "," & cur�������)
            Call yh_feedetailtrans(mstr˳���, str��������, str����, rsExse("ҽ����Ŀ����"), _
                rsExse("�շ�����"), dbl����, dbl���, str����, str���, "", strҽ��, str����, str�����, strҽ��, str�Ǽ�ʱ��, _
                cur�Ը�����, cur�Ը����, cur�������, mstrErr)
            If mstrErr <> "0000" Then
                MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
                'ҽ�����ݿ�ع�
                Call yh_transaction("1", mstr˳���, str�����, "0", mstrErr)
                Exit Function
            End If
            cur�������� = cur�������� + rsExse("���")
            
            '��ȡ�÷��ü�¼��ID
            gstrSQL = "Select ID From סԺ���ü�¼ " & _
                " Where NO='" & rsExse!NO & "' And ��¼����=" & rsExse!��¼���� & " And ��¼״̬=" & rsExse!��¼״̬ & " And ���=" & rsExse!���
            Call OpenRecordset(rsTemp, "��ȡ�÷��ü�¼��ID")
            
            If rsTemp.RecordCount <> 0 Then
                gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rsTemp("ID") & "," & cur������� & ",'" & cur�Ը����� & "|" & cur�Ը���� & "|" & cur������� & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
        '����9�汾���Զ�����û��ʱ��������(2007-12-06)
        Else
            Call DebugTool("�ϴ���ϸ����" & mstr˳��� & "," & str�������� & "," & str���� & "," & rsExse("ҽ����Ŀ����") & "," & _
                rsExse("�շ�����") & "," & dbl���� & "," & dbl��� & "," & str���� & "," & str��� & "," & "" & "," & strҽ�� & "," & str���� & "," & str����� & "," & strҽ�� & "," & str�Ǽ�ʱ�� & "," & _
                cur�Ը����� & "," & cur�Ը���� & "," & cur�������)
            Call yh_feedetailtrans(mstr˳���, str��������, str����, rsExse("ҽ����Ŀ����"), _
                rsExse("�շ�����"), dbl����, dbl���, str����, str���, "", strҽ��, str����, str�����, strҽ��, str�Ǽ�ʱ��, _
                cur�Ը�����, cur�Ը����, cur�������, mstrErr)
            If mstrErr <> "0000" Then
                MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
                'ҽ�����ݿ�ع�
                Call yh_transaction("1", mstr˳���, str�����, "0", mstrErr)
                Exit Function
            End If
            cur�������� = cur�������� + rsExse("���")

            '��ȡ�÷��ü�¼��ID
            gstrSQL = "Select ID From סԺ���ü�¼ " & _
                " Where NO='" & rsExse!NO & "' And ��¼����=" & rsExse!��¼���� & " And ��¼״̬=" & rsExse!��¼״̬ & " And ���=" & rsExse!���
            Call OpenRecordset(rsTemp, "��ȡ�÷��ü�¼��ID")

            If rsTemp.RecordCount <> 0 Then
                gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rsTemp("ID") & "," & cur������� & ",'" & cur�Ը����� & "|" & cur�Ը���� & "|" & cur������� & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
        End If
        rsExse.MoveNext
    Loop
        
    str����� = Get�����
    If str����� = "" Then
        Exit Function
    End If
    '�������
    Dim str�����־ As String
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, str��ʼ������ As String
    Dim cur������Ա�Ը� As Double, cur������Աͳ�� As Double, cur����Աͳ�� As Double, str��Ա״̬ As String, cur����ҹ����� As Double
    
    str��ʼ������ = Space(4)
    mstrErr = Space(4)
    str�����־ = "0" '�������
    Call yh_virtualbalance(mstr˳���, str�����־, "", str�����, curȫ�Ը�, cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, _
        cur�����Ը�, cur��ͳ��, cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, cur����Աͳ��, str��Ա״̬, str��ʼ������, cur����ҹ�����, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    '������ʱ���ݣ�Ϊ���������׼��
    'Modified By ZYB 20030812
'    gstrSQL = "Select �ʻ���� From �����ʻ� Where ����=" & intInsure & " And ����ID=" & lng����ID
'    Call OpenRecordset(rsTemp, "��ȡҽ�����˵�ҽ����")
'    cur�����ʻ� = rsTemp!�ʻ����
    cur�����ʻ� = �������_����(lng����ID, mstrҽ����, 10, intinsure)
    If cur������Աͳ�� > 0 Then
        cur�Ը��ܶ� = cur������Ա�Ը�
    Else
        cur�Ը��ܶ� = cur�������� - (curͳ��֧�� + cur��ͳ�� + cur����Աͳ�� + cur������Աͳ��)
    End If
    cur�����ʻ� = IIf(CDbl(Format(cur�����ʻ�, "#####0.00")) >= CDbl(Format(cur�Ը��ܶ�, "#####0.00")), cur�Ը��ܶ�, cur�����ʻ�)
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then cur�����ʻ� = 0
    
    With g��������
        .����ID = lng����ID
        .�������ý�� = cur��������
    End With
    
    סԺ�������_���� = "�����ʻ�;" & cur�����ʻ� & ";1" '�����޸�
    סԺ�������_���� = סԺ�������_���� & "|ҽ������;" & curͳ��֧�� & ";0"
    סԺ�������_���� = סԺ�������_���� & "|��ͳ��;" & cur��ͳ�� & ";0"
    סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & cur����Աͳ�� & ";0"
    סԺ�������_���� = סԺ�������_���� & "|���ⲹ��;" & cur������Աͳ�� & ";0"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, str˳��� As String, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    Dim str����� As String   '������ƺ�
    Dim str������ As String, strҽ�� As String
    Dim str�����־ As String, strSelfNo As String
    Dim cur�����ʻ� As Currency
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, str��ʼ������ As String, str��Ա״̬ As String, cur����ҹ����� As Double
    Dim cur������Ա�Ը� As Double, cur������Աͳ�� As Double, cur����Աͳ�� As Double
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date, lng����ID As Long, rsTemp As New ADODB.Recordset
    
    Dim str���� As String, STR���� As String, str�Ա� As String, str���֤�� As String, lng���� As Double, str��Ⱥ��� As String
    str��ʼ������ = Space(4)
    On Error GoTo errHandle
    '��鲡���Ƿ��Ѿ���Ժ�����������Ժ״̬��ֱ���˳�ϵͳ(20071101)
    gstrSQL = "select a.��Ժ���� " & _
              " from ������ҳ a,������Ϣ b " & _
              " Where a.����ID = b.����ID And a.��ҳID = b.סԺ���� " & _
              " and a.��Ժ���� is null and a.����id = " & lng����ID
    Call OpenRecordset(rsTemp, "�жϳ�Ժ����")
    If rsTemp.RecordCount <> 0 Then
        Err.Raise 9000, gstrSysName, "����δ��Ժ�����ܰ�����㣬���Ȱ����Ժ��"
        Exit Function
    End If
    'ȡ��Ժ�Ǽ���֤�����ص�˳���
    mstr˳��� = str˳���
    str����� = Get�����
    If str����� = "" Then
        Exit Function
    End If
    '���ý���:���ʡ�Ϊ�˴ﵽ��;���ʵ�Ŀ�ģ�û��ʹ�ý��㺯��
    '�ȶ�ȡҽ����
    gstrSQL = "Select ҽ���� From �����ʻ� Where ����=" & intinsure & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡҽ�����˵�ҽ����")
    strSelfNo = rsTemp!ҽ����
    '��ȡ���θ����ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & intinsure & _
        " And A.���㷽ʽ in ('�����ʻ�') And A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ���θ����ʻ�֧����")
    cur�����ʻ� = 0
    If Not rsTemp.EOF Then
        cur�����ʻ� = rsTemp!�����ʻ�
    End If
    
    mstrErr = Space(4)
    str�����־ = "1"   '����
    Call yh_virtualbalance(mstr˳���, str�����־, lng����ID, str�����, curȫ�Ը�, cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, _
        cur�����Ը�, cur��ͳ��, cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, cur����Աͳ��, str��Ա״̬, str��ʼ������, cur����ҹ�����, mstrErr)
    mstrErr = TrimStr(mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        'ҽ�����ݿ�ع�
        Call yh_transaction("2", mstr˳���, str�����, "0", mstrErr)
        Exit Function
    End If
    Call yh_transaction("2", mstr˳���, str�����, "1", mstrErr)
    '��д�����
    curDate = zlDatabase.Currentdate
    '�����ò��˱��ν���Ĳ�����Ϣ
    gstrSQL = "Select nvl(����ID,0) ����ID From �����ʻ� A Where A.����=" & intinsure & " and A.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���ս���")
    If rsTemp.EOF = False Then
        lng����ID = rsTemp("����ID")
    End If
    'дIC����������������һ�������д�����˿�дʧ��ʱ����Ȼ��������
    str������ = Left(strSelfNo, 1)
    str��ʼ������ = Space(4)
    strҽ�� = LeftDB(UserInfo.����, 8)
    If CDbl(cur�����ʻ�) <> 0 Then
        '������Ҫ���¿�ǰ�������cardinfo()
'        str���� = Space(18)
'        str���� = Space(60)
'        str�Ա� = Space(3)
'        str���֤�� = Space(20)
'        mstrErr = Space(4)
'        Call yh_cardinfo(str������, mcur�ʻ����, str����, str����, str�Ա�, str���֤��, lng����, str��Ⱥ���, mstrErr)
'       If Trim(mstrErr) <> "0000" Then
'           MsgBox "���ν���ò��˲��ÿ�֧��!!!!", vbOKOnly
'       Else
        mstrErr = Space(4)
        Call yh_cardpay(str������, mstr˳���, strҽ��, "סԺ����", CDbl(cur�����ʻ�), str��ʼ������, mstrErr)
        mstrErr = TrimStr(mstrErr)
        If mstrErr <> "0000" Then
            Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure) & "������¿�ʧ��,�벹���ֽ�" & Format(cur�����ʻ�, "#####0.00") & "��"
            cur�����ʻ� = 0
        End If
       'End If
    End If
    
    '��ʱ��ȡ����
    Dim cur���� As Double
    gstrSQL = "select nvl(����,0) as ���� from �����ʻ� where ˳���='" & mstr˳��� & "' and ����=" & intinsure & " and  ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ����")
    cur���� = strVal(rsTemp!����)
    '���±��ս����¼:
    If mstrErr <> "0000" Then
       cur�����ʻ� = 0
       '������Ա�Ը�"cur������Ա�Ը�"��Ϊ"cur�����Ը�"
      gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & Get���ֱ���(lng����ID) & "," & cur�����Ը� & "," & _
        g��������.�������ý�� & "," & curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ� & ",'" & mstr˳��� & "'," & g��������.��ҳID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    Else
      gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & Get���ֱ���(lng����ID) & "," & cur�����Ը� & "," & _
        g��������.�������ý�� & "," & curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ� & ",'" & mstr˳��� & "'," & g��������.��ҳID & ")"
      Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End If
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
    '���� curͳ���ۼ� ������Ŀ����Ϊ�˵���API�����ͼ���
    Dim curͳ���ۼ� As Double, cur����ͳ���޶� As Double, cur���ͳ���޶� As Double
    Dim cur�����ۼ� As Double, str������Ϣ As String, str������־���� As String, str��ҩ���� As String
    '������ҽ��֧�ֲ�ѯ֧���ۼ�
    mstrErr = Space(4)
    Call yh_RyspInfo(mstr˳���, cur����, curͳ���ۼ�, cur����ͳ���޶�, cur���ͳ���޶�, cur�����ۼ�, str������Ϣ, str������־����, str��ҩ����, mstrErr)
    curͳ�ﱨ���ۼ� = curͳ���ۼ�
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & cur�����ۼ� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    '����������Ϣ
'    gstrSQL = "zl_����������Ϣ_insert(" & lng����ID & "," & intInsure & "," & Year(curDate) & "," & _
'        mstr˳��� & "," & cur���� & "," & curͳ���ۼ� & "," & _
'        cur����ͳ���޶� & "," & cur���ͳ���޶� & "," & cur�����ۼ� & "," & str������Ϣ & ",'" & str������־���� & "',null)"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "����������Ϣ")
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    סԺ����_���� = True
    '�ж��Ƿ���Ҫ���ó�Ժ���㣨���HIS�ѳ�Ժ�Ҳ�����δ����ã�
    Dim lng��ҳID As Long
    'ȡ����ҳID
    gstrSQL = "Select Nvl(סԺ����,0) ��ҳID From ������Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "ȡ��ҳID")
    lng��ҳID = rsTemp!��ҳID
    
    If Not ����δ�����(lng����ID, lng��ҳID) And ҽ�������Ѿ���Ժ(lng����ID) Then
        gstrSQL = "Select A.��Ժ����,A.��Ժ����,Decode(A.��Ժ��ʽ,'����',0,'����',1,'תԺ',2,9) as ��Ժ��ʽ,B.����,D.סԺ��,Sysdate as ����ʱ��," & _
                " C.����,C.ҽ����,C.����,C.˳��� " & _
                " From ������ҳ A,���ű� B,�����ʻ� C,������Ϣ D " & _
                " Where A.����ID=D.����ID And A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
                " And A.��Ժ����ID=B.ID And A.����ID=C.����ID And C.����=" & intinsure
        Call OpenRecordset(rsTemp, "ȡ˳���")
    
        If rsTemp.EOF Then
            Err.Raise 9000 + vbExclamation, gstrSysName, "û�д˲��˻�˲��˲���ҽ�����ˣ��޷������Ժ������������ҽ���ʻ��а������Ժ������"
            Exit Function
        End If
        If IsNull(rsTemp!˳���) Then
            Err.Raise 9000, gstrSysName, "δ���ָò��˵�סԺ����˳���,�޷������Ժ������������ҽ���ʻ��а������Ժ������"
            Exit Function
        End If
        
        Call ��Ժ�Ǽ�_����(lng����ID, lng��ҳID, rsTemp!˳���, intinsure, True)
    End If
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_����(lng����ID As Long, ByVal intinsure As Integer) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str˳��� As String, cur�����ʻ� As Currency
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curDate As Date
    Dim str������Ⱥ��־ As String
    
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & intinsure & " and ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    str˳��� = rsTemp("֧��˳���")
    str������Ⱥ��־ = Nvl(rsTemp("������Ⱥ��־"), 0)
'    mstrErr = Space(4)
'    Call yh_recedefeebalance(str˳���, "1", lng����ID, String(18, "1"), mstrErr) '1��ʾ���ѽ���
'         mstrErr = Trim(mstrErr)
'    If mstrErr <> "0000" Then
'       Err.Raise 9000,gstrSysName, GetErrInfo(mstrErr, intInsure)
'        Exit Function
'    End If
    mstrErr = Space(4)
    Call yh_recedefeebalance(str˳���, "0", lng����ID, String(18, "1"), mstrErr) 'Ŀǰ������Ԥ�����ڴ���,��ʾ0����Ԥ����
       mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        Exit Function
    End If
    
    '�������ʻ�֧������˵����ţ������Ǳ���סԺ��֧��ȫ�����ˣ�
    cur�����ʻ� = Nvl(rsTemp("�����ʻ�֧��"), 0)
    If CDbl(cur�����ʻ�) <> 0 Then
        mstrErr = Space(4)
        Call yh_recedefeebalance(str˳���, "4", lng����ID, String(18, "1"), mstrErr) 'Ŀǰ������Ԥ�����ڴ���
        mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            Err.Raise 9000, gstrSysName, "�����ʻ�����ʧ�ܵ�������˳ɹ�������HIS����ϵ��" & vbCrLf & "��ϸ����" & GetErrInfo(mstrErr, intinsure)
        End If
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(intinsure, rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    '�ⶥ�߱����м������룬���Բ�ȡ��
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & rsTemp("����ID") & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & rsTemp("����") * -1 & "," & rsTemp("�ⶥ��") & "," & _
        rsTemp("ʵ������") * -1 & "," & rsTemp("�������ý��") * -1 & "," & rsTemp("ȫ�Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & "," & rsTemp("���Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & _
        cur�����ʻ� * -1 & ",'" & str˳��� & "'," & rsTemp("��ҳID") & ",null,null,'" & str������Ⱥ��־ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")

    סԺ�������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ������Ϣ_����(ByVal lngErrCode As Long) As String
'���ܣ����ݴ���ŷ��ش�����Ϣ

End Function

Private Function LeftDB(ByVal strText As String, ByVal lngLength As Long)
'���ܣ������ݿ�ĳ��ȼ��㷽ʽ�õ��ַ�����ʵ�ʿ����Ӵ�
    LeftDB = StrConv(MidB(StrConv(strText, vbFromUnicode), 1, lngLength), vbUnicode)
End Function
Private Function strVal(ByVal strText As String)
'���ַ��ͻ�Ϊ������
       strVal = Val(strText)
End Function

Private Function Get�����() As String
    Dim str����� As String
    
    On Error GoTo errHandle
    
    str����� = Space(18)
    Call yh_gettranssequence(str�����) '������ô��ݺͽ��������������
    str����� = TrimStr(str�����)
    If str����� = "" Then
        MsgBox "��ȡ������ƺ�ʧ�ܡ�", vbInformation, gstrSysName
    End If
    
    Get����� = str�����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Is����ȷ(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ��ж϶������Ŀ��Ƿ����Ҫ�����Ĳ��˵�
    Dim rsTemp As New ADODB.Recordset
    Dim str����_�� As String, str���� As String, str������ As String
    
    Dim cur��� As Double, STR���� As String, str�Ա� As String
    Dim str���֤�� As String, lng���� As Double, str��Ⱥ��� As String
    
    On Error GoTo errHandle
    
    gstrSQL = "select ����,ҽ���� from �����ʻ� where ����=" & intinsure & " and ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    str����_�� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    str������ = Left(rsTemp("ҽ����"), 1)
    
    str���� = Space(20)
    STR���� = Space(60)
    str�Ա� = Space(3)
    str���֤�� = Space(20)
    
    mstrErr = Space(4)
    Call yh_cardinfo(str������, cur���, str����, STR����, str�Ա�, str���֤��, lng����, str��Ⱥ���, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
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

Private Function Get�����(ByVal strҽ���� As String, ����� As Currency, ByVal intinsure As Integer) As Boolean
'���ܣ��õ������
    Dim cur��� As Double, STR���� As String, str�Ա� As String, str���� As String
    Dim str���֤�� As String, lng���� As Double, str������ As String, str��Ⱥ��� As String
    
    str������ = Left(strҽ����, 1)
    
    str���� = Space(20)
    STR���� = Space(60)
    str�Ա� = Space(3)
    str���֤�� = Space(20)
    
    mstrErr = Space(4)
    Call yh_cardinfo(str������, cur���, str����, STR����, str�Ա�, str���֤��, lng����, str��Ⱥ���, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
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
    Call OpenRecordset(rsTemp, "����ҽ��")
    
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

Public Function ��������Ǽ�_����(ByVal str˳��� As String, ByVal intinsure As Integer) As Boolean
'���ܣ���������Ǽ�
    Dim rsTemp As New ADODB.Recordset
    Dim str����� As String   '������ƺ�
    Dim strҽ�� As String, str���� As String
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, str��ʼ������ As String
    Dim cur������Ա�Ը� As Double, cur������Աͳ�� As Double, cur����Աͳ�� As Double, str��Ա״̬ As String, cur����ҹ�֧������ As Double
    
    '���������(ʡ��ҽ������һ������)
    Dim strSicks As String          '�������صĲ����б�ע���������ⲡĿǰֻ��ѡ��Ѫ͸
    Dim strFeeBalanceType As String, dblXTCS As Double      '���㺯���ķ��س��Σ�֧���������Ѫ͸����
    Dim strSickSel As String        '����Աѡ��Ĳ��ֱ���
    Dim str��ע As String
    
    On Error GoTo errHandle
    str��ʼ������ = Space(4)
    
    gstrSQL = "Select ֧��˳��� from ���ս����¼ where ֧��˳���='" & str˳��� & "'"
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    If rsTemp.EOF = False Then
        MsgBox "�ò��˵ļ��ｻ���Ѿ��ɹ���ɣ����ܳ�����ֻ�����ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ɾ��ǰ�÷�����������δ����ϸ
    mstrErr = Space(4)
    Call yh_rollbackdetail(str˳���, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    '��Ժ�Ǽ���ͨ�����ý��㽻����ɡ�����ý���
    str����� = Get�����
    If str����� = "" Then
        
    End If
    
    mstrErr = Space(4)
    Call yh_feebalance(mstr˳���, strҽ��, str����, strSickSel, str�����, _
            curȫ�Ը�, cur�ҹ��Ը�, curͳ��֧��, curͳ���Ը�, cur�����Ը�, cur�����Ը�, cur��ͳ��, _
            cur���Ը�, cur������Ա�Ը�, cur������Աͳ��, cur����Աͳ��, str��Ա״̬, str��ʼ������, cur����ҹ�֧������, _
            strFeeBalanceType, dblXTCS, mstrErr)
       mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        'ҽ�����ݿ�ع�
        Call yh_transaction("2", str˳���, str�����, "0", mstrErr)
        Exit Function
    End If
    
    ��������Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get�������(ByVal str��Ŀ���� As String, ByVal intinsure As Integer) As String
    Dim rsClass As New ADODB.Recordset
    '��ȡĳ��ҽ����Ŀ�Ĵ������
    gstrSQL = "Select ��ע From ����֧����Ŀ Where ����=" & intinsure & " And �շ�ϸĿID=" & str��Ŀ����
    Call OpenRecordset(rsClass, "��ȡĳ��ҽ����Ŀ�Ĵ������")
    If rsClass.RecordCount = 0 Then
        MsgBox "ҽ������Ϊ��" & str��Ŀ���� & "����Ŀ�ڱ�����Ŀ���в����ڣ�", vbInformation, gstrSysName
        Exit Function
    End If
    Get������� = Nvl(rsClass!��ע)
End Function
