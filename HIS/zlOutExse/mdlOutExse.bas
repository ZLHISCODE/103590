Attribute VB_Name = "mdlOutExse"
Option Explicit 'Ҫ���������
'=======ϵͳ������ر���============
Public Enum �����֤Enum
    id�����շ� = 0
    id��Ժ�Ǽ� = 1
    id�ʻ����� = 2
    id�Һ� = 3
    id���� = 4
    id����ȷ�� = 5
End Enum

Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    'support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29      '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    support�൥���շ� = 30          '�Ƿ�֧�ֶ൥���շ�
    
    support�����շѴ�Ϊ���۵� = 31  '�������շѵ�תΪ���۵����棬�޸���ǰ�̶��ж�ĳ��ҽ���ķ�ʽ
    
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    support�൥���շѱ���ȫ�� = 39  '�൥���շѱ���ȫ��
    
    supportҽ���ӿڴ�ӡƱ�� = 46    'HIS��ֻ��Ʊ�ݺŵ�������ӡ��ҽ���ӿ�(����)�д�ӡ
    support�൥��һ�ν��� = 47      '�൥��Ԥ����ʱ��ҽ���ӿڽ������һ�ε���ʱ���ؽ�������HIS���ٷ�̯��ÿ�ŵ�����
    
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    supportҽ��ȷ���������� = 48
    supportʵʱ��� = 60             '�Ƿ����÷���ʵʱ���
    '���˺�:27536 20100119
    support�����ѽɿ���� = 64            '���շ�ʱ,����շѲ�����"�����нɿ�������ۼƿ���"Ϊtrueʱ,ͬʱ��ҽ������ʱû������ɿ���ʱ�������û�
    support�˷Ѻ��ӡ�ص� = 65   'ҽ�������Ƿ��˷Ѻ��ӡ�ص�:����
    support����_���ֵ��ݽ��� = 80               'Ԥ���㡢���㶼ֻ����һ��ҽ������:һ��ͨͬ������
    
    support�ҺŲ���ȡ������ = 81    '�ڹҺ�ʱ����ʹ��ҽ����ȡ������

    support������ȫ�� = 82 '�����˷�ʱ�������ݽ����˷ѣ�86176
    support�൥�ݷֵ��ݽ��� = 83 '�൥��һ�ν��㰴���ݽ���ҽ��������86321
    supportһ�ν���ֵ����˷� = 85 '��һ�ν������ҽ���ӿڣ����������˷�,91602
End Enum
Public Type Ty_InsurePatiPara
    ��������ҽ����Ŀ As Boolean
    �����շѴ�Ϊ���۵� As Boolean
    �����ѽɿ���� As Boolean
    ������봫����ϸ As Boolean
    ҽ���ӿڴ�ӡƱ�� As Boolean
    ҽ��ȷ���������� As Boolean
    �൥��һ�ν��� As Boolean
    ���������շ� As Boolean
    ����Ԥ���� As Boolean
    �൥���շ� As Boolean
    �ֱҴ��� As Boolean
    ʵʱ��� As Boolean
    ���Ը� As Boolean
    ȫ�Ը� As Boolean
    blnOnlyBjYb As Boolean '���ؽ�֧�ֱ���ҽ��:���˺�
    �˷Ѻ��ӡ�ص� As Boolean '
End Type
Public Enum ����Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum
Public gobjCustBill As Object               '�Զ�����ʵ�����
Public gobjRegist As Object                 '�ҺŲ���
Public gobjPatient As Object                 '���˹���ϵͳ����
Public gclsInsure As New clsInsure
Public gbytWarn As Byte '���ʱ�������ֵ
Public gstrModiNO As String '�޸ĺ�������µ��ݺ�
'============����ϵͳ����=====================

Public gstrҽ���������� As String 'ҽ����������ķ�������
Public gstr���ѷ������� As String '���Ѳ�������ķ�������
Public gstrCustomerAppellation As String    '�������ߵĳƺ�:����,�ͻ�
Public gbln�˷�����ģʽ As Boolean '�˷��Ƿ�ʹ���������ģʽ
Public gbln�����л� As Boolean '35242
'�������뷽ʽ
Public gblnInputName As Boolean '������������
Public gblnInputID As Boolean '������ID����
Public gblnInputCard As Boolean '����ˢ������
Public gblnInputNO As Boolean '����Һŵ�����
Public gblnUnPopPriceBill As Boolean '���������۵�ѡ��
Public gobjSquare As SquareCard  '�����㲿��  42301
'���￨
'Public gbytCardNOLen As Byte '���￨�ų���
'Public gblnShowCard As Boolean '�Ƿ���￨����ʾΪ��������
Public gobjPublicExpense As Object
Public gintPriceGradeStartType As Integer
Public gstrҩƷ�۸�ȼ� As String
Public gstr���ļ۸�ȼ� As String
Public gstr��ͨ�۸�ȼ� As String

Public gdblԤ��������鿨 As Double 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
Public gbytԤ����˷��鿨 As Byte 'Ԥ����˷�ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
Public gbln���ѿ��˷��鿨 As Boolean '���ѿ��˷�ʱ�Ƿ�ˢ����֤

'Ʊ�ݿ���
Public gobjBillPrint As Object '������Ʊ�ݴ�ӡ����
Public gblnBillPrint As Boolean '������Ʊ�ݴ�ӡ�����Ƿ����

Public gblnStrictCtrl As Boolean '�Ƿ��ϸ�Ʊ�ݹ���
Public gbytFactLength As Byte 'Ʊ�ݺ��볤��
Public gblnSharedInvoice As Boolean
Public glngFactNormal As Long       '��ͨ��Ʊ��ʽ
Public glngFactMediCare As Long     'ҽ����Ʊ��ʽ

'ҩ����������ؿ���
Public gblnStock As Boolean 'ָ��ҩ��ʱ�Ƿ��޶�����ҩƷ�Ŀ��
Public gbytAssign As Byte '��ҩ���ڶ�̬���䷽ʽ(0,1)
Public gbln�շѺ��Զ���ҩ As Boolean '�Ƿ��Զ���ҩ��ҩ
Public gbln�����Զ����� As Boolean '�����շѻ����,���ʻ��۵���˺��Զ�����
Public gbytMediOutMode As Byte '����ҩƷ���ⷽʽ��0-�������Ƚ��ȳ���1-��Ч������ȳ�,Ч����ͬ�����ٰ������Ƚ��ȳ�
Public gbln���������ɿ� As Boolean '��ȡ���۵����Ƿ������ɿ�:39253
Public gbln����ʾ�޿������ As Boolean
'�����������
Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gblnMyStyle As Boolean 'ʹ�ø��Ի����

Public gbln�շ���� As Boolean '�Ƿ������������
Public gblnFeeKindCode As Boolean '�������ʱ,��λ�����շ�������
Public gstrMatchMode As String  '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
Public gbln�������� As Boolean '�����������,�������ʱѡ��������,�򱣴�ʱ,���ټ��
Public gblnPrePayPriority As Boolean '����ʹ��Ԥ����
Public gbytAutoSplitBill As Byte '���ݰ�����ִ�п����Զ�����
Public gbytҽ�������� As Byte '0-�����м�顢1-��鲢����δ������Ŀ��2-��鲢��ֹδ������Ŀ
Public gcurMaxMoney As Currency '���ʷ���������ѽ��

'��ҩ������
Public grsABCNum As ADODB.Recordset
Public gstrABC As String '��������Ŀ����ĸ
'���˺� ����:????    ����:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '����С������
Public gstrFeePrecisionFmt As String '����С����ʽ:0.00000

'���ü������
Public gBytMoney As Byte '�շѷֱҴ�����
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"
Public gbln��������ۿ� As Boolean '������Ŀ���ܼ����ۿ�
Public gblnMultiBalance As Boolean  '���ŵ���ʹ�ö��ֽ��㷽ʽģʽ
Public glngAddedItem As Long    '�Զ����չҺŷѵ��շ���ĿID

'��������
Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gbln�����������۷��� As Boolean '���ʱ����������۷���
 

'==============���ñ�������===============
'Ʊ�ݿ���
Public gobjTax As Object '˰�ش�ӡ�ӿڶ���
Public gblnTax As Boolean '�����Ƿ�ʹ��˰�ش�ӡ
Public gstrTax As String
Public gint�շ��嵥 As Integer      '0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ
Public gint����֪ͨ�� As Integer    '0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ

'ҩ�������ڿ���
Public glng��ҩ�� As Long 'ָ������ҩ��,0Ϊ��̬����
Public glng��ҩ�� As Long 'ָ������ҩ��,0Ϊ��̬����
Public glng��ҩ�� As Long 'ָ���ĳ�ҩ��,0Ϊ��̬����
Public glng���ϲ��� As Long 'ָ�������ķ��ϲ���,0Ϊ��̬����

Public gstr���� As String  'ָ������ҩ����ҩ����,��Ϊ��̬����
Public gstr�д� As String 'ָ������ҩ����ҩ����,��Ϊ��̬����
Public gstr�ɴ� As String  'ָ���ĳ�ҩ����ҩ����,��Ϊ��̬����
Public gblnҩ���ϰల�� As Boolean     '�Ƿ�������ҩ���ϰల��

Public glng���ϸĿID As Long           '�Ƿ������������
Public gstr�������� As String         '��������(���㷽ʽ�е���������)

Public gstr����վݷ�Ŀ As String
'���뷢ҩʱҪ������ҩ��
Public gstr��ҩ�� As String
Public gstr��ҩ�� As String
Public gstr��ҩ�� As String

Public gbln����ҩ�� As Boolean '�Ƿ���ʾ����ҩ�����
Public gbln����ҩ�� As Boolean '�Ƿ���ʾ����ҩ����
Public gbyt�����ʾ��ʽ As Byte   '����:31936: ����ҩ����ҩ��(�ǲ���Ա��������)�Ŀ����ʾ��ʽ:0-ֱ����ʾ���;1-��ʾ����

'�������
Public gstrCardPass As String 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'
Public gbytCode As Byte '�������ɷ�ʽ��0-ƴ��,1-���,2-����
Public gstr�շ���� As String '��������շ����
Public gblnPay As Boolean '��ҩ�Ƿ����븶��
Public gblnTime As Boolean '�����Ŀ�Ƿ���������
Public gbyt����ҽ�� As Byte '0-������ȷ����������,1-��������ȷ��������,2-�����˺Ϳ��������໥����
Public gbln��ȱʡ������ As Boolean
Public gblnȱʡ�������� As Boolean
Public gbln�����俪���� As Boolean
Public gcurMax As Currency '������������������

Public gstr�ѱ� As String 'ȱʡ�Ĳ��˷ѱ�
Public gstr���㷽ʽ As String  'ȱʡʹ�õĽ��㷽ʽ
Public gstr��������ID As String '�洢��ǰ����Ա����������ID,
Public gbln�Ա� As Boolean '����Ƿ񾭹�����Ŀ
Public gbln���� As Boolean
Public gbln�ѱ� As Boolean
Public gblnҽ�Ƹ��� As Boolean
Public gbln�Ӱ� As Boolean
Public gbln�������� As Boolean
Public gbln������ As Boolean
Public gbyt��������ʾ As Byte

Public gblnSeekName As Boolean '�Ƿ�ͨ����������ģ������
Public gblnOnlyUnitPatient As Boolean '�������ʱ,��������ʱֻ���Լ��λ����
Public gintNameDays As Integer 'ͨ������ģ����������
Public gblnSeekBill As Boolean '�Ƿ��Զ���Ѱ���۵���
Public gintSeekDays As Integer '�Զ���¼���ݵ�����
Public gblnCheckRegeventDept As Boolean '��鲡�˹Һſ���
Public gbytUnRegevent As Byte      'δ�ҺŲ����շ�,0-����,1-����,2-��ֹ
Public gbln����¼������ʹ�õĿ����� As Boolean '92727

'LED�������ۿ���
Public gblnLED As Boolean '�Ƿ�ʹ��Led��ʾ
Public gbln�ֹ����� As Boolean 'ʹ��Led��,�Ƿ��ֹ�����
Public gblnLedDispDetail As Boolean 'ʹ��Led��,ÿ����һ�е����Ƿ����豸����ʾ�շ���ϸ
Public gblnLedWelcome As Boolean    'ʹ��Led��,���շ�ʱ,�����²��˻��뻮�۵�ʱ,�Ƿ���ʾ��ӭ��Ϣ������

'��������
Public gint������Դ As Integer '�շѣ�����ʱ�Ĳ�����Դ(1-����,2-סԺ)
Public gbln������Դ��Ȩ�޿��� As Boolean '�Ƿ�������Ĳ�����Դ

Public gblnҩ����λ As Boolean '����,����,�շ�ʱ�Ƿ������ﵥλ������ʾ������,�շ�Ҳ���ܰ�סԺ��λ
Public gstrҩ����λ As String '���ݲ�����Դ������"���ﵥλ"��"סԺ��λ"
Public gstrҩ����װ As String '���ݲ�����Դ������"�����װ"��"סԺ��װ"

Public gblnCheckTest As Boolean '���Ƥ�Խ��
Public gbln�ۼ� As Boolean '�շ��Ƿ���ʾ�ۼ�
Public gbln��ʿ As Boolean '�շѻ����Ƿ���ʾ��ʿ
Public gint����ϼ� As Integer '0-���վݷ�Ŀ,1-��������Ŀ
Public gblnMulti As Boolean '�Ƿ�֧�ֶ൥���շ�
Public gblnShowErr As Boolean '�鿴�շѵ���ʱ�Ƿ���ʾ������
Public gintDelPrice As Integer '���뻮�۹���ʱ,ɾ��n���ڵĻ��۵�
Public gbln���ʴ�ӡ As Boolean
Public gbln���۴�ӡ As Boolean
Public gbln��˴�ӡ As Boolean

Public Type TY_Reg_Para  '�Һ���ز���
    bytNODaysGeneral As Byte    '��ͨ�Һ���Ч����
    bytNoDayseMergency As Byte '����Һ���Ч����
End Type
'------------------------------------------------------------------------------------------
'��֧�����
Public Type gTY_PayMoney
    lngҽ�ƿ����ID As Long
    bln���ѿ� As Boolean
    str���㷽ʽ As String
    str���� As String
    strˢ������ As String
    strˢ������ As String
    lng���ѿ�ID As Long
    str������� As String   '�ϴ�ˢ���������
    dbl��ˢ��� As Double '�Ѿ�ˢ�����ѿ����
    str������ˮ�� As String
    str����˵�� As String
    bln���� As Boolean
    bln��������  As Boolean
    intҽ�ƿ����� As Integer
    bln֧Ʊ As Boolean
    bln���ƿ� As Boolean
    blnOneCard As Boolean '�Ƿ�һ��ͨ����
    int���� As Integer '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����;<0 ��ʾ������֧��
    strNo As String
    lngID As Long 'Ԥ��ID
    lng����ID As Long
    dbl�ʻ���� As Double
End Type
Public gtyPrePatiPay As gTY_PayMoney '�ϴβ��˵�֧����ʽ
'-------------------------------------------------------------------------------------------------------------------------------------------------
'--����ϵͳ����
'����:27990
Private Type Ty_System_Para
     bytҩƷ������ʾ As Byte   'ҩƷ������ʾ�������浥����ϸ������������桢ֱ�ӽ����ҩƷѡ����ʱ��ҩƷ������ʾ����0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
     byt����ҩƷ��ʾ As Byte  '����ҩƷ��ʾ��ͨ��������뷽ʽ����ѡ����ʱҩƷ���Ƶ���ʾ����0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
     byt��������ʶ����� As Byte   '�Ƿ������ʶ��::1-����ͨ��ɨ��¼���¼������;0-�����ƣ�����ͨ������Ȳ���
     Sy_Reg  As TY_Reg_Para     '�Һ����:34717
End Type
Public gTy_System_Para As Ty_System_Para

'-------------------------------------------------------------------------------------------------------------------------------------------------
'���˺�:ģ���������
Private Type Ty_Module_Para
    byt�ɿ���� As Byte    '�ɿ����:0-�������нɿ�������ۼƿ���,1-��������ɿ��Ž��������ۼ�(�ı䲡�˳���)��2-�շ�ʱ����Ҫ����ɿ���
    int����ʣ��Ʊ������ As Integer      '�շ�ʱ,Ʊ����ʣ��X�ź�ʼ�����շ�Ա:-1��������
    bytƱ�ݷ������ As Byte   '25187:Ʊ�ݷ������:0-����ʵ�ʴ�ӡ����Ʊ��;1-����ϵͳԤ���������;2-�����û��Զ���������
    bytƱ�ݻ������� As Byte            'bytƱ�ݷ������>=1ʱ:��Ʊ�ݷ��������ܵ�����:25187
    blnƱ�ݷֵ��� As Boolean    '�ֵ��� :25187
    intִ�п��� As Integer        'N��ִ�п��ҷ�ҳ:25187
    int�վݷ�Ŀ As Integer        'N���վݷ�Ŀ��ҳ:25187
    int�շ�ϸĿ As Integer        'N���շ�ϸĿ��ҳ:25187
    bln�ֱ��ӡ As Boolean      '���ŵ����շѷֱ��ӡ:25187,�ϲ�
    bytƱ�����ɷ�ʽ As Byte     '0-����Ŀ��ӡ,1-��ϸĿ��ӡ,10-�Ȱ�ִ�п��ҷֱ��,�ٰ���Ŀ��ӡ,11-�Ȱ�ִ�п��ҷֱ��,�ٰ�ϸĿ��ӡ:25187,�ϲ�
    byt�����վ��д� As Byte     '�շ��վ����д�:25187,�ϲ�
    blnһ��Ʊ�� As Boolean '�շ�һ��ֻ��һ��Ʊ��:25187,�ϲ�
    bln������ As Boolean '�Ƿ���ȡ������:25187,�ϲ�
    bln���ռ��Ʊ�� As Boolean  '25187,�ϲ�
    blnʹ�üӼ��л�  As Boolean '47457
    bytҩƷ��ҩ�˷ѷ�ʽ As Byte '47400
    byt�˷�ȱʡѡ��ʽ As Byte '87489
    bytˢ��ȱʡ������ As Byte '86853
    blnֻ��ҽ������ɹ������շ� As Boolean '91665
    str�����շ�ִ�п��� As String '96357����ʽ������ID1,����ID2,����ID3,...
    str�������շ�ִ�п��� As String '�������˵ı����շ�ִ�п��ң���ʽ������ID1,����ID2,����ID3,...
    blnҽ��������ȱʡ��λ As Boolean
    bln�ֽ��˿�ȱʡ��ʽ As Boolean
    bln���ֱ��ӡ As Boolean '104983
    strȱʡ���� As String '112753
End Type

Public gTy_Module_Para As Ty_Module_Para
'-------------------------------------------------------------------------------------------------------------------------------------------------
Public gstrMatchMethod As String
Private mlng���ű���ƽ������ As Long
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public grsTotal As ADODB.Recordset  'ͳ�ƻ��ܵ��ۼƽ��
Public grs�շ���� As ADODB.Recordset
Public gobjPlugIn As Object
Public gobjPublicDrug As Object 'ҩƷ��������,105872
Public gblnUserIsClinic As Boolean

Public Function zlGet�շ����() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ����
    '����:�����շ����
    '����:���˺�
    '����:2013-02-21 17:08:51
    '����:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�Ȼ��浽����
    On Error GoTo errHandle
    gstrSQL = "Select  ����,���� From �շ���Ŀ���"
    If grs�շ���� Is Nothing Then
        Set grs�շ���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ����")
    ElseIf grs�շ����.State <> 1 Then
        Set grs�շ���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ����")
    End If
    Set zlGet�շ���� = grs�շ����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetDefaultWindow(ByVal str��� As String, ByVal lngҩ��ID As Long) As String
'����:��ȡȱʡ��ҩ����������
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    
    Select Case str���
        Case "5"
            If InStr(gstr����, ":") > 0 Then '������û�д�ҩ��ID
                 strTmp = gstr����
            ElseIf glng��ҩ�� > 0 And gstr���� <> "" Then
                strTmp = glng��ҩ�� & ":" & gstr����
            End If
        Case "6"
            If InStr(gstr�д�, ":") > 0 Then
                 strTmp = gstr�д�
            ElseIf glng��ҩ�� > 0 And gstr�д� <> "" Then
                 strTmp = glng��ҩ�� & ":" & gstr�д�
            End If
        Case "7"
            If InStr(gstr�д�, ":") > 0 Then
                 strTmp = gstr�ɴ�
            ElseIf glng��ҩ�� > 0 And gstr�ɴ� <> "" Then
                 strTmp = glng��ҩ�� & ":" & gstr�ɴ�
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str���
                Case "5"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    GetDefaultWindow = strTmp
End Function


Public Function Get��ҩ����(ByVal Curdate As Date, ByVal lngҩ��ID As Long, ByVal str��� As String, _
    str���� As String, str�ɴ� As String, str�д� As String) As String
'���ܣ���ȡҩƷ��Ӧ�ķ�ҩ����
'������lngҩ��ID=ִ�в���ID,curDate=��ǰʱ��
'˵������ͬһ������ҩ���ķ�ҩ������ƽ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ָ��ʱ�̶�����(ָ����ָû�ж�Ӧҩ���ϰ�ʱָ��)
    Select Case str���
        Case "5"
            If str���� <> "" Then
                Get��ҩ���� = str����
            ElseIf glng��ҩ�� > 0 Then
                Get��ҩ���� = GetDefaultWindow(str���, lngҩ��ID)
                str���� = Get��ҩ����
            End If
        Case "6"
            If str�ɴ� <> "" Then
                Get��ҩ���� = str�ɴ�
            ElseIf glng��ҩ�� > 0 Then
                Get��ҩ���� = GetDefaultWindow(str���, lngҩ��ID)
                str�ɴ� = Get��ҩ����
            End If
        Case "7"
            If str�д� <> "" Then
                Get��ҩ���� = str�д�
            ElseIf glng��ҩ�� > 0 Then
                Get��ҩ���� = GetDefaultWindow(str���, lngҩ��ID)
                str�д� = Get��ҩ����
            End If
    End Select
    
    
    If Get��ҩ���� <> "" Then
        strSQL = "Select ���� From ��ҩ���� Where �ϰ��=1 And ҩ��ID=[1] And ����=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngҩ��ID, Get��ҩ����)
        If rsTmp.EOF Then Get��ҩ���� = ""
        Exit Function
    End If
    
    '��̬�����ϰ�ķ�ר�Ҵ���,98876
    strSQL = "Select Zl_Get��ҩ����([1],[2],[3]) As ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ҩ����", lngҩ��ID, gbytAssign, Curdate)
    If Not rsTmp.EOF Then
        Get��ҩ���� = Nvl(rsTmp!����)
    End If
    
    If Get��ҩ���� <> "" Then
        Select Case str���
            Case "5"
                str���� = Get��ҩ����
            Case "6"
                str�ɴ� = Get��ҩ����
            Case "7"
                str�д� = Get��ҩ����
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Public Function Get������() As Detail
'���ܣ���ȡ�����ѵ��շ�ϸĿID
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.*,C.���� as ������� " & _
        " From �շ���ĿĿ¼ A,�շ��ض���Ŀ B,�շ���Ŀ��� C " & _
        " Where B.�շ�ϸĿID=A.ID And C.����=A.��� " & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.�ض���Ŀ='������'"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlOutExse")
    If Not rsTmp.EOF Then
        Set Get������ = New Detail
        With Get������
            .ID = rsTmp!ID
            .��� = rsTmp!���
            .������� = rsTmp!�������
            .���� = rsTmp!����
            .���� = rsTmp!����
            .��� = Nvl(rsTmp!���)
            .���㵥λ = Nvl(rsTmp!���㵥λ)
            .��� = False '���ɱ��
            .�Ӱ�Ӽ� = False '���Ӱ�Ӽ�
            .���ηѱ� = True '��ѱ��޹�
            .˵�� = Nvl(rsTmp!˵��)
            .ִ�п��� = Nvl(rsTmp!ִ�п���, 3) 'ȱʡΪ����Ա���ڿ���
        End With
    Else
        Set Get������ = Nothing
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set Get������ = Nothing
End Function

Public Function isSimple(strNo As String) As Boolean
'���ܣ��жϵ����Ƿ�Ϊ���շѲ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Distinct ��ҩ����,�շ����,Nvl(����,1) as ����" & _
        " From ������ü�¼ Where ��¼״̬ IN(1,3)" & _
        " And ��¼����=1 And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If rsTmp.RecordCount = 1 Then
        If rsTmp!���� = 1 And rsTmp!�շ���� = "Z" And Nvl(rsTmp!��ҩ����) = "Z" Then isSimple = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillingWarn(strPrivs As String, str���� As String, str���ò��� As String, _
    rsWarn As ADODB.Recordset, cur��� As Currency, cur���ն� As Currency, _
    cur���ݽ�� As Currency, cur���� As Currency, str��� As String, _
    ByVal str����� As String, ByRef str�ѱ���� As String, Optional bln�ಡ�� As Boolean) As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:
'     str����=��������,������ʾ
'     str���ò���=���ݲ�����ݷ��صļ��ʱ������÷���
'     rsWarn=��ǰ�������ʱ������ü�¼
'     cur���=�������,�����ۼƱ���
'     cur���ն�=���˵��շ����ķ��ö�,����ÿ�ձ���
'     cur���ݽ��=���˵���������ķ���
'     cur����=���˵������ö�,�����ۼƱ���
'     str���=��ǰҪ�������,���ڷ��౨��
'     str�����=�������,������ʾ
'����:0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
'     str�������="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
    Dim i As Integer, byt��־ As Byte
    Dim bln�ѱ��� As Boolean
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    '�����������
    rsWarn.Filter = "����ID=0 And ���ò���='" & str���ò��� & "'"
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str���) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str����� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str���) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str����� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str���) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str����� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    If bln�ಡ�� Then
        'ʾ����",��:-,��:DEF,��:567,��567"
        '������־2ʾ����",��:-��,��:DEF��,��:567��,��567��"
        bln�ѱ��� = str�ѱ���� & "," Like "*," & str���� & ":-*,*" _
            Or str�ѱ���� & "," Like "*," & str���� & ":*" & str��� & "*,*"
    Else
        'ʾ����"-" �� ",ABC,567,DEF"
        '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
        bln�ѱ��� = InStr(str�ѱ����, str���) > 0 Or str�ѱ���� Like "-*"
    End If
    
    If bln�ѱ��� Then
        If byt��־ = 2 Then
            If bln�ಡ�� Then
                arrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(arrTmp)
                    If "," & arrTmp(i) & "," Like "*," & str���� & ":-*,*" _
                        Or "," & arrTmp(i) & "," Like "*," & str���� & ":*" & str��� & "*,*" Then
                        byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                        'Exit For  '˵����סԺģ��
                    End If
                Next
            Else
                If str�ѱ���� Like "-*" Then
                    byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
                Else
                    arrTmp = Split(str�ѱ����, ",")
                    For i = 0 To UBound(arrTmp)
                        If InStr(arrTmp(i), str���) > 0 Then
                            byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                            'Exit For '˵����סԺģ��
                        End If
                    Next
                End If
            End If
        Else
            Exit Function
        End If
    End If
    
    If str����� <> "" Then str����� = """" & str����� & """����"
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        '��ֻ������:1.ǿ�Ƽ���,��Ȩ��ʱ,��ֹ����
                        Call MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ����", vbInformation + vbOKOnly, gstrSysName)
                        BillingWarn = 3
                        '--26349
'                        If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                            BillingWarn = 2
'                        Else
'                            BillingWarn = 1
'                        End If

                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & " ����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If cur��� + cur���� - cur���ݽ�� < 0 Then
                        byt��ʽ = 2
                        If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                            MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�", vbInformation, gstrSysName
                            BillingWarn = 3
                        Else
                            MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���", vbInformation, gstrSysName
                            BillingWarn = 4
                        End If
                    ElseIf cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                            If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1
                            End If
                        Else
                            MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If cur��� + cur���� - cur���ݽ�� < 0 Then
                            byt��ʽ = 2
                            If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                                MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�", vbInformation, gstrSysName
                                BillingWarn = 3
                            Else
                                MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���", vbInformation, gstrSysName
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", vbInformation, gstrSysName
                        BillingWarn = 3
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ն� + cur���ݽ�� > rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        Call MsgBox(str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ����.", vbOKOnly + vbInformation, gstrSysName)
                        BillingWarn = 3
'                        If MsgBox(str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                            BillingWarn = 2
'                        Else
'                            BillingWarn = 1
'                        End If
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ն� + cur���ݽ�� > rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        MsgBox str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", vbInformation, gstrSysName
                        BillingWarn = 3
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־3
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStock(ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long, Optional ByVal lng���� As Long = -1) As Double
'���ܣ���ȡָ��ҩ��ָ��ҩƷ���(�����۵�λ)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '��ȡ���(�����������),ҩ��������(����=0,����Ϊҩ��)����Ч��
    If lng���� = -1 Or lng���� = 0 Then
        strSQL = _
            " Select Nvl(Sum(A.��������),0) as ��� From ҩƷ��� A" & _
            " Where (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
            " And A.����=1 And A.ҩƷID=[1] And A.�ⷿID=[2]"
    Else
        strSQL = _
            " Select Nvl(Sum(A.��������),0) as ��� From ҩƷ��� A" & _
            " Where Nvl(A.����,0)= [3] And (A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
            " And A.����=1 And A.ҩƷID=[1] And (A.�ⷿID = [2] Or A.�ⷿID In (Select ����ⷿid From ����ⷿ���� Where ����id = [2] And Rownum < 2))  "
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngҩƷID, lngҩ��ID, lng����)
    If Not rsTmp.EOF Then GetStock = rsTmp!���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMultiStock(ByVal lngҩƷID As Long, ByVal strҩ��IDs As String) As Double
'���ܣ���ȡָ��ҩ��ָ��ҩƷ���(�����۵�λ)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '��ȡ���(�����������),ҩ��������(����=0,����Ϊҩ��)����Ч��
    strSQL = _
        " Select Nvl(Sum(A.��������),0) as ��� From ҩƷ��� A" & _
        " Where (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
        " And A.����=1 And A.ҩƷID=[1] And instr([2],','|| A.�ⷿID ||',')>0"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngҩƷID, "," & strҩ��IDs & ",")
    If Not rsTmp.EOF Then GetMultiStock = rsTmp!���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPlace(ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long, _
    Optional bln���� As Boolean = False) As String
'���ܣ���ȡָ��ҩƷ��ָ��ҩ���Ŀⷿ��λ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    If bln���� Then
        strSQL = "Select �ⷿ��λ From ���ϴ����޶� Where ����ID=[1] And �ⷿID=[2]"
    Else
        strSQL = "Select �ⷿ��λ From ҩƷ�����޶� Where ҩƷID=[1] And �ⷿID=[2]"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngҩƷID, lngҩ��ID)
    If Not rsTmp.EOF Then GetPlace = Nvl(rsTmp!�ⷿ��λ)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub GetDoctor(lng����ID As Long, ByRef rsTmp As ADODB.Recordset, _
    Optional ByVal bln������Ա���� As Boolean, Optional ByVal str�������� As String = "�ٴ�")
    '���ܣ���ȡָ���������ҵ�ҽ����ʿ,���δָ���������ң����ȡ����ҽ����ʿ
    '��Σ�
    '     bln������Ա����-����Ա�����������µ���Ա
    Dim strSQL As String, bln��ʿ As Boolean
    Dim strWhere As String
    
    On Error GoTo errH
    If lng����ID = 0 And bln������Ա���� Then
        strWhere = " And Exists (Select 1 From ������Ա M, ������Ա N" & _
                    " Where m.����id = n.����id And m.��Աid = a.Id And n.��Աid = [1])"
    End If
    
    str�������� = Replace(str��������, "'", "")
    If str�������� <> "" Then
        If InStr(1, str��������, ",") > 0 Then
            strWhere = strWhere & " And Instr(','||[2]||',',','||d.��������||',')>0"
        Else
            strWhere = strWhere & " And d.�������� = [2]"
        End If
    End If
    
    '�����Ź�������Ϊ���ٴ���ҽ����ʿ,��Ϊ���ܸ���������һ�������Ƿ�ĩ������.
    If rsTmp Is Nothing Then
        strSQL = _
            "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
            " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��,A.רҵ����ְ��,B.ȱʡ" & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C,��������˵�� D" & _
            " Where A.ID = B.��ԱID And A.ID=C.��ԱID And B.����ID=D.����ID And C.��Ա���� IN('ҽ��','��ʿ') " & _
            " And D.������� IN(" & gint������Դ & ",3) And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & strWhere & vbNewLine & _
            " Order by " & IIf(gbyt��������ʾ = 1, "����", "���") & ",ȱʡ Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", UserInfo.ID, str��������)
    End If
   
    '�����������л�ʿ�������շ���Ŀ�������Ƽ�����ʱ�Ŷ�ȡ��ʿ
    bln��ʿ = gbln��ʿ And (gstr�շ���� = "" Or gstr�շ���� Like "*'E'*" Or gstr�շ���� Like "*'M'*" Or gstr�շ���� Like "*'4'*")
    If lng����ID = 0 Then
        rsTmp.Filter = IIf(bln��ʿ, "", "��Ա����='ҽ��'")
    Else
        rsTmp.Filter = "����ID=" & lng����ID & IIf(bln��ʿ, "", " And ��Ա����='ҽ��'")
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub GetDoctorDept(ByRef rsTmp As ADODB.Recordset, _
    Optional ByVal bln������Ա���� As Boolean, _
    Optional ByVal str���� As String = "'�ٴ�','����'", _
    Optional ByVal lngDeptID As Long)
    '���ܣ���ȡ���п�������
    '��Σ�
    '     bln������Ա����-����Ա����������
    '     str����='�ٴ�','����','��ҩ��',...,����Ϊ��
    '      lngDeptID-��ǰ����ID�����ID
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errH
    If bln������Ա���� Then
        strWhere = " And Exists (Select 1 From ������Ա Where ����id = a.Id And ��Աid = [1])"
    End If
    
    str���� = Replace(str����, "'", "")
    If str���� <> "" Then
        If InStr(1, str����, ",") > 0 Then
            strWhere = strWhere & " And Instr(','||[2]||',',','||B.��������||',')>0"
        Else
            strWhere = strWhere & " And B.�������� = [2]"
        End If
    End If
    If lngDeptID <> 0 Then strWhere = strWhere & " And A.ID=[3]"
    
    strSQL = _
        "Select A.ID, A.����, A.����, A.����, 0 As ȱʡ, B.��������, Decode(D.���ȼ�, 1, (Decode(C.���ȼ�, 1, 1, 2)), 3) ���ȼ�" & vbNewLine & _
        "From ���ű� A, ��������˵�� B," & vbNewLine & _
        "     (Select ����id, Max(Decode(Instr('���,����,����,����,Ӫ��,���', ��������), 0, 1, 2)) As ���ȼ�" & vbNewLine & _
        "       From ��������˵�� Where ������� <> 0" & vbNewLine & _
        "       Group By ����id) C, (Select ����id, Max(Decode(�������, 1, 1, 2)) As ���ȼ� From ��������˵�� Where ������� <> 0 Group By ����id) D" & vbNewLine & _
        "Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And A.ID = B.����id And" & vbNewLine & _
        "      B.����id = C.����id And B.����id = D.����id And B.������� In (" & gint������Դ & ", 3)" & vbNewLine & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & strWhere & vbNewLine & _
        "Order By ���ȼ�,����"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", UserInfo.ID, str����, lngDeptID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function Read��ҩ����(lngID As Long) As ADODB.Recordset
'���ܣ���ȡָ��ҩ���ķ�ҩ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ��ҩ���� Where ҩ��ID=[1] Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngID)
    Set Read��ҩ���� = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWindow(ByVal lngҩ��ID As Long, ByRef rs��ҩ���� As ADODB.Recordset) As Boolean
'���ܣ�ȷ��ҽԺָ��ҩ���Ƿ�ʹ�÷�ҩ����
'˵������Ϊר�Ҵ���ָ��,����̬����,�����ſ�
    Dim strSQL As String
    
    On Error GoTo errH
    
    If rs��ҩ���� Is Nothing Then
        strSQL = "Select ����,ҩ��ID From ��ҩ���� Where Nvl(ר��,0)=0"
        Set rs��ҩ���� = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rs��ҩ����, strSQL, "mdlOutExse")
    End If
    
    If lngҩ��ID = 0 Then
        rs��ҩ����.Filter = ""
    Else
        rs��ҩ����.Filter = "ҩ��ID=" & lngҩ��ID
    End If
    If Not rs��ҩ����.EOF Then ExistWindow = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFeeKind() As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ����, ����, ���� From �շ���Ŀ���"
    Set GetFeeKind = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շ����")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillExistReplenishData(intType As Integer, Optional lngBalance As Long, Optional strNos As String, Optional lng��ӡID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ���ڶ��ν���
    '����:True-���ڶ��ν������� False-�����ڶ��ν�������
    '���:intType:0-�շ����ݣ�ʹ��lngBalanceΪ�������
    '     intType:1-�շ����ݣ�ʹ��strNosΪ���ݺ�
    '     intType:2-���ݴ�ӡID���ж��Ƿ�ʹ�ò�����
    '     lng��ӡID-��ӡID>0ʱ������ʱ����ȡ��
    '����:������
    '����:2014-10-15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, cllPro As Collection
    Dim strValue(0 To 0) As String
    Dim varData() As Variant
    
    On Error GoTo errHandle
        
    If lng��ӡID > 0 And intType = 2 Then
        strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From ���ò����¼ A, (Select Distinct ����id From ������ü�¼ B,��ʱƱ�ݴ�ӡ���� C Where B.NO=C.NO and mod(B.��¼����,10)=1 and C.����=1 and C.ID=[2]) B" & vbNewLine & _
        " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2 And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", lngBalance, lng��ӡID)
    ElseIf intType = 0 Then
        strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From ���ò����¼ A, (Select Distinct ����id From ����Ԥ����¼ Where ������� = [1]) B" & vbNewLine & _
        " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2 And Rownum < 2"
        strSQL = strSQL & " Union " & _
        " Select 1 From ���ò����¼ Where ������� = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", lngBalance, lng��ӡID)
    Else
        If Len(strNos) <= 4000 Then
            strSQL = "" & _
            " Select /*+ Rule */ 1" & vbNewLine & _
            " From ���ò����¼ A," & vbNewLine & _
            "      (Select Distinct ����id" & vbNewLine & _
            "       From ������ü�¼" & vbNewLine & _
            "       Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(f_Str2list([1])))) B" & vbNewLine & _
            " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2 And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", strNos)
        Else
            If zlGetSplitString4000(strNos, cllPro) = False Then Exit Function
            If zlFromCollectBulidSQL(cllPro, strSQL, varData) = False Then Exit Function
            strSQL = "With ������Ϣ as (" & strSQL & ")" & vbCrLf
            strSQL = strSQL & _
            "      Select Distinct A1.����id" & vbNewLine & _
            "       From ������ü�¼ A1,������Ϣ A2" & vbNewLine & _
            "       Where Mod(A1.��¼����, 10) = 1 And A1.NO=A2.NO " & vbNewLine
            strSQL = "" & _
            "   Select 1" & vbNewLine & _
            "   From ���ò����¼ A," & vbNewLine & _
            "        (" & strSQL & ") B " & vbNewLine & _
            "   Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2 And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecordByArray(strSQL, "�����ν���", varData)
        End If
    End If
    
    If rsTmp.EOF Then
        CheckBillExistReplenishData = False
    Else
        CheckBillExistReplenishData = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function InitSysPar() As Boolean
    '���ܣ���ʼ��ϵͳ����
    '���أ���-����ɹ�
    Dim strValue As String
    On Error Resume Next
    
    '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    '����:35242
    gbln�����л� = IIf(Val(zlDatabase.GetPara("����ƥ�䷽ʽ�л�", , , 1)) = 1, 1, 0) = 1
    
    '������ʾ��ʽ
    'gblnShowCard = zlDatabase.GetPara(12, glngSys) = "0"
    
    '���˺� ����:????    ����:2010-12-06 23:38:53
    '���õ��۱���λ��
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    '�շѷֱҴ���ʽ
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    gBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 2, 1)))

    '�������뷽ʽ
    strValue = zlDatabase.GetPara(17, glngSys, , "1111")
    gblnInputName = (Mid(strValue, 1, 1) = "1")
    gblnInputCard = (Mid(strValue, 2, 1) = "1")
    gblnInputNO = (Mid(strValue, 3, 1) = "1")
    gblnInputID = (Mid(strValue, 4, 1) = "1")
    
    'ָ��ҩ��ʱ���ƿ��
    gblnStock = zlDatabase.GetPara(18, glngSys) = "1"
    
    '���ڷ��䷽ʽ
    gbytAssign = Val(zlDatabase.GetPara(19, glngSys, , 0))
        
    'Ʊ�ݺ��볤�ȡ����￨�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytFactLength = Val(Split(strValue, "|")(0))
    'gbytCardNOLen = Val(Split(strValue, "|")(4))
    
    '�Һ���Ч����
    '���˺�:34717
    '��λ:ǰһλ���ܹҺ�;��һλ����Һ�
    strValue = zlDatabase.GetPara(21, glngSys, , "01") & "1"
    gTy_System_Para.Sy_Reg.bytNODaysGeneral = Val(Left(strValue, 1))
    gTy_System_Para.Sy_Reg.bytNoDayseMergency = Val(Mid(strValue, 2, 1))
    'If gTy_System_Para.Sy_Reg.bytNODaysGeneral = 0 Then gTy_System_Para.Sy_Reg.bytNODaysGeneral = 1
    ' If gTy_System_Para.Sy_Reg.bytNoDayseMergency = 0 Then gTy_System_Para.Sy_Reg.bytNoDayseMergency = 1
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys, , 0))
    
    'Ʊ���ϸ����
    gblnStrictCtrl = Mid(zlDatabase.GetPara(24, glngSys, , "00000"), 1, 1) = "1"
        
    'һ��ͨ����ˢ������
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    gdblԤ��������鿨 = Val(Split(strValue, "|")(0))
    gbytԤ����˷��鿨 = Val(Split(strValue, "|")(1))
    gbln���ѿ��˷��鿨 = zlDatabase.GetPara(282, glngSys) = "1"
    
    'ҽ����������
    gstrҽ���������� = "'" & Replace(zlDatabase.GetPara(41, glngSys), "|", "','") & "'"

    '���ѷ�������
    gstr���ѷ������� = "'" & Replace(zlDatabase.GetPara(42, glngSys), "|", "','") & "'"
    
    '�շ���Ŀ�������ƥ�䷽ʽ
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    
    '�Զ���ҩ��ҩ
    gbln�շѺ��Զ���ҩ = zlDatabase.GetPara(45, glngSys) = "1"
            
    'ˢ��Ҫ����������
    gstrCardPass = zlDatabase.GetPara(46, glngSys, , "0000000000")
        
      
    '���ն�����
    gbytҽ�������� = Val(zlDatabase.GetPara(59, glngSys, , 1))
    
    
    
    
    '���ʷ���������ѽ��
    gcurMaxMoney = Val(zlDatabase.GetPara(60, glngSys))
    
    '�Ƿ�Ҫ�������������
    gbln�շ���� = zlDatabase.GetPara(72, glngSys) = "1"
    
    '����ҩ�Ƿ���Ʒ����ʾ
    'gbln��Ʒ�� = zlDatabase.GetPara(74, glngSys) = "1"
        
    
    '���ŵ���ʹ�ö��ֽ��㷽ʽģʽ
    gblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
        
      
    '�Զ�����
    gbln�����Զ����� = zlDatabase.GetPara(92, glngSys) = "1"
    
    '������Ŀ���ܼ����ۿ�
    gbln��������ۿ� = zlDatabase.GetPara(93, glngSys) = "1"
    

    '���ʱ����������۷���
    gbln�����������۷��� = zlDatabase.GetPara(98, glngSys) = "1"
        
    '���������ʱ,���������Ŀʱ,��λ����������
    gblnFeeKindCode = zlDatabase.GetPara(144, glngSys) = "1" And Not gbln�շ����
    
    gbytMediOutMode = Val(zlDatabase.GetPara(150, glngSys))
    gbln�˷�����ģʽ = IIf(zlDatabase.GetPara(151, glngSys) = "1", True, False)
    gbln����ʾ�޿������ = zlDatabase.GetPara(316, glngSys) = "1"
    'e.����ȫ�ֲ���
    '-------------------------------------------------------------------------------------------------
    '����:27990
    With gTy_System_Para
        .byt����ҩƷ��ʾ = Val(zlDatabase.GetPara("����ҩƷ��ʾ")) '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
        .bytҩƷ������ʾ = Val(zlDatabase.GetPara("ҩƷ������ʾ"))  '��0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
        .byt��������ʶ����� = Val(zlDatabase.GetPara(320, glngSys, , "0"))      '1-����ͨ��ɨ��¼���¼������;0-�����ƣ�����ͨ������Ȳ���
    End With
    
    '��ǰ����Ա�Ƿ�Ϊ�ٴ�������Ա
    gblnUserIsClinic = UserIsClinic(UserInfo.ID)
    InitSysPar = True
End Function

Public Sub InitLocPar(lngModul As Long)
'���ܣ���ʼ������ģ�����
    Dim strValue As String, intType As Integer
    Dim arrTmp As Variant
    
    On Error Resume Next
    
    'a.����ע���洢��ģ�����
    '----------------------------------------------------------------------------------------
    If lngModul = 1120 Or lngModul = 1121 Or lngModul = 1122 Then
        gblnLED = Val(GetSetting("ZLSOFT", "����ȫ��", "ʹ��", 0)) <> 0
    End If
    'b.���ݿ�洢�Ĺ���ȫ�ֲ���
    '----------------------------------------------------------------------------------------
    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    gbytCode = Val(zlDatabase.GetPara("���뷽ʽ"))
    gblnMyStyle = zlDatabase.GetPara("ʹ�ø��Ի����") = "1"
     'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------
    If lngModul = 1120 Or lngModul = 1121 Or lngModul = 1122 Then
        gstr���㷽ʽ = zlDatabase.GetPara("ȱʡ���㷽ʽ", glngSys, lngModul)
        gstr�ѱ� = zlDatabase.GetPara("ȱʡ�ѱ�", glngSys, lngModul)
        
        If glngSys Like "8??" Then
            gstr�շ���� = "'5','6','7'"
        Else
            gstr�շ���� = zlDatabase.GetPara("�շ����", glngSys, lngModul)
        End If
        
        gbln�ֹ����� = zlDatabase.GetPara("�ֹ�����", glngSys, lngModul) = "1"
        gblnLedDispDetail = zlDatabase.GetPara("LED��ʾ�շ���ϸ", glngSys, lngModul, "1") = "1"
        gblnLedWelcome = zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, lngModul, "1") = "1"
        
        gblnSharedInvoice = zlDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, lngModul) = "1"
        gblnPay = zlDatabase.GetPara("��ҩ����", glngSys, lngModul) = "1"
        gblnTime = zlDatabase.GetPara("�������", glngSys, lngModul) = "1"
        gbln��ʿ = zlDatabase.GetPara("��ʾ��ʿ", glngSys, lngModul) = "1"
    
        gblnҩ����λ = zlDatabase.GetPara("ҩƷ��λ", glngSys, lngModul) = "1"
    
        glng��ҩ�� = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, lngModul)
        glng��ҩ�� = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, lngModul)
        glng��ҩ�� = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, lngModul)
        glng���ϲ��� = zlDatabase.GetPara("ȱʡ���ϲ���", glngSys, lngModul)
        gbln����ҩ�� = zlDatabase.GetPara("��ʾ����ҩ�����", glngSys, lngModul) = "1"
        gbln����ҩ�� = zlDatabase.GetPara("��ʾ����ҩ����", glngSys, lngModul) = "1"
        If lngModul = 1121 Then
            gbln���������ɿ� = zlDatabase.GetPara("��ȡ���ۺ������ɿ�", glngSys, lngModul) = "1"
        End If

        
        '31936:����ҩ����ҩ��(�ǲ���Ա��������)�Ŀ����ʾ��ʽ:0-ֱ����ʾ���;1-��ʾ����
        '         �˲�����Ҫ���:��ʾ����ҩ��������ʾ����ҩ����Ĳ�������������,�����ѡ������һ��,�ò�����������
        If lngModul = 1121 Then
            gbyt�����ʾ��ʽ = 0: '�����շѲ���
        Else
            gbyt�����ʾ��ʽ = Val(zlDatabase.GetPara("�����ʾ��ʽ", glngSys, lngModul))
        End If
        
        '���뷢ҩʱ�ļ��
        gstr��ҩ�� = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, lngModul)
        gstr��ҩ�� = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, lngModul)
        gstr��ҩ�� = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, lngModul)
        
        gbyt����ҽ�� = Val(zlDatabase.GetPara("����ҽ��", glngSys, lngModul, 0))
        gbyt��������ʾ = Val(zlDatabase.GetPara("��������ʾ��ʽ", glngSys, lngModul, 1))
        gblnSeekName = zlDatabase.GetPara("����ģ������", glngSys, lngModul) = "1"
        gintNameDays = Val(zlDatabase.GetPara("������������", glngSys, lngModul))
        gbln����¼������ʹ�õĿ����� = Val(zlDatabase.GetPara("����¼������ʹ�õĿ�����", glngSys, lngModul)) = 1
    End If
    
    gbln������Դ��Ȩ�޿��� = False
    
    If lngModul = 1120 Or lngModul = 1121 Then
        gcurMax = Val(zlDatabase.GetPara("�����", glngSys, lngModul))
        gstr�д� = zlDatabase.GetPara("��ҩ������", glngSys, lngModul)
        gstr���� = zlDatabase.GetPara("��ҩ������", glngSys, lngModul)
        gstr�ɴ� = zlDatabase.GetPara("��ҩ������", glngSys, lngModul)
        
        '��꾭����Ŀ
        gbln�Ա� = zlDatabase.GetPara("�Ա�", glngSys, lngModul) = "1"
        gbln���� = zlDatabase.GetPara("����", glngSys, lngModul) = "1"
        gbln�ѱ� = zlDatabase.GetPara("�ѱ�", glngSys, lngModul) = "1"
        gblnҽ�Ƹ��� = zlDatabase.GetPara("ҽ�Ƹ���", glngSys, lngModul) = "1"
        gbln�Ӱ� = zlDatabase.GetPara("�Ӱ�", glngSys, lngModul) = "1"
        gbln�������� = zlDatabase.GetPara("��������", glngSys, lngModul) = "1"
        gbln������ = zlDatabase.GetPara("������", glngSys, lngModul) = "1"
        
        gbln��ȱʡ������ = zlDatabase.GetPara("��ʹ��ȱʡ������", glngSys, lngModul) = "1"
        gbln�����俪���� = zlDatabase.GetPara("����Ҫ���뿪����", glngSys, lngModul) = "1"
        gblnȱʡ�������� = zlDatabase.GetPara("ȱʡ��������", glngSys, lngModul) = "1"
        gint����ϼ� = Val(zlDatabase.GetPara("����ϼƷ�ʽ", glngSys, lngModul))
        If lngModul = 1120 Then
            gint����֪ͨ�� = Val(zlDatabase.GetPara("����֪ͨ����ӡ��ʽ", glngSys, lngModul))
            gintDelPrice = Val(zlDatabase.GetPara("ȡ�����۵�", glngSys, lngModul))
        ElseIf lngModul = 1121 Then
            '���˺�:��������ģ�����,��������,½���ڽ��д���
            Dim strTmp As String
            With gTy_Module_Para
                'gbln�ɿ���� = zlDatabase.GetPara("�շѽɿ��������", glngSys, lngModul) = "1"
                .byt�ɿ���� = Val(zlDatabase.GetPara("�շѽɿ��������", glngSys, lngModul))   '����:22343;51670
                strTmp = Trim(zlDatabase.GetPara("Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", glngSys, lngModul, "0|10"))
                If Val(Split(strTmp & "|", "|")(0)) = 0 Then
                    .int����ʣ��Ʊ������ = -1
                Else
                    .int����ʣ��Ʊ������ = Val(Split(strTmp & "|", "|")(1))     '����:26948
                End If
            End With
            '25187
            '���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
            strTmp = Trim(zlDatabase.GetPara("Ʊ�ݷ������", glngSys, lngModul, "0||0;0;0;0;0"))
            arrTmp = Split(strTmp & "||", "||")
            With gTy_Module_Para
                .bytƱ�ݷ������ = Val(arrTmp(0))
                arrTmp = Split(arrTmp(1) & ";;;;;", ";")
                .blnƱ�ݷֵ��� = Val(arrTmp(0)) = 1
                .intִ�п��� = Val(arrTmp(1))
                .int�վݷ�Ŀ = Val(arrTmp(2))
                .int�շ�ϸĿ = Val(arrTmp(3))
                .bytƱ�ݻ������� = Val(arrTmp(4))
                .bln�ֱ��ӡ = zlDatabase.GetPara("���ŵ����շѷֱ��ӡ", glngSys, lngModul) = "1"
                .bytƱ�����ɷ�ʽ = Val(zlDatabase.GetPara("�շ�Ʊ�����ɷ�ʽ", glngSys, lngModul))
                .byt�����վ��д� = Val(zlDatabase.GetPara("�շ��վ����д�", glngSys, lngModul, 3))
                .blnһ��Ʊ�� = zlDatabase.GetPara("�շ�ÿ��ֻ��һ��Ʊ��", glngSys, lngModul) = "1"
                .bln������ = Val(zlDatabase.GetPara("�վݼ��չ�����", glngSys, lngModul, "0")) = 1
                .bln���ռ��Ʊ�� = Val(zlDatabase.GetPara("����ʹ��Ʊ��", glngSys, lngModul, "0")) = "0"
                .blnʹ�üӼ��л� = Val(zlDatabase.GetPara("ʹ�üӼ��л�֧����ʽ", glngSys, lngModul, "1")) = "1"
                .bytҩƷ��ҩ�˷ѷ�ʽ = Val(zlDatabase.GetPara("ҩƷ��ҩ�˷ѷ�ʽ", glngSys, lngModul, "0"))  '47400
                .byt�˷�ȱʡѡ��ʽ = Val(zlDatabase.GetPara("�˷�ȱʡѡ��ʽ", glngSys, lngModul, "0")) '87489
                .bytˢ��ȱʡ������ = Val(zlDatabase.GetPara("ˢ��ȱʡ������", glngSys, 1151, "0")) '86853
                .blnֻ��ҽ������ɹ������շ� = Val(zlDatabase.GetPara("ֻ��ҽ������ɹ������շ�", glngSys, lngModul, "0")) '91665
                .str�����շ�ִ�п��� = zlDatabase.GetPara("�����շ�ִ�п���", glngSys, lngModul)
                .str�������շ�ִ�п��� = zlGet�������շ�ִ�п���(lngModul)
                .blnҽ��������ȱʡ��λ = Val(zlDatabase.GetPara("ҽ��������ȱʡ��λ", glngSys, lngModul, "0")) = "1"
                .bln�ֽ��˿�ȱʡ��ʽ = Val(zlDatabase.GetPara("�ֽ��˿�ȱʡ��ʽ", glngSys, lngModul, "0")) = "1"
                .bln���ֱ��ӡ = zlDatabase.GetPara("��첡�˷ֵ��ݴ�ӡ", glngSys, lngModul) = "1"
                .strȱʡ���� = zlDatabase.GetPara("���������˷�ȱʡ��ʽ", glngSys, lngModul, "")
            End With
            
            gbln�ۼ� = zlDatabase.GetPara("��ʾ�ۼ�", glngSys, lngModul) = "1"
            gblnCheckTest = zlDatabase.GetPara("���Ƥ�Խ��", glngSys, lngModul) = "1"
            gblnPrePayPriority = zlDatabase.GetPara("����ʹ��Ԥ����", glngSys, lngModul) = "1"
            
            
            
            glngAddedItem = Val(zlDatabase.GetPara("�Զ����չҺŷ�", glngSys, lngModul))
            gblnShowErr = zlDatabase.GetPara("��ʾ������", glngSys, lngModul) = "1"
            gblnMulti = zlDatabase.GetPara("�൥���շ�", glngSys, lngModul) = "1"
            
            gblnSeekBill = zlDatabase.GetPara("��Ѱ���۵���", glngSys, lngModul) = "1"
            gintSeekDays = Val(zlDatabase.GetPara("��Ѱ��������", glngSys, lngModul))
            gblnUnPopPriceBill = zlDatabase.GetPara("���������۵�ѡ��", glngSys, lngModul) = "1"
        
            gblnCheckRegeventDept = zlDatabase.GetPara("��鲡�˹Һſ���", glngSys, lngModul) = "1"
            gbytUnRegevent = Val(zlDatabase.GetPara("δ�ҺŲ����շ�", glngSys, lngModul))
            gbytAutoSplitBill = Val(zlDatabase.GetPara("�Զ���ϵ���", glngSys, lngModul))
            gint�շ��嵥 = Val(zlDatabase.GetPara("�շ��嵥��ӡ��ʽ", glngSys, lngModul))
            glngFactNormal = Val(zlDatabase.GetPara("��ͨ��Ʊ��ʽ", glngSys, lngModul))
            glngFactMediCare = Val(zlDatabase.GetPara("ҽ����Ʊ��ʽ", glngSys, lngModul))
        End If
    ElseIf lngModul = 1122 Then
        gint����ϼ� = 0        '���˺�:�����������û�иò�������,���,ֻ�ܰ��վݷ�Ŀ��ͳ��
        gblnOnlyUnitPatient = zlDatabase.GetPara("ֻ���Һ�Լ��λ����", glngSys, lngModul) = "1"
        gbln���ʴ�ӡ = zlDatabase.GetPara("���ʴ�ӡ", glngSys, lngModul) = "1"
        gbln���۴�ӡ = zlDatabase.GetPara("���۴�ӡ", glngSys, lngModul) = "1"
        gbln��˴�ӡ = zlDatabase.GetPara("��˴�ӡ", glngSys, lngModul) = "1"
    End If
    
    If lngModul = 1120 Or lngModul = 1121 Or lngModul = 1122 Then
        gint������Դ = Val(zlDatabase.GetPara("������Դ", glngSys, lngModul, , , , intType))
        '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
        gbln������Դ��Ȩ�޿��� = IIf(intType = 1 Or intType = 3 Or intType = 15, True, False)
        
        If gint������Դ = 1 Then
            gstrҩ����λ = "���ﵥλ": gstrҩ����װ = "�����װ"
        Else
            gstrҩ����λ = "סԺ��λ": gstrҩ����װ = "סԺ��װ"
        End If
    End If
    
    'd.���ݻ�������
    '-------------------------------------------------------------------------------------------------
    If lngModul = 1120 Or lngModul = 1121 Or lngModul = 1122 Then
        gblnҩ���ϰల�� = Checkҩ���ϰల��
    End If
    
    If lngModul = 1121 Then
        Call GetErrorItem(glng���ϸĿID, gstr����վݷ�Ŀ)
        gstr�������� = zlGet��������
    End If
    
    If lngModul = 1124 Then
        gstr�������� = zlGet��������
        gblnҩ����λ = zlDatabase.GetPara("ҩƷ��λ��ʾ", glngSys, lngModul) = "1"
        gstrҩ����λ = "���ﵥλ": gstrҩ����װ = "�����װ"
        gblnSeekName = Split(zlDatabase.GetPara("����ģ�����ҷ�ʽ", glngSys, lngModul, "0|0"), "|")(0) = "1"
        gintNameDays = Val(Split(zlDatabase.GetPara("����ģ�����ҷ�ʽ", glngSys, lngModul, "0|0"), "|")(1))
        With gTy_Module_Para
            .blnʹ�üӼ��л� = Val(zlDatabase.GetPara("ʹ�üӼ��л�֧����ʽ", glngSys, lngModul, "1")) = "1"
            .bytҩƷ��ҩ�˷ѷ�ʽ = Val(zlDatabase.GetPara("ҩƷ��ҩ�˷ѷ�ʽ", glngSys, lngModul, "0"))  '47400
        End With
    End If
End Sub

Private Function zlGet�������շ�ִ�п���(ByVal lngModul As Long) As String
    '��ȡ�����õ������շ�ִ�п���
    '���أ�
    '   ��ʽ:����ID1,����ID2,����ID3,...
    Dim strSQL As String, rsTemp As Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select b.����ֵ" & vbNewLine & _
            " From zlParameters A, zlUserParas B" & vbNewLine & _
            " Where a.Id = b.����id And Nvl(a.ϵͳ, 0) = [1] And Nvl(a.ģ��, 0) = [2]" & vbNewLine & _
            "       And a.������ = [3] And b.����ֵ Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����շ�ִ�п��Ҳ���", glngSys, lngModul, "�����շ�ִ�п���")
    If rsTemp Is Nothing Then Exit Function
    Do While Not rsTemp.EOF
        strTemp = strTemp & "," & Nvl(rsTemp!����ֵ)
        rsTemp.MoveNext
    Loop
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    zlGet�������շ�ִ�п��� = strTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBalanceSet() As ADODB.Recordset
    '���ܣ�����һ�������¼������
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Fields.Append "�������", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "���㷽ʽ", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "������", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "��������", adBigInt, , adFldIsNullable '1-ҽ��;2-���ѿ�;3-ҽ�ƿ�;0-����
    rsTmp.Fields.Append "�����ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "����", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "������ˮ��", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "����˵��", adVarChar, 500, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set GetBalanceSet = rsTmp
End Function

Public Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str������ As String, ByVal str�������� As String, _
    ByVal intMode As Integer, ByVal intPrice As Integer, Optional ByVal intPage As Integer, Optional ByVal lngRow As Long) As ADODB.Recordset
'���ܣ����ݵ��ݶ������ݴ���һ����ϸ��¼����Ϣ(���ۼ۵�λ)
'�ֶΣ�����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������,ִ�п���ID,
'         �������ʣ�1-�շѵ�,2-���ʵ�),�Ƿ񻮼�(1-����;0-�������շѼ����ʵ�)
'������intPage=ָ���ĵ���,lngRow=ָ�����У���ָ��ʱ�������е��ݵ�������
'         intMode:�������ʣ�1-�շѵ�,2-���ʵ�)
'         intPrice:�Ƿ񻮼�(1-����;0-�������շѼ����ʵ�)
    Dim i As Integer, j As Integer, p As Integer, strSQL As String
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl���� As Double, curʵ�� As Currency
    Dim rsTmp As New ADODB.Recordset, rsPrice As ADODB.Recordset
    Dim intStartPage  As Integer, intPages As Integer
    
    On Error GoTo errHandle
    
    rsTmp.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    '69788:���ϴ�,2014-6-5,�����������ֶδ�С����20��Ϊ100
    '79420,���ϴ�,2014/11/10:������¼���ֶδ�С
    rsTmp.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    '131048,����,2018-9-5,��CheckChargeItem���ӿ��е�rsDetail ��¼���������ֶ�:
    '                                  ִ�п���ID���������ʣ�1-�շѵ�,2-���ʵ�)���Ƿ񻮼�(1-����,0-�������շѼ����ʵ�)
    rsTmp.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��������", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�Ƿ񻮼�", adBigInt, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    intStartPage = IIf(intPage <= 0, 1, intPage)
    intPages = IIf(intPage <= 0, objBill.Pages.Count, intPage)
    For p = intStartPage To intPages
        If objBill.Pages(p).NO <> "" Then '��ȡ���۵�
            strSQL = "Select ����ID, NULL as ��ҳID, �շ����, �շ�ϸĿid, Avg(���� * Nvl(����, 1)) ����," & vbNewLine & _
                    "        Sum(��׼����) As ����, Sum(ʵ�ս��) ʵ�ս��,Max(ִ�в���id) as ִ�в���id" & vbNewLine & _
                    " From ������ü�¼" & vbNewLine & _
                    " Where NO = [1] And ��¼���� = 1 And ��¼״̬ = 0" & vbNewLine & _
                    " Group By �շ�ϸĿid, ����id, �շ����"
            Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���۵�", objBill.Pages(p).NO)
            With rsPrice
                For i = 1 To .RecordCount
                    rsTmp.Filter = "�շ�ϸĿID=" & !�շ�ϸĿID
                    If rsTmp.RecordCount = 0 Then
                        rsTmp.AddNew
                        
                        rsTmp!����ID = Nvl(!����ID, objBill.����ID)
                        rsTmp!��ҳID = Nvl(!��ҳID, objBill.��ҳID)
                        rsTmp!�շ���� = !�շ����
                        rsTmp!�շ�ϸĿID = !�շ�ϸĿID
                        
                        rsTmp!���� = !����
                        rsTmp!���� = !����
                        rsTmp!ʵ�ս�� = !ʵ�ս��
                        
                        rsTmp!������ = str������
                        rsTmp!�������� = str��������
                        rsTmp!ִ�п���ID = !ִ�в���ID
                        rsTmp!�������� = intMode
                        rsTmp!�Ƿ񻮼� = intPrice
                    Else
                        rsTmp!���� = rsTmp!���� + !����
                        rsTmp!���� = (rsTmp!���� + !����) / 2
                        rsTmp!ʵ�ս�� = rsTmp!ʵ�ս�� + !ʵ�ս��
                    End If
                    
                    rsTmp.Update
                    .MoveNext
                Next
            End With
        Else
            If lngRow = 0 Then
                intB = 1
                intE = objBill.Pages(p).Details.Count
            Else
                intB = lngRow
                intE = lngRow
            End If
            
            For i = intB To intE
                dbl���� = 0: curʵ�� = 0
                With objBill.Pages(p).Details(i)
                    If lngRow = 0 Then
                        rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID
                        blnNew = rsTmp.RecordCount = 0
                    Else
                        blnNew = True
                    End If
                                    
                    If blnNew Then
                        rsTmp.AddNew
                        
                        rsTmp!����ID = objBill.����ID
                        rsTmp!��ҳID = objBill.��ҳID
                        
                        rsTmp!�շ���� = .�շ����
                        rsTmp!�շ�ϸĿID = .�շ�ϸĿID
                        rsTmp!ִ�п���ID = .ִ�в���ID
                        rsTmp!�������� = intMode
                        rsTmp!�Ƿ񻮼� = intPrice
                        
                        For j = 1 To .InComes.Count
                            dbl���� = dbl���� + .InComes(j).��׼����
                            curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                        Next
                        If InStr(",5,6,7,", .�շ����) > 0 And gblnҩ����λ Then
                            '��ҩ����λת��Ϊ�ۼ۵�λ
                            rsTmp!���� = IIf(.���� = 0, 1, .����) * .���� * .Detail.ҩ����װ
                            rsTmp!���� = Format(dbl���� / .Detail.ҩ����װ, gstrFeePrecisionFmt)
                        Else
                            rsTmp!���� = IIf(.���� = 0, 1, .����) * .����
                            rsTmp!���� = Format(dbl����, gstrFeePrecisionFmt)
                        End If
                        rsTmp!ʵ�ս�� = Format(curʵ��, gstrDec)
                        
                        rsTmp!������ = str������
                        rsTmp!�������� = str��������
                    Else
                        For j = 1 To .InComes.Count
                            dbl���� = dbl���� + .InComes(j).��׼����
                            curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                        Next
                        If InStr(",5,6,7,", .�շ����) > 0 And gblnҩ����λ Then
                            '��ҩ����λת��Ϊ�ۼ۵�λ
                            rsTmp!���� = rsTmp!���� + IIf(.���� = 0, 1, .����) * .���� * .Detail.ҩ����װ
                            rsTmp!���� = Format((rsTmp!���� + Format(dbl���� / .Detail.ҩ����װ, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                        Else
                            rsTmp!���� = rsTmp!���� + IIf(.���� = 0, 1, .����) * .����
                            rsTmp!���� = Format((rsTmp!���� + Format(dbl����, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                        End If
                        rsTmp!ʵ�ս�� = rsTmp!ʵ�ս�� + Format(curʵ��, gstrDec)
                    End If
                    
                    rsTmp.Update
                End With
            Next
        End If
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetApply(ByVal strNos As String, ByVal bytFlag As Byte) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�˷���������
    '���:strNOs-���ݺ�,����ö��ŷ���
    '     bytFlag-��¼����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-08-05 11:59:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    strSQL = "" & _
    "   Select NO,������,����ʱ��,����ԭ��,�����,���ԭ��,Nvl(״̬,0) As ״̬" & _
    "   From �����˷����� " & _
    "   Where NO IN  (Select Column_value From Table(f_str2List([1]))) And ��¼���� = [2]"
    On Error GoTo errH
    Set GetApply = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNos, bytFlag)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ImportBill(strNo As String, blnModi As Boolean, bytFlag As Byte, _
    Optional ByVal int���� As Integer, Optional ByVal bln������ As Boolean = True, _
    Optional ByVal bln�����巨 As Boolean, Optional ByVal strҩƷ�۸�ȼ� As String, _
    Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ�۸�ȼ� As String) As ExpenseBill
'���ܣ���ȡ���õ��ݵ����ݶ�����(Ŀǰ���Դ�����Ŀ,����������Ŀ),�����޸Ļ���ʱ��
'������
'      strNO=���ݺ�
'      bytFlag=��¼����,'0-�շ�,1-����,2-�������
'      int����=�Ƿ���(�޸�)ҽ���շѵ���(�����֤��ͨ��)
'      bln�����巨   �򵥼��ʵ��޸ĵ���ʱ���ö��巨
'���أ���ŵ�����Ϣ�ĵ��ݶ���
'˵������Ϊ������ʱ��Ŀ�۸���Ϣ��������,���Է�������������¼���
'      �ù��̽������޸Ķ��룬�޸�ʱҪ�ſ��������(�������շ���)
'      �����ǵ��뻹���޸ĵ���,����Ӧ������ͣ���շ�ϸĿ
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim intCurNo As Integer, strInfo As String, strAdvance As String
    Dim int��� As Integer, blnDo As Boolean, i As Integer
    Dim dblAllTime As Double, dblCurTime As Double
    Dim dbl�Ӱ�Ӽ��� As Double
    Dim dblStock As Double, strҩ��IDs As String, strͣ����Ŀ��� As String
    Dim colSerial As New Collection
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim strժҪ As String, strWherePriceGrade
     
    On Error GoTo errH
    '�۸�ȼ�
    If strҩƷ�۸�ȼ� <> "" Or str���ļ۸�ȼ� <> "" Or str��ͨ�۸�ȼ� <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [4])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [5])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And d.�۸�ȼ� = [6])" & vbNewLine & _
            "            Or (d.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From �շѼ�Ŀ" & vbNewLine & _
            "                                Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [4])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [5])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And �۸�ȼ� = [6])))))"
    Else
        strWherePriceGrade = " And d.�۸�ȼ� Is Null"
    End If
    '�շѼ�Ŀ����:�¼���۸�,����ж���۸�,��һ���շ�ϸĿID�оͻ��ж��������ͬ�ļ�¼
    '�۸񸸺� is NULL:ֻȡÿ���շ�ϸĿID�ĵ�һ��(ҩƷֻ��һ��),��ΪҪ����۸�
    strSQL = _
        " Select X.ҩƷID,W.����ID,W.��������," & _
        "       A.��� As ���,A.��������,A.NO,A.��¼����,A.��¼״̬,0 as  �ಡ�˵�,A.Ӥ����,A.�ѱ�,A.����,A.�Ա�,A.����," & _
        "       A.���ʽ ,A.��ʶ��,A.����ID, 0 as ��ҳID,0 as ���˲���ID,A.���˿���ID,A.��������ID,A.�����־,A.�Ӱ��־," & _
        "       A.���ӱ�־,A.�շ����,A.�շ�ϸĿID,A.��ҩ����,Nvl(����,1) as ����,Nvl(A.����,0) as ����," & _
        "       A.��׼����,A.ִ�в���ID,A.������,A.������,A.����Ա���,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��,A.ժҪ," & _
        "       B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(F.����,B.����) as ����,E1.���� as ��Ʒ��,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
        "       B.���ηѱ�,B.˵��,B.ִ�п���,Nvl(A.��������,B.��������) ��������,D.�ּ�,D.ԭ��,D.ȱʡ�۸�,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
        "       E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.�շ����,'4',1,X." & gstrҩ����װ & ") as ҩ����װ," & _
        "       Decode(A.�շ����,'4',B.���㵥λ,X." & gstrҩ����λ & ") as ҩ����λ," & _
        "       Decode(A.�շ����,'4',Nvl(W.���÷���,0),Nvl(X.ҩ������,0)) as ����,Nvl(A.ҽ�����,0) ҽ�����,B.¼������,A.����,M1.���� as ��������" & _
        " From ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E," & _
        "          �շ���Ŀ���� F,�շ���Ŀ���� E1,�������� W,ҩƷ��� X,������ĿĿ¼ M1" & _
        " Where Nvl(A.���ӱ�־,0)<>9 And A.��¼����=[2] And Instr([3],','||A.��¼״̬||',')>0" & _
                IIf(Not bln������, " And Nvl(A.���ӱ�־,0)<>8", "") & " And A.NO=[1]" & _
        "       And A.�۸񸸺� Is Null And A.�շ�ϸĿID=B.ID And A.�շ�ϸĿID=D.�շ�ϸĿID" & _
        "       And A.�շ����=C.���� And D.������ĿID=E.ID And A.�շ�ϸĿID=W.����ID(+) " & _
        "       And A.�շ�ϸĿID=X.ҩƷID(+) And X.ҩ��ID=M1.ID(+)" & _
        "       And A.�շ�ϸĿID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.�����־ " & IIf(gint������Դ = 1, " IN(1,4)", "= 2") & _
        "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "       And Sysdate Between D.ִ������ And Nvl(D.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
            strWherePriceGrade
        
    strSQL = "Select * From (" & strSQL & ") Order by ���"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, IIf(bytFlag = 0, 1, bytFlag), _
        IIf(blnModi, ",0,1,", ",0,1,3,"), strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ�۸�ȼ�)
    
    'û�м�¼���ǿյ���
    Set objBill = New ExpenseBill
    Set objBill.Pages(1).Details = New BillDetails
    If rsTmp.RecordCount <> 0 Then
        With rsTmp
            i = 1
NextRecord: Do While Not .EOF
                '����շ���Ŀ�Ƿ�ͣ�û���������ﲡ��
                '����ͣ��ʱ,��������
                If Not blnModi And InStr(",5,6,7,", !�շ����) = 0 Then
                    If InStr(1, strͣ����Ŀ��� & ",", "," & !�������� & ",") > 0 Then
                        .MoveNext
                        GoTo NextRecord
                    Else
                        If Not CheckFeeItemAvailable(!�շ�ϸĿID, 1) Then
                            strͣ����Ŀ��� = strͣ����Ŀ��� & "," & !���
                            MsgBox "����[" & strNo & "]�е�" & !��� & "���շ���Ŀ:" & !���� & "" & vbCrLf & _
                                "��ͣ�û��ٷ����ڲ���,�����ᱻ����." & IIf(IsNull(!��������), "����д�����Ŀ,Ҳ���ᱻ����.", ""), vbInformation, gstrSysName
                            .MoveNext
                            GoTo NextRecord
                        End If
                    End If
                End If
            
                '����������=====================================================
                If i = 1 Then
                    objBill.NO = !NO
                    objBill.Pages(1).NO = "" 'Ҫ����Ա��޸�ʱ������ֱ������ķ���
                    objBill.Pages(1).��������ID = IIf(IsNull(!��������ID), 0, !��������ID)
                    objBill.Pages(1).������ = IIf(IsNull(!������), "", !������)
                    objBill.Pages(1).ҽ����� = !ҽ�����
                    
                    objBill.����ID = IIf(IsNull(!����ID), 0, !����ID)
                    objBill.��ҳID = IIf(IsNull(!��ҳID), 0, !��ҳID)
                    objBill.����ID = IIf(IsNull(!���˲���ID), 0, !���˲���ID)
                    objBill.����ID = IIf(IsNull(!���˿���id), 0, !���˿���id)
                    objBill.���� = IIf(IsNull(!����), "", !����)
                    objBill.�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
                    objBill.���� = IIf(IsNull(!����), "", !����)
                    objBill.��ʶ�� = IIf(IsNull(!��ʶ��), 0, !��ʶ��)
                    objBill.���� = "" & !���ʽ
                    objBill.�ѱ� = IIf(IsNull(!�ѱ�), "", !�ѱ�)
                    objBill.�����־ = IIf(IsNull(!�����־), 0, !�����־)
                    objBill.�Ӱ��־ = IIf(IsNull(!�Ӱ��־), 0, !�Ӱ��־)
                    objBill.Ӥ���� = IIf(IsNull(!Ӥ����), 0, !Ӥ����)
                    objBill.������ = IIf(IsNull(!������), "", !������)
                    objBill.����Ա��� = IIf(IsNull(!����Ա���), "", !����Ա���)
                    objBill.����Ա���� = IIf(IsNull(!����Ա����), "", !����Ա����)
                    objBill.����ʱ�� = !����ʱ��
                    objBill.�Ǽ�ʱ�� = !�Ǽ�ʱ��
                    objBill.�ಡ�˵� = (IIf(IsNull(!�ಡ�˵�), 0, !�ಡ�˵�) = 1)
                End If
                
                '�����շ�ϸĿ=====================================================
                Set objBillDetail = New BillDetail
                Set objBillDetail.Detail = New Detail
                            
                '�������,��������
                intCurNo = intCurNo + 1
                objBillDetail.��� = intCurNo 'ʵ�����к�
                colSerial.Add intCurNo, "_" & !��� '��¼ԭ������ڵ��к�
                If Not IsNull(!��������) Then
                    objBillDetail.�������� = colSerial("_" & !��������)
                End If
                                                                    
                'ʹ��ԭ���Ķ�̬�ѱ�
                objBillDetail.�ѱ� = IIf(IsNull(!�ѱ�), "", !�ѱ�)
                objBillDetail.�շ���� = !�շ����
                objBillDetail.�շ�ϸĿID = !�շ�ϸĿID
                objBillDetail.���㵥λ = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                
                objBillDetail.���� = Nvl(!����, 1)
                If InStr(",5,6,7,", !�շ����) > 0 And gblnҩ����λ Then
                    objBillDetail.���� = Nvl(!����, 0) / Nvl(!ҩ����װ, 1)
                Else
                    objBillDetail.���� = Nvl(!����, 0)
                End If
                objBillDetail.ԭʼ���� = objBillDetail.���� * objBillDetail.����
                
                objBillDetail.��ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ����)
                
                objBillDetail.���ӱ�־ = IIf(IsNull(!���ӱ�־), 0, !���ӱ�־)
                
                objBillDetail.ժҪ = IIf(IsNull(!ժҪ), "", !ժҪ)
                
                objBillDetail.ִ�в���ID = IIf(IsNull(!ִ�в���ID), 0, !ִ�в���ID)
                
                objBillDetail.ԭʼִ�в���ID = objBillDetail.ִ�в���ID     '�����޸�ʱ�����жϿ��
                
                objBillDetail.Detail.ID = !�շ�ϸĿID
                objBillDetail.Detail.���� = !����
                objBillDetail.Detail.��� = (IIf(IsNull(!�Ƿ���), 0, !�Ƿ���) = 1)
                objBillDetail.Detail.�������� = 0 '!!!Ŀǰ���Դ�����Ŀ,����������Ŀ
                objBillDetail.Detail.���д��� = 0 '!!!Ŀǰ���Դ�����Ŀ,����������Ŀ
                objBillDetail.Detail.��� = IIf(IsNull(!���), "", !���)
                objBillDetail.Detail.���㵥λ = Nvl(!���㵥λ)
                
                objBillDetail.Detail.ҩ����λ = Nvl(!ҩ����λ)
                objBillDetail.Detail.ҩ����װ = Nvl(!ҩ����װ, 1)
                
                If InStr(",4,5,6,7,", !�շ����) > 0 Then
                    dblStock = GetStock(!�շ�ϸĿID, !ִ�в���ID)
                Else
                    dblStock = 0
                End If

                If InStr(",5,6,7,", !�շ����) > 0 And gblnҩ����λ Then dblStock = dblStock / Nvl(!ҩ����װ, 1)
                If blnModi Then
                    If InStr(",5,6,7,", !�շ����) > 0 Or !�շ���� = "4" And Nvl(!��������, 0) = 1 Then dblStock = dblStock + objBillDetail.ԭʼ����
                End If
                objBillDetail.Detail.��� = dblStock
                
                objBillDetail.Detail.�Ӱ�Ӽ� = (IIf(IsNull(!�Ӱ�Ӽ�), 0, !�Ӱ�Ӽ�) = 1)
                objBillDetail.Detail.��� = IIf(IsNull(!���), "", !���)
                objBillDetail.Detail.������� = IIf(IsNull(!�������), "", !�������)
                objBillDetail.Detail.���� = IIf(IsNull(!����), "", !����)
                objBillDetail.Detail.��Ʒ�� = Nvl(!��Ʒ��)
                objBillDetail.Detail.���ηѱ� = (IIf(IsNull(!���ηѱ�), 0, !���ηѱ�) = 1)
                objBillDetail.Detail.˵�� = IIf(IsNull(!˵��), "", !˵��)
                objBillDetail.Detail.ִ�п��� = IIf(IsNull(!ִ�п���), 0, !ִ�п���)
                objBillDetail.Detail.���� = IIf(IsNull(!��������), "", !��������)
                objBillDetail.Detail.�������� = Nvl(!��������)
                objBillDetail.Detail.��ҩ��̬ = Val(Nvl(!����))
                
                If InStr(",5,6,7,", !�շ����) > 0 Then
                    objBillDetail.Detail.����ְ�� = Get����ְ��(objBillDetail.Detail.ID)
                    objBillDetail.Detail.�������� = Get��������(objBillDetail.Detail.ID)
                End If
                objBillDetail.Detail.¼������ = Val("" & !¼������)
                
                objBillDetail.Detail.ҩ��ID = IIf(IsNull(!ҩ��ID), 0, !ҩ��ID)
                objBillDetail.Detail.��� = IIf(IsNull(!�Ƿ���), 0, !�Ƿ���) = 1
                objBillDetail.Detail.���� = IIf(IsNull(!����), 0, !����) = 1
                objBillDetail.Detail.�������� = Nvl(!��������, 0) = 1
                
                '����:41136
                strժҪ = objBillDetail.ժҪ
                If Not blnModi Then '90304
                    strժҪ = gclsInsure.GetItemInfo(int����, objBill.����ID, objBillDetail.�շ�ϸĿID, strժҪ, 1, , "|1")
                    objBillDetail.ժҪ = strժҪ
                End If
                
                '����۸񲿷�=====================================================
                Set objBillDetail.InComes = New BillInComes
                Do
                    '�������еļ۸��������¼���
                    If InStr(",5,6,7,", !�շ����) > 0 Or (!�շ���� = "4" And Nvl(!��������, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        'ʱ��ҩƷ����۸�(�����ɲ�����)
                        dblAllTime = !���� * !���� '�������ۼ�����
                        If dblAllTime <> 0 Or Nvl(!�Ƿ���, 0) = 0 Then
                            Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                        "��ȡҩƷ��ǰ�ۼ�", CLng(!�շ�ϸĿID), objBillDetail.ִ�в���ID, dblAllTime)
                            If rsPrice.EOF Then
                                '��ȡ�۸�ʧ��
                                If !�շ���� = "4" Then
                                    MsgBox "��������""" & Nvl(!����) & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                                Else
                                    MsgBox "ҩƷ""" & Nvl(!����) & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                                End If
                                objBillIncome.��׼���� = 0
                            Else
                                strPrice = Nvl(rsPrice!Price) & "|||"
                                varPrice = Split(strPrice, "|")
                                objBillIncome.��׼���� = Val(varPrice(0))
                                dblʣ������ = Val(varPrice(2))
                                
                                If dblʣ������ <> 0 And Nvl(!�Ƿ���, 0) = 1 Then
                                    '����δ�ֽ����
                                    If !�շ���� = "4" Then
                                        MsgBox "ʱ����������""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    Else
                                        MsgBox "ʱ��ҩƷ""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.��׼���� = 0
                                End If
                            End If
                        Else
                            objBillIncome.��׼���� = 0
                        End If
                    ElseIf Nvl(!�Ƿ���, 0) = 1 Then
                        If Abs(!��׼����) > Abs(Val(Nvl(!�ּ�))) Then
                            objBillIncome.��׼���� = Val(Nvl(!ȱʡ�۸�))
                        Else
                            objBillIncome.��׼���� = !��׼����
                        End If
                    Else
                        objBillIncome.��׼���� = !�ּ�
                    End If
                                        
                    If InStr(",5,6,7,", !�շ����) > 0 And gblnҩ����λ Then
                        objBillIncome.��׼���� = Format(objBillIncome.��׼���� * Nvl(!ҩ����װ, 1), gstrFeePrecisionFmt)
                    Else
                        objBillIncome.��׼���� = Format(objBillIncome.��׼����, gstrFeePrecisionFmt)
                    End If
                    objBillIncome.�ּ� = IIf(IsNull(!�ּ�), 0, !�ּ�) '�ּ�ԭ�۶�ҩƷ�������
                    objBillIncome.ԭ�� = IIf(IsNull(!ԭ��), 0, !ԭ��)
                    objBillIncome.������ĿID = IIf(IsNull(!������ID), 0, !������ID)
                    objBillIncome.������Ŀ = IIf(IsNull(!������Ŀ), "", !������Ŀ)
                    objBillIncome.�վݷ�Ŀ = IIf(IsNull(!�ַ�Ŀ), "", !�ַ�Ŀ)
                    
                    'Ӧ�ս��=����*����*����
                    objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * objBillDetail.���� * objBillDetail.����
                    
                    '�������������ü���(����������Ŀ)
                    If IIf(IsNull(!���ӱ�־), 0, !���ӱ�־) = 1 And IIf(IsNull(!�շ����), "", !�շ����) = "F" Then
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * IIf(IsNull(!�����շ���), 1, !�����շ��� / 100)
                    End If
                    
                    '�Ӱ�����ʼ���
                    dbl�Ӱ�Ӽ��� = 0
                    If IIf(IsNull(!�Ӱ��־), 0, !�Ӱ��־) = 1 And IIf(IsNull(!�Ӱ�Ӽ�), 0, !�Ӱ�Ӽ�) = 1 Then
                        dbl�Ӱ�Ӽ��� = IIf(IsNull(!�Ӱ�Ӽ���), 0, !�Ӱ�Ӽ��� / 100)
                        objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� + objBillIncome.Ӧ�ս�� * dbl�Ӱ�Ӽ���
                    End If
                    objBillIncome.Ӧ�ս�� = Format(objBillIncome.Ӧ�ս��, gstrDec)
                    
                    '����ʵ�ս��
                    If IIf(IsNull(!���ηѱ�), 0, !���ηѱ�) = 1 Then
                        objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                    Else
                        'ʹ��ԭ���Ķ�̬�ѱ�
                        objBillIncome.ʵ�ս�� = ActualMoney(objBillDetail.�ѱ�, !������ID, objBillIncome.Ӧ�ս��, _
                            objBillDetail.�շ�ϸĿID, objBillDetail.ִ�в���ID, objBillDetail.ԭʼ����, dbl�Ӱ�Ӽ���)
                    End If
                    
                    With objBillIncome
                        '��ȡ��Ŀ������Ϣ,��ҽ�����˲���
                        If int���� <> 0 And bytFlag = 0 Then
                            strAdvance = objBillDetail.ժҪ & "||" & objBillDetail.ԭʼ����
                            strInfo = gclsInsure.GetItemInsure(objBill.����ID, objBillDetail.�շ�ϸĿID, .ʵ�ս��, True, int����, strAdvance)
                            If strInfo <> "" Then
                                objBillDetail.������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                                objBillDetail.���մ���ID = Val(Split(strInfo, ";")(1))
                                .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                                objBillDetail.���ձ��� = CStr(Split(strInfo, ";")(3))
                                
                                If UBound(Split(strInfo, ";")) >= 4 Then
                                    If CStr(Split(strInfo, ";")(4)) <> "" Then objBillDetail.ժҪ = CStr(Split(strInfo, ";")(4))
                                    If UBound(Split(strInfo, ";")) >= 5 Then
                                        If Split(strInfo, ";")(5) <> "" Then objBillDetail.Detail.���� = Split(strInfo, ";")(5)
                                    End If
                                End If
                            End If
                        End If
                    
                        objBillDetail.InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
                    End With
                    
                    '�ж���һ����¼�Ƿ����ڵ�ǰ��
                    blnDo = False
                    int��� = !���
                    .MoveNext
                    If Not .EOF Then blnDo = (int��� = !���)
                    i = i + 1
                Loop While blnDo And Not .EOF
               
                With objBillDetail
                    objBill.Pages(1).Details.Add .�ѱ�, .Detail, .�շ�ϸĿID, .���, .��������, .�շ����, .���㵥λ, .��ҩ����, _
                        .����, .����, .���ӱ�־, .ִ�в���ID, .InComes, , .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ, .ԭʼ����, .ԭʼִ�в���ID
                    
                    '���ù�����
                    If objBill.Pages(1).Details(objBill.Pages(1).Details.Count).���ӱ�־ = 8 Then
                        objBill.Pages(1).Details(objBill.Pages(1).Details.Count).������ = True
                    End If
                End With
            Loop
        End With
        
        '��ȡ��ҩ�巨
        If Not bln�����巨 Then
            strSQL = "Select ��� From ҩƷ�շ���¼ Where NO=[1] And ����=[2]"  '8-�շѴ�����ҩ��9-���ʵ�������ҩ��
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, IIf(bytFlag = 0 Or bytFlag = 1, 8, 9))
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!���) Then
                    objBill.Pages(1).�巨 = rsTmp!���
                End If
            End If
        End If
        
    End If
    
    Set ImportBill = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BillOnlyFactMoney(ByVal strNo As String) As Boolean
'���ܣ��ж�һ���շѵ����Ƿ���й�����
'������strNO=F0000001,����'��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(ID) as ����,Sum(Decode(Nvl(���ӱ�־,0),8,1,0)) as ������ From ������ü�¼" & _
        " Where ��¼����=1 And ��¼״̬=1 And �۸񸸺� is Null And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "OutExse", strNo)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!������, 0) = 1 And Nvl(rsTmp!����, 0) = 1 Then
            BillOnlyFactMoney = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMaxFact(ByVal strNo As String) As String
'���ܣ���ȡָ���շѵ���(����Ϊ�����е�һ��)���������Ʊ�ݺ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    
    'Ӧȡ���һ�δ�ӡ��������
    strSQL = "Select Max(ID) From Ʊ�ݴ�ӡ���� Where ��������=1 And NO=[1]"
    strSQL = "Select Max(A.����) as ���� " & _
        " From Ʊ��ʹ����ϸ A,Ʊ�ݴ�ӡ���� B" & _
        " Where B.��������=1 And B.ID=(" & strSQL & ")" & _
        " And A.��ӡID=B.ID And A.Ʊ��=1 And A.����=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then GetMaxFact = Nvl(rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function RePrintCharge(ByVal bytType As Byte, strNos As String, frmParent As Object, _
                            ByRef lng����ID As Long, ByVal strReclaimInvoice As String, _
                            Optional blnDelOpt As Boolean, Optional DateDel As Date, _
                            Optional intPrintFormat As Integer, Optional blnVirtualPrint As Boolean, _
                            Optional ByVal blnDelRecord As Boolean, _
                            Optional lngShareUseID As Long, Optional strUseType As String = "", _
                            Optional blnOnePatiPrint As Boolean = False, _
                            Optional strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ǰ�տ��¼���´�ӡһ��Ʊ��
    '���:1-�ش�;2-����
    '       strNOs -ָ��Ҫ�ش�ĵ��ݺţ������ţ������Ƕ�����ݺţ�Ϊ"'AAA','BBB',..."����ʽ
    '       lng����ID-�ϴ�ʹ�õ�����ID,��һ��ʹ�û����嵥����������ʱû��
    '       strReclaimInvoice-ʵ���ջصķ�Ʊ��(ֻ��Ʊ�ݷ������Ϊ1��2�Ŵ���)
    '       blnDelOpt-�˷��ش��������
    '       DateDel-�˷�ʱ��
    '       intPrintFormat-��ӡ��ʽ���
    '       intPrintOldFormat-�ϰ�Ʊ�ݴ�ӡ��ʽ
    '       blnVirtualPrint-ҽ���ӿڴ�ӡƱ�ݣ�HIS������ӡֻ��Ʊ��
    '       blnDelRecord-�ش�ʱ���Ƿ��Ƕ��˷Ѽ�¼�����ش�(Ŀǰֻ�б���ҽ��(ҽ���ӿڴ�ӡƱ��)������)
    '       lngShareUseID-����Ʊ������ID(27559)
    '       blnOnePatiPrint-������һ�δ�ӡ
    '       strPriceGrade-�۸�ȼ������ڼ��㹤����
    '����:��ӡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 11:50:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInvoice As String, strInfo As String
    Dim j As Integer, blnValid As Boolean, blnInput As Boolean
    Dim blnDo As Boolean, blnHaveInvoice As Boolean
    Dim strRptName As String, int��Ʊ���� As Integer
    Dim bytPrintType As Byte, lng��ӡID As Long
    Dim bln�ֱ��ӡ As Boolean
    
    blnHaveInvoice = lng����ID <> 0     '��Ҫ���˷���,����������õ�,�������շ�Ʊ,Ȼ���ش�Ʊ:30386
    
    If blnHaveInvoice = False And blnDelOpt Then
        blnHaveInvoice = zlCheckIsPrintInvoice(strNos)
    End If
    
    strRptName = "ZL" & glngSys \ 100 & "_BILL_1121_1"
   
    '����ϸ����Ʊ��ʹ��
    If gblnStrictCtrl Then
        '��ʱֻ�ж��Ƿ���,��ӡ֮ǰ�ٸ��������ж��Ƿ���
        int��Ʊ���� = 1
        If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ And blnOnePatiPrint = False Then
            int��Ʊ���� = UBound(Split(strNos, ",")) + 1
            If int��Ʊ���� = -1 Then int��Ʊ���� = 1
        End If
        If zlCheckInvoiceValied(lng����ID, int��Ʊ����, strInvoice, lngShareUseID, strUseType) = False Then Exit Function
    End If

    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        If blnDelOpt Then
            blnDo = True
        Else
            If intPrintFormat = 0 Then   '��ȱʡƱ�ݸ�ʽ��ʾ
                intPrintFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
            End If
            SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", intPrintFormat
            '����û�и�ʽ�Ĵ���,���,��Ҫǿ��ȱʡ��ָ����ʽ
            blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
            'ȡ��ѡ��ĸ�ʽ
            intPrintFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
    End If
    
    If blnDo Then
        'ȡ��һ��Ʊ�ݺ���
        If Not gblnStrictCtrl Then
            
            '�п����ǵ�һ��ʹ��
            Do
                blnInput = False
                '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
                strInvoice = UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1121, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                    vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlstr.Increase(strInvoice)
                    strInvoice = UCase(InputBox("��ȷ��" & IIf(bytType = 1, "�ش�", "����") & "ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
                    
                '�û�ȡ������,�����ӡ
                If strInvoice = "" Then
                    If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnValid = True
                Else
                    '���������Ч��
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> gbytFactLength Then
                            MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '����Ʊ�����ö�ȡ
                blnInput = False
                strInvoice = GetNextBill(lng����ID)
                If strInvoice = "" Then
                    '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                    '30386:��ӡ�˷�Ʊ��,�����ش��ٷ���
                    If frmInputBox.InputBox(frmParent, "��ʼ��Ʊ��", "�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                    vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                Else
                    '30386
                    If frmInputBox.InputBox(frmParent, "��ʼ��Ʊ��", "��ȷ��" & IIf(bytType = 1, "�ش�", "����") & "ʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                End If
                
                '�û�ȡ������,����ӡ
                If strInvoice = "" Then Exit Function
                
                '���������Ч��
                If blnInput Then
                    If zlCheckInvoiceValied(lng����ID, 1, strInvoice, lngShareUseID, strUseType) Then blnValid = True
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If
        
        bytPrintType = IIf(blnDelOpt, 3, 2)
        If bytType = 2 And gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
            bytPrintType = 4  '����
        End If
        
        
        If blnOnePatiPrint Then
            '�����ٽ�����
            If zlSaveTempPrintData(Replace(strNos, "'", ""), lng����ID, strInvoice, lng��ӡID) = False Then Exit Function
        End If
        
        bln�ֱ��ӡ = gTy_Module_Para.bln�ֱ��ӡ And Check��쵥��(Replace(strNos, "'", "")) = False And blnOnePatiPrint = False
        
        '1-�µ���ӡ,2-�ش�,3-�˷Ѵ�ӡ,4-����Ʊ��(ֻ��:2-��ϵͳԤ�������3-�û��Զ�����ʱ��ת��)
        If blnDelOpt Then
            Call frmPrint.ReportPrint(bytPrintType, strNos, "", strReclaimInvoice, lng����ID, lngShareUseID, strInvoice, _
                DateDel, , , bln�ֱ��ӡ, intPrintFormat, blnVirtualPrint, , strUseType, , blnOnePatiPrint, lng��ӡID, strPriceGrade)
        Else
            Call frmPrint.ReportPrint(bytPrintType, strNos, "", strReclaimInvoice, lng����ID, lngShareUseID, strInvoice, _
                zlDatabase.Currentdate, , , bln�ֱ��ӡ, intPrintFormat, blnVirtualPrint, blnDelRecord, strUseType, , _
                blnOnePatiPrint, lng��ӡID, strPriceGrade)
        End If
        RePrintCharge = True
    End If
End Function

Public Function Check��쵥��(ByVal strNos As String) As Boolean
    'ֻҪ��һ�������,����Ϊȫ������쵥��
    'strNOs -ָ��Ҫ�ش�ĵ��ݺţ������Ƕ�����ݺţ�Ϊ"AAA,BBB,..."����ʽ
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errH:
    If strNos = "" Then Exit Function
    strSQL = "Select /*+cardinality(b,10)*/ 1" & vbNewLine & _
            " From ������ü�¼ A, Table(Cast(f_Str2list([1]) As t_Strlist)) B" & vbNewLine & _
            " Where a.No = b.Column_Value And Mod(a.��¼����, 10) = 1 And a.�����־ = 4 And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ����������", strNos)
    Check��쵥�� = Not rsTemp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintDelCharge(ByVal lng������� As Long, frmParent As Object, _
                            ByRef lng����ID As Long, Optional ByVal blnDelOpt As Boolean, Optional DateDel As Date, _
                            Optional intPrintFormat As Integer, Optional blnVirtualPrint As Boolean, _
                            Optional ByVal blnDelRecord As Boolean, _
                            Optional lngShareUseID As Long, Optional strUseType As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ�˷ѷ�Ʊ(��Ʊ)
    '���:
    '       lng����ID -ָ��Ҫ�ش�ĵ��ݵĽ������
    '       lng����ID-�ϴ�ʹ�õ�����ID,��һ��ʹ�û����嵥����������ʱû��
    '       DateDel-�˷�ʱ��
    '       blnDelOpt-�Ƿ����ش�򲹴����
    '       intPrintFormat-��ӡ��ʽ���
    '       blnVirtualPrint-ҽ���ӿڴ�ӡƱ�ݣ�HIS������ӡֻ��Ʊ��
    '       blnDelRecord-�ش�ʱ���Ƿ��Ƕ��˷Ѽ�¼�����ش�(Ŀǰֻ�б���ҽ��(ҽ���ӿڴ�ӡƱ��)������)
    '       lngShareUseID-����Ʊ������ID(27559)
    '����:��ӡ�ɹ�,����true,���򷵻�False
    '����:
    '����:2016-05-27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInvoice As String, strInfo As String
    Dim j As Integer, blnValid As Boolean, blnInput As Boolean
    Dim blnDo As Boolean, blnHaveInvoice As Boolean
    Dim lng����IDTemp As Long
    Dim strRptName As String
    Dim bytPrintType As Byte, lng��ӡID As Long
    Dim bln�ֱ��ӡ As Boolean
    
    strRptName = "ZL" & glngSys \ 100 & "_BILL_1121_7"
   
    '����ϸ����Ʊ��ʹ��
    If gblnStrictCtrl Then
        '��ʱֻ�ж��Ƿ���,��ӡ֮ǰ�ٸ��������ж��Ƿ���
        lng����ID = GetInvoiceGroupID(1, 1, lng����ID, lngShareUseID, , strUseType)
        Select Case lng����ID
            Case -1
                MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
        End Select
        If lng����ID <= 0 Then Exit Function
    End If

    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        If blnDelOpt Then
            blnDo = True
        Else
            If intPrintFormat = 0 Then   '��ȱʡƱ�ݸ�ʽ��ʾ
                intPrintFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
            End If
            SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", intPrintFormat
            '����û�и�ʽ�Ĵ���,���,��Ҫǿ��ȱʡ��ָ����ʽ
            blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
            'ȡ��ѡ��ĸ�ʽ
            intPrintFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
    End If
    
    If blnDo Then
        'ȡ��һ��Ʊ�ݺ���
        If Not gblnStrictCtrl Then
            '�п����ǵ�һ��ʹ��
            Do
                blnInput = False
                '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
                strInvoice = UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1121, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                    vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlCommFun.IncStr(strInvoice)
                    strInvoice = UCase(InputBox("��ȷ��ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
                    
                '�û�ȡ������,�����ӡ
                If strInvoice = "" Then
                    If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnValid = True
                Else
                    '���������Ч��
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> gbytFactLength Then
                            MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '����Ʊ�����ö�ȡ
                blnInput = False
                strInvoice = GetNextBill(lng����ID)
                If strInvoice = "" Then
                    '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                    '30386:��ӡ�˷�Ʊ��,�����ش��ٷ���
                    If frmInputBox.InputBox(frmParent, "��ʼ��Ʊ��", "�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                    vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                                     blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                Else
                    '30386
                    If frmInputBox.InputBox(frmParent, "��ʼ��Ʊ��", "��ȷ��ʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                                    blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                End If
                
                '�û�ȡ������,����ӡ
                If strInvoice = "" Then Exit Function
                
                '���������Ч��
                If blnInput Then
                    lng����IDTemp = GetInvoiceGroupID(1, 1, lng����ID, lngShareUseID, strInvoice, strUseType)
                    If lng����IDTemp = -3 Then
                        MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                    Else
                        lng����ID = lng����IDTemp
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If
        
        '1-�µ���ӡ,2-�ش�,3-�˷Ѵ�ӡ,4-����Ʊ��(ֻ��:2-��ϵͳԤ�������3-�û��Զ�����ʱ��ת��),6-�˷�Ʊ��(��Ʊ)��ӡ
        If blnDelOpt Then
            Call frmPrint.ReportPrint(6, lng�������, "", "", lng����ID, lngShareUseID, strInvoice, _
                DateDel, , , False, intPrintFormat, blnVirtualPrint, blnDelRecord, strUseType)
        Else
            Call frmPrint.ReportPrint(6, lng�������, "", "", lng����ID, lngShareUseID, strInvoice, _
                 zlDatabase.Currentdate, , , False, intPrintFormat, blnVirtualPrint, blnDelRecord, strUseType)
        End If
        PrintDelCharge = True
    End If
End Function

Public Function zlGetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "", _
    Optional ByRef lngRestNum As Long, _
    Optional ByRef lngNextUseID As Long, _
    Optional ByRef bytErr As Byte) As Long
    '���ܣ���ȡ�������ò���ָ��Ʊ��������÷�Χ�ڵ�����ID
    '��Σ�
    '    bytKind        =   Ʊ��
    '    intNum         =   Ҫ��ӡ��Ʊ������
    '    lngLastUseID   =   �ϴ�ʹ�õ�����ID
    '    lngShareUseID  =   ���ز���ָ���Ĺ���ID
    '    strBill        =   ��ǰƱ�ݺţ����ڼ���������ε�Ʊ�ݷ�Χ
    '    strUseType     =   ʹ�����
    '���Σ�
    '    lngRestNum     =   �ϴ�ʹ������ʣ���Ʊ����
    '    lngNextUseID   =   ��һ���������ε�����ID
    '    bytErr         =
    '                    1 - û������(���������������һ��Ҳ��������δ����),δ���ù���
    '                    2 - û������(����򲻹�����δ����),���õĹ���������򲻹�
    '                    3 - ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
    '                    4 - ָ�����ε�Ʊ��������
    '                    5 - ָ�����ε�Ʊ�ݲ�����,��һ����������������Ʊ��,����ID��¼��lngNextUseID��
    '                    6 - ָ�����ε�Ʊ�ݲ�����,��һ�����������ǹ���Ʊ��,����ID��¼��lngNextUseID��
    '���أ�
    '    >0   =   �ɹ������õ�����ID
    '    =0   =   ʧ��
    '�޸�:Ƚ����,�Զ�����Ʊ������,���ٷ�Ʊ���˷�
    '�޸�����:2015-04-21
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    Dim lngShortCount As Long '��ǰ����ȱ��Ʊ������������һ������ȡ
    Dim blnNext As Boolean '�Ƿ���һ������
    
    On Error GoTo errH
    lngShortCount = intNum
    lngRestNum = 0: lngNextUseID = 0: bytErr = 0
    '1.�ϴε����������Ƿ���ò�����
    If lngLastUseID > 0 Then
        strSQL = "" & _
        "   Select ǰ׺�ı�,��ʼ����,��ֹ����,ʣ������" & _
        "   From Ʊ�����ü�¼ " & _
        "   Where Ʊ��=[1] And ʣ������>0 And ID=[2]  " & _
        "           And (Nvl(ʹ�����,'LXH')=[3] Or  ʹ����� Is NULL) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, lngLastUseID, IIf(Trim(strUseType) = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then 'Ŀǰ��Ʊ�ݺſ��ܺ��ϴβ�ͬ��������Ҫ��鷶Χ
                lngShortCount = lngShortCount - Val(Nvl(!ʣ������))
                If strBill <> "" Then
                    blnTmp = False
                    strPre = "" & !ǰ׺�ı�
                    If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                        blnTmp = True
                    End If
                End If
                If strBill = "" Or strBill <> "" And Not blnTmp Then
                    blnNext = True
                    lngRestNum = Val(Nvl(!ʣ������))
                    zlGetInvoiceGroupID = lngLastUseID
                    bytErr = IIf(lngRestNum < intNum, 5, 0)
'                    Exit Function'�������۵�ǰ�����Ƿ��㹻�����˳���
                                    '����ΪԤ�������Ʊ�ż��ʱ����֪����Ҫ������Ʊ�ݣ�ʼ�մ������Ķ���1�ţ�
                                    '�޷�֪����ǰ�����Ƿ��㹻������ʼ��ȡ����һ�����Σ������У�
                Else
                    bytErr = 3
                End If
            ElseIf intNum > 1 Then  '����ȷ���������ε���ʱ,��ǰָ������������
                blnNext = True
                lngRestNum = Val(Nvl(!ʣ������))
                zlGetInvoiceGroupID = lngLastUseID
                bytErr = 4
            End If
        End With
    End If
    
    '2.�ϴε��������β����û򲻿���ʱ,ȡ������Ĳ������õ�
    '  �ж��������ʹ�õ�����,�ٵ�����,��������
    strSQL = "" & _
    "   Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������" & _
    "   From Ʊ�����ü�¼" & _
    "   Where Ʊ�� = [1] And ʣ������ >0 And ������ = [2]  " & _
    "           And (Nvl(ʹ�����,'LXH')=[3] Or  ʹ����� Is NULL ) " & _
    "           And ʹ�÷�ʽ = 1" & _
    IIf(lngLastUseID > 0, "           And ID <> [4]", "") & _
    "   Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,ʹ����� desc, ��ʼ����"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, UserInfo.����, IIf(strUseType = "", "LXH", strUseType), lngLastUseID)
    With rsTmp
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                lngShortCount = lngShortCount - Nvl(!ʣ������)
                If strBill <> "" Then
                    blnTmp = False
                    strPre = "" & !ǰ׺�ı�
                    If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                        blnTmp = True
                    End If
                End If
                If blnNext Or strBill = "" Or strBill <> "" And Not blnTmp Then '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
                    If blnNext Then
                        If lngShortCount > 0 Then
                            bytErr = 5
                        Else
                            bytErr = IIf(lngRestNum < intNum, 5, bytErr)
                            lngNextUseID = Nvl(!ID)
                            Exit Function
                        End If
                    Else
                        blnNext = True
                        lngRestNum = Val(Nvl(!ʣ������))
                        zlGetInvoiceGroupID = Nvl(!ID)
                        bytErr = IIf(lngRestNum < intNum, 5, 0)
'                        Exit Function'����������������Ƿ��㹻�����˳���
                                    '����ΪԤ�������Ʊ�ż��ʱ����֪����Ҫ������Ʊ�ݣ�ʼ�մ������Ķ���1�ţ�
                                    '�޷�֪����ǰ�����Ƿ��㹻������ʼ��ȡ����һ�����Σ������У�
                    End If
                Else
                    bytErr = 3
                End If
                .MoveNext
            Next
        Else
            bytErr = IIf(lngShortCount > 0, IIf(blnNext, 5, 1), bytErr)
        End If
    End With
        
    '3.û�����õ�,ʹ�ñ��ز���ָ���Ĺ�������
    If lngShareUseID > 0 And lngShareUseID <> lngLastUseID Then
        strSQL = "" & _
        "   Select ǰ׺�ı�,��ʼ����,��ֹ����,ʣ������" & _
        "   From Ʊ�����ü�¼  " & _
        "   Where Ʊ��=[1] And ʣ������>0 And ID=[2] " & _
        "   And (Nvl(ʹ�����,'LXH')=[3] Or  ʹ����� Is NULL) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ������", bytKind, lngShareUseID, IIf(strUseType = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then
                lngShortCount = lngShortCount - Nvl(!ʣ������)
                If strBill <> "" Then
                    blnTmp = False
                    strPre = "" & !ǰ׺�ı�
                    If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(!��ʼ����) And UCase(strBill) <= UCase(!��ֹ����) And Len(strBill) = Len(!��ʼ����)) Then
                        blnTmp = True
                    End If
                End If
                If blnNext Or strBill = "" Or strBill <> "" And Not blnTmp Then '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
                    If blnNext Then
                        If lngShortCount > 0 Then
                            bytErr = 6
                        Else
                            bytErr = IIf(lngRestNum < intNum, 6, bytErr)
                            lngNextUseID = lngShareUseID
                            Exit Function
                        End If
                    Else
                        If lngShortCount > 0 Then
                            bytErr = 2
                        Else
                            bytErr = 0
                            lngRestNum = Val(Nvl(!ʣ������))
                            zlGetInvoiceGroupID = lngShareUseID
                            Exit Function
                        End If
                    End If
                Else
                    bytErr = 3
                End If
            Else
                bytErr = 2
            End If
        End With
    Else
        bytErr = IIf(lngShortCount > 0, IIf(blnNext, 6, 1), bytErr)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlCheckInvoiceValied(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "", _
    Optional ByVal lngShareUseID As Long, Optional strUseType As String = "", _
    Optional ByRef lngRestNum As Long, _
    Optional ByRef lngNextUseID As Long, Optional ByRef strNextInvoiceNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:
    '    lng����ID = ����id
    '    intNum = ҳ��
    '    strInvoiceNO = ����ķ�Ʊ��
    '    lngShareUseID = ���ز���ָ���Ĺ���ID
    '    strUseType = ʹ�����
    '����:lng����ID-����ID
    '    lngRestNum     =   �ϴ�ʹ������ʣ���Ʊ����
    '    lngNextUseID = ��һ���������ε�����ID
    '    strNextInvoiceNO = ��һ���������ε���һ��Ʊ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 15:36:59
    '�޸�:Ƚ����
    '�޸�����:2015-04-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim bytErr As Byte
    
    strNextInvoiceNO = ""
    lng����ID = zlGetInvoiceGroupID(1, intNum, lng����ID, lngShareUseID, strInvoiceNO, strUseType, _
                                        lngRestNum, lngNextUseID, bytErr)
    If lngNextUseID <> 0 Then strNextInvoiceNO = GetNextBill(lngNextUseID)
    If lng����ID > 0 And bytErr = 0 Then zlCheckInvoiceValied = True: Exit Function
    'bytErr =
    '         1 - û������(���������������һ��Ҳ��������δ����),δ���ù���
    '         2 - û������(����򲻹�����δ����),���õĹ���������򲻹�
    '         3 - ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
    '         4 - ָ�����ε�Ʊ��������
    '         5 - ָ�����ε�Ʊ�ݲ�����,��һ����������������Ʊ��,����ID��¼��lngNextUseID��
    '         6 - ָ�����ε�Ʊ�ݲ�����,��һ�����������ǹ���Ʊ��,����ID��¼��lngNextUseID��
    Select Case bytErr
        Case 1
            If Trim(strUseType) = "" Then
                MsgBox "��û�����ú͹��õ��շ�Ʊ�ݣ���������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "��û�����ú͹��õġ�" & strUseType & "���շ�Ʊ�ݣ���������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
        Case 2
            If Trim(strUseType) = "" Then
                MsgBox "���صĹ���Ʊ���Ѿ����꣬��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "���صĹ���Ʊ�ݵġ�" & strUseType & "���շ�Ʊ���Ѿ����꣬��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
        Case 3
            MsgBox "��ǰƱ�ݺ��� " & strInvoiceNO & " ���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ����������룡", vbInformation, gstrSysName
        Case 4 'ָ�����ε�Ʊ��������
            If strNextInvoiceNO <> "" Then
                If MsgBox("��ǰ����Ʊ����ʹ���꣬�Ƿ�ʹ�ÿ�ʼƱ��Ϊ��" & strNextInvoiceNO & "������һ��Ʊ��������ɴ�ӡ��" & vbCrLf & vbCrLf & _
                    "ע�⣺��˶���һ��Ʊ�����εĿ�ʼƱ���Ƿ�Ϊ��" & strNextInvoiceNO & "���������ǣ���ѡ�񡰷񡱡���ѡ���ǡ����뼰ʱ������Ʊ��", _
                    vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    zlCheckInvoiceValied = True: Exit Function
                End If
            End If
        Case 5, 6
            If strNextInvoiceNO <> "" Then
                If MsgBox("��ǰ��ӡ����Ҫ " & intNum & " ��Ʊ�ݣ�����ǰƱ�ݺ��� " & strInvoiceNO & " �������ε���ЧƱ��ֻ�� " & lngRestNum & " �š�" & vbCrLf & _
                    "�Ƿ�ʹ�ÿ�ʼƱ��Ϊ��" & strNextInvoiceNO & "������һ��" & IIf(bytErr = 5, "����Ʊ������", "����Ʊ������") & "��ɴ�ӡ��" & vbCrLf & vbCrLf & _
                    "ע�⣺��˶���һ��Ʊ�����εĿ�ʼƱ���Ƿ�Ϊ��" & strNextInvoiceNO & "���������ǣ���ѡ�񡰷񡱡���ѡ���ǡ����뼰ʱ������Ʊ��", _
                    vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    zlCheckInvoiceValied = True: Exit Function
                End If
            Else 'û����һ�����õ�����Ʊ����
                MsgBox "��ǰƱ��ʣ������ " & lngRestNum & " �Ų��㱾�δ�ӡ�������� " & intNum & " �ţ���������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݡ�", vbInformation, gstrSysName
            End If
        Case Else
            MsgBox "Ʊ��������Ϣ����ʧ�ܣ�����������Խ����ش򵥾ݣ�", vbInformation, gstrSysName
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckDisable(objBill As ExpenseBill) As String
'���ܣ���鵥���е�ҩƷ�Ľ������
'���أ�ҩƷ���������ʾ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim strGroup As String, strIDs As String
    Dim blnStop As Boolean, strMsg As String
    
    Err = 0: On Error GoTo errH:
    For p = 1 To objBill.Pages.Count
        strInfo = "": strIDs = ""
        For i = 1 To objBill.Pages(p).Details.Count
            If InStr(",5,6,7,", objBill.Pages(p).Details(i).�շ����) > 0 Then
                strIDs = strIDs & "," & objBill.Pages(p).Details(i).�շ�ϸĿID
            End If
        Next
        strIDs = Mid(strIDs, 2)
        If Not (strIDs = "" Or UBound(Split(strIDs, ",")) < 1) Then
            strSQL = _
                " Select /*+ RULE */  A.����,Count(Distinct A.��ĿID) as ������" & _
                " From ���ƻ�����Ŀ A,ҩƷ��� B," & _
                "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
                " Where A.��ĿID=B.ҩ��ID And B.ҩƷID  = j.Column_Value" & _
                " Having Count(Distinct A.��ĿID)>1  " & _
                "  Group by A.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strIDs)
            
            If Not rsTmp.EOF Then
                strGroup = ""
                For i = 1 To rsTmp.RecordCount
                    strGroup = strGroup & "," & rsTmp!����
                    rsTmp.MoveNext
                Next
                strGroup = Mid(strGroup, 2)
                
                For i = 0 To UBound(Split(strGroup, ","))
                    strSQL = _
                        "Select /*+ RULE */   Distinct C.����,C.����,D.����,D.����,D.���" & _
                        " From ҩƷ��� A,������ĿĿ¼ B,���ƻ�����Ŀ C,�շ���ĿĿ¼ D," & _
                        "          (Select Column_Value From Table(Cast(f_num2list([2]) As Zltools.t_Numlist ))) J " & _
                        " Where A.ҩ��ID=B.ID And B.ID=C.��ĿID And A.ҩƷID=D.ID" & _
                        "           And C.����=[1]" & _
                        "           And A.ҩƷID  = j.Column_Value" & _
                        " Order by C.����,C.����,D.����"
                        
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", Val(Split(strGroup, ",")(i)), strIDs)
                    If Not rsTmp.EOF Then
                        rsTmp.Filter = "����=1"
                        If rsTmp.RecordCount > 1 Then
                            k = k + 1
                            strInfo = strInfo & vbCrLf & "�� " & k & " ��(��������)��" & vbCrLf
                            For j = 1 To rsTmp.RecordCount
                                strInfo = strInfo & "[" & rsTmp!���� & "]" & rsTmp!���� & IIf(IsNull(rsTmp!���), "", "(" & rsTmp!��� & ")") & "                 " & vbCrLf
                                rsTmp.MoveNext
                            Next
                        End If
                        rsTmp.Filter = "����=2"
                        If rsTmp.RecordCount > 1 Then
                            blnStop = True
                            k = k + 1
                            strInfo = strInfo & vbCrLf & "�� " & k & " ��(�������)��" & vbCrLf
                            For j = 1 To rsTmp.RecordCount
                                strInfo = strInfo & "[" & rsTmp!���� & "]" & rsTmp!���� & IIf(IsNull(rsTmp!���), "", "(" & rsTmp!��� & ")") & "                 " & vbCrLf
                                rsTmp.MoveNext
                            Next
                        End If
                        rsTmp.Filter = 0
                    End If
                Next
                If strInfo <> "" Then
                    If objBill.Pages.Count = 1 Then
                        strMsg = strMsg & vbCrLf & "����������ҩƷ������û����ã�" & vbCrLf & strInfo
                    Else
                        strMsg = strMsg & vbCrLf & "����" & p & "������ҩƷ������û����ã�" & vbCrLf & strInfo
                    End If
                End If
            End If
        End If
    Next
    If strMsg <> "" Then
        If blnStop Then
            strMsg = strMsg & vbCrLf & "���޸Ľ���ҩƷ���ټ�����"
        Else
            strMsg = strMsg & vbCrLf & "Ҫ������"
        End If
    End If
    CheckDisable = Mid(strMsg, 3)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'���ܣ��ж��Ƿ����ָ�������������͵�������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ���ID From ҩƷ�������� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", bytBill)
    If Not rsTmp.EOF Then
        ExistIOClass = IIf(IsNull(rsTmp!���ID), 0, rsTmp!���ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetChargeTotal() As Currency
'���ܣ���ȡ��ǰ����Ա�����ڵ��շ��ܶ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    '�շ��еķ�ҽ��������֮��
    strSQL = "" & _
    "   Select Sum(��Ԥ��) as ��� From ����Ԥ����¼ a" & _
    "   Where ��¼����=3 And ����Ա����=[1]" & _
    "       And �տ�ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate+1)" & _
    "       And Not Exists(Select 'X' From ���㷽ʽ b Where ���� IN(3,4) And a.���㷽ʽ=B.����)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", UserInfo.����)
    If Not rsTmp.EOF Then
        GetChargeTotal = IIf(IsNull(rsTmp!���), 0, rsTmp!���)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub zlInitȱʡ����()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ����ǰ����Ա��ȱʡ����
    '���ƣ����˺�
    '���ڣ�2010-08-16 16:28:21
    '˵����31936
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If gstr��������ID <> "" Then Exit Sub
    strSQL = "Select ����ID From ������Ա Where ��ԱID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ա����������", UserInfo.ID)
    Do While Not rsTmp.EOF
        gstr��������ID = gstr��������ID & "," & Nvl(rsTmp!����ID)
        rsTmp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function GetStockInfo(lngҩƷID As Long, blnҩ�� As Boolean, blnҩ�� As Boolean) As String
'���ܣ���ȡҩƷ�ڸ���ҩ����ҩ��Ŀ����Ϣ
'������"blnҩ��/blnҩ��"����Ҫ��һ������Ϊ��
'���أ�������Ϣ
    Dim strSQL As String, strSQL2 As String, i As Integer
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Call zlInitȱʡ����
    
    If blnҩ�� And blnҩ�� Then
        strSQL = "'��ҩ��','��ҩ��','��ҩ��','��ҩ��','��ҩ��','��ҩ��'"
    ElseIf blnҩ�� Then
        strSQL = "'��ҩ��','��ҩ��','��ҩ��'"
    ElseIf blnҩ�� Then
        strSQL = "'��ҩ��','��ҩ��','��ҩ��'"
    End If
    
    '�ų�������ʵ����,���������סԺ
    strSQL = _
        " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.�������� IN(" & strSQL & ")"
    'ҩ��������ҩƷ����Ч��
    strSQL2 = "Select ����ID From ��������˵�� Where �������� IN('��ҩ��','��ҩ��','��ҩ��')"
    '�����������ҩƷ
    strSQL = _
        " Select B.����,B.����,A.�ⷿID," & _
        " Nvl(Sum(A.��������),0)" & IIf(gblnҩ����λ, "/Nvl(C." & gstrҩ����װ & ",1)", "") & " as ���" & _
        " From ҩƷ��� A,(" & strSQL & ") B,ҩƷ��� C" & _
        " Where A.�ⷿID=B.ID And A.ҩƷID=C.ҩƷID" & _
        " And ((A.Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
        "   Or (Nvl(C.ҩ������,0)=0 And A.�ⷿID IN(" & strSQL2 & ")))" & _
        " And A.����=1 And A.ҩƷID=[1]" & _
        " Group by B.����,B.����,A.�ⷿID,Nvl(C." & gstrҩ����װ & ",1)" & _
        " Having Sum(Nvl(A.��������,0))<>0" & _
        " Order By B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngҩƷID)
    strSQL = ""
    Do While Not rsTmp.EOF
        If InStr(1, gstr��������ID & ",", "," & Nvl(rsTmp!�ⷿid) & ",") > 0 Or gbyt�����ʾ��ʽ = 0 Then
            '��ʾ�����:�����ⷿ,�������ⷿ��ʾΪ�����ʱ
            strSQL = strSQL & "," & rsTmp!���� & ":" & rsTmp!���
        Else
            '�ǲ���Ա�ⷿ,����ʾΪ����
            strSQL = strSQL & "," & rsTmp!���� & ":" & IIf(Val(Nvl(rsTmp!���)) > 0, "��", "��") & "���."
        End If
        rsTmp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    GetStockInfo = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ChargeExistInsure(ByVal strNos As String, _
    Optional lng����ID As Long, Optional lng����ID As Long, _
    Optional bln���� As Boolean, Optional ByVal bln�˷� As Boolean) As Integer
'���ܣ��ж��շ�(���˷�)��¼���Ƿ����ָ����ҽ�����㷽ʽ
'������strNO=�շѵ��ݺ�
'���أ���������򷵻ص��ݵ�ʱ�����༰����ID,����ID,�Ƿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errH
            
    lng����ID = 0: lng����ID = 0: bln���� = False
    strWhere = " And A.��¼״̬ " & IIf(bln�˷�, "= 2", " IN(1,3)")
    strSQL = "" & _
        " Select /*+cardinality(j,10)*/ b.��¼id, b.����, b.����id, a.�Ƿ���" & vbNewLine & _
        " From ������ü�¼ A, ���ս����¼ B, Table(f_Str2list([1])) J" & vbNewLine & _
        " Where a.����id = b.��¼id And Mod(a.��¼����, 10) = 1 And a.No = j.Column_Value And b.���� = 1" & strWhere & vbNewLine & _
        "       And Rownum < 2"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", Replace(strNos, "'", ""))
    If Not rsTmp.EOF Then
        lng����ID = Nvl(rsTmp!����ID, 0)
        lng����ID = Nvl(rsTmp!��¼ID, 0)
        bln���� = Nvl(rsTmp!�Ƿ���, 0) = 1
        ChargeExistInsure = Nvl(rsTmp!����, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Public Sub Load�ѱ�(cbo As ComboBox, ByVal lng����ID As Long, ByVal bln���� As Boolean, ByRef rsTmp As ADODB.Recordset)
'���ܣ���ȡ�������Ψһ�Էѱ���д
'������lng����ID-��������ID,0��ʾȡ���������п��ҵķѱ�,bln����-�Ƿ�������޳���ķѱ�,rsTmp-�ѱ��¼��

    Dim strSQL As String, i As Integer
    
    cbo.Clear
    On Error GoTo errH
    If rsTmp Is Nothing Then
        strSQL = _
                    " Select a.����,a.����,a.����,Nvl(a.ȱʡ��־,0) as ȱʡ,Nvl(a.���޳���,0) as ����,Nvl(b.����ID,0) as ����ID" & _
                    " From �ѱ� a,�ѱ����ÿ��� b Where a.����=b.�ѱ�(+)" & _
                    " And Nvl(a.�������,3) IN(1,3) And a.����=1 And " & _
                    " Trunc(Sysdate) Between Nvl(a.��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD')) And Nvl(a.��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by a.����"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlOutExse")
    End If
    '���ÿ���:1-ȫ����2-ָ��
    strSQL = IIf(bln����, "", " And ����=0")
    rsTmp.Filter = IIf(lng����ID = 0, "����ID=0" & strSQL, _
                        "(����ID=0 " & strSQL & ") OR (����ID=" & lng����ID & strSQL & ")")

    For i = 1 To rsTmp.RecordCount
        cbo.AddItem rsTmp!���� & "-" & rsTmp!����
        If gstr�ѱ� = rsTmp!���� Then
            cbo.ListIndex = cbo.NewIndex
            cbo.ItemData(cbo.NewIndex) = 1
        End If
        If rsTmp!ȱʡ = 1 Then
            If cbo.ListIndex = -1 Then cbo.ListIndex = cbo.NewIndex
            cbo.ItemData(cbo.NewIndex) = 1
        End If
        '���޳��ﲻ���Ǳ��غ�ϵͳȱʡ
        If rsTmp!���� = 1 Then cbo.ItemData(cbo.NewIndex) = 2
        rsTmp.MoveNext
    Next
    'Load�ѱ� = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Load��̬�ѱ�(����ID As Long) As String
'���ܣ�Ȩ��ָ�����Ҷ�ȡ��ǰ��Ч�Ķ�̬�ѱ�
'���أ��ѱ�="���˽�,��һ��"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    'If ����ID = 0 Then Exit Function  'Ϊ0ʱ���������п���
    On Error GoTo errH
    strSQL = _
        " Select ����,����,���� From �ѱ�" & _
        " Where Nvl(����,1)=2 And Nvl(���ÿ���,1)=1 And Nvl(�������,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Union ALL" & _
        " Select Distinct A.����,A.����,A.����" & _
        " From �ѱ� A,�ѱ����ÿ��� B" & _
        " Where A.����=B.�ѱ� And B.����ID=[1]" & _
        " And Nvl(A.����,1)=2 And Nvl(A.���ÿ���,1)=2 And Nvl(A.�������,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(A.��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(A.��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", ����ID)
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    Load��̬�ѱ� = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillCanModi(strNo As String, bytFlag As Byte) As Boolean
'���ܣ��ж�һ�ŵ����Ƿ�����޸�
'������bytFlag=��¼����
'˵������������д��ڷ�����ʱ��ҩƷ,�������޸�(��Ϊ��������)

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.NO" & _
        " From ������ü�¼ A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
        " Where A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID" & _
        " And A.��¼״̬ IN(0,1,3) And A.�շ���� IN('5','6','7')" & _
        " And (Nvl(B.ҩ������,0)=1 Or Nvl(C.�Ƿ���,0)=1)" & _
        " And A.NO=[1] And A.��¼����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    BillCanModi = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadPatiCardObj(ByVal lng����ID As Long, strNo As String) As Detail
'���ܣ���ȡָ�����˵ľ��￨���ۼ�¼
'���أ�strNO=���۵��ݺ�
'      ���￨��Ŀ����(δ����۸�)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objDetail As Detail
    
    On Error GoTo errH
    
    strSQL = "Select A.NO,B.ID,B.���,D.���� as �������,B.����,B.����," & _
        " B.���,B.���㵥λ,B.��������,B.ִ�п���,Nvl(B.���ηѱ�,0) as ���ηѱ�," & _
        " Nvl(B.�Ӱ�Ӽ�,0) as �Ӱ�Ӽ�,Nvl(B.�Ƿ���,0) as �Ƿ���" & _
        " From ������ü�¼ A,�շ���ĿĿ¼ B,�շ��ض���Ŀ C,�շ���Ŀ��� D" & _
        " Where A.�շ�ϸĿID+0=B.ID And A.��¼����=1 And A.��¼״̬=0" & _
        " And A.�շ�ϸĿID+0=C.�շ�ϸĿID And C.�ض���Ŀ='���￨'" & _
        " And A.����Ա���� is NULL And A.����ID=[1] And B.���=D.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID)
    If Not rsTmp.EOF Then
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .���� = rsTmp!����
            .��� = IIf(IsNull(rsTmp!���), "", rsTmp!���)
            .���㵥λ = IIf(IsNull(rsTmp!���㵥λ), "", rsTmp!���㵥λ)
            
            .��� = rsTmp!�Ƿ��� = 1
            .�Ӱ�Ӽ� = rsTmp!�Ӱ�Ӽ� = 1
            .���ηѱ� = rsTmp!���ηѱ� = 1
            
            .��� = rsTmp!���
            .������� = rsTmp!�������
            .���� = rsTmp!����
            
            .ִ�п��� = IIf(IsNull(rsTmp!ִ�п���), 0, rsTmp!ִ�п���)
            .���� = IIf(IsNull(rsTmp!��������), "", rsTmp!��������)
        End With
        Set ReadPatiCardObj = objDetail
        strNo = rsTmp!NO
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillCanDelete(ByVal strNo As String, ByVal int��¼���� As Byte, _
    Optional ByRef blnHaveExe As Boolean, Optional ByVal strTime As String, Optional ByRef blnFlagPrint As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ��ж�һ�ŵ����Ƿ�����˷ѻ�����
    '������strNO=���ݺ�,,int��¼����=��¼����
    '˵���������˷ѻ����ʵ�����
    '    1.����δ��ȫִ��(ִ��״̬=0,2)
    '    2.ʣ��������<>0
    '    3.���������ſ�������
    '���أ�
    '   blnHaveExe=�Ƿ������(��ȫ/����)ִ�е�����
    '   -1=����ʧ��
    '    0=�����˷ѻ�����
    '    1=�õ��ݲ�����
    '    2=�Ѿ�ȫ����ȫִ��(ִ��״̬=1)
    '    3=δ��ȫִ�в���ʣ������Ϊ0
    '    blnFlagPrint=����Ӧ�������Ƿ��Ѵ�ӡ(����ҽ���еĲɼ���ʽ��ִ��)
    '����:���˺�
    '����:2014-07-15 09:56:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    
    On Error GoTo errH
    strNo = Replace(strNo, "'", "")
    strWhere = " And A.��¼����=[2] "
    If int��¼���� = 1 Then
        strWhere = " And mod(A.��¼����,10)=[2]"
    End If
    '1.����δ��ȫִ��(ִ��״̬=0,2)
    strSQL = "Select Distinct Nvl(A.ִ��״̬,0) as ִ��״̬,B.��������" & _
        " From ������ü�¼ A,����ҽ������ B" & vbNewLine & _
        " Where Nvl(A.���ӱ�־,0)<>9 And A.NO=[1]   And A.��¼״̬ IN(0,1,3) " & strWhere & vbNewLine & _
        " And A.ҽ�����=B.ҽ��ID(+) And A.NO=B.NO(+) And A.��¼����=B.��¼����(+) And Nvl(A.����״̬,0)<>1 " & _
        IIf(strTime <> "", " And A.�Ǽ�ʱ��=[3]", "")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int��¼����, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int��¼����)
    End If
    
    '���ݲ�����
    If rsTmp.EOF Then BillCanDelete = 1: Exit Function
    blnFlagPrint = Not IsNull(rsTmp!��������)
    
    '�����Ѿ�ȫ����ȫִ��
    rsTmp.Filter = "ִ��״̬<>1"
    If rsTmp.EOF Then BillCanDelete = 2 ': Exit Function
    
    '�Ƿ������(��ȫ/����)ִ�е�����
    rsTmp.Filter = "ִ��״̬<>0"
    blnHaveExe = Not rsTmp.EOF

    
    'δ��ȫִ�в���ʣ��������<>0
    '��ԭʼ��������δ��ȫִ�е��д�(������ҩ���˷Ѻ�ִ��״̬=1,���˷Ѽ�¼ִ��״̬<>1)
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    '��¼״̬=1,3ʱ��0:δִ��;1:��ȫִ��;2:����ִ�У���¼״̬=2ʱ��-x:��x���˷�
    strSQL = "" & _
    "   Select Nvl(�۸񸸺�,���) as ���" & _
    "   From ������ü�¼" & _
    "   Where Nvl(���ӱ�־,0)<>9 And NO=[1]  And Nvl(ִ��״̬,0)<>1 And ��¼״̬ IN(0,1,3)" & Replace(strWhere, "A.", "") & _
            IIf(strTime <> "", " And �Ǽ�ʱ��=[3]", "")
            
    strSQL = _
    "   Select ���,�շ�ϸĿID,Sum(����) as ʣ����  " & _
    "   From ( Select ��¼����,��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���,�շ�ϸĿID," & _
    "                 Avg(Nvl(����,1)*����) as ���� " & _
    "          From ������ü�¼" & _
    "          Where Nvl(���ӱ�־,0)<>9 And NO=[1] " & Replace(strWhere, "A.", "") & _
    "                And Nvl(ִ��״̬,0)<>1 And Nvl(�۸񸸺�,���) IN(" & strSQL & ")" & _
    "          Group by ��¼����,��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���),�շ�ϸĿID,����ID)" & _
    "   Group by ���,�շ�ϸĿID  " & _
    "   Having Sum(����)<>0"
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int��¼����, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int��¼����)
    End If
    If rsTmp.EOF Then BillCanDelete = 3
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    BillCanDelete = -1
End Function

Public Function BillDeleteAll(strNo As String, bytFlag As Byte, blnHaveExcutePrice As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ�����ݱ����Ƿ������ȫ�˷�(����ʣ������Ϊ��Ϊ�ж�����׼����=ʣ����)
    '���:strNO-���ݺ�
    '       bytFlag-��¼����(1-�շ�;2-����)
    '����:ȫ�˷���true,���򷵻�False
    '����:���˺�
    '����:2013-04-25 17:13:02
    '˵��:Ҫ��Ͻ������Ƿ�ȫ������˷��ж��Ƿ���������ȫ�˷�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
          
    On Error GoTo errH
    '���˺�45685,58077
    '��ȡҩƷ�շ���¼�е�׼����
    strSQL1 = _
    "   Select ����ID,Sum(Nvl(����,1)*ʵ������) as ׼������" & _
    "   From ҩƷ�շ���¼" & _
    "   Where NO=[1] And MOD(��¼״̬,3)=1 And ����� is NULL" & _
    "             And ���� IN([3],[4])  " & _
    "   Group by ����ID" & _
    "   Union ALL "
    If blnHaveExcutePrice Then
            '60735:��ҽ��ִ�мƼ��д�������ʱ,��ҽ��ִ�мƼ���ȡ��
            '77686,���ϴ�,2014/9/18,�����������
            strSQL1 = strSQL1 & _
            " Select Max(ID) As ����id, Decode(Sign(Sum(����)), -1, 0, Sum(����)) As ׼����" & vbNewLine & _
            " From ( Select Decode(a.��¼״̬, 2, 0, a.Id) As ID, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(a.����, 1) As ����," & vbNewLine & _
            "              Decode(a.��¼״̬, 2, 0, Nvl(a.����, 1) * Nvl(a.����, 1)) As ԭʼ����" & vbNewLine & _
            "       From ������ü�¼ A, ����ҽ����¼ M" & vbNewLine & _
            "       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Instr('5,6,7', a.�շ����) = 0 And" & vbNewLine & _
            "             a.No = [1] And mod(a.��¼����,10) = [3] And a.��¼״̬ In (1, 2, 3)��and a.�۸񸸺� is null" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, -1 * b.���� As ��ִ��, 0 ԭʼ����" & vbNewLine & _
            "       From ������ü�¼ A, ҽ��ִ�мƼ� B, ����ҽ����¼ M" & vbNewLine & _
            "       Where a.ҽ����� = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid + 0 And a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0" & vbNewLine & _
            "           And Instr('5,6,7', a.�շ����) = 0" & vbNewLine & _
            "           And (Exists (Select 1  From ����ҽ��ִ��  Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And b.Ҫ��ʱ�� = Ҫ��ʱ�� And Nvl(ִ�н��, 0) = 1)" & vbNewLine & _
            "                Or Exists (Select 1 From ����ҽ������ Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And Nvl(ִ��״̬, 0) = 1))" & vbNewLine & _
            "          And a.No = [1] And mod(a.��¼����,10) = [3] And a.��¼״̬ In (1, 3)��and a.�۸񸸺� is null" & vbNewLine & _
            "       ) Q1" & vbNewLine & _
            " Where Not Exists (Select 1 From ҩƷ�շ���¼ Where ����id = Q1.Id And instr( ',8,9,10,21,24,25,26,',','||����||',')>0) " & vbNewLine & _
            " Group by ҽ��ID,�շ�ϸĿID  Having Max(ID)<>0"
    Else
         strSQL1 = strSQL1 & _
         " Select Max(ID) as ����ID,decode(sign(Sum(����)),-1,0,Sum(����)) as ׼���� " & _
         " From (   Select J.ID,J.ҽ����� as ҽ��ID,J.�շ�ϸĿID,nvl(J.����,1)*nvl(J.����,1)  as ���� " & _
         "               From  ������ü�¼ J,����ҽ����¼ M " & _
         "               Where  J.ҽ�����=M.ID  " & _
         "                       And Exists(Select 1 From   ����ҽ������ where ҽ��ID=J.ҽ����� and  Nvl( ִ��״̬, 0) <> 1 And No||''=[1] and ��¼����+0=[2]) " & _
         "                       And Exists(Select 1 From   ����ҽ���Ƽ� A Where   A.ҽ��ID=J.ҽ����� and A.�շ�ϸĿID=J.�շ�ϸĿID And A.��������=0  and  Nvl( A.�շѷ�ʽ, 0) =0 ) " & _
         "                       And J.No=[1] and mod(J.��¼����,10)=[2] And J.��¼״̬ in (1,2,3) and J.�۸񸸺� is null   " & _
         "                       And Instr('5,6,7', j.�շ����) = 0 And  Not Exists  (Select 1  From ��������  Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
         "                       And  instr(',C,D,F,G,K,',','||M.�������||',')=0  " & _
         "               Union all  " & _
         "               Select j.id, A.ҽ��ID,a.�շ�ϸĿID,-1*nvl(a.����,1)*nvl(C.��������,1) as ���� " & _
         "                              From ����ҽ���Ƽ� A,����ҽ������ B,����ҽ��ִ�� C,������ü�¼ J,����ҽ����¼ M " & _
         "               where  A.ҽ��ID=b.ҽ��id  and b.ҽ��id=c.ҽ��id and b.���ͺ�=c.���ͺ� And a.ҽ��id=M.ID " & _
         "                       And Nvl(C.ִ�н��, 1) =1  And A.��������=0 and  Nvl( A.�շѷ�ʽ, 0) =0  And Nvl(b.ִ��״̬, 0) <> 1 And B.No||''=[1] and B.��¼����+0=[2]  " & _
         "                       And a.ҽ��id=J.ҽ����� and a.�շ�ϸĿid=j.�շ�ϸĿid  " & _
         "                       And J.No=[1] and mod(J.��¼����,10)=[2] And J.��¼״̬ in (1,3) and J.�۸񸸺� is null   " & _
         "                       And Instr('5,6,7', j.�շ����) = 0 And  Not Exists  (Select 1  From ��������  Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
         "                       And  instr(',C,D,F,G,K,',','||M.�������||',')=0   "
        
        '58077��Ҫ�ſ�ҽ���ƻ��в���������ȡ�ķ���:
             '   0-������ȡ��1-�����Թܷ��ã�2-һ�η���ֻ��ȡһ�Σ�3-����ֻ��ȡһ�Σ�4-����δִ����ȡһ�Σ�5-����ֻ��ȡһ�Σ��ų�������Ŀ��6-����δִ����ȡһ�Σ��ų�������Ŀ��7-ÿ���״β���ȡ
        strSQL1 = strSQL1 & " Union All" & _
         "               Select j.Id, a.ҽ��id, a.�շ�ϸĿid, 0 As ���� " & _
         "               From ����ҽ���Ƽ� A,������ü�¼ J , ����ҽ����¼ M " & _
         "               Where  a.ҽ��id = M.ID and a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid  And A.��������=0  and Nvl(a.�շѷ�ʽ, 0) <> 0  " & _
         "                            And j.No =[1] And mod(j.��¼����,10) = [2]  And nvl(J.ִ��״̬,0)=2 " & _
         "                            And j.��¼״̬ In (1, 3) And  j.�۸񸸺� Is Null And Instr('5,6,7', j.�շ����) = 0  " & _
         "                            And Not Exists(Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
         "                            And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0  " & _
         "               ) " & _
         " group by ҽ��ID,�շ�ϸĿID Having Max(ID)<>0"
      End If
    
    '���ŷ��õ�����ʣ������Ϊ0����(��ϸ��ÿһ��)
    'ִ��״̬ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
    strSQL = _
    " Select Sum(A.ID) as ID,Sum(A.ִ��״̬) as ִ��״̬," & _
    "               A.���,A.�շ����,Sum(����) as ʣ������" & _
    " From (  Select    Decode(A.��¼״̬,2,0,A.ID) as ID," & _
    "                           Decode(A.��¼״̬,2,0,Nvl(A.ִ��״̬,0)) as ִ��״̬," & _
    "                           A.���,A.�շ����,Nvl(A.����,1)*A.���� as ����" & _
    "               From ������ü�¼ A" & _
    "               Where A.�۸񸸺� is NULL And Nvl(A.���ӱ�־,0)<>9 And mod(A.��¼����,10)=[2] And A.NO=[1]" & _
    "               ) A" & _
    "   Group by A.���,A.�շ����" & _
    "   Having Nvl(Sum(����),0)<>0"
                
    '��ʣ��������׼�������������������
        '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
        '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        'Decode(Instr(',4,5,6,7,',A.�շ����),0,A.ʣ������,
    strSQL = _
    "   Select A.��� " & _
    "   From ( Select A.���,A.ʣ������," & _
    "                           Decode(A.ִ��״̬,1,0,Nvl(B.׼������,A.ʣ������) ) as ׼������" & _
    "               From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
    "               Where A.ID=B.����ID(+)" & _
    "               ) A" & _
    " Where Nvl(A.׼������,0)<>A.ʣ������"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag, IIf(bytFlag = 2, "9", "8"), IIf(bytFlag = 2, "25", "24"))
    BillDeleteAll = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillDeleteAllNew(strNo As String, bytFlag As Byte) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ�����ݱ����Ƿ������ȫ�˷�(����ʣ������Ϊ��Ϊ�ж�����׼����=ʣ����)
    '���:strNO-���ݺ�
    '     bytFlag-��¼����(1-�շ�;2-����)
    '����:ȫ�˷���true,���򷵻�False
    '����:Ƚ����
    '����:2016-10-09 17:13:02
    '˵��:Ҫ��Ͻ������Ƿ�ȫ������˷��ж��Ƿ���������ȫ�˷�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
          
    On Error GoTo errH
    '���˺�45685,58077,99715
    '��ȡҩƷ�շ���¼�е�׼����
    strSQL1 = _
        " Select ����ID,Sum(Nvl(����,1)*ʵ������) as ׼������" & vbNewLine & _
        " From ҩƷ�շ���¼" & vbNewLine & _
        " Where NO=[1] And Mod(��¼״̬,3)=1 And ����� Is NULL And ���� IN([3],[4])" & vbNewLine & _
        " Group by ����ID"
        
    '��������ص�׼����
    strSQL1 = strSQL1 & vbNewLine & _
        " Union ALL " & vbNewLine & _
        " Select Max(ID) As ����ID, Nvl(Sum(����), 0) As ׼����" & vbNewLine & _
        " From(Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Decode(c.ִ��״̬, 0, 1, 0) * c.���� As ����" & vbNewLine & _
        "      From ������ü�¼ A, ����ҽ������ B, ҽ��ִ�мƼ� C, ����ҽ����¼ M" & vbNewLine & _
        "      Where a.ҽ����� = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.ҽ��ID = m.ID" & vbNewLine & _
        "            And b.���ͺ� = c.���ͺ� And a.�շ�ϸĿid = c.�շ�ϸĿid + 0 And a.�۸񸸺� Is Null" & vbNewLine & _
        "            And Instr(',5,6,7,', ',' || a.�շ���� || ',') = 0" & vbNewLine & _
        "            And Not Exists(Select 1 From �������� C Where a.�շ�ϸĿid = c.����id And c.�������� = 1)" & vbNewLine & _
        "            And Instr(',C,D,F,G,K,',','||m.�������||',')=0" & vbNewLine & _
        "            And a.No = [1] And a.��¼���� = [2] And a.��¼״̬ In (1, 3) And b.��¼���� = [2]" & vbNewLine & _
        "     )" & vbNewLine & _
        " Group By ҽ��ID, �շ�ϸĿID" & vbNewLine & _
        " Having Max(ID) <> 0"
    
    '���ŷ��õ�����ʣ������Ϊ0����(��ϸ��ÿһ��)
    'ִ��״̬ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
    '*��ҽ��ִ�мƼ۵Ĳ����˷��޷��ж�׼���������������˷�(ִ��״̬����Ϊ1)
    strSQL = _
        " Select Max(A.ID) as ID,Max(A.ִ��״̬) as ִ��״̬," & _
        "        A.���,A.�շ����,Sum(����) as ʣ������" & _
        " From (Select Decode(A.��¼״̬,2,0,A.ID) as ID," & _
        "              Decode(A.��¼״̬,2,0,Nvl(A.ִ��״̬,0)) as ִ��״̬," & _
        "              A.���,A.�շ����,Nvl(A.����,1)*A.���� as ����" & _
        "       From ������ü�¼ A" & _
        "       Where mod(A.��¼����,10)=[2] And A.NO=[1] And Nvl(A.���ӱ�־,0)<>9" & _
        "       Union All" & _
        "       Select 0 As ID,1 As ִ��״̬,a.���,a.�շ����,0 As ����" & vbNewLine & _
        "       From ������ü�¼ A" & vbNewLine & _
        "       Where A.��¼���� = [2] And A.NO=[1] And A.��¼״̬ In (1, 3) And Nvl(A.ִ��״̬, 0) = 2" & vbNewLine & _
        "             And Not Exists(Select 1" & vbNewLine & _
        "                            From ����ҽ������ B, ҽ��ִ�мƼ� C" & vbNewLine & _
        "                            Where b.ҽ��id = A.ҽ����� And b.No = A.No" & vbNewLine & _
        "                                  And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ�" & vbNewLine & _
        "                                  And c.�շ�ϸĿid + 0 = A.�շ�ϸĿid And b.��¼���� = [2])" & vbNewLine & _
        "             And Instr('5,6,7', A.�շ����) = 0" & vbNewLine & _
        "             And Not Exists(Select 1 From �������� Where ����id = A.�շ�ϸĿid And Nvl(��������, 0) = 1)" & _
        "      ) A" & _
        " Group by A.���,A.�շ����" & _
        " Having Nvl(Sum(����),0)<>0"
                
    '��ʣ��������׼�������������������
        '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
        '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
    strSQL = _
        " Select A.��� " & _
        " From(Select A.���,A.ʣ������," & _
        "             Decode(A.ִ��״̬,1,0,Nvl(B.׼������,A.ʣ������) ) as ׼������" & _
        "      From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
        "      Where A.ID=B.����ID(+)" & _
        "     ) A" & _
        " Where Nvl(A.׼������,0)<>A.ʣ������"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag, _
        IIf(bytFlag = 2, "9", "8"), IIf(bytFlag = 2, "25", "24"))
    BillDeleteAllNew = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BillExistDelete(ByVal strNo As String, ByVal int��¼���� As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ�������Ƿ����(����)�˷ѻ����ʵ�����
    '���:strNO-ָ�����ݺ�
    '     int��¼����-��¼����
    '����:�����˷ѵ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-18 10:38:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select NO From ������ü�¼ Where NO=[1] And ��¼����=[2] And ��¼״̬=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int��¼����)
    BillExistDelete = Not rsTemp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function BillExistMoney(ByVal strNos As String, ByVal int��¼���� As Integer, _
    Optional bln���շ� As Boolean, Optional lng��ӡID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ�����ݵ���Ŀ�Ƿ��Ѿ�ȫ������(ʣ������=0)
    '���:strNOs=����Ϊһ�ŵ���,Ҳ����Ϊ���ŵ���(�൥���շѲ�����),��ʽΪ:"'AAA','BBB','CCC',..."
    '     int��¼����-��¼����
    '     lng��ӡID-
    '����:
    '����: True=û��ȫ�����꣬��ʾ�����˷�
    '      False=��ȫ������
    '����:���˺�
    '����:2014-06-18 10:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String, strTmp As String
    Dim strWhere As String
    
    On Error GoTo errH
        
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    'ʵ�ж൥��һ�ν����,ҽ�������˷Ѻ�������¼����Ϊ11�ļ�¼,"NO,��¼״̬,ִ��״̬,���"���ظ�,AVG����������������,����Ҫ����"��¼����"
    
    If lng��ӡID > 0 Then
        '����ʱ����ȡ��
        
        strTmp = Replace(strNos, "'", "")
        If int��¼���� = 1 Then
            strWhere = "And Mod(A.��¼����,10)=[1]"
        Else
            strWhere = "And A.��¼���� =[1]"
        End If
        If bln���շ� Then strWhere = strWhere & " And A.��¼״̬<>0 "
        
        strSQL = _
        " Select NO,���,Sum(����) as ʣ������" & _
        " From ( Select A.NO,A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
        "               Avg(Nvl(A.����, 1) * A.����) As ����" & _
        "       From ������ü�¼ A,(select NO From ��ʱƱ�ݴ�ӡ���� where ID=[3] and ����=[1]) J" & _
        "       Where Nvl(A.���ӱ�־,0)<>9 And A.NO=J.NO " & strWhere & _
        "       Group by A.NO,A.��¼����,A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���))" & _
        " Group by NO,��� " & _
        " Having Sum(����)<>0 "
        
          
          Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", int��¼����, strTmp, lng��ӡID)
    ElseIf UBound(Split(strNos, ",")) > 0 Then
        strTmp = Replace(strNos, "'", "")
        If int��¼���� = 1 Then
            strWhere = "And Mod(A.��¼����,10)=[1]"
        Else
            strWhere = "And A.��¼���� =[1]"
        End If
        If bln���շ� Then strWhere = strWhere & " And A.��¼״̬<>0 "
        
        strSQL = _
        " Select /*+ rule */  NO,���,Sum(����) as ʣ������" & _
        " From ( Select A.NO,A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
        "               Avg(Nvl(A.����, 1) * A.����) As ����" & _
        "       From ������ü�¼ A,(Select Column_Value From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) J " & _
        "       Where Nvl(A.���ӱ�־,0)<>9 And A.NO=J.Column_Value " & strWhere & _
        "       Group by A.NO,A.��¼����,A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���))" & _
        " Group by NO,��� " & _
        " Having Sum(����)<>0 "
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", int��¼����, strTmp)
    Else
        If int��¼���� = 1 Then
            strWhere = "And Mod(��¼����,10)=[2]"
        Else
            strWhere = "And ��¼���� =[2]"
        End If
        If bln���շ� Then strWhere = strWhere & " And ��¼״̬<>0 "
            
        strSQL = _
        " Select NO,���,Sum(����) as ʣ������" & _
        " From ( Select NO,��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���," & _
        "               Avg(Nvl(����, 1) * ����) As ����" & _
        "        From ������ü�¼" & _
        "        Where Nvl(���ӱ�־,0)<>9 And NO=[1] " & strWhere & _
        "        Group by NO,��¼����,��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by NO,��� " & _
        " Having Sum(����)<>0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", Replace(strNos, "'", ""), int��¼����)
    End If
    BillExistMoney = Not rsTemp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetBillRows(strNo As String, bytFlag As Byte) As Integer
'���ܣ���ȡһ�ŷ��õ�����δ���ϵķ�������
'������bytFlag=��¼����
'˵���������˷�/����ʱ�жϲ����˷�/����,�˷�ʱҪ�ſ�������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    strSQL = _
    " Select ���,Sum(����) as ʣ������" & _
    " From ( Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���," & _
    "               Avg(Nvl(����, 1) * ����) As ����" & _
    "        From ������ü�¼" & _
    "        Where Nvl(���ӱ�־,0)<>9 And NO=[1] And ��¼����=[2]" & _
    "        Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
    " Group by ��� " & _
    " Having Sum(����)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    If Not rsTmp.EOF Then GetBillRows = rsTmp.RecordCount
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckLimit(ByVal objBill As ExpenseBill, Optional intPage As Integer, Optional intRow As Integer) As Boolean
'���ܣ����õ���ҩƷ�����������,�����ڼ��ʵ�,�շ�
'������intPage,intRow=ָ����ĳ�ŵ���ĳ�н��м��,����Ϊȫ�����
'˵����
'   1.ȫ��û���������������棻���г���ҩƷ�����ں�������ʾ�������ؼ١�
'   2.���ʱ���Ϊÿ�����˵������
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim tmpDetail As BillDetail, curDetail As BillDetail
    Dim dblTime As Double, strItemIDs As String '�Ѿ������˵�ҩƷ
    Dim dbl���� As Double, i As Integer, p As Integer
    Dim strҩƷ������ʾ As String
    
    CheckLimit = True
    Err = 0: On Error GoTo errH:
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, objBill.Pages.Count, intPage)
        '�ռ�����
        strItemIDs = ""
        For i = 1 To objBill.Pages(p).Details.Count
            If intRow = 0 Or (intRow > 0 And i = intRow) Then
                With objBill.Pages(p).Details(i)
                    '�ռ�ҩƷID
                    If InStr(strItemIDs & ",", "," & .�շ�ϸĿID & ",") = 0 And InStr(",5,6,7,", .�շ����) > 0 Then
                        strItemIDs = strItemIDs & "," & .�շ�ϸĿID
                    End If
                End With
            End If
        Next
        If strItemIDs <> "" Then
            strItemIDs = Mid(strItemIDs, 2)
            strSQL = "Select A.ҩƷID,A.����ϵ��,B.���㵥λ as ������λ" & _
                " From ҩƷ��� A,������ĿĿ¼ B" & _
                " Where A.ҩ��ID=B.ID And A.ҩƷID IN (" & strItemIDs & ")"
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlOutExse")
            strItemIDs = ""
            For i = 1 To objBill.Pages(p).Details.Count
                If intRow = 0 Or (intRow > 0 And i = intRow) Then
                    Set tmpDetail = objBill.Pages(p).Details(i)
                    If InStr(",5,6,7,", tmpDetail.�շ����) > 0 And tmpDetail.Detail.�������� > 0 Then
                        If InStr(strItemIDs, "," & tmpDetail.�շ�ϸĿID) = 0 Then
                            dblTime = 0
                            For Each curDetail In objBill.Pages(p).Details
                                If InStr(",5,6,7,", curDetail.�շ����) > 0 And tmpDetail.�շ�ϸĿID = curDetail.�շ�ϸĿID Then
                                    dblTime = dblTime + curDetail.���� * curDetail.����
                                End If
                            Next
                            rsTmp.Filter = "ҩƷID=" & tmpDetail.�շ�ϸĿID
                            If Not rsTmp.EOF Then
                                If gblnҩ����λ Then
                                    dbl���� = dblTime * tmpDetail.Detail.ҩ����װ * rsTmp!����ϵ��
                                Else
                                    dbl���� = dblTime * rsTmp!����ϵ��
                                End If
                                If dbl���� > tmpDetail.Detail.�������� Then
                                    strҩƷ������ʾ = IIf(objBill.Pages.Count = 1, "", "����" & p & "��") & "ҩƷ """ & tmpDetail.Detail.���� & """ ���ܼ��� " & _
                                        FormatEx(dbl����, 5) & rsTmp!������λ & "(" & FormatEx(dblTime, 5) & IIf(gblnҩ����λ, tmpDetail.Detail.ҩ����λ, tmpDetail.Detail.���㵥λ) & ") ������������ " & _
                                        FormatEx(tmpDetail.Detail.��������, 5) & rsTmp!������λ & " ��" & vbCrLf & vbCrLf & "�Ƿ�����?�������˳�����֮ǰ���ٽ��д������������ѡ��[ȡ��]"
                                
                                        strҩƷ������ʾ = MsgBox(strҩƷ������ʾ, vbYesNoCancel + vbDefaultButton3 + vbInformation, gstrSysName)
                                        If strҩƷ������ʾ = vbNo Then
                                            CheckLimit = False: Exit Function
                                        End If
                                        If strҩƷ������ʾ = vbCancel Then
                                           gbln�������� = True
                                        End If
                                End If
                            End If
                            strItemIDs = strItemIDs & "," & tmpDetail.�շ�ϸĿID
                        End If
                    End If
                End If
            Next
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��ҩ����IDs(strNo As String, Optional strType As String) As String
'���ܣ���ȡ�շѵ��ݵ�ҩ��ID�����ĵķ��ϲ���ID
'���أ�"23,45,656..."
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    If strType = "" Then   'ҩƷ
        strSQL = "Select Distinct ִ�в���ID From ������ü�¼" & _
        " Where ��¼����=1 And ��¼״̬ IN(0,1) And �շ���� IN('5','6','7') And NO=[1]"
    Else                   '����
        strSQL = "Select Distinct a.ִ�в���ID From ������ü�¼ a,�������� b" & _
        " Where a.�շ�ϸĿid=b.����id And b.��������=1 And " & _
        "(a.��¼����=1 And a.��¼״̬ In (0,1) Or a.��¼����=2 And a.��¼״̬=1) And a.�շ����='4' And a.NO=[1]"
        '��¼����=1 �շ�ʱ,��ֱ���շѻ��Ữ�۵��շ�,����a.��¼״̬ In (0,1)
        '��¼����=2 ������� ������ʱ����,��¼״̬=1,���ʻ��۵������ͨ��zl_������ʼ�¼_Verify�����Զ�����
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!ִ�в���ID) Then
            Get��ҩ����IDs = Get��ҩ����IDs & "," & rsTmp!ִ�в���ID
        End If
        rsTmp.MoveNext
    Loop
    Get��ҩ����IDs = Mid(Get��ҩ����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillRepeat(lng����ID As Long, bytƱ�� As Byte, strFactNO As String) As Boolean
'���ܣ���ʹ����Ʊ��֮ǰ������Ƿ��ظ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From Ʊ��ʹ����ϸ" & _
        " Where ����ID=[1] And Ʊ��=[2] And ����=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID, bytƱ��, strFactNO)
    CheckBillRepeat = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlCheckIsInvoiceListPrinted(ByVal strNo As String, Optional blnNOMoved As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ���ϸ��ӡ��
    '���:strNo-���ݺ�
    '       blnNOMoved-�Ƿ�����ʷ������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-04-16 09:52:05
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long, strPrintTable As String, strSQL1 As String
    
    On Error GoTo errHandle
    If gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Function
    
    strPrintTable = "Ʊ�ݴ�ӡ��ϸ"
    'Ӧ�������һ�δ�ӡ���������
    strSQL = "" & _
    "   Select NO  " & _
    "   From " & strPrintTable & " A, Table( f_Str2list([1])) J  " & _
    "   Where A.Ʊ��=1  And A.NO = J.Column_Value And Rownum=1"
    If blnNOMoved Then
        strSQL = Replace(strSQL, strPrintTable, "H" & strPrintTable)
        'strSql = strSql & " Union ALL " & strSQL1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If rsTmp.EOF Then Exit Function
    zlCheckIsInvoiceListPrinted = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetMultiNOs(ByVal strNo As String, _
    Optional lng��ӡID As Long, _
    Optional blnNOMoved As Boolean, _
    Optional bln��������ŷ��� As Boolean, _
    Optional bln��ʷ��ͬ���� As Boolean = False) As String
    '���ܣ�����һ���շѵ��ݵ�NO������ͬһ�δ�ӡ�Ķ���NO
    '������blnNoMoved�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
    '          bln��������ŷ���-�Ƿ񰴽�����ŷ���
    '          bln��ʷ��ͬ����-�Ƿ�������ʷ��һ���ѯ
    '���أ���ʽ��"'AAA','BBB','CCC',..."
    '      ���ָ����"lng��ӡID",�򷵻�
    '˵�������ڶ൥���շ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
    Dim i As Long, strNos As String
    
    On Error GoTo errHandle
    lng��ӡID = 0
    If bln��������ŷ��� Then
        strSQL = "Select Distinct A.NO,0 as ID" & vbNewLine & _
                " From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
                " Where A.����ID = B.����ID" & vbNewLine & _
                "       And B.������� In (Select Max(A.�������)" & vbNewLine & _
                "                          From  ����Ԥ����¼ A, ������ü�¼ B" & vbNewLine & _
                "                          Where A.����ID = B.����ID And Mod(b.��¼����, 10) =1 And B.��¼״̬<>2 And B.NO = [1] )"
        If blnNOMoved Then
            strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
            strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        ElseIf bln��ʷ��ͬ���� Then
            strSQL1 = Replace(strSQL, "������ü�¼", "H������ü�¼")
            strSQL1 = Replace(strSQL1, "����Ԥ����¼", "H����Ԥ����¼")
            strSQL = strSQL & " Union ALL " & vbNewLine & strSQL1
        End If
    Else
        'Ӧ�������һ�δ�ӡ���������
        strSQL = "Select ID, NO" & vbNewLine & _
                " From Ʊ�ݴ�ӡ����" & vbNewLine & _
                " Where �������� = 1" & vbNewLine & _
                "       And ID In (Select ID" & vbNewLine & _
                "                 From (Select b.Id" & vbNewLine & _
                "                       From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B" & vbNewLine & _
                "                       Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = [1]" & vbNewLine & _
                "                       Order By a.ʹ��ʱ�� Desc)" & vbNewLine & _
                "                 Where Rownum < 2)"
        If blnNOMoved Then
            strSQL = Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
            strSQL = Replace(strSQL, "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
        ElseIf bln��ʷ��ͬ���� Then
            strSQL1 = Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
            strSQL1 = Replace(strSQL1, "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
            strSQL = strSQL & " Union ALL " & vbNewLine & strSQL1
        End If
    End If
    strSQL = strSQL & vbNewLine & _
            " Order by NO"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    
    If Not rsTmp.EOF Then
        lng��ӡID = Nvl(rsTmp!ID, 0) '����û��
        For i = 1 To rsTmp.RecordCount
            strNos = strNos & ",'" & rsTmp!NO & "'"
            rsTmp.MoveNext
        Next
        GetMultiNOs = Mid(strNos, 2)
    Else
        GetMultiNOs = "'" & strNo & "'"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub GetErrorItem(ByRef lng���ϸĿid As Long, ByRef str�վݷ�Ŀ As String, _
    Optional ByVal strPricrGrade As String)
'���ܣ���ȡ�������շ�ϸĿid,�վݷ�Ŀ
'���ã��շ�ʱ����Ƿ����������,�Ե��뻮�۵�,����Ʊ������
'˵��������Ŀ��Ӧ������ҲӦΪ�����Ŀ(δ��)��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    Set rsTmp = zlGetSpecialItemFee("�����", strPricrGrade)
    If Not rsTmp.EOF Then
        lng���ϸĿid = Val(Nvl(rsTmp!�շ�ϸĿID))
        str�վݷ�Ŀ = Nvl(rsTmp!�վݷ�Ŀ)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function SeekPatiBill(ByVal lng����ID As Long) As Long
'���ܣ����ݲ�����Ѱ���˵Ļ��۵���(ϵͳָ��������)
'���أ���������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strDeptLimitWhere As String
    
    On Error GoTo errH
    '96357
    If gTy_Module_Para.str�����շ�ִ�п��� <> "" Then
        strDeptLimitWhere = vbNewLine & _
            " And Exists(Select 1" & vbNewLine & _
            "      From ������ü�¼ M" & vbNewLine & _
            "      Where m.��¼���� = a.��¼���� And m.No = a.No" & vbNewLine & _
            "           And Instr('," & gTy_Module_Para.str�����շ�ִ�п��� & ",', ','||m.ִ�в���id||',') > 0)"
    ElseIf gTy_Module_Para.str�������շ�ִ�п��� <> "" Then
        strDeptLimitWhere = vbNewLine & _
            " And Not Exists(Select 1" & vbNewLine & _
            "      From ������ü�¼ M" & vbNewLine & _
            "      Where m.��¼���� = a.��¼���� And m.No = a.No" & vbNewLine & _
            "           And Instr('," & gTy_Module_Para.str�������շ�ִ�п��� & ",', ','||m.ִ�в���id||',') > 0)"
    End If
    
    strSQL = "Select Count(a.NO) as ������ From ������ü�¼ A" & vbNewLine & _
            " Where a.��¼����=1 And a.��¼״̬=0" & vbNewLine & _
            "       And a.������ is Not NULL And a.����Ա���� IS NULL" & vbNewLine & _
            "       And a.����ID=[1] And a.�Ǽ�ʱ��+0>=Sysdate-" & gintSeekDays & _
           strDeptLimitWhere
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID)
    If Not rsTmp.EOF Then SeekPatiBill = Nvl(rsTmp!������, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Checkҩ���ϰల��() As Boolean
'���ܣ����ҽԺ��ҩ���Ƿ�ʹ�����ϰల��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Count(B.����ID) as NUM From ��������˵�� A,���Ű��� B" & _
        " Where A.����ID=B.����ID And A.�������� IN('��ҩ��','��ҩ��','��ҩ��')"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlOutExse")
    If Not rsTmp.EOF Then
        Checkҩ���ϰల�� = Nvl(rsTmp!Num, 0) <> 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�շ�ִ�п���ID( _
    ByVal str��� As String, ByVal lng��Ŀid As Long, ByVal intִ�п������� As Integer, _
    ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, Optional ByVal int��Χ As Integer = 1, _
    Optional ByVal lng��ҩ�� As Long, Optional ByVal lng��ҩ�� As Long, Optional ByVal lng��ҩ�� As Long, _
    Optional ByVal lngִ�п���ID As Long, Optional ByVal lng���˲���ID As Long) As Long
'���ܣ���ȡ�շ���Ŀ��ִ�п���
'������int��Χ=1.����,2-סԺ
'      lngִ�п���ID=ָ����ȱʡִ�п���ID(����ҩƷ������)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strҩ�� As String, lngҩ�� As Long
    Dim bytDay As Byte
    
    On Error GoTo errH
    
    If str��� = "4" Then
        '��ִ�п�������ʱ
        strSQL = _
            " Select Distinct" & _
            "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            "       And B.������� IN([1],3) And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[2]  " & _
            "               Or Exists(select 1 From �������Ҷ�Ӧ M where A.��������ID=M.����ID And M.����ID=[2] ))" & _
            "       And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", int��Χ, lng���˿���ID, lng��Ŀid)
        If Not rsTmp.EOF Then
            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID    '3:�����û�У��򷵻ص�һ�����õ�ִ�п���(��ҽ��վ��ͬ)
            
            '1:ȱʡΪָ����(ҽ����)ִ�п���,�����Ƿ�����ڲ��˿���
            rsTmp.Filter = "ִ�п���ID=" & lngִ�п���ID
            '2:�����ɷ����ڲ��˿��ҵ�ִ�п���
            If rsTmp.EOF Then
                '2.1:����ȱʡΪ���˿���
                If lngִ�п���ID <> lng���˿���ID Then
                    rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˿���ID
                End If
                '2.2:����ȱʡΪ���˲���
                If rsTmp.EOF Then
                    If lng���˲���ID <> 0 And lng���˲���ID <> lng���˿���ID And lng���˲���ID <> lngִ�п���ID Then
                        rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˲���ID
                    End If
                End If
            End If
            '2.3:�ɷ����ڲ��˿��ҵ�һ��ִ�п���
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        If str��� = "5" Then
            strҩ�� = "��ҩ��": lngҩ�� = lng��ҩ��
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��": lngҩ�� = lng��ҩ��
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��": lngҩ�� = lng��ҩ��
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not gblnҩ���ϰల�� Then
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                "       And B.������� IN([2],3) And B.����ID=C.ID" & _
                "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                "       And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                "       And B.������� IN([2],3) And B.����ID=C.ID" & _
                "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                "       And D.����ID=C.ID And D.����=[5]" & _
                "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                "       And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strҩ��, int��Χ, lng���˿���ID, lng��Ŀid, bytDay)
        If Not rsTmp.EOF Then
            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
            rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lngִ�п���ID
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0 And ִ�п���ID=" & lngִ�п���ID
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lngҩ��
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0 And ִ�п���ID=" & lngҩ��
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    Else
        Select Case intִ�п�������
            Case 0 '0-����ȷ����
                Get�շ�ִ�п���ID = UserInfo.����ID
            Case 1 '1-�������ڿ���
                Get�շ�ִ�п���ID = lng���˿���ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get�շ�ִ�п���ID = lng���˿���ID
                Else
                    Get�շ�ִ�п���ID = lng���˲���ID
                End If
            Case 3 '3-����Ա���ڿ���
                Get�շ�ִ�п���ID = UserInfo.����ID
            Case 4 '4-ָ������
                strSQL = "" & _
                "   Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                "   From �շ�ִ�п���  A,���ű� C" & _
                "   Where A.�շ�ϸĿID=[1]  And A.ִ�п���ID+0=C.ID  " & _
                "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                "   Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng��Ŀid, int��Χ, lng���˿���ID)
                If Not rsTmp.EOF Then
                     'ȱʡȡ����Ա���ڿ���
                    rsTmp.Filter = "��������ID=" & UserInfo.����ID
                    If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
                    If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
                    If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                End If
            Case 6 '6-���������ڿ���
                Get�շ�ִ�п���ID = lng��������ID
        End Select
        If Get�շ�ִ�п���ID = 0 Then Get�շ�ִ�п���ID = UserInfo.����ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ժҪ(ByVal strNo As String, ByVal bytFlag As Byte, ByVal int��� As Integer) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ�������ݵ�ժҪ
    '��Σ�strNO-���ݺ�
    '      bytFlag-��¼����
    '      int���-�к�
    '���Σ�
    '���أ�ժҪ
    '���ƣ����˺�
    '���ڣ�2010-03-03 15:19:36
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ժҪ From ������ü�¼" & _
        " Where ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=[2] And Nvl(�۸񸸺�,���)=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag, int���)
    If Not rsTmp.EOF Then Get����ժҪ = Nvl(rsTmp!ժҪ)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillisAdviceMoney(ByVal strNo As String, ByVal bytFlag As Byte, _
    lngҽ��ID As Long, lng���ͺ� As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��ж�һ�ŵ����Ƿ�ҽ���ĸ��ӷ���
    '��Σ�int��¼����=��Ӧ������ü�¼.��¼����
    '���Σ�ҽ��ID
    '      ���ͺ�
    '���أ���ҽ������,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-03-03 15:22:06
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    lngҽ��ID = 0: lng���ͺ� = 0
    
    On Error GoTo errH
            
    strSQL = " Select ҽ����� From ������ü�¼ Where Rownum=1 And ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    
    If Not rsTmp.EOF Then lngҽ��ID = Nvl(rsTmp!ҽ�����, 0)
    If lngҽ��ID <> 0 Then
    
        strSQL = "Select ���ͺ� From ����ҽ������  Where ҽ��ID=[3] And NO=[1] And ��¼����=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag, lngҽ��ID)
        
        If Not rsTmp.EOF Then lng���ͺ� = rsTmp!���ͺ�
    End If
    
    If lngҽ��ID <> 0 And lng���ͺ� <> 0 Then
        BillisAdviceMoney = True
    Else
        lngҽ��ID = 0: lng���ͺ� = 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistDrug(ByVal strNo As String, ByVal bytFlag As Byte) As Long
'���ܣ��ж�һ�ŵ������Ƿ����δ���䷢ҩ���ڵ�ҩƷ,�����ҩ��ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "" & _
    " Select ִ�в���ID From ������ü�¼" & _
    " Where �շ���� IN('5','6','7') And NO=[1] And ��¼״̬ IN(0,1) And ��¼����=[2]" & _
    "       And ִ�в���ID is Not NULL And ��ҩ���� is NULL" & _
    " Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    If Not rsTmp.EOF Then
        BillExistDrug = rsTmp!ִ�в���ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillDrugDept(ByVal strNo As String, lng��ҩ�� As Long, lng��ҩ�� As Long, lng��ҩ�� As Long) As Boolean
'���ܣ���ȡһ���շѵ��ݵ�ҩ��ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "" & _
    " Select �շ����,ִ�в���ID " & _
    " From ������ü�¼" & _
    " Where �շ���� IN('5','6','7') And NO=[1] And ��¼״̬ IN(0,1,3) And ��¼����=1" & _
    "       And ִ�в���ID is Not NULL" & _
    " Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        If rsTmp!�շ���� = "5" Then
            lng��ҩ�� = rsTmp!ִ�в���ID
        ElseIf rsTmp!�շ���� = "6" Then
            lng��ҩ�� = rsTmp!ִ�в���ID
        ElseIf rsTmp!�շ���� = "7" Then
            lng��ҩ�� = rsTmp!ִ�в���ID
        End If
    End If
    BillDrugDept = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Read�����ʻ�����(ByVal lng����ID As Long) As Currency
'���ܣ���ȡָ�������¼�и����ʻ�֧���Ľ��(��Ԥ����ʽ����)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ��Ԥ�� From ����Ԥ����¼ A,���㷽ʽ B" & _
        " Where A.���㷽ʽ=B.���� And B.����=3" & _
        " And A.��¼���� Not IN(1,11) And A.����ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID)
    If Not rsTmp.EOF Then
        Read�����ʻ����� = Nvl(rsTmp!��Ԥ��, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Billδ�շ�(ByVal strNo As String, _
    ByVal bytFlag As Byte) As Boolean
    '���ܣ����ָ���������Ƿ����δ�շ�(δ���)�ķ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select id From ������ü�¼ Where NO=[1] And ��¼����=[2] And ��¼״̬=0"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    Billδ�շ� = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillPay(ByVal strNo As String, _
    ByRef cur��Ԥ���� As Currency, ByRef curӦ�ɽ�� As Currency) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѵ���Ԥ�����ϼ�,���ɿ�ϼ�
    '���:
    '����:cur��Ԥ����-���س�Ԥ��;curӦ�ɽ��-����Ӧ�ɽ��(�ֽ�ĺͷ�ҽ���ࣨ����������))
    '����: 0(��ʱδ����)
    '����:���˺�
    '����:2014-06-18 10:50:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "" & _
    " Select Sum(Decode(Mod(b.��¼����, 10), 1, b.��Ԥ��, 0)) ��Ԥ��," & vbNewLine & _
    "        Sum(Decode(b.��¼����,3,Decode(c.����, 1, b.��Ԥ��, 2, b.��Ԥ��, 0),0)) �ɿ���" & vbNewLine & _
    " From ����Ԥ����¼ B, ���㷽ʽ C" & vbNewLine & _
    " Where b.���㷽ʽ = c.����" & vbNewLine & _
    "       And b.����id In (Select ����id From ������ü�¼ Where Mod(��¼����, 10) = 1 And NO = [1])"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        cur��Ԥ���� = Val("" & rsTmp!��Ԥ��)
        curӦ�ɽ�� = Val("" & rsTmp!�ɿ���)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceName(ByVal strNo As String) As String
'���ܣ���ȡ�շѵ���ԭ��ҽ�����㷽ʽ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "" & _
    " Select C.���� " & _
    " From ������ü�¼ A,����Ԥ����¼ B,���㷽ʽ C" & _
    " Where A.����ID=B.����ID And B.��¼����=3 And B.���㷽ʽ=C.����" & _
    "       And Nvl(C.����,1) IN(1,2) And A.��¼����=1 And A.��¼״̬ IN(1,3)" & _
    "       And A.NO=[1] And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then GetBalanceName = Nvl(rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDelBalanceID(ByVal strNo As String) As Long
'���ܣ���ȡ�˷Ѽ�¼�Ľ���ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ����ID From ������ü�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=2 And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then GetDelBalanceID = Val("" & rsTmp!����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillIdentical(ByVal strNo As String) As Boolean
'���ܣ��ж�ָ���ļ��ʵ����е�״̬�Ƿ�һ��,���Ƿ�ͬʱ������˺�δ��˵�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    BillIdentical = True
    
    On Error GoTo errH

    strSQL = _
    " Select Count(Distinct �Ǽ�ʱ��) as ʱ����," & _
    "        Sum(Decode(��¼״̬,0,1,0)) as δ���," & _
    "        Sum(Decode(��¼״̬,0,0,1)) as �����" & _
    " From ������ü�¼" & _
    " Where ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=2"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!δ���, 0) <> 0 And Nvl(rsTmp!�����, 0) <> 0 Then
            BillIdentical = False
        ElseIf Nvl(rsTmp!ʱ����, 0) > 1 Then
            BillIdentical = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AuditingWarn(ByVal strPrivs As String, ByVal strNo As String, ByVal str��� As String) As Boolean
'���ܣ���˻��۵�ʱ���Է��ý��б���
'������str���=ָ��������Ҫ��˵��к�,Ϊ�ձ�ʾ������
    Dim rsWarn As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String, j As Long, str���s As String
    Dim cur���ն� As Currency, cur��� As Currency, cur��� As Currency
    Dim strWarn As String, intWarn As Integer
    
    strSQL = "" & _
    " Select A.�����־, A.����, A.����id , E.Ԥ����� - E.������� As ���, B.������, C.���� As ������," & vbNewLine & _
    "        A.�շ����, D.���� As �������, Sum(A.ʵ�ս��) As ���, Zl_Patiwarnscheme(A.����id) As ���ò���" & vbNewLine & _
    " From ������ü�¼ A, ������Ϣ B, ҽ�Ƹ��ʽ C, �շ���Ŀ��� D, ������� E" & vbNewLine & _
    " Where A.��¼���� = 2 And A.��¼״̬ = 0 And A.NO = [1] And A.�շ���� = D.���� And A.����id = E.����id(+) And" & vbNewLine & _
    "       E.����(+) = 1 And A.����id = B.����id And B.ҽ�Ƹ��ʽ = C.����(+)" & vbNewLine & _
            IIf(str��� <> "", " And Instr([2],','||Nvl(A.�۸񸸺�,A.���)||',')>0", "") & _
    " Group By Nvl(A.�۸񸸺�, A.���), A.�����־, A.����, A.����id,  B.������, E.Ԥ�����, E.�������, C.����," & vbNewLine & _
    "         A.�շ����, D.����"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, "," & str��� & ",")
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If InStr(str���s, rsTmp!�շ���� & rsTmp!�������) = 0 Then
                str���s = str���s & "," & rsTmp!�շ���� & rsTmp!�������
            End If
            cur��� = cur��� + rsTmp!���
            rsTmp.MoveNext
        Loop
        rsTmp.MoveFirst
        str���s = Mid(str���s, 2)
        
        If cur��� > 0 Then
            Set rsWarn = GetUnitWarn(rsTmp!���ò���, "0")
                        
            cur���ն� = GetPatiDayMoney(rsTmp!����ID)
            cur��� = Nvl(rsTmp!���, 0)
            If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(0, rsTmp!����ID) + cur���
            '���౨��
            For j = 0 To UBound(Split(str���s, ","))
                intWarn = BillingWarn(strPrivs, rsTmp!����, rsTmp!���ò���, rsWarn, _
                    cur���, cur���ն�, cur���, Nvl(rsTmp!������, 0), _
                    Left(Split(str���s, ",")(j), 1), Mid(Split(str���s, ",")(j), 2), strWarn)
                If intWarn = 2 Or intWarn = 3 Then Exit Function
            Next
        End If
    End If
    
    AuditingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckValidity(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, Optional ByVal blnAsk As Boolean = True) As Boolean
'���ܣ�����������ϵ����Ч���Ƿ����
'˵����blnAsk=��ʾ�Ƿ�ѯ���Ƿ����,����Ϊ����
    Dim rsTmp As ADODB.Recordset
    Dim Curdate As Date, minDate As Date
    Dim strSQL As String, strName As String
    
    CheckValidity = True
    
    '��һ���Բ��ϲ��ж�
    '��Ϊ���ܸ��������Ч�ڲ�ͬ,���Ҫ�õ�����������С��Ч��
    strSQL = _
        " Select C.����,Nvl(B.����,0) as ����," & _
        " B.�������� as ���,B.���Ч��,Sysdate as ʱ��" & _
        " From �������� A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
        " Where A.����ID=B.ҩƷID And A.����ID=C.ID And A.һ���Բ���=1" & _
        " And B.����=1 And Nvl(B.��������,0)>0 And A.���Ч�� is Not NULL" & _
        " And A.����ID=[1] And B.�ⷿID=[2]" & _
        " Order by Nvl(B.����,0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID, lng�ⷿID)
    If Not rsTmp.EOF Then
        strName = rsTmp!����
        Curdate = rsTmp!ʱ��
        minDate = CDate("3000-01-01")
            
        Do While Not rsTmp.EOF
            If rsTmp!���Ч�� < minDate Then
                minDate = rsTmp!���Ч��
            End If
            If Nvl(rsTmp!���, 0) < dbl���� Then
                dbl���� = dbl���� - Nvl(rsTmp!���, 0)
            Else
                dbl���� = 0
            End If
            If dbl���� = 0 Then Exit Do
            rsTmp.MoveNext
        Loop

        If Curdate > minDate Then
            If blnAsk Then
                If MsgBox("��������""" & strName & """�����Ч��""" & Format(minDate, "yyyy-MM-dd") & """�ѹ���,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckValidity = False
                End If
            Else
                MsgBox "���ѣ�" & vbCrLf & vbCrLf & "��������""" & strName & """�����Ч��""" & Format(minDate, "yyyy-MM-dd") & """�ѹ��ڡ�", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistRegist(ByVal lng����ID As Long, Optional ByRef blnPeisPriceBill As Boolean, _
    Optional ByVal blnSaveBillCheck As Boolean) As Boolean
    '���ܣ��ж�ָ�������Ǵ�����Ч�ĹҺż�¼
    '���:
    '   blnSaveBillCheck - �Ƿ񱣴�ʱ��飬�����
    '���Σ�
    '   blnPeisPriceBill-�ò����Ƿ������컮�۵�
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    '����������ʱ��ʾ�����,��Ȼ����
    If gTy_System_Para.Sy_Reg.bytNODaysGeneral = 0 And gTy_System_Para.Sy_Reg.bytNoDayseMergency = 0 Then ExistRegist = True: Exit Function
    
    If blnSaveBillCheck = False Then
        '��첡�˲������Ƿ�Һ�
        '102660,�ж��Ƿ���첡�ˣ�����ͨ��������ҽ����¼"���ж��Ƿ���첡���ˣ���Ҫ���ݻ��۵����Ƿ�������������ж�
        strSQL = "Select 1" & vbNewLine & _
                " From ������ü�¼" & vbNewLine & _
                " Where ��¼���� = 1 And ��¼״̬ = 0 And Nvl(�����־, 0) = 4 And ����id = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID)
        If Not rsTmp.EOF Then
            blnPeisPriceBill = True
            ExistRegist = True: Exit Function
        End If
    End If
    
    strSQL = "Select 1 From ���˹Һż�¼  " & _
            " Where RowNum<2 And ����ID=[1] and ��¼����=1 and ��¼״̬=1 " & zlGetRegEventsCons(, , True)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID)
    ExistRegist = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckRegisted(ByVal lng����ID As Long, Optional ByRef blnPeisPriceBill As Boolean, _
    Optional ByVal blnSaveBillCheck As Boolean) As Boolean
'����:�շ�ʱ,��鲡���Ƿ�Һ�
'���:
'   blnSaveBillCheck - �Ƿ񱣴�ʱ��飬�����
'���Σ�
'   blnPeisPriceBill-�ò����Ƿ������컮�۵�
    blnPeisPriceBill = False
    
    If gbytUnRegevent = 0 Then CheckRegisted = True: Exit Function
    If lng����ID <> 0 Then
        If Not ExistRegist(lng����ID, blnPeisPriceBill, blnSaveBillCheck) Then lng����ID = 0
    End If
           
    If lng����ID = 0 Then
        If gbytUnRegevent = 1 Then
            If MsgBox("����û�йҺ�,��ȷ��Ҫ�����շ���?", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
                CheckRegisted = False: Exit Function
            End If
        Else
            Call MsgBox("����û�йҺ�,����������շ�.", vbInformation, gstrSysName)
            CheckRegisted = False: Exit Function
        End If
    End If
    CheckRegisted = True
End Function


Public Function CheckAddedItem(ByVal lng����ID As Long, Optional str�������� As String) As Boolean
'���ܣ��ж�ָ�������Ƿ���Ҫ�Զ�����ָ���շ���Ŀ
'      �������Զ����յ�����:
'       1.û���趨�Һŵ���Ч����(Ϊ��)
'       2.û���趨�Զ�������Ŀ,�����ĿID�Ѳ�����
'       3.�Һŵ���Ч������,û�йҺ�,û���Զ�����
'       4.���ز��������˲��������롰���������
    Dim rsTmp As ADODB.Recordset, strSQL As String, strID As String
    Dim strWhere As String
    On Error GoTo errH
    If (gTy_System_Para.Sy_Reg.bytNODaysGeneral = 0 And gTy_System_Para.Sy_Reg.bytNoDayseMergency) Or glngAddedItem = 0 Then Exit Function
    If gstr�շ���� <> "" And InStr(1, "," & gstr�շ���� & ",", ",'Z',") = 0 Then Exit Function
    
    If lng����ID = 0 Then
        strID = "And ���� = [2]"
    Else
        strID = "And ����id + 0 = [1]"
    End If
    
    strWhere = "   (nvl(����,0) =0 and ����ʱ�� > Trunc(Sysdate - " & gTy_System_Para.Sy_Reg.bytNODaysGeneral & ")) "
    strWhere = strWhere & "  Or  (nvl(����,0) =1 and ����ʱ�� > Trunc(Sysdate - " & gTy_System_Para.Sy_Reg.bytNoDayseMergency & ")) "
    strWhere = " And (" & strWhere & ") "
    
    strSQL = "" & _
    " Select 1" & vbNewLine & _
    " From Dual" & vbNewLine & _
    " Where     Not Exists (Select 1 From ���˹Һż�¼ Where ��¼����=1 and ��¼״̬=1 and  Rownum < 2 " & strID & strWhere & ")" & vbNewLine & _
    "       And Not Exists (Select 1" & vbNewLine & _
    "                       From ������ü�¼" & vbNewLine & _
    "                       Where Rownum < 2 " & strID & " And �Ǽ�ʱ�� > Trunc(Sysdate - " & gTy_System_Para.Sy_Reg.bytNODaysGeneral & ") " & vbNewLine & _
    "                             And �շ�ϸĿid + 0 = [3] And ��¼���� = 1 And ��¼״̬ = 1) " & _
    "       And Exists (Select 1 From �շ���ĿĿ¼ Where ID = [3])"
    
    '���˺� ����:34717    ����:2010-12-20 16:09:26
    '��������������޷������Ƿ��Ѿ�������ָ�����շ���Ŀ,���,������ͨ�Һ���Ч����Ϊ׼
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID, str��������, glngAddedItem)
    CheckAddedItem = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemUnderSet(ByVal str��� As String, ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal dbl��� As Double) As Boolean
'���ܣ����ָ��ҩƷ/������ָ���ⷿ�Ŀ���Ƿ���ڴ�������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '�޼�¼���ֶ�Ϊ��/�ձ�ʾδ����
    If str��� = "4" Then
        strSQL = "Select �ⷿID,����ID,����,����,�̵�����,�ⷿ��λ From ���ϴ����޶� Where ����ID=[1] And �ⷿID=[2] And Nvl(����,0)<>0 And ����>[3]"
    Else
        strSQL = "Select �ⷿID,ҩƷID,����,����,�̵�����,�ⷿ��λ From ҩƷ�����޶� Where ҩƷID=[1] And �ⷿID=[2] And Nvl(����,0)<>0 And ����>[3]"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngҩƷID, lng�ⷿID, dbl���)
    If Not rsTmp.EOF Then ItemUnderSet = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistFact(strNo As String) As Boolean
'���ܣ��ж�ָ���������Ƿ���ڹ�����
'������strNO=���ŵ����е�һ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    'Ӧȡ���һ�δ�ӡ��������
    strSQL = "Select Max(ID) From Ʊ�ݴ�ӡ���� Where ��������=1 And NO=[1]"
    strSQL = "Select Count(A.ID) as NUM From ������ü�¼ A,Ʊ�ݴ�ӡ���� B" & _
        " Where A.NO=B.NO And A.��¼����=1 And A.��¼״̬=1 And Nvl(A.���ӱ�־,0)=8 And B.��������=1 And B.ID=(" & strSQL & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        BillExistFact = Nvl(rsTmp!Num, 0) > 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckSingleBalance(ByVal strNo As String) As Boolean
'���ܣ��ж�ָ���������Ƿ�ֻ��һ�ַ�ҽ�����㷽ʽ(��Ԥ������)
'       :strNO(��ʽΪ"E01,E02"):����:34035
'������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strNo = Replace(strNo, "'", "")
    CheckSingleBalance = True
    strSQL = "" & _
    " Select /*+ rule */ Count(Distinct A.���㷽ʽ) num" & vbNewLine & _
    " From ����Ԥ����¼ A, ���㷽ʽ B,Table( f_Str2list([1])) J" & vbNewLine & _
    " Where   A.��¼���� = 3 And A.��¼״̬ In (1, 3) " & _
    "           And A.���㷽ʽ = B.���� And B.���� In (1, 2)  And A.NO = J.Column_Value"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If rsTmp!Num > 1 Then CheckSingleBalance = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckTest(ByVal strNos As String, Optional ByVal DatBegin As Date, Optional ByVal DatEnd As Date) As Boolean
    '���ܣ���鵱ǰѡ��Ҫ�շѵĵ������Ƿ����Ƥ�Խ��Ϊ���Ի�û��Ƥ�Եļ�¼
    '������strNOs=���ݺ�(��ʽΪ"E01,E02")',���ŵ���ʱ,�봫��DatBegin��DatEnd����
    '���أ�ΪFalseʱ��ʾ����������շ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strIF As String
    Dim i As Long, strInfo As String, strTmp As String, blnHaveMateria As Boolean
    Dim str������ĿID As String, rsTest As ADODB.Recordset
    Dim strTable As String
    CheckTest = True
    
    '62020:�е��ݣ��Ͳ�Ӧ�ô������ڷ�Χ�Ĳ���
'    If InStr(1, strNos, ",") > 0 Then
'        strIF = " And A.�Ǽ�ʱ�� Between [2] And [3] And Instr(','||[1]||',',','||A.NO||',')>0"
'    Else
'        strIF = " And A.NO = [1]"
'    End If
'
    strSQL = _
    " Select /*+ rule */ Distinct B.ID,B.ҽ������,B.Ƥ�Խ��,a.�շ����,B.������ĿID" & _
    " From ������ü�¼ A,����ҽ����¼ B,������ĿĿ¼ C,Table(f_Str2list([1])) J" & _
    " Where A.��¼����=1 And A.��¼״̬=0  And A.������ IS Not NULL And A.����Ա���� IS NULL" & _
    "       And A.NO = J.Column_Value " & _
    "       And A.ҽ�����=B.ID And B.������ĿID=C.ID And C.���='E' And C.��������='1'" & _
    "       And (B.Ƥ�Խ�� is NULL Or B.Ƥ�Խ��='(+)')"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNos, DatBegin, DatEnd)
    If Not rsTmp.EOF Then
        '����:33110
        str������ĿID = ""
        For i = 1 To rsTmp.RecordCount
            If InStr(1, str������ĿID & ",", "," & Nvl(rsTmp!������ĿID) & ",") = 0 Then
                    str������ĿID = str������ĿID & "," & Nvl(rsTmp!������ĿID)
            End If
            rsTmp.MoveNext
        Next
        If Len(str������ĿID) > 2000 Then
            strTable = " (Select distinct B.��ĿID as ����ID,B.�÷�ID From ������ĿĿ¼ A,�����÷����� B Where A.ID=B.�÷�ID and nvl(B.����,0)=0 and  A.ID In (" & Mid(str������ĿID, 2) & ") ) J"
        ElseIf str������ĿID <> "" Then
            str������ĿID = Mid(str������ĿID, 2)
            strTable = " (Select distinct B.��ĿID as ����ID,B.�÷�ID From Table(Cast(f_num2list([4]) As Zltools.t_Numlist )) A ,�����÷����� B Where a.Column_Value=B.�÷�ID and nvl(B.����,0)=0) J"
        Else
            strTable = " (Select -1 as ����ID,0 as �÷�ID From dual) J"
        End If
        strSQL = "" & _
        "   Select /*+ rule */ distinct  M.����,M.����,M.���,J.�÷�ID,J.����ID" & _
        "   From ������ü�¼ A,�շ���ĿĿ¼ M,ҩƷ��� B,������ĿĿ¼ C,Table(f_Str2list([1])) M ," & _
                    strTable & _
        "   Where A.��¼����=1 And A.��¼״̬=0 And  A.������ IS Not NULL And A.����Ա���� IS NULL " & _
        "              And A.�շ���� In('5','6','7')  And A.NO = M.Column_Value" & _
        "              And A.�շ�ϸĿID =M.ID " & _
        "              And a.�շ�ϸĿID=b.ҩƷID and B.ҩ��ID=C.ID And  B.ҩ��ID=J.����ID "
        Set rsTest = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNos, DatBegin, DatEnd, str������ĿID)
        If rsTest.RecordCount > 0 Then blnHaveMateria = True
        
        If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            strTmp = rsTmp!ҽ������ & "��" & IIf(IsNull(rsTmp!Ƥ�Խ��), "��Ƥ�Խ��", "���Ϊ����(+)")
            If InStr(1, strInfo, strTmp) = 0 Then
                strInfo = strInfo & vbCrLf & strTmp & "����"
                rsTest.Filter = "�÷�ID=" & Val(Nvl(rsTmp!������ĿID))
                If Not rsTest.EOF Then
                    strInfo = strInfo & "ҩƷ:" & Nvl(rsTest!����) & "����"
                End If
            End If
            rsTmp.MoveNext
        Next
        If blnHaveMateria Then
            strInfo = "�շѵ����У�����Ƥ���޽����Ϊ���ԣ�" & vbCrLf & strInfo & vbCrLf & vbCrLf & "�������շ�!"
            MsgBox strInfo, vbInformation, gstrSysName
        Else
            strInfo = "�շѵ����У�����Ƥ���޽����Ϊ���ԣ�" & vbCrLf & strInfo & vbCrLf & vbCrLf & " �Ƿ�����շ�?"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckTest = False
                Exit Function
            End If
        End If
        CheckTest = Not blnHaveMateria
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetPriceBills(ByVal lngPatient As Long, ByVal lngRegDept As Long, _
                ByVal DatBegin As Date, ByVal DatEnd As Date, _
                Optional blnAddDiagnose As Boolean = False, _
                Optional bytDefaultSel As Byte = 1) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ָ������ʱ�䷶Χ�ڵĻ��۵�
    '���: lngPatient-����ID
    '        lngRegDept-�Һſ���,ͨ�����۵�����,���Ҳ���Ҫ����Һſ���ʱ,�Żᴫ��Һſ���
    '        blnAddDiagnose-�������(����:33685)
    '        bytDefaultSel-ȱʡѡ��(0-����;1-��Ч����;2-����)
    '����:���۵���¼��
    '����:���˺�
    '����:2011-03-14 10:40:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strSelected As String, strIF As String
    Dim strSubTable As String, strTbDiagnose As String
    Dim blnDeptLimit As Boolean
    Dim strDeptLimitWhere As String
    
    On Error GoTo errH
    '96357
    If gTy_Module_Para.str�����շ�ִ�п��� <> "" Then
        strDeptLimitWhere = vbNewLine & _
            " And Exists(Select 1" & vbNewLine & _
            "      From ������ü�¼ M" & vbNewLine & _
            "      Where m.��¼���� = a.��¼���� And m.No = a.No" & vbNewLine & _
            "           And Instr('," & gTy_Module_Para.str�����շ�ִ�п��� & ",', ','||m.ִ�в���id||',') > 0)"
    ElseIf gTy_Module_Para.str�������շ�ִ�п��� <> "" Then
        strDeptLimitWhere = vbNewLine & _
            " And Not Exists(Select 1" & vbNewLine & _
            "      From ������ü�¼ M" & vbNewLine & _
            "      Where m.��¼���� = a.��¼���� And m.No = a.No" & vbNewLine & _
            "           And Instr('," & gTy_Module_Para.str�������շ�ִ�п��� & ",', ','||m.ִ�в���id||',') > 0)"
    End If
    
    'decode(A.ִ��״̬,-1,NULL,'��'):����:38281
    '�ϴ�δѡ���,ִ��״̬Ϊ-1��ʾ��ִͣ��
    If lngRegDept > 0 Then
        strSelected = "Decode(A.��������ID," & lngRegDept & ",decode(A.ִ��״̬,-1,NULL,'��'),NULL)"
    Else
        If bytDefaultSel = 0 Then
            '36438:gintSeekDays
            strSelected = "Decode(Sign(Trunc(Sysdate)-Trunc(A.�Ǽ�ʱ��)),0,decode(A.ִ��״̬,-1,NULL,'��'),-1,decode(A.ִ��״̬,-1,NULL,'��'),NULL)"
        ElseIf bytDefaultSel = 1 Then   'ȱʡ��Ч����
            '36438:gintSeekDays
            strSelected = "Decode(Sign( Sysdate-" & gintSeekDays & "-A.�Ǽ�ʱ��),0,decode(A.ִ��״̬,-1,NULL,'��'),-1,decode(A.ִ��״̬,-1,NULL,'��'),NULL)"
        Else
            strSelected = " decode(A.ִ��״̬,-1,NULL,'��') "
        End If
    End If
    strIF = " And A.����ID=[1] And A.�Ǽ�ʱ�� Between [2] And [3] "
    strTbDiagnose = ""
    '��:Wmsys.Wm_Concat��Ϊ��f_List2Str(Cast(collect ()))�ķ�ʽ.ԭ����oracle10gĿǰֻ�ǲ��԰�
    '����:38528
    If blnAddDiagnose Then '����ҽ��
        strTbDiagnose = "" & _
        "         ,( Select distinct A.NO,  f_List2str(Cast(COLLECT(distinct Q.������� ) as t_Strlist))  as ���" & vbNewLine & _
        "           From A  ,�������ҽ�� J,������ϼ�¼ Q  " & vbNewLine & _
        "           Where  A.ҽ�����=J.ҽ��ID and J.���ID=Q.ID " & vbNewLine & _
        "           Group by  A.NO  ) C"
    End If
    strSubTable = "" & _
                " Select " & strSelected & " as ѡ�� ,A.��������ID, " & vbNewLine & _
                "       A.NO,A.������,A.����,A.�Ա�,A.����,A.Ӧ�ս��,A.ʵ�ս��," & vbNewLine & _
                "       A.������,A.�Ǽ�ʱ��  As ����ʱ��,nvl(B.���ID,A.ҽ����� ) as ҽ�����, " & vbNewLine & _
                "       decode(C.ID,NULL,0,1) as Ƥ��" & vbNewLine & _
                " From ������ü�¼ A,����ҽ����¼ B,������ĿĿ¼ C" & vbNewLine & _
                " Where A.��¼����=1 And A.��¼״̬=0 And A.ҽ�����=B.ID(+)" & vbNewLine & _
                "       And  B.������ĿID=C.ID(+) And C.���(+)='E' And C.��������(+)='1' " & strIF & vbNewLine & _
                strDeptLimitWhere
   
   strSQL = _
        " Select * From ( with A as ( " & strSubTable & ")   " & vbNewLine & _
        " Select " & IIf(blnAddDiagnose, "       nvl(Max(C.���),' ') as ���,", "") & vbNewLine & _
        "       A.ѡ��,A.NO as ���ݺ�,B.���� as ��������,A.������ as ҽ��,Ltrim(A.����) as ����,A.�Ա�,A.����," & vbNewLine & _
        "       ltrim(To_Char(Sum(A.Ӧ�ս��),'99999" & gstrDec & "')) as Ӧ�ս��," & vbNewLine & _
        "       ltrim(To_Char(Sum(A.ʵ�ս��),'99999" & gstrDec & "')) as ʵ�ս��," & vbNewLine & _
        "       A.������,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��, " & vbNewLine & _
        "       Decode(nvl(Max(A.Ƥ��),0),1,'��','') as Ƥ��" & vbNewLine & _
        " From A,���ű� B" & vbCrLf & strTbDiagnose & vbNewLine & _
        " Where  A.��������ID=B.ID  " & IIf(blnAddDiagnose, " And A.NO=C.NO(+)", "") & vbNewLine & _
        " Group by A.ѡ��,A.NO,B.����,A.������,A.����,A.�Ա�,A.����,A.������,A.����ʱ��,A.��������ID) " & vbNewLine & _
        " Order by ���ݺ� Desc"
        '" Order by  ����ʱ�� Desc"'102748,����ʱ�併�����Ϊ�����ݺŽ���
    Set GetPriceBills = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngPatient, DatBegin, DatEnd)
    
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReadABCNum(ByVal strPrivs As String) As Boolean
'����:��ȡ��ҩ������
'������strPrivs=���ڸ���Ȩ�޿����Ƿ��ȡС����ݲ���
'���أ���������Ŀ����ĸ
    Dim strSQL As String
        
    On Error GoTo errH
    
    If InStr(strPrivs, "ҩƷ����С��") > 0 Then
        strSQL = "Select Upper(����) as ����,��ֵ From ��ҩ������ Order by ����"
    Else
        strSQL = "Select Upper(����) as ����,��ֵ From ��ҩ������ Where Trunc(��ֵ)=��ֵ Order by ����"
    End If
    
    Set grsABCNum = New ADODB.Recordset 'Filter��Newʱ���
    Call zlDatabase.OpenRecordset(grsABCNum, strSQL, "mdlPublic")
    
    '��ȡ��������Ŀ����ĸ
    gstrABC = ""
    Do While Not grsABCNum.EOF
        '�����ĸֻ��һλ,��ֻ��Ϊ��ĸ
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Left(grsABCNum!����, 1)) > 0 Then
            gstrABC = gstrABC & Left(grsABCNum!����, 1)
        End If
        grsABCNum.MoveNext
    Loop
    ReadABCNum = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ConvertABCtoNUM(ByVal strInput As String) As String
'���ܣ�������ҩ�����ݶ��壬������ת��Ϊ����
'����1.�����ĸǰ��������������
'      2.����ͬʱ�����������ĸ
'      3.�����ĸ����������,С��������
'      4.���������Ϲ����򷵻�0
    Dim strBit As String, strNum As String, i As Long
    Dim blnABC As Boolean, blnNum As Boolean
    
    If strInput = "" Then ConvertABCtoNUM = "": Exit Function
    strInput = UCase(strInput)
    
    For i = 1 To Len(strInput)
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(strInput, i, 1)) > 0 Then
            If blnNum Or blnABC Then strNum = "0": Exit For
            
            grsABCNum.Filter = "����='" & Mid(strInput, i, 1) & "'"
            If Not grsABCNum.EOF Then
                strBit = FormatEx(grsABCNum!��ֵ, 5)
            Else
                strNum = "0": Exit For
            End If
            
            blnABC = True
        ElseIf InStr("0123456789.", Mid(strInput, i, 1)) > 0 Then
            If blnABC Then strNum = "0": Exit For
            strBit = Mid(strInput, i, 1)
            blnNum = True
        Else
            strBit = Mid(strInput, i, 1)
        End If
        strNum = strNum & strBit
    Next
    ConvertABCtoNUM = strNum
End Function

Public Function CheckDeptIsMedTech(ByVal lngDeptID As Long) As Boolean
'���ܣ����ָ���Ĳ����Ƿ���ҽ�����������
'����:
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    
    strSQL = "Select 1 From ��������˵�� Where ����id = [1] And �������� In('���','����','����','����','Ӫ��','���')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngDeptID)
    CheckDeptIsMedTech = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckMediCareItem(ByVal lng�շ�ϸĿID As Long, ByVal int���� As Integer, _
    ByVal str�շ���Ŀ���� As String, ByVal bln���� As Boolean, Optional ByVal strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��շ���Ŀ�Ƿ������˱���֧����Ŀ
    '���:bln����-��ǰ�Ƿ�Ϊ����
    '     dbl�۸�-���۵ļ۸�
    '����:
    '����:1.����ķ���true,���򷵻�False
    '     2.�����ҽ������,����true
    '     3.���۵��Ҽ۸�=0�Ĳ����
    '����:���˺�
    '����:2010-01-07 14:44:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset, strSQL As String, rs�۸� As ADODB.Recordset, dbl�۸� As Double
    Dim strWherePriceGrade As String
        
    CheckMediCareItem = True
    If gbytҽ�������� = 0 Then Exit Function
    On Error GoTo errH
    
    '���˺� ����:27286 ���۵ļ۸�Ϊ��Ĳ����м����� ����:2010-01-07 15:13:45
    If bln���� Then
        If strPriceGrade <> "" Then
            strWherePriceGrade = _
                "      And (b.�۸�ȼ� = [2]" & vbNewLine & _
                "          Or (b.�۸�ȼ� Is Null" & vbNewLine & _
                "              And Not Exists(Select 1" & vbNewLine & _
                "                             From �շѼ�Ŀ" & vbNewLine & _
                "                             Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
                "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
        Else
            strWherePriceGrade = " And b.�۸�ȼ� Is Null"
        End If
        strSQL = _
            "Select b.�ּ�" & vbNewLine & _
            "From �շѼ�Ŀ B" & vbNewLine & _
            "Where b.�շ�ϸĿid = [1]" & vbNewLine & _
            "      And Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
        Set rs�۸� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ǰ�۸�", lng�շ�ϸĿID, strPriceGrade)
        If rs�۸�.EOF = False Then
            dbl�۸� = Val(Nvl(rs�۸�!�ּ�))
        Else
            dbl�۸� = 0
        End If
        If dbl�۸� = 0 Then Exit Function
    End If
     
    strSQL = "Select �շ�ϸĿID From ����֧����Ŀ Where �շ�ϸĿID=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng�շ�ϸĿID, int����)
        
    If rsTmp.RecordCount = 0 Then
        If gbytҽ�������� = 1 Then
            If MsgBox("û������""" & str�շ���Ŀ���� & """��Ӧ�ı�����Ŀ,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckMediCareItem = False
            End If
        ElseIf gbytҽ�������� = 2 Then
            MsgBox "û������""" & str�շ���Ŀ���� & """��Ӧ�ı�����Ŀ!", vbInformation, gstrSysName
            CheckMediCareItem = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional blnδ�ҵ����� As Boolean = False, Optional strOra���� As String, Optional strWhere As String, _
    Optional blnվ�� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '     blnվ��-�Ƿ����վ������
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str���� As String, str���� As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    str���� = strKey
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   And ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  "
    End If
    gstrSQL = gstrSQL & strWhere & IIf(blnվ��, zl_��ȡվ������, "") & _
    "   order by ����"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If blnδ�ҵ����� Then
            If zlCommFun.IsCharChinese(str����) = False Then GoTo NOAdd::
            If MsgBox("ע��:" & vbCrLf & _
                   "     δ�ҵ���ص�" & strTable & ",�Ƿ����ӡ�" & str���� & "����", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If zl_AutoAddBaseItem(strTable, str����, str����, strTable & "����", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str����, str���� & "-" & str����)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = str����
                    End If
                End With
            Else
                If zlControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str����, str���� & "-" & str����)
                objCtl.Tag = str����
                zlCommFun.PressKey vbKeyTab
            End If
            zl_SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgbox "û���ҵ�����������" & strTable & ",����!"
            If objCtl.Enabled Then objCtl.SetFocus
            If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, Nvl(rsTemp!����), Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!����)
            Else
                .Text = IIf(blnOnlyName, Nvl(rsTemp!����), Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����))
            End If
        End With
    Else
        If zlControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = Nvl(rsTemp!����)
        objCtl.Tag = Nvl(rsTemp!����)
        zlCommFun.PressKey vbKeyTab
    End If
    zl_SelectAndNotAddItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_AutoAddBaseItem(ByVal strTable As String, str���� As String, str���� As String, _
    Optional strTittle As String = "������Ŀ", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Զ�������Ŀ��Ϣ(ֻ����б���,���Ƶ���Ϣ����(ֻ���ӣ����������,����)
    '--�����:
    '--������:
    '--��  ��:���ӳɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int���� As Integer, strCode As String, strSpecify As String
    zl_AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("û���ҵ��������" & strTable & "����Ҫ��������" & strTable & "����", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(����)), 2) As Length FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int���� = rsTemp!length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int���� = Len(strCode)
    strCode = strCode + 1
    
    If int���� >= Len(strCode) Then
    strCode = String(int���� - Len(strCode), "0") & strCode
    End If
    strSpecify = zlCommFun.SpellCode(str����)
    
    
    gstrSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str���� & "','" & strSpecify & "')"
    zlDatabase.ExecuteProcedure gstrSQL, strTittle
    str���� = strCode
    zl_AutoAddBaseItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

 
Public Function zl_��ȡվ������(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str���� As String = "") As String
    '����:��ȡվ����������:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str���� = "", "", str���� & ".") & "վ��"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_��ȡվ������ = strWhere
End Function
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '����:���ؼ�ƥ�䴮%dd%,�����Ǵ�д
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper = False Then
        GetMatchingSting = strLeft & strString & strRight
    Else
        GetMatchingSting = strLeft & UCase(strString) & strRight
    End If
End Function
Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False, Optional blnNotTran As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNotTran-����������
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNotTran = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub
Public Function zlIsExistsSquareCard(ByVal strNos As String, Optional bln����ȫ�� As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ�Ϊ�����㵥��
    '���:strNos-���ݺ�(����Ϊ����,�ö��ŷ���)
    '        bln����ȫ��-true:��ʾֻ����Ƿ����ȫ�˵ĵ���;False-ֻ���ˢ����¼
    '����:
    '����:����,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    Dim intType As Integer
    On Error GoTo errHandle
    intType = -1
    If bln����ȫ�� Then intType = 0
    
    '55064
    strNoIns = Replace(strNos, "'", "")
    strSQL = "" & _
    "Select /*+ rule */  B.ID  " & _
    " From ����Ԥ����¼ B, Table(f_Str2list([1])) J " & _
    " Where B.NO = J.Column_Value And  " & _
    "       (  (Nvl(B.���㿨���, 0) <> 0 And Exists(Select 1 From ���ѿ����Ŀ¼ Where ���=nvl(B.���㿨���,0) And nvl(�Ƿ�ȫ��,0)<>[2] And Nvl(�Ƿ�����,0)=0 ) ) " & _
    "          Or (Nvl(B.�����id, 0) <> 0 And Exists(Select 1 From ҽ�ƿ���� Where ID=nvl(B.�����ID,0) And nvl(�Ƿ�ȫ��,0)<>[2]  And nvl(�Ƿ�����,0)=0) )  " & _
    "       ) And B.��¼���� = 3 And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����շѵ��Ƿ����ˢ����¼", strNoIns, intType)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsOnly����ҽ��() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ��Ϊ�б���ҽ��
    '����:�ǵ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-12 09:42:08
    '����:27331
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrTmp As Variant, i As Long, strTemp As String
    arrTmp = Split(GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", ""), ",")
    strTemp = ""
    For i = 0 To UBound(arrTmp)
        If IsNumeric(arrTmp(i)) Then
            strTemp = strTemp & "," & Val(arrTmp(i))
            'If CheckMCOutMode(arrTmp(i)) Then mlngOutModeMC = Val(arrTmp(i)): Exit For '������ģʽ
        End If
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    zlIsOnly����ҽ�� = strTemp = "920"  '������:����:26982
End Function

Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲿����Ϣ�Ƿ���ر���
    '����:��ʾ����,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 13:11:01
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle

    If mlng���ű���ƽ������ = 0 Then
        strSQL = "Select Avg(length(����)) As ���� From ���ű�"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡ���ű����ƽ������")
        mlng���ű���ƽ������ = Val(Nvl(rsTemp!����))
    End If
    '���ڱ��볤�ȿ��ܹ���,�޷���ʾ���ŵ�����,����Զ���ʾ�Ͳ���ʾ����,������5ʱ,����ʾ.С��5ʱ,��ʾ
   zlIsShowDeptCode = mlng���ű���ƽ������ <= 5
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���в��� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str���в���-���в�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���в��� <> "" Then
        str���� = zlCommFun.SpellCode(str���в���)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str���в���) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ,�������," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function

Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Աѡ��ѡ����
    '���:cboSel-ָ���Ĳ���ѡ�񲿼�
    '     rsPerson-ָ������Ա��Ϣ(ID,���,����,����)
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str����-��������(������,���в���Ա��)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String, strLike As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���� <> "" Then
        str���� = zlCommFun.SpellCode(str����)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strSearch Then lngID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!���) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!���)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!ID))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboSel, lngID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboSel
End Function
Public Function zlCheckIsPrintInvoice(ByVal strNos As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����Ʊ���Ƿ���ڴ�ӡ���
    '��Σ�strNOs       =   ָ��Ҫ�ش�ĵ��ݺţ������ţ������Ƕ�����ݺţ�Ϊ"'AAA','BBB',..."����ʽ
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-05-27 22:04:21
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim cllPro As Collection, varData() As Variant
    
    On Error GoTo errHandle
    strNos = Replace(strNos, "'", "")
    
    If Len(strNos) <= 4000 Then
        strSQL = "" & _
        "   Select /*+ rule */ Max(A.ID) as ID " & _
        "   From Ʊ�ݴ�ӡ���� A,Table( f_Str2list([1])) J " & _
        "   Where A.��������=1   And A.NO=J.Column_Value"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNos)
    Else
        If zlGetSplitString4000(strNos, cllPro) = False Then Exit Function
        If zlFromCollectBulidSQL(cllPro, strSQL, varData) = False Then Exit Function
        strSQL = "With ������Ϣ as (" & strSQL & ")" & vbCrLf
        
        strSQL = "" & strSQL & _
        "   Select Max(A.ID) as ID " & _
        "   From Ʊ�ݴ�ӡ���� A,������Ϣ J " & _
        "   Where A.��������=1   And A.NO=J.NO"
        
        strSQL = "Select * From (" & strSQL & ")"
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "mdlOutExse", varData)
    End If
    If rsTemp.RecordCount <> 0 Then
        zlCheckIsPrintInvoice = Val(Nvl(rsTemp!ID)) <> 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet�����շѶ���(ByVal strҽ����� As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ��ID,��ȡ��Ӧ�������շѶ�����Ŀ
    '���:strҽ�����-���ʱ���ö��ŷָ�
    '����:�����շѶ��յ����ݼ�
    '����:���˺�
    '����:2010-11-03 14:12:35
    '����:33634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error GoTo errHandle
     strSQL = "" & _
    "   Select /*+ RULE */ A.Id as ҽ�����,B.������ĿID,B.�շ���ĿID as �շ�ϸĿID,b.�շ�����,b.���ж���,b.������Ŀ " & _
    "   From  ����ҽ����¼ A,�����շѹ�ϵ B,Table(f_num2list([1])) J" & _
    "   Where   a.ID=J.Column_Value   And a.������Ŀid=b.������ĿID And nvl(b.���ж���,0)=1"
    Set zlGet�����շѶ��� = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ������Ӧ���շѹ�ϵ", strҽ�����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub

Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
End Sub


Public Function zlGetRegEventsCons(Optional strFieldName As String = "����", _
    Optional strAliaName As String = "", Optional bln����ʱ�� As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�Һ���Ŀ����������
    '���:strFieldName-�������⻹�ֶ�(�缱��)
    '       strAliaName:����
    '����:
    '����:��������
    '����:���˺�
    '����:2010-12-20 16:33:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strTimeName As String
    strFieldName = IIf(strAliaName <> "", strAliaName & ".", "") & strFieldName
    strTimeName = IIf(strAliaName <> "", strAliaName & ".", "") & IIf(bln����ʱ��, "����ʱ��", "�Ǽ�ʱ��")
    
    With gTy_System_Para.Sy_Reg
        strWhere = ""
        If .bytNODaysGeneral <> 0 Or .bytNoDayseMergency <> 0 Then
            If .bytNODaysGeneral <> 0 Then
                strWhere = strWhere & " Or ( nvl(" & strFieldName & ",0)=0  And " & strTimeName & ">Trunc(Sysdate-" & .bytNODaysGeneral & "))"
            Else
                strWhere = strWhere & " Or  nvl(" & strFieldName & ",0)=0   "
            End If
            If .bytNoDayseMergency <> 0 Then
                strWhere = strWhere & " Or ( nvl(" & strFieldName & ",0)=1  And " & strTimeName & ">Trunc(Sysdate-" & .bytNoDayseMergency & "))"
            Else
                strWhere = strWhere & " Or nvl(" & strFieldName & ",0)=1  "
            End If
        End If
        If strWhere <> "" Then
            strWhere = " And  (" & Mid(strWhere, 4) & ")"
        End If
    End With
    zlGetRegEventsCons = strWhere
End Function


Public Function ImportWholeSet(frmParent As Object, ByVal intInsure As Integer, rsSel As ADODB.Recordset, _
    lng��ҩ�� As Long, lng��ҩ�� As Long, lng��ҩ�� As Long, _
    Optional lng����ID As Long = 0, Optional int���� As Integer = 0, _
    Optional blnҩ����λ As Boolean, Optional lng��������ID As Long = 0, Optional bytӤ���� As Byte, _
    Optional int�����־ As Integer, Optional bln�Ӱ�Ӽ� As Boolean = False, _
    Optional ByVal lngUnitID As Long, Optional int��Χ As Integer, _
    Optional str������ As String = "", Optional str������ As String = "", _
    Optional ByVal strҩƷ�۸�ȼ� As String, _
    Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ�۸�ȼ� As String, _
    Optional ByVal lng��ҳID As Long, Optional ByVal lng����ID As Long, Optional ByVal lng����ID As Long) As ExpenseBill
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���õ��ݵ����ݶ�����
    '���:rsSel-ѡ�еĳ�����Ŀ
    '       lngUnitID    ��ǰ��������ID
    '      int��Χ=1.����,2-סԺ
    '����:
    '����:��ŵ�����Ϣ�ĵ��ݶ���
    '����:���˺�
    '����:2010-09-02 16:17:54
    '˵��:��Ϊ������ʱ��Ŀ�۸���Ϣ��������,���Է�������������¼���
    '       ��������ͣ���շ�ϸĿ
    '����:    '����:34465
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue(0 To 10) As String, strSubItem As String, str�շ�ϸĿID As String, j As Long
    Dim rsItems As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lng���˿���ID As Long, strժҪ As String
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset
    Dim rsMoney As ADODB.Recordset
    Dim lngDoUnit As Long
    Dim i As Long, intCurNo As Integer
    Dim int��� As Integer, blnDo As Boolean, dblStock As Double
    Dim blnLoad As Boolean, strSQL As String, strҩ��IDs As String, strͣ����Ŀ��� As String, strPrivs As String
    Dim curModiMoney As Currency
    Dim strAdvance As String, strInfo As String
    Dim dblAllTime As Double, dblCurTime As Double, dbl�Ӱ�Ӽ��� As Double, lngLastPati As Long
    Dim colSerial As New Collection
    Dim bytType As Byte '0-����;1-סԺ;2-�����סԺ
    Dim strTable  As String
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    Set colSerial = New Collection
    '�۸�ȼ�
    If strҩƷ�۸�ȼ� <> "" Or str���ļ۸�ȼ� <> "" Or str��ͨ�۸�ȼ� <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [14])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [15])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And d.�۸�ȼ� = [16])" & vbNewLine & _
            "            Or (d.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From �շѼ�Ŀ" & vbNewLine & _
            "                                Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [14])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [15])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And �۸�ȼ� = [16])))))"
    Else
        strWherePriceGrade = " And d.�۸�ȼ� Is Null"
    End If
    
    With rsSel
        str�շ�ϸĿID = "": j = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Len(str�շ�ϸĿID) > 1990 And j <= 10 Then
                strValue(j) = Mid(str�շ�ϸĿID, 2)
                strSubItem = strSubItem & " Union ALL " & _
                " Select Column_Value as �շ�ϸĿID From Table(f_Num2List([" & j + 1 & "])) B "
                str�շ�ϸĿID = "": j = j + 1
            End If
            str�շ�ϸĿID = str�շ�ϸĿID & "," & Val(Nvl(!�շ�ϸĿID))
            .MoveNext
        Loop
    End With
    
    If str�շ�ϸĿID <> "" Then
        If j > 10 Then
             strSubItem = strSubItem & " UNION ALL Select ID From �շ���ĿĿ¼ Where id in (" & Mid(str�շ�ϸĿID, 2) & ")"
        Else
            strValue(j) = Mid(str�շ�ϸĿID, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as �շ�ϸĿID From Table(f_Num2List([" & j + 1 & "])) B "
        End If
    End If
    
    gstrSQL = "" & _
       "   Select /*+ rule */ A.����id, A.����id, A.���д���, A.�������� " & _
       "   From �շѴ�����Ŀ A, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.����id = D.�շ�ϸĿid "
    Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOutExse", strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    strSubItem = Mid(strSubItem, 11)
    strTable = " Select [13] as ����ID,�շ�ϸĿID From (" & strSubItem & ")"

    If lng��ҳID = 0 Then
        gstrSQL = "" & _
        " Select  X.ҩƷID,W.����ID,W.��������," & _
        "       F.�ѱ�,F.����,F.�Ա�,F.����,F.������," & _
        "       '' as ����,F.����� as ��ʶ��,F.����ID,0 as ��ҳID,0 as ���˲���ID,0 as ���˿���ID," & _
        "       B.��� as �շ����,A.�շ�ϸĿID," & _
        "       B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(H.����,B.����) as ����,E1.���� as ��Ʒ��,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
        "       B.���ηѱ�,B.˵��,B.ִ�п���,B.�������, B.��������  ��������,D.�ּ�,D.ԭ��,D.ȱʡ�۸�,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
        "       E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
        "       Decode(B.���,'4',1,X." & gstrҩ����װ & ") as ҩ����װ," & _
        "       Decode(B.���,'4',B.���㵥λ,X." & gstrҩ����λ & ") as ҩ����λ," & _
        "       Decode(b.���,'4',Nvl(W.���÷���,0),Nvl(X.ҩ������,0)) as ����,B.¼������, " & _
        "       M1.���� as ���Ʊ���,M1.���� as ��������,X.��ҩ��̬,x.����ϵ��,M1.���㵥λ as ������λ" & _
        "   From  (" & strTable & ") A ,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E,������Ϣ F, " & _
        "          �շ���Ŀ���� H,�շ���Ŀ���� E1,�������� W,ҩƷ��� X,������ĿĿ¼ M1" & _
        " Where  A.�շ�ϸĿID=D.�շ�ϸĿID And A.�շ�ϸĿID=B.ID " & _
        "       And b.���=C.���� And A.�շ�ϸĿID=X.ҩƷID(+) and X.ҩ��ID=M1.ID(+) And A.�շ�ϸĿID=W.����ID(+) And D.������ĿID=E.ID" & _
        "       And A.�շ�ϸĿID=H.�շ�ϸĿID(+) And H.����(+)=1 And H.����(+)=[12]" & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.����ID=F.����ID(+)" & _
        "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
        "       And Sysdate Between D.ִ������ And Nvl(D.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                strWherePriceGrade
    Else
        gstrSQL = "" & _
        " Select  X.ҩƷID,W.����ID,W.��������," & _
        "       Nvl(G.�ѱ�,F.�ѱ�) As �ѱ�,Nvl(G.����,F.����) As ����,Nvl(G.�Ա�,F.�Ա�) As �Ա�,Nvl(G.����,F.����) As ����,F.������," & _
        "       G.��Ժ���� as ����,F.����� as ��ʶ��,F.����ID,G.��ҳID,G.��ǰ����ID as ���˲���ID,G.��Ժ����ID as ���˿���ID," & _
        "       B.��� as �շ����,A.�շ�ϸĿID," & _
        "       B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(H.����,B.����) as ����,E1.���� as ��Ʒ��,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
        "       B.���ηѱ�,B.˵��,B.ִ�п���,B.�������, B.��������  ��������,D.�ּ�,D.ԭ��,D.ȱʡ�۸�,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
        "       E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
        "       Decode(B.���,'4',1,X." & gstrҩ����װ & ") as ҩ����װ," & _
        "       Decode(B.���,'4',B.���㵥λ,X." & gstrҩ����λ & ") as ҩ����λ," & _
        "       Decode(b.���,'4',Nvl(W.���÷���,0),Nvl(X.ҩ������,0)) as ����,B.¼������, " & _
        "       M1.���� as ���Ʊ���,M1.���� as ��������,X.��ҩ��̬,x.����ϵ��,M1.���㵥λ as ������λ" & _
        "   From  (" & strTable & ") A ,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E,������Ϣ F, " & _
        "       ������ҳ G,�շ���Ŀ���� H,�շ���Ŀ���� E1,�������� W,ҩƷ��� X,������ĿĿ¼ M1" & _
        " Where  A.�շ�ϸĿID=D.�շ�ϸĿID And A.�շ�ϸĿID=B.ID " & _
        "       And b.���=C.���� And A.�շ�ϸĿID=X.ҩƷID(+) and X.ҩ��ID=M1.ID(+) And A.�շ�ϸĿID=W.����ID(+) And D.������ĿID=E.ID" & _
        "       And A.�շ�ϸĿID=H.�շ�ϸĿID(+) And H.����(+)=1 And H.����(+)=[12]" & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.����ID=F.����ID(+) And F.����ID=G.����ID(+) And G.��ҳID(+) = [17]" & _
        "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
        "       And Sysdate Between D.ִ������ And Nvl(D.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                strWherePriceGrade
    End If
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOutExse", strValue(0), strValue(1), strValue(2), _
        strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10), _
        IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1), lng����ID, strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ�۸�ȼ�, lng��ҳID)
     'û�м�¼���ǿյ���
    Set objBill = New ExpenseBill
 
    Set objBill.Pages(1).Details = New BillDetails
    With rsSel
            i = 1
            If .RecordCount <> 0 Then .MoveFirst
NextRecord: Do While Not .EOF
            '����շ���Ŀ�Ƿ�ͣ�û���������ﲡ��
            '����ͣ��ʱ,��������
            rsItems.Filter = "�շ�ϸĿID=" & Val(Nvl(!�շ�ϸĿID))
            If rsItems.EOF Then 'δ�ҵ�.������
                 .MoveNext
                GoTo NextRecord:
            End If

            '����շ���Ŀ�Ƿ�ͣ�û���������ﲡ��
            '����ͣ��ʱ,��������
            If InStr(",5,6,7,", rsItems!�շ����) = 0 Then
                If InStr(1, strͣ����Ŀ��� & ",", "," & !�������� & ",") > 0 Then
                    .MoveNext
                    GoTo NextRecord
                Else
                    If Not CheckFeeItemAvailable(!�շ�ϸĿID, 1) Then
                        strͣ����Ŀ��� = strͣ����Ŀ��� & "," & !���
                        MsgBox "�����շ���Ŀ�еĵ�" & !��� & "���շ���Ŀ:" & rsItems!���� & "" & vbCrLf & _
                            "��ͣ�û��ٷ����ڲ���,�����ᱻ����." & IIf(IsNull(!��������), "����д�����Ŀ,Ҳ���ᱻ����.", ""), vbInformation, gstrSysName
                        .MoveNext
                        GoTo NextRecord
                    End If
                End If
            End If
        
            '����������=====================================================
            If i = 1 Then
                objBill.NO = ""
                objBill.Pages(1).NO = "" 'Ҫ����Ա��޸�ʱ������ֱ������ķ���
                objBill.Pages(1).��������ID = lng��������ID
                objBill.Pages(1).������ = str������
                objBill.Pages(1).ҽ����� = 0
                
                objBill.����ID = Val(Nvl(rsItems!����ID))
                objBill.��ҳID = Val(Nvl(rsItems!��ҳID))
                objBill.����ID = IIf(lng����ID = 0, Val(Nvl(rsItems!���˲���ID)), lng����ID)
                objBill.����ID = IIf(lng����ID = 0, Val(Nvl(rsItems!���˿���id)), lng����ID)
                objBill.���� = Nvl(rsItems!����)
                objBill.�Ա� = Nvl(rsItems!�Ա�)
                objBill.���� = Nvl(rsItems!����)
                objBill.��ʶ�� = Val(Nvl(rsItems!��ʶ��))
                objBill.���� = Nvl(rsItems!����)
                objBill.�ѱ� = Nvl(rsItems!�ѱ�)
                objBill.�����־ = int�����־
                objBill.�Ӱ��־ = IIf(bln�Ӱ�Ӽ�, 1, 0)
                objBill.Ӥ���� = bytӤ����
                objBill.������ = str������
                objBill.����Ա��� = UserInfo.���
                objBill.����Ա���� = UserInfo.����
                objBill.����ʱ�� = zlDatabase.Currentdate
                objBill.�Ǽ�ʱ�� = objBill.����ʱ��
                objBill.�ಡ�˵� = 0
            End If
            
            '�����շ�ϸĿ=====================================================
            Set objBillDetail = New BillDetail
            Set objBillDetail.Detail = New Detail
                        
            '�������,��������
            intCurNo = intCurNo + 1
            objBillDetail.��� = intCurNo 'ʵ�����к�
            colSerial.Add Array(Val(Nvl(!�շ�ϸĿID)), intCurNo), "_" & !���
            objBillDetail.�������� = Nvl(!��������, 0) '��Ϊ������������,�ȼ�¼ԭ����,�����ٴ���
            
            'ʹ��ԭ���Ķ�̬�ѱ�
            objBillDetail.�ѱ� = Nvl(rsItems!�ѱ�)
            objBillDetail.�շ���� = Nvl(rsItems!�շ����)
            objBillDetail.�շ�ϸĿID = Nvl(rsItems!�շ�ϸĿID)
            objBillDetail.���㵥λ = Nvl(rsItems!���㵥λ)
            objBillDetail.���� = IIf(Val(Nvl(!����)) = 0, 1, Val(Nvl(!����)))
            
            If InStr(",5,6,7,", rsItems!�շ����) > 0 And gblnҩ����λ Then
                objBillDetail.���� = Nvl(!����, 0) / Nvl(rsItems!ҩ����װ, 1)
            Else
                objBillDetail.���� = Nvl(!����, 0)
            End If
            objBillDetail.ԭʼ���� = objBillDetail.���� * objBillDetail.����
            
            objBillDetail.��ҩ���� = ""     '��Ҫ��һ��ȷ��
            
            objBillDetail.���ӱ�־ = 0
            
            objBillDetail.ժҪ = ""
            
            '���ĺ�ҩƷ����
            '����ִ�п���ȱʡΪ���˲���,�������ָ����,��Ϊָ������
            If objBillDetail.�շ���� = "4" Then
                lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, objBill.����ID)
                If lngDoUnit = 0 Then lngDoUnit = lng��������ID
            End If
            
            '���˿���ID
            lng���˿���ID = objBill.����ID
            If lng���˿���ID = 0 Then lng���˿���ID = lng��������ID
            objBillDetail.Detail.ִ�п��� = IIf(IsNull(rsItems!ִ�п���), 0, rsItems!ִ�п���)
            
            lngDoUnit = Get�շ�ִ�п���ID(objBillDetail.�շ����, objBillDetail.�շ�ϸĿID, _
                objBillDetail.Detail.ִ�п���, lng���˿���ID, lng��������ID, int��Χ, _
                IIf(lng��ҩ�� = 0, glng��ҩ��, lng��ҩ��), _
                IIf(lng��ҩ�� = 0, glng��ҩ��, lng��ҩ��), _
                IIf(lng��ҩ�� = 0, glng��ҩ��, lng��ҩ��), _
                lngDoUnit, lngUnitID)
            
            objBillDetail.ִ�в���ID = lngDoUnit
            
            objBillDetail.ԭʼִ�в���ID = objBillDetail.ִ�в���ID     '�����޸�ʱ�����жϿ��
            
            objBillDetail.Detail.ID = !�շ�ϸĿID
            objBillDetail.Detail.���� = Nvl(rsItems!����)
            objBillDetail.Detail.��� = (Val(Nvl(rsItems!�Ƿ���)) = 1)
            objBillDetail.Detail.�������� = 0 '!!!Ŀǰ���Դ�����Ŀ,����������Ŀ
            objBillDetail.Detail.���д��� = 0 '!!!Ŀǰ���Դ�����Ŀ,����������Ŀ
            objBillDetail.Detail.��� = Nvl(rsItems!���)
            objBillDetail.Detail.���㵥λ = Nvl(rsItems!���㵥λ)
            
            objBillDetail.Detail.ҩ����λ = Nvl(rsItems!ҩ����λ)
            objBillDetail.Detail.ҩ����װ = Nvl(rsItems!ҩ����װ, 1)
            
            If InStr(",4,5,6,7,", rsItems!�շ����) > 0 Then
                dblStock = GetStock(Val(Nvl(!�շ�ϸĿID)), objBillDetail.ִ�в���ID)
            Else
                dblStock = 0
            End If

            If InStr(",5,6,7,", rsItems!�շ����) > 0 And gblnҩ����λ Then dblStock = dblStock / objBillDetail.Detail.ҩ����װ
            objBillDetail.Detail.��� = dblStock
            
            
            objBillDetail.Detail.�Ӱ�Ӽ� = (Val(Nvl(rsItems!�Ӱ�Ӽ�)) = 1)
            objBillDetail.Detail.��� = Nvl(rsItems!���)
            objBillDetail.Detail.������� = Nvl(rsItems!�������)
            objBillDetail.Detail.���� = Nvl(rsItems!����)
            objBillDetail.Detail.��Ʒ�� = Nvl(rsItems!��Ʒ��)
            objBillDetail.Detail.���ηѱ� = (Val(Nvl(rsItems!���ηѱ�)) = 1)
            objBillDetail.Detail.˵�� = ""
            objBillDetail.Detail.���� = IIf(IsNull(rsItems!��������), "", rsItems!��������)
            objBillDetail.Detail.�������� = Nvl(rsItems!��������)
            objBillDetail.Detail.��ҩ��̬ = Val(Nvl(rsItems!��ҩ��̬))
            
            If objBillDetail.�������� <> 0 Then
                'A.����id, A.����id, A.���д���, A.�������� "
                rsOthers.Filter = "����ID=" & colSerial("_" & !��������)(0) & " And ����ID=" & objBillDetail.�շ�ϸĿID
                If Not rsOthers.EOF Then
                    objBillDetail.Detail.�������� = Val(Nvl(rsOthers!��������))
                    objBillDetail.Detail.���д��� = Val(Nvl(rsOthers!���д���))
                End If
            End If
            
            If InStr(",5,6,7,", rsItems!�շ����) > 0 Then
                objBillDetail.Detail.����ְ�� = Get����ְ��(objBillDetail.Detail.ID)
                objBillDetail.Detail.�������� = Get��������(objBillDetail.Detail.ID)
            End If
            objBillDetail.Detail.¼������ = Val(Nvl(rsItems!¼������))
            
            objBillDetail.Detail.ҩ��ID = Val(Nvl(rsItems!ҩ��ID))
            objBillDetail.Detail.��� = Val(Nvl(rsItems!�Ƿ���)) = 1
            objBillDetail.Detail.���� = Val(Nvl(rsItems!����)) = 1
            objBillDetail.Detail.�������� = Val(Nvl(rsItems!��������)) = 1
            objBillDetail.Detail.������λ = Nvl(rsItems!������λ)
            objBillDetail.Detail.����ϵ�� = Val(Nvl(rsItems!����ϵ��))
            '����:41136
            strժҪ = objBillDetail.ժҪ
'            If lng����ID <> 0 And intInsure <> 0 Then '90304
                strժҪ = gclsInsure.GetItemInfo(intInsure, lng����ID, objBillDetail.�շ�ϸĿID, strժҪ, 1, , "|1")
                objBillDetail.ժҪ = strժҪ
'            Else
'                objBillDetail.ժҪ = ""
'            End If
            
            '����۸񲿷�=====================================================
            If rsItems.RecordCount > 0 Then rsItems.MoveFirst
            Set objBillDetail.InComes = New BillInComes
            Do While Not rsItems.EOF
                '�������еļ۸��������¼���
                If InStr(",5,6,7,", rsItems!�շ����) > 0 Or (rsItems!�շ���� = "4" And Nvl(rsItems!��������, 0) = 1) Then
                    '----------------------------------------------------------------------------------------------
                    'ʱ��ҩƷ����۸�(�����ɲ�����)
                    dblAllTime = Val(Nvl(!����))     '�������ۼ�����
                    If dblAllTime <> 0 Or Val(Nvl(rsItems!�Ƿ���)) = 1 Then
                        Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                    "��ȡҩƷ��ǰ�ۼ�", CLng(!�շ�ϸĿID), objBillDetail.ִ�в���ID, dblAllTime)
                        If rsPrice.EOF Then
                            '��ȡ�۸�ʧ��
'                            If !�շ���� = "4" Then
'                                MsgBox "��������""" & Nvl(!����) & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
'                            Else
'                                MsgBox "ҩƷ""" & Nvl(!����) & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
'                            End If
                            objBillIncome.��׼���� = 0
                        Else
                            strPrice = Nvl(rsPrice!Price) & "|||"
                            varPrice = Split(strPrice, "|")
                            objBillIncome.��׼���� = Val(varPrice(0))
                            dblʣ������ = Val(varPrice(2))
                            
                            If dblʣ������ <> 0 And Val(Nvl(rsItems!�Ƿ���)) = 1 Then
                                '����δ�ֽ����
'                                If rsItems!�շ���� = "4" Then
'                                    MsgBox "ʱ����������""" & rsItems!���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
'                                Else
'                                    MsgBox "ʱ��ҩƷ""" & rsItems!���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
'                                End If
                                objBillIncome.��׼���� = 0
                            End If
                        End If
                    Else
                        objBillIncome.��׼���� = 0
                    End If
                ElseIf Val(Nvl(rsItems!�Ƿ���)) = 1 Then
                    If Abs(Val(Nvl(!����))) > Abs(Val(Nvl(rsItems!�ּ�))) Or Abs(Val(Nvl(!����))) = 0 Then
                        objBillIncome.��׼���� = Val(Nvl(rsItems!ȱʡ�۸�))
                    Else
                        objBillIncome.��׼���� = Val(Nvl(!����))
                    End If
                Else
                objBillIncome.��׼���� = Val(Nvl(rsItems!�ּ�))
                End If
                                    
                If InStr(",5,6,7,", rsItems!�շ����) > 0 And gblnҩ����λ Then
                    objBillIncome.��׼���� = Format(objBillIncome.��׼���� * Nvl(rsItems!ҩ����װ, 1), gstrFeePrecisionFmt)
                Else
                    objBillIncome.��׼���� = Format(objBillIncome.��׼����, gstrFeePrecisionFmt)
                End If
                objBillIncome.�ּ� = Val(Nvl(rsItems!�ּ�))  '�ּ�ԭ�۶�ҩƷ�������
                objBillIncome.ԭ�� = Val(Nvl(rsItems!ԭ��))
                objBillIncome.������ĿID = Val(Nvl(rsItems!������ID))
                objBillIncome.������Ŀ = Nvl(rsItems!������Ŀ)
                objBillIncome.�վݷ�Ŀ = Nvl(rsItems!�ַ�Ŀ)
                
                'Ӧ�ս��=����*����*����
                objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * objBillDetail.���� * objBillDetail.����
                
                '�������������ü���(����������Ŀ)
                If 0 = 1 And Nvl(rsItems!�շ����) = "F" Then
                    objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� * IIf(Val(Nvl(rsItems!�����շ���)) = 0, 1, Val(Nvl(rsItems!�����շ���)) / 100)
                End If
                
                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If bln�Ӱ�Ӽ� And Val(Nvl(rsItems!�Ӱ�Ӽ�)) = 1 Then
                    dbl�Ӱ�Ӽ��� = Val(Nvl(rsItems!�Ӱ�Ӽ�)) / 100
                    objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� + objBillIncome.Ӧ�ս�� * dbl�Ӱ�Ӽ���
                End If
                objBillIncome.Ӧ�ս�� = Format(objBillIncome.Ӧ�ս��, gstrDec)
                
                '����ʵ�ս��
                If Val(Nvl(rsItems!���ηѱ�)) = 1 Then
                    objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                Else
                    'ʹ��ԭ���Ķ�̬�ѱ�
                    objBillIncome.ʵ�ս�� = ActualMoney(objBillDetail.�ѱ�, Val(Nvl(rsItems!������ID)), objBillIncome.Ӧ�ս��, _
                        objBillDetail.�շ�ϸĿID, objBillDetail.ִ�в���ID, objBillDetail.ԭʼ����, dbl�Ӱ�Ӽ���)
                End If
                
                With objBillIncome
                    '��ȡ��Ŀ������Ϣ,��ҽ�����˲���
                    If int���� <> 0 Then
                        strAdvance = objBillDetail.ժҪ & "||" & objBillDetail.ԭʼ����
                        strInfo = gclsInsure.GetItemInsure(objBill.����ID, objBillDetail.�շ�ϸĿID, .ʵ�ս��, True, int����, strAdvance)
                        If strInfo <> "" Then
                            objBillDetail.������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                            objBillDetail.���մ���ID = Val(Split(strInfo, ";")(1))
                            .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                            objBillDetail.���ձ��� = CStr(Split(strInfo, ";")(3))
                            
                            If UBound(Split(strInfo, ";")) >= 4 Then
                                If CStr(Split(strInfo, ";")(4)) <> "" Then objBillDetail.ժҪ = CStr(Split(strInfo, ";")(4))
                                If UBound(Split(strInfo, ";")) >= 5 Then
                                    If Split(strInfo, ";")(5) <> "" Then objBillDetail.Detail.���� = Split(strInfo, ";")(5)
                                End If
                            End If
                        End If
                    End If
                    objBillDetail.InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
                End With
                '�ж���һ����¼�Ƿ����ڵ�ǰ��
                int��� = !���
                i = i + 1
                rsItems.MoveNext
            Loop
           
            With objBillDetail
                objBill.Pages(1).Details.Add .�ѱ�, .Detail, .�շ�ϸĿID, .���, .��������, .�շ����, .���㵥λ, .��ҩ����, _
                    .����, .����, .���ӱ�־, .ִ�в���ID, .InComes, , .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ, .ԭʼ����, .ԭʼִ�в���ID
                
                '���ù�����
                If objBill.Pages(1).Details(objBill.Pages(1).Details.Count).���ӱ�־ = 8 Then
                    objBill.Pages(1).Details(objBill.Pages(1).Details.Count).������ = True
                End If
            End With
            .MoveNext
        Loop
    End With
     '�����´����������
     With objBill.Pages(1)
        For i = 1 To .Details.Count
            If .Details(i).�������� <> 0 Then
                 .Details(i).�������� = colSerial("_" & .Details(i).��������)(1)
            End If
        Next
    End With
    Set ImportWholeSet = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo Errhand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
Errhand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("������������", , , True)) = 1
    Err = 0: On Error GoTo Errhand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
Errhand:
End Function



Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln������� As Boolean = True, Optional bln���� As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     bln�������     �Ƿ���и������
    '     bln����         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    zlDblIsValid = zlCommFun.DblIsValid(strInput, intMax, bln�������, bln����, hWnd, str��Ŀ)
End Function
Public Function CheckNegative(ByVal lngPatient As Long, ByVal lngPage As Long, ByVal lngItem As Long, ByVal lngExecuteDept As Long, _
    ByVal dblNum As Double, ByVal dblҩ����װ As Double, ByVal strPrivs As String, Optional strStartDate As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡�˱��δ�סԺ���շ���Ŀ�������ϼ��Ƿ��㹻����
    '���:lngNum-����ĸ��������������ҩƷ�����ݲ���ת�����ۼ۵�λ�ٴ������ͬһ����������ͬ����Ŀ��ִ�п��ҵ��ж��У���ʱ����飬����֮ǰ�ټ��
    '       strPrivs-Ȩ�޴�
    '       strStartDate-��ѯ�����ڷ�Χ�Ŀ�ʼʱ�䵽��ǰʱ��
    '����:
    '����:�㹻����Ȩ�޳帺��ʱ,����true,���򷵻�False
    '����:���˺�
    '����:2011-03-18 11:43:09
    '����:36558
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblδ�� As Double, dbl�ѽ� As Double
    Dim strSQL As String, rsTmp As ADODB.Recordset
    'Ŀǰֻ֧���������۲��˵ļ��
    If strStartDate = "" Then strStartDate = "2000-01-01"
    '�ݲ����ø�Ȩ��
'    If InStr(1, strPrivs, ";�������ʲ���鷢����Ŀ;") > 0 Then
'        '���ڸ�������ʱ����鱾��סԺ��������Ŀ����,�д�Ȩ��,����¼�벡��δ�������ķ�����Ŀ���г���,�����鱾��סԺ��������Ŀ�������ܳ���
'        CheckNegative = True: Exit Function
'    End If
    
    '��¼���� In(2,3)ȡ���������ϵ����:  :28029
    On Error GoTo errH
    CheckNegative = True
    
    strSQL = "" & _
            "   Select Nvl(Sum(Nvl(����, 1) * ����),0) As ����," & vbNewLine & _
            "           Sum(decode(����ID,NULL,0,1)* Nvl( ����,1)* ����) as ��������  " & _
            "   From ������ü�¼" & vbNewLine & _
            "   Where  ��¼���� =2 and ���ʷ��� = 1 And �۸񸸺� Is Null  And ��¼״̬<>0  " & _
            "               And ����id = [1] " & vbNewLine & _
            "               And �շ�ϸĿid+0 = [3] And ִ�в���id+0 = [4] And �Ǽ�ʱ��+0>=[5]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngPatient, lngPage, lngItem, lngExecuteDept, CDate(strStartDate))
    
    If Not rsTmp.EOF Then
        If RoundEx(Abs(dblNum), 8) > RoundEx(Val(Nvl(rsTmp!����)), 8) Then
                MsgBox "�����������ڸò����ڵ�ǰִ�п��ҵļ�������" & FormatEx(rsTmp!���� / IIf(gblnҩ����λ, dblҩ����װ, 1), 5) & "��", vbInformation, gstrSysName
                CheckNegative = False: Exit Function
        End If
        
        '�ݲ���
'        Select Case gbytBillOpt '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
'        Case 0  '����
'        Case 1   '����
'            dblδ�� = RoundEx((Val(Nvl(rstmp!����)) - Val(Nvl(rstmp!��������))) / IIf(gblnҩ����λ, dblҩ����װ, 1), 8)
'            dbl�ѽ� = RoundEx(Val(Nvl(rstmp!��������)) / IIf(gblnҩ����λ, dblҩ����װ, 1), 8)
'            If RoundEx(Abs(dblNum), 8) > RoundEx(dblδ��, 8) Then
'                If MsgBox("��������(" & FormatEx(RoundEx(Abs(dblNum) / IIf(gblnҩ����λ, dblҩ����װ, 1), 8), 5) & _
'                        ") �а������Ѿ����ʲ���(δ��:" & FormatEx(dblδ��, 5) & "; �ѽ�:" & FormatEx(dbl�ѽ�, 5) & ") ��" & vbCrLf & _
'                    " �Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                    CheckNegative = False: Exit Function
'                End If
'            End If
'        Case 2   '��ֹ
'                dblδ�� = RoundEx((Val(Nvl(rstmp!����)) - Val(Nvl(rstmp!��������))) / IIf(gblnҩ����λ, dblҩ����װ, 1), 8)
'                dbl�ѽ� = RoundEx(Val(Nvl(rstmp!��������)) / IIf(gblnҩ����λ, dblҩ����װ, 1), 8)
'                If RoundEx(Abs(dblNum), 8) > RoundEx(dblδ��, 8) Then
'                    Call MsgBox("��������(" & FormatEx(RoundEx(Abs(dblNum) / IIf(gblnҩ����λ, dblҩ����װ, 1), 8), 5) & _
'                        ") �а������Ѿ����ʲ���(δ��:" & FormatEx(dblδ��, 5) & "; �ѽ�:" & FormatEx(dbl�ѽ�, 5) & ") ,���ܼ�����" & vbCrLf & _
'                    "", vbInformation + vbOKOnly, gstrSysName)
'                    CheckNegative = False: Exit Function
'                End If
'        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    CheckNegative = False
    Call SaveErrLog
End Function

Public Function zl_GetInvoicePrintFormat(ByVal lngModule As Long, _
    Optional strʹ����� As String = "", Optional ByRef intPrintFormatOld As Integer, _
    Optional ByRef blnPatiPrintBill As Boolean = False, Optional ByVal blnDelFeePrintBill As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ӡ��ʽ
    '���:blnPatiPrintBill-��ȡ�����˲���Ʊ�ݸ�ʽ
    '   blnDelFeePrintBill - ��ȡ�˷ѷ�Ʊ��ʽ(91998)
    '����:intPrintFormatOld-������Ʊ�ݴ�ӡ��ʽ(Ʊ�ŷ������Ϊ����ʵ�ʴ�ӡ��ƱƱ�ŷ�ʽ����ӡ�ĸ�ʽ)(56963)
    '����:��ӡ��ʽ(���)
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
    Dim intFormat As Integer, intFormat1 As Integer
    Dim intNewPrintFormat As Integer
    
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    If blnDelFeePrintBill Then
        strShareTypeFormat = Trim(zlDatabase.GetPara("�˷ѷ�Ʊ��ʽ", glngSys, lngModule, ""))
    ElseIf blnPatiPrintBill Then
        strShareTypeFormat = Trim(zlDatabase.GetPara("�����˲���Ʊ��ʽ", glngSys, lngModule, ""))
    Else
        strShareTypeFormat = Trim(zlDatabase.GetPara("�շѷ�Ʊ��ʽ", glngSys, lngModule, ""))
    End If
    '��ʽ:ʹ�����1,��ʽ1|ʹ�����2,��ʽ2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        intFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intFormat1 = intFormat
        If Trim(varTemp(0)) = strʹ����� And intFormat <> 0 Then
            intNewPrintFormat = intFormat: GoTo GetOLdFormat:
        End If
    Next
    intNewPrintFormat = intFormat1
    '��ȡ�ɷ�Ʊ��ʽ(56963)
GetOLdFormat:
    If gTy_Module_Para.bytƱ�ݷ������ = 0 Or blnDelFeePrintBill Then
        '����ʵ�ʴ�ӡ����Ʊ��ʱ,����ԭ����ʽ��ӡ���ô���
        intPrintFormatOld = intNewPrintFormat
        zl_GetInvoicePrintFormat = intNewPrintFormat
        Exit Function
    End If
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    strShareTypeFormat = Trim(zlDatabase.GetPara("�շѷ�Ʊ��ʽ(��)", glngSys, lngModule, ""))
    '��ʽ:ʹ�����1,��ʽ1|ʹ�����2,��ʽ2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        intFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intFormat1 = intFormat
        If Trim(varTemp(0)) = strʹ����� And intFormat <> 0 Then
            intPrintFormatOld = intFormat
            zl_GetInvoicePrintFormat = intNewPrintFormat
            Exit Function
        End If
    Next
    intPrintFormatOld = intFormat1
    zl_GetInvoicePrintFormat = intNewPrintFormat
End Function

Public Function zl_GetInvoicePrintMode(ByVal lngModule As Long, _
    Optional strʹ����� As String = "", Optional ByVal blnDelFeePrintBill As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ӡ��ʽ
    '����:int��ӡ��ʽ-��ӡ��ʽ()
    '   blnDelFeePrintBill - ��ȡ�˷ѷ�Ʊ��ʽ(91998)
    '����:0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
     Dim intPrintMode As Long, intPrintMode1 As Long
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    If blnDelFeePrintBill Then
        strShareTypeFormat = Trim(zlDatabase.GetPara("�˷ѷ�Ʊ��ӡ��ʽ", glngSys, lngModule, ""))
    Else
        strShareTypeFormat = Trim(zlDatabase.GetPara("�շѷ�Ʊ��ӡ��ʽ", glngSys, lngModule, ""))
    End If
    '��ʽ:ʹ�����1,��ӡ��ʽ1|ʹ�����2,��ӡ��ʽ2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,", ",")
        intPrintMode = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
        If Trim(varTemp(0)) = strʹ����� Then
            zl_GetInvoicePrintMode = intPrintMode: Exit Function
        End If
    Next
    zl_GetInvoicePrintMode = intPrintMode1
End Function

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

Public Function GetChargeBalance(ByVal strNos As String, _
    Optional ByVal lng������� As Long = 0, Optional lng����ID As Long, _
    Optional blnHistory As Boolean = False, _
    Optional strDelTime As String, Optional intSign As Integer = 1) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѵ���ؽ��㷽ʽ
    '���:�ж�˳��:strNos-->lng�������-->lng����ID
    '       strDelTime-�˷�ʱ���
    '       intSign:ͳ�Ʒ�:1��-1;
    '����:�շ���صĽ��㷽ʽ(����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������)
    '       �ֶ�:A.����ID,A.NO,A.����,A.��������,A.���㷽ʽ,A.������,
    '               A.�����ID,A.����,A.�Ƿ�ȫ��,A.�Ƿ�����,A.�������,A.����,A.������ˮ��,
    '               A.����˵��,A.�������,A.У�Ա�־
    '����:���˺�
    '����:2011-08-28 21:38:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable  As String, strFeeTab As String, strPreTab As String
    Dim dtDelDate As Date
    strFeeTab = IIf(blnHistory, "H������ü�¼", "������ü�¼")
    strPreTab = IIf(blnHistory, "H����Ԥ����¼", "����Ԥ����¼")
    
    On Error GoTo errHandle
    
   dtDelDate = CDate("1991-01-01")
    If strDelTime <> "" Then dtDelDate = CDate(strDelTime)
    If strNos <> "" Then
        strTable = "Select distinct M.NO,M.����ID From " & strFeeTab & " M," & strPreTab & " J,Table( f_Str2list([2])) Q Where M.����ID=J.����ID and M.NO=Q.Column_value And M.��¼����=1"
        strTable = strTable & "" & _
        IIf(strDelTime = "", " And M.��¼״̬ In(1,3)", " And M.��¼״̬=2") & _
        IIf(strDelTime <> "", " And M.�Ǽ�ʱ��=[3]", "")
    Else
        strTable = "Select distinct M.NO,M.����ID From " & strFeeTab & " M," & strPreTab & " J,Table( f_Num2list([1])) Q Where M.����ID=J.����ID and " & IIf(lng������� = 0, "J.����ID", "J.�������") & "=Q.Column_value And M.��¼����=1"
    End If
    
    strSQL = "" & _
    "   Select   W.NO,A.����ID, " & _
    "       Case  When Mod(A.��¼����,10)=1 then 1  " & _
    "                 When nvl(A.�����ID,0)>0 then   3 " & _
    "                 When nvl(A.���㿨���,0)>0 then 4 " & _
    "                 When  nvl(B.����,0)=3 or nvl(B.����,0)=4 then 2 " & _
    "                 When  C.���㷽ʽ Is not null   then 5 else 0 End  as ����," & _
    "       nvl(b.����,1) as ��������,B.Ӧ����, " & _
    "       Decode(Mod(A.��¼����,10),1,'Ԥ���',A.���㷽ʽ) as ���㷽ʽ, " & _
    "       " & intSign & "*nvl(A.��Ԥ��,0) as ��Ԥ��," & _
    "       Nvl(I.�Ƿ��˿��鿨,0) as �Ƿ��˿��鿨," & _
    "       nvl(nvl(A.�����ID,A.���㿨���),0) as �����ID, nvl(I.����,L.����) as ����, " & _
    "       nvl(nvl(I.�Ƿ�ȫ��,L.�Ƿ�ȫ��) ,0) as �Ƿ�ȫ��,nvl(nvl(I.�Ƿ�����,L.�Ƿ�����),0) as �Ƿ�����," & _
    "       A.�������,A.ժҪ,nvl(A.����,A.��λ�ʺ�) as ����,nvl(A.������ˮ��,A.�������) as ������ˮ��,A.����˵��,A.�������,C.ҽԺ����,nvl(A.У�Ա�־,0) У�Ա�־" & _
    "   From " & strPreTab & " A,  (" & strTable & ") W, " & _
    "           ҽ�ƿ���� I,���ѿ����Ŀ¼ L ,���㷽ʽ B,  " & _
    "           (Select ���㷽ʽ ,ҽԺ���� From һ��ͨĿ¼ Where ����=1 ) C " & _
    "   where A.����ID=W.����ID   " & _
    "               And A.�����ID=I.ID(+) and A.���㿨���=L.���(+) " & _
    "               And A.���㷽ʽ=B.����(+)   " & _
    "               And A.���㷽ʽ=C.���㷽ʽ(+)"
    strSQL = "" & _
    "   Select /*+ Rule*/  A.����ID,A.NO,A.����,A.��������,A.Ӧ����,A.���㷽ʽ,nvl(sum(A.��Ԥ��),0) as ������, " & _
    "           A.�Ƿ��˿��鿨,A.�����ID,A.����,A.�Ƿ�ȫ��,A.�Ƿ�����,A.�������,Max(A.ժҪ) as ժҪ,A.����,A.������ˮ��," & _
    "           A.����˵��, A.�������,A.У�Ա�־ " & _
    "   From (" & strSQL & ") A" & _
    "   Group by  A.����ID,A.NO,A.����,A.Ӧ����,A.���㷽ʽ,A.�����ID ,A.����,A.�Ƿ�ȫ��,A.�Ƿ�����,A.�������, " & _
    "            A.�Ƿ��˿��鿨,A.����,A.������ˮ��,A.����˵��,A.������� ,A.��������,A.ҽԺ����, A.У�Ա�־" & _
    "   Having nvl(sum(A.��Ԥ��),0)<>0"
    '�쳣���ݵĽ��㷽ʽ(����Ԥ����)
    Set GetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շѽ��㷽ʽ", IIf(lng������� = 0, lng����ID, lng�������), Replace(strNos, "'", ""), dtDelDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub
Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

Public Function isCheckExiseSingularity(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ���շ��Ƿ�����쳣�����ϵ���
    '���:strNo-���ݺ�
    '����:
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2012-03-01 12:02:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  1  " & _
    "   From ����Ԥ����¼ B,������ü�¼ A,����Ԥ����¼ C  " & _
    "   Where  B.����ID=A.����ID and B.�������=C.�������  " & _
    "               And C.NO||''<>[1] And C.��¼״̬<>1  And A.NO=[1] And Rownum=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ������һ���շ��Ƿ�����Ѿ����ϵĵ���", strNo)
    isCheckExiseSingularity = rsTemp.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlExeCuteBillNoSplit(ByVal blnģ����� As Boolean, ByVal int�������� As Integer, _
    ByVal lng����ID As Long, ByVal strNos As String, ByVal lng����ID As Long, _
    ByVal str��ʼ��Ʊ�� As String, ByVal datFeeDate As Date, Optional ByVal bytƱ�� As Byte = 1, _
    Optional ByRef str��Ʊ�� As String, Optional int��Ʊ���� As Integer, _
    Optional ByVal lngNext����ID As Long, Optional ByVal strNext��ʼ��Ʊ�� As String, Optional lng��ӡID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�з���Ʊ�ŵĹ���
    '���::1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
    '       strNos-���ݺ�(�ö��ŷֿ�)
    '       str��Ʊ��-��Ʊ��(����ö��ŷ���):(3-�ش�;4-�����˷�ʱ,����)
    '       lng��ӡID- lng��ӡID<>0ʱ����ʾ������ʱ����ʱƱ�ݴ�ӡ���ݡ�����Ӧ��NO������Ʊ��
    '����:str��Ʊ��-���δ�ӡ�ķ�Ʊ��(����ö��ŷ���)
    '       int��Ʊ����-��Ʊ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-03-27 10:10:41
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCmd As ADODB.Command, i As Long
    Dim objPara(0 To 13) As ADODB.Parameter
    Dim varParaName As Variant, varParaValue As Variant, varTemp As Variant
    Dim strParaName As String, varValue As Variant
    Dim intDataType As ADODB.DataTypeEnum
    Dim Paradt As ParameterDirectionEnum
    Dim intMaxSize As Integer
    Dim strLog As String
   
    On Error GoTo errHandle
    Set objCmd = New ADODB.Command
    '  Zl_Invoice_Autoallot(
    '  ��������_In   Number,
    '  ģ�����_In   Number,
    '  Ʊ��_In       Ʊ��ʹ����ϸ.Ʊ��%Type,
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    '  ����id_In     ������ü�¼.����id%Type,
    '  Nos_In        Varchar2,
    '  ��ʼ��Ʊ��_In ������ü�¼.ʵ��Ʊ��%Type,
    '  ʹ����_In     Ʊ��ʹ����ϸ.ʹ����%Type,
    '  ʹ��ʱ��_In   Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
    '  Next����id_In Ʊ��ʹ����ϸ.����id%Type := 0,
    '  NextƱ�ݺ�_In Ʊ��ʹ����ϸ.����%Type := Null,
    '  ��Ʊ��_In     In Out Varchar2,
    '  ��Ʊ����_In   Out Number
   varParaName = Split("��������_In|N|IN,ģ�����_In|N|IN,Ʊ��_In|N|IN,����id_In|N|IN," & _
                    "����id_In|N|IN,Nos_In|C|IN,��ʼ��Ʊ��_In|C|IN,ʹ����_In|C|IN," & _
                    "ʹ��ʱ��_In|D|IN,Next����id_In|N|IN,NextƱ�ݺ�_In|C|IN,��Ʊ��_In|C|INOUT,��Ʊ����_In|N|OUT,��ӡid_In|N|IN", ",")
                    
                    
   varParaValue = Split(int�������� & ";" & IIf(blnģ�����, 1, 0) & ";1;" & lng����ID & ";" & _
                    lng����ID & ";" & strNos & ";" & str��ʼ��Ʊ�� & ";" & UserInfo.���� & ";" & _
                    Format(datFeeDate, "YYYY-MM-DD HH:MM:SS") & ";" & lngNext����ID & ";" & strNext��ʼ��Ʊ�� & ";" & str��Ʊ�� & ";" & int��Ʊ���� & ";" & lng��ӡID, ";")
  
                               
    For i = 0 To UBound(varParaName)
        '������|��������|�����
        varTemp = Split(varParaName(i) & "||||", "|")
        strParaName = varTemp(0)    '������
        Select Case Trim(varTemp(1))
        Case "C" '�ַ�
             varValue = Replace(CStr(varParaValue(i)), "'", "")
            ' ����ð󶨱���ת��ΪRAWʱ����2000���ַ�Ҫ��adLongVarChar
            intMaxSize = LenB(StrConv(varValue, vbFromUnicode))
            If intMaxSize <= 2000 Then
                intMaxSize = IIf(intMaxSize <= 200, 200, 2000)
                intDataType = adVarChar
            Else
                If intMaxSize < 4000 Then intMaxSize = 4000
                intDataType = adLongVarChar
            End If
            strLog = strLog & ",'" & varValue & "'"
        Case "D"
             intDataType = adDBTimeStamp
             varValue = CDate(varParaValue(i))
             strLog = strLog & ",to_date('" & varParaValue(i) & "','yyyy-mm-dd hh24:mi:ss') "
        Case Else
             intDataType = adVarNumeric
             varValue = CLng(varParaValue(i))
             strLog = strLog & "," & varValue
             intMaxSize = 30
        End Select
        Select Case Trim(varTemp(2))
        Case "IN" '�ַ�
             Paradt = adParamInput
        Case "INOUT"
             Paradt = adParamInputOutput
        Case Else
             Paradt = adParamOutput
        End Select
        
        If varTemp(1) = "D" Then
            Set objPara(i) = objCmd.CreateParameter(strParaName, _
              intDataType, Paradt)
        Else
            Set objPara(i) = objCmd.CreateParameter(strParaName, _
              intDataType, Paradt, intMaxSize)
        End If
        If Paradt <> adParamOutput Then
          objPara(i).Value = varValue
        End If
        objCmd.Parameters.Append objPara(i)
    Next
    If strLog <> "" Then strLog = Mid(strLog, 2)
    strLog = "Zl_Invoice_Autoallot(" & strLog & ")"
    objCmd.CommandText = "Zl_Invoice_Autoallot"
    objCmd.CommandType = adCmdStoredProc
    Set objCmd.ActiveConnection = gcnOracle
    Call SQLTest(App.ProductName, "Ʊ�ŷ���", strLog)
    objCmd.Execute
    Call SQLTest
    str��Ʊ�� = Nvl(objPara(UBound(varParaName) - 2).Value)
    int��Ʊ���� = Val(Nvl(objPara(UBound(varParaName) - 1).Value))
    zlExeCuteBillNoSplit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInvoiceFromNOs(ByVal strInvioceNos As String, _
    Optional bln����ʷ��ռ� As Boolean = False, _
    Optional ByRef str������� As String = "", Optional cllInvoiceNoInfor As Collection) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ�Ʊ��,��ȡ��Ӧ�ĵ��ݺ�
    '���:strInvioceNos-��Ʊ��,��������ö��ŷ���:A0001,A0002
    '����:str�������-�������ŵ��ݵĶ���������(����÷�Ʊ�漰����շѵ�)
    '       cllInvoiceNoInfor-array(No,���)
    '����:�ɹ����ش���ķ�Ʊ���漰�ĵ��ݺ�
    '����:���˺�
    '����:2013-04-12 15:59:32
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strNos As String
    Dim strSQL1 As String, strSQL As String
    
    On Error GoTo errHandle
    
    Set cllInvoiceNoInfor = New Collection
    str������� = ""
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
        strSQL = "" & _
        "   Select  /*+ RULE */  A.NO,Max(A.���) as ���,Max(C.�������) as �������" & _
        "   From Ʊ�ݴ�ӡ��ϸ A,������ü�¼ B,����Ԥ����¼ C,Table( f_Str2list([1])) J" & _
        "   Where A.Ʊ��=J.Column_Value and Ʊ��=1 and A.�Ƿ����<>1" & _
        "           And A.No=B.NO And B.��¼����=1  And nvl(B.��¼״̬,0)<>2 And B.����ID=C.����ID" & _
        "   Group by A.NO"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ӧ���ݵķ�Ʊ��", strInvioceNos)
        strNos = ""
        With rsTemp
            Do While Not .EOF
                strNos = strNos & "," & Nvl(!NO)
                str������� = str������� & "," & Val(Nvl(!�������))
                cllInvoiceNoInfor.Add Array(Nvl(!NO), Nvl(!���))
                .MoveNext
            Loop
            If str������� <> "" Then str������� = Mid(str�������, 2)
            If strNos <> "" Then
                zlInvoiceFromNOs = Mid(strNos, 2)
                Exit Function
            End If
        End With
    End If
    strSQL = "" & _
    "   Select NO  " & _
    "   From Ʊ�ݴ�ӡ���� A, " & _
    "           (   Select Max(M.��ӡID) as ��ӡID " & _
    "               From  Ʊ��ʹ����ϸ M ,Table( f_Str2list([1])) J  " & _
    "               Where M.Ʊ��=1 And M.����=1 And M.����=J.Column_Value  " & _
    "               Group by M.����" & _
    "               )  Q" & _
    "   Where A.��������=1  And ID=Q.��ӡID "
    strSQL1 = Replace(Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����"), "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
    
    strSQL = "" & _
    "   Select  /*+ RULE */   Distinct NO " & _
    "   From (" & strSQL & " Union ALL " & strSQL1 & ") " & _
    "   Order by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ӧ���ݵķ�Ʊ��", strInvioceNos)
    
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(!NO)
            .MoveNext
        Loop
        If strNos <> "" Then
            zlInvoiceFromNOs = Mid(strNos, 2)
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetReclaimInvoice(ByVal strNoInfor As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ҫ���յ�Ʊ��
    '���:strNoInfor-���ݺ���Ϣ(���ݺ�1:���1(1..n);���ݺ�2:���2(1..n)
    '       str���-�����е����,����ö��ŷ���
    '����:��ȡ�ɹ�,���ر�����Ҫ���յ�Ʊ��
    '����:���˺�
    '����:2013-03-27 18:27:13
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoices As String, rsTemp As ADODB.Recordset, strSQL As String
    Dim strNos As String, varData As Variant, varTemp As Variant
    Dim i As Long, blnFind As Boolean, j As Long
    Dim str������� As String, rsInvoice As ADODB.Recordset
    Dim cllNos As Collection
    Dim strNo As String
    Dim varValue() As Variant
    
    On Error GoTo errHandle
    
    '����ʵ�ʴ�ӡ����Ʊ��ʱ,�����ؾ����Ʊ��
    If gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Function

    
    Set rsInvoice = New ADODB.Recordset
    rsInvoice.Fields.Append "��Ʊ��", adVarChar, 50, adFldIsNullable
    rsInvoice.CursorLocation = adUseClient
    rsInvoice.LockType = adLockOptimistic
    rsInvoice.CursorType = adOpenStatic
    rsInvoice.Open

    If strNoInfor = "" Then Exit Function
    strNoInfor = Replace(strNoInfor, "'", "")
    varData = Split(strNoInfor, ";")
  
    Set cllNos = New Collection
    strNos = ""
    For i = 0 To UBound(varData)
        strNo = Split(varData(i) & ":", ":")(0)
        If Len(strNos & "," & strNo) > 4000 Then
            strNos = Mid(strNos, 2)
            cllNos.Add strNos
            strNos = ""
        End If
        strNos = strNos & "," & strNo
    Next
    If strNos <> "" Then
        strNos = Mid(strNos, 2)
        cllNos.Add strNos
    End If
    
    If cllNos.Count <= 1 Then
        strSQL = "" & _
        "   Select  /*+ RULE */  A.NO, A.Ʊ��,A.���,A.����Ʊ�����" & _
        "   From Ʊ�ݴ�ӡ��ϸ A,Table( f_Str2list([1])) J" & _
        "   Where A.NO=Column_Value and Ʊ��=1 and �Ƿ����<>1  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ӧ���ݵķ�Ʊ��", strNos)
    Else
        If zlFromCollectBulidSQL(cllNos, strSQL, varValue) = False Then Exit Function
        strSQL = "With ������Ϣ as (" & strSQL & ")" & vbCrLf
        strSQL = strSQL & " Select * From ������Ϣ "
        
        strSQL = "" & _
        "   Select  A.NO, A.Ʊ��,A.���,A.����Ʊ�����" & _
        "   From Ʊ�ݴ�ӡ��ϸ A,(" & strSQL & ") J" & _
        "   Where A.NO=J.NO and Ʊ��=1 and �Ƿ����<>1  "
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "��ȡ��Ӧ���ݵķ�Ʊ��", varValue)
    End If


    If rsTemp.RecordCount = 0 Then Exit Function
    With rsTemp
        Do While Not .EOF
            If InStr(strInvoices & ",", "," & Nvl(!Ʊ��)) = 0 Then
                blnFind = False
                For i = 0 To UBound(varData)
                    varTemp = Split(varData(i) & ":", ":")
                    If varTemp(0) = Nvl(!NO) Then
                         If varTemp(1) = "" Then
                            If Val(Nvl(!����Ʊ�����)) <> 0 And InStr(str������� & ",", "," & Val(Nvl(!����Ʊ�����)) & ",") = 0 Then
                                str������� = str������� & "," & Val(Nvl(!����Ʊ�����))
                            End If
                            blnFind = True: Exit For
                          End If
                         varTemp = Split(varTemp(1), ",")
                         For j = 0 To UBound(varTemp)
                             If InStr("," & Nvl(!���) & ",", "," & varTemp(j) & ",") > 0 Then
                                    If Val(Nvl(!����Ʊ�����)) <> 0 And InStr(str������� & ",", "," & Val(Nvl(!����Ʊ�����)) & ",") = 0 Then
                                        str������� = str������� & "," & Val(Nvl(!����Ʊ�����))
                                    End If
                                    blnFind = True: Exit For
                             End If
                         Next
                    End If
                Next
                If blnFind Then
                    rsInvoice.AddNew
                    rsInvoice!��Ʊ�� = Nvl(!Ʊ��)
                    rsInvoice.Update
                    strInvoices = strInvoices & "," & Nvl(!Ʊ��)
                End If
            End If
            .MoveNext
        Loop
        If strInvoices <> "" Then strInvoices = Mid(strInvoices, 2)
    End With
    If str������� <> "" Then
            '������Ʊ��ҲҪ��ʾ����
            str������� = Mid(str�������, 2)
            gstrSQL = "" & _
             "   Select  /*+ RULE */ distinct  A.Ʊ��" & _
             "   From Ʊ�ݴ�ӡ��ϸ A,Table( f_Num2list([1])) J" & _
             "   Where A.����Ʊ�����=J.Column_Value And A.Ʊ��=1 and A.�Ƿ����<>1 " & _
             "              And A.Ʊ�� Not In(Select Column_Value From Table( f_Str2list([2])) )"
             Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ӧ���ݵķ�Ʊ��", str�������, strInvoices)
             With rsTemp
                Do While Not .EOF
                    rsInvoice.AddNew
                    rsInvoice!��Ʊ�� = Nvl(!Ʊ��)
                    rsInvoice.Update
                    strInvoices = strInvoices & "," & Nvl(!Ʊ��)
                    .MoveNext
                Loop
             End With
     End If
     '����
     rsInvoice.Sort = "��Ʊ��"
     With rsInvoice
        If .RecordCount <> 0 Then .MoveFirst
        strInvoices = ""
        Do While Not .EOF
            strInvoices = strInvoices & "," & Nvl(!��Ʊ��)
            .MoveNext
        Loop
        .Close
     End With
     If strInvoices <> "" Then strInvoices = Mid(strInvoices, 2)
     Set rsInvoice = Nothing
    zlGetReclaimInvoice = strInvoices
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Public Function zlGetFromNoTOInvoice(ByVal strNos As String) As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����������Ӧ�ķ�Ʊ��
    '���:strNos-���ݺ�,����Ϊ���,���ʱ�ö��ŷ���
    '����:�ɹ�����ָ����������Ӧ�ķ�Ʊ�ŵļ�¼��
    '����:���˺�
    '����:2013-05-06 16:17:50
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    'If strNos <> "" Then strNos = Replace(Mid(strNos, 2), "'", "")
    strSQL = "" & _
    "   With C_���� as (Select Column_Value as NO From Table( f_Str2list([1])) )" & _
    "   Select  /*+ RULE */  A.NO, A.Ʊ��,A.���,A.����Ʊ�����" & _
    "   From Ʊ�ݴ�ӡ��ϸ A,C_���� J" & _
    "   Where A.NO=J.NO and Ʊ��=1 and �Ƿ����<>1  " & _
    "   Union ALL " & _
    "   Select A.NO, A.Ʊ��,A.���,A.����Ʊ�����" & _
    "   From Ʊ�ݴ�ӡ��ϸ A, " & _
    "               (Select ����Ʊ����� From Ʊ�ݴ�ӡ��ϸ A,C_���� M  " & _
    "                Where A.NO=M.NO and A.Ʊ��=1 and A.�Ƿ����<>1 ) J" & _
    "   Where A.����Ʊ�����=J.����Ʊ����� And A.Ʊ��=1 and A.�Ƿ����<>1 "
   strSQL = "" & _
   "    Select /*+ RULE */ distinct  A.NO, A.Ʊ��,A.���,A.����Ʊ�����  " & _
   "    From (" & strSQL & ") A"
    Set zlGetFromNoTOInvoice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ӧ���ݵķ�Ʊ��", strNos)
End Function
Public Function zlCheckDrugIsPutDrug(ByVal strNos As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ҩƷ��ҩ���м��
    '���:strNos-���ݺ�,����ö��ŷָ�
    '����:û�а�ҩ���ҩ����ѡ��Ϊtrueʱ,����true,���򷵻�False
    '����:���˺�
    '����:2013-04-16 12:44:49
    '����:47400
    ' ����:�˷�ʱ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    If gTy_Module_Para.bytҩƷ��ҩ�˷ѷ�ʽ = 0 Then zlCheckDrugIsPutDrug = True: Exit Function
    
    strSQL = "Select  /*+ rule */  1 From δ��ҩƷ��¼ A,Table(f_str2List([1])) J Where A.NO=J.Column_Value And A.���� in (8,24) And A.��ҩ�� Is NOT NULL And Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ҩƷ�Ƿ��ҩ!", strNos)
    If rsTemp.EOF Then zlCheckDrugIsPutDrug = True: Exit Function
    If gTy_Module_Para.bytҩƷ��ҩ�˷ѷ�ʽ = 1 Then
        '��ֹ�˷�
        MsgBox "���˷ѵ������Ѿ����ڰ�ҩ�ĵ���,��������Ѿ���ҩ�ĵ��ݽ����˷�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '��ʾ
    If MsgBox("���˷ѵ������Ѿ����ڰ�ҩ�ĵ���,�Ƿ���Ѿ���ҩ�ĵ��ݽ����˷�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    zlCheckDrugIsPutDrug = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckIsExcuteData(ByVal strNos As String, ByVal byt��¼���� As Byte) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ��ִ�мƼ��Ƿ�������
    '���:strNOs-�շѵ���,����ö��ŷ���
    '����:�����ݷ���true,���򷵻�False
    '����:���˺�
    '����:2013-04-25 13:57:56
    '����:60735
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select /*+ rule */  1 " & _
    "   From ������ü�¼ A, ҽ��ִ�мƼ� B, Table(f_Str2list([2])) J " & _
    "   Where a.ҽ����� = b.ҽ��id And mod(a.��¼����,10) = [1] And a.No = j.Column_Value And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ҽ��ִ�мƼ��Ƿ�������", byt��¼����, strNos)
    zlCheckIsExcuteData = Not rsTemp.EOF
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlStringSort(ByVal strSort As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ַ�������
    '���:strSort-�ö��ŷ���
    '����:
    '����:�����������ָ���
    '����:���˺�
    '����:2013-05-07 18:23:34
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    Dim strTemp As String, intCount As Integer
    Dim i As Long, j As Long
    varTemp = Split(strSort, ",")
    intCount = UBound(varTemp)
    For i = 0 To intCount
        For j = i + 1 To intCount
            If varTemp(i) > varTemp(j) Then
                strTemp = varTemp(i)
                varTemp(i) = varTemp(j)
                varTemp(j) = strTemp
            End If
        Next
    Next
    strTemp = ""
    For i = 0 To intCount
        strTemp = strTemp & "," & varTemp(i)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    zlStringSort = strTemp
End Function
 
Public Sub zlDebugWriteFile(ByVal strLogText As String)
    Dim objLogFile As FileSystemObject
    Dim objLogText As TextStream
    Dim strTmp As String
    If OS.IsDesinMode = False Then Exit Sub
    
    Set objLogFile = New FileSystemObject
    On Local Error Resume Next
    Set objLogText = objLogFile.OpenTextFile(gstrDBUser & "_" & Format(Date, "yyyyMMdd") & ".log", ForAppending, True, TristateFalse)
    On Local Error GoTo 0
    If Not objLogText Is Nothing Then
        strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
        objLogText.WriteLine strTmp
        objLogText.WriteLine strLogText
    End If
    objLogText.Close
    Set objLogText = Nothing
    Set objLogFile = Nothing
End Sub

Public Function zlBillErrIsCanDel(ByVal lng����ID As Long, ByRef bytDelErrType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ָ�����쳣�շѵ����Ƿ�������
    '����:bytErrType-�˷Ѵ�������:1-�쳣�շѵ�����;2-�������˷�
    '����:���˺�
    '����:2012-03-01 01:04:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnDel As Boolean, strNo As String
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select NO, Max(��¼״̬) as ��¼״̬" & _
    "   From ������ü�¼  " & _
    "   Where ����ID=[1] And nvl(����״̬,0)=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�쳣�շѵ���", lng����ID)
    blnDel = False
    If Not rsTemp.EOF Then
        blnDel = Val(Nvl(rsTemp!��¼״̬)) = 2
        strNo = Nvl(rsTemp!NO)
        bytDelErrType = 2
        If blnDel Then
            strSQL = "" & _
             "   Select  1 " & _
             "   From ������ü�¼  " & _
             "   Where NO=[1] And ��¼����=1 And ��¼״̬ in (1,3) " & _
             "     And nvl(����״̬,0)=1 "
             
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�쳣�շѵ���", strNo)
            If Not rsTemp.EOF Then bytDelErrType = 1
        End If
    End If
    zlBillErrIsCanDel = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str�������� As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�������,���ò�ͬ�������͵���ʾ��ɫ
    '���:objPatiControl-���˿ؼ�(�ı���,��ǩ)
    '    str��������-��������
    '    lngDefaultColor-ȱʡ���˵���ʾ��ɫ
    '����:True-������ɫ�ɹ���False-ʧ��
    '����:���ϴ�
    '����:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str�������� <> "" Then
        lngColor = zlDatabase.GetPatiColor(str��������)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function
Public Function zlFormatNum(ByVal dblMoney As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ʽ����(����:.03 ��ʽΪ0.03,123��ʽΪ123)
    '���:dblMoney-��ʽ�����
    '����:���ظ�ʽ����(����:.03 ��ʽΪ0.03,123��ʽΪ123)
    '����:���˺�
    '����:2014-07-30 15:29:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    Dim strMoney As String
'    If dblMoney = 0 Then Exit Function
    strTemp = Format(dblMoney, "###0.00######;-###0.00######;;")
    If strTemp = "" Then Exit Function
    strMoney = strTemp
    For i = Len(strTemp) To 1 Step -1
        If Val(Mid(strTemp, i, 1)) <> 0 Or Mid(strTemp, i, 1) = "." Then Exit For
        strMoney = Mid(strTemp, 1, i - 1)
    Next
    If Right(strMoney, 1) = "." Then strMoney = Mid(strMoney, 1, Len(strMoney) - 1)
    zlFormatNum = strMoney
End Function

Public Function GetMedicareStr(colBalance As Collection, Optional ByVal intPage As Integer, _
    Optional ByVal intBeforePage As Integer) As String
'���ܣ����ر��ս��㷽ʽ��,"���㷽ʽ|���||...."
'������intPage=�Ƿ�ָ������,����Ϊ���е���
'      intBeforePage=����õ��ݼ���ǰ�ĵ���
'˵�����ú�����colBalanceΪ׼����,����ҽ�������շ�Ҳ��
    Dim i As Integer, p As Integer, strTmp As String
    Dim varData As Variant, curMoney As Currency
    Dim rsTemp As New ADODB.Recordset, strBalance As String, varBalance As Variant
    
    Err = 0: On Error GoTo Errhand:
    rsTemp.Fields.Append "���㷽ʽ", adVarChar, 20, adFldIsNullable
    rsTemp.Fields.Append "���", adCurrency, , adFldIsNullable
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, IIf(intBeforePage = 0, colBalance.Count, intBeforePage), intPage)
        For i = 0 To UBound(colBalance(p))
            '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
            varData = Split(colBalance(p)(i), ";")
            If varData(0) <> "" Then
                If InStr(strBalance & ";", ";" & varData(0) & ";") = 0 Then
                    strBalance = strBalance & ";" & varData(0)
                End If
                rsTemp.AddNew
                rsTemp!���㷽ʽ = varData(0)
                rsTemp!��� = Val(varData(3))
                rsTemp.Update
            End If
        Next
    Next
    If strBalance <> "" Then
        strBalance = Mid(strBalance, 2)
        varBalance = Split(strBalance, ";")
        For i = 0 To UBound(varBalance)
            curMoney = 0
            rsTemp.Filter = "���㷽ʽ='" & varBalance(i) & "'"
            Do While Not rsTemp.EOF
                curMoney = curMoney + Nvl(rsTemp!���)
                rsTemp.MoveNext
            Loop
            strTmp = strTmp & "||" & varBalance(i) & "|" & curMoney
        Next
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    GetMedicareStr = Mid(strTmp, 3)
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub SetBalanceVal(colBalance As Collection, ByVal intPage As Integer, _
    ByVal strBalance As String, Optional ByRef strNone As String)
    '���ܣ�����ָ����ŵ���ָ�����ս��㷽ʽ����Чֵ
    '������
    '       strBalance-���ݽ��㷽ʽ�ַ������ý��㷽ʽ��¼������ʽ�����㷽ʽ1|���1||���㷽ʽ2|���2||...
    '˵�����ú�����colBalanceΪ׼����,����ҽ�������շ�Ҳ��
    '˵������������ҽ���շ��޸ı��ս���������۵�ҽ���շ����ø����ʻ��Ƚ�����
    Dim arrValue As Variant, arrPage As Variant
    Dim strTmp As String, i As Long, j As Long
    Dim varBalance As Variant, varTemp As Variant
    Dim strItem As String, curVal As Currency
    Dim blnFind As Boolean, rs���㷽ʽ As ADODB.Recordset
    
    If strBalance = "" Then Exit Sub
    
    Set rs���㷽ʽ = Get���㷽ʽ("�շ�")
    arrPage = Array()
    varBalance = Split(strBalance, "||")
    For i = 0 To UBound(varBalance)
        varTemp = Split(varBalance(i) & "|||", "|")
        strItem = varTemp(0): curVal = Val(varTemp(1))
        '���������øý��㷽ʽ,��Ϊҽ����Ľ��㷽ʽ
        rs���㷽ʽ.Filter = "����='" & strItem & "' And ����<>1 And ����<>2"
        If rs���㷽ʽ.EOF Then
            '��¼ҽ���е�����û�еĽ��㷽ʽ
            If InStr(strNone & ",", "," & strItem & ",") = 0 Then
                strNone = strNone & "," & strItem
            End If
        Else
            If colBalance.Count > 0 Then
                If UBound(colBalance(intPage)) >= 0 Then
                    blnFind = False
                    For j = 0 To UBound(colBalance(intPage))
                        '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
                        arrValue = Split(colBalance(intPage)(j), ";")
                        If arrValue(0) = strItem Then blnFind = True
                        If arrValue(0) = strItem And arrValue(3) <> curVal Then
                            strTmp = arrValue(0) & ";" & arrValue(1) & ";" & arrValue(2) & ";" & Format(curVal, "0.00")
                        Else
                            strTmp = arrValue(0) & ";" & arrValue(1) & ";" & arrValue(2) & ";" & arrValue(3)
                        End If
                        
                        ReDim Preserve arrPage(UBound(arrPage) + 1)
                        arrPage(UBound(arrPage)) = strTmp
                    Next
                    
                    If Not blnFind Then
                        ReDim Preserve arrPage(UBound(arrPage) + 1)
                        '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
                        arrPage(UBound(arrPage)) = strItem & ";" & curVal & ";" & Val(varTemp(2)) & ";" & curVal
                    End If
                Else
                     '������ʱǿ������:��֧��Ԥ�����ҽ�������շ�ʱ��
                    ReDim Preserve arrPage(UBound(arrPage) + 1)
                    arrPage(UBound(arrPage)) = strItem & ";" & Format(curVal, "0.00") & ";" & Val(varTemp(2)) & ";" & Format(curVal, "0.00")
                End If
            Else
                 '������ʱǿ������:��֧��Ԥ�����ҽ�������շ�ʱ��
                ReDim Preserve arrPage(UBound(arrPage) + 1)
                arrPage(UBound(arrPage)) = strItem & ";" & Format(curVal, "0.00") & ";" & Val(varTemp(2)) & ";" & Format(curVal, "0.00")
            End If
        End If
    Next

    colBalance.Remove intPage '����Ԫ�ز���ֱ���޸�
    If colBalance.Count >= intPage Then
        colBalance.Add arrPage, , intPage
    Else
        colBalance.Add arrPage
    End If
End Sub

Public Function WriteInforToCard(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String, _
        ByVal objSquareCard As Object, ByVal lngCardTypeID As Long, _
        ByVal strNos As String) As Boolean
    '����:��������Ϣд�뿨��
    '��Σ�
    '    frmMain - ���ô���
    '    lngModul - ģ���
    '    strPrivs - Ȩ�޴�
    '    objSquareCard - ҽ�ƿ�����
    '    strNOs - ���ݺţ���ʽ��'A0001','A0002','A0003',...��A0001,A0002,A0003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng����ID As Long, lng������� As Long
    
    Err = 0: On Error GoTo errH:
    '����:56615
    If InStr(strPrivs, ";������Ϣд��;") = 0 Then Exit Function
    
    strSQL = "Select Distinct A.����ID,B.�������" & _
        " From ������ü�¼ A,����Ԥ����¼ B,Table( f_Str2list([1])) J" & _
        " Where A.����ID=B.����ID And A.NO=J.Column_Value And  Nvl(A.���ӱ�־,0)<>9 And A.��¼���� = 1 " & _
        "       And A.��¼״̬ in(1,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ݽ������", Replace(strNos, "'", ""))
    If rsTemp.EOF Then Exit Function
    Do While Not rsTemp.EOF
        lng����ID = Val(Nvl(rsTemp!����ID))
        lng������� = Val(Nvl(rsTemp!�������))
        '���ý�����д���ӿ�
        If lng����ID <> 0 And lng������� <> 0 Then
            Call objSquareCard.zlMzInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng����ID, lng�������)
        End If
        rsTemp.MoveNext
    Loop
    
    WriteInforToCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function zlSaveTempPrintData(ByVal strNos As String, ByVal lng����ID As Long, ByVal strFactNO As String, ByRef lng��ӡID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʱ�Ĵ�ӡ����
    '���:strNos-���ݺ�
    '     lng����ID-����ID
    '     strFactNo-��ʼ��Ʊ��
    '����:lng��ӡID-���ش�ӡID
    '����:
    '����:���˺�
    '����:2016-05-03 16:44:43
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, blnTrans As Boolean
    
    On Error GoTo errHandle
        '������ʱ����
    Set cllPro = New Collection
    blnTrans = True
    If SaveTempPrintDataTocCllPro(strNos, strFactNO, lng����ID, lng��ӡID, cllPro) = False Then Exit Function
    zlExecuteProcedureArrAy cllPro, "������ʱƱ�ݴ�ӡ����"
    zlSaveTempPrintData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function SaveTempPrintDataTocCllPro(ByVal strNos As String, ByVal str��ʼ��Ʊ�� As String, ByVal lng����ID As Long, _
    ByRef lng��ӡID As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݺ���Ϣ���ֽ�ɱ��浽��ʱƱ�ݴ�ӡ������
    '���:strNos-���ݺ��ַ���������ö��ŷָ�
    '     str��Ʊ��-��ʼ��Ʊ��
    '     lng����ID-����ID
    '����:cllPro-������ʱƱ�ݴ�ӡ���ݵĹ���.
    '     lng��ӡID-���ش�ӡID
    '����:�ɹ�����true,���򷵻�false
    '����:���˺�
    '����:2016-04-27 17:48:42
    '˵����������ʱ��Ʊ�ݴ�ӡ���ݣ���Ҫ����Ϊ�����˲���Ʊ��ʱ�������򵥾ݺų���4000�����Զ��屨���������ƣ���ˣ���Ϊ����ʱ��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim cllData As Collection
    
    On Error GoTo errHandle
    
    lng��ӡID = zlDatabase.GetNextId("Ʊ�ݴ�ӡ����")
    Set cllData = New Collection
    If zlGetSplitString4000(strNos, cllData) = False Then Exit Function
    
    For i = 1 To cllData.Count
        '    Zl_��ʱƱ�ݴ�ӡ����_Insert
        strSQL = "Zl_��ʱƱ�ݴ�ӡ����_Insert("
        '    ��ӡid_In     Ʊ�ݴ�ӡ����.Id%Type,
        strSQL = strSQL & lng��ӡID & ","
        '    No_In         Varchar2,
        strSQL = strSQL & "'" & cllData(i) & "',"
        '    ����id_In     ��ʱƱ�ݴ�ӡ����.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '    ��ʼ��Ʊ��_In ��ʱƱ�ݴ�ӡ����.��ʼƱ��%Type
        strSQL = strSQL & "'" & str��ʼ��Ʊ�� & "')"
        zlAddArray cllPro, strSQL
    Next
    SaveTempPrintDataTocCllPro = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetSplitString4000(ByVal strSplitData As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��strSplitData���ݣ���4000���ַ��ֽ⣬�Ա㱣�������ݿ�(ÿ���ָ����ַ�ҪС��10)
    '���:strSplitData-Ҫ�ֽ������,�ö��ŷ���
    '����:cllPro-���ظ���
    '����:�ֽ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-05-04 09:43:32
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    On Error GoTo errHandle
    Set cllPro = New Collection
    If Len(strSplitData) <= 4000 Then
        cllPro.Add strSplitData: zlGetSplitString4000 = True: Exit Function
    End If
    
    Do While True
        If Len(strSplitData) < 4000 Then
            cllPro.Add strSplitData: zlGetSplitString4000 = True: Exit Function
        End If
        
        i = InStr(3950, strSplitData, ",")
        If i = 0 Then Exit Do
        
        strTemp = Mid(strSplitData, 1, i - 1)
        strSplitData = Mid(strSplitData, i + 1)
        cllPro.Add strTemp
    Loop
    If strSplitData <> "" Then cllPro.Add strSplitData
    zlGetSplitString4000 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Sub zlDeleteTempPrintData(lng��ӡID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ����ʱ��ӡ����
    '���:lng��ӡID
    '����:���˺�
    '����:2016-05-03 14:36:47
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error Resume Next
    strSQL = "Zl_��ʱƱ�ݴ�ӡ����_Delete( " & lng��ӡID & ")"
    zlDatabase.ExecuteProcedure strSQL, "ɾ����ʱƱ�ݴ�ӡ��������"
    Err = 0
End Sub

Public Function zlIsOnePatiPrint(ByVal strNo As String, ByRef strPrintNos As String, ByRef blnOnePatiPrint As Boolean, Optional ByVal blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�������Ƿ񰴲���һ�δ�ӡ��
    '���:strNo-��Ҫ�ش�NO
    '     blnNOMoved-�Ƿ�ת����ʷ��ռ�
    '����:���ر���һ�δ�ӡ��NO,����ö��ŷ���
    '     blnOnePatiPrint-����ǰ�����һ�δ�ӡ������true,���򷵻�False
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-05-03 17:12:20
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strNos As String
    strSQL = "" & _
    "   Select A2.NO,Max(nvl(A2.��ӡ����,0)) as ��ӡ����  " & _
    "   From  Ʊ�ݴ�ӡ���� A1,Ʊ�ݴ�ӡ���� A2  " & _
    "   Where A1.ID=A2.ID and A1.��������=A2.��������  And A1.NO=[1] And A1.��������=1" & _
    "   Group By A2.NO"
    If blnNOMoved Then strSQL = Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ָ�����ݻ�ȡһ���ӡ�����е���", strNo)
    blnOnePatiPrint = False
    
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            If blnOnePatiPrint = False And Val(Nvl(rsTemp!��ӡ����)) = 1 Then blnOnePatiPrint = True
            .MoveNext
        Loop
    End With
    strPrintNos = strNos
    If strPrintNos <> "" Then strPrintNos = Mid(strPrintNos, 2)
    zlIsOnePatiPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlFromCollectBulidSQL(ByVal cllData As Collection, ByRef strBoundSQL As String, ByRef varData() As Variant, _
    Optional ByVal strAliaName As String = "NO", Optional ByVal blnNumber As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ϣ����������ֵ��SQL
    '���:cllData-����
    '     strAliaName-����
    '     blnNumber-�Ƿ�����
    '����:strBoundSQL-�󶨵�SQL
    '       varData-����ֵ
    '����:��ϳɹ�������True,���򷵻�False
    '����:���˺�
    '����:2016-05-04 11:58:37
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varValue() As Variant
    Dim strSQL As String, strTable As String, i As Long
    

    On Error GoTo errHandle
    ReDim varValue(0 To cllData.Count - 1)
    For i = 1 To cllData.Count
        If blnNumber Then
            strTable = "Table(f_Num2list([" & i & "]))"
        Else
            strTable = "Table(f_Str2list([" & i & "]))"
        End If
        strSQL = strSQL & _
        " UNION ALL " & vbCrLf & _
        " Select Column_Value as " & strAliaName & " From " & strTable
        If blnNumber Then
            varValue(i - 1) = Val(cllData(i))
        Else
            varValue(i - 1) = CStr(cllData(i))
        End If
    Next
    If strSQL <> "" Then strSQL = Mid(strSQL, 11)
    varData = varValue
    strBoundSQL = strSQL
    zlFromCollectBulidSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlChargeBillIsAllDel(ByVal strNos As String, Optional ByVal lng��ӡID As Long = 0, Optional ByRef blnAllDel As Boolean, Optional strNotDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ����е��ݶ�ȫ����
    '���:strNos-(lng��ӡID=0ʱʹ��)ָ���ĵ��ݺ�,����ö��ŷָ�������:A0001,A0002,...
    '     lng��ӡID -���ݴ�ӡID�����
    '����:blnAllDel-ȫ�����꣬����true,���򷵻�False
    '     strNotDelNos-δ����ĵ��ݺ�
    '����:�����ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-05-05 11:21:31
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim varValue() As Variant, cllBound As Collection
    
    On Error GoTo errHandle
    If lng��ӡID > 0 Then
        strSQL = " " & _
        "   Select b.No,B.���, Sum(Nvl(b.����, 0) * Nvl(b.����, 0)) As ����  " & _
        "   From ������ü�¼ B, (Select Distinct NO From ��ʱƱ�ݴ�ӡ���� Where ID = [1]) A " & _
        "   Where Mod(b.��¼����, 10) = 1 And b.No = a.No And b.�۸񸸺� Is Null " & _
        "   Having Sum(Nvl(b.����, 0) * Nvl(b.����, 0)) <> 0 " & _
        "   Group By b.No,B.���"
        
        strSQL = "Select Distinct  NO From (" & strSQL & ") Where nvl(����,0)<>0  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ�ݴ�ӡ��������ȡ�����Ƿ�ȫ������", lng��ӡID)
    ElseIf Len(strNos) <= 4000 Then
        strSQL = " " & _
        "   Select b.No,B.���, Sum(Nvl(b.����, 0) * Nvl(b.����, 0)) As ����  " & _
        "   From ������ü�¼ B " & _
        "   Where Mod(b.��¼����, 10) = 1  And b.�۸񸸺� Is Null " & _
        "         And b.No in (select Column_Value From Table(f_Str2list([1])))" & _
        "   Having Sum(Nvl(b.����, 0) * Nvl(b.����, 0)) <> 0 " & _
        "   Group By b.No,B.���"
        
        strSQL = "Select Distinct  NO From (" & strSQL & ") Where nvl(����,0)<>0  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ�ݴ�ӡ��������ȡ�����Ƿ�ȫ������", strNos)
    Else
        If zlGetSplitString4000(strNos, cllBound) = False Then Exit Function
        If zlFromCollectBulidSQL(cllBound, strSQL, varValue) = False Then Exit Function
    
        strSQL = " With ������Ϣ as (" & strSQL & ") "
        strSQL = strSQL & vbCrLf & _
            "   Select b.No,B.���, Sum(Nvl(b.����, 0) * Nvl(b.����, 0)) As ����  " & _
            "   From ������ü�¼ B,������Ϣ A" & _
            "   Where Mod(b.��¼����, 10) = 1  And b.�۸񸸺� Is Null " & _
            "         And b.No =A.NO " & _
            "   Having Sum(Nvl(b.����, 0) * Nvl(b.����, 0)) <> 0 " & _
            "   Group By b.No,B.���"
        strSQL = "Select Distinct  NO From (" & strSQL & ") Where nvl(����,0)<>0  "
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "����Ʊ�ݴ�ӡ��������ȡ�����Ƿ�ȫ������", varValue)
     End If
         
     With rsTemp
        blnAllDel = True
        strNotDelNos = ""
        Do While Not .EOF
            If blnAllDel Then blnAllDel = False
            strNotDelNos = strNotDelNos & "," & Nvl(rsTemp!NO)
            .MoveNext
        Loop
        If strNotDelNos <> "" Then strNotDelNos = Mid(strNotDelNos, 2)
     End With
     zlChargeBillIsAllDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Upgradeҽ��ִ�мƼ�ִ��״̬(ByVal strNos As String) As Boolean
    '���ܣ�����"ҽ��ִ�мƼ�.ִ��״̬"
    '��Σ�
    '   strNos ���ݺţ���ʽ:A001,A002,A003,...
    '���أ��������򷵻�True�����򷵻�False
    '�����:99715
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strҽ��IDs As String, varҽ��IDs As Variant
    Dim i As Long
    
    On Error GoTo errHandler
    strҽ��IDs = ""
    strSQL = " Select /*+cardinality(j,10)*/ Distinct a.ҽ����� As ҽ��ID, a.No" & vbNewLine & _
        " From ������ü�¼ A, ����ҽ������ B, ҽ��ִ�мƼ� C, Table(f_Str2list([1])) J" & vbNewLine & _
        " Where b.ҽ��id = a.ҽ����� And b.No = a.No" & vbNewLine & _
        "       And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And c.�շ�ϸĿid + 0 = a.�շ�ϸĿid" & vbNewLine & _
        "       And a.No = j.Column_Value And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And a.ҽ����� Is Not Null" & vbNewLine & _
        "       And b.��¼���� = 1 And c.ִ��״̬ Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж�ҽ��ִ�мƼ�ִ��״̬�Ƿ�������", strNos)
    If rsTemp.RecordCount = 0 Then
        Upgradeҽ��ִ�мƼ�ִ��״̬ = True
        Exit Function
    End If
    
    '�Ѽ�ҽ��ID
    Do While Not rsTemp.EOF
        If InStr(";" & strҽ��IDs & ";", ";" & Nvl(rsTemp!ҽ��id) & "," & Nvl(rsTemp!NO) & ";") = 0 Then
            strҽ��IDs = strҽ��IDs & ";" & Nvl(rsTemp!ҽ��id) & "," & Nvl(rsTemp!NO)
        End If
        rsTemp.MoveNext
    Loop
    If strҽ��IDs = "" Then
        Upgradeҽ��ִ�мƼ�ִ��״̬ = True
        Exit Function
    End If
    
    '��������
    strҽ��IDs = Mid(strҽ��IDs, 2)
    varҽ��IDs = Split(strҽ��IDs, ";")
    For i = 0 To UBound(varҽ��IDs)
        'Zl_ҽ��ִ�мƼ�_����(
        strSQL = "Zl_ҽ��ִ�мƼ�_����("
        '  ҽ��id_In   ����ҽ��ִ��.ҽ��id%Type,
        strSQL = strSQL & "" & Split(varҽ��IDs(i), ",")(0) & ","
        '  No_In       ����ҽ������.No%Type,
        strSQL = strSQL & "'" & Split(varҽ��IDs(i), ",")(1) & "',"
        '  ��¼����_In ����ҽ������.��¼����%Type
        strSQL = strSQL & "" & "1" & ")"
        zlDatabase.ExecuteProcedure strSQL, "��������"
    Next
    
    Upgradeҽ��ִ�мƼ�ִ��״̬ = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreatePlugIn(ByVal lngMod As Long) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugIn = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    If gobjPlugIn Is Nothing Then
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    End If
    If gobjPlugIn Is Nothing Then Exit Function
    
    Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "Initialize")
        Set gobjPlugIn = Nothing
        Exit Function
    End If
    
    CreatePlugIn = True
End Function

Public Function zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String) As Boolean
'���ܣ���Ҳ���������ͬʱ�ж��Ƿ�Ϊ�ǽӿڷ��������ڵĴ���
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & _
            objErr.Number & vbCrLf & _
            objErr.Description, vbInformation, gstrSysName
    End If
    Err.Clear
End Function

Public Function CreatePublicDrug(ByVal lngSys As Long, _
    cnOracle As ADODB.Connection, ByVal strDBUser As String) As Boolean
    '���ܣ���̬����ҩƷ��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    If Not gobjPublicDrug Is Nothing Then CreatePublicDrug = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjPublicDrug = CreateObject("zlPublicDrug.clsPublicDrug")
    
    Err = 0: On Error GoTo errHandler
    If gobjPublicDrug Is Nothing Then
        MsgBox "ҩƷ����������zlPublicDrug������ʧ�ܣ�����ϵͳ��Ա��ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    'Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If gobjPublicDrug.zlInitCommon(lngSys, cnOracle, strDBUser) = False Then
        MsgBox "ҩƷ����������zlPublicDrug����ʼ��ʧ�ܣ�����ϵͳ��Ա��ϵ��", vbInformation, gstrSysName
        Set gobjPublicDrug = Nothing: Exit Function
    End If
    CreatePublicDrug = True
    Exit Function
errHandler:
    Set gobjPublicDrug = Nothing
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckChargeItemByPlugIn(objPlugIn As Object, _
    lngSys As Long, ByVal lngModule As Long, _
    ByVal intType As Integer, ByVal intMode As Integer, _
    ByRef rsDetail As ADODB.Recordset, Optional strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ҳ������շ���Ŀ��Ч�Խ��м��
    '���:lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '     intType:0-����;1-סԺ
    '     intMode:0-¼����ϸʱ�ĳ�����;1-���浥��ǰ�Ļ��ܼ��
    '     rsDetail-����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������,
    '                  ִ�п���ID���������ʣ�1-�շѵ�,2-���ʵ�)���Ƿ񻮼�(1-����;0-�������շѼ����ʵ�)
    '     strExpend-���Ժ���չ��������
    '����:strExpend-���Ժ���չ��������
    '����:���ݺϷ�����true,���򷵻�False
    '����:Ƚ����
    '����:2017-04-19 10:09:26
    '�����:105189
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '1.û����Ҳ���ʱ����Ϊ���ͨ��
    '2.��Ҳ�������CheckChargeItem�ӿڣ�Ҳ��Ϊ���ͨ��
    If objPlugIn Is Nothing Then CheckChargeItemByPlugIn = True: Exit Function
    
    On Error Resume Next
    If objPlugIn.CheckChargeItem(lngSys, lngModule, intType, intMode, rsDetail, strExpend) = False Then
        'ע�⣬�ӿڲ�����ʱҲ�����
        If Err <> 0 Then
            If Err.Number = 438 Then '�ӿڲ����ڣ���Ϊ���ͨ��
                CheckChargeItemByPlugIn = True
                Exit Function
            End If
            Call zlPlugInErrH(Err, "CheckChargeItem")
        End If
        Exit Function
    End If
    CheckChargeItemByPlugIn = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PatiValiedCheckByPlugIn(ByVal lngModule As Long, ByVal lng����ID As Long) As Boolean
    '������ҽӿ� PatiValiedCheck ��鲡����Ϣ
    '�����:102234,138602
    '˵����
    '   1.û����Ҳ���ʱ����Ϊ���ͨ��
    '   2.��Ҳ�������PatiValiedCheck�ӿڣ�Ҳ��Ϊ���ͨ��
    '   3.δ�������˲����
    
    If gobjPlugIn Is Nothing Then PatiValiedCheckByPlugIn = True: Exit Function
    If lng����ID = 0 Then PatiValiedCheckByPlugIn = True: Exit Function
    
    On Error Resume Next
    'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
        ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
        ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
    '���ܣ���鵱ǰ�����Ƿ���ָ�������ⲡ��
    '���أ�trueʱ�������������Falseʱ���������
    '������
    '      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '      lngType �������ͣ�1������Һţ�2��סԺ��Ժ��3�������շѣ�4��סԺ���ʡ�
    '      lngPatiID-����ID: �½����ģ�Ϊ0,�����뽨������ID
    '      lngPageID-��ҳID: �½����ģ�Ϊ0,�����뽨����ҳID(סԺ������ҳID) ����˵������ lngType=4 ʱ�Ŵ��� lngPageID����������0
    '      strPatiInforXML-������Ϣ:���δ�������˴��룬"�������Ա����䣬�������ڣ�ҽ���ţ����֤��"���������� ��ʽ:2016-11-11 12:12:12
    '                      �̶���ʽ��<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH>
    '      strReserve=��������,������չʹ��
    On Error Resume Next
    If gobjPlugIn.PatiValiedCheck(glngSys, lngModule, 3, lng����ID, 0, "") = False Then
        'ע�⣬�ӿڲ�����ʱҲ�����
        If Err <> 0 Then
            If Err.Number = 438 Then '�ӿڲ����ڣ���Ϊ���ͨ��
                PatiValiedCheckByPlugIn = True: Exit Function
            End If
            Call zlPlugInErrH(Err, "PatiValiedCheck")
        End If
        Exit Function
    End If
    PatiValiedCheckByPlugIn = True
End Function

Public Function GetPriceGradeFromNos(ByVal strNos As String, Optional ByVal lng����ID As Long) As String
    '���ܣ����ݵ��ݺż�վ���ȡ��ͨ��Ŀ�۸�ȼ�
    '��Σ�
    '   strNos ���ݺţ������ţ������Ƕ�����ݺţ�Ϊ"'AAA','BBB',..."����ʽ
    Dim strPriceGrade As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varNos As Variant
    
    On Error GoTo errHandler
    If lng����ID = 0 Then
        varNos = Split(Replace(strNos, "'", ""), ",")
        If UBound(varNos) = -1 Then Exit Function
        
        strSQL = _
            "Select a.����id, b.���� As ���ʽ" & vbNewLine & _
            "From ������ü�¼ A, ҽ�Ƹ��ʽ B" & vbNewLine & _
            "Where a.���ʽ = b.���� And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And a.No = [1] And Rownum < 2"
    Else
        strSQL = _
            "Select a.����id, b.���� As ���ʽ" & vbNewLine & _
            "From ������ü�¼ A, ҽ�Ƹ��ʽ B" & vbNewLine & _
            "Where a.���ʽ = b.���� And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And a.����ID = [2] And Rownum < 2"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ID��ҽ�Ƹ��ʽ", CStr(varNos(0)), lng����ID)
    If rsTemp.EOF Then Exit Function
    
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(rsTemp!����ID)), 0, Nvl(rsTemp!���ʽ), , , strPriceGrade)

    GetPriceGradeFromNos = strPriceGrade
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CreatePublicExpenseObject(ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ò���
    '���:
    '����:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    gstrҩƷ�۸�ȼ� = "": gstr���ļ۸�ȼ� = "": gstr��ͨ�۸�ȼ� = ""
    If gobjPublicExpense Is Nothing Then
        Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If gobjPublicExpense Is Nothing Then Exit Sub
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If gobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    gintPriceGradeStartType = gobjPublicExpense.zlGetPriceGradeStartType()
    If gintPriceGradeStartType = 0 Then Exit Sub
    '��ȡվ��۸�ȼ�
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
End Sub

Public Function UserIsClinic(ByVal lng��ԱID As Long) As Boolean
    '�жϵ�ǰ����Ա�Ƿ�Ϊ�ٴ�������Ա
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    strSQL = _
        "Select 1" & vbNewLine & _
        "From ������Ա A,���ű� B, ��������˵�� C" & vbNewLine & _
        "Where a.����id = b.Id And b.id = c.����id" & vbNewLine & _
        "      And c.�������� In ('�ٴ�', '���', '����', '����', '����', '����')" & vbNewLine & _
        "      And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
        "      And (b.վ��='" & gstrNodeNo & "' Or b.վ�� is Null) " & vbNewLine & _
        "      And a.��Աid = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng��ԱID)
    If rsTemp.RecordCount = 0 Then Exit Function
    UserIsClinic = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSelectWholeItems(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
     ByRef rsOutSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ŀѡ����(ѡ�������)
    '���:lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:rsOutSel-�ɹ�ʱ,����ѡ��ĳ�����Ŀ(���ֶ�:ϸĿID,����,����,���,��������,ִ�п���....)
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2017-11-08 16:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPublicExpense Is Nothing Then
        Call CreatePublicExpenseObject(lngModule)
    End If
    If gobjPublicExpense Is Nothing Then Exit Function
    zlSelectWholeItems = gobjPublicExpense.zlSelectWholeItems(frmMain, lngModule, strPrivs, rsOutSel)
End Function

Public Sub ZlShowBillFormat(ByVal lngModule As Long, lblFormat As Label, ByVal intFormat As Integer)
    '���ܣ���ʾƱ�ݸ�ʽ����
    '��Σ�
    '   lngModule - ģ���
    '   lblFormat - ��ʾƱ�ݸ�ʽ�ı�ǩ����
    '   intFormat - Ʊ�ݸ�ʽ���
    '���أ�Ʊ�ݸ�ʽ������
    Dim strFormatName As String
    
    On Error GoTo errHandler
    strFormatName = ZlGetBillFormat(lngModule, intFormat)
    If strFormatName = "" Then
        lblFormat.Visible = False
    Else
        lblFormat.Caption = "Ʊ��:" & strFormatName
        lblFormat.Visible = True
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ZlGetBillFormat(ByVal lngModule As Long, ByVal intFormat As Integer) As String
    '���ܣ���ȡƱ�ݸ�ʽ����
    '��Σ�
    '   lngModule - ģ���
    '   intFormat - Ʊ�ݸ�ʽ���
    '���أ�Ʊ�ݸ�ʽ������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strRptName As String
    
    On Error GoTo errHandler
    If lngModule = 1124 Then
        strRptName = "ZL" & glngSys \ 100 & "_BILL_1124"
    Else
        strRptName = "ZL" & glngSys \ 100 & "_BILL_1121_1"
    End If
    
    If intFormat = 0 Then '��ȱʡƱ�ݸ�ʽ��ʾ
        intFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
    End If
    
    strSQL = _
        "Select b.˵��" & vbNewLine & _
        "From zlReports A, zlRPTFMTs B" & vbNewLine & _
        "Where a.Id = b.����id And a.��� = [1] And b.��� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����ʽ˵��", strRptName, intFormat)
    If rsTmp.EOF Then Exit Function
    
    ZlGetBillFormat = Nvl(rsTmp!˵��)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlChargeSaveValied_Plugin(ByVal lngModule As Long, ByVal int��¼���� As Integer, ByVal bln���� As Boolean, _
    ByVal bln���۵� As Boolean, ByVal strNos As String, ByVal rsSaveItems As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ң���鱣�����ݵĺϷ���
    '���:lngModule-ģ���
    '     int��¼����-1-�շѵ�;2-���ʵ�
    '     bln���۵�-�Ƿ�ǰ�Ǳ���Ļ��۵�
    '     strNOs-�����շ�ʱ������Ļ��۵��ţ��Ա����շѵĻ��۵���)
    '     rsSaveItems=��ǰ�������Ŀ����(�ֶ� :����ID����ҳID,�������, ���,�۸񸸺�,�շ�ϸĿID��������Ŀid������ �����Σ���׼���ۣ�Ӧ�ս�� ��
    '                                            ʵ�ս�����ʱ�䣬��Ŀ���룬��Ŀ���ƣ��������,��������ID,������,ִ�в���ID)
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2017-12-13 17:55:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1.û����Ҳ���ʱ����Ϊ���ͨ��
    '2.��Ҳ�������CheckChargeItem�ӿڣ�Ҳ��Ϊ���ͨ��
    If gobjPlugIn Is Nothing Then zlChargeSaveValied_Plugin = True: Exit Function
    
    On Error Resume Next
    If gobjPlugIn.ChargeSaveValied(glngSys, lngModule, int��¼����, bln����, bln���۵�, strNos, rsSaveItems) = False Then
        'ע�⣬�ӿڲ�����ʱҲ�����
        If Err <> 0 Then
            If Err.Number = 438 Then '�ӿڲ����ڣ���Ϊ���ͨ��
                zlChargeSaveValied_Plugin = True
                Exit Function
            End If
            Call zlPlugInErrH(Err, "ChargeSaveValied")
            Err = 0: On Error GoTo 0
        End If
        Exit Function
    End If
    zlChargeSaveValied_Plugin = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlChargeSaveAfter_Plugin(ByVal lngModule As Long, ByVal lng����ID, ByVal lng��ҳID As Long, _
                                    ByVal bln���� As Boolean, ByVal int��¼���� As Integer, ByVal strNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ң���鱣�����ݵĺϷ���
    '���:     lngSys , lngModual = ��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    '   lng����ID�����ʱ�ʱ������0)
    '   lng��ҳID�����ʱ�ʱ������0)
    '   bln���� -�Ƿ��������
    '   int��¼����-1-�շ�;2-����
    '   strNOs-���ݺ�,����ö��ŷָ�
    '����:���˺�
    '����:2017-12-13 17:55:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1.û����Ҳ���ʱ����Ϊ���ͨ��
    '2.��Ҳ�������CheckChargeItem�ӿڣ�Ҳ��Ϊ���ͨ��
    If gobjPlugIn Is Nothing Then Exit Sub
    
    On Error Resume Next
    Call gobjPlugIn.ChargeSaveAfter(glngSys, lngModule, lng����ID, lng��ҳID, bln����, int��¼����, strNos)
    If Err = 0 Then Exit Sub
    
    'ע�⣬�ӿڲ�����ʱҲ�����
    If Err.Number = 438 Then Exit Sub  '�ӿڲ����ڣ���Ϊ���ͨ��
    Call zlPlugInErrH(Err, "ChargeSaveAfter")
    Err = 0: On Error GoTo 0
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function zlGetSaveDataItems_Plugin(ByVal objBills As ExpenseBill, ByRef str����Nos As String, ByRef rsItems As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��Ҫ�����������ϸ(�˹��̣���Ҫ��Ӧ������ҽӿ�,���û����Һţ���ֱ�ӷ���True,��¼������Nothing)
    '���:objBills-���ݶ���
    '����:str����Nos-���ص�ǰ�շ����漰�Ļ��۵�
    '     rsItems-���ص�ǰ��Ҫ��������ݼ�(�ֶ� :����ID����ҳID,�������, ���,�۸񸸺�,�շ�ϸĿID��������Ŀid������ �����Σ���׼���ۣ�Ӧ�ս�� ��
    '                                            ʵ�ս�����ʱ�䣬��Ŀ���룬��Ŀ���ƣ��������,��������ID,������,ִ�в���ID)
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2017-12-14 11:41:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objBillDetail As BillDetail  '���ݵ��շ�ϸĿ����
    Dim objBillIncome As BillInCome
    Dim int�۸񸸺� As Integer
    Dim p As Long, int��� As Integer
    
    On Error GoTo errHandle
    
    Set rsItems = Nothing
    str����Nos = ""
    
    If gobjPlugIn Is Nothing Then zlGetSaveDataItems_Plugin = True: Exit Function
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "�������", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "���", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "�۸񸸺�", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "�շ���ĿID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "����", adDouble, , adFldIsNullable
    rsItems.Fields.Append "����", adDouble, , adFldIsNullable
    rsItems.Fields.Append "��׼����", adDouble, , adFldIsNullable
    rsItems.Fields.Append "Ӧ�ս��", adDouble, , adFldIsNullable
    rsItems.Fields.Append "ʵ�ս��", adDouble, , adFldIsNullable
    rsItems.Fields.Append "����ʱ��", adVarChar, 20, adFldIsNullable
    rsItems.Fields.Append "��Ŀ����", adVarChar, 30, adFldIsNullable
    rsItems.Fields.Append "��Ŀ����", adVarChar, 200, adFldIsNullable
    rsItems.Fields.Append "�������", adVarChar, 2, adFldIsNullable
    rsItems.Fields.Append "��������ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "������", adVarChar, 20, adFldIsNullable
    rsItems.Fields.Append "ִ�в���ID", adBigInt, , adFldIsNullable
    
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
    
     '��ÿ�ŵ��ݶ���ִ�б���
    For p = 1 To objBills.Pages.Count
       
        If objBills.Pages(p).NO = "" Then
            int��� = 0
            For Each objBillDetail In objBills.Pages(p).Details
                If objBillDetail.���� <> 0 Then
                    int�۸񸸺� = 0
                    For Each objBillIncome In objBillDetail.InComes
                      int��� = int��� + 1 '��ǰ��¼���
                       rsItems.AddNew
                       rsItems!����ID = objBills.����ID
                       rsItems!��ҳID = objBills.��ҳID
                       rsItems!������� = p
                       rsItems!��� = int���
                       rsItems!�۸񸸺� = IIf(int�۸񸸺� = 0, Null, int���)
                       rsItems!�շ���ĿID = objBillDetail.�շ�ϸĿID
                       rsItems!������ĿID = objBillIncome.������ĿID
                       rsItems!���� = objBillDetail.����
                       rsItems!���� = objBillDetail.����
                       rsItems!��׼���� = objBillIncome.��׼����
                       rsItems!Ӧ�ս�� = objBillIncome.Ӧ�ս��
                       rsItems!ʵ�ս�� = objBillIncome.ʵ�ս��
                       rsItems!����ʱ�� = Format(objBills.����ʱ��, "yyyy-mm-dd HH:MM:SS")
                       rsItems!��Ŀ���� = objBillDetail.Detail.����
                       rsItems!��Ŀ���� = objBillDetail.Detail.����
                       rsItems!������� = objBillDetail.�շ����
                       rsItems!ִ�в���ID = objBillDetail.ִ�в���ID
                       rsItems!��������ID = objBills.Pages(p).��������ID
                       rsItems!������ = objBills.Pages(p).������
                       rsItems.Update
                      If int�۸񸸺� = 0 Then int�۸񸸺� = int���
                    Next     'ÿһ���շ���Ŀ
                End If
            Next
        Else
            str����Nos = str����Nos & "," & objBills.Pages(p).NO
        End If
    Next  '��һ�ŵ�
    If str����Nos <> "" Then str����Nos = Mid(str����Nos, 2)
    
    zlGetSaveDataItems_Plugin = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


