Attribute VB_Name = "mdlInExse"
Option Explicit 'Ҫ���������
Public Enum gBalanceBill
    g_Ed_������� = 0
    g_Ed_סԺ���� = 1
    g_Ed_���½��� = 2
    g_Ed_ȡ������ = 3
    g_Ed_�������� = 4
    g_Ed_�������� = 5
    g_Ed_���ݲ鿴 = 6
End Enum

Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
'============����ϵͳ����=====================
'ҽ������f
Public gclsInsure As New clsInsure
Public gstrҽ���������� As String 'ҽ����������ķ�������
Public gstr���ѷ������� As String '���Ѳ�������ķ�������
Public gbytҽ�������� As Byte '0-�����м�顢1-��鲢����δ������Ŀ��2-��鲢��ֹδ������Ŀ
Public gbln�����л� As Boolean '35242
'ˢ������
Public gdblԤ��������鿨 As Double 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
Public gbytԤ����˷��鿨 As Byte 'Ԥ����˷�ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
Public gbln���ѿ��˷��鿨 As Boolean '���ѿ��˷�ʱ�Ƿ�ˢ����֤

'LED����
Public gblnLED As Boolean        '����ʱ�Ƿ�����LED�豸����
Public gblnLedWelcome As Boolean '�Ƿ��ڽ������겡�˺���ʾ��ӭ��Ϣ
Public gobjKernel As Object
'Ʊ�ݿ���
Public gblnStrictCtrl As Boolean '�Ƿ��ϸ�Ʊ�ݹ���
Public gbytFactLength As Byte 'Ʊ�ݺ��볤��

Public gobjBillPrint As Object '������Ʊ�ݴ�ӡ����
Public gblnBillPrint As Boolean '������Ʊ�ݴ�ӡ�����Ƿ����

Public gobjTax As Object '˰�ش�ӡ�ӿڶ���
Public gblnTax As Boolean '�����Ƿ�ʹ��˰�ش�ӡ
Public gstrTax As String
Public gblnNurseStation As Boolean
Public gblnPrintByPatient As Boolean '��Լ��λ���ʰ����˷ֱ��ӡƱ��
Public gbytInvoiceKind As Byte      '0-סԺҽ�Ʒ��վ�,1-����ҽ�Ʒ��վ�
Public gbytFeePrintSet As Byte      '0-����ӡ;1-��ӡ��ʾ;2-��ӡ������ʾ

'���ü������
Public gBytMoney As Byte '�շѷֱҴ�����
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"
Public gbln��������ۿ� As Boolean '������Ŀ���ܼ����ۿ�
Public gblnסԺ��λ As Boolean      'ҩƷ��סԺ��λ,�����ۼ۵�λ
Public gcurMaxMoney As Currency '���ʷ���������ѽ��


'ҩ����ؿ���
Public gblnStock As Boolean 'ָ��ҩ��ʱ�Ƿ��޶�����ҩƷ�Ŀ��
Public gbln���뷢ҩ As Boolean '�Ƿ������շ��뷢ҩ����
Public gint���ķ��Ͽ��� As Integer    '������ɺ��Ƿ��Զ�����:0-���Զ����ϣ�1-�Զ����ϣ�2-�����ҿ���ʱ�Զ�����

'ҩ�������ڿ���
Public glng��ҩ�� As Long 'ָ������ҩ��,0Ϊ��̬����
Public glng��ҩ�� As Long 'ָ������ҩ��,0Ϊ��̬����
Public glng��ҩ�� As Long 'ָ���ĳ�ҩ��,0Ϊ��̬����
Public glng���ϲ��� As Long 'ָ�������ķ��ϲ���,0Ϊ��̬����
Public gblnҩ���ϰల�� As Boolean '�Ƿ��������ϰల��
Public gbytSendMateria As Byte '0-���ʺ󲻷�ҩ,1-�Զ���ҩ,2-��ʾ��ҩ
Public gbytMediOutMode As Byte '����ҩƷ���ⷽʽ��0-�������Ƚ��ȳ���1-��Ч������ȳ�,Ч����ͬ�����ٰ������Ƚ��ȳ�
Public gbln����ʾ�޿������ As Boolean

'���뷢ҩʱҪ������ҩ��
Public gstr��ҩ�� As String
Public gstr��ҩ�� As String
Public gstr��ҩ�� As String

Public gbln����ҩ�� As Boolean '�Ƿ���ʾ����ҩ�����
Public gbln����ҩ�� As Boolean '�Ƿ���ʾ����ҩ����


'�����������
Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gbytCode As Byte '�������ɷ�ʽ��0-ƴ��,1-���,2-����
Public gblnMyStyle As Boolean 'ʹ�ø��Ի����

Public gint������� As Integer
Public gbln�շ���� As Boolean '�Ƿ������������
Public gblnFeeKindCode As Boolean '�������ʱ,��λ�����շ�������

Public gbln������ As Boolean '�����Ƿ�������뿪����
Public gbln������ As Boolean '�����Ƿ�����������ƵĿ�����
Public gstrMatchMode As String  '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��

'��ҩ������
Public grsABCNum As ADODB.Recordset
Public gstrABC As String '��������Ŀ����ĸ

'��������
Public gbytBilling As Byte '0-����,1-����,2-���
Public gstrModiNO As String '�޸ĺ�������µ��ݺ�
Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gbytWarn As Byte '���ʱ�������ֵ
Public gbln�����������۷��� As Boolean '���ʱ����������۷���

Public gblnPrice As Boolean '�Ƿ��������Ϊ���۵�
Public gbln�������� As Boolean '���۲��˼���
Public gblnסԺ���� As Boolean
Public gblnÿ��סԺ��סԺ�� As Boolean
Public gobjPati As Object

'ҽ�����
Public gblnҩ�ƻ��۵� As Boolean
Public gbln�������۵� As Boolean
Public gblnִ�к���� As Boolean

'�������
Public gstr�շ���� As String '��������շ����
Public gblnPay As Boolean '��ҩ�Ƿ����븶��
Public gblnTime As Boolean '����Ƿ����븶��
Public gbln��ʿ As Boolean '�������Ƿ���ʾ��ʿ
Public gblnFromDr As Boolean '������ȷ������
Public gbyt��������ʾ As Byte
'���˺� ����:????    ����:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '����С������
Public gstrFeePrecisionFmt As String '����С����ʽ:0.00000

'��ӡ����
Public gbln���ʴ�ӡ As Boolean
Public gbln���۴�ӡ As Boolean
Public gbln��˴�ӡ As Boolean
 
'ҽ��ִ��
Public gbln����ִ�� As Boolean 'ִ���߱��˵Ǽ�
Public gblnExeҽ�� As Boolean
Public gstrExe��Դ As String
Public gstrExe��� As String
Public gbytExe���ﵥ������ As Byte
Public gbytExeסԺ�������� As Byte
Public gbytExe��쵥������ As Byte
Public gbytExe��ӡ��ʽ As Byte
Public gblnִ�к��� As Boolean            '�������õ�������ִ�к��Զ�����,ȡ��ִ�к��Զ�����

Public gobjCustBill As Object               '�Զ�����ʵ�����
Public gbln�������� As Boolean             '�����������,�������ʱѡ��������,�򱣴�ʱ,���ټ��
Public grs�շ���� As ADODB.Recordset
'-------------------------------------------------------------------------------------------------------------------------------------------------
'--����ϵͳ����
Private Type TY_System_para_Balance
    blnˢ���������� As Boolean  '�Ƿ�ˢ����������
    bln��Ժ��׼���� As Boolean '1-��Ժ��׼����,0-��Ժ�������
    bytAuditing As Byte  '����δ��˵��ݵĽ��ʴ���:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    byt���δִ�� As Byte    '��Ժ�ͽ��ʳ�Ժʱ����Ƿ���δִ����Ŀ��δ��ҩƷ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    byt���δ��ҩ As Byte   '�ڳ�Ժ���ʼ�������������г�Ժʱ�Ƿ��鲡�˵�δ��ҩƷ��Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    byt������δִ�� As Byte    '�������ʱ����Ƿ���δִ����Ŀ��δ��ҩƷ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    byt������δ��ҩ As Byte   '�������ʱ�Ƿ��鲡�˵�δ��ҩƷ��Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    blnҽ��������ܳ�Ժ As Boolean 'ҽ���´��Ժҽ���������˳�Ժ
End Type

Private Type Ty_System_Para
     bytҩƷ������ʾ As Byte   'ҩƷ������ʾ�������浥����ϸ������������桢ֱ�ӽ����ҩƷѡ����ʱ��ҩƷ������ʾ����0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
     byt����ҩƷ��ʾ As Byte  '����ҩƷ��ʾ��ͨ��������뷽ʽ����ѡ����ʱҩƷ���Ƶ���ʾ����0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
     int���ݲ�¼ʱ�� As Integer '���ݲ�¼ʱ��
     byt������˷�ʽ As Byte '49501:������˷�ʽ:0-δ��˲�������ʣ�ȱʡΪ0;1-���ʱ����������ú�ҽ��������ҽ�������ͷ��õ�����
     blnδ��ƽ�ֹ���� As Boolean '51612
     byt��������ʶ����� As Byte   '�Ƿ������ʶ��::1-����ͨ��ɨ��¼���¼������;0-�����ƣ�����ͨ������Ȳ���
     strCardPass As String 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'
     TY_Balance As TY_System_para_Balance
End Type
Public gTy_System_Para As Ty_System_Para


'==============���ʲ���===============
'���ʲ���
Public gblnAutoOut As Boolean '��Ժ���˽��ʺ��Ƿ��Զ���Ժ
Public gblnҽ��������ܳ�Ժ As Boolean 'ҽ���´��Ժҽ���������˳�Ժ
Public gintOutDay As Integer '���ʿ�ѡ���Ժ��������
Public gbytAuditing As Byte  '����δ��˵��ݵĽ��ʴ���:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gblnZero As Boolean '����ʱ�Ƿ��������
Public gstrCardPass As String 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'
Public gbyt���δִ�� As Byte    '��Ժ�ͽ��ʳ�Ժʱ����Ƿ���δִ����Ŀ��δ��ҩƷ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gbyt���δ��ҩ As Byte   '�ڳ�Ժ���ʼ�������������г�Ժʱ�Ƿ��鲡�˵�δ��ҩƷ��Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gint����ʱ�� As Integer '0-���Ǽ�ʱ��,1-������ʱ��
Public gbln��Ժ��׼���� As Boolean '1-��Ժ��׼����,0-��Ժ�������
Public gbln����ָ��Ԥ���� As Boolean  '��ʹ��ָ��סԺ������Ԥ����
Public gbln���סԺ������������ As Boolean '�ж��סԺ���õĲ����Զ�������������
Public gbyt���ʼ����տ��� As Byte '��Ժ����ʱ��鲡�˵Ĵ��տ���,0-��ֹ,1-����
Public gbln��;������Ԥ�� As Boolean '��;����ȱʡ��Ԥ����
Public gstr���㷽ʽ��ʾ˳�� As String   '32322
Public gbyt����ʱ��Ѫ�Ѽ�� As Byte   '34260
'=======ϵͳ������ر���============
Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
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
    support��Ժ��ʵ�ʽ��� = 29       '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    support�����ֳ�����ϸ = 32    '�������סԺ���ʴ�����ÿ����ϸ���в��ֳ���
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    supportסԺ�������� = 34        'HISʼ����ΪסԺ֧�ֽ������ϣ������֧����ҽ���ӿ��ڲ��������ؼټ��ɣ����Ӹò�����Ϊ�����GetCapability�����������ֽ��㷽ʽ�Ƿ�֧��ȫ��
    support�������� = 35            '�Ƿ����������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
    support����_ָ��סԺ���� = 36   '�Ƿ�֧��ָ��סԺ��������ҽ������
    support����_ָ�����ڷ�Χ = 37   '�Ƿ�֧��ָ���������ڷ�Χ����ҽ������
    support����_����Ӥ�������� = 38 '�Ƿ���������Ӥ��������
    
    support������� = 41            '�Ƿ�֧������ҽ�����˵ļ��ʷ���ʹ��������������
    support����_ָ������ = 42           '�Ƿ������ڽ������ý�����ָ������
    support����_ָ��������Ŀ = 43       '�Ƿ������ڽ������ý�����ָ��������Ŀ
    support����_�������ú���ýӿ� = 44 '���Ϊ�����ڽ������ú�ŵ���סԺ������㣬֮ǰ������
    support����_ָ���������� = 45        '�Ƿ������ڽ������ý�����ָ����������
    supportҽ���ӿڴ�ӡƱ�� = 46    'HIS��ֻ��Ʊ�ݺŵ�������ӡ��ҽ���ӿ�(����)�д�ӡ
    support�൥��һ�ν��� = 47      '�൥��Ԥ����ʱ��ҽ���ӿڽ������һ�ε���ʱ���ؽ�������HIS���ٷ�̯��ÿ�ŵ�����
        
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    support�������_�������ú���ýӿ� = 49          '�Ƿ��ڽ������ú�ŵ��������������ӿ�,�˲���Ϊtrue����ʾ�����������ʱ���н����������ã������ú�����������ӿ�
    supportʵʱ��� = 60             '�Ƿ����÷���ʵʱ���
    support�˷Ѻ��ӡ�ص� = 65   'ҽ�������Ƿ��˷Ѻ��ӡ�ص�:����
    support�������Ϻ��ӡ�ص� = 66      '�������Ϻ��ӡ�ص�
    support������;���� = 84   'ҽ���Ƿ�֧��������;����:81661
    support����һ�ν���סԺ���� = 88  '����סԺ���˶Զ��סԺ���ý���һ�ν���,�����:114915
End Enum

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
Private mlng���ű���ƽ������ As Long
Public gobjSquare As SquareCard  '�����㲿��  42301
Public gobjPlugIn As Object
Public gobjPublicDrug As Object 'ҩƷ��������,105875
Public gobjPublicExpense As Object  '���ù�������
Public gobjPublicExpenseBillOperation As Object
Public gintPriceGradeStartType As Integer
Public gstrҩƷ�۸�ȼ� As String
Public gstr���ļ۸�ȼ� As String
Public gstr��ͨ�۸�ȼ� As String
Public gobjCharge As Object '������ò��� zl9OutExse.clsOutExse

Public glngInstanceCount As Long '��������

Public Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str������ As String, ByVal str�������� As String, _
    ByVal intMode As Integer, ByVal intPrice As Integer, Optional ByVal lngRow As Long) As ADODB.Recordset
'���ܣ����ݵ��ݶ������ݴ���һ����ϸ��¼����Ϣ(���ۼ۵�λ)
'�ֶΣ�����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������,ִ�п���ID,
'          �������ʣ�1-�շѵ�,2-���ʵ�),�Ƿ񻮼�(1-����;0-�������շѼ����ʵ�)
'������intPage=ָ���ĵ���,lngRow=ָ�����У���ָ��ʱ�������е��ݵ�������
'          intMode:�������ʣ�1-�շѵ�,2-���ʵ�)
'          intPrice:�Ƿ񻮼�(1-����;0-�������շѼ����ʵ�)

    Dim i As Integer, j As Integer
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl���� As Double, curʵ�� As Currency
    Dim rsTmp As New ADODB.Recordset
    
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
    
    
    If lngRow = 0 Then
        intB = 1
        intE = objBill.Details.Count
    Else
        intB = lngRow
        intE = lngRow
    End If
    
    For i = intB To intE
        dbl���� = 0: curʵ�� = 0
        With objBill.Details(i)
            If lngRow = 0 Then
                If .����ID = 0 Then
                    rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID
                Else    '���ʱ�
                    rsTmp.Filter = "����ID=" & .����ID & " And �շ�ϸĿID=" & .�շ�ϸĿID
                End If
                blnNew = rsTmp.RecordCount = 0
            Else
                blnNew = True
            End If
                            
            If blnNew Then
                rsTmp.AddNew
                
                If .����ID = 0 Then
                    rsTmp!����ID = objBill.����ID
                    rsTmp!��ҳID = objBill.��ҳID
                Else    '���ʱ�
                    rsTmp!����ID = .����ID
                    rsTmp!��ҳID = .��ҳID
                End If
                
                rsTmp!�շ���� = .�շ����
                rsTmp!�շ�ϸĿID = .�շ�ϸĿID
                rsTmp!ִ�п���ID = .ִ�в���ID
                rsTmp!�������� = intMode
                rsTmp!�Ƿ񻮼� = intPrice
                
                For j = 1 To .InComes.Count
                    dbl���� = dbl���� + .InComes(j).��׼����
                    curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                Next
                If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                    '��ҩ����λת��Ϊ�ۼ۵�λ
                    rsTmp!���� = IIf(.���� = 0, 1, .����) * .���� * .Detail.סԺ��װ
                    rsTmp!���� = Format(dbl���� / .Detail.סԺ��װ, gstrFeePrecisionFmt)
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
                If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                    '��ҩ����λת��Ϊ�ۼ۵�λ
                    rsTmp!���� = rsTmp!���� + IIf(.���� = 0, 1, .����) * .���� * .Detail.סԺ��װ
                    rsTmp!���� = Format((rsTmp!���� + Format(dbl���� / .Detail.סԺ��װ, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                Else
                    rsTmp!���� = rsTmp!���� + IIf(.���� = 0, 1, .����) * .����
                    rsTmp!���� = Format((rsTmp!���� + Format(dbl����, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                End If
                rsTmp!ʵ�ս�� = rsTmp!ʵ�ս�� + Format(curʵ��, gstrDec)
            End If
            
            rsTmp.Update
        End With
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
End Function

Public Function GetVBalance(ByVal bytҵ������ As Byte, strPrivs As String, int���� As Integer, lng����ID As Long, Optional strTime As String, _
     Optional DateBegin As Date, Optional DateEnd As Date, Optional blnOnlyYbUpData As Boolean, _
     Optional blnDateMoved As Boolean, Optional bytBaby As Byte, Optional blnOnly���� As Boolean, Optional bytKind As Byte, _
     Optional strItem As String, Optional strDeptIDs As String, Optional strClass As String, Optional strChargeType As String = "") As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ������δ����ϸĿ��ϸ(���շ�ϸĿ)
    '��Σ�lng����ID-����ID,
    '      strTime�� ҽ������ֻ������סԺ�����ͷ����ڼ� [strTime=סԺ������,"0,1,2,3",0��ʾ����]
    '      DateBegin,DateEnd�� ���ʷ����ڼ�,���Ǽ�ʱ�����ʱ��,ȱʡֵΪCDate("0:00:00")
    '      blnOnlyYbUpData���Ƿ�ֻ�������ϴ�����
    '      blnDateMoved�����˵Ǽ�ʱ���Ƿ���ת������֮ǰ
    '      bytBaby��0-���з���,1-���˷���,2�Լ���-��bytBaby-1��Ӥ������]
    '      blnOnly�����������ʷ���
    '      bytKind��0-����ͨ����,1-��������,2-��ͨ���ú�������
    '      strItem:�վݷ�Ŀ��,'��ҩ��','��ҩ��',...
    '      strDeptIds����������ID��,"1,2,3",�ձ�ʾ����
    '      strClass�����з���(��δ����),"'����1','����2',..."
    '      bytҵ������-0-����ҵ��;1-סԺҵ��
    '���Σ�
    '���أ��ɹ�=��¼��,ʧ��=Nothing
    '���ƣ����˺�
    '���ڣ�2010-03-06 10:39:50
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCond As String, blnRelation As Boolean
    Dim strTable As String, bytType As Byte '0-����,1-סԺ,2-�����סԺ
    Dim strWherePage As String 'סԺ��������
    Dim strWhereMzPage As String
    On Error GoTo errH
    
    strPrivs = ";" & strPrivs & ";"
    'Modified by ZYB 2002-10-30
    blnRelation = gclsInsure.GetCapability(support��������ҽ����Ŀ, lng����ID, int����)
     
    strCond = " And A.����ID=[1]"
    strWherePage = IIf(strTime = "", "", " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0")
    strWhereMzPage = IIf(strTime = "", "", " And Instr([2],',0,')>0")   '36004
    
    If DateBegin <> CDate("0:00:00") Then
        strCond = strCond & " And " & IIf(gint����ʱ�� = 0, "A.�Ǽ�ʱ��", "A.����ʱ��") & " Between [3] And [4]"
        DateBegin = CDate(Format(DateBegin, "yyyy-MM-dd 00:00:00"))
        DateEnd = CDate(Format(DateEnd, "yyyy-MM-dd 23:59:59"))
    End If
 
    '���˺�:2010-03-06 11:23:52: Or A.����ID is Not NULL ����������ѽ��ʵ���ϸ,���ҵķ�������,�Ǵ��,����û��˵ҽ��������,����ݲ�����!
    If blnOnlyYbUpData Then strCond = strCond & " And (A.�Ƿ��ϴ�=1 And A.����ID is Null Or A.����ID is Not NULL)"
    
    strCond = strCond & IIf(bytBaby = 0, "", IIf(bytBaby = 1, " And Nvl(A.Ӥ����,0)=0", " And A.Ӥ����=[6]"))
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.�վݷ�Ŀ||''',')>0")
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([8],','||A.��������ID||',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([10],','''||A.�շ����||''',')>0")   '34260
    
    If bytKind = 1 Then
        strCond = strCond & " And A.�����־=4"
    Else
        If InStr(strPrivs, ";סԺ���ý���;") = 0 Or blnOnly���� Then strCond = strCond & " And A.�����־<>2"
        If InStr(strPrivs, ";������ý���;") = 0 Then strCond = strCond & " And A.�����־<>1"
        If bytKind = 0 Then strCond = strCond & " And A.�����־<>4"
    End If
    If bytҵ������ = 0 Or bytKind = 1 Then
        bytType = 0
    Else
        bytType = 1 '42027
    End If
''    '��ȡ���û�ȡ��Χ����
''    If bytKind = 1 Then '��������
''        bytType = 0
''    ElseIf (InStr(strPrivs, ";סԺ���ý���;") = 0 Or blnOnly����) Then  '���ﲿ�ֵĴ���
''            If InStr(strPrivs, ";������ý���;") = 0 Then
''                '��Ȩ��,�ִ�������������ݵ�:
''                ' a: 3-����(���￨�ȶ�����շ�);4-���
''                bytType = IIf(bytKind = 0, 1, 0) '����Ǿ��￨,�Ͷ�סԺ���ü�¼,�����������ü�¼
''            Else
''                '���������Ȩ��
''                'a: 1-����,3-����(���￨�ȶ�����շ�);4-���
''                bytType = IIf(bytKind = 0, 2, 0)
''            End If
''    ElseIf InStr(strPrivs, ";������ý���;") = 0 Then    'סԺ����,�����ܽ��������
''        '2-סԺ;3-����(���￨�ȶ�����շ�);4-���
''        bytType = IIf(bytKind = 0, 1, 2)
''    Else  '�����סԺ
''        '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
''        bytType = 2
''    End If
    '����Ҫ�󣺼�¼����,��¼״̬,NO����š��շ�����շ�ϸĿID,�շ����ơ����㵥λ���������š���񡢲��ء��������۸񡢽�ҽ��,
    '          ����ʱ��,�Ǽ�ʱ��,Ӥ����,ҽ����Ŀ���롢���մ���ID��������Ŀ���Ƿ��ϴ�,�Ƿ���
    'ע�⣺���ڽ���ֻ������б�����Ŀ�����,�������뱣��֧����Ŀ����ʱ����(+)
    '   ����Ϊ��ָ���ѱ����������Ĵ��۳����¼,���൥����ϸ���ϴ�
    
    '��ʱ���ģ��������Ϻ��SQL����,������ʱ��Ϊ"���/����"
    If blnOnly���� Then
        '�������
        '�������ݺ�,ֻ�������շ�ʱȡ���۵����õ����ݺ�
        'һ��һ���ĳ������ò���,��Ȼ�����ڵ��ʲ��ֽ��ʵ����,����ȻҪ��sum(ʵ�ս��)-sum(���ʽ��),��Ϊ�������ϲ����ļ�¼û��ʵ�ս��
        If bytType = 2 Then
            strTable = "" & _
            "       Select NO, ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, ժҪ, �Ƿ���, ��������id, ִ�в���id," & vbNewLine & _
            "              Avg(Nvl(����, 0) * ����) As ����, Avg(��׼����) As ����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ʵ�ս��," & vbNewLine & _
            "              Sum(ͳ����) As ͳ����,��������" & vbNewLine & _
            "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A") & vbNewLine & _
            "       Where ��¼״̬ <> 0 And ���ʷ��� = 1 And A.���� <> 0 " & strCond & strWhereMzPage & vbNewLine & _
            "       Group By NO, Mod(��¼����, 10), ��¼״̬, Nvl(�۸񸸺�, ���), ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, ժҪ, �Ƿ���," & vbNewLine & _
            "                ��������id, ִ�в���id,��������" & vbNewLine & _
            "       Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0 " & _
            "       UNION ALL " & vbCrLf & _
            "       Select NO, ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, ժҪ, �Ƿ���, ��������id, ִ�в���id," & vbNewLine & _
            "              Avg(Nvl(����, 0) * ����) As ����, Avg(��׼����) As ����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ʵ�ս��," & vbNewLine & _
            "              Sum(ͳ����) As ͳ����,��������" & vbNewLine & _
            "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & vbNewLine & _
            "       Where ��¼״̬ <> 0 And ���ʷ��� = 1 And A.���� <> 0 " & strCond & strWherePage & vbNewLine & _
            "       Group By NO, Mod(��¼����, 10), ��¼״̬, Nvl(�۸񸸺�, ���), ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, ժҪ, �Ƿ���," & vbNewLine & _
            "                ��������id, ִ�в���id,��������" & vbNewLine & _
            "       Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0 "
            
        Else
            If bytType = 0 Then
                strTable = IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A")
            Else
                strTable = IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A")
            End If
            
            strTable = "" & _
            "       Select NO, ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, ժҪ, �Ƿ���, ��������id, ִ�в���id," & vbNewLine & _
            "              Avg(Nvl(����, 0) * ����) As ����, Avg(��׼����) As ����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ʵ�ս��," & vbNewLine & _
            "              Sum(ͳ����) As ͳ����,��������" & vbNewLine & _
            "       From " & strTable & vbNewLine & _
            "       Where ��¼״̬ <> 0 And ���ʷ��� = 1 And A.���� <> 0" & IIf(bytType = 1, " And A.��ҳID Is Not Null ", "") & strCond & IIf(bytType = 0, strWhereMzPage, strWherePage) & vbNewLine & _
            "       Group By NO, Mod(��¼����, 10), ��¼״̬, Nvl(�۸񸸺�, ���), ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, ժҪ, �Ƿ���," & vbNewLine & _
            "                ��������id, ִ�в���id,��������" & vbNewLine & _
            "       Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0 "
            
            If bytType = 0 Then
                strTable = strTable & " Union " & _
                "       Select NO, ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, ժҪ, �Ƿ���, ��������id, ִ�в���id," & vbNewLine & _
                "              Avg(Nvl(����, 0) * ����) As ����, Avg(��׼����) As ����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ʵ�ս��," & vbNewLine & _
                "              Sum(ͳ����) As ͳ����,��������" & vbNewLine & _
                "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & vbNewLine & _
                "       Where ��¼״̬ <> 0 And ���ʷ��� = 1 And A.���� <> 0 And (Mod(A.��¼����,10) = 5 And A.��ҳID Is Null)" & strCond & vbNewLine & _
                "       Group By NO, Mod(��¼����, 10), ��¼״̬, Nvl(�۸񸸺�, ���), ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, ժҪ, �Ƿ���," & vbNewLine & _
                "                ��������id, ִ�в���id,��������" & vbNewLine & _
                "       Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0 "
            End If
        End If
        
        strSQL = "" & _
        " Select Sysdate As ����ʱ��, A.����id, A.�շ����, A.�վݷ�Ŀ, A.���㵥λ, A.�շ�ϸĿid," & vbNewLine & _
        "       B.����id ����֧������id, B.�Ƿ�ҽ�� �Ƿ�ҽ��, B.��Ŀ���� ���ձ���, Sum(A.����) As ����, Avg(A.����) As ����," & vbNewLine & _
        "       Sum(A.ʵ�ս��) As ʵ�ս��, Sum(A.ͳ����) As ͳ����, Max(A.ժҪ) ժҪ, Max(A.�Ƿ���) �Ƿ���," & vbNewLine & _
        "       Max(A.��������id) ��������id, Max(A.ִ�в���id) ִ�в���id, Max(A.������) ������,Max(A.��������) ��������" & vbNewLine & _
        " From ( " & strTable & ") A, ����֧����Ŀ B, �շ���ĿĿ¼ C " & vbNewLine & _
        " Where A.�շ�ϸĿid = C.ID And A.�շ�ϸĿid = B.�շ�ϸĿid" & IIf(blnRelation, "(+)", "") & " And B.����" & IIf(blnRelation, "(+)", "") & " = [5] " & vbNewLine & _
                    IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
        " Group By A.�շ�ϸĿid, A.����id, A.�շ����, A.�վݷ�Ŀ, A.���㵥λ, B.����id, B.�Ƿ�ҽ��, B.��Ŀ����" & vbNewLine & _
        " Having Sum(A.ʵ�ս��) <> 0"
    Else
        'סԺ����
        If bytType = 2 Then
            strTable = ""
            If InStr(1, "," & strTime & ",", ",0,") > 0 Or strTime = "" Then  '��������:33189
                    strTable = "" & _
                    "       Select Mod(A.��¼����, 10) As ��¼����, A.��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���) As ���, A.�����־, A.����id," & vbNewLine & _
                    "              -1*NULL as ��ҳid, Nvl(A.Ӥ����, 0) As Ӥ����, A.������ As ҽ��, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ," & vbNewLine & _
                    "              A.���ձ���, Nvl(A.���մ���id, 0) As ���մ���id, Avg(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
                    "              Avg(A.��׼����) As ��׼����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ���, A.����ʱ��," & vbNewLine & _
                    "              A.�Ǽ�ʱ��, Nvl(A.�Ƿ��ϴ�, 0) As �Ƿ��ϴ�, Nvl(A.�Ƿ���, 0) As �Ƿ���,Nvl(A.������Ŀ��, 0) As ������Ŀ��, A.ժҪ,A.��������" & vbNewLine & _
                    "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A") & ", ������Ŀ B" & vbNewLine & _
                    "       Where A.��¼״̬ <> 0 And A.���ʷ��� = 1 And A.������Ŀid = B.ID And A.���� <> 0 " & strCond & strWhereMzPage & vbNewLine & _
                    "       Group By Mod(A.��¼����, 10), A.��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���), A.�����־, A.����id," & vbNewLine & _
                    "                Nvl(A.Ӥ����, 0), A.������, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ, A.���ձ���," & vbNewLine & _
                    "                Nvl(A.���մ���id, 0), A.����ʱ��, A.�Ǽ�ʱ��, Nvl(A.�Ƿ��ϴ�, 0), Nvl(A.�Ƿ���, 0), Nvl(A.������Ŀ��, 0),��A.ժҪ,A.��������" & vbNewLine & _
                    "       Having Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) <> 0��" & _
                    "       UNION ALL " & vbCrLf
            End If
            strTable = strTable & "" & _
            "       Select Mod(A.��¼����, 10) As ��¼����, A.��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���) As ���, A.�����־, A.����id," & vbNewLine & _
            "              A.��ҳid, Nvl(A.Ӥ����, 0) As Ӥ����, A.������ As ҽ��, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ," & vbNewLine & _
            "              A.���ձ���, Nvl(A.���մ���id, 0) As ���մ���id, Avg(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
            "              Avg(A.��׼����) As ��׼����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ���, A.����ʱ��," & vbNewLine & _
            "              A.�Ǽ�ʱ��, Nvl(A.�Ƿ��ϴ�, 0) As �Ƿ��ϴ�, Nvl(A.�Ƿ���, 0) As �Ƿ���,��Nvl(A.������Ŀ��, 0) As ������Ŀ��, A.ժҪ,A.��������" & vbNewLine & _
            "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " , ������Ŀ B" & vbNewLine & _
            "       Where A.��¼״̬ <> 0 And A.���ʷ��� = 1 And A.������Ŀid = B.ID And A.���� <> 0 " & strCond & strWherePage & vbNewLine & _
            "       Group By Mod(A.��¼����, 10), A.��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���), A.�����־, A.����id, A.��ҳid," & vbNewLine & _
            "                Nvl(A.Ӥ����, 0), A.������, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ, A.���ձ���," & vbNewLine & _
            "                Nvl(A.���մ���id, 0), A.����ʱ��, A.�Ǽ�ʱ��, Nvl(A.�Ƿ��ϴ�, 0), Nvl(A.�Ƿ���, 0), Nvl(A.������Ŀ��, 0),A.ժҪ,A.��������" & vbNewLine & _
            "       Having Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) <> 0"
        Else
            If bytType = 0 Then
                strTable = IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A")
            Else
                strTable = IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A")
            End If
            strTable = "" & _
            "       Select Mod(A.��¼����, 10) As ��¼����, A.��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���) As ���, A.�����־, A.����id," & vbNewLine & _
            "              " & IIf(bytType = 0, "-1*NULL", "A.��ҳid") & " as ��ҳid, Nvl(A.Ӥ����, 0) As Ӥ����, A.������ As ҽ��, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ," & vbNewLine & _
            "              A.���ձ���, Nvl(A.���մ���id, 0) As ���մ���id, Avg(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
            "              Avg(A.��׼����) As ��׼����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ���, A.����ʱ��," & vbNewLine & _
            "              A.�Ǽ�ʱ��, Nvl(A.�Ƿ��ϴ�, 0) As �Ƿ��ϴ�, Nvl(A.�Ƿ���, 0) As �Ƿ���,Nvl(A.������Ŀ��, 0) As ������Ŀ��, A.ժҪ,A.��������" & vbNewLine & _
            "       From " & strTable & " , ������Ŀ B" & vbNewLine & _
            "       Where A.��¼״̬ <> 0 And A.���ʷ��� = 1" & IIf(bytType = 1, " And A.��ҳID Is Not Null ", "") & " And A.������Ŀid = B.ID And A.���� <> 0 " & strCond & IIf(bytType = 0, strWhereMzPage, strWherePage) & vbNewLine & _
            "       Group By Mod(A.��¼����, 10), A.��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���), A.�����־, A.����id" & IIf(bytType = 0, "", ", A.��ҳid") & "," & vbNewLine & _
            "                Nvl(A.Ӥ����, 0), A.������, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ, A.���ձ���," & vbNewLine & _
            "                Nvl(A.���մ���id, 0), A.����ʱ��, A.�Ǽ�ʱ��, Nvl(A.�Ƿ��ϴ�, 0), Nvl(A.�Ƿ���, 0), Nvl(A.������Ŀ��, 0),��A.ժҪ,A.��������" & vbNewLine & _
            "       Having Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) <> 0��"
            
            If bytType = 0 Then
                strTable = strTable & " Union " & _
                "       Select Mod(A.��¼����, 10) As ��¼����, A.��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���) As ���, A.�����־, A.����id," & vbNewLine & _
                "              " & IIf(bytType = 0, "-1*NULL", "A.��ҳid") & " as ��ҳid, Nvl(A.Ӥ����, 0) As Ӥ����, A.������ As ҽ��, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ," & vbNewLine & _
                "              A.���ձ���, Nvl(A.���մ���id, 0) As ���մ���id, Avg(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
                "              Avg(A.��׼����) As ��׼����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ���, A.����ʱ��," & vbNewLine & _
                "              A.�Ǽ�ʱ��, Nvl(A.�Ƿ��ϴ�, 0) As �Ƿ��ϴ�, Nvl(A.�Ƿ���, 0) As �Ƿ���,Nvl(A.������Ŀ��, 0) As ������Ŀ��, A.ժҪ,A.��������" & vbNewLine & _
                "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " , ������Ŀ B" & vbNewLine & _
                "       Where A.��¼״̬ <> 0 And A.���ʷ��� = 1 And (Mod(A.��¼����,10) = 5 And A.��ҳID Is Null) And A.������Ŀid = B.ID And A.���� <> 0 " & strCond & vbNewLine & _
                "       Group By Mod(A.��¼����, 10), A.��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���), A.�����־, A.����id" & IIf(bytType = 0, "", ", A.��ҳid") & "," & vbNewLine & _
                "                Nvl(A.Ӥ����, 0), A.������, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ, A.���ձ���," & vbNewLine & _
                "                Nvl(A.���մ���id, 0), A.����ʱ��, A.�Ǽ�ʱ��, Nvl(A.�Ƿ��ϴ�, 0), Nvl(A.�Ƿ���, 0), Nvl(A.������Ŀ��, 0),��A.ժҪ,A.��������" & vbNewLine & _
                "       Having Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) <> 0��"
            End If
        End If
        
        strSQL = "Select A.��¼����, A.��¼״̬, A.NO, A.���, A.�����־, A.����id, A.��ҳid, A.Ӥ����, C.��Ŀ���� As ҽ����Ŀ����," & vbNewLine & _
                "       A.���ձ���, A.���մ���id, A.�շ����, A.�շ�ϸĿid, Nvl(E.����, B.����) As �շ�����, A.���㵥λ," & vbNewLine & _
                "       X.���� As ��������, B.���, B.����, A.����, A.��׼���� As �۸�, A.���," & vbNewLine & _
                "       A.ҽ��, A.����ʱ��, A.�Ǽ�ʱ��, A.�Ƿ��ϴ�, A.�Ƿ���, A.������Ŀ��, A.ժҪ,A.��������" & vbNewLine & _
                "From ( " & strTable & ") A, �շ���ĿĿ¼ B, ����֧����Ŀ C, �շ���Ŀ���� E,���ű� X" & vbNewLine & _
                "Where A.�շ�ϸĿid = B.ID And B.ID = C.�շ�ϸĿid" & IIf(blnRelation, "(+)", "") & " And C.����" & IIf(blnRelation, "(+)", "") & " = [5] And A.��������id = X.ID " & vbNewLine & _
                        IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.��������,Nvl(B.��������,'��'))||''',')>0") & _
                "      And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = " & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1)
    End If
    Set GetVBalance = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, "," & strTime & ",", DateBegin, DateEnd, int����, bytBaby - 1, "," & strItem & ",", "," & strDeptIDs & ",", "," & strClass & ",", "," & strChargeType & ",")
    '����:strDeptIDs:42478
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetBalance(ByVal bytҵ������ As Byte, strPrivs As String, lng����ID As Long, Optional strTime As String, Optional strDeptIDs As String, _
    Optional strClass As String, Optional DateBegin As Date, Optional DateEnd As Date, Optional bytBaby As Byte, Optional strItem As String, _
    Optional blnOnlyYbUpData As Boolean, Optional ByVal blnZero As Boolean, Optional blnDateMoved As Boolean, _
    Optional blnOnly���� As Boolean, Optional bytKind As Byte, _
    Optional bln���ѿ����� As Boolean = False, Optional strChargeType As String = "") As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ������δ���ʽ��ϸĿ(��ÿ������Ŀ��)
    '��Σ�lng����ID-����ID,
    '      strTime��סԺ������,"0,1,2,3",0��ʾ����
    '      strDeptIds����������ID��,"1,2,3",�ձ�ʾ����
    '      strClass��""-���з���(��δ����),"'����1','����2',..."
    '      strItem���վݷ�Ŀ��,'��ҩ��','��ҩ��',...
    '      bytBaby��0-���з���,1-���˷���,2�Լ���-��bytBaby-1��Ӥ������
    '      DateBegin,DateEnd�����ʷ����ڼ�,���Ǽ�ʱ�����ʱ��,ȱʡֵΪCDate("0:00:00")
    '      blnOnlyYbUpData���Ƿ�ֻ�������ϴ�����
    '      blnZero���Ƿ��ȡ�����
    '      blnDateMoved�����˵Ǽ�ʱ���Ƿ���ת������֮ǰ
    '      blnOnly�����������ʷ���
    '      bytKind��  0-����ͨ����,1-��������,2-��ͨ���ú�������
    '      bln���ѿ�����-���������ѿ�(��Ҫ����һЩ�ֶ�:A.)
    '     strChargeType:""��ʾ���з���,����Ϊָ���շ����ķ���;��:5,6,7��  '34260
    '     bytҵ������-0-����ҵ��;1-סԺҵ��
    '���Σ�
    '���أ��ɹ�=��¼��,ʧ��=Nothing
    '���ƣ����˺�
    '���ڣ�2010-03-06 13:21:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, strCond As String, strCond2 As String
    Dim strTable As String, bytType As Byte '0-����,1-סԺ,2-�����סԺ
    Dim strWherePage As String 'סԺ��������
    Dim strWhereMzPage As String
        
    strCond = " And A.����ID=[1]"
    
    strWherePage = IIf(strTime = "", "", " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0")
    strWhereMzPage = IIf(strTime = "", "", " And Instr([2],',0,')>0")   '����
    
    If Not DateBegin = CDate("0:00:00") Then
        strCond = strCond & " And " & IIf(gint����ʱ�� = 0, "A.�Ǽ�ʱ��", "A.����ʱ��") & " Between [3] And [4]"
        DateBegin = CDate(Format(DateBegin, "yyyy-MM-dd 00:00:00"))
        DateEnd = CDate(Format(DateEnd, "yyyy-MM-dd 23:59:59"))
    End If
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([5],','||A.��������ID||',')>0")
    strCond = strCond & IIf(bytBaby = 0, "", IIf(bytBaby = 1, " And Nvl(A.Ӥ����,0)=0", " And A.Ӥ����=[6]"))
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.�վݷ�Ŀ||''',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([9],','''||A.�շ����||''',')>0")   '34260
    
    If bytKind = 1 Then
        strCond = strCond & " And A.�����־=4"
    Else
        If InStr(strPrivs, ";סԺ���ý���;") = 0 Or blnOnly���� Then strCond = strCond & " And A.�����־<>2"
        If InStr(strPrivs, ";������ý���;") = 0 Then strCond = strCond & " And A.�����־<>1"
        If bytKind = 0 Then strCond = strCond & " And A.�����־<>4"
    End If
        
    
    strCond2 = strCond   '�Ѿ�����ʵ�,�����Ƿ��ϴ���Ҫȡ,�����Ȱ����������¼����,�ڶ����Ӳ�ѯ��
    If blnOnlyYbUpData Then
        strCond = strCond & " And (A.�Ƿ��ϴ�=1 And A.����ID is Null) "
    Else
        strCond = strCond & " And A.����ID Is Null "
    End If
    
    ' bytKind:0-����ͨ����,1-��������,2-��ͨ���ú�������
    If bytҵ������ = 0 Or bytKind = 1 Then
        bytType = 0
    Else
         bytType = 1
    End If

''    '��ȡ���û�ȡ��Χ����
''    If bytKind = 1 Then '��������
''        bytType = 0
''    ElseIf (InStr(strPrivs, "סԺ���ý���") = 0 Or blnOnly����) Then  '���ﲿ�ֵĴ���
''            If InStr(strPrivs, "������ý���") = 0 Then
''                '��Ȩ��,�ִ�������������ݵ�:
''                ' a: 3-����(���￨�ȶ�����շ�);4-���
''                bytType = IIf(bytKind = 0, 1, 0) '����Ǿ��￨,�Ͷ�סԺ���ü�¼,�����������ü�¼
''            Else
''                '���������Ȩ��
''                'a: 1-����,3-����(���￨�ȶ�����շ�);4-���
''                bytType = IIf(bytKind = 0, 2, 0)
''            End If
''    ElseIf InStr(strPrivs, "������ý���") = 0 Then    'סԺ����,�����ܽ��������
''        '2-סԺ;3-����(���￨�ȶ�����շ�);4-���
''        bytType = IIf(bytKind = 0, 1, 2)
''    Else  '�����סԺ
''        '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
''        bytType = 2
''    End If
    
    
    'סԺ,����,ʱ��,[���ݺ�],��Ŀ,��Ŀ,Ӥ����,[ID],[���],[��¼����],[��¼״̬],[ִ��״̬],[A.��ҳID],[A.��������ID],[�Ǽ�ʱ��],δ����,���ʽ��,[����]
    If blnZero Then
        If bytType = 2 Then
            strTable = "" & _
            " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,1 as ��־,'����' as סԺ," & _
            "                -1*NULL as ��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
            "                Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,Nvl(A.ʵ�ս��,0) as δ����,��������, A.�շ����" & _
                     IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.���� * A.���� As ����, A.��׼����, A.ͳ����, A.���մ���id", "") & _
            " From ������ü�¼ A " & _
            " Where A.��¼״̬<>0 And A.���ʷ���=1" & strCond & strWhereMzPage & _
            " Union all " & _
            " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,2 as ��־,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��') as סԺ," & _
            "                A.��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
            "                Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,Nvl(A.ʵ�ս��,0) as δ����,��������, A.�շ����" & _
                     IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.���� * A.���� As ����, A.��׼����, A.ͳ����, A.���մ���id", "") & _
            " From סԺ���ü�¼ A " & _
            " Where A.��¼״̬<>0 And A.���ʷ���=1" & strCond & strWherePage & _
            ""
        Else
            If bytType = 0 Then
                    strTable = " From ������ü�¼ A Where A.��¼״̬<>0 And A.���ʷ���=1" & strCond & strWhereMzPage
            Else
                    strTable = " From סԺ���ü�¼ A Where A.��¼״̬<>0 And A.���ʷ���=1" & strCond & strWherePage
            End If
            strTable = "" & _
            " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬," & IIf(bytType = 0, "1 as ��־,'����'", "2 as ��־,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��')") & " as סԺ," & _
            "               " & IIf(bytType = 0, " -1*NULL", "A.��ҳID") & " as ��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
            "                Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,Nvl(A.ʵ�ս��,0) as δ����,��������, A.�շ����" & _
                     IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.���� * A.���� As ����, A.��׼����, A.ͳ����, A.���մ���id", "") & _
            " " & strTable & IIf(bytType = 1, " And A.��ҳID Is Not Null ", "")
            If bytType = 0 Then
                strTable = strTable & " Union " & _
                " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬," & IIf(bytType = 0, "1 as ��־,'����'", "2 as ��־,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��')") & " as סԺ," & _
                "               " & IIf(bytType = 0, " -1*NULL", "A.��ҳID") & " as ��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
                "                Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,Nvl(A.ʵ�ս��,0) as δ����,��������, A.�շ����" & _
                         IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.���� * A.���� As ����, A.��׼����, A.ͳ����, A.���մ���id", "") & _
                " From סԺ���ü�¼ A Where A.��¼״̬<>0 And (Mod(A.��¼����,10) = 5 And A.��ҳID Is Null) And A.���ʷ���=1" & strCond
            End If
        End If
    Else
    
        '���Ӳ�ѯ���ڹ��˵���һ�ν���ʱһ��һ���ķ���
        '��������ʱ,��ʹһ��һ��,ҲҪ�ó�������
        If bytType = 2 Then
            strTable = ""
            If InStr(1, "," & strTime & ",", ",0,") > 0 Or strTime = "" Then  '��������
                strTable = "" & _
                " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,1 as ��־,'����'  as סԺ," & _
                "               -1*NULL as ��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
                "                Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,Nvl(A.ʵ�ս��,0) as δ����,��������, A.�շ����" & _
                             IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.���� * A.���� As ����, A.��׼����, A.ͳ����, A.���մ���id", "") & _
                " From  ������ü�¼ A," & _
                "      ( Select A.NO,A.���,A.��¼����,Nvl(Sum(A.ʵ�ս��),0) as ʵ�ս��" & _
                "        From  ������ü�¼ A" & _
                "        Where A.��¼״̬<>0  And A.���ʷ���=1 And Nvl(A.ʵ�ս��,0)<>0" & strCond & _
                "        Group by A.NO,A.���,A.��¼���� Having Nvl(Sum(A.ʵ�ս��),0)<>0 ) B " & _
                " Where A.NO=B.NO And A.���=B.��� And A.��¼����=B.��¼���� And A.����ID Is Null" & strCond & _
                " Union ALL "
            End If
            strTable = strTable & "" & _
            " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,2 as ��־,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��') as סԺ," & _
            "         A.��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
            "         Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,Nvl(A.ʵ�ս��,0) as δ����,��������, A.�շ����" & _
                      IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.���� * A.���� As ����, A.��׼����, A.ͳ����, A.���մ���id", "") & _
            " From  סԺ���ü�¼ A," & _
            "      ( Select A.NO,A.���,A.��¼����,Nvl(Sum(A.ʵ�ս��),0) as ʵ�ս��" & _
            "        From  סԺ���ü�¼ A" & _
            "        Where A.��¼״̬<>0  And A.���ʷ���=1 And Nvl(A.ʵ�ս��,0)<>0" & strCond & strWherePage & _
            "        Group by A.NO,A.���,A.��¼���� Having Nvl(Sum(A.ʵ�ս��),0)<>0 ) B " & _
            " Where A.NO=B.NO And A.���=B.��� And A.��¼����=B.��¼���� And A.����ID Is Null" & strCond & strWherePage
            
        Else
            strTable = "" & _
            " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬," & IIf(bytType = 0, "1 as ��־,'����'", "2 as ��־,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��')") & " as סԺ," & _
            "                " & IIf(bytType = 0, "-1*NULL", "A.��ҳID") & " as ��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
            "                Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,Nvl(A.ʵ�ս��,0) as δ����,��������, A.�շ����" & _
                         IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.���� * A.���� As ����, A.��׼����, A.ͳ����, A.���մ���id", "") & _
            " From  " & IIf(bytType = 0, "������ü�¼ A", " סԺ���ü�¼ A") & "," & _
            "      ( Select A.NO,A.���,A.��¼����,Nvl(Sum(A.ʵ�ս��),0) as ʵ�ս��" & _
            "        From " & IIf(bytType = 0, "������ü�¼ A", " סԺ���ü�¼ A") & _
            "        Where A.��¼״̬<>0  And A.���ʷ���=1" & IIf(bytType = 1, " And A.��ҳID Is Not Null ", "") & " And Nvl(A.ʵ�ս��,0)<>0" & strCond & IIf(bytType = 0, strWhereMzPage, strWherePage) & _
            "        Group by A.NO,A.���,A.��¼���� Having Nvl(Sum(A.ʵ�ս��),0)<>0 ) B " & _
            " Where A.NO=B.NO And A.���=B.��� And A.��¼����=B.��¼���� And A.����ID Is Null" & strCond & IIf(bytType = 0, strWhereMzPage, strWherePage)
            
            If bytType = 0 Then
                strTable = strTable & " Union " & _
                " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬," & IIf(bytType = 0, "1 as ��־,'����'", "2 as ��־,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��')") & " as סԺ," & _
                "                " & IIf(bytType = 0, "-1*NULL", "A.��ҳID") & " as ��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
                "                Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,Nvl(A.ʵ�ս��,0) as δ����,��������, A.�շ����" & _
                             IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.���� * A.���� As ����, A.��׼����, A.ͳ����, A.���մ���id", "") & _
                " From   סԺ���ü�¼ A" & "," & _
                "      ( Select A.NO,A.���,A.��¼����,Nvl(Sum(A.ʵ�ս��),0) as ʵ�ս��" & _
                "        From סԺ���ü�¼ A" & _
                "        Where A.��¼״̬<>0 And (Mod(A.��¼����,10) = 5 And A.��ҳID Is Null)  And A.���ʷ���=1 And Nvl(A.ʵ�ս��,0)<>0" & strCond & _
                "        Group by A.NO,A.���,A.��¼���� Having Nvl(Sum(A.ʵ�ս��),0)<>0 ) B " & _
                " Where A.NO=B.NO And A.���=B.��� And A.��¼����=B.��¼���� And A.����ID Is Null" & strCond
            End If
        End If
    End If
    
    If bytType = 2 Then
        strSQL = ""
        If InStr(1, "," & strTime & ",", ",0,") > 0 Or strTime = "" Then  '��������
            strSQL = "" & _
            "        SELECT 0 as ID,A.NO,A.���,Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.ִ��״̬,1 as ��־," & _
            "                '����'  as סԺ,-1*NULL as ��ҳID," & _
            "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
            "               Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
            "               Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ����,A.�������� , max(A.�շ����) as �շ����" & _
                            IIf(bln���ѿ�����, ",max(A.�ѱ�) as �ѱ�, max(A.ִ�в���id) as ִ�в���id, max(A.������) as ������,avg(A.���� * nvl(A.����,1)) As ����, avg(A.��׼����) as ��׼����, avg(A.ͳ����) as ͳ����,max( A.���մ���id) as ���մ���id ", "") & _
            "        FROM " & IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A") & " " & _
            "        Where A.����id Is Not Null And A.��¼״̬<>0 And A.���ʷ���=1 And (Nvl(A.ʵ�ս��, 0)<>Nvl(A.���ʽ��, 0) Or Nvl(A.���ʽ��, 0)=0)" & strCond2 & strWhereMzPage & _
            "        Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0  " & _
            "                    Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and Sum(Nvl(A.���ʽ��,0)) =0 And Mod(Count(*),2)=0) " & _
            "                    Or Sum(Nvl(A.���ʽ��, 0))=0 And Sum(Nvl(A.Ӧ�ս��,0))<>0 And Mod(Count(*),2)=0" & _
            "        Group by A.NO,A.���,Mod(A.��¼����,10),A.��¼״̬,A.ִ��״̬," & _
            "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.�Ǽ�ʱ��,Decode(Nvl(A.Ӥ����,0),0,'','��'),A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,A.��������" & _
            "         Union all  "
        End If
        strSQL = strSQL & _
        "        SELECT 0 as ID,A.NO,A.���,Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.ִ��״̬,2 as ��־," & _
        "               Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��') as סԺ,A.��ҳID," & _
        "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
        "               Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
        "               Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ����,A.��������, max(A.�շ����) as �շ���� " & _
                        IIf(bln���ѿ�����, ",max(A.�ѱ�) as �ѱ�, max(A.ִ�в���id) as ִ�в���id, max(A.������) as ������,avg(A.���� * nvl(A.����,1)) As ����, avg(A.��׼����) as ��׼����, avg(A.ͳ����) as ͳ����,max( A.���մ���id) as ���մ���id ", "") & _
        "        FROM " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " " & _
        "        Where A.����id Is Not Null And A.��¼״̬<>0 And A.���ʷ���=1 And (Nvl(A.ʵ�ս��, 0)<>Nvl(A.���ʽ��, 0) Or Nvl(A.���ʽ��, 0)=0)" & strCond2 & strWherePage & _
        "        Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0 " & _
        "                    Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and Sum(Nvl(A.���ʽ��,0)) =0 And Mod(Count(*),2)=0) " & _
        "                    Or Sum(Nvl(A.���ʽ��, 0))=0 And Sum(Nvl(A.Ӧ�ս��,0))<>0 And Mod(Count(*),2)=0" & _
        "        Group by A.NO,A.���,Mod(A.��¼����,10),A.��¼״̬,A.ִ��״̬,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��'),A.��ҳID," & _
        "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.�Ǽ�ʱ��,Decode(Nvl(A.Ӥ����,0),0,'','��'),A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,A.��������" & _
        ""
    Else
        If bytType = 0 Then
            strSQL = IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼", 2, "", True, ""), "������ü�¼")
        Else
            strSQL = IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼", 2, "", True, ""), "סԺ���ü�¼")
        End If
        
        strSQL = "" & _
        "        SELECT 0 as ID,A.NO,A.���,Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.ִ��״̬," & _
        "               " & IIf(bytType = 0, "1 as ��־,'����'", "2 as ��־,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��')") & " as סԺ," & IIf(bytType = 0, "-1*NULL", "A.��ҳID") & " as ��ҳID," & _
        "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
        "               Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
        "               Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ����,A.��������, max(A.�շ����) as �շ���� " & _
                        IIf(bln���ѿ�����, ",max(A.�ѱ�) as �ѱ�, max(A.ִ�в���id) as ִ�в���id, max(A.������) as ������,avg(A.���� * nvl(A.����,1)) As ����, avg(A.��׼����) as ��׼����, avg(A.ͳ����) as ͳ����,max( A.���մ���id) as ���մ���id ", "") & _
        "        FROM " & strSQL & " A" & _
        "        Where A.����id Is Not Null And A.��¼״̬<>0 And A.���ʷ���=1 " & IIf(bytType = 1, " And A.��ҳID Is Not Null ", "") & " And (Nvl(A.ʵ�ս��, 0)<>Nvl(A.���ʽ��, 0) Or Nvl(A.���ʽ��, 0)=0)" & strCond2 & IIf(bytType = 0, strWhereMzPage, strWherePage) & _
        "        Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0 " & _
        "                    Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and  Sum(Nvl(A.���ʽ��,0))=0  And Mod(Count(*),2)=0) " & _
        "                     Or Sum(Nvl(A.���ʽ��, 0))=0 And Sum(Nvl(A.Ӧ�ս��,0))<>0 And Mod(Count(*),2)=0" & _
        "        Group by A.NO,A.���,Mod(A.��¼����,10),A.��¼״̬,A.ִ��״̬," & IIf(bytType = 0, "", "Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��'),A.��ҳID,") & _
        "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.�Ǽ�ʱ��,Decode(Nvl(A.Ӥ����,0),0,'','��'),A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,A.��������" & _
        ""
        If bytType = 0 Then
            strSQL = strSQL & " Union " & _
            "        SELECT 0 as ID,A.NO,A.���,Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.ִ��״̬," & _
            "               " & IIf(bytType = 0, "1 as ��־,'����'", "2 as ��־,Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��')") & " as סԺ," & IIf(bytType = 0, "-1*NULL", "A.��ҳID") & " as ��ҳID," & _
            "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
            "               Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
            "               Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ����,A.��������, max(A.�շ����) as �շ���� " & _
                            IIf(bln���ѿ�����, ",max(A.�ѱ�) as �ѱ�, max(A.ִ�в���id) as ִ�в���id, max(A.������) as ������,avg(A.���� * nvl(A.����,1)) As ����, avg(A.��׼����) as ��׼����, avg(A.ͳ����) as ͳ����,max( A.���մ���id) as ���մ���id ", "") & _
            "        FROM " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼", 2, "", True, ""), "סԺ���ü�¼") & " A" & _
            "        Where A.����id Is Not Null And A.��¼״̬<>0 And A.���ʷ���=1 And (Mod(A.��¼����,10) = 5 And A.��ҳID Is Null) And (Nvl(A.ʵ�ս��, 0)<>Nvl(A.���ʽ��, 0) Or Nvl(A.���ʽ��, 0)=0)" & strCond2 & _
            "        Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0 " & _
            "                    Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and  Sum(Nvl(A.���ʽ��,0))=0  And Mod(Count(*),2)=0) " & _
            "                     Or Sum(Nvl(A.���ʽ��, 0))=0 And Sum(Nvl(A.Ӧ�ս��,0))<>0 And Mod(Count(*),2)=0" & _
            "        Group by A.NO,A.���,Mod(A.��¼����,10),A.��¼״̬,A.ִ��״̬," & IIf(bytType = 0, "", "Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��'),A.��ҳID,") & _
            "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.�Ǽ�ʱ��,Decode(Nvl(A.Ӥ����,0),0,'','��'),A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,A.��������" & _
            ""
        End If
    End If
    
    '
    '����:48305,61527: ������ And Mod(Count(*),2)=0
    '   Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and  Sum(Nvl(A.���ʽ��,0))=0 ) " & _

    strTable = strTable & " Union ALL " & strSQL
    
    strSQL = _
        "Select A.��־,A.סԺ,Nvl(B.����,'δ֪') as ����,A.ʱ��,A.NO as ���ݺ� ,Nvl(E.����,C.����) as ��Ŀ,A.�վݷ�Ŀ as ��Ŀ, A.Ӥ����,A.ID,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,A.��ҳID,A.��������ID,A.�Ǽ�ʱ��," & _
        "       Nvl(A.δ����,0) δ����,Nvl(A.δ����,0) ���ʽ��,Nvl(A.��������,C.��������) as ����, A.�շ����" & _
                IIf(bln���ѿ�����, ",A.�ѱ�, A.ִ�в���id, A.������,A.����, A.��׼���� as �۸�, A.ͳ����, A.���մ���id,A.�շ�ϸĿID,C.���㵥λ", "") & _
        " From (  " & strTable & ") A,���ű� B,�շ���ĿĿ¼ C,������Ŀ D,�շ���Ŀ���� E " & _
        " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID And A.������ĿID=D.ID " & IIf(strClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
        "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        " Order by A.ʱ�� Desc,A.סԺ,A.NO Desc,A.��¼����,A.���"
    
    'Mod(Count(*),2)=1��Ϊ��������ۺ�ʵ�ս��Ϊ��ķ����ڽ��ʺ��Ƿ����ϻ��ٴν���
    On Error GoTo errH
    Set GetBalance = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, "," & strTime & ",", DateBegin, DateEnd, _
                    "," & strDeptIDs & ",", bytBaby - 1, "," & strItem & ",", "," & strClass & ",", "," & strChargeType & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�������(lng����ID As Long, bytType As Byte, Optional intԤ����� As Integer = 2) As Currency
'���ܣ���ȡָ�����˵�Ԥ�������
'������bytType:0-�������,1-Ԥ�����
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    strSQL = "Select Nvl(sum(�������),0) as �������,Nvl(sum(Ԥ�����),0) as Ԥ�����" & _
        " From ������� Where ����=1 And ����ID=[1]  " & IIf(intԤ����� = 0, "", " And ����=[2] ")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, intԤ�����)
    If Not rsTmp.EOF Then Get������� = IIf(bytType = 0, rsTmp!�������, rsTmp!Ԥ�����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function Chk�������(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ��жϲ����Ƿ������
'������
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Nvl(��˱�־,0) as ��˱�־" & _
        " From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    
    '49501
    If gTy_System_Para.byt������˷�ʽ = 0 Then
        Chk������� = (rsTmp!��˱�־ >= 1)
    Else
        Chk������� = (rsTmp!��˱�־ > 1)
    End If

    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function checkҽ���´��Ժҽ��(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ��жϲ����Ƿ���Ԥ��Ժ״̬,�Ҵ�����Ч�ĳ�Ժ(תԺ������)ҽ���������Ժ(��Ч��ҽ����ָ��ʼִ��ʱ����Ԥ��Ժʱ����ͬ���Ҵ����ѷ���״̬[ҽ��״̬=8])��
'������
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select a.Id" & vbNewLine & _
            "From ����ҽ����¼ a, ���˱䶯��¼ b, ������ҳ c, ������ĿĿ¼ d" & vbNewLine & _
            "Where a.����id = [1] And a.��ҳid = [2] And a.ҽ��״̬ = 8 And a.����id = b.����id And a.��ҳid = b.��ҳid And" & vbNewLine & _
            "           a.��ʼִ��ʱ�� = b.��ʼʱ��+0 And b.��ʼԭ�� = 10 And b.����id = c.����id And b.��ҳid = c.��ҳid And" & vbNewLine & _
            "           c.״̬ = 3 And d.���='Z' And d.�������� In ('5', '6', '11') And a.������Ŀid = d.Id"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    checkҽ���´��Ժҽ�� = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillingWarn(strPrivs As String, str���� As String, lng����ID As Long, str���ò��� As String, _
    rsWarn As ADODB.Recordset, cur��� As Currency, cur���ն� As Currency, _
    cur���ݽ�� As Currency, cur���� As Currency, str��� As String, _
    ByVal str����� As String, ByRef str�ѱ���� As String, _
    Optional bln�ಡ�� As Boolean, Optional ByVal blnPrice As Boolean, _
    Optional curItemMoney As Currency = 0, _
    Optional blnNotCheck��� As Boolean = False) As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:
'     str����=��������,������ʾ
'     lng����ID=���˲���ID,����ѡ�����õĲ����������ã�0��ʾû��ȷ�������������в����ı�����������
'     str���ò���=���ݲ�����ݷ��صļ��ʱ������÷���
'     rsWarn=��ǰ�������ʱ������ü�¼
'     cur���=�������,�����ۼƱ���
'     cur���ն�=���˵��շ����ķ��ö�,����ÿ�ձ���
'     cur���ݽ��=���˵���������ķ���
'     cur����=���˵������ö�,�����ۼƱ���
'     str���=��ǰҪ�������,���ڷ��౨��
'     str�����=�������,������ʾ
'     blnPrice=Ƿ��ʱ�Ƿ�����ǿ�Ʊ���Ϊ���۵�,���ڼ��ʻ򻮼���
'     curItemMoney-���ʽ��(�������<>0 ,����Ҫ�жϵ������,����������,�������û�����,������ݱ�����ʽ����):���˺�:24491
'     blnNotCheck���:���������м��(��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
'����:0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
'     5.������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
'     str�������="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
    Dim i As Integer, byt��־ As Byte
    Dim bln�ѱ��� As Boolean
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    '�����������
    If rsWarn.State = 0 Then Exit Function '20030709
    rsWarn.Filter = "���ò���='" & str���ò��� & "' And ����ID=" & lng����ID
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str���) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str����� = "" '�������ʱ,������ʾ��������
        
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־1 <> "-" And blnNotCheck��� Then Exit Function
    End If
    
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str���) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str����� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־2 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str���) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str����� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־3 <> "-" And blnNotCheck��� Then Exit Function
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
                        '����ѱ���ʽ��������:���ڱ���ֵʱ,Ԥ���ľ�ʱ,��ֻ�����һ��,��Ϊ���һ�����������
                        'Exit For
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
                            '����ѱ�����������:���ڱ���ֵʱ,Ԥ���ľ�ʱ,��ֻ�����һ��,��Ϊ���һ�����������
                            'Exit For
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
                    '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                     If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > Val(Nvl(rsWarn!����ֵ)) Then
                         'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                         '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                        If MsgBox("ע��" & vbCrLf & _
                                   "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ", �Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������" & _
                            IIf(gbytBilling = 0 And blnPrice, vbCrLf & vbCrLf & "��ʾ:�����ѡ���������,������ǰ���ݱ���Ϊ���۵�,�Ȳ��˽ɷѺ�����ˡ�", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = IIf(gbytBilling = 0, 5, 1)  '1  :����:28515
                        End If
                    Else
                        If gbytBilling = 0 And blnPrice Then
                            MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & _
                                Format(cur��� + cur���� - cur���ݽ��, "0.00") & " ����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��" & _
                                vbCrLf & vbCrLf & "��ʾ:�����ѡ�񽫵�ǰ���ݱ���Ϊ���۵�,�Ȳ��˽ɷѺ�����ˡ�", vbInformation, gstrSysName
                        Else
                            MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & " ����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If cur��� + cur���� - cur���ݽ�� < 0 Then
                    
                        '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                         If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > 0 Then
                             'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                             '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                            If MsgBox("ע��" & vbCrLf & _
                                       "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�,�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                            End If
                            Exit Function
                         End If
                         
                        byt��ʽ = 2
                        If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                            If blnPrice Then
                                If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�" & _
                                    vbCrLf & vbCrLf & IIf(gbytBilling = 0, "Ҫ����ǰ���ݱ���Ϊ���۵���", "Ҫǿ�Ʊ��滮�۵���"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    BillingWarn = 2
                                Else
                                    BillingWarn = IIf(gbytBilling = 0, 5, 1)
                                End If
                            Else
                                MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�", vbInformation, gstrSysName
                                BillingWarn = 3
                            End If
                        Else
                            If gbytBilling = 0 And blnPrice Then
                                MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���" & _
                                    vbCrLf & vbCrLf & "��ʾ:�����ѡ�񽫵�ǰ���ݱ���Ϊ���۵�,�Ȳ��˽ɷѺ�����ˡ�", vbInformation, gstrSysName
                            Else
                                MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���", vbInformation, gstrSysName
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                        '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                         If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > Val(Nvl(rsWarn!����ֵ)) Then
                             'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                             '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                            If MsgBox("ע��" & vbCrLf & _
                                       "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ", �Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                            End If
                            Exit Function
                         End If
                        
                        byt��ʽ = 1
                        If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                            If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������" & _
                                IIf(gbytBilling = 0 And blnPrice, vbCrLf & vbCrLf & "��ʾ:�����ѡ���������,������ǰ���ݱ���Ϊ���۵�,�Ȳ��˽ɷѺ�����ˡ�", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1
                            End If
                        Else
                            If gbytBilling = 0 And blnPrice Then
                                MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & _
                                    Format(cur��� + cur���� - cur���ݽ��, "0.00") & " ����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��" & _
                                    vbCrLf & vbCrLf & "��ʾ:�����ѡ�񽫵�ǰ���ݱ���Ϊ���۵�,�Ȳ��˽ɷѺ�����ˡ�", vbInformation, gstrSysName
                            Else
                                MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If cur��� + cur���� - cur���ݽ�� < 0 Then
                            '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                             If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > 0 Then
                                 'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                                 '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                                If MsgBox("ע��" & vbCrLf & _
                                           "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�,�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                    BillingWarn = 2
                                Else
                                    BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                                End If
                                Exit Function
                             End If
                             
                            byt��ʽ = 2
                            If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                                If blnPrice Then
                                    If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�" & _
                                        vbCrLf & vbCrLf & IIf(gbytBilling = 0, "Ҫ����ǰ���ݱ���Ϊ���۵���", "Ҫǿ�Ʊ��滮�۵���"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        BillingWarn = 2
                                    Else
                                        BillingWarn = IIf(gbytBilling = 0, 5, 1)
                                    End If
                                Else
                                    MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�", vbInformation, gstrSysName
                                    BillingWarn = 3
                                End If
                            Else
                                If gbytBilling = 0 And blnPrice Then
                                    MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���" & _
                                        vbCrLf & vbCrLf & "��ʾ:�����ѡ�񽫵�ǰ���ݱ���Ϊ���۵�,�Ȳ��˽ɷѺ�����ˡ�", vbInformation, gstrSysName
                                Else
                                    MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���", vbInformation, gstrSysName
                                End If
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
                    '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                     If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > Val(Nvl(rsWarn!����ֵ)) Then
                         'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                         '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                        If MsgBox("ע��" & vbCrLf & _
                                   "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�,�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
            
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        If blnPrice Then
                            If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�" & _
                                vbCrLf & vbCrLf & IIf(gbytBilling = 0, "Ҫ����ǰ���ݱ���Ϊ���۵���", "Ҫǿ�Ʊ��滮�۵���"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = IIf(gbytBilling = 0, 5, 1)
                            End If
                        Else
                            MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", vbInformation, gstrSysName
                            BillingWarn = 3
                        End If
                    Else
                        If gbytBilling = 0 And blnPrice Then
                            MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��" & _
                                vbCrLf & vbCrLf & "��ʾ:�����ѡ�񽫵�ǰ���ݱ���Ϊ���۵�,�Ȳ��˽ɷѺ�����ˡ�", vbInformation, gstrSysName
                        Else
                            MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ն� + cur���ݽ�� > rsWarn!����ֵ Then
                    '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                     If curItemMoney <> 0 And cur���ն� + cur���ݽ�� - curItemMoney < Val(Nvl(rsWarn!����ֵ)) Then
                         'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                         '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                        If MsgBox("ע��" & vbCrLf & _
                                   "    ���ˡ�" & str���� & "�� ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                     
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        If MsgBox(str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1
                        End If
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ն� + cur���ݽ�� > rsWarn!����ֵ Then
                    '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                     If curItemMoney <> 0 And cur���ն� + cur���ݽ�� - curItemMoney < Val(Nvl(rsWarn!����ֵ)) Then
                         'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                         '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                        If MsgBox("ע��" & vbCrLf & _
                                   "    ���ˡ�" & str���� & "�� ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                     
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
    If BillingWarn = 1 Or BillingWarn = 4 Or BillingWarn = 5 Then
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


Public Function GetDrugTotal(ByVal objBill As ExpenseBill, ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long) As Double
'���ܣ���ȡ������ָ��ҩƷ��ͬһҩ�����е�������
'������ lngҩ��ID-0��ʾ���뷢ҩʱ,���޶�ҩ�����
    Dim i As Integer, dblCount As Double
    
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).�շ�ϸĿID = lngҩƷID Then
            If IIf(lngҩ��ID <> 0, objBill.Details(i).ִ�в���ID = lngҩ��ID, 1 = 1) Then
                dblCount = dblCount + objBill.Details(i).���� * objBill.Details(i).����
            End If
        End If
    Next
    GetDrugTotal = dblCount
End Function

Public Function GetOriginalTotal(ByVal objBill As ExpenseBill, ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long) As Double
'���ܣ���ȡ������ָ��ҩƷ��ͬһҩ�����е�ԭʼ������
'������ lngҩ��ID-0��ʾ���뷢ҩʱ,���޶�ҩ�����
    Dim i As Integer, dblCount As Double
    
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).�շ�ϸĿID = lngҩƷID Then
            If IIf(lngҩ��ID <> 0, objBill.Details(i).ԭʼִ�в���ID = lngҩ��ID, 1 = 1) Then
                dblCount = dblCount + objBill.Details(i).ԭʼ����
            End If
        End If
    Next
    GetOriginalTotal = dblCount
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

Public Sub NurseDeposit(ByVal frmMain As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    Optional bln����˿� As Boolean = True, Optional ByVal bytPrepayType As Byte = 2)
    '����Ԥ�������
    '��Σ�
    '   bytPrepayType-Ԥ������(0-�����סԺ;1-����;2-סԺ)
    On Error GoTo errH
    If gobjPati Is Nothing Then
        Set gobjPati = CreateObject("zl9Patient.clsPatient")

    End If
    If gobjPati Is Nothing Then Exit Sub
    
    Call gobjPati.NurseDeposit(glngSys, gcnOracle, frmMain, gstrDBUser, lng����ID, lng��ҳID, bln����˿�, bytPrepayType)
    Set gobjPati = Nothing
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetMultiStock(ByVal lngҩƷID As Long, ByVal strҩ��IDs As String) As Double
'���ܣ���ȡָ��ҩ��ָ��ҩƷ���(�����۵�λ)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '��ȡ���(�����������),ҩ��������(����=0,����Ϊҩ��)����Ч��
    strSQL = _
        " Select Nvl(Sum(A.��������),0) as ��� From ҩƷ��� A" & _
        " Where (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
        " And A.����=1 And A.ҩƷID=[1] And Instr([2],','||A.�ⷿID||',')>0"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngҩƷID, "," & strҩ��IDs & ",")
    If Not rsTmp.EOF Then GetMultiStock = rsTmp!���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckDisable(objBill As ExpenseBill) As String
'���ܣ���鵥���е�ҩƷ�Ľ������
'���أ�ҩƷ���������ʾ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim i As Long, j As Long, k As Long
    Dim strGroup As String, strIDs As String
    Dim blnStop As Boolean
    
    Err = 0: On Error GoTo errH:
    For i = 1 To objBill.Details.Count
        If InStr(",5,6,7,", objBill.Details(i).�շ����) > 0 Then
            strIDs = strIDs & "," & objBill.Details(i).�շ�ϸĿID
        End If
    Next
    strIDs = Mid(strIDs, 2)
    If strIDs = "" Or UBound(Split(strIDs, ",")) < 1 Then Exit Function
    
    strSQL = _
        " Select /*+ RULE */  A.����,Count(Distinct A.��ĿID) as ������" & _
        " From ���ƻ�����Ŀ A,ҩƷ��� B," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        " Where A.��ĿID=B.ҩ��ID And B.ҩƷID  = j.Column_Value" & _
        " Having Count(Distinct A.��ĿID)>1  " & _
        "  Group by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strIDs)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strGroup = strGroup & "," & rsTmp!����
            rsTmp.MoveNext
        Next
        strGroup = Mid(strGroup, 2)
        
        For i = 0 To UBound(Split(strGroup, ","))
            strSQL = _
            "Select /*+ RULE */ Distinct C.����,C.����,D.����,D.����,D.���" & _
            " From ҩƷ��� A,������ĿĿ¼ B,���ƻ�����Ŀ C,�շ���ĿĿ¼ D," & _
            "          (Select Column_Value From Table(Cast(f_num2list([2]) As Zltools.t_Numlist ))) J " & _
            " Where A.ҩ��ID=B.ID And B.ID=C.��ĿID And A.ҩƷID=D.ID" & _
            "           And C.����=[1]" & _
            "           And A.ҩƷID=  j.Column_Value " & _
            " Order by C.����,C.����,D.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", Val(Split(strGroup, ",")(i)), strIDs)
            
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
            If blnStop Then
                CheckDisable = "���ֵ���������ҩƷ������û����ã�" & vbCrLf & strInfo & vbCrLf & "���޸Ľ���ҩƷ���ټ�����"
            Else
                CheckDisable = "���ֵ���������ҩƷ������û����ã�" & vbCrLf & strInfo & vbCrLf & "Ҫ������"
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetDeposit(lng����ID As Long, _
    Optional blnDateMoved As Boolean, Optional strTime As String, _
    Optional ByVal bln����תסԺ As Boolean = False, _
    Optional ByVal strPepositDate As String = "", _
    Optional intԤ����� As Integer = 0, _
    Optional rs���㷽ʽ As ADODB.Recordset) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ʣ��Ԥ������ϸ
    '���:strTime-סԺ����,��:1,2,3
    '        bln����תסԺ-�Ƿ��������תסԺ(ֻ�ܳ�ָ����Ԥ��)
    '        strPepositDate-ָ����Ԥ������
    '       intԤ�����-0-�����סԺ;1-����;2- סԺ
    '����:
    '����: Ԥ����ϸ����
    '����:���˺�
    '����:2011-03-31 14:58:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strSub1 As String
    Dim strWherePage As String, strTable As String
    Dim strWhere As String, strDate As String
    Dim strPara As String, intPara As Integer
    Dim str���� As String, strData() As String
    Dim int��ʽ As Integer, rsDeposit As ADODB.Recordset
    Dim i As Integer
    Dim str�������� As String, int�������� As Integer
    Dim str�������� As String, strDecode As String
    Dim strHead As String
    On Error GoTo errH

    strSQL = ""
    If intԤ����� = 1 Then strTime = ""    '69500
    
    strWherePage = IIf(strTime = "", "", " And instr(','||[2]||',',','||Nvl(A.��ҳID,0)||',')>0")
    strTable = IIf(blnDateMoved, zlGetFullFieldsTable("����Ԥ����¼"), "����Ԥ����¼ A")
    strWhere = "": strDate = "2000-01-01 00:00:00"
    
    strPara = zlDatabase.GetPara("��Ԥ��ȱʡ˳��", glngSys, 1137, "0")
    intPara = Val(Split(strPara & "|", "|")(0))
    
    If strPepositDate <> "" Then
        If IsDate(strPepositDate) Then
            strDate = strPepositDate
            strWhere = " And A.�տ�ʱ��=[3]"
        End If
    End If
    
    If intԤ����� <> 0 Then
        strWhere = strWhere & " And A.Ԥ����� =[4]"
    End If
    
    If bln����תסԺ Then
        strWhere = strWhere & " And A.ժҪ='����תסԺԤ��'"
    End If

    If intPara = 0 Then
        'Ĭ������
        '����=5:���۷�
        strSQL = "" & _
        " Select a.No, a.Ʊ�ݺ�, a.Id, a.���, a.��¼״̬, a.Ԥ��id, a.����, a.���㷽ʽ, " & vbNewLine & _
        "       a.�����id, a.���㿨���, decode(nvl(A.���㿨���,0),0,0,1)  as �Ƿ����ѿ�,a.����, a.������ˮ��, a.����˵��, " & vbNewLine & _
        "       c.�Ƿ�ת�ʼ����� As ת�ʼ�����,  Nvl(c.����, q.����)  As ���������, Nvl(C.�Ƿ�����, Q.�Ƿ�����)  As �Ƿ�����, " & vbNewLine & _
        "       Nvl(C.�Ƿ�ȫ��, Q.�Ƿ�ȫ��) As �Ƿ�ȫ��, c.�Ƿ�ȱʡ����," & vbNewLine & _
        "       b.���� As ��������,  Sign(Nvl(a.���, 0)) As ��־" & vbNewLine & _
        " From (Select a.No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬," & vbNewLine & _
        "              Min(Decode(a.����id, Null, a.Id, 0) * Decode(a.��¼״̬, 1, 1, 0)) As ID," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.Id), 0)) As Ԥ��id," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.ʵ��Ʊ��), Null)) As Ʊ�ݺ�," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, To_Char(a.�տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss')), Null)) As ����," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.���㷽ʽ), Null)) As ���㷽ʽ," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.�����id), Null)) As �����id," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.���㿨���), Null)) As ���㿨���," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.����), Null)) As ����," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.������ˮ��), Null)) As ������ˮ��," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.����˵��), Null)) As ����˵��" & vbNewLine & _
        "       From " & strTable & vbNewLine & _
        "       Where a.��¼���� In (1, 11) And a.����id = [1] " & strWhere & strWherePage & _
        "       Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0" & vbNewLine & _
        "       Group By a.No) A, ���㷽ʽ B, ҽ�ƿ���� C,���ѿ����Ŀ¼ Q" & vbNewLine & _
        " Where a.���㷽ʽ = b.����(+) And a.�����id = c.Id(+) And a.���㿨��� = q.���(+) And b.���� <> 5" & vbNewLine & _
        " Order By ��־, ����, NO"
    Else
        '�������������
        strData = Split(Split(strPara & "|", "|")(1), ",")
        int�������� = 1
        For i = 0 To UBound(strData)
            int��ʽ = Val(Split(strData(i) & ":", ":")(1))
            If Split(strData(i) & ":", ":")(0) = "�ֽ������" Then
                If int�������� = 1 Then
                    str�������� = "Decode(Nvl(b.����,0),1,1"
                Else
                    str�������� = str�������� & ",1," & int��������
                End If
                
                Select Case int��ʽ
                Case 0
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 1, A.����, Null) As �����ֽ�"
                    str�������� = str�������� & ",�����ֽ�"
                Case 1
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 1, A.���, Null) As �����ֽ�"
                    str�������� = str�������� & ",�����ֽ�"
                Case 2
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 1, A.���, Null) As �����ֽ�"
                    str�������� = str�������� & ",�����ֽ� Desc"
                End Select
                
                int�������� = int�������� + 1
            End If
            
            If Split(strData(i) & ":", ":")(0) = "���������" Then
                If int�������� = 1 Then
                    str�������� = "Decode(Nvl(b.����,0),2,1,3,1,4,1,6,1,7,1"
                Else
                    str�������� = str�������� & ",2," & int�������� & ",3," & int�������� & ",4," & int�������� & ",6," & int�������� & ",7," & int��������
                End If
                
                Select Case int��ʽ
                Case 0
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 2, A.����, 3, A.����, 4, A.����, 6, A.����, 7, A.����, Null) As ��������"
                    str�������� = str�������� & ",��������"
                Case 1
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 2, A.���, 3, A.���, 4, A.���, 6, A.���, 7, A.���, Null) As ��������"
                    str�������� = str�������� & ",��������"
                Case 2
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 2, A.���, 3, A.���, 4, A.���, 6, A.���, 7, A.���, Null) As ��������"
                    str�������� = str�������� & ",�������� Desc"
                End Select
                
                int�������� = int�������� + 1
            End If
            
            If Split(strData(i) & ":", ":")(0) = "�����������" Then
                If int�������� = 1 Then
                    str�������� = "Decode(Nvl(b.����,0),8,1"
                Else
                    str�������� = str�������� & ",8," & int��������
                End If
                
                Select Case int��ʽ
                Case 0
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 8, A.����, Null) As ��������"
                    str�������� = str�������� & ",��������"
                Case 1
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 8, A.���, Null) As ��������"
                    str�������� = str�������� & ",��������"
                Case 2
                    strDecode = strDecode & ",Decode(NVL(B.����, 0), 8, A.���, Null) As ��������"
                    str�������� = str�������� & ",�������� Desc"
                End Select
                
                int�������� = int�������� + 1
            End If
        Next i
        str�������� = str�������� & ",Null) As ��������"
        
        str���� = "Order By ��־,��������" & str�������� & ",No"
        
        strSQL = "" & _
        "Select " & str�������� & strDecode & ", a.No, a.Ʊ�ݺ�, a.Id, a.���, a.��¼״̬, a.Ԥ��id, a.����, a.���㷽ʽ, " & _
        "       a.�����id, a.���㿨���, decode(nvl(A.���㿨���,0),0,0,1)  as �Ƿ����ѿ�,a.����, a.������ˮ��, a.����˵��, " & vbNewLine & _
        "       c.�Ƿ�ת�ʼ����� As ת�ʼ�����,  Nvl(c.����, q.����)  As ���������, Nvl(C.�Ƿ�����, Q.�Ƿ�����)  As �Ƿ�����, " & vbNewLine & _
        "        Nvl(C.�Ƿ�ȫ��, Q.�Ƿ�ȫ��) As �Ƿ�ȫ��, c.�Ƿ�ȱʡ����,b.���� As ��������,  Sign(Nvl(a.���, 0)) As ��־" & vbNewLine & _
        " From (Select a.No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬," & vbNewLine & _
        "              Min(Decode(a.����id, Null, a.Id, 0) * Decode(a.��¼״̬, 1, 1, 0)) As ID," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.Id), 0)) As Ԥ��id," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.ʵ��Ʊ��), Null)) As Ʊ�ݺ�," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, To_Char(a.�տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss')), Null)) As ����," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.���㷽ʽ), Null)) As ���㷽ʽ," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.�����id), Null)) As �����id," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.���㿨���), Null)) As ���㿨���," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.����), Null)) As ����," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.������ˮ��), Null)) As ������ˮ��," & vbNewLine & _
        "              Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, Null, a.����˵��), Null)) As ����˵��" & vbNewLine & _
        "       From " & strTable & vbNewLine & _
        "       Where a.��¼���� In (1, 11) And a.����id = [1] " & strWhere & strWherePage & _
        "       Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0" & vbNewLine & _
        "       Group By a.No) A, ���㷽ʽ B, ҽ�ƿ���� C,���ѿ����Ŀ¼ Q" & vbNewLine & _
        " Where a.���㷽ʽ = b.����(+) And a.�����id = c.Id(+) And a.���㿨��� = q.���(+) And b.���� <> 5" & vbNewLine & str����
    End If

    Set rsDeposit = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, strTime, strDate, intԤ�����)
    
    Set GetDeposit = rsDeposit
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetMaxDate(lng����ID As Long, lng��ҳID As Long, Optional intԭ�� As Integer) As Date
'���ܣ���ȡת�Ʋ��������ϴα䶯ʱ��
'������intԭ��=�����ϴα䶯��ԭ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    GetMaxDate = #1/1/1900#
    intԭ�� = 0
    
    strSQL = " Select ��ʼʱ��,��ʼԭ�� From ���˱䶯��¼" & _
             " Where ��ʼʱ�� is Not NULL And ��ֹʱ�� is NULL And ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        GetMaxDate = IIf(IsNull(rsTmp!��ʼʱ��), GetMaxDate, rsTmp!��ʼʱ��)
        intԭ�� = Nvl(rsTmp!��ʼԭ��, 0)
    End If
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
    '�շѷֱҴ���ʽ
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    gBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 3, 1)))
    
    gbln���뷢ҩ = zlDatabase.GetPara(16, glngSys) = "1"
    gblnStock = zlDatabase.GetPara(18, glngSys) = "1"
    
    'Ʊ�ݺ��볤�ȡ����￨�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytFactLength = Val(Split(strValue, "|")(2))
    
    gbyt���δִ�� = Val(zlDatabase.GetPara(22, glngSys))
    gbyt���δ��ҩ = Val(zlDatabase.GetPara(154, glngSys))  '33048
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    'Ʊ���ϸ����
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnStrictCtrl = Mid(strValue, 3, 1) = "1"
    
    gbln��Ժ��׼���� = zlDatabase.GetPara(31, glngSys) = "1"
    
    'һ��ͨ������֤
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    gdblԤ��������鿨 = Val(Split(strValue, "|")(0))
    gbytԤ����˷��鿨 = Val(Split(strValue, "|")(1))
    gbln���ѿ��˷��鿨 = zlDatabase.GetPara(282, glngSys) = "1"
    
    gblnִ�к��� = True ' zlDatabase.GetPara(33, glngSys) = "1"
    
    strValue = zlDatabase.GetPara(41, glngSys)
    gstrҽ���������� = "'" & Replace(strValue, "|", "','") & "'"
    strValue = zlDatabase.GetPara(42, glngSys)
    gstr���ѷ������� = "'" & Replace(strValue, "|", "','") & "'"
    
    gblnҽ��������ܳ�Ժ = zlDatabase.GetPara(43, glngSys) = "1"
    
    '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
            
    gstrCardPass = zlDatabase.GetPara(46, glngSys, , "0000000000")
    gbln����ִ�� = zlDatabase.GetPara(51, glngSys) = "1"
    gbln������ = zlDatabase.GetPara(52, glngSys) = "1"
    gbln������ = zlDatabase.GetPara(53, glngSys) = "1"
    gbytAuditing = Val(zlDatabase.GetPara(58, glngSys))
    
    gbytҽ�������� = Val(zlDatabase.GetPara(59, glngSys))
    gcurMaxMoney = Val(zlDatabase.GetPara(60, glngSys))
    gint���ķ��Ͽ��� = Val(zlDatabase.GetPara(63, glngSys)) '
    gint������� = Val(zlDatabase.GetPara(65, glngSys, , 1)) Mod 10
    
    gbln�շ���� = zlDatabase.GetPara(72, glngSys) = "1"
    gblnҩ�ƻ��۵� = zlDatabase.GetPara(79, glngSys) = "1"
    gbln�������۵� = zlDatabase.GetPara(80, glngSys) = "1"
    ' 81����:�ò���������10.03��ǰ�ʹ��ڣ�δ�ҵ�BUG�š���˻��۵���Ŀ����ȷ�Ϸ��ã�ִ��֮�������ȷ�Ϸ��ã��ͻ���Ҫ�˹�����ȥ��˻��۵�����ҵ��������˵��������û�б�Ҫ���ڣ�Ӧ�ö�����Ϊִ�к��Զ���˻��۵���������ؿ��ư����ϴ˲������д���
    gblnִ�к���� = True ' zlDatabase.GetPara(81, glngSys) = "1"
    
    gbln��������ۿ� = zlDatabase.GetPara(93, glngSys) = "1"
    gbln�����������۷��� = zlDatabase.GetPara(98, glngSys) = "1"
                    
    '���˺� ����:????    ����:2010-12-06 23:38:53
    '���õ��۱���λ��
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    '���������ʱ,���������Ŀʱ,��λ����������
    gblnFeeKindCode = zlDatabase.GetPara(144, glngSys) = "1" And Not gbln�շ����
    
    gblnÿ��סԺ��סԺ�� = zlDatabase.GetPara(145, glngSys)
    gbytMediOutMode = Val(zlDatabase.GetPara(150, glngSys))
    gbln����ʾ�޿������ = zlDatabase.GetPara(316, glngSys) = "1"
    
    '����ȫ�ֲ���
    '-------------------------------------------------------------------------------------------------
    '����:27990
    With gTy_System_Para
        .byt����ҩƷ��ʾ = Val(zlDatabase.GetPara("����ҩƷ��ʾ")) '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
        .bytҩƷ������ʾ = Val(zlDatabase.GetPara("ҩƷ������ʾ"))  '��0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
        .int���ݲ�¼ʱ�� = Val(zlDatabase.GetPara(158, glngSys, , "24"))    '����:33744
        .byt������˷�ʽ = Val(zlDatabase.GetPara(185, glngSys, , "0"))
        .blnδ��ƽ�ֹ���� = Val(zlDatabase.GetPara(215, glngSys, , "0")) = 1    '51612
        .byt��������ʶ����� = Val(zlDatabase.GetPara(320, glngSys, , "0"))      '1-����ͨ��ɨ��¼���¼������;0-�����ƣ�����ͨ������Ȳ���
        With .TY_Balance
            .blnˢ���������� = Mid(gstrCardPass, 7, 1) = "1"
            .bytAuditing = Val(zlDatabase.GetPara(58, glngSys))
            .byt���δִ�� = Val(zlDatabase.GetPara(22, glngSys))
            .byt���δ��ҩ = Val(zlDatabase.GetPara(154, glngSys))
            .byt������δ��ҩ = Val(zlDatabase.GetPara(265, glngSys))
            .byt������δִ�� = Val(zlDatabase.GetPara(266, glngSys))
            .blnҽ��������ܳ�Ժ = zlDatabase.GetPara(43, glngSys) = "1"
            .bln��Ժ��׼���� = zlDatabase.GetPara(31, glngSys) = "1"
        End With
    End With
    InitSysPar = True
End Function

Public Sub zlInitҩ��()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҩ������ز���
    '����:���˺�
    '����:2010-01-25 21:29:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glng��ҩ�� = Val(zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, 1150))
    glng��ҩ�� = Val(zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, 1150))
    glng��ҩ�� = Val(zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, 1150))
    glng���ϲ��� = Val(zlDatabase.GetPara("ȱʡ���ϲ���", glngSys, 1150))
    
    gbln����ҩ�� = zlDatabase.GetPara("��ʾ����ҩ�����", glngSys, 1150) = "1"
    gbln����ҩ�� = zlDatabase.GetPara("��ʾ����ҩ����", glngSys, 1150) = "1"
    
    '���뷢ҩʱ�ļ��
    gstr��ҩ�� = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, 1150)
    gstr��ҩ�� = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, 1150)
    gstr��ҩ�� = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, 1150)
End Sub

Public Sub InitLocPar(lngModul As Long)
'���ܣ���ʼ��ģ�����
'��������
    Dim strValue As String
    On Error Resume Next
   
   'a.����ע���洢��ģ�����
    '----------------------------------------------------------------------------------------
    If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Or lngModul = 1137 Then
        gblnLED = Val(GetSetting("ZLSOFT", "����ȫ��", "ʹ��", 0)) <> 0
    End If
    
    'b.���ݿ�洢�Ĺ���ȫ�ֲ���
    '----------------------------------------------------------------------------------------
    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    strValue = zlDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    gbytCode = Val(zlDatabase.GetPara("���뷽ʽ"))
    gblnMyStyle = zlDatabase.GetPara("ʹ�ø��Ի����") = "1"
        
        
        
    'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    If lngModul = 1137 Then '����
        gbytInvoiceKind = Val(zlDatabase.GetPara("סԺ����Ʊ������", glngSys, lngModul, "0"))
        'glngShareUseID = Val(zlDatabase.GetPara("���ý���Ʊ������", glngSys, lngModul, "0"))
        gbytFeePrintSet = Val(zlDatabase.GetPara("������ϸ��ӡ", glngSys, lngModul, "0"))
        gbyt����ʱ��Ѫ�Ѽ�� = Val(zlDatabase.GetPara("����ʱ��Ѫ�Ѽ��", glngSys, lngModul, "0"))
    ElseIf lngModul = 1142 Then
        strValue = zlDatabase.GetPara("ҽ��������Դ", glngSys, lngModul, "111")
        '���������
        If Len(strValue) = 1 Then
            If strValue = "0" Then
                strValue = "111"
            ElseIf strValue = "1" Then
                strValue = "101"
            Else
                strValue = "010"
            End If
        End If
        gstrExe��Դ = strValue
        
        gstrExe��� = zlDatabase.GetPara("ҽ��ִ�����", glngSys, lngModul)
        gbytExe���ﵥ������ = Val(zlDatabase.GetPara("ҽ�����ﵥ������", glngSys, lngModul, "2"))
        gbytExeסԺ�������� = Val(zlDatabase.GetPara("ҽ��סԺ��������", glngSys, lngModul, "2"))
        gbytExe��쵥������ = Val(zlDatabase.GetPara("ҽ����쵥������", glngSys, lngModul, "2"))
        gbytExe��ӡ��ʽ = Val(zlDatabase.GetPara("ִ�еǼǵ���ӡ��ʽ", glngSys, lngModul, "2"))
    End If
        
    If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Or lngModul = 1137 Then
        gblnLedWelcome = zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, lngModul, "1") = "1"
    End If
        
    
    If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Then
        gstr�շ���� = zlDatabase.GetPara("�շ����", glngSys, 1150)
        gbln�������� = zlDatabase.GetPara("�������۲��˼���", glngSys, 1150) = "1"
        gblnסԺ���� = zlDatabase.GetPara("סԺ���۲��˼���", glngSys, 1150) = "1"
        
        gintOutDay = Val(zlDatabase.GetPara("��Ժ��������", glngSys, 1150))
        
        
        Call zlInitҩ��
        
    
        gbyt��������ʾ = IIf(zlDatabase.GetPara("��������ʾ��ʽ", glngSys, 1150) = "2", 2, 1)
        gblnFromDr = zlDatabase.GetPara("����ҽ��", glngSys, 1150) = "0"
        
        gblnPrice = zlDatabase.GetPara("������Ϊ���۵�", glngSys, 1150) = "1"
        gblnסԺ��λ = zlDatabase.GetPara("����ҩƷ��λ", glngSys, 1150) = "1"
        gbytSendMateria = Val(zlDatabase.GetPara("���ʺ�ҩ", glngSys, 1150))
    
        gblnPay = zlDatabase.GetPara("��ҩ����", glngSys, 1150) = "1"
        gblnTime = zlDatabase.GetPara("�������", glngSys, 1150) = "1"
        gbln��ʿ = zlDatabase.GetPara("��ʾ��ʿ", glngSys, 1150) = "1"
        
        '��ӡ����
        gbln���ʴ�ӡ = zlDatabase.GetPara("���ʴ�ӡ", glngSys, lngModul) = "1"  '���ʴ�ӡ����1150�Ĳ���
        gbln���۴�ӡ = zlDatabase.GetPara("���۴�ӡ", glngSys, 1150) = "1"
        gbln��˴�ӡ = zlDatabase.GetPara("��˴�ӡ", glngSys, 1150) = "1"
        
        
        gblnҩ���ϰల�� = Checkҩ���ϰల��
        
    ElseIf lngModul = 1137 Then
        gintOutDay = Val(zlDatabase.GetPara("��Ժ��������", glngSys, lngModul))
        
        'gint���ʴ�ӡ = Val(zlDatabase.GetPara("��ͨ���˽��ʴ�ӡ", glngSys, lngModul))
        gblnPrintByPatient = zlDatabase.GetPara("��Լ��λ�����˴�ӡ", glngSys, lngModul) = "1"
        gbln��;������Ԥ�� = zlDatabase.GetPara("��;������Ԥ��", glngSys, lngModul) = "1"
        gblnAutoOut = zlDatabase.GetPara("��Ժ���˽��ʺ��Զ���Ժ", glngSys, lngModul) = "1"
        gblnZero = zlDatabase.GetPara("���������", glngSys, lngModul) = "1"
        gbln����ָ��Ԥ���� = zlDatabase.GetPara("����ָ��Ԥ����", glngSys, lngModul) = "1"
        gbln���סԺ������������ = zlDatabase.GetPara("���סԺ������������", glngSys, lngModul) = "1"
        gint����ʱ�� = IIf(zlDatabase.GetPara("���ʷ���ʱ��", glngSys, lngModul) = "1", 1, 0)
        gbyt���ʼ����տ��� = zlDatabase.GetPara("���ʼ����տ���", glngSys, lngModul, , "0")
        '32322
        gstr���㷽ʽ��ʾ˳�� = Trim(zlDatabase.GetPara("���㷽ʽ��ʾ˳��", glngSys, lngModul, "��ҽ������-�н��;��ҽ������-�޽��;ҽ������-�н���������޸�;ҽ������-�޽���������޸�;ҽ������-�н���Ҳ������޸�;ҽ������-�޽���Ҳ������޸�"))
    
    ElseIf lngModul = 1142 Then
        'ҽ��ִ�в���
        gblnExeҽ�� = zlDatabase.GetPara("ҽ��ҽ������", glngSys, lngModul) = "1"
    End If
End Sub


Public Function ImportBill(strNO As String, blnBat As Boolean, frmParent As Object, _
    Optional blnModi As Boolean, Optional blnסԺ��λ As Boolean, Optional ByVal bln�����巨 As Boolean, _
    Optional ByVal lngUnitID As Long, Optional ByVal bln����ִ������ As Boolean = True, _
    Optional ByVal strҩƷ�۸�ȼ� As String, _
    Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ�۸�ȼ� As String) As ExpenseBill
'���ܣ���ȡ���õ��ݵ����ݶ�����(Ŀǰ���Դ�����Ŀ,����������Ŀ)
'������
'      strNO=���ݺ�
'      blnBat=�Ƿ�ಡ�˵�(�Լ��ʵ���Ч)
'      blnInHos=�Ƿ�ֻ���뵥������Ժ���˼�¼(��Ҫ���ڼ��ʱ���)  '�˲�����ȡ��,��Ϊ���޸�Ϊ���������ﲡ�˵ĵ���
'      blnModi=�Ƿ����޸ĵ���ʱ���øù���(����Ϊ����)
'      bln�����巨  �򵥼��ʵȲ����巨
'      lngUnitID    ��ǰ��������ID
'      bln����ִ������   ֻ�м��ʵ�ʱ,����Ҫȡִ������
'���أ���ŵ�����Ϣ�ĵ��ݶ���
'˵������Ϊ������ʱ��Ŀ�۸���Ϣ��������,���Է�������������¼���
'      �����ǵ��뻹���޸ĵ���,����Ӧ������ͣ���շ�ϸĿ

    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset
    Dim rsPrice As ADODB.Recordset
    Dim rsMoney As ADODB.Recordset
    Dim i As Long, intCurNo As Integer
    Dim int��� As Integer, blnDo As Boolean, dblStock As Double
    Dim blnLoad As Boolean, strSQL As String, strҩ��IDs As String, strͣ����Ŀ��� As String, strPrivs As String
    Dim curModiMoney As Currency
    
    Dim dblAllTime As Double, dblCurTime As Double, dbl�Ӱ�Ӽ��� As Double, dblPriceSingle As Double, lngLastPati As Long
    Dim colSerial As New Collection, dblPrice As Double
    Dim bytType As Byte '0-����;1-סԺ;2-�����סԺ
    Dim strTable  As String
    Dim strժҪ As String, strWherePriceGrade As String
    
    On Error GoTo errH
    '�۸�ȼ�
    If strҩƷ�۸�ȼ� <> "" Or str���ļ۸�ȼ� <> "" Or str��ͨ�۸�ȼ� <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [5])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [6])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And d.�۸�ȼ� = [7])" & vbNewLine & _
            "            Or (d.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From �շѼ�Ŀ" & vbNewLine & _
            "                                Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [5])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [6])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And �۸�ȼ� = [7])))))"
    Else
        strWherePriceGrade = " And d.�۸�ȼ� Is Null"
    End If
    
    If blnBat Or blnModi Then  '�ಡ�˵������Ҫ�޸ĵĵ�����,�϶���סԺ��
        strTable = "" & _
        "   Select A.��¼����,A.���,A.��������,A.NO,A.��¼״̬,A.�ಡ�˵�,A.Ӥ����,A.��������ID,A.�����־,A.�Ӱ��־," & _
        "          A.���ӱ�־,A.�շ�ϸĿID,A.��ҩ����,Nvl(����,1) as ����,Nvl(A.����,0) as ����," & _
        "          A.��׼����,A.ִ�в���ID,A.������,A.������,A.����Ա���,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��,A.ժҪ," & _
        "          A.�շ����,A.��������,A.����ID,A.����" & _
        "   From סԺ���ü�¼ A " & _
        "   Where A.��¼���� in (1,2)  And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� Is Null " & IIf(Not blnModi, " And Nvl(A.����,0)>=0", "") & _
        "         And A.NO=[1] And Nvl(A.�ಡ�˵�,0)=[3] " & _
        ""
    Else
        strTable = "" & _
        "   Select A.��¼����,A.���,A.��������,A.NO,A.��¼״̬,0 as �ಡ�˵�,A.Ӥ����,A.��������ID,A.�����־,A.�Ӱ��־," & _
        "          A.���ӱ�־,A.�շ�ϸĿID,A.��ҩ����,Nvl(����,1) as ����,Nvl(A.����,0) as ����," & _
        "          A.��׼����,A.ִ�в���ID,A.������,A.������,A.����Ա���,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��,A.ժҪ," & _
        "          A.�շ����,A.��������,A.����ID,A.����" & _
        "   From ������ü�¼ A " & _
        "   Where A.��¼���� in (1,2)  And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� Is Null " & IIf(Not blnModi, " And Nvl(A.����,0)>=0", "") & _
        "         And A.NO=[1]  " & _
        "   Union ALL " & _
        "   Select A.��¼����,A.���,A.��������,A.NO,A.��¼״̬,A.�ಡ�˵�,A.Ӥ����,A.��������ID,A.�����־,A.�Ӱ��־," & _
        "          A.���ӱ�־,A.�շ�ϸĿID,A.��ҩ����,Nvl(����,1) as ����,Nvl(A.����,0) as ����," & _
        "          A.��׼����,A.ִ�в���ID,A.������,A.������,A.����Ա���,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��,A.ժҪ," & _
        "          A.�շ����,A.��������,A.����ID,A.����" & _
        "   From סԺ���ü�¼ A " & _
        "   Where A.��¼���� in (1,2)  And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� Is Null " & IIf(Not blnModi, " And Nvl(A.����,0)>=0", "") & _
        "         And A.NO=[1] And Nvl(A.�ಡ�˵�,0)=[3] "
        
    End If
    
    '�շѼ�Ŀ����:�¼���۸�,����ж���۸�,��һ���շ�ϸĿID�оͻ��ж��������ͬ�ļ�¼
    '�۸񸸺� is NULL:ֻȡÿ���շ�ϸĿID�ĵ�һ��(ҩƷֻ��һ��),��ΪҪ����۸�
    strSQL = _
        " Select F.����,X.ҩƷID,W.����ID,W.��������," & _
        "       A.��� As ���,A.��������,A.NO,A.��¼����,A.��¼״̬,A.�ಡ�˵�,A.Ӥ����,G.�ѱ�,F.����,F.�Ա�,F.����,F.������," & _
        "       G.��Ժ���� as ����,F.סԺ�� as ��ʶ��,F.����ID,G.��ҳID,G.��ǰ����ID as ���˲���ID,G.��Ժ����ID as ���˿���ID,A.��������ID,A.�����־,A.�Ӱ��־," & _
        "       G.��������,A.���ӱ�־,A.�շ����,A.�շ�ϸĿID,A.��ҩ����,Nvl(����,1) as ����,Nvl(A.����,0) as ����," & _
        "       A.��׼����,A.ִ�в���ID,A.������,A.������,A.����Ա���,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��,A.ժҪ," & _
        "       B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(H.����,B.����) as ����,E1.���� as ��Ʒ��,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
        "       B.���ηѱ�,B.˵��,B.ִ�п���,B.�������,Nvl(A.��������,B.��������) ��������,D.�ּ�,D.ԭ��,D.ȱʡ�۸�,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
        "       E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.�շ����,'4',1,X.סԺ��װ) as סԺ��װ,Decode(A.�շ����,'4',B.���㵥λ,X.סԺ��λ) as סԺ��λ," & _
        "       Decode(A.�շ����,'4',Nvl(W.���÷���,0),Nvl(X.ҩ������,0)) as ����,B.¼������,A.����,M1.���� as ��������,X.��ҩ��̬,x.����ϵ��,M1.���㵥λ as ������λ" & _
        " From (" & strTable & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E,������Ϣ F, " & _
        "          ������ҳ G,�շ���Ŀ���� H,�շ���Ŀ���� E1,�������� W,ҩƷ��� X,������ĿĿ¼ M1" & _
        " Where  A.�շ�ϸĿID=D.�շ�ϸĿID And A.�շ�ϸĿID=B.ID " & _
        "       And A.�շ����=C.���� And A.�շ�ϸĿID=X.ҩƷID(+) and X.ҩ��ID=M1.ID(+) And A.�շ�ϸĿID=W.����ID(+) And D.������ĿID=E.ID" & _
        "       And A.�շ�ϸĿID=H.�շ�ϸĿID(+) And H.����(+)=1 And H.����(+)=[4]" & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.����ID=F.����ID(+) And F.����ID=G.����ID(+) And F.��ҳID=G.��ҳID(+)" & _
        "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
        "       And Sysdate Between D.ִ������ And Nvl(D.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) " & _
                strWherePriceGrade
        
    If blnBat And Not blnModi Then
        strPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
        If InStr(1, strPrivs, ";���в���;") = 0 Then strSQL = strSQL & " And G.��ǰ����ID = [2]"
    End If
        
    If Not gbln���뷢ҩ Then
        strSQL = "Select * From (" & strSQL & ")" & IIf(blnBat, " Order by LPAD(����,10,' '),����ID,���", " Order by ���")
    Else
        '���뷢ҩʱ�ſ�ʱ�ۺͷ���ҩƷ������
        strSQL = "Select * From (" & strSQL & ") Where Not(Instr(',5,6,7,',�շ����)>0 And (����=1 Or �Ƿ���=1))" & _
            IIf(blnBat, " Order by LPAD(����,10,' '),����ID,���", " Order by ���")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, lngUnitID, IIf(blnBat, 1, 0), _
        IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1), strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ�۸�ȼ�)
    
    'û�м�¼���ǿյ���
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    If rsTmp.RecordCount <> 0 Then
        '������������,Ҳ��סԺ�ĵ���,��ֻ��ѡ��һ��
        If Not blnModi Then
            rsTmp.Filter = "��¼����=1"
            i = rsTmp.RecordCount
            If i > 0 Then
                rsTmp.Filter = "��¼����=2"
                If rsTmp.RecordCount > 0 Then
                    If zlCommFun.ShowMsgbox("���ݵ���", "�ҵ����ŵ��ݺ�Ϊ[" & strNO & "]�ĵ���,������Ҫ����", _
                            "!סԺ����(&Z),���ﵥ��(&M)", frmParent, vbInformation) = "���ﵥ��" Then
                        rsTmp.Filter = "��¼����=1"
                    End If
                Else
                    rsTmp.Filter = ""   '���ﵥ��
                End If
            Else
                rsTmp.Filter = ""       'סԺ����
            End If
        Else    '�޸ĵ���ֻ����סԺ��
            rsTmp.Filter = "��¼����=2"
        End If
        
        rsTmp.MoveFirst
        
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
                        If Not CheckFeeItemAvailable(!�շ�ϸĿID, 2) Then
                            strͣ����Ŀ��� = strͣ����Ŀ��� & "," & !���
                            MsgBox "����[" & strNO & "]�е�" & !��� & "���շ���Ŀ:" & !���� & "" & vbCrLf & _
                                "��ͣ�û��ٷ����ڲ���,�����ᱻ����." & IIf(IsNull(!��������), "����д�����Ŀ,Ҳ���ᱻ����.", ""), vbInformation, gstrSysName
                            .MoveNext
                            GoTo NextRecord
                        End If
                    End If
                End If
                If blnBat And Not blnModi Then
                    If InStr(1, strPrivs, ";���в���;") > 0 And lngUnitID <> 0 And lngLastPati <> Val(!����ID) Then
                        lngLastPati = !����ID
                        If InStr(1, "," & lngUnitID & ",", "," & !���˲���ID & ",") = 0 Then
                            If MsgBox("����""" & !���� & """��ǰ�����ڵ�ǰ�������Ƿ���ò��˷���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                GoTo NextRecord
                            End If
                        End If
                    End If
                End If
                
                
                '����������=====================================================
                If i = 1 Then
                    objBill.NO = !NO
                    objBill.����ID = IIf(IsNull(!����ID), 0, !����ID)
                    objBill.��ҳID = IIf(IsNull(!��ҳID), 0, !��ҳID)
                    objBill.����ID = IIf(IsNull(!���˲���ID), 0, !���˲���ID)
                    objBill.����ID = IIf(IsNull(!���˿���id), 0, !���˿���id)
                    objBill.���� = IIf(IsNull(!����), "", !����)
                    objBill.�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
                    objBill.���� = IIf(IsNull(!����), "", !����)
                    objBill.��ʶ�� = IIf(IsNull(!��ʶ��), 0, !��ʶ��)
                    objBill.���� = "" & !����
                    objBill.�ѱ� = IIf(IsNull(!�ѱ�), "", !�ѱ�)
                    objBill.�����־ = IIf(IsNull(!�����־), 0, !�����־)
                    objBill.�Ӱ��־ = IIf(IsNull(!�Ӱ��־), 0, !�Ӱ��־)
                    objBill.Ӥ���� = IIf(IsNull(!Ӥ����), 0, !Ӥ����)
                    objBill.��������ID = IIf(IsNull(!��������ID), 0, !��������ID)
                    objBill.������ = IIf(IsNull(!������), "", !������)
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
                                
                '������źʹ�������
                intCurNo = intCurNo + 1
                objBillDetail.��� = intCurNo 'ʵ�����к�
                colSerial.Add intCurNo, "_" & !��� '��¼ԭ������ڵ��к�
                objBillDetail.�������� = Nvl(!��������, 0) '��Ϊ������������,�ȼ�¼ԭ����,�����ٴ���
                
                objBillDetail.����ID = IIf(IsNull(!����ID), 0, !����ID)
                objBillDetail.��ҳID = IIf(IsNull(!��ҳID), 0, !��ҳID)
                objBillDetail.Ӥ���� = IIf(IsNull(!Ӥ����), 0, !Ӥ����) '���ʱ�ʱ,ÿ�����˲�ͬ
                objBillDetail.����ID = IIf(IsNull(!���˲���ID), 0, !���˲���ID)
                objBillDetail.����ID = IIf(IsNull(!���˿���id), 0, !���˿���id)
                objBillDetail.���� = IIf(IsNull(!����), "", !����)
                objBillDetail.�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
                objBillDetail.���� = IIf(IsNull(!����), "", !����)
                objBillDetail.סԺ�� = IIf(IsNull(!��ʶ��), 0, !��ʶ��)
                objBillDetail.���� = "" & !����
                objBillDetail.�ѱ� = IIf(IsNull(!�ѱ�), "", !�ѱ�)
                objBillDetail.������ = IIf(IsNull(!������), 0, !������)
                
                'Ŀǰ�����ڼ��ʱ�
                objBillDetail.ҽ�Ƹ��� = Get����ҽ�Ƹ��ʽ(IIf(IsNull(!����ID), 0, !����ID), IIf(IsNull(!��ҳID), 0, !��ҳID))
                
                objBillDetail.�շ���� = IIf(IsNull(!�շ����), "", !�շ����)
                objBillDetail.�շ�ϸĿID = IIf(IsNull(!�շ�ϸĿID), 0, !�շ�ϸĿID)
                objBillDetail.���㵥λ = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                
                objBillDetail.���� = Nvl(!����, 1)
                If InStr(",5,6,7,", !�շ����) > 0 And blnסԺ��λ Then
                    objBillDetail.���� = Nvl(!����, 0) / Nvl(!סԺ��װ, 1)
                Else
                    objBillDetail.���� = Nvl(!����, 0)
                End If
                objBillDetail.ԭʼ���� = objBillDetail.���� * objBillDetail.����
                                
                If blnBat Then
                    objBillDetail.��ҩ���� = IIf(IsNull(!����), "", !����)
                Else
                    objBillDetail.��ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ����)
                End If
                
                objBillDetail.���ӱ�־ = IIf(IsNull(!���ӱ�־), 0, !���ӱ�־)
                objBillDetail.ժҪ = IIf(IsNull(!ժҪ), "", !ժҪ)
            
                If InStr(",5,6,7,", !�շ����) > 0 And gbln���뷢ҩ Then
                    objBillDetail.ִ�в���ID = 0
                Else
                    objBillDetail.ִ�в���ID = IIf(IsNull(!ִ�в���ID), 0, !ִ�в���ID)
                End If
                objBillDetail.ԭʼִ�в���ID = objBillDetail.ִ�в���ID
                
                '������ܼ�¼�޸ĸõ��ݵ�ԭ���ݺ�,������ȴҪ���ڴ�ż��ʱ��˵ķ������
                If blnBat Then
                    blnLoad = objBill.Details.Count = 0
                    If Not blnLoad Then
                        blnLoad = objBillDetail.����ID <> objBill.Details(objBill.Details.Count).����ID
                    End If
                    If blnLoad Then
                        '������Ϣ
                        Set rsMoney = Nothing
                        If blnModi Then
                            '�޸�ǰ�ĵ�ǰ���ݵĲ��˷��ý��
                            If gbytBilling = 0 Then
                                'int��Դ-1-����,2-סԺ
                                curModiMoney = GetBillMoney(2, strNO, objBillDetail.����ID)
                            End If
                            
                            Set rsMoney = GetMoneyInfo(objBillDetail.����ID, CDbl(curModiMoney), , 2)
                        Else
                            Set rsMoney = GetMoneyInfo(objBillDetail.����ID, , , 2)
                        End If
                        If Not rsMoney Is Nothing Then
                            objBillDetail.���￨�� = rsMoney!Ԥ����� & "," & rsMoney!������� & "," & rsMoney!Ԥ����� - rsMoney!�������
                        Else
                            objBillDetail.���￨�� = "0,0,0"
                        End If
                        '���շ��ö�
                        objBillDetail.���￨�� = objBillDetail.���￨�� & "," & GetPatiDayMoney(objBillDetail.����ID)
                    Else
                        objBillDetail.���￨�� = objBill.Details(objBill.Details.Count).���￨��
                    End If
                End If
                
                objBillDetail.Detail.ID = !�շ�ϸĿID
                objBillDetail.Detail.���� = !����
                objBillDetail.Detail.��� = (IIf(IsNull(!�Ƿ���), 0, !�Ƿ���) = 1)
                objBillDetail.Detail.�������� = 0 '!!!Ŀǰ���Դ�����Ŀ,����������Ŀ
                objBillDetail.Detail.���д��� = 0 '!!!Ŀǰ���Դ�����Ŀ,����������Ŀ
                objBillDetail.Detail.��� = IIf(IsNull(!���), "", !���)
                objBillDetail.Detail.���㵥λ = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                
                objBillDetail.Detail.סԺ��λ = Nvl(!סԺ��λ)
                objBillDetail.Detail.סԺ��װ = Nvl(!סԺ��װ, 1)
                                
                
                If Not gbln���뷢ҩ And InStr(",4,5,6,7,", !�շ����) > 0 Then
                    dblStock = GetStock(!�շ�ϸĿID, !ִ�в���ID)
                Else
                    dblStock = 0
                End If
                If InStr(",5,6,7,", !�շ����) > 0 And gbln���뷢ҩ Then
                    strҩ��IDs = Decode(!�շ����, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                    If strҩ��IDs <> "" Then dblStock = GetMultiStock(!�շ�ϸĿID, strҩ��IDs)
                End If
                If InStr(",5,6,7,", !�շ����) > 0 And blnסԺ��λ Then dblStock = dblStock / Nvl(!סԺ��װ, 1)
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
                objBillDetail.Detail.������� = IIf(IsNull(!�������), 0, !�������)
                objBillDetail.Detail.���� = IIf(IsNull(!��������), "", !��������)
                objBillDetail.Detail.�������� = Nvl(!��������)
                
                If InStr(",5,6,7,", !�շ����) > 0 Then
                    objBillDetail.Detail.����ְ�� = Get����ְ��(objBillDetail.Detail.ID)
                    objBillDetail.Detail.�������� = Get��������(objBillDetail.Detail.ID)
                End If
                objBillDetail.Detail.¼������ = Val("" & !¼������)
                                    
                objBillDetail.Detail.ҩ��ID = IIf(IsNull(!ҩ��ID), 0, !ҩ��ID)
                objBillDetail.Detail.��� = IIf(IsNull(!�Ƿ���), 0, !�Ƿ���) = 1
                objBillDetail.Detail.���� = IIf(IsNull(!����), 0, !����) = 1
                objBillDetail.Detail.�������� = Nvl(!��������, 0) = 1
                objBillDetail.Detail.Ҫ������ = 0
                objBillDetail.Detail.��ҩ��̬ = Val(Nvl(!����))
                objBillDetail.Detail.������λ = Nvl(!������λ)
                objBillDetail.Detail.����ϵ�� = Val(Nvl(!����ϵ��))
                
                '����:41136
                strժҪ = objBillDetail.ժҪ
                '90304
                If Not blnModi Then
                    strժҪ = gclsInsure.GetItemInfo(Val(Nvl(!����)), objBill.����ID, objBillDetail.�շ�ϸĿID, strժҪ, 2, , "|1")
                    objBillDetail.ժҪ = strժҪ
                End If
                
                '����۸񲿷�=====================================================
                Set objBillDetail.InComes = New BillInComes
                Do
                    '�������еļ۸��������¼���'***
                    If !�Ƿ��� = 1 Then
                        If InStr(",5,6,7,", !�շ����) > 0 Or (!�շ���� = "4" And Nvl(!��������, 0) = 1) Then
                            '----------------------------------------------------------------------------------------------
                            'ʱ��ҩƷ����۸�(�����ɲ�����)
                            dblAllTime = !���� * !����
                            If dblAllTime <> 0 Then
                                dblPrice = Getʱ��ҩƷӦ�ս��(objBillDetail.ִ�в���ID, CLng(!�շ�ϸĿID), dblAllTime, gstrDec, dblPriceSingle)
                                If dblAllTime <> 0 Then
                                    '����δ�ֽ����
                                    If !�շ���� = "4" Then
                                        MsgBox "ʱ����������""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    Else
                                        MsgBox "ʱ��ҩƷ""" & !���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.��׼���� = 0
                                Else
                                    'ע�⣺���������ֻ�ܱ���4λС��,�Ҳ���������,������Ҫ�ֹ�����;�����������ڼ��㾫������������
                                    objBillIncome.��׼���� = IIf(dblPriceSingle = 0, Format(dblPrice / (!���� * !����), gstrFeePrecisionFmt), dblPriceSingle) '�������ۼۼ۸�
                                End If
                            Else
                                objBillIncome.��׼���� = 0
                            End If
                            '----------------------------------------------------------------------------------------------
                        Else
                            If Abs(!��׼����) > Abs(IIf(IsNull(!�ּ�), 0, !�ּ�)) Then
                                objBillIncome.��׼���� = IIf(IsNull(!ȱʡ�۸�), 0, !ȱʡ�۸�)
                            Else
                                objBillIncome.��׼���� = !��׼����
                            End If
                        End If
                    Else
                        objBillIncome.��׼���� = !�ּ�
                    End If

                    If InStr(",5,6,7,", !�շ����) > 0 And blnסԺ��λ Then
                        objBillIncome.��׼���� = Format(objBillIncome.��׼���� * Nvl(!סԺ��װ, 1), gstrFeePrecisionFmt)
                    Else
                        objBillIncome.��׼���� = Format(objBillIncome.��׼����, gstrFeePrecisionFmt)
                    End If
                    objBillIncome.�ּ� = IIf(IsNull(!�ּ�), 0, !�ּ�) '�ּ�ԭ�۶�ҩƷ�������
                    objBillIncome.ԭ�� = IIf(IsNull(!ԭ��), 0, !ԭ��)
                    objBillIncome.������ĿID = IIf(IsNull(!������ID), 0, !������ID)
                    objBillIncome.������Ŀ = IIf(IsNull(!������Ŀ), "", !������Ŀ)
                    objBillIncome.�վݷ�Ŀ = IIf(IsNull(!�ַ�Ŀ), "", !�ַ�Ŀ)
                    
                    'Ӧ�ս��=����*����*����
                    If !�Ƿ��� = 1 And (InStr(",5,6,7,", !�շ����) > 0 Or !�շ���� = "4" And Nvl(!��������, 0) = 1) Then
                        objBillIncome.Ӧ�ս�� = dblPrice '��֤Ӧ�ս�������۽��û�����
                    Else
                        objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * objBillDetail.���� * objBillDetail.����
                    End If
                    
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
                        objBillIncome.ʵ�ս�� = ActualMoney(objBillDetail.�ѱ�, !������ID, objBillIncome.Ӧ�ս��, _
                            objBillDetail.�շ�ϸĿID, objBillDetail.ִ�в���ID, objBillDetail.ԭʼ����, dbl�Ӱ�Ӽ���)
                    End If
                    
                    With objBillIncome
                        objBillDetail.InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��
                    End With
                    
                    '�ж���һ����¼�Ƿ����ڵ�ǰ��
                    blnDo = False
                    int��� = !���
                    .MoveNext
                    If Not .EOF Then blnDo = (int��� = !���)
                    i = i + 1
                Loop While blnDo And Not .EOF
               
                With objBillDetail
                    objBill.Details.Add .Detail, .�շ�ϸĿID, .���, .��������, .����ID, .��ҳID, .����ID, .����ID, .����, .�Ա�, .����, .סԺ��, .����, _
                        .�ѱ�, .��������, .�շ����, .���㵥λ, .��ҩ����, .����, .����, .���ӱ�־, .ִ�в���ID, .InComes, .���￨��, , .������, .ҽ�Ƹ���, , , , .ժҪ, .ԭʼ����, .ԭʼִ�в���ID, .Ӥ����
                    '���뷢ҩʱ,Key����Ϊ1,��ʾ�༭ʱִ�п����в��ɽ���
                    If InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ Then
                        objBill.Details(objBill.Details.Count).Key = 1
                    End If
                End With
            Loop
        End With
        
        '�����´����������
        For i = 1 To objBill.Details.Count
            If objBill.Details(i).�������� <> 0 Then
                objBill.Details(i).�������� = colSerial("_" & objBill.Details(i).��������)
            End If
        Next
    End If
    
    If Not bln�����巨 Then
        If blnModi And Not blnBat Then  '�����ʼ����ʻ��۵����޸�ʱ(û���ſ��򵥼���,ֻ�е���ȥ��һ�ض�)
                    '��ȡ��ҩ�巨
            strSQL = "Select ��� From ҩƷ�շ���¼ Where NO=[1] And ����=9" '9-���ʵ�������ҩ��
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!���) Then
                    objBill.�巨 = rsTmp!���
                End If
            End If
        End If
    End If
    
    If Not bln����ִ������ Then
        '���˺� ����:27383 ����:2010-02-01 16:58:14
        strSQL = "Select max(����) as ���� From ҩƷ�շ���¼ Where NO=[1] And ���� =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, IIf(Not blnBat, 9, 10))
        objBill.ִ������ = 0
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!����) Then
                objBill.ִ������ = Mid(Nvl(rsTmp!����) & "00", 2, 1)
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
Public Function zlGetBalancePati(ByVal lng����ID As Long, ByRef lng����ID As Long, lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ʵĲ���ID����ҳID
    '���:
    '����:lng����ID,lng��ҳID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-03 18:48:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    
    strSQL = "" & _
        "   Select ����ID,Max(��ҳID) as ��ҳID From ( " & _
        "   Select distinct ����ID,��ҳID From סԺ���ü�¼ Where ����ID=[1]  Union " & _
        "   Select distinct ����ID,0 as ��ҳID From ������ü�¼ Where ����ID=[1]    ) Group by ����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    lng����ID = 0: lng��ҳID = 0
    If rsTemp.RecordCount > 0 Then
        lng����ID = Nvl(rsTemp!����ID, 0): lng��ҳID = Nvl(rsTemp!��ҳID, 0)
    End If
    zlGetBalancePati = rsTemp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function RePrintBalance(strNO As String, frmParent As Object, lng����ID As Long, _
       Optional intInsure As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ǰ�տ��¼���´�ӡһ��Ʊ��
    '���:blnMediCare-�Ƿ�Ϊ���ս���Ʊ��
    '����:
    '����:��ӡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-02 10:48:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoice As String
    Dim i As Long, lng����ID As Long, lngPatientID As Long, lngPatientCount As Long
    Dim blnDo As Boolean, rsTmp As ADODB.Recordset, strRptName As String, intFormat As Integer
    Dim strUseType As String, intPrintMode As Integer
    Dim lngShareUseID As Long, lng����ID As Long, lng��ҳID As Long
    Dim objInvoice As clsInvoice, objFact As clsFactProperty
    Dim bytInvoiceKind As Byte
    Dim strKind As String, rsKind As ADODB.Recordset, bytKind As Byte
    
    '��Լ��λ
    Set rsTmp = GetBanlancePatients(lng����ID)
    If rsTmp Is Nothing Then Exit Function
    lngPatientCount = rsTmp.RecordCount
    Set objInvoice = New clsInvoice
    Set objFact = New clsFactProperty
    strKind = "Select Nvl(��������,0) As ���� From ���˽��ʼ�¼ Where ID = [1] And Rownum < 2"
    Set rsKind = zlDatabase.OpenSQLRecord(strKind, "��������", lng����ID)
    If Not rsKind.EOF Then bytKind = Val(rsKind!����)
    If bytKind = 1 Then
        bytInvoiceKind = Val(zlDatabase.GetPara("�������Ʊ������", glngSys, 1137, "0"))
    Else
        bytInvoiceKind = Val(zlDatabase.GetPara("סԺ����Ʊ������", glngSys, 1137, "0"))
    End If
    
    Call objInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser)

    If lngPatientCount > 1 Then
        '��Լ��λ����
        Call objInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), 0, 0, 0, objFact, , , bytKind)
        objFact.ʹ����� = zlDatabase.GetPara("��Լ��λ���ʴ�ӡ", glngSys, 1137)
    Else
        Call zlGetBalancePati(lng����ID, lng����ID, lng��ҳID)
        Call objInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), lng����ID, lng��ҳID, intInsure, objFact, , , bytKind)
    End If
    
    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        'bytInvoiceKind:����Ʊ������,0-סԺƱ��;1-����Ʊ��
        strRptName = IIf(bytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
        If objFact.��ӡ��ʽ = 0 Then   '��ȱʡƱ�ݸ�ʽ��ʾ
            objFact.��ӡ��ʽ = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
        SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", objFact.��ӡ��ʽ
        '����û�и�ʽ�Ĵ���,���,��Ҫǿ��ȱʡ��ָ����ʽ
        blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
        'ȡ��ѡ��ĸ�ʽ
        objFact.��ӡ��ʽ = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
    End If
    
    If blnDo Then
        If gblnPrintByPatient Then
            '��Լ��λ
            lngPatientCount = rsTmp.RecordCount
        Else
            lngPatientCount = 1
        End If
        
        Call GetNextInvoice(frmParent, objInvoice, objFact, lngPatientCount, lng����ID, strInvoice)
        If objFact.�ϸ���� And strInvoice = "" Then Exit Function
        objFact.LastUseID = lng����ID
        
        If gblnPrintByPatient And lngPatientCount > 1 Then
            For i = 1 To rsTmp.RecordCount
                lngPatientID = rsTmp!����ID
                '��ҩ��λ,����ͨסԺ���˴�ӡƱ��
                Call frmPrint.ReportPrint(2, strNO, lng����ID, objFact, strInvoice, , , , lngPatientID, objFact.��ӡ��ʽ)
                If i < rsTmp.RecordCount Then
                    strInvoice = ""
                    Call GetNextInvoice(frmParent, objInvoice, objFact, lngPatientCount + 1 - i, lng����ID, strInvoice, i = 1)
                    If objFact.�ϸ���� And strInvoice = "" Then Exit Function
                End If
                rsTmp.MoveNext
            Next
        Else
            Call frmPrint.ReportPrint(2, strNO, lng����ID, objFact, strInvoice, , , , , objFact.��ӡ��ʽ)
        End If
        RePrintBalance = True
    End If
End Function

Public Sub GetNextInvoice(ByRef frmParent As Object, ByVal objInvoice As clsInvoice, ByRef objFact As clsFactProperty, _
    ByVal lngLeastNum As Long, ByRef lng����ID As Long, ByRef strInvoice As String, _
    Optional ByRef blnFirst As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش����Ʊ��ʱ,��ȡ��һƱ�ݺ�
    '���:blnFirst-�����˴�ӡʱ���Ƿ��״δ�ӡ�����״δ�ӡ��ʾȷ��Ʊ�ݺţ�
    '����:lng����ID-��������ID
    '        strInvoice-���ط�Ʊ��
    '����:���˺�
    '����:2011-05-03 17:21:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean, blnInput As Boolean
    '����ϸ����Ʊ��ʹ��
    If objFact.�ϸ���� Then
        If objInvoice.GetInvoiceGroupID(UserInfo.����, objFact.Ʊ��, lngLeastNum, objFact.LastUseID, objFact.��������ID, "", objFact.ʹ�����, lng����ID) = False Then Exit Sub
        Select Case lng����ID
            Case -1
                If objFact.ʹ����� <> "" Then
                    MsgBox "��û�����ú͹��á�" & objFact.ʹ����� & "���Ľ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
            Case -2
                If objFact.ʹ����� <> "" Then
                    MsgBox "���صĹ��á�" & objFact.ʹ����� & "���Ľ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
        End Select
        If lng����ID <= 0 Then Exit Sub
    End If
        
    'ȡ��һ��Ʊ�ݺ���
    If Not objFact.�ϸ���� Then
        '�п����ǵ�һ��ʹ��
        Do
            blnInput = False
            '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
            strInvoice = UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, ""))
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                "", frmParent.Left + 1500, frmParent.Top + 1500))
                blnInput = True
            Else
                strInvoice = zlStr.Increase(strInvoice)
                If blnFirst Then
                    strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
            End If
                
            '�û�ȡ������,�����ӡ
            If strInvoice = "" Then
                If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                blnValid = True
            Else
                '���������Ч��
                If blnInput Then
                    If zlCommFun.ActualLen(strInvoice) <> gbytFactLength Then
                        MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & objFact.Ʊ�ų��� & " λ��", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
        Exit Sub
    End If
    Do
        '����Ʊ�����ö�ȡ
        blnInput = False
        Call objInvoice.zlGetNextBill(1137, lng����ID, strInvoice)
        If strInvoice = "" Then
            '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
            strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                "", frmParent.Left + 1500, frmParent.Top + 1500))
            blnInput = True
        ElseIf blnFirst Then
            strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                            strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
            blnInput = True
        End If
        '�û�ȡ������,����ӡ
        If strInvoice = "" Then Exit Sub
        
        '���������Ч��
        If blnInput Then
            If objInvoice.GetInvoiceGroupID(UserInfo.����, objFact.Ʊ��, lngLeastNum, objFact.LastUseID, objFact.��������ID, strInvoice, objFact.ʹ�����, lng����ID) = False Then Exit Sub
            If lng����ID = -3 Then
                MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
            Else
                blnValid = True
            End If
        Else
            blnValid = True
        End If
    Loop While Not blnValid
End Sub

Public Function GetBanlancePatients(lng����ID As Long) As ADODB.Recordset
'���ܣ��ж�һ�ż��ʵ����Ƿ��������ʵ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "" & _
        "   Select ����ID From סԺ���ü�¼ Where ����ID=[1] Group by ����ID Union " & _
        "   Select ����ID From ������ü�¼ Where ����ID=[1] Group by ����ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    
    Set GetBanlancePatients = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillisBatch(strNO As String) As Boolean
'���ܣ��ж�һ�ż��ʵ����Ƿ��������ʵ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(�ಡ�˵�,0) as �ಡ�˵� From סԺ���ü�¼ Where ��¼����=2 And ��¼״̬ IN(0,1,3) And NO=[1] And RowNum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    If Not rsTmp.EOF Then BillisBatch = (rsTmp!�ಡ�˵� = 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillisSimple(strNO As String, Optional bytType As Byte = 2) As Boolean
'���ܣ��ж�һ�ż��ʵ����Ƿ�Ϊ��ģʽ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Distinct ��ҩ����,�շ����,Nvl(����,1) as ����" & _
        " From סԺ���ü�¼ Where ��¼״̬ IN(0,1,3)" & _
        " And ��¼����=[2] And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytType)
    If rsTmp.RecordCount = 1 Then
        If rsTmp!���� = 1 And rsTmp!�շ���� = "Z" And Nvl(rsTmp!��ҩ����) = "Z" Then BillisSimple = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceID(strNO As String) As Long
'���ܣ���ȡһ�Ž��ʵ��ݵ�ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From ���˽��ʼ�¼ Where ��¼״̬=1 And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    If Not rsTmp.EOF Then GetBalanceID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceDeposit(ByVal lngBalanceID As Long, ByVal blnNOMoved As Boolean) As ADODB.Recordset
    '���ܣ���ȡһ�Ž��ʵ��ݵĳ�Ԥ����¼
    Dim strSQL As String
    On Error GoTo errH
    strSQL = " " & _
    "   Select a.Id, a.���ݺ�, a.Ʊ�ݺ�, To_Char(Max(b.�տ�ʱ��), 'YYYY-MM-DD') As ����, a.���㷽ʽ, " & _
    "          LTrim(To_Char(Max(a.��Ԥ��), '9999999990.00')) As ���, nvl(b.�����id, b.���㿨���) as �����Id ,min(decode(nvl(b.���㿨���,0),0,0,1)) as �Ƿ����ѿ�, Min(Nvl(c.����, q.����)) As ���������, " & _
    "          Min(Nvl(q.�Ƿ�����, c.�Ƿ�����)) As �Ƿ�����, Min(Nvl(q.�Ƿ�ȫ��, c.�Ƿ�ȫ��)) As �Ƿ�ȫ��, Min(c.�Ƿ�ȱʡ����) As �Ƿ�ȱʡ����,min(C.�Ƿ�ת�ʼ�����) as �Ƿ�ת�ʼ�����, Min(b.����) As ����, " & _
    "          Min(b.������ˮ��) As ������ˮ��, Min(b.����˵��) As ����˵�� " & _
    "   From (Select ID, NO As ���ݺ�, ʵ��Ʊ�� As Ʊ�ݺ�, To_Char(�տ�ʱ��, 'YYYY-MM-DD') As ����, ���㷽ʽ, Nvl(��Ԥ��, 0) As ��Ԥ�� " & _
    "          From ����Ԥ����¼ " & _
    "          Where Mod(��¼����, 10) = 1 And ����id = [1] And Nvl(��Ԥ��, 0) <> 0) A, ����Ԥ����¼ B, ҽ�ƿ���� C, ���ѿ����Ŀ¼ Q " & _
    "   Where a.���ݺ� = b.No And b.��¼���� = 1 And b.�����id = c.Id(+) And b.���㿨��� = q.���(+) " & _
    "   Group By a.Id, a.���ݺ�, a.Ʊ�ݺ�, a.���㷽ʽ,nvl(b.�����id, b.���㿨���) " & _
    "   Order By ����, a.���㷽ʽ"
    If blnNOMoved Then strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
    Set GetBalanceDeposit = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngBalanceID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalancePay(ByVal lngBalanceID As Long, ByVal blnNOMoved As Boolean) As ADODB.Recordset
'���ܣ���ȡһ�Ž��ʵ��ݵĽ����¼
    Dim strSQL As String
    On Error GoTo errH
    strSQL = _
            "Select A.���㷽ʽ,Ltrim(To_Char(A.��Ԥ��,'9999999990.00')) as ���," & _
            " A.�������,Nvl(B.����,0) as ���� From " & IIf(blnNOMoved, "H", "") & "����Ԥ����¼ A,���㷽ʽ B" & _
            " Where mod(A.��¼����,10)=2 And A.����ID=[1]" & _
            " And A.���㷽ʽ=B.����(+) Order by A.���㷽ʽ"
        
    Set GetBalancePay = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngBalanceID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!���ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistErrRecord(lngID As Long) As Boolean
'���ܣ���������ʱ�жϽ���ʱ�Ƿ���������,���û��,��Ҫ��ȡ���ݺ�,�������������ü�¼
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select NO From סԺ���ü�¼ Where Nvl(���ӱ�־,0)=9 And ����ID=[1] Union " & _
             "Select NO From ������ü�¼ Where Nvl(���ӱ�־,0)=9 And ����ID=[1] "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngID)
    ExistErrRecord = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsureInfo(lng����ID As Long) As String
'���ܣ���ȡסԺ���˱����ʻ���Ϣ
'���أ�"������;ҽ����"
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '���Ӳ�����ҳ,ȷ������סԺ�Ǳ��ղ���,����һ����Ժ
    strSQL = "Select A.����,B.ҽ����" & _
        " From ������� A,�����ʻ� B,������Ϣ C,������ҳ D" & _
        " Where A.���=B.���� And B.����ID=C.����ID" & _
        " And B.����=D.���� And C.����ID=D.����ID" & _
        " And D.��ҳID=C.��ҳID And C.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    If Not rsTmp.EOF Then GetInsureInfo = rsTmp!���� & ";" & rsTmp!ҽ����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetMinMaxDate(ByVal lngID As Long, dMin As Date, dMax As Date, Optional ByVal blnNOMoved As Boolean) As Boolean
'���ܣ����ݽ���ID��ȡ�����С�Ǽ�/����ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If gint����ʱ�� = 0 Then
        
        strSQL = "Select Max(�Ǽ�ʱ��) as ���,Min(�Ǽ�ʱ��) as ��С From " & IIf(blnNOMoved, "H", "") & "������ü�¼ Where ����ID=[1] Union all " & _
                 "Select Max(�Ǽ�ʱ��) as ���,Min(�Ǽ�ʱ��) as ��С From " & IIf(blnNOMoved, "H", "") & "סԺ���ü�¼ Where ����ID=[1]"
    Else
        strSQL = "Select Max(����ʱ��) as ���,Min(����ʱ��) as ��С From " & IIf(blnNOMoved, "H", "") & "������ü�¼ Where ����ID=[1] Union all " & _
                 "Select Max(����ʱ��) as ���,Min(����ʱ��) as ��С From " & IIf(blnNOMoved, "H", "") & "סԺ���ü�¼ Where ����ID=[1]"
    End If
    
    strSQL = "Select Max(���) as ���,Min(��С) as ��С From ( " & strSQL & ")"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!���) Or IsNull(rsTmp!��С) Then Exit Function
        dMax = rsTmp!���
        dMin = rsTmp!��С
        GetMinMaxDate = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDept(lng����ID As Long) As Long
'���ܣ����ز�����������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.��Ժ����ID From ������Ϣ A,������ҳ B Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    If Not rsTmp.EOF Then GetPatiDept = IIf(IsNull(rsTmp!��Ժ����ID), 0, rsTmp!��Ժ����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get����ҩ�嵥(strNO As String, strTime As String, bln���ʱ� As Boolean) As ADODB.Recordset
'���ܣ����ݷ��õ��ݺ�,�Ǽ�ʱ��,��ȡ����ҩƷ�嵥
'˵������ͨ��ҩʱΪ���˿��ң����ҽ����Ϊ�������ҡ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.ID,A.�ⷿID,A.�Է�����ID" & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B" & _
        " Where A.NO=[1] And A.����=[2] And Mod(A.��¼״̬,3)=1 And A.����� is NULL" & _
        " And A.NO=B.NO And A.����ID=B.ID And B.��¼״̬<>0 And B.�Ǽ�ʱ��+0=[3]" & _
        " Order by A.ҩƷID"
    If strTime <> "" Then
        Set Get����ҩ�嵥 = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, IIf(bln���ʱ�, 10, 9), CDate(strTime))
    Else
        Set Get����ҩ�嵥 = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, IIf(bln���ʱ�, 10, 9))
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetStockInfo(lngҩƷID As Long, blnҩ�� As Boolean, blnҩ�� As Boolean, Optional ByVal blnסԺ��λ As Boolean) As String
'���ܣ���ȡҩƷ�ڸ���ҩ����ҩ��Ŀ����Ϣ
'������"blnҩ��/blnҩ��"����Ҫ��һ������Ϊ��
'���أ�������Ϣ
    Dim strSQL As String, strSQL2 As String, i As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
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
        " Nvl(Sum(A.��������),0)" & IIf(blnסԺ��λ, "/Nvl(C.סԺ��װ,1)", "") & " as ���" & _
        " From ҩƷ��� A,(" & strSQL & ") B,ҩƷ��� C" & _
        " Where A.�ⷿID=B.ID And A.ҩƷID=C.ҩƷID" & _
        " And ((A.Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
        " Or (Nvl(C.ҩ������,0)=0 And A.�ⷿID IN(" & strSQL2 & ")))" & _
        " And A.����=1 And A.ҩƷID=[1]" & _
        " Group by B.����,B.����,A.�ⷿID,Nvl(C.סԺ��װ,1)" & _
        " Having Sum(Nvl(A.��������,0))<>0" & _
        " Order By B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngҩƷID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!���� & ":" & rsTmp!���
        rsTmp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    GetStockInfo = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetItemLog(ByVal int��Դ As Integer, strNO As String, bytFlag As Byte, ��� As Integer, Optional blnNOMoved As Boolean) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ�����еĽ���
    '��Σ�int��Դ-1-����;2-סԺ
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-06 17:07:09
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If int��Դ = 1 Then
        strSQL = "Select ���� From " & IIf(blnNOMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼") & _
                " Where ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=[2] And ���=[3]"
    Else
        strSQL = "Select ���� From " & IIf(blnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼") & _
                " Where ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=[2] And ���=[3]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, ���)
    
    If Not rsTmp.EOF Then GetItemLog = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckNegative(ByVal lngPatient As Long, ByVal lngPage As Long, ByVal lngItem As Long, ByVal lngExecuteDept As Long, _
    ByVal dblNum As Double, ByVal dblסԺ��װ As Double, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡�˱��δ�סԺ���շ���Ŀ�������ϼ��Ƿ��㹻����
    '���:lngNum-����ĸ��������������ҩƷ�����ݲ���ת�����ۼ۵�λ�ٴ������ͬһ����������ͬ����Ŀ��ִ�п��ҵ��ж��У���ʱ����飬����֮ǰ�ټ��
    '     strPrivs-Ȩ�޴�
    '����:
    '����:�㹻����Ȩ�޳帺��ʱ,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-29 12:01:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblδ�� As Double, dbl�ѽ� As Double
    Dim strSQL As String, rsTmp As ADODB.Recordset
    '����:26951
    If InStr(1, strPrivs, ";�������ʲ���鷢����Ŀ;") > 0 Then
        '���ڸ�������ʱ����鱾��סԺ��������Ŀ����,�д�Ȩ��,����¼�벡��δ�������ķ�����Ŀ���г���,�����鱾��סԺ��������Ŀ�������ܳ���
        CheckNegative = True: Exit Function
    End If
    
    '��¼���� In(2,3)ȡ���������ϵ����:  :28029
    On Error GoTo errH
    CheckNegative = True
    
   ' strSQL = "" & _
            "   Select Nvl(Sum(Nvl(����, 1) * ����),0) As ����," & vbNewLine & _
            "           Sum(decode(����ID,NULL,0,1)* Nvl( ����,1)* ����) as ��������  " & _
            "   From סԺ���ü�¼" & vbNewLine & _
            "   Where  ��¼���� In(2,3) and ���ʷ��� = 1 And �۸񸸺� Is Null" & _
                    IIf(gbytBilling = 0, " And ��¼״̬<>0", "") & " And ����id = [1] And ��ҳid = [2]" & vbNewLine & _
            "      And �շ�ϸĿid+0 = [3] And ִ�в���id+0 = [4]"
    '����:39836
    strSQL = " " & _
    "   Select Nvl(Sum(Decode(A.��¼����, 2, 1, 3, 1, 0) * Nvl(A.����, 1) * A.����), 0) As ����, " & _
    "          Sum(Decode(nvl(Mod(M.��¼״̬, 3),1), 0, 1, 1, 1, -1) * Decode(A.����id, Null, 0, 1) * Nvl(A.����, 1) * A.����) As �������� " & _
    "   From סԺ���ü�¼ A, ���˽��ʼ�¼ M " & _
    "   Where  A.����id = M.ID(+) And A.���ʷ��� = 1 And A.�۸񸸺� Is Null  " & IIf(gbytBilling = 0, " And A.��¼״̬<>0", "") & _
    "         And A.����id = [1] And A.��ҳid = [2] And " & _
    "         A.�շ�ϸĿid + 0 = [3] And A.ִ�в���id + 0 = [4] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngPatient, lngPage, lngItem, lngExecuteDept)
    
    If Not rsTmp.EOF Then
        '����:32106
        If Val(FormatEx(Abs(dblNum), 8)) > Val(FormatEx(Val(Nvl(rsTmp!����)), 8)) Then
                MsgBox "�����������ڸò��˱���סԺ�ڵ�ǰִ�п��ҵļ�������" & FormatEx(rsTmp!���� / IIf(gblnסԺ��λ, dblסԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                CheckNegative = False: Exit Function
        End If
        Select Case gbytBillOpt '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
        Case 0  '����
        Case 1   '����
            dblδ�� = Val(FormatEx((Val(Nvl(rsTmp!����)) - Val(Nvl(rsTmp!��������))) / IIf(gblnסԺ��λ, dblסԺ��װ, 1), 8))
            dbl�ѽ� = Val(FormatEx(Val(Nvl(rsTmp!��������)) / IIf(gblnסԺ��λ, dblסԺ��װ, 1), 8))
            If Val(FormatEx(Abs(dblNum), 8)) > Val(FormatEx(dblδ��, 8)) Then
                If MsgBox("��������(" & FormatEx(FormatEx(Abs(dblNum) / IIf(gblnסԺ��λ, dblסԺ��װ, 1), 8), 5) & _
                        ") �а������Ѿ����ʲ���(δ��:" & FormatEx(dblδ��, 5) & "; �ѽ�:" & FormatEx(dbl�ѽ�, 5) & ") ��" & vbCrLf & _
                    " �Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    CheckNegative = False: Exit Function
                End If
            End If
        Case 2   '��ֹ
                dblδ�� = Val(FormatEx((Val(Nvl(rsTmp!����)) - Val(Nvl(rsTmp!��������))) / IIf(gblnסԺ��λ, dblסԺ��װ, 1), 8))
                dbl�ѽ� = Val(FormatEx(Val(Nvl(rsTmp!��������)) / IIf(gblnסԺ��λ, dblסԺ��װ, 1), 8))
                If Val(FormatEx(Abs(dblNum), 8)) > Val(FormatEx(dblδ��, 8)) Then
                    Call MsgBox("��������(" & FormatEx(FormatEx(Abs(dblNum) / IIf(gblnסԺ��λ, dblסԺ��װ, 1), 8), 5) & _
                        ") �а������Ѿ����ʲ���(δ��:" & FormatEx(dblδ��, 5) & "; �ѽ�:" & FormatEx(dbl�ѽ�, 5) & ") ,���ܼ�����" & vbCrLf & _
                    "", vbInformation + vbOKOnly, gstrSysName)
                    CheckNegative = False: Exit Function
                End If
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    CheckNegative = False
    Call SaveErrLog
End Function

Public Function GetPatientFeeItemTotal(ByVal lngPatient As Long, ByVal lngPage As Long, ByVal strNO As String) As ADODB.Recordset
'���ܣ���ȡָ�����ݵ��շ���Ŀ�ļ������ݼ���
'������
    Dim strSQL As String

    On Error GoTo errH
    '��¼���� In(2,3)-�ų��������ϵ�����
    strSQL = "Select A.�շ�ϸĿid, A.ִ�в���id, Sum(Nvl(A.����, 1) * A.����" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ") As ����" & vbNewLine & _
            "From סԺ���ü�¼ A,ҩƷ��� X" & vbNewLine & _
            "Where A.��¼���� In(2,3) And A.���ʷ��� = 1 And A.�۸񸸺� Is Null And A.����id = [1] And A.��ҳid = [2] And A.�շ�ϸĿID=X.ҩƷID(+)" & _
            IIf(gbytBilling = 0, " And A.��¼״̬<>0", "") & " And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From סԺ���ü�¼ B" & vbNewLine & _
            "       Where NO = [3] And ��¼���� In(2,3) And A.�շ�ϸĿid = B.�շ�ϸĿid + 0 And A.ִ�в���id = B.ִ�в���id + 0)" & vbNewLine & _
            "Group By �շ�ϸĿid, ִ�в���id"
    Set GetPatientFeeItemTotal = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngPatient, lngPage, strNO)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    
    Call SaveErrLog
End Function

Public Function GetNOFeeItem(ByVal strNO As String, ByVal bytFlag As Byte, Optional ByVal strRows As String) As ADODB.Recordset
'���ܣ���ȡָ�����ݵķ����е��շ���Ŀ��ִ�п���
'������
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select ���,�շ����,�շ�ϸĿid, ִ�в���id" & vbNewLine & _
            "From סԺ���ü�¼ A" & vbNewLine & _
            "Where NO = [1] And ��¼���� = [2] And �۸񸸺� Is Null" & IIf(strRows = "", "", " And Instr(','||[3]||',',','||���||',')>0")
    Set GetNOFeeItem = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, strRows)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillCanBeOperate(ByVal strNO As String, ByVal strPriv As String, _
    ByVal strNote As String, Optional ByVal strTime As String, _
    Optional str����IDs As String, Optional ByVal bytType As Byte = 2, _
    Optional ByVal byt������Դ As Byte) As Boolean
'���ܣ����ݵ��ݵĲ�����Ϣ�ж��Ƿ���Ȩ�޲����õ���
'������strNote=������������,������ʾ������ʱ�����⴦��
'      str����IDs=����ʱ��������������Ĳ���ID��,��Ϊ���в���
'      byt������Դ 0-סԺ,1-����
'˵������Ҫ�ǲ��˳�Ժ(��Ԥ��Ժ)��,���û��Ȩ��,���������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnOut As Boolean
    Dim strInfo As String
    
    str����IDs = ""
    
    If InStr(strPriv, ";��Ժδ��ǿ�Ƽ���;") > 0 _
        And InStr(strPriv, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        BillCanBeOperate = True: Exit Function
    End If
    
    On Error GoTo errH
    
    '����޶�Ӧ��ҳ,�����ѳ�Ժ����(�����ﲡ��ҽ������)
    If strNote Like "*����" Then
        '���ʲ���ʱ,ֻ�Կ������ʲ������ݽ����ж�
        strSQL = _
            " Select ��� From סԺ���ü�¼" & _
            " Where ��¼����=[2] And NO=[1] And Nvl(ִ��״̬,0)<>1 And �۸񸸺� is NULL" & _
            " Group by ��� Having Nvl(Sum(Nvl(����,1)*����),0)<>0"
    ElseIf strNote Like "*���" Then
        '��˲���ʱ,ֻ��δ��˲������ݽ����ж�
        strSQL = _
            " Select ��� From סԺ���ü�¼" & _
            " Where ��¼����=2 And �۸񸸺� is NULL And ��¼״̬=0 And NO=[1]"
    End If
    strSQL = "Select Distinct ����,����ID,��ҳID From סԺ���ü�¼" & _
        " Where ��¼����=[2] And NO=[1] And ��¼״̬ IN(0,1,3)" & _
        IIf(strTime <> "", " And �Ǽ�ʱ��=[3]", "") & _
        IIf(strSQL <> "", " And Nvl(�۸񸸺�,���) IN(" & strSQL & ")", "")

    strSQL = "Select B.����ID,B.����," & _
    " Decode(A.����ID,NULL,Sysdate,A.��Ժ����) as ��Ժ����," & _
    " Nvl(A.״̬,0) as ״̬,Nvl(C.�������,0) as ���" & _
    " From ������ҳ A,(" & strSQL & ") B,������� C" & _
    " Where B.����ID=A.����ID(+) And C.����(+)=1 And C.����(+)=2  And B.��ҳID=A.��ҳID(+) And B.����ID=C.����ID(+) And C.����(+)=1 And C.����(+)=2 "
    
    If byt������Դ = 1 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytType, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytType)
    End If
    
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!��Ժ����) Or rsTmp!״̬ = 3 Then
            If rsTmp!��� = 0 And InStr(strPriv, ";��Ժ����ǿ�Ƽ���;") = 0 Then
                strInfo = strInfo & vbCrLf & "����""" & rsTmp!���� & """�ѳ�Ժ(��Ԥ��Ժ)�ҷ����Ѿ����塣"
            ElseIf rsTmp!��� <> 0 And InStr(strPriv, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
                strInfo = strInfo & vbCrLf & "����""" & rsTmp!���� & """�ѳ�Ժ(��Ԥ��Ժ)�ҷ�����δ���塣"
            Else
                str����IDs = str����IDs & "," & rsTmp!����ID
            End If
        Else
            str����IDs = str����IDs & "," & rsTmp!����ID
        End If
        rsTmp.MoveNext
    Loop
    str����IDs = Mid(str����IDs, 2)
        
    'ֻ�м��ʱ����ʿ��Բ��ݼ���
    If str����IDs = "" Or (strInfo <> "" And strNote <> "����") Then
        MsgBox Mid(strInfo, 3) & vbCrLf & "��û��Ȩ�޶Ե���""" & strNO & """����" & strNote & "��", vbInformation, gstrSysName
        Exit Function
    Else
        If UBound(Split(str����IDs, ",")) + 1 = rsTmp.RecordCount Then str����IDs = ""
        If strInfo <> "" Then
            MsgBox Mid(strInfo, 3) & vbCrLf & "��ֻ�ܶԵ������������˵ķ��ý���" & strNote & "��", vbInformation, gstrSysName
        End If
    End If
    
    BillCanBeOperate = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function BillCanModi(strNO As String, bytFlag As Byte) As Boolean
'���ܣ��ж�һ�ŵ����Ƿ�����޸�
'������bytFlag=��¼����
'˵������������д��ڷ�����ʱ��ҩƷ,�������޸�(��Ϊ��������)
'***
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.ID" & _
        " From סԺ���ü�¼ A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
        " Where A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID" & _
        " And A.��¼״̬ IN(0,1,3) And (Nvl(B.ҩ������,0)=1 Or Nvl(C.�Ƿ���,0)=1)" & _
        " And A.NO=[1] And A.��¼����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    BillCanModi = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadҩƷ��Ϣ(lngҩƷID As Long) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.* From ҩƷ���� A,ҩƷ��� B Where A.ҩ��ID=B.ҩ��ID And B.ҩƷID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngҩƷID)
    If Not rsTmp.EOF Then Set ReadҩƷ��Ϣ = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function BillCanDelete(ByVal strNO As String, ByVal bytFlag As Byte, _
    Optional ByVal blnBat As Boolean, Optional ByVal strTime As String, _
    Optional ByVal strPrivs As String, Optional ByRef blnFlagPrint As Boolean, _
    Optional ByVal byt������Դ As Byte) As Integer
'���ܣ��ж�һ�ŵ����Ƿ�����˷ѻ�����
'������strNO=���ݺ�,bytFlay=��¼����,blnBat=�Ƿ�ಡ�˵�,strTime=���ݵĵǼ�ʱ��
'      strPrivs=������룬�������жϷ���ҩƷ���ʻ���������Ȩ��(ҽ������,�򵥼��ʿ��Բ���)
'      byt������Դ 0-סԺ,1-����
'˵���������˷ѻ����ʵ�����
'    1.����δ��ȫִ��(ִ��״̬=0,2)
'    2.ʣ��������<>0
'���أ�
'   -1=����ʧ��
'    0=�����˷ѻ�����
'    1=���ݲ����ڻ�û�и�����շ���Ŀ������Ȩ��
'    2=�Ѿ�ȫ����ȫִ��(ִ��״̬=1)
'    3=δ��ȫִ�в���ʣ������Ϊ0
'    blnFlagPrint=����Ӧ�������Ƿ��Ѵ�ӡ(����ҽ���еĲɼ���ʽ��ִ��)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strFeeKind As String
    
    On Error GoTo errH
    '֮ǰ�Ѽ��,������һ������Ȩ��
    If strPrivs <> "" Then
        '55380
        Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
        blnYP = zlStr.IsHavePrivs(";" & strPrivs & ";", "ҩƷ����")
        blnZL = zlStr.IsHavePrivs(";" & strPrivs & ";", "��������")
        blnWC = zlStr.IsHavePrivs(";" & strPrivs & ";", "��������")
        If blnYP And blnWC And blnZL Then
            '����,������
        ElseIf blnYP And blnWC And Not blnZL Then
            strFeeKind = " And �շ����   In('4','5','6','7')"
        ElseIf blnYP And Not blnWC And blnZL Then
            strFeeKind = " And �շ����   <>'4'"
        ElseIf blnYP And Not blnWC And Not blnZL Then
            strFeeKind = " And �շ���� In('5','6','7')"
        ElseIf Not blnYP And blnWC And blnZL Then
            strFeeKind = " And �շ���� Not In('5','6','7')"
        ElseIf Not blnYP And Not blnWC And blnZL Then
            strFeeKind = " And �շ���� Not In('4','5','6','7')"
        ElseIf Not blnYP And blnWC And Not blnZL Then
            strFeeKind = " And �շ���� ='4'"
        End If
    End If
    
    '1.����δ��ȫִ��(ִ��״̬=0,2)
    strSQL = "Select Distinct Nvl(A.ִ��״̬,0) as ִ��״̬,B.��������" & _
        " From סԺ���ü�¼ A,����ҽ������ B" & vbNewLine & _
        " Where A.NO=[1] And A.��¼����=[2] And A.��¼״̬ IN(0,1,3)" & IIf(byt������Դ = 0, " And Nvl(�ಡ�˵�,0)=[3]", "") & vbNewLine & _
        " And A.ҽ�����=B.ҽ��ID(+) And A.NO=B.NO(+) And A.��¼����=B.��¼����(+)" & _
        IIf(strTime <> "", " And �Ǽ�ʱ��=[4]", "") & strFeeKind
    If byt������Դ = 1 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, IIf(blnBat, 1, 0), CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, IIf(blnBat, 1, 0))
    End If
    
    If rsTmp.EOF Then BillCanDelete = 1: Exit Function '���ݲ����ڻ�û�и�����շ���Ŀ������Ȩ��
    blnFlagPrint = Not IsNull(rsTmp!��������)
    
    '�����Ѿ�ȫ����ȫִ��
    rsTmp.Filter = "ִ��״̬<>1"
    If rsTmp.EOF Then BillCanDelete = 2 ': Exit Function
    
    
    'δ��ȫִ�в���ʣ��������<>0
    '��ԭʼ��������δ��ȫִ�е��д�(������ҩ���˷Ѻ�ִ��״̬=1,���˷Ѽ�¼ִ��״̬<>1)
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    strSQL = _
        " Select Nvl(�۸񸸺�,���) as ���" & _
        " From סԺ���ü�¼" & _
        " Where Nvl(ִ��״̬,0)<>1 And NO=[1] And ��¼����=[2] And ��¼״̬ IN(0,1,3)" & _
                IIf(strTime <> "", " And �Ǽ�ʱ��=[3]", "")
    strSQL = _
        " Select ���,�շ�ϸĿID,Sum(����) as ʣ���� " & _
        " From ( Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���,�շ�ϸĿID," & _
        "               Avg(Nvl(����,1)*����) as ���� " & _
        "        From סԺ���ü�¼" & _
        "        Where NO=[1] And ��¼����=[2] And Nvl(ִ��״̬,0)<>1 And Nvl(�۸񸸺�,���) IN(" & strSQL & ")" & _
        "        Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���),�շ�ϸĿID " & _
        "       )" & _
        " Group by ���,�շ�ϸĿID  " & _
        " Having Sum(����)<>0"
    If byt������Դ = 1 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    
    If rsTmp.EOF Then BillCanDelete = 3
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    BillCanDelete = -1
End Function

Public Function BillExistDelete(strNO As String, bytFlag As Byte) As Boolean
'���ܣ��ж�ָ�������Ƿ����(����)�˷ѻ����ʵ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select NO From סԺ���ü�¼ Where NO=[1] And ��¼����=[2] And ��¼״̬=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    BillExistDelete = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistMoney(strNO As String, bytFlag As Byte) As Boolean
'���ܣ��ж�ָ�����ݵ���Ŀ�Ƿ��Ѿ�ȫ������(ʣ������=0)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    strSQL = _
        " Select ���,Sum(����) as ʣ������" & _
        " From (" & _
            " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���,Avg(Nvl(����, 1) * ����) As ����" & _
            " From סԺ���ü�¼" & _
            " Where NO=[1] And ��¼����=[2]" & _
            " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    BillExistMoney = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckMediCareItem(ByVal lng�շ�ϸĿID As Long, ByVal int���� As Integer, ByVal str�շ���Ŀ���� As String, ByVal bln���� As Boolean, _
    Optional blnErrShowInsureName As Boolean = False, Optional ByVal strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��շ���Ŀ�Ƿ������˱���֧����Ŀ
    '���:lng�շ�ϸĿID-�շ�ϸĿID
    '     int����-����
    '     str�շ���Ŀ����-�շ���Ŀ����
    '     blnErrShowInsureName-������ʾʱ,�Ƿ���ʾ��������
    '����:���ڶ��뷵��true,���򷵻�False
    '����:���˺�
    '����:2015-07-09 11:45:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String, rs�۸� As ADODB.Recordset, dbl�۸� As Double
    Dim strInsureName As String, strWherePriceGrade As String
    
    CheckMediCareItem = True
    If gbytҽ�������� = 0 Then Exit Function
    
    If gclsInsure.GetCapability(support��������ҽ����Ŀ, , int����) Then Exit Function
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
        " Select  B.�ּ� " & _
        " From �շѼ�Ŀ B " & _
        " Where   ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
        "       And B.�շ�ϸĿID=[1]" & vbNewLine & _
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
        strInsureName = ""
        If blnErrShowInsureName Then
            strSQL = "Select ����  From ������� where ���=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", int����)
            If Not rsTmp.EOF Then
                strInsureName = "��" & Nvl(rsTmp!����) & "��"
            End If
        End If
        If gbytҽ�������� = 1 Then
            If MsgBox(strInsureName & "û������""" & str�շ���Ŀ���� & """��Ӧ�ı�����Ŀ,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckMediCareItem = False
            End If
        ElseIf gbytҽ�������� = 2 Then
            MsgBox strInsureName & "û������""" & str�շ���Ŀ���� & """��Ӧ�ı�����Ŀ!", vbInformation, gstrSysName
            CheckMediCareItem = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveNOAuditing(ByVal lng����ID As Long, Optional ByVal strHosTimes As String) As Boolean
'���ܣ��жϲ���δ��������Ƿ����δ��˼��ʷ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '77686,���ϴ�,2014/9/18,�����������
    If strHosTimes = "" Then
        strSQL = _
            "Select 1 From סԺ���ü�¼ A" & _
                " Where ���ʷ���=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And ����ID=[1] And Not Exists" & _
                " (Select 1 From ҩƷ�շ���¼ C Where A.ID = C.����ID And Mod(C.��¼״̬, 3) = 1 And Nvl(C.ժҪ,'��һ')='�ܷ�' And instr( ',8,9,10,21,24,25,26,',','||C.����||',')>0) And Not Exists" & _
                " (Select 1 From ����ҽ������ B Where A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ�����=B.ҽ��ID And B.ִ��״̬ = 2) And Rownum=1"
    Else
        strSQL = _
        "Select /*+ rule*/ 1 From סԺ���ü�¼ A,Table(f_num2list([2])) B" & _
            " Where ���ʷ���=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And ����ID=[1] And Not Exists" & _
            " (Select 1 From ҩƷ�շ���¼ C Where A.ID = C.����ID And Mod(C.��¼״̬, 3) = 1 And Nvl(C.ժҪ,'��һ')='�ܷ�' And instr( ',8,9,10,21,24,25,26,',','||C.����||',')>0) And Not Exists" & _
            " (Select 1 From ����ҽ������ B Where A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ�����=B.ҽ��ID And B.ִ��״̬ = 2) And Rownum=1 And A.��ҳID=B.COLUMN_VALUE"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, strHosTimes)
    HaveNOAuditing = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BalanceExistInsure(strNO As String, Optional ByRef bytFlag As Byte, Optional ByRef lng����ID As Long) As Integer
'���ܣ��жϽ��ʼ�¼���Ƿ����ָ����ҽ�����㷽ʽ
'������strNO=�շѵ��ݺ�,bytFlag-ҽ����������:1-���2-סԺ
'���أ��������,�򷵻ز�������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    lng����ID = 0
    On Error GoTo errH
    
    strSQL = "Select B.����,B.����,nvl(A.����ID,B.����ID) as ����ID  From ���˽��ʼ�¼ A,���ս����¼ B" & _
       " Where A.��¼״̬ IN(1,3) And A.NO=[1]" & _
       "    And A.ID=B.��¼ID And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    
    If Not rsTmp.EOF Then
        BalanceExistInsure = Val(IIf(IsNull(rsTmp!����), 0, rsTmp!����))
        lng����ID = Val(Nvl(rsTmp!����ID))
        bytFlag = Val("" & rsTmp!����)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetBillInsures(strInsure As String, ByVal strNO As String, _
    Optional ByVal strTime As String, Optional ByVal blnAuditing As Boolean, _
    Optional ByVal blnGetNoneInsure As Boolean, Optional ByVal bytFlag As Byte = 2, _
    Optional ByVal byt������Դ As Byte) As Boolean
'���ܣ���ȡ���ʱ��е����മ"10,20,30,...",Ҳ�����ڼ��ʵ�
'������strNO=���ʵ��ݺ�
'      blnAuditing=�Ƿ����ڼ������,ֻ���δ��˵Ĳ�������
'      blnGetNoneInsure=�Ƿ񽫷Ǳ��շ��÷���Ϊ0����
'      byt������Դ 0-סԺ,1-����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strInsure = ""
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(B.����,0) as ����" & _
        " From סԺ���ü�¼ A,������ҳ B" & _
        " Where A.��¼����=[2] And A.��¼״̬" & IIf(blnAuditing, "=0", " IN(0,1,3)") & _
            IIf(blnGetNoneInsure, "", " And B.���� is Not NULL") & _
        " And A.NO=[1] And A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
        IIf(strTime <> "", " And A.�Ǽ�ʱ��=[3]", "")
    If byt������Դ = 1 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    
    Do While Not rsTmp.EOF
        strInsure = strInsure & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    strInsure = Mid(strInsure, 2)
    GetBillInsures = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDelPriv(ByVal strNO As String, ByVal strPrivs As String, Optional ByVal strTime As String, _
        Optional ByVal bytFlag As Byte = 2, Optional ByVal bytMode As Byte = 1, _
        Optional ByVal byt������Դ As Byte) As Boolean
'���ܣ�����Ƿ�Ȩ�޳���סԺ���ʵ�
'��Σ�
'      byt������Դ 0-סԺ,1-����
'����: bytMode,����Ȩ�޲���ʱ�Ƿ����ʾ,1-�������,������,0-���ʼ���,���ؼ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ֻ�ж�δ���ʷ�����
    strSQL = "Select Nvl(Sum(Decode(�շ����,'5',1,'6',1,'7',1,0)),0) as ҩƷ��," & _
        " Nvl(Sum(Decode(�շ����,'4',1,0)),0) as ������," & _
        " Nvl(Sum(Decode(�շ����,'4',0,'5',0,'6',0,'7',0,1)),0) as ������" & _
        " From סԺ���ü�¼" & _
        " Where ��¼����=[2] And ��¼״̬ IN(0,1) And NO=[1]" & _
        IIf(strTime <> "", " And �Ǽ�ʱ��=[3]", "")
    If byt������Դ = 1 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    
    If rsTmp.EOF Then CheckDelPriv = True: Exit Function
    'û��סԺ����Ȩ��ʱ,�˵��Ͱ�ť������Ϊ���ɼ�
    '55380
    '55380
    Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
    Dim strNotPrivs As String, strNote As String
    
    blnYP = zlStr.IsHavePrivs(";" & strPrivs & ";", "ҩƷ����")
    blnZL = zlStr.IsHavePrivs(";" & strPrivs & ";", "��������")
    blnWC = zlStr.IsHavePrivs(";" & strPrivs & ";", "��������")
    
    If blnYP = False And blnZL = False And blnWC = False Then
        MsgBox "��û��ҩƷ���ʻ��������ʻ��������ʵ�Ȩ��,���ܶԵ���[" & strNO & "]�������ʣ�", vbInformation, gstrSysName
        Exit Function
    End If
    strNotPrivs = ""
    If Not blnYP Then strNotPrivs = strNotPrivs & "��ҩƷ����"
    If Not blnWC Then strNotPrivs = strNotPrivs & "����������"
    If Not blnZL Then strNotPrivs = strNotPrivs & "����������"
    strNotPrivs = Mid(strNotPrivs, 2)
    strNote = ""
    
    If blnYP Then strNote = strNote & "��ҩƷ����"
    If blnWC Then strNote = strNote & "����������"
    If blnZL Then strNote = strNote & "����������"
    strNote = Mid(strNote, 2)
    
    If rsTmp!ҩƷ�� > 0 And Not blnYP Then
        MsgBox "��û��" & strNotPrivs & "Ȩ��,ֻ�ܶԵ���[" & strNO & "]�е�" & strNote & "�������ʣ�", vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    If rsTmp!������ > 0 And Not blnWC Then
        MsgBox "��û��" & strNotPrivs & "Ȩ��,ֻ�ܶԵ���[" & strNO & "]�е�" & strNote & "�������ʣ�", vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    If rsTmp!������ > 0 And Not blnZL Then
        MsgBox "��û��" & strNotPrivs & "Ȩ��,ֻ�ܶԵ���[" & strNO & "]�е�" & strNote & "�������ʣ�", vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    CheckDelPriv = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get�������(lng�շ�ϸĿID As Long) As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ������� From �շ���ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng�շ�ϸĿID)
    If Not rsTmp.EOF Then Get������� = IIf(IsNull(rsTmp!�������), 0, rsTmp!�������)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check���۲���(ByVal strNO As String, ByVal strPrivs As String, _
    Optional ByVal strTime As String, Optional ByVal bytFlag As Byte = 2, _
    Optional ByVal byt������Դ As Byte) As String
'���ܣ������Ƿ���������۲����˽��м���,�Լ��ʵ�/����м��
'��Σ�
'      byt������Դ 0-סԺ,1-����
'˵������Ҫ���ڼ��ʵ�/���޸�,���ʡ����ڼ��ʱ�,ֻҪ����һ�����۲�����Ȩ��,��������ֹ
'���أ�û��Ȩ�޵����۲���,��"���۲���","�������۲���","סԺ���۲���"
    Dim rsTmp As ADODB.Recordset
    Dim bln�������� As Boolean
    Dim blnסԺ���� As Boolean
    Dim strSQL As String
    
    bln�������� = gbln�������� And InStr(strPrivs, ";�������ۼ���;") > 0
    blnסԺ���� = gblnסԺ���� And InStr(strPrivs, ";סԺ���ۼ���;") > 0
        
    If bln�������� And blnסԺ���� Then Exit Function
    
    If Not bln�������� And Not blnסԺ���� Then
        strSQL = "1,2"
    ElseIf Not bln�������� Then
        strSQL = "1"
    ElseIf Not blnסԺ���� Then
        strSQL = "2"
    End If
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(B.��������,0) as ��������" & _
        " From סԺ���ü�¼ A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
        " And A.NO=[1] And A.��¼����=[2]" & _
        " And Nvl(B.��������,0) IN(" & strSQL & ") And A.��¼״̬ IN(0,1,3)" & _
        IIf(strTime <> "", " And A.�Ǽ�ʱ��=[3]", "")
    If byt������Դ = 1 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    If Not rsTmp.EOF Then
        If rsTmp.RecordCount = 2 Then
            Check���۲��� = "���۲���"
        ElseIf rsTmp!�������� = 1 Then
            Check���۲��� = "�������۲���"
        ElseIf rsTmp!�������� = 2 Then
            Check���۲��� = "סԺ���۲���"
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillRows(strNO As String, bytFlag As Byte) As Integer
'���ܣ���ȡһ�ŷ��õ�����δ���ϵķ�������
'������bytFlag=��¼����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    strSQL = _
        " Select ���,Sum(����) as ʣ������" & _
        " From (" & _
            " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���,Avg(Nvl(����, 1) * ����) As ����" & _
            " From סԺ���ü�¼" & _
            " Where NO=[1] And ��¼����=[2]" & _
            " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    If Not rsTmp.EOF Then GetBillRows = rsTmp.RecordCount
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetMaxBedLen(Optional lng����ID As Long, Optional bln���� As Boolean) As Integer
'���ܣ���ȡָ�����ŵĴ�λ�ŵ���󳤶�
'������lng����ID=����ID�����ID,Ϊ0��ʾ���в��������
'      blnռ��=�Ƿ�ֻ�ܱ�ռ�õĴ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not bln���� Or lng����ID = 0 Then
        strSQL = "Select Max(Lengthb(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIf(lng����ID = 0, " is Not NULL", "=[1]")
    Else
        strSQL = "Select Max(Lengthb(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIf(lng����ID = 0, " is Not NULL", "=[1]")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!����), 0, rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckLimit(ByVal objBill As ExpenseBill, Optional intRow As Integer, Optional ByVal blnסԺ��λ As Boolean) As Boolean
'���ܣ����õ���ҩƷ�����������,�����ڼ��ʵ�/��
'˵����
'   1.ȫ��û���������������棻���г���ҩƷ�����ں�������ʾ�������ؼ١�
'   2.���ʱ���Ϊÿ�����˵������
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim tmpDetail As BillDetail, curDetail As BillDetail
    Dim i As Integer, j As Integer, dblTime As Double
    Dim dbl���� As Double, strItemIDs As String '�Ѿ������˵�ҩƷ
    Dim strPatiIDs As String, arrPati As Variant '�Ѿ������˵Ĳ���
    Dim lng����ID As Long, str���� As String
    Dim strҩƷ������ʾ As String
    
    CheckLimit = True
    If objBill.Details.Count = 0 Then Exit Function
    Err = 0: On Error GoTo errH:
    '�ռ�����
    For i = 1 To objBill.Details.Count
        If intRow = 0 Or (intRow > 0 And i = intRow) Then
            With objBill.Details(i)
                '�ռ�ҩƷID
                If InStr(strItemIDs & ",", "," & .�շ�ϸĿID & ",") = 0 And InStr(",5,6,7,", .�շ����) > 0 Then
                    strItemIDs = strItemIDs & "," & .�շ�ϸĿID
                End If
                '�ռ�������Ϣ
                If InStr(strPatiIDs & ";", ";" & .����ID & "," & .���� & ";") = 0 Then
                    strPatiIDs = strPatiIDs & ";" & .����ID & "," & .����
                End If
            End With
        End If
    Next
    If strItemIDs = "" Then Exit Function
    strItemIDs = Mid(strItemIDs, 2)
    arrPati = Split(Mid(strPatiIDs, 2), ";")
        
    strSQL = "Select A.ҩƷID,A.����ϵ��,B.���㵥λ as ������λ" & _
        " From ҩƷ��� A,������ĿĿ¼ B" & _
        " Where A.ҩ��ID=B.ID And A.ҩƷID IN (" & strItemIDs & ")"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    
    For i = 0 To UBound(arrPati)
        lng����ID = Val(Split(arrPati(i), ",")(0))
        str���� = CStr(Split(arrPati(i), ",")(1))
        strItemIDs = ""
        For j = 1 To objBill.Details.Count
            If intRow = 0 Or (intRow > 0 And j = intRow) Then
                Set tmpDetail = objBill.Details(j)
                If InStr(",5,6,7,", tmpDetail.�շ����) > 0 And tmpDetail.Detail.�������� > 0 And tmpDetail.����ID = lng����ID Then
                    If InStr(strItemIDs, "," & tmpDetail.�շ�ϸĿID) = 0 Then
                        dblTime = 0
                        For Each curDetail In objBill.Details
                            If InStr(",5,6,7,", curDetail.�շ����) > 0 And tmpDetail.�շ�ϸĿID = curDetail.�շ�ϸĿID And curDetail.����ID = lng����ID Then
                                dblTime = dblTime + curDetail.���� * curDetail.����
                            End If
                        Next
                        rsTmp.Filter = "ҩƷID=" & tmpDetail.�շ�ϸĿID
                        If Not rsTmp.EOF Then
                            If blnסԺ��λ Then
                                dbl���� = dblTime * tmpDetail.Detail.סԺ��װ * rsTmp!����ϵ��
                            Else
                                dbl���� = dblTime * rsTmp!����ϵ��
                            End If
                            If dbl���� > tmpDetail.Detail.�������� Then
                                strҩƷ������ʾ = IIf(str���� = "", "", """" & str���� & """ ��") & "ҩƷ """ & tmpDetail.Detail.���� & """ ���ܼ��� " & _
                                    FormatEx(dbl����, 5) & rsTmp!������λ & "(" & FormatEx(dblTime, 5) & IIf(blnסԺ��λ, tmpDetail.Detail.סԺ��λ, tmpDetail.Detail.���㵥λ) & ") ������������ " & _
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
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDiagnosticInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                                  ByVal str������� As String, ByVal str��¼��Դ As String) As ADODB.Recordset
'���ܣ���ȡָ�����˵���ϼ�¼'
'����:
'�������-1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
'        11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���
'��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����(����ҽ��վ,���ժҪ);

    On Local Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " Select �������,��¼��Դ,�������,����ID,���ID,��Ժ���,�Ƿ����� From ������ϼ�¼ " & _
             " Where ����ID=[1] And Nvl(��ҳID,0)=[2]" & _
             " And ��ϴ���=1 And instr([3],','||�������||',')>0 And ��¼��Դ in (" & str��¼��Դ & ")" & _
             " Order by ��¼���� Desc"
            '��ϴ���-��Ժʱ,������ҳ�����п�����д��Ҫ���,��Ҫ��ϵȶ�����¼
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID, "," & str������� & ",")
    
    If Not rsTmp.EOF Then Set GetDiagnosticInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDepCharacter(ByVal lngDepID As Long) As String
'���ܣ���ȡ���Ź�������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select �������� From ��������˵�� Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngDepID)
    
    Do While Not rsTmp.EOF
        If InStr(1, GetDepCharacter & ",", "," & rsTmp!�������� & ",") = 0 Then
            GetDepCharacter = GetDepCharacter & "," & rsTmp!��������
        End If
        rsTmp.MoveNext
    Loop
    
    If GetDepCharacter <> "" Then GetDepCharacter = Mid(GetDepCharacter, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckInhibitiveByNurse(ByRef objBill As ExpenseBill, ByRef rs������ As ADODB.Recordset) As Boolean
'���ܣ��ж�ָ���������Ƿ��л�ʿ��ֹ���������
    Dim bln��ʿ As Boolean, i As Integer
    
    CheckInhibitiveByNurse = False
    If objBill.������ <> "" Then
        Call GetOperatorInfo(rs������, objBill.������, bln��ʿ)
        If Not bln��ʿ Then Exit Function
        
        For i = 1 To objBill.Details.Count
            If InStr(",E,M,4,", objBill.Details(i).�շ����) = 0 Then
                CheckInhibitiveByNurse = True: Exit Function
            End If
        Next
    End If
End Function

Public Function CheckErrorItem() As Boolean
'���ܣ�������ڴ�����С��������Ŀ�Ƿ�������ȷ
'˵��������Ŀ��Ӧ������ҲӦΪ�����Ŀ(δ��)��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.ID,B.���,B.����,B.����" & _
        " From �շ��ض���Ŀ A,�շ���ĿĿ¼ B" & _
        " Where A.�ض���Ŀ='�����' And A.�շ�ϸĿID=B.ID" & _
        " And (B.����ʱ�� is NULL Or B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    CheckErrorItem = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Init�����˿�������(ByRef cbo������ As ComboBox, ByRef cbo�������� As ComboBox, _
                            ByRef rs������ As ADODB.Recordset, ByRef rs�������� As ADODB.Recordset, _
                            ByVal strPrivs As String, ByVal bytUseType As Byte, ByVal lngDeptID As Long _
                            ) As Boolean
'����:��ʼ��������,���������б�,������Click�¼�
'����:lngDeptID-��ǰ�����Ĳ���ID,���в���ʱΪ0

    '1.�����˾�����������,��ȱʡ������(���ǽ���һ��)
    If gblnFromDr Then
        Call FillDoctor(cbo������, rs������)
        If cbo������.ListCount = 1 Then Call zlControl.CboSetIndex(cbo������.hWnd, 0)
        
        Call FillDept(cbo��������, rs��������, rs������, strPrivs, bytUseType, lngDeptID)
        If cbo��������.ListCount = 0 Then
            MsgBox "û�г�ʼ��סԺ�ٴ�����,���ȵ����Ź��������ã�", vbInformation, gstrSysName
            Exit Function
        End If
        If cbo��������.ListIndex = -1 And cbo��������.ListCount = 1 Then Call zlControl.CboSetIndex(cbo��������.hWnd, 0)
    
    '2.�������Ҿ���������,��ʾȱʡ��������
    Else
        Call FillDept(cbo��������, rs��������, rs������, strPrivs, bytUseType, lngDeptID)
        If cbo��������.ListCount = 0 Then
            MsgBox "û�г�ʼ��סԺ�ٴ�����,���ȵ����Ź��������ã�", vbInformation, gstrSysName
            Exit Function
        End If
        
        'ȱʡ��ʾ��ǰ����,�����ǰ�����в���,����ʾ��һ��
        If lngDeptID <> 0 Then Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lngDeptID))
        If cbo��������.ListIndex = -1 Then Call zlControl.CboSetIndex(cbo��������.hWnd, 0)
        
        Call FillDoctor(cbo������, rs������, cbo��������.ItemData(cbo��������.ListIndex))
        If cbo������.ListCount = 1 Then Call zlControl.CboSetIndex(cbo������.hWnd, 0)
    End If
    
    If cbo��������.ListCount > 0 Then Call SetWidth(cbo��������.hWnd, GetWidth(cbo��������.hWnd) * 1.2)
    Init�����˿������� = True
End Function


Public Sub Set�����˿�������(ByRef cbo������ As ComboBox, ByRef cbo�������� As ComboBox, _
           ByRef rs������ As ADODB.Recordset, ByRef rs�������� As ADODB.Recordset, _
           ByVal str������ As String, ByVal lng��������ID As Long _
           )
'���ܣ�����ϵͳ�������ÿ����˺Ϳ������ң�����������Click�¼�
'       ��ҪĿ�����ڽ�ֹ��ʽ����Clickʱ�Կ����ˣ��������ҵ��໥Ӱ�죬���������Ӱ��(���磺��ı䵥�ݶ����еĶ�Ӧֵ)
    Dim lng��ԱID As Long, str�������� As String
    
    'a.�����˶���������
    If gblnFromDr Then
        Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True))
        
        If cbo������.ListIndex = -1 And str������ <> "" Then
            lng��ԱID = GetPersonnelID(str������, rs������)
            cbo������.AddItem str������
            cbo������.ItemData(cbo������.NewIndex) = lng��ԱID
            Call zlControl.CboSetIndex(cbo������.hWnd, cbo������.NewIndex)
        End If
        
        If cbo������.ListIndex <> -1 Then
            cbo��������.Clear
            Call FillDept(cbo��������, rs��������, rs������, "", 0, 0, cbo������.ItemData(cbo������.ListIndex))
        End If
        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        If cbo��������.ListIndex = -1 And lng��������ID <> 0 Then
            str�������� = GET��������(lng��������ID, rs��������)
            If str�������� <> "" Then
                cbo��������.AddItem str��������
                cbo��������.ItemData(cbo��������.NewIndex) = lng��������ID
                Call zlControl.CboSetIndex(cbo��������.hWnd, cbo��������.NewIndex)
            End If
        End If
        
    'b.�������Ҷ�������
    Else
        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        If cbo��������.ListIndex = -1 And lng��������ID <> 0 Then
            str�������� = GET��������(lng��������ID, rs��������)
            If str�������� <> "" Then
                cbo��������.AddItem str��������
                cbo��������.ItemData(cbo��������.NewIndex) = lng��������ID
                Call zlControl.CboSetIndex(cbo��������.hWnd, cbo��������.NewIndex)
            End If
        End If
        
        If cbo��������.ListIndex <> -1 Then
            cbo������.Clear
            Call FillDoctor(cbo������, rs������, lng��������ID)
        End If
        Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True))
        If cbo������.ListIndex = -1 And str������ <> "" Then
            lng��ԱID = GetPersonnelID(str������, rs������)
            cbo������.AddItem str������
            cbo������.ItemData(cbo������.NewIndex) = lng��ԱID
            Call zlControl.CboSetIndex(cbo������.hWnd, cbo������.NewIndex)
        End If
    End If
End Sub

Public Function SetDefaultDept(ByRef cbo�������� As ComboBox, ByRef rs�������� As ADODB.Recordset, _
                               ByRef rs������ As ADODB.Recordset, ByVal lng������ID As Long _
                                ) As Boolean
'����:���ݿ���������ȱʡ�Ŀ�������,��������Click�¼�
'˵��:ȱʡ����Ϊ"ֻ������סԺ"ʱ�����Զ�λȱʡ
'     ���߿����˵����п��Ҷ�Ϊͬһ�������򼶱�ʱ(�綼�Ǽ������������סԺ��)�����Զ�λȱʡ
'     ����,����������,ȡ��һ��

    Dim i As Long, lng��������ID As Long, blnDo As Boolean, lng���ȼ� As Long
    
    If cbo��������.ListCount = 1 Then
        Call zlControl.CboSetIndex(cbo��������.hWnd, 0)
    Else
        rs������.Filter = "ȱʡ=1 And ID=" & lng������ID
        If rs������.RecordCount > 0 Then lng��������ID = rs������!����ID
        
        If rs��������.RecordCount > 1 And lng��������ID > 0 Then
            rs��������.MoveFirst
            For i = 1 To rs��������.RecordCount
                If lng��������ID = rs��������!ID And rs��������!���ȼ� = 1 Then blnDo = True: Exit For
                rs��������.MoveNext
            Next
            
            If Not blnDo Then
                blnDo = True
                rs��������.MoveFirst
                For i = 1 To rs��������.RecordCount
                    If lng���ȼ� <> rs��������!���ȼ� And lng���ȼ� <> 0 Then blnDo = False: Exit For
                    lng���ȼ� = rs��������!���ȼ�
                    rs��������.MoveNext
                Next
            End If
            
            If blnDo Then Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        End If
        
        If cbo��������.ListIndex = -1 Then Call zlControl.CboSetIndex(cbo��������.hWnd, 0)
    End If
End Function

Public Sub FillDoctor(ByRef cbo������ As ComboBox, ByRef rs������ As Recordset, Optional ByVal lng����ID As Long)
'���ܣ�����ָ���Ŀ�������ID��ȡ����дҽ���б�,����ȱʡҽ��
    Dim strOldID As String
    
    cbo������.Clear
    Call GetDoctor(lng����ID, gbln��ʿ And (gstr�շ���� = "" _
        Or gstr�շ���� Like "*'E'*" Or gstr�շ���� Like "*'M'*" Or gstr�շ���� Like "*'4'*"), rs������)
    
    Do While Not rs������.EOF
        '70857:������,2014-03-07,�����˼���һ��ʱ���ڼ����ظ�������
        If InStr("," & strOldID & ",", "," & rs������!ID & ",") = 0 Then
            If gbyt��������ʾ = 1 Then
                cbo������.AddItem rs������!���� & "-" & rs������!����
            Else
                cbo������.AddItem rs������!��� & "-" & rs������!����
            End If
            cbo������.ItemData(cbo������.NewIndex) = rs������!ID
            strOldID = strOldID & rs������!ID & ","
        End If
        rs������.MoveNext
    Loop
End Sub

Public Sub FillDept(ByRef cbo�������� As ComboBox, ByRef rs�������� As Recordset, ByRef rs������ As Recordset, _
                   ByVal strPrivs As String, ByVal bytUseType As Byte, ByVal lngDeptID As Long, Optional ByVal lng��ԱID As Long)
'���ܣ���ȡ�����ؿ����б�,����ȱʡ����
'������ lngDeptID-��ǰ�����Ĳ���
'       lng��ԱID=ֻ��ȡָ����Ա���ڿ���(������ȱʡ��)
    
    Dim strSQL As String, i As Long, lngOldDepID As Long
    Dim strDepts As String  'ָ����Ա�����Ķ������
        
    cbo��������.Clear
    If rs�������� Is Nothing Then Call GetDoctorDept(rs��������, strPrivs, bytUseType, lngDeptID)
   
    If lng��ԱID <> 0 Then
        If Not rs������ Is Nothing Then
            rs������.Filter = "ID=" & lng��ԱID
            For i = 1 To rs������.RecordCount
                strDepts = strDepts & " OR ID=" & rs������!����ID      'filter��֧��in
                rs������.MoveNext
            Next
        End If
        If strDepts <> "" Then
            rs��������.Filter = Mid(strDepts, 4)
        Else
            rs��������.Filter = "ID=0" '��Աû�����ò���,����ʾ��������
        End If
    Else
        rs��������.Filter = ""
    End If
    
    If rs��������.RecordCount > 0 Then
        For i = 1 To rs��������.RecordCount
            If lngOldDepID <> rs��������!ID Then   'һ�����ſ���ͬʱ���ڲ��ƺ��ٴ�,��������ͬ��
                cbo��������.AddItem IIf(zlIsShowDeptCode, rs��������!���� & "-", "") & rs��������!����
                cbo��������.ItemData(cbo��������.NewIndex) = rs��������!ID
                lngOldDepID = rs��������!ID
            End If
            rs��������.MoveNext
        Next
    End If
End Sub


Public Function GetOperatorInfo(ByVal rs������ As Recordset, ByVal str���� As String, _
                                Optional ByRef bln��ʿ As Boolean, Optional intְ�� As Integer) As Boolean
'���ܣ���ȡָ������������(ҽ����ʿ)�����ʻ�ְ��
'���أ�intְ��:0-δ���ã�bln��ʿ:�Ƿ�ֻ�ǻ�ʿ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    bln��ʿ = False: intְ�� = 0
    If Not rs������ Is Nothing Then
        rs������.Filter = "����='" & str���� & "' " & IIf(gbln��ʿ, "", " And ��Ա����<>'��ʿ'")
        If rs������.RecordCount > 0 Then
            intְ�� = rs������!ְ��
            strSQL = rs������!��Ա����
            If strSQL = "��ʿ" Then bln��ʿ = True
            If strSQL = "ҽ��" Then bln��ʿ = False
        End If
    Else
        strSQL = _
            " Select Nvl(A.Ƹ�μ���ְ��,0) as ְ��,B.��Ա���� From ��Ա�� A,��Ա����˵�� B" & _
            " Where A.ID=B.��ԱID And B.��Ա���� IN('ҽ��','��ʿ') And A.����=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", str����)
        If Not rsTmp.EOF Then
            intְ�� = rsTmp!ְ��
            Do While Not rsTmp.EOF
                If rsTmp!��Ա���� = "��ʿ" Then bln��ʿ = True
                If rsTmp!��Ա���� = "ҽ��" Then bln��ʿ = False: Exit Do
                rsTmp.MoveNext
            Loop
        End If
    End If
    GetOperatorInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetDoctor(ByVal lng����ID As Long, ByVal bln��ʿ As Boolean, ByRef rsTmp As ADODB.Recordset)
'���ܣ���ȡָ�����ҵ�ҽ��
'������lng����ID=ָ������ID,bln��ʿ=�Ƿ�Ҳ��ȡ��ʿ
    'Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '�����Ź�������Ϊ���ٴ���ҽ����ʿ,��Ϊ���ܸ���������һ�������Ƿ�ĩ������.
    If rsTmp Is Nothing Then
        strSQL = _
            "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
            " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��,B.ȱʡ" & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C,��������˵�� D" & _
            " Where A.ID = B.��ԱID And A.ID=C.��ԱID And B.����ID=D.����ID And C.��Ա���� IN('ҽ��','��ʿ') " & _
            " And D.������� IN(2,3) And D.�������� IN('�ٴ�','����') And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by " & IIf(gbyt��������ʾ = 1, "����", "���") & ",ȱʡ Desc"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    End If
   
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

Public Sub GetDoctorDept(ByRef rs�������� As ADODB.Recordset, ByRef strPrivs As String, _
                        ByRef bytUseType As Byte, ByRef lngDeptID As Long)
'���ܣ���ȡ���п�������
'������strPrivs-�����ж��Ƿ����"�������ۼ���"��"���п���"��Ȩ��
'      bytUseType-'���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
'      lngDeptID-��ǰ����ID�����ID
    Dim strSQL As String
    
    On Error GoTo errH
    '��ѡ��������(�����ҽ������,����������סԺ��)
    If (InStr(GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "�������ۼ���") > 0 And gbln��������) Or bytUseType = 2 Then
        strSQL = "1,2,3"
    Else
        strSQL = "2,3"
    End If
    If bytUseType = 0 Or bytUseType = 1 Then
        strSQL = _
            "Select A.ID, A.����, A.����, A.����, 0 As ȱʡ, B.��������, D.���ȼ�" & vbNewLine & _
            "From ���ű� A, ��������˵�� B," & vbNewLine & _
            "     (Select ����id, Max(Decode(�������, 2, 1, 2)) As ���ȼ� From ��������˵�� Where ������� <> 0 Group By ����id) D" & vbNewLine & _
            "Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And A.ID = B.����id" & vbNewLine & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " And B.����id = D.����id And (B.������� IN(" & strSQL & ") AND B.�������� IN('�ٴ�','����') Or b.��������='����')" & vbNewLine & _
            "Order By ���ȼ�,����"
    ElseIf bytUseType = 2 Then
        'ҽ�����Ҽ���
        strSQL = _
            "Select A.ID, A.����, A.����, A.����, 0 As ȱʡ, B.��������, D.���ȼ�" & vbNewLine & _
            "From ���ű� A, ��������˵�� B," & vbNewLine & _
            "     (Select ����id, Max(Decode(�������, 2, 1, 2)) As ���ȼ� From ��������˵�� Where ������� <> 0 Group By ����id) D" & vbNewLine & _
            "Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And A.ID = B.����id" & vbNewLine & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " And B.����id = D.����id And (B.������� IN(" & strSQL & ") AND B.�������� IN('���','����','����','����','Ӫ��') Or b.��������='����')" & vbNewLine & _
            IIf(InStr(strPrivs, ";���п���;") > 0, "", " And A.ID=[1] ") & vbNewLine & _
            "Order By ���ȼ�,����"
    End If
    Set rs�������� = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngDeptID)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




Public Function GetLastAdviceTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Date
'���ܣ���ȡָ���������һ����Ч��ҽ����ʱ��
'˵�������ڲ��˳�Ժʱ�жϳ�Ժʱ�������ڸ�ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    GetLastAdviceTime = CDate("1900-01-01")
    
    On Error GoTo errH
    
    '�Գ������ִ��ʱ��Ϊ׼�ж�,��ʱ�ſ������Գ���
    '��������Ժ��ҩ�����,����"��Ժ"ҽ��Ϊ׼,��Ժʱ�䱾���ͱ�����ڸñ䶯ʱ�䡣
    strSQL = "Select Max(Nvl(ִ����ֹʱ��,Nvl(�ϴ�ִ��ʱ��,��ʼִ��ʱ��))) as ʱ��" & _
        " From ����ҽ����¼" & _
        " Where Nvl(ҽ����Ч,0)=0 And ҽ��״̬ Not IN(1,2,4)" & _
        " And Not (ִ��ʱ�䷽�� is NULL And (Nvl(Ƶ�ʴ���, 0) = 0 Or Nvl(Ƶ�ʼ��, 0) = 0 Or �����λ is NULL))" & _
        " And ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!ʱ��) Then
            GetLastAdviceTime = rsTmp!ʱ��
        End If
    End If
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
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    If Not rsTmp.EOF Then
        Checkҩ���ϰల�� = Nvl(rsTmp!Num, 0) <> 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�շ�ִ�п���ID(ByVal str��� As String, ByVal lng��Ŀid As Long, _
    ByVal intִ�п������� As Integer, ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, Optional ByVal int��Χ As Integer = 2, _
    Optional ByVal lng���ϲ���ID As Long, Optional lng���˲���ID As Long, _
    Optional lng����ȱʡִ�п��� As Long = 0) As Long
    
    '���ܣ���ȡ��ҩ�շ���Ŀ��ִ�п���
    '������int��Χ=1.����,2-סԺ
    '       lng���ϲ���ID=ָ����ȱʡִ�п���ID���˲���ID(Ŀǰ����������)
    '       lng����ȱʡִ�п���-������ĿĬ�ϵ�ִ�п���:27327
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strҩ�� As String, lngҩ�� As Long
    Dim bytDay As Byte, strIDs As String
    
    On Error GoTo errH
    If str��� = "4" Then
        strSQL = _
        " Select Distinct" & _
        "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
        "       And B.������� IN([1],3) And B.����ID=C.ID" & _
        "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
        "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
        "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
        "       And ( A.��������ID is NULL Or A.��������ID=[2]   " & _
        "             Or Exists(select 1 From �������Ҷ�Ӧ M where A.��������ID=M.����ID And M.����ID=[2]))" & _
        "       And A.�շ�ϸĿID=[3]" & _
        " Order by B.�������,C.����"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", int��Χ, lng���˿���ID, lng��Ŀid)
        If Not rsTmp.EOF Then
            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID    '3:�����û�У��򷵻ص�һ�����õ�ִ�п���(��ҽ��վ��ͬ)
            '1:ȱʡΪָ����(ҽ����)ִ�п���,�����Ƿ�����ڲ��˿���
            rsTmp.Filter = "ִ�п���ID=" & lng���ϲ���ID
            '2:�����ɷ����ڲ��˿��ҵ�ִ�п���
            If rsTmp.EOF Then
                '2.0 ��������д���ȱʡ��ִ�п���,��ȱʡΪ����ָ����ȱʡ����
                If lng����ȱʡִ�п��� <> 0 Then
                    rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                    If Not rsTmp.EOF Then
                            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                    End If
                End If
                
                '2.1:����ȱʡΪ���˿���
                If lng���ϲ���ID <> lng���˿���ID Then
                    rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˿���ID
                End If
                '2.2:����ȱʡΪ���˲���
                If rsTmp.EOF Then
                    If lng���˲���ID <> 0 And lng���˲���ID <> lng���˿���ID And lng���˲���ID <> lng���ϲ���ID Then
                        rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˲���ID
                    End If
                End If
            End If
            '2.3:�ɷ����ڲ��˿��ҵ�һ��ִ�п���
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        If gbln���뷢ҩ Then Exit Function
        
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = glng��ҩ��
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = glng��ҩ��
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = glng��ҩ��
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strҩ��, int��Χ, lng���˿���ID, lng��Ŀid, bytDay)
        If Not rsTmp.EOF Then
            strIDs = ""
            If lng����ȱʡִ�п��� <> 0 Then
                rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                If Not rsTmp.EOF Then
                        Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                End If
            End If
            rsTmp.Filter = "��������ID=" & lng���˿���ID
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
            For i = 1 To rsTmp.RecordCount
                strIDs = strIDs & "," & rsTmp!ִ�п���ID '�ռ����ڶ�̬����
                If i = 1 Or rsTmp!ִ�п���ID = lngҩ�� Then
                    Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                    If rsTmp!ִ�п���ID = lngҩ�� Then
                        strIDs = "": Exit For
                    End If
                End If
                rsTmp.MoveNext
            Next
            strIDs = Mid(strIDs, 2)
            If UBound(Split(strIDs, ",")) <= 0 Then strIDs = ""
         End If
    Else
        Select Case intִ�п�������
            Case 0 '0-����ȷ����
                '1 ������Ŀѡ���Ҵ���ȱʡ��ִ�п��ҵ� ������Ŀ��ִ�в���ID
                If lng����ȱʡִ�п��� <> 0 Then
                    Get�շ�ִ�п���ID = lng����ȱʡִ�п���: Exit Function
                End If
                '101736,�ֹ�����ȱʡִ�п���
                '2 �շ���Ŀ.ȱʡ����(�ֹ�����ȱʡִ�п���)
                If int��Χ = 2 Then
                    strSQL = "Select a.ִ�п���id" & vbNewLine & _
                            " From �շ�ִ�п��� A, ���ű� C" & vbNewLine & _
                            " Where a.ִ�п���id + 0 = c.Id And a.�շ�ϸĿid = [1]" & vbNewLine & _
                            "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
                            "       And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null)" & vbNewLine & _
                            "       And a.������Դ = [2] And a.��������id Is Null"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng��Ŀid, 2)
                    If Not rsTmp.EOF Then
                        If Val(Nvl(rsTmp!ִ�п���ID)) <> 0 Then
                            Get�շ�ִ�п���ID = Val(Nvl(rsTmp!ִ�п���ID)): Exit Function
                        End If
                    End If
                    '3 ���˿���
                    If lng���˿���ID <> 0 Then Get�շ�ִ�п���ID = lng���˿���ID: Exit Function
                    '4 ��������
                    If lng��������ID <> 0 Then Get�շ�ִ�п���ID = lng��������ID: Exit Function
                End If
                '5 ����Ա��������ID
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
                "   From �շ�ִ�п��� A,���ű� C" & _
                "   Where A.ִ�п���ID+0=C.ID And  A.�շ�ϸĿID=[1]" & _
                "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng��Ŀid, int��Χ, lng���˿���ID)
                If Not rsTmp.EOF Then
                    If lng����ȱʡִ�п��� <> 0 Then
                         rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                         If Not rsTmp.EOF Then
                                 Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                         End If
                     End If
                    'ȱʡȡ����Ա���ڿ���
                    rsTmp.Filter = "��������ID=" & UserInfo.����ID
                    If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
                    If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
                    If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                End If
            Case 5 'Ժ��ִ��(Ԥ��,������δ��)
            Case 6 '�����˿���
               Get�շ�ִ�п���ID = lng��������ID
        End Select
        If Get�շ�ִ�п���ID = 0 Then Get�շ�ִ�п���ID = UserInfo.����ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal int�����־ As Integer) As String
'���ܣ���鲡����ҽ�������Ƿ���δִ�����(δִ�л�����ִ��)����Ŀ
'���أ�ҽ����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(2,[1],[2],-1,0,[3]) as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitExe", lng����ID, lng��ҳID, int�����־)
    
    If Not rsTmp.EOF Then
        ExistWaitExe = Nvl(rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitDrug(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal lng�����Ժ��ҩ As Long = 0) As String
'���ܣ���鲡����ҩ���Ƿ���δ��ҩ��ҩƷ������
'���أ�ҩ���ͷ��ϲ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(1,[1],[2],-1,[3]) as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitDrug", lng����ID, lng��ҳID, lng�����Ժ��ҩ)
    
    If Not rsTmp.EOF Then
        ExistWaitDrug = Nvl(rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillisAdviceMoney(ByVal strNO As String, ByVal bytFlag As Byte, _
    lngҽ��ID As Long, lng���ͺ� As Long) As Boolean
'���ܣ��ж�һ�ŵ����Ƿ�ҽ���ĸ��ӷ���
'������int��¼����=��ӦסԺ���ü�¼.��¼����
'���أ�ҽ��ID,���ͺ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    lngҽ��ID = 0: lng���ͺ� = 0
    
    On Error GoTo errH
            
    strSQL = "Select ҽ����� From סԺ���ü�¼" & _
        " Where Rownum=1 And ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    If Not rsTmp.EOF Then lngҽ��ID = Nvl(rsTmp!ҽ�����, 0)
    If lngҽ��ID <> 0 Then
        strSQL = "Select ���ͺ� From ����ҽ������" & _
            " Where ҽ��ID=[3] And NO=[1] And ��¼����=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, lngҽ��ID)
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

Public Function PatiCanBilling(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
    '���ܣ����ָ�������Ƿ�������Ȩ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(b.����, a.����) As ����, b.��Ժ����, b.״̬, Nvl(Sum(c.���), 0) As ���,b.�������� " & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ����δ����� C" & vbNewLine & _
            "Where a.����id = [1] And a.����id = b.����id And b.��ҳid = [2] And b.����id = c.����id(+) And b.��ҳid = c.��ҳid(+)" & vbNewLine & _
            "Group By Nvl(b.����, a.����), b.��Ժ����, b.״̬,b.��������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If Val(Nvl(rsTmp!��������)) = 1 And Not (gbln�������� And InStr(strPrivs, ";�������ۼ���;") > 0) Then
            strMsg = """" & rsTmp!���� & """Ϊ�������۲��ˣ���û��Ȩ�޶�����м��ʲ�����"
        End If
        If Val(Nvl(rsTmp!��������)) = 2 And Not (gblnסԺ���� And InStr(strPrivs, ";סԺ���ۼ���;") > 0) Then
            strMsg = """" & rsTmp!���� & """ΪסԺ���۲��ˣ���û��Ȩ�޶�����м��ʲ�����"
        End If
        
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
        
        If IsNull(rsTmp!��Ժ����) And Nvl(rsTmp!״̬, 0) <> 3 Then PatiCanBilling = True: Exit Function
        If InStr(strPrivs, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
            If Nvl(rsTmp!���, 0) <> 0 Then
                strMsg = """" & rsTmp!���� & """�ķ���δ���壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If InStr(strPrivs, ";��Ժ����ǿ�Ƽ���;") = 0 Then
            If Nvl(rsTmp!���, 0) = 0 Then
                strMsg = """" & rsTmp!���� & """�ķ����ѽ��壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    PatiCanBilling = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMaxFact(ByVal strNO As String) As String
'���ܣ���ȡָ�����ʵ��ݷ��������Ʊ�ݺ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    
    'Ӧȡ���һ�δ�ӡ��������
    strSQL = "Select Max(ID) From Ʊ�ݴ�ӡ���� Where ��������=3 And NO=[1]"
    strSQL = "Select Max(����) as ���� From Ʊ��ʹ����ϸ Where Ʊ��=3 And ����=1 And ��ӡID=(" & strSQL & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    If Not rsTmp.EOF Then GetMaxFact = Nvl(rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiHaveStorage(ByVal lng����ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    strSQL = "Select A.���㷽ʽ,Sum(A.���) as ���" & _
        " From ����Ԥ����¼ A,���㷽ʽ B" & _
        " Where A.��¼����=1 And A.���㷽ʽ=B.���� And B.����=5 And A.����ID=[1]" & _
        " Group by A.���㷽ʽ Having Sum(A.���)<>0"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            strMsg = strMsg & vbCrLf & rsTmp!���㷽ʽ & "��" & Format(rsTmp!���, "0.00")
            rsTmp.MoveNext
        Loop
    End If
    If strMsg <> "" Then
        If gbyt���ʼ����տ��� = 1 Then
            If MsgBox("�������´��շ���û���˻����ˣ�" & vbCrLf & strMsg & vbCrLf & vbCrLf & "Ҫ����������?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                PatiHaveStorage = True
            Else
                PatiHaveStorage = False
            End If
        Else
            MsgBox "�������´��շ���û���˻����ˣ�" & vbCrLf & strMsg & vbCrLf & vbCrLf & "���Ƚ������˻��������ٽ��ʡ�", vbInformation, gstrSysName
            PatiHaveStorage = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillIdentical(ByVal strNO As String) As Boolean
'���ܣ��ж�ָ���ļ��ʵ����е�״̬�Ƿ�һ��,���Ƿ�ͬʱ������˺�δ��˵�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    BillIdentical = True
    
    On Error GoTo errH

    strSQL = _
        " Select Count(Distinct �Ǽ�ʱ��) as ʱ����," & _
        " Sum(Decode(��¼״̬,0,1,0)) as δ���," & _
        " Sum(Decode(��¼״̬,0,0,1)) as �����" & _
        " From סԺ���ü�¼" & _
        " Where ��¼״̬ IN(0,1,3) And NO=[1] And ��¼����=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
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

Public Function AuditingWarn(ByVal strPrivs As String, ByRef rsWarn As ADODB.Recordset, _
                            ByVal strNO As String, ByVal str��� As String) As Boolean
'���ܣ���˻��۵�ʱ���Է��ý��б���
'������str���=ָ��������Ҫ��˵��к�,Ϊ�ձ�ʾ������
    Dim rsTmp As ADODB.Recordset
    Dim rsFee As ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, lng����ID As Long
    Dim str����IDs As String, str���s As String
    Dim cur���ն� As Currency, cur��� As Currency, cur��� As Currency
    Dim strWarn As String, intWarn As Integer, blnҽ�� As Boolean
            
    strSQL = "Select A.�����־, A.����, A.����id, A.��ҳid, A.���˲���id ����id," & vbNewLine & _
            "       Decode(B.������, Null, B.������, Zl_Patientsurety(A.����id, A.��ҳid)) ������," & vbNewLine & _
            "       Zl_Patiwarnscheme(A.����id, A.��ҳid) As ���ò���, C.�Ƿ�ҽ�� As ������, A.�շ����, D.���� As �������," & vbNewLine & _
            "       Sum(A.ʵ�ս��) As ���" & vbNewLine & _
            "From סԺ���ü�¼ A, ������Ϣ B, ҽ�Ƹ��ʽ C, �շ���Ŀ��� D" & vbNewLine & _
            "Where A.��¼���� = 2 And A.��¼״̬ = 0 And A.NO = [1] And A.�շ���� = D.���� And A.����id = B.����id And" & vbNewLine & _
            "      B.ҽ�Ƹ��ʽ = C.����(+)" & vbNewLine & _
            IIf(str��� <> "", " And Instr([2],','||Nvl(A.�۸񸸺�,A.���)||',')>0", "") & _
            "Group By Nvl(A.�۸񸸺�, A.���), A.�����־, A.����, A.����id, A.��ҳid, A.���˲���id, B.������, C.�Ƿ�ҽ��, A.�շ����," & vbNewLine & _
            "         D.����" & vbNewLine & _
            "Order By A.����id"
            
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, "," & str��� & ",")
    For i = 1 To rsTmp.RecordCount
        If InStr(str����IDs & ",", "," & rsTmp!����ID & ",") = 0 Then
            str����IDs = str����IDs & "," & rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
            
    If str����IDs <> "" Then
        str����IDs = Mid(str����IDs, 2)
        For i = 0 To UBound(Split(str����IDs, ","))
            lng����ID = Val(Split(str����IDs, ",")(i))
            rsTmp.Filter = "����ID=" & lng����ID
            
            'ȡ�������ͽ��
            str���s = "": cur��� = 0
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
                blnҽ�� = ("" & rsTmp!������ = "1")
                            
                cur���ն� = GetPatiDayMoney(lng����ID)
                Set rsFee = GetMoneyInfo(lng����ID, 0, blnҽ��, 2)
                If Not rsFee Is Nothing Then cur��� = rsFee!Ԥ����� - rsFee!�������
                
                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(1, lng����ID) + cur���
                
                '���౨��
                For j = 0 To UBound(Split(str���s, ","))
                    intWarn = BillingWarn(strPrivs, rsTmp!����, Val("" & rsTmp!����ID), rsTmp!���ò���, rsWarn, _
                        cur���, cur���ն�, cur���, Nvl(rsTmp!������, 0), _
                        Left(Split(str���s, ",")(j), 1), Mid(Split(str���s, ",")(j), 2), strWarn)
                    If intWarn = 2 Or intWarn = 3 Then Exit Function
                Next
            End If
        Next
    End If
    AuditingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckExistsGathering(strNO As String) As Boolean
'����:�ж�ָ�����ʵ��Ƿ����Ӧ�տ�Ľɿ��¼
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select A.NO From ���˽��ʼ�¼ A, ���˽ɿ���� B Where A.NO = [1] And A.ID = B.����id And Rownum = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    CheckExistsGathering = rsTmp.RecordCount > 0
    
Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatientDue(lng����ID As Long) As Currency
'����:��ȡָ�����˵�Ӧ�տ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Zl_Patientdue([1]) Due From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    If rsTmp.RecordCount > 0 Then GetPatientDue = rsTmp!Due
    
Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub LoadPatientBaby(ByRef cboBaby As ComboBox, ByVal lngPatient As Long, lngPatientPage As Long)
    Dim rsTmp As ADODB.Recordset, i As Long
    
    cboBaby.Clear
    cboBaby.AddItem "0-���˱���"
    cboBaby.ItemData(cboBaby.NewIndex) = 0
    Call zlControl.CboSetIndex(cboBaby.hWnd, 0)
    
    If lngPatient <> 0 Then
        Set rsTmp = GetPatientBaby(lngPatient, lngPatientPage)
        With rsTmp
            For i = 1 To .RecordCount
                If Not IsNull(!Ӥ������) Then
                    cboBaby.AddItem !��� & "-" & !Ӥ������
                Else
                    cboBaby.AddItem !��� & "-��" & !��� & "��Ӥ��"
                End If
                cboBaby.ItemData(cboBaby.NewIndex) = !���
                .MoveNext
            Next
        End With
    End If
End Sub

Public Function GetPatientBaby(ByVal lngPatient As Long, lngPatientPage As Long) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select ���, Ӥ������ From ������������¼ Where ����id = [1] And ��ҳID = [2]"
    On Error GoTo errH
    Set GetPatientBaby = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������¼", lngPatient, lngPatientPage)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckValidity(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, Optional ByVal blnAsk As Boolean = True) As Boolean
'���ܣ�����������ϵ����Ч���Ƿ����
'˵����blnAsk=��ʾ�Ƿ�ѯ���Ƿ����,����Ϊ����
    Dim rsTmp As ADODB.Recordset
    Dim Curdate As Date, MinDate As Date
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng�ⷿID)
    If Not rsTmp.EOF Then
        strName = rsTmp!����
        Curdate = rsTmp!ʱ��
        MinDate = CDate("3000-01-01")
            
        Do While Not rsTmp.EOF
            If rsTmp!���Ч�� < MinDate Then
                MinDate = rsTmp!���Ч��
            End If
            If Nvl(rsTmp!���, 0) < dbl���� Then
                dbl���� = dbl���� - Nvl(rsTmp!���, 0)
            Else
                dbl���� = 0
            End If
            If dbl���� = 0 Then Exit Do
            rsTmp.MoveNext
        Loop

        If Curdate > MinDate Then
            If blnAsk Then
                If MsgBox("��������""" & strName & """�����Ч��""" & Format(MinDate, "yyyy-MM-dd") & """�ѹ���,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckValidity = False
                End If
            Else
                MsgBox "���ѣ�" & vbCrLf & vbCrLf & "��������""" & strName & """�����Ч��""" & Format(MinDate, "yyyy-MM-dd") & """�ѹ��ڡ�", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckRecalcRecord(ByVal strNO As String, Optional ByVal byt������Դ As Byte) As Boolean
'���ܣ��ж�ָ�����˵�ָ�������Ƿ���ڰ��ѱ�����ĳ����¼(����Ϊ0�ļ�¼)
'��Σ�
'   byt������Դ 0-סԺ,1-����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(A.ID) Num" & vbNewLine & _
            "From סԺ���ü�¼ A," & vbNewLine & _
            "     (Select ����id, ��ҳid, ���˲���id, ���˿���id, �շ�ϸĿid, ������Ŀid, ��������id, ִ�в���id, ����ʱ��" & vbNewLine & _
            "       From סԺ���ü�¼" & vbNewLine & _
            "       Where NO = [1] And ���ʷ��� = 1" & vbNewLine & _
            "       Group By ����id, ��ҳid, ���˲���id, ���˿���id, �շ�ϸĿid, ������Ŀid, ��������id, ִ�в���id, ����ʱ��) B" & vbNewLine & _
            "Where A.��¼���� = 2 And A.���� = 0 And A.����id = B.����id And A.��ҳid = B.��ҳid And" & vbNewLine & _
            "      A.���˲���id + 0 = B.���˲���id And A.���˿���id + 0 = B.���˿���id And A.�շ�ϸĿid + 0 = B.�շ�ϸĿid And" & vbNewLine & _
            "      A.������Ŀid + 0 = B.������Ŀid And A.��������id + 0 = B.��������id And A.ִ�в���id + 0 = B.ִ�в���id And" & vbNewLine & _
            "      A.����ʱ�� = B.����ʱ��"
    If byt������Դ = 1 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    If rsTmp.RecordCount > 0 Then CheckRecalcRecord = rsTmp!Num > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckNONegative(ByVal strNO As String, Optional ByVal bytType As Byte = 2, _
    Optional ByVal byt������Դ As Byte) As Boolean
'���ܣ��ж�ָ�������Ƿ����������ϸ
'��Σ�
'   byt������Դ 0-סԺ,1-����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 1 From סԺ���ü�¼ Where NO = [1] And ��¼���� = [2] And ��¼״̬ = 1 And ���� < 0"
    If byt������Դ = 1 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytType)
    If rsTmp.RecordCount > 0 Then CheckNONegative = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function inBlackList(ByVal lng����ID As Long) As String
'���ܣ��жϲ����Ƿ��ں�������,�����ؼ���ԭ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '������:47663
    strSQL = "Select ���, ����ԭ�� From ���ⲡ�� Where ����ʱ�� is Null And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    If Not rsTmp.EOF Then inBlackList = rsTmp!��� & "-" & rsTmp!����ԭ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function ReadABCNum(ByVal strPrivs As String) As Boolean
'����:��ȡ��ҩ������
'������strPrivs=���ڸ���Ȩ�޿����Ƿ��ȡС����ݲ���
'���أ���������Ŀ����ĸ
    Dim strSQL As String
        
    On Error GoTo errH
    
    If InStr(strPrivs, ";ҩƷ����С��;") > 0 Then
        strSQL = "Select Upper(����) as ����,��ֵ From ��ҩ������ Order by ����"
    Else
        strSQL = "Select Upper(����) as ����,��ֵ From ��ҩ������ Where Trunc(��ֵ)=��ֵ Order by ����"
    End If
    
    Set grsABCNum = New ADODB.Recordset 'Filter��Newʱ���
    Call zlDatabase.OpenRecordset(grsABCNum, strSQL, "mdlInExse")
    
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


Public Function GetExamineItem(ByVal strItems As String, ByVal lngMediCareID As Long) As ADODB.Recordset
'����:����ָ��������շ���ĿҪ�������ļ�¼��
'����:strItems-�շ�ϸĿID��,����:"2369,2367,2368"
'     lngMediCareID-����,����:901
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select A.�շ�ϸĿid" & vbNewLine & _
            "From ����֧����Ŀ A ,Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B" & vbNewLine & _
            "Where A.���� = [1] And A.Ҫ������ = 1 And A.�շ�ϸĿid = B.Column_Value"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngMediCareID, strItems)
    
    Set GetExamineItem = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRowByFeeItemID(ByRef ObjBillDetails As BillDetails, ByRef lngItemID As Long, Optional lngPatientID As Long) As Long
'����:�����շ���ĿID�������ڵ����е��к�,������ظ���,ֻ���ص�һ��
    Dim i As Long
    
    For i = 1 To ObjBillDetails.Count
        If lngItemID = ObjBillDetails(i).�շ�ϸĿID And (ObjBillDetails(i).����ID = lngPatientID Or lngPatientID = 0) Then
            GetRowByFeeItemID = i: Exit Function
        End If
    Next
End Function

Public Function CheckExamine(ByRef ObjBillDetails As BillDetails, _
    ByRef rsMedAudit As ADODB.Recordset, _
    ByRef lngMediCareID As Long, _
    Optional ByVal str���� As String = "") As Boolean
    '����:���ݸ������շ���Ŀ���󼯺Ͳ���������Ŀ��¼�������Ӧ���շ���Ŀ�Ƿ���Ҫ����
    '���:str����-Ϊ��ʱ,��ʾΪ��ǰ����,��Ϊ��ʱ,��������ʾ����
    Dim i As Long, j As Long, strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    For i = 1 To ObjBillDetails.Count
        strTmp = strTmp & "," & ObjBillDetails(i).�շ�ϸĿID
    Next
    Set rsTmp = GetExamineItem(Mid(strTmp, 2), lngMediCareID)
    
    strTmp = ""
    For i = 1 To rsTmp.RecordCount
        rsMedAudit.Filter = "��ĿID=" & rsTmp!�շ�ϸĿID
        If rsMedAudit.RecordCount = 0 Then
            strTmp = strTmp & "," & GetRowByFeeItemID(ObjBillDetails, rsTmp!�շ�ϸĿID)
        ElseIf Not IsNull(rsMedAudit!��������) Then
            j = GetRowByFeeItemID(ObjBillDetails, rsTmp!�շ�ϸĿID)
            If ObjBillDetails(j).���� * ObjBillDetails(j).���� * IIf(gblnסԺ��λ, ObjBillDetails(j).Detail.סԺ��װ, 1) > rsMedAudit!�������� Then
                MsgBox IIf(str���� <> "", "����:" & str���� & "��", "") & "��" & j & "���շ���Ŀ�����γ�������׼��ʹ������" & FormatEx(rsMedAudit!�������� / IIf(gblnסԺ��λ, ObjBillDetails(j).Detail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                CheckExamine = False: Exit Function
            End If
        End If
        rsTmp.MoveNext
    Next
    
    If strTmp <> "" Then
        If str���� <> "" Then
            MsgBox "����:" & str���� & "�ڵ�" & Mid(strTmp, 2) & "���շ���ĿҪ������,��δ����׼ʹ��!", vbInformation, gstrSysName
        Else
            MsgBox "��" & Mid(strTmp, 2) & "���շ���ĿҪ������,��ǰ����δ����׼ʹ��!", vbInformation, gstrSysName
        End If
        CheckExamine = False: Exit Function
    End If
    CheckExamine = True
End Function

Public Function CheckFeeItemLimitDept(ByVal lngFeeItem As Long, ByVal lngPatientUnit As Long, ByVal lngPatientDept As Long) As Boolean
'����:����շ���Ŀ,���������,�Ƿ������ڵ�ǰ���˿��һ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select ����id From �շ����ÿ��� Where ��Ŀid = [1] And (Select Count(����id) From �շѴ�����Ŀ Where ����id = [1]) > 0"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItem)
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!����ID = lngPatientUnit Or rsTmp!����ID = lngPatientDept Then
                CheckFeeItemLimitDept = True
                Exit For
            End If
            rsTmp.MoveNext
        Next
    Else
        CheckFeeItemLimitDept = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillBeforIN(ByVal strNO As String) As Boolean
'���ܣ������ʵ��Ƿ����ڱ���סԺ֮ǰ
    Dim rsTmp As ADODB.Recordset, strSQL As String
     
    strSQL = "Select 1" & vbNewLine & _
        "From ���˽��ʼ�¼ A, ������Ϣ B" & vbNewLine & _
        "Where A.NO = [1] And A.��¼״̬ = 1 And A.����id = B.����id And B.��Ժʱ�� > A.�շ�ʱ�� And B.��Ժʱ�� Is Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    CheckBillBeforIN = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillExistReplenishData(intTYPE As Integer, Optional lngBalance As Long, Optional strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ���ڶ��ν���
    '����:True-���ڶ��ν������� False-�����ڶ��ν�������
    '���:intType:0-�շ����ݣ�ʹ��lngBalanceΪ�������
    '     intType:1-�շ����ݣ�ʹ��strNosΪ���ݺ�
    '����:������
    '����:2014-10-15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If intTYPE = 0 Then
        strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From ���ò����¼ A, (Select Distinct ����id From ����Ԥ����¼ Where ������� = [1]) B" & vbNewLine & _
        " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2 And Rownum < 2"
        strSQL = strSQL & " Union " & _
        " Select 1 From ���ò����¼ Where ������� = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", lngBalance)
    Else
        strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From ���ò����¼ A," & vbNewLine & _
        "      (Select Distinct ����id" & vbNewLine & _
        "       From ������ü�¼" & vbNewLine & _
        "       Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(f_Str2list([1])))) B" & vbNewLine & _
        " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2 And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����ν���", strNos)
    End If
    If rsTmp.EOF Then
        CheckBillExistReplenishData = False
    Else
        CheckBillExistReplenishData = True
    End If
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
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
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


Public Function zlCreateFeeListStruc(ByRef rsFeelists As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������صķ��ü�¼���ṹ
    '���:
    '����:rsFeelists-���ر��ؼ�¼���ṹ,ͬʱ���˼�¼����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-05 16:18:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set rsFeelists = New ADODB.Recordset
    rsFeelists.Fields.Append "�������", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "�ѱ�", adVarChar, 50, adFldIsNullable
    rsFeelists.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsFeelists.Fields.Append "ʵ��Ʊ��", adVarChar, 20, adFldIsNullable
    rsFeelists.Fields.Append "����ʱ��", adDBTimeStamp, , adFldIsNullable
    rsFeelists.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsFeelists.Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
    rsFeelists.Fields.Append "���㵥λ", adVarChar, 50, adFldIsNullable
    '69788:���ϴ�,2014-6-5,�����������ֶδ�С����20��Ϊ100
    rsFeelists.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsFeelists.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "����", adDouble, , adFldIsNullable
    rsFeelists.Fields.Append "����", adDouble, , adFldIsNullable
    rsFeelists.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    rsFeelists.Fields.Append "ͳ����", adCurrency, , adFldIsNullable
    rsFeelists.Fields.Append "����֧������ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "�Ƿ�ҽ��", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "���ձ���", adVarChar, 50, adFldIsNullable
    rsFeelists.Fields.Append "ժҪ", adVarChar, 4000, adFldIsNullable
    rsFeelists.Fields.Append "�Ƿ���", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "��������ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "ִ�в���ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "���ν���", adDouble, , adFldIsNullable
    rsFeelists.CursorLocation = adUseClient
    rsFeelists.LockType = adLockOptimistic
    rsFeelists.CursorType = adOpenStatic
    rsFeelists.Open
    zlCreateFeeListStruc = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

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
    Err = 0: On Error GoTo ErrHand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
ErrHand:
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
    Err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
ErrHand:
End Function



Public Function ImportWholeSet(frmParent As Object, ByVal intInsure As Integer, rsSel As ADODB.Recordset, Optional lng����ID As Long = 0, _
     Optional blnסԺ��λ As Boolean, Optional lng��������ID As Long = 0, Optional bytӤ���� As Byte, _
     Optional int�����־ As Integer, Optional bln�Ӱ�Ӽ� As Boolean = False, _
     Optional ByVal lngUnitID As Long, Optional int��Χ As Integer, _
     Optional str������ As String = "", Optional str������ As String = "", _
     Optional lngPatiNums As Long = 1, Optional blnNurseStation As Boolean = False, _
    Optional ByVal strҩƷ�۸�ȼ� As String, _
    Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ�۸�ȼ� As String, _
    Optional ByVal lng��ҳID As Long, Optional ByVal lng����ID As Long, Optional ByVal lng����ID As Long) As ExpenseBill
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���õ��ݵ����ݶ�����
    '���:rsSel-ѡ�еĳ�����Ŀ
    '       lngUnitID    ��ǰ��������ID
    '      int��Χ=1.����,2-סԺ
    '      lngPatiNums-������(����������Ч)
    '����:
    '����:��ŵ�����Ϣ�ĵ��ݶ���
    '����:���˺�
    '����:2010-09-02 16:17:54
    '˵��:��Ϊ������ʱ��Ŀ�۸���Ϣ��������,���Է�������������¼���
    '       ��������ͣ���շ�ϸĿ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue(0 To 10) As String, strSubItem As String, str�շ�ϸĿID As String, j As Long
    Dim rsItems As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lng���˿���ID As Long, strժҪ As String
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset, rsPrice As ADODB.Recordset
    Dim rsMoney As ADODB.Recordset
    Dim lngDoUnit As Long
    Dim i As Long, intCurNo As Integer
    Dim int��� As Integer, blnDo As Boolean, dblStock As Double
    Dim blnLoad As Boolean, strSQL As String, strҩ��IDs As String, strͣ����Ŀ��� As String, strPrivs As String
    Dim curModiMoney As Currency
    
    Dim dblAllTime As Double, dblCurTime As Double, dbl�Ӱ�Ӽ��� As Double, dblPriceSingle As Double, lngLastPati As Long
    Dim colSerial As New Collection, dblPrice As Double
    Dim bytType As Byte '0-����;1-סԺ;2-�����סԺ
    Dim strTable  As String, strWherePriceGrade As String
    
    On Error GoTo errH
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
       "   Select A.����id, A.����id, A.���д���, A.�������� " & _
       "   From �շѴ�����Ŀ A, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.����id = D.�շ�ϸĿid "
    Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInExse", strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    strSubItem = Mid(strSubItem, 11)
    strTable = " Select [13] as ����ID,�շ�ϸĿID From (" & strSubItem & ")"
    
    gstrSQL = "" & _
    " Select  X.ҩƷID,W.����ID,W.��������," & _
    "       G.�ѱ�,F.����,F.�Ա�,F.����,F.������," & _
    "       G.��Ժ���� as ����,F.סԺ�� as ��ʶ��,F.����ID,G.��ҳID,G.��ǰ����ID as ���˲���ID,G.��Ժ����ID as ���˿���ID," & _
    "       G.��������,B.��� as �շ����,A.�շ�ϸĿID," & _
    "       B.���㵥λ,B.���,C.���� as �������,B.����,Nvl(H.����,B.����) as ����,E1.���� as ��Ʒ��,B.���,Nvl(B.�Ƿ���,0) as �Ƿ���,B.�Ӱ�Ӽ�," & _
    "       B.���ηѱ�,B.˵��,B.ִ�п���,B.�������, B.��������  ��������,D.�ּ�,D.ԭ��,D.ȱʡ�۸�,D.������ĿID as ������ID,E.���� as ������Ŀ," & _
    "       E.�վݷ�Ŀ as �ַ�Ŀ,D.�Ӱ�Ӽ���,D.�����շ���,Nvl(W.����ID,X.ҩ��ID) as ҩ��ID," & _
    "       Decode(B.���,'4',1,X.סԺ��װ) as סԺ��װ,Decode(B.���,'4',B.���㵥λ,X.סԺ��λ) as סԺ��λ," & _
    "       Decode(b.���,'4',Nvl(W.���÷���,0),Nvl(X.ҩ������,0)) as ����,B.¼������, " & _
    "       M1.���� as ���Ʊ���,M1.���� as ��������,X.��ҩ��̬,x.����ϵ��,M1.���㵥λ as ������λ" & _
    "   From  (" & strTable & ") A ,�շ���ĿĿ¼ B,�շ���Ŀ��� C,�շѼ�Ŀ D,������Ŀ E,������Ϣ F, " & _
    "          ������ҳ G,�շ���Ŀ���� H,�շ���Ŀ���� E1,�������� W,ҩƷ��� X,������ĿĿ¼ M1" & _
    " Where  A.�շ�ϸĿID=D.�շ�ϸĿID And A.�շ�ϸĿID=B.ID " & _
    "       And b.���=C.���� And A.�շ�ϸĿID=X.ҩƷID(+) and X.ҩ��ID=M1.ID(+) And A.�շ�ϸĿID=W.����ID(+) And D.������ĿID=E.ID" & _
    "       And A.�շ�ϸĿID=H.�շ�ϸĿID(+) And H.����(+)=1 And H.����(+)=[12]" & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
    "       And A.����ID=F.����ID(+) And F.����ID=G.����ID(+) And " & IIf(lng��ҳID <> 0, " G.��ҳID(+) = [17]", " F.��ҳID=G.��ҳID(+) ") & _
    "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
    "       And Sysdate Between D.ִ������ And Nvl(D.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) " & _
            strWherePriceGrade
    
    If Not gbln���뷢ҩ Then
        gstrSQL = "Select * From (" & gstrSQL & ")"
    Else
        '���뷢ҩʱ�ſ�ʱ�ۺͷ���ҩƷ������
        gstrSQL = "Select * From (" & gstrSQL & ") Where Not( Instr(',5,6,7,',�շ����)>0 And (����=1 Or �Ƿ���=1))"
    End If
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, "mdlExse", strValue(0), strValue(1), strValue(2), strValue(3), _
        strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10), _
        IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1), lng����ID, strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ�۸�ȼ�, lng��ҳID)
    'û�м�¼���ǿյ���
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    
    With rsSel
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
NextRecord: Do While Not .EOF
            '����շ���Ŀ�Ƿ�ͣ�û���������ﲡ��
            '����ͣ��ʱ,��������
            rsItems.Filter = "�շ�ϸĿID=" & Val(Nvl(!�շ�ϸĿID))
            If rsItems.EOF Then 'δ�ҵ�.������
                 .MoveNext
                GoTo NextRecord:
            End If
            If InStr(",5,6,7,", rsItems!�շ����) = 0 Then
                If InStr(1, strͣ����Ŀ��� & ",", "," & !�������� & ",") > 0 Then
                    .MoveNext
                    GoTo NextRecord
                Else
                    If Not CheckFeeItemAvailable(!�շ�ϸĿID, 2) Then
                        strͣ����Ŀ��� = strͣ����Ŀ��� & "," & !���
                        MsgBox "�����շ���Ŀ�еĵ�" & !��� & "���շ���Ŀ:" & rsItems!���� & "" & vbCrLf & _
                            "��ͣ�û��ٷ����ڲ���,�����ᱻ����." & IIf(IsNull(!��������), "����д�����Ŀ,Ҳ���ᱻ����.", ""), vbInformation, gstrSysName
                        .MoveNext
                        GoTo NextRecord
                    End If
                End If
            Else
                If blnNurseStation Then
                    MsgBox "�����շ���Ŀ�еĵ�" & !��� & "���շ���Ŀ:" & rsItems!���� & "" & vbCrLf & _
                        "ΪҩƷ��Ŀ,��ʿվ��������ʱ�����ᱻ����.", vbInformation, gstrSysName
                    .MoveNext
                    GoTo NextRecord
                End If
            End If
            
            If i = 1 Then
                objBill.NO = ""
                objBill.����ID = Val(Nvl(rsItems!����ID))
                objBill.��ҳID = Val(Nvl(rsItems!��ҳID))
                objBill.����ID = IIf(lng����ID = 0, Val(Nvl(rsItems!���˲���ID)), lng����ID)
                objBill.����ID = IIf(lng����ID = 0, Val(Nvl(rsItems!���˿���id)), lng����ID)
                objBill.���� = Nvl(rsItems!����)
                objBill.�Ա� = Nvl(rsItems!�Ա�)
                objBill.���� = Nvl(rsItems!����)
                objBill.��ʶ�� = Val(Nvl(rsItems!��ʶ��))
                objBill.���� = "" & rsItems!����
                objBill.�ѱ� = Nvl(rsItems!�ѱ�)
                objBill.�����־ = int�����־
                objBill.�Ӱ��־ = IIf(bln�Ӱ�Ӽ�, 1, 0)
                objBill.Ӥ���� = bytӤ����
                objBill.��������ID = lng��������ID
                objBill.������ = str������
                objBill.������ = str������
                objBill.����Ա��� = UserInfo.���
                objBill.����Ա���� = UserInfo.����
                objBill.����ʱ�� = zlDatabase.Currentdate   ' !����ʱ��
                objBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
                objBill.�ಡ�˵� = 0
                
            End If
            '�����շ�ϸĿ=====================================================
            Set objBillDetail = New BillDetail
            Set objBillDetail.Detail = New Detail
                
            '������źʹ�������
            intCurNo = intCurNo + 1
            objBillDetail.��� = intCurNo
            colSerial.Add Array(Val(Nvl(!�շ�ϸĿID)), intCurNo), "_" & !���  '��¼ԭ������ڵ��к�
            objBillDetail.�������� = Nvl(!��������, 0) '��Ϊ������������,�ȼ�¼ԭ����,�����ٴ���
            
            objBillDetail.����ID = Val(Nvl(rsItems!����ID))
            objBillDetail.��ҳID = Val(Nvl(rsItems!��ҳID))
            objBillDetail.Ӥ���� = 0
            objBillDetail.����ID = Val(Nvl(rsItems!���˲���ID))
            objBillDetail.����ID = Val(Nvl(rsItems!���˿���id))
            objBillDetail.���� = Nvl(rsItems!����)
            objBillDetail.�Ա� = Nvl(rsItems!�Ա�)
            objBillDetail.���� = Nvl(rsItems!����)
            objBillDetail.סԺ�� = Val(Nvl(rsItems!��ʶ��))
            objBillDetail.���� = "" & rsItems!����
            objBillDetail.�ѱ� = Nvl(rsItems!�ѱ�)
            objBillDetail.������ = Val(Nvl(rsItems!������))
            
            'Ŀǰ�����ڼ��ʱ�
            objBillDetail.ҽ�Ƹ��� = Get����ҽ�Ƹ��ʽ(objBillDetail.����ID, objBillDetail.��ҳID)
            
            objBillDetail.�շ���� = Nvl(rsItems!�շ����)
            objBillDetail.�շ�ϸĿID = Val(Nvl(!�շ�ϸĿID))
            objBillDetail.���㵥λ = Nvl(rsItems!���㵥λ)
            objBillDetail.���� = IIf(Val(Nvl(!����)) = 0, 1, Val(Nvl(!����)))
            If InStr(",5,6,7,", rsItems!�շ����) > 0 And blnסԺ��λ Then
                objBillDetail.���� = Nvl(!����, 0) / Nvl(rsItems!סԺ��װ, 1)
            Else
                objBillDetail.���� = Nvl(!����, 0)
            End If
            
            objBillDetail.ԭʼ���� = objBillDetail.���� * objBillDetail.����
            objBillDetail.��ҩ���� = ""
            
            objBillDetail.���ӱ�־ = 0 ' IIf(IsNull(!���ӱ�־), 0, !���ӱ�־)
            '���ĺ�ҩƷ����
            '����ִ�п���ȱʡΪ���˲���,�������ָ����,��Ϊָ������
            If objBillDetail.�շ���� = "4" Then
                lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, objBillDetail.����ID)
                If lngDoUnit = 0 Then lngDoUnit = lng��������ID
            End If
            
            '���˿���ID
            lng���˿���ID = objBillDetail.����ID
            If lng���˿���ID = 0 Then lng���˿���ID = lng��������ID
            objBillDetail.Detail.ִ�п��� = IIf(IsNull(rsItems!ִ�п���), 0, rsItems!ִ�п���)
            objBillDetail.ִ�в���ID = Val(Nvl(!ִ�п���ID))
            lngDoUnit = Get�շ�ִ�п���ID(objBillDetail.�շ����, objBillDetail.�շ�ϸĿID, _
                 objBillDetail.Detail.ִ�п���, lng���˿���ID, lng��������ID, int��Χ, lngDoUnit, objBillDetail.����ID, objBillDetail.ִ�в���ID)
            
            objBillDetail.ִ�в���ID = lngDoUnit

            If InStr(",5,6,7,", rsItems!�շ����) > 0 And gbln���뷢ҩ Then
                objBillDetail.ִ�в���ID = 0
            End If
            objBillDetail.ԭʼִ�в���ID = objBillDetail.ִ�в���ID
            
            objBillDetail.Detail.ID = !�շ�ϸĿID
            objBillDetail.Detail.���� = Nvl(rsItems!����)
            objBillDetail.Detail.��� = (Val(Nvl(rsItems!�Ƿ���)) = 1)
            objBillDetail.Detail.�������� = 0
            objBillDetail.Detail.���д��� = 0
            
            If Not gbln���뷢ҩ And InStr(",4,5,6,7,", rsItems!�շ����) > 0 Then
                dblStock = GetStock(Val(Nvl(!�շ�ϸĿID)), objBillDetail.ִ�в���ID)
            Else
                dblStock = 0
            End If
            
            If InStr(",5,6,7,", rsItems!�շ����) > 0 And gbln���뷢ҩ Then
                strҩ��IDs = Decode(rsItems!�շ����, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                If strҩ��IDs <> "" Then dblStock = GetMultiStock(!�շ�ϸĿID, strҩ��IDs)
            End If
            If InStr(",5,6,7,", rsItems!�շ����) > 0 And blnסԺ��λ Then dblStock = dblStock / Nvl(rsItems!סԺ��װ, 1)
            objBillDetail.Detail.��� = dblStock
            
            
            If objBillDetail.�������� <> 0 Then
                'A.����id, A.����id, A.���д���, A.�������� "
                rsOthers.Filter = "����ID=" & colSerial("_" & !��������)(0) & " And ����ID=" & objBillDetail.�շ�ϸĿID
                If Not rsOthers.EOF Then
                    objBillDetail.Detail.�������� = Val(Nvl(rsOthers!��������))
                    objBillDetail.Detail.���д��� = Val(Nvl(rsOthers!���д���))
                End If
            End If
            
            objBillDetail.Detail.��� = Nvl(rsItems!���)
            objBillDetail.Detail.���㵥λ = Nvl(rsItems!���㵥λ)
            
            objBillDetail.Detail.סԺ��λ = Nvl(rsItems!סԺ��λ)
            objBillDetail.Detail.סԺ��װ = Val(Nvl(rsItems!סԺ��װ))
            
            objBillDetail.Detail.�Ӱ�Ӽ� = 0 ' (IIf(IsNull(!�Ӱ�Ӽ�), 0, !�Ӱ�Ӽ�) = 1)
            objBillDetail.Detail.��� = Nvl(rsItems!���)
            objBillDetail.Detail.������� = Nvl(rsItems!�������)
            objBillDetail.Detail.���� = Nvl(rsItems!����)
            objBillDetail.Detail.��Ʒ�� = Nvl(rsItems!��Ʒ��)
            objBillDetail.Detail.���ηѱ� = (Val(Nvl(rsItems!���ηѱ�)) = 1)
            objBillDetail.Detail.˵�� = ""
            objBillDetail.Detail.������� = IIf(IsNull(rsItems!�������), 0, rsItems!�������)
            objBillDetail.Detail.���� = IIf(IsNull(rsItems!��������), "", rsItems!��������)
            objBillDetail.Detail.�������� = Nvl(rsItems!��������)
            
            If InStr(",5,6,7,", rsItems!�շ����) > 0 Then
                objBillDetail.Detail.����ְ�� = Get����ְ��(objBillDetail.Detail.ID)
                objBillDetail.Detail.�������� = Get��������(objBillDetail.Detail.ID)
            End If
            objBillDetail.Detail.¼������ = Val(Nvl(rsItems!¼������))
            objBillDetail.Detail.ҩ��ID = Val(Nvl(rsItems!ҩ��ID))
            objBillDetail.Detail.��� = Val(Nvl(rsItems!�Ƿ���)) = 1
            objBillDetail.Detail.���� = Val(Nvl(rsItems!����)) = 1
            objBillDetail.Detail.�������� = Val(Nvl(rsItems!��������)) = 1
            objBillDetail.Detail.Ҫ������ = 0
            objBillDetail.Detail.��ҩ��̬ = Val(Nvl(rsItems!��ҩ��̬))
            objBillDetail.Detail.������λ = Nvl(rsItems!������λ)
            objBillDetail.Detail.����ϵ�� = Val(Nvl(rsItems!����ϵ��))
         
            '����:41136
            strժҪ = objBillDetail.ժҪ
'            If lng����ID <> 0 Then '90304
                strժҪ = gclsInsure.GetItemInfo(intInsure, lng����ID, objBillDetail.�շ�ϸĿID, strժҪ, 2, , "|1")
                objBillDetail.ժҪ = strժҪ
'            Else
'                objBillDetail.ժҪ = ""
'            End If
             '����۸񲿷�=====================================================
             rsItems.MoveFirst
            Set objBillDetail.InComes = New BillInComes
            Do While Not rsItems.EOF
                '�������еļ۸��������¼���'***
                If Val(Nvl(rsItems!�Ƿ���)) = 1 Then
                    If InStr(",5,6,7,", rsItems!�շ����) > 0 Or (rsItems!�շ���� = "4" And Nvl(rsItems!��������, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        'ʱ��ҩƷ����۸�(�����ɲ�����)
                        dblAllTime = Val(Nvl(!����)) * IIf(Val(Nvl(!����)) = 0, 1, Val(Nvl(!����))) * lngPatiNums
                        If dblAllTime <> 0 Then
                            dblPrice = Getʱ��ҩƷӦ�ս��(objBillDetail.ִ�в���ID, CLng(Nvl(!�շ�ϸĿID)), dblAllTime, gstrDec, dblPriceSingle)
                            If dblAllTime <> 0 Then
                                If Val(Nvl(!����)) = 0 Then
                                    '����δ�ֽ����
                                    If rsItems!�շ���� = "4" Then
                                        MsgBox "ʱ����������""" & Nvl(rsItems!����) & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    Else
                                        MsgBox "ʱ��ҩƷ""" & Nvl(rsItems!����) & """��治��,�޷�����۸�", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.��׼���� = 0
                                Else
                                    objBillIncome.��׼���� = Val(Nvl(!����))
                                End If
                            Else
                                'ע�⣺���������ֻ�ܱ���4λС��,�Ҳ���������,������Ҫ�ֹ�����;�����������ڼ��㾫������������
                                objBillIncome.��׼���� = IIf(dblPriceSingle = 0, Format(dblPrice / (Val(Nvl(!����))), gstrFeePrecisionFmt), dblPriceSingle)  '�������ۼۼ۸�
                            End If
                        Else
                            objBillIncome.��׼���� = 0
                        End If
                        '----------------------------------------------------------------------------------------------
                    Else
                        
                        If Abs(Val(Nvl(!����))) > Val(Nvl(rsItems!�ּ�)) Or Abs(Val(Nvl(!����))) = 0 Then
                            objBillIncome.��׼���� = Val(Nvl(rsItems!ȱʡ�۸�))
                        Else
                            objBillIncome.��׼���� = Val(Nvl(!����))
                        End If
                    End If
                Else
                    objBillIncome.��׼���� = Val(Nvl(rsItems!�ּ�))
                End If

                If InStr(",5,6,7,", rsItems!�շ����) > 0 And blnסԺ��λ Then
                    objBillIncome.��׼���� = Format(objBillIncome.��׼���� * Nvl(rsItems!סԺ��װ, 1), gstrFeePrecisionFmt)
                Else
                    objBillIncome.��׼���� = Format(objBillIncome.��׼����, gstrFeePrecisionFmt)
                End If
                
                objBillIncome.�ּ� = Val(Nvl(rsItems!�ּ�))  '�ּ�ԭ�۶�ҩƷ�������
                objBillIncome.ԭ�� = Val(Nvl(rsItems!ԭ��))
                objBillIncome.������ĿID = Val(Nvl(rsItems!������ID))
                objBillIncome.������Ŀ = Nvl(rsItems!������Ŀ)
                objBillIncome.�վݷ�Ŀ = Nvl(rsItems!�ַ�Ŀ)
                
                'Ӧ�ս��=����*����*����
                If Val(Nvl(rsItems!�Ƿ���)) = 1 And (InStr(",5,6,7,", rsItems!�շ����) > 0 Or rsItems!�շ���� = "4" And Nvl(rsItems!��������, 0) = 1) Then
                    objBillIncome.Ӧ�ս�� = dblPrice '��֤Ӧ�ս�������۽��û�����
                Else
                    objBillIncome.Ӧ�ս�� = objBillIncome.��׼���� * objBillDetail.���� * objBillDetail.����
                End If
                
                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If bln�Ӱ�Ӽ� And Val(Nvl(rsItems!�Ӱ�Ӽ�)) = 1 Then
                    dbl�Ӱ�Ӽ��� = Val(Nvl(rsItems!�Ӱ�Ӽ�)) / 100
                    objBillIncome.Ӧ�ս�� = objBillIncome.Ӧ�ս�� + objBillIncome.Ӧ�ս�� * dbl�Ӱ�Ӽ���
                End If
                objBillIncome.Ӧ�ս�� = Format(objBillIncome.Ӧ�ս��, gstrDec)
                
                '����ʵ�ս��
                If lng����ID = 0 Then   '��������(�������),���Դ˴�������ʵ�ս��
                    objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                Else
                    If Val(Nvl(rsItems!���ηѱ�)) = 1 Then
                        objBillIncome.ʵ�ս�� = objBillIncome.Ӧ�ս��
                    Else
                        objBillIncome.ʵ�ս�� = ActualMoney(objBillDetail.�ѱ�, Val(Nvl(rsItems!������ID)), objBillIncome.Ӧ�ս��, _
                            objBillDetail.�շ�ϸĿID, objBillDetail.ִ�в���ID, objBillDetail.ԭʼ����, dbl�Ӱ�Ӽ���)
                    End If
                End If
                With objBillIncome
                    objBillDetail.InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��
                End With
                '�ж���һ����¼�Ƿ����ڵ�ǰ��
                int��� = !���
                i = i + 1
                rsItems.MoveNext
            Loop
            
            With objBillDetail
                objBill.Details.Add .Detail, .�շ�ϸĿID, .���, .��������, .����ID, .��ҳID, .����ID, .����ID, .����, .�Ա�, .����, .סԺ��, .����, _
                    .�ѱ�, .��������, .�շ����, .���㵥λ, .��ҩ����, .����, .����, .���ӱ�־, .ִ�в���ID, .InComes, .���￨��, , .������, .ҽ�Ƹ���, , , , .ժҪ, .ԭʼ����, .ԭʼִ�в���ID, .Ӥ����
                '���뷢ҩʱ,Key����Ϊ1,��ʾ�༭ʱִ�п����в��ɽ���
                If InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ Then
                    objBill.Details(objBill.Details.Count).Key = 1
                End If
            End With
            .MoveNext
        Loop
    End With
     '�����´����������
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).�������� <> 0 Then
            objBill.Details(i).�������� = colSerial("_" & objBill.Details(i).��������)(1)
        End If
    Next
    Set ImportWholeSet = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub zlExecPrintSingleBill(ByVal frmMain As Object, ByVal lng����ID As Long, _
    ByVal strPrivs As String, Optional str��ֹ���� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡָ�����˵Ĵ߿
    '����:���˺�
    '����:2010-10-29 16:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�߿��� As Double, bytType As Byte
    If frmPatiPressMoney.zlPatiPressMoney(frmMain, glngModul, strPrivs, 0, "", lng����ID, IIf(bytType = 2, 1, 2)) = False Then Exit Sub
End Sub
Public Sub zlPrintBedCard(ByVal frmMain As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ��ͷ��
    '����:���˺�
    '����:2010-10-29 17:10:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_2", frmMain) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_2", frmMain, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, 2)
    End If
End Sub

Public Sub zlPrintDayDetail(ByVal frmMain As Object, ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, _
    Optional bln��ʾ�˷� As Boolean = False, Optional bln��ʾ��� As Boolean = False, Optional bln����ʱ�� As Boolean = True, _
    Optional lng��ҳID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡһ���嵥
    '����:int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    '����:���˺�
    '����:2010-10-29 17:16:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If int���� = 1 Then
        With frmDailyPrint
            .mlng����ID = lng����ID
            .mlng����ID = lng����ID
            .mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pһ���嵥)
            .Show 1, frmMain
        End With
        Exit Sub
    End If
    If Not frmDailyListAsk Is Nothing Then Unload frmDailyListAsk
    frmDailyListAsk.mlngModul = 1141    '��Ȼ��һ���嵥ģ��Ĳ���Ϊ׼
    frmDailyListAsk.mbytInFun = 1
    frmDailyListAsk.mlng����ID = lng����ID
    frmDailyListAsk.mlngPageID = lng��ҳID
    frmDailyListAsk.Show vbModal, frmMain
    If frmDailyListAsk.mblnAskOk Then
        ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", frmMain, "����ID=" & lng����ID, _
            "��ʼʱ��=" & Format(frmDailyListAsk.mdatBegin, "YYYY-MM-DD HH:MM:SS"), _
            "����ʱ��=" & Format(frmDailyListAsk.mdatEnd, "YYYY-MM-DD HH:MM:SS"), _
            "��ʾ�˷�=" & IIf(bln��ʾ�˷�, "1", "0"), _
            "��ʾ�����=" & IIf(bln��ʾ���, "1", "0"), _
            "���˲���=" & lng����ID, _
            "��ҳID=" & frmDailyListAsk.mlngPageID, _
            "����ʱ��=" & IIf(bln����ʱ��, "����ʱ��", "�Ǽ�ʱ��"), 1
    End If
End Sub

Public Sub zlPrintAccountPage(ByVal frmMain As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ��ҳ
    '����:���˺�
    '����:2010-11-01 10:03:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ReportPrintSet gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1139_2", frmMain
End Sub
Public Function zlGetPatiInsure(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��������Ϣ
    '����:ҽ��������Ϣ��Ϣ��
    '����:���˺�
    '����:2010-11-01 10:09:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = _
        " Select A.�Ǽ�ʱ��, B.����, E.����, Nvl(E.ҽ����, D.��Ϣֵ) As ҽ����,b.��������" & vbNewLine & _
        " From ������Ϣ A, ������ҳ B, ������ҳ�ӱ� D, ҽ�����˵��� E, ҽ�����˹����� F" & vbNewLine & _
        " Where B.����id = [1] And B.��ҳid = [2] And A.����id = B.����id And B.����id = D.����id(+)" & _
        "       And B.��ҳid = D.��ҳid(+) And D.��Ϣ��(+) = 'ҽ����' And" & vbNewLine & _
        "       A.����id = F.����id(+) And F.��־(+) = 1 And F.ҽ���� = E.ҽ����(+)" & _
        "       And F.���� = E.����(+) And F.���� = E.����(+)"
    On Error GoTo errH
    Set zlGetPatiInsure = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��������Ϣ��", lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlPreBalance(ByVal frmMain As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ԥ���㹦��
    '�ɹ�:����true,���򷵻�False
    '����:���˺�
    '����:2010-11-01 10:08:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int���� As Integer, strҽ���� As String, str���� As String
    Dim rsTmp As ADODB.Recordset, str������� As String
    Dim blnDateMoved As Boolean, dat�Ǽ�ʱ�� As Date, bln�������� As Boolean
    
    zlPreBalance = False
    Set rsTmp = zlGetPatiInsure(lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        With rsTmp
            int���� = Val(!����)
            strҽ���� = "" & !ҽ����
            str���� = "" & !����
            dat�Ǽ�ʱ�� = !�Ǽ�ʱ��
            bln�������� = Val(Nvl(!��������)) = 1
        End With
    End If
    
    If int���� = 0 Then
        MsgBox "��ȡ����ҽ�������Ϣʧ��!", vbExclamation, gstrSysName
        Exit Function
    End If
    If gclsInsure.GetCapability(support����_�������ú���ýӿ�, lng����ID, int����) Then
        MsgBox "��ҽ���ӿڲ�֧�ֽ�������ǰԤ����!", vbExclamation, gstrSysName
        Exit Function
    End If
    blnDateMoved = zlDatabase.DateMoved(dat�Ǽ�ʱ��, , , "��ȡ��ʷ��Ϣ")
    Screen.MousePointer = 11
    If bln�������� Then
        Set rsTmp = GetVBalance(0, "������ý���", int����, lng����ID, , , , , blnDateMoved)
    Else
        Set rsTmp = GetVBalance(1, "סԺ���ý���", int����, lng����ID, , , , , blnDateMoved)
    End If
    Screen.MousePointer = 0
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ò���û��δ���ʵı�����Ŀ����!", vbInformation, gstrSysName
    Else
        str������� = gclsInsure.WipeoffMoney(rsTmp, lng����ID, strҽ����, "0", int����, "|0") '������;����
        MsgBox "Ԥ����ɹ�!" & str�������, vbInformation, gstrSysName '�ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
        zlPreBalance = True
    End If
End Function

Public Sub zlPreBalanceAll(ByVal frmMain As Object, ByVal lng����ID As Long)
    Dim rsTemp As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim str������� As String, i As Integer, strSQL As String
    Dim lng����ID As Long, int���� As Integer, blnDateMoved As Boolean
    Dim strҽ���� As String, str���� As String, str���� As String, str�Ǽ�ʱ�� As Date, bln�������� As Boolean
    
    Err = 0: On Error GoTo ErrHand:
    If lng����ID = 0 Then
        MsgBox "δѡ����,���ܽ�������Ԥ��,��ѡ��һ������!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    strSQL = "" & _
    "   Select distinct A.����ID, B.��ҳID,B.סԺ��,B.��Ժ���� as ����, A.�Ǽ�ʱ��,B.����, " & vbNewLine & _
    "       E.����, Nvl(b.����, a.����) As ����,E.ҽ����,b.��������" & vbNewLine & _
    "   From ������Ϣ A, ������ҳ B, ҽ�����˵��� E, ҽ�����˹����� F,��Ժ���� C " & vbNewLine & _
    "   Where A.����ID = B.����ID  And Nvl(B.��ҳID, 0) <> 0 " & vbNewLine & _
    "               And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3 And  A.����ID=C.����ID  " & _
    "               And A.����ID = F.����ID(+) And F.��־(+) = 1 " & vbNewLine & _
    "               And F.ҽ���� = E.ҽ���� And F.���� = E.���� And F.���� = E.����(+)   " & vbNewLine & _
    "             " & vbNewLine & _
      IIf(lng����ID = 0, " Order by ����,סԺ�� Desc", "  And B.��ǰ����ID =[1] And C.����ID=[1] Order by ����,LPAD(����,10,' ')")
      
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ԤԼ����", lng����ID)
    If rsTemp.EOF Then
        MsgBox IIf(lng����ID = 0, "", "��ǰ����") & "û�з�����Ժ��ҽ�����ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("�ò�������" & IIf(lng����ID = 0, "���в���", "��ǰ����") & "�е�������Ժҽ������(����" & rsTemp.RecordCount & "��)����Ԥ����," & _
        vbCrLf & "����ܻỨ�ѽϳ���ʱ��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    With rsTemp
        Do While Not .EOF
            str���� = Nvl(!����)
            lng����ID = Val(Nvl(!����ID))
            int���� = Val(Nvl(!����))
            strҽ���� = Nvl(!ҽ����)
            str���� = Nvl(!����)
            str�Ǽ�ʱ�� = Format(!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
            bln�������� = Val(Nvl(!��������)) = 1
            
            If Not gclsInsure.GetCapability(support����_�������ú���ýӿ�, lng����ID, int����) Then
                blnDateMoved = zlDatabase.DateMoved(str�Ǽ�ʱ��, , , "����Ԥ��")
                Call zlCommFun.ShowFlash("���ڴ���ҽ������""" & str���� & """ ...", frmMain)
                If Not frmMain Is Nothing Then frmMain.Refresh
                If bln�������� Then
                    Set rsTmp = GetVBalance(0, "������ý���", int����, lng����ID, , , , , blnDateMoved)
                Else
                    Set rsTmp = GetVBalance(1, "סԺ���ý���", int����, lng����ID, , , , , blnDateMoved)
                End If
                If Not rsTmp Is Nothing Then
                    If Not rsTmp.RecordCount = 0 Then
                        str������� = gclsInsure.WipeoffMoney(rsTmp, lng����ID, strҽ����, "0", int����, "|0") '������;����
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    Call zlCommFun.StopFlash
    MsgBox "Ԥ��ɹ�!", vbInformation + vbOKOnly, gstrSysName
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Function zlCheckPatiFeeRenewValied(ByVal lng����ID As Long, _
    lng��ҳID As Long, lng����ID As Long, lng����ID As Long, _
    ByRef str���ת��ʱ�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡�˲����Ƿ񳬹�ʱ��
    '����:str���ת��ʱ��-���ת��ʱ��(yyyy-mm-dd hh:mm:ss)
    '����:true-�Ϸ���¼����;False-���ܲ�¼����
    '����:���˺�
    '����:2010-12-10 11:04:04
    '����:33744
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtDate As Date, strTemp As String
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select ��ֹʱ��,��ֹԭ��" & _
    "   From (  Select ��ֹʱ��,��ֹԭ�� From ���˱䶯��¼  " & _
    "               Where ����id = [1] and ��ҳID=[2] and ����ID=[3] and ����ID=[4]  " & _
    "                           And (��ֹԭ�� = 3 or ��ֹԭ��=15 or ��ֹԭ��=10 or ��ֹԭ��=1)  " & _
    "               Order By ��ֹʱ�� Desc, ��ʼԭ��) " & _
    "   Where Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鲹�ѵ�ʱ�޼��", lng����ID, lng��ҳID, lng����ID, lng����ID)
    If rsTemp.EOF Then rsTemp.Close: Exit Function
    
    dtDate = CDate(Format(rsTemp!��ֹʱ��, "yyyy-mm-dd HH:MM:SS"))
    If dtDate + 1 / 24 * gTy_System_Para.int���ݲ�¼ʱ�� < zlDatabase.Currentdate Then
        If gTy_System_Para.int���ݲ�¼ʱ�� = 0 Then
            strTemp = IIf(Val(Nvl(rsTemp!��ֹԭ��)) = 10, "Ԥ��Ժ", IIf(Val(Nvl(rsTemp!��ֹԭ��)) = 1, "��Ժ", "��ת�ƻ�ת����"))
            ShowMsgbox "ע��:" & vbCrLf & "    �ò���" & strTemp & ",ϵͳ����Ϊ��������в�¼���ò�����"
        Else
            ShowMsgbox "ע��:" & vbCrLf & "    �ò���ת�ƻ�ת�����Ѿ�������" & gTy_System_Para.int���ݲ�¼ʱ�� & "Сʱ,���ܽ��в�¼����!"
        End If
        rsTemp.Close: Exit Function
        Exit Function
    End If
    rsTemp.Close
    str���ת��ʱ�� = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    zlCheckPatiFeeRenewValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlExecBilling_Mulit(ByVal int���� As Integer, _
    ByVal frmMain As Object, _
    ByVal lng����ID As Long, _
    ByVal lng����ID As Long, bln��Ժ As Boolean, ByVal bln���� As Boolean, _
    Optional strUnitIDs As String = "", Optional lng��ҳID As Long = 0, _
    Optional bln���� As Boolean = False, Optional lng����ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ʵ�
    '���:int����- 0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS);6-���ò�ѯ����
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-01 11:01:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim lngModule As Long, strPrivs As String
    
    zlExecBilling_Mulit = False
    If InStr(GetInsidePrivs(Enum_Inside_Program.pסԺ����), "���в���") = 0 Then
        If strUnitIDs = "" Then
            '���»�ȡ����Ա�����ڲ���
            strUnitIDs = GetUserUnits
        End If
        If InStr("," & strUnitIDs & ",", "," & lng����ID & ",") = 0 And bln���� = False Then
            MsgBox "��û�����в�����Ȩ�ޣ����ܶ����������Ĳ��˼��ʣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If int���� = 6 Then
        If lng����ID = 0 Then
            MsgBox "δѡ���������ʵĲ��������ܽ����������ʣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    lng����ID = 0
    If bln���� Then
        If InStr(1, "012", int����) > 0 Then '    int���� = 0 - ҽ��վ����, 1 - ��ʿվ����, 2 - ҽ��վ����(PACS / LIS))
            lng����ID = lng����ID
        End If
    End If
    If lng����ID = 0 And lng����ID = 0 Then
        MsgBox "δѡ���������ʵĲ�����ҽ�����ţ����ܽ����������ʣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Err.Clear: On Error Resume Next
    
    
    gbytBilling = 0
    lngModule = Enum_Inside_Program.pסԺ����
    strPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
    
    If frmChargeBat.ShowMe(frmMain, lngModule, strPrivs, 0, lng����ID, lng����ID, lng����ID, bln����) = False Then Exit Function
    zlExecBilling_Mulit = True
 
End Function

Public Function zlExecBilling(ByVal int���� As Integer, ByVal frmMain As Object, _
    ByVal lng����ID As Long, _
    ByVal lng����ID As Long, bln��Ժ As Boolean, ByVal bln���� As Boolean, _
    Optional strUnitIDs As String = "", Optional lng��ҳID As Long = 0, _
    Optional bln���� As Boolean = False, Optional lng����ID As Long = 0, _
    Optional lngҽ��ID As Long = 0, Optional ByVal bln�������۲��� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ļ��ʵ�
    '���:int����- 0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS);9-���ò�ѯ����
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-01 11:01:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, str���ת��ʱ�� As String
    
    zlExecBilling = False
    If InStr(GetInsidePrivs(Enum_Inside_Program.pסԺ����), "���в���") = 0 Then
        If strUnitIDs = "" Then
            '���»�ȡ����Ա�����ڲ���
            strUnitIDs = GetUserUnits
        End If
        If InStr("," & strUnitIDs & ",", "," & lng����ID & ",") = 0 And bln���� = False Then
            MsgBox "��û�����в�����Ȩ�ޣ����ܶ����������Ĳ��˼��ʣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    lng����ID = 0
    If bln���� And int���� <> 6 Then
        If InStr(1, "012", int����) > 0 Then '    int���� = 0 - ҽ��վ����, 1 - ��ʿվ����, 2 - ҽ��վ����(PACS / LIS))
                lng����ID = lng����ID
        End If
        '���Ѽ���Ƿ񳬹�ʱ��
        If zlCheckPatiFeeRenewValied(lng����ID, lng��ҳID, lng����ID, lng����ID, str���ת��ʱ��) = False Then Exit Function
    End If
    
    '��Ժ���˼���Ȩ��
    If bln��Ժ Then
        If bln���� And InStr(GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "��Ժ����ǿ�Ƽ���") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷����Ѿ�����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Function
        ElseIf Not bln���� And InStr(GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "��Ժδ��ǿ�Ƽ���") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷�����δ����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�������۲��˵����������
    If bln�������۲��� Then
        If Not (gbln�������� And InStr(GetInsidePrivs(Enum_Inside_Program.p���ʲ���), ";�������ۼ���;") > 0) Then
            MsgBox "��û��Ȩ�޶��������۲��˽��м��ʲ�����", vbInformation, gstrSysName
            Exit Function
        End If
        zlExecBilling = ZLShowChargeWindow(frmMain, 2, 0, lng����ID, lng��ҳID, _
            lng����ID, lng����ID, bln����, lngҽ��ID, str���ת��ʱ��)
        Exit Function
    End If
    
    Err.Clear: On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����)
    frmCharge.mbytInState = 0
    frmCharge.mbytUseType = 0
    frmCharge.mlngDeptID = lng����ID
    frmCharge.mlngUnitID = lng����ID
    frmCharge.mlngModule = 1133
    frmCharge.mlng����ID = lng����ID
    frmCharge.mbln���� = bln����
    frmCharge.mlng����ҽ�� = lngҽ��ID
    frmCharge.mlng��ҳID = lng��ҳID
    frmCharge.mstr���ת��ʱ�� = str���ת��ʱ��
    frmCharge.Show IIf(frmMain Is Nothing, 0, 1), frmMain
    If gblnOK Then zlExecBilling = True
End Function

Public Function ZLShowChargeWindow(frmMain As Object, _
    ByVal bytFun As Byte, ByVal bytInState As Byte, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lngDeptID As Long, ByVal lngUnitID As Long, _
    ByVal bln���� As Boolean, ByVal lng����ҽ�� As Long, _
    ByVal str���ת��ʱ�� As String, Optional ByVal strInNO As String) As Boolean
    '����������ù���
    '��Σ�
    '   bytFun 0-�շ�,1-����,2-�������
    '   bytInState 0-ִ��(���޸�),1-���,2-����,3-�˷�(�շѡ����ʲ����˷�),4-�����շ�;5-�쳣��������;11-���Ƶ���
    '   lngUnitID As Long '��ǰ���ʲ���,Ϊ0ʱ��ʾ���в���
    '   lngDeptID As Long '��ǰ���ʿ���,Ϊ0ʱ��ʾ���п���
    '   bln���� As Boolean '33744
    '   strInNO ���뵥�ݣ����˺͸��Ƶ���ʱ���루���������ʱ��Ч��
    Dim strCommon As String, intAtom As Integer, blnOk As Boolean
    
    On Error Resume Next
    If gobjCharge Is Nothing Then
        Set gobjCharge = CreateObject("zl9OutExse.clsOutExse")
        If gobjCharge Is Nothing Then Exit Function
    End If
    
    Err.Clear: On Error GoTo 0
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    blnOk = gobjCharge.Charge(frmMain, gcnOracle, glngSys, gstrDBUser, bytFun, bytInState, lng����ID, lng��ҳID, _
        lngDeptID, lngUnitID, bln����, lng����ҽ��, str���ת��ʱ��, strInNO)
    Call GlobalDeleteAtom(intAtom)
    ZLShowChargeWindow = blnOk
End Function

Public Function ZlIsOutpatientObserve(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '�ж��Ƿ�Ϊ�������۲���
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    strSQL = "Select 1 From ������ҳ Where �������� = 1 And ����id = [1] And ��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�Ϊ�������۲���", lng����ID, lng��ҳID)
    ZlIsOutpatientObserve = Not rsTemp.EOF
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlWrite_Off_ApplyAndVerfy(ByVal frmMain As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                                            ByVal bln���� As Boolean, Optional ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��������������
    '���:bln����:true�������룬false-�������
    '����:
    '����:����ɹ�,����ture,���򷵻�False
    '����:���˺�
    '����:2010-11-01 11:31:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If lng����ID = 0 Then
        MsgBox "��ѡ���˲�����", vbInformation, gstrSysName
        Exit Function
    End If
    With frmReCharge
        .mbytFun = IIf(bln����, 0, 1)
        .mbytUseType = 0
        .mlngDeptID = lng����ID
        .mlngPatientID = lng����ID
        .mstrPrivs = GetInsidePrivs(Enum_Inside_Program.pסԺ����, True)
        .mstrInNO = strNO
        .Show 1, frmMain
    End With
    If gblnOK Then zlWrite_Off_ApplyAndVerfy = True
End Function
Public Function zlGet�շ����() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ����
    '����:���˺�
    '����:2010-11-25 14:22:18
    '����:34260
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If grs�շ���� Is Nothing Then
        gstrSQL = "Select ����,���,ϵͳ��־,�����༭ From �շ����"
        Set grs�շ���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ����")
    ElseIf grs�շ����.State <> 1 Then
        gstrSQL = "Select ����,���,ϵͳ��־,�����༭ From �շ����"
        Set grs�շ���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ����")
    End If
    Set zlGet�շ���� = grs�շ����
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckPatiIsDeath(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ����Ѿ�����.
    '���:
    '����:
    '����:�Ѿ�����,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-22 14:32:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select 1 From ������ҳ  where ����ID=[1] and ��Ժ��ʽ like '%����%' and RowNum <=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ��Ѿ�����", lng����ID)
    zlCheckPatiIsDeath = Not rsTemp.EOF
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckPatiIsMemo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡�˱�ע��Ϣ�Ƿ����
    '����:�������,����true,���򷵻�Flase
    '����:���˺�
    '����:2010-12-24 09:43:14
    '����:34763
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select 1 From ���˱�ע��Ϣ where ����ID=[1] and nvl(��ҳID,0)=[2] and Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鲡���Ƿ���ڲ��˱�ע��Ϣ", lng����ID, lng��ҳID)
    zlCheckPatiIsMemo = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCallPatiMemoWriteAndRead(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByRef objInPati As Object, Optional blnOnlyReadMemo As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò��˱�ע�޸ĺ���ʾ�ӿ�
    '���:objInPati-���˲���
    '       lng����ID-����ID
    '       blnOnlyReadMemo-��ֻ��,���ܱ༭(�ݲ���,�Ժ���ܴ��ڵ���)
    '����:
    '����:���óɹ��򲻴��ڲ�����Ϣ����,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-24 09:50:03
    '����:34763
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If objInPati Is Nothing Then
        Set objInPati = CreateObject("zl9InPatient.clsInPatient")
    End If
    If objInPati Is Nothing Then zlCallPatiMemoWriteAndRead = True: Exit Function
    Err = 0: On Error GoTo errHandle
    'zlPatiMemoReadAndWrite(ByVal frmMain As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String, Optional ByVal blnEdit As Boolean = False)
    Call objInPati.zlPatiMemoReadAndWrite(frmMain, gcnOracle, lng����ID, lng��ҳID, strPrivs)      ' , Not blnOnlyReadMemo
    zlCallPatiMemoWriteAndRead = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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

Public Sub zlRptControlToVsGrid(ByVal objRpt As ReportControl, ByRef objGrid As VSFlexGrid)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��rptControl������װ�䵽����ؼ���
    '���:objRpt-ReportControl
    '     intPrintType-��ӡ����
    '����:���˺�
    '����:2011-01-31 13:19:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, objRow As ReportRow
    On Error GoTo errHandle
    objGrid.Cols = 1: objGrid.Rows = 2
    With objRpt
        j = 0
        For i = 0 To .Columns.Count - 1
            If .Columns(i).Visible Then
                objGrid.TextMatrix(0, j) = .Columns(i).Caption
                objGrid.ColKey(j) = Trim(objGrid.TextMatrix(0, j))
                objGrid.ColWidth(j) = .Columns(i).Width * Screen.TwipsPerPixelX
                Select Case .Columns(i).Alignment
                Case xtpAlignmentCenter
                    objGrid.ColAlignment(j) = flexAlignCenterCenter
                Case xtpAlignmentLeft
                    objGrid.ColAlignment(j) = flexAlignCenterCenter
                Case xtpAlignmentRight
                    objGrid.ColAlignment(j) = flexAlignRightCenter
                End Select
                objGrid.FixedAlignment(j) = flexAlignCenterCenter
                objGrid.Cols = objGrid.Cols + 1
                j = j + 1
            End If
        Next
        For Each objRow In .Rows
            If objRow.GroupRow = False Then
                For j = 0 To .Columns.Count - 1
                    If .Columns(j).Visible Then
                        '����65471,������:����������֮��,�������������ͷ�����ϵ�����
                        objGrid.TextMatrix(objGrid.Rows - 1, objGrid.ColIndex(.Columns(j).Caption)) = objRow.Record(.Columns(j).ItemIndex).Value
                    End If
                Next
                objGrid.Rows = objGrid.Rows + 1
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlGet���ʳ���ID(ByVal lng����ID As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ʳ���ID
    '���:lng����ID
    '����:
    '����:���ʳ�����ID
    '����:���˺�
    '����:2011-02-10 11:59:36
    '����:35554
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������")
    If rsTemp.EOF Then
        zlGet���ʳ���ID = 0
    Else
        zlGet���ʳ���ID = rsTemp!ID '�������ݵ�ID
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetMultiNOs(ByVal strNO As String, Optional lng��ӡID As Long, Optional blnNOMoved As Boolean) As String
'���ܣ�����һ���շѵ��ݵ�NO������ͬһ�δ�ӡ�Ķ���NO
'������blnNoMoved�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
'���أ���ʽ��"'AAA','BBB','CCC',..."
'      ���ָ����"lng��ӡID",�򷵻�
'˵�������ڶ൥���շ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strNos As String
    Dim i As Long
    
    On Error GoTo errH
            
    lng��ӡID = 0
    
    'Ӧ�������һ�δ�ӡ���������
    strSQL = "Select ID,NO From " & IIf(blnNOMoved, "H", "") & "Ʊ�ݴ�ӡ���� Where ��������=1" & _
        " And ID=(Select Max(ID) From " & IIf(blnNOMoved, "H", "") & "Ʊ�ݴ�ӡ���� Where ��������=1 And NO=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNO)
    If Not rsTmp.EOF Then
        lng��ӡID = Nvl(rsTmp!ID, 0) '����û��
        For i = 1 To rsTmp.RecordCount
            strNos = strNos & ",'" & rsTmp!NO & "'"
            rsTmp.MoveNext
        Next
        GetMultiNOs = Mid(strNos, 2)
    Else
        GetMultiNOs = "'" & strNO & "'"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetDelBalanceID(ByVal strNO As String) As Long
'���ܣ���ȡ�˷Ѽ�¼�Ľ���ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ����ID From ������ü�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=2 And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNO)
    If Not rsTmp.EOF Then GetDelBalanceID = Val("" & rsTmp!����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_GetInvoicePrintFormat(ByVal lngModule As Long, _
    Optional strʹ����� As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ӡ��ʽ
    '����:int��ӡ��ʽ-��ӡ��ʽ(0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ)
    '����:��ӡ��ʽ(���)
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
    Dim lngFormat As Long, lngFormat1 As Long
    
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    strShareTypeFormat = Trim(zlDatabase.GetPara("���ʷ�Ʊ��ʽ", glngSys, lngModule, ""))
    '��ʽ:ʹ�����1,��ʽ1|ʹ�����2,��ʽ2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,", ",")
        lngFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
        If Trim(varTemp(0)) = strʹ����� Then
            zl_GetInvoicePrintFormat = lngFormat: Exit Function
        End If
    Next
    zl_GetInvoicePrintFormat = lngFormat1
End Function

Public Function zl_GetInvoicePrintMode(ByVal lngModule As Long, _
    Optional strʹ����� As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ӡ��ʽ
    '����:int��ӡ��ʽ-��ӡ��ʽ()
    '����:0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
     Dim intPrintMode As Long, intPrintMode1 As Long
    '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
    strShareTypeFormat = Trim(zlDatabase.GetPara("���˽��ʴ�ӡ", glngSys, lngModule, ""))
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
Public Function zlisCheckOperatorICU() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ����Ա�Ƿ�ΪICU���ŵ���Ա
    '����:��ICU,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-05 23:29:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = _
    " Select 1" & _
    " From  ��������˵�� B,������Ա C" & _
    " Where  B.����ID=C.����ID And B.��������='ICU' and C.��ԱID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ����Ա�Ƿ�ICU������Ա", UserInfo.ID)
    If rsTemp.RecordCount <> 0 Then
        zlisCheckOperatorICU = True
    End If
    rsTemp.Close
    Set rsTemp = Nothing
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlisCheckDeptICU(ByVal lngDeptID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�������Ƿ�ΪICU����
    '����:��ICU,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-05 23:29:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    strSQL = _
    " Select 1" & _
    " From  ��������˵�� B " & _
    " Where  B.����ID=[1] And B.��������='ICU'  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ����Ա�Ƿ�ICU������Ա", lngDeptID)
    If rsTemp.RecordCount <> 0 Then
        zlisCheckDeptICU = True
    End If
    rsTemp.Close
    Set rsTemp = Nothing
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsAllowFeeChange(lng����ID As Long, lng��ҳID As Long, _
   Optional int״̬ As Integer = -1, Optional str���� As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�������ñ䶯
    '���:int״̬-(-1��ʾ�����ݿ��ж�ȡ��˱�־�����ж�;>0��ʾ,ֱ�Ӹ��ݸ�״̬�����ж�)
    '����:����䶯����true,���򷵻�False
    '����:���˺�
    '����:2012-05-21 15:44:47
    '����:49501,51612
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    If gTy_System_Para.byt������˷�ʽ = 0 And gTy_System_Para.blnδ��ƽ�ֹ���� = False Then
        ''����Ǹ��
        zlIsAllowFeeChange = True: Exit Function
    End If
   
    strSQL = "" & _
    " Select Nvl(��˱�־,0) as ��˱�־,nvl(״̬,0) as ״̬" & _
    " From ������ҳ " & _
    " Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        
        MsgBox "δ�ҵ���Ӧ�Ĳ�����Ϣ" & IIf(str���� <> "", "(����:" & str���� & ")", "") & ",��������м�¼����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '���δ��Ʋ��˲��������
    If gTy_System_Para.blnδ��ƽ�ֹ���� And Val(Nvl(rsTemp!״̬)) = 1 Then
        '51612
        MsgBox "����δ���(" & IIf(str���� <> "", "����:" & str����, "") & "��" & lng��ҳID & "��סԺ) ,���ܶԸò��˽��м��˻����˲�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����ؼ��
    If gTy_System_Para.byt������˷�ʽ = 0 Then zlIsAllowFeeChange = True: Exit Function
    If int״̬ < 0 Then
        int״̬ = Val(Nvl(rsTemp!��˱�־))
    End If
    '������״̬
    If int״̬ = 1 Then
        MsgBox "����" & IIf(str���� <> "", ":" & str����, "") & "�ڵ�" & lng��ҳID & "��סԺ���Ѿ���ʼ��˷���,���ܶԸò��˽��з��ñ䶯��", vbInformation, gstrSysName
        Exit Function
    End If
    If int״̬ = 2 Then
        MsgBox "�Ѿ�����˶Բ���" & IIf(str���� <> "", ":" & str����, "") & "��" & lng��ҳID & "��סԺ���õ����,���ܶԸò��˽��з��ñ䶯��", vbInformation, gstrSysName
        Exit Function
    End If
    zlIsAllowFeeChange = True
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
    

Public Function GetMzBalanceData(lng����ID As Long, Optional strDeptIDs As String, _
    Optional strClass As String, Optional dtStartDate As Date, Optional dtEndDate As Date, _
    Optional strItem As String, _
    Optional blnOnlyYbUpData As Boolean, Optional ByVal blnZero As Boolean, Optional blnDateMoved As Boolean, _
    Optional bytKind As Byte, Optional strChargeType As String = "", _
    Optional bln����ʱ�� As Boolean, Optional strTime As String, Optional strChargeTypeNot As String = "", Optional strDiag As String = "") As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�������������������
    '��Σ�lng����ID-����ID,
    '      strDeptIds����������ID��,"1,2,3",�ձ�ʾ����
    '      strClass��""-���з���(��δ����),"'����1','����2',..."
    '      strItem���վݷ�Ŀ��,'��ҩ��','��ҩ��',...
    '      DateBegin,DateEnd�����ʷ����ڼ�,���Ǽ�ʱ�����ʱ��,ȱʡֵΪCDate("0:00:00")
    '      blnOnlyYbUpData���Ƿ�ֻ�������ϴ�����
    '      blnZero���Ƿ��ȡ�����
    '      blnDateMoved�����˵Ǽ�ʱ���Ƿ���ת������֮ǰ
    '      bytKind��  0-����ͨ����,1-��������,2-��ͨ���ú�������
    '      strChargeType:""��ʾ���з���,����Ϊָ���շ����ķ���;��:5,6,7��  '34260
    '      bln����ʱ��-�ǰ�����ʱ��ͳ��:true-����ʱ��;false-���Ǽ�ʱ��ͳ��
    '      strTime-�������۲������۴���,0,1,2...(0��ʾ�����ۼ�������)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 17:57:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCond As String, strCond2 As String
    Dim strTable As String, strWherePage As String, strConO As String
    Dim strDiagCondition As String
    On Error GoTo errHandle
    strWherePage = IIf(strTime = "", "", " And Instr([8],','||Nvl(A.��ҳID,0)||',')>0")
         
    strCond = " And A.����ID=[1]"
    If Not dtStartDate = CDate("0:00:00") Then
        strConO = strCond
        strCond = strCond & " And " & IIf(Not bln����ʱ��, "A.�Ǽ�ʱ��", "A.����ʱ��") & " Between [2] And [3]"
        dtStartDate = CDate(Format(dtStartDate, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(dtEndDate, "yyyy-MM-dd 23:59:59"))
    End If
    
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([4],','||A.��������ID||',')>0")
    strCond = strCond & IIf(strItem = "", "", " And Instr([5],','''||A.�վݷ�Ŀ||''',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([6],','''||A.�շ����||''',')>0")   '34260
    strCond = strCond & IIf(strChargeTypeNot = "", "", " And Instr([9],','||A.�շ����|| ',')=0")
    
    strConO = strConO & IIf(strDeptIDs = "", "", " And Instr([4],','||A.��������ID||',')>0")
    strConO = strConO & IIf(strItem = "", "", " And Instr([5],','''||A.�վݷ�Ŀ||''',')>0")
    strConO = strConO & IIf(strChargeType = "", "", " And Instr([6],','''||A.�շ����||''',')>0")   '34260
    strConO = strConO & IIf(strChargeTypeNot = "", "", " And Instr([9],','||A.�շ����|| ',')=0")
    
    If Not (strDiag = "" Or strDiag = "�������") Then
        strDiagCondition = " And Exists (Select 1 From �������ҽ�� K,������ϼ�¼ L Where K.ҽ��ID = A.ҽ����� And K.���ID = L.ID And ������� = [10])"
    End If
    
    
    '0-����ͨ����,1-��������,2-��ͨ���ú�������
    If bytKind = 1 Then
        '��������
        strCond = strCond & " And A.�����־=4"
        strConO = strConO & " And A.�����־=4"
    Else
        strCond = strCond & " And A.�����־<>2"
        If bytKind = 0 Then strCond = strCond & " And A.�����־<>4"
        strConO = strConO & " And A.�����־<>2"
        If bytKind = 0 Then strConO = strConO & " And A.�����־<>4"
    End If
    
    strCond2 = strCond   '�Ѿ�����ʵ�,�����Ƿ��ϴ���Ҫȡ,�����Ȱ����������¼����,�ڶ����Ӳ�ѯ��
    If blnOnlyYbUpData Then
        strCond = strCond & " And (A.�Ƿ��ϴ�=1 And A.����ID is Null) "
        strConO = strConO & " And A.�Ƿ��ϴ�=1  "
    Else
        strCond = strCond & " And A.����ID Is Null "
    End If
    
 
     
    'סԺ,����,ʱ��,[���ݺ�],��Ŀ,��Ŀ,Ӥ����,[ID],[���],[��¼����],[��¼״̬],[ִ��״̬],[A.��ҳID],[A.��������ID],[�Ǽ�ʱ��],δ����,���ʽ��,[����]
    If blnZero Then
        strTable = "" & _
        "   SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,1 as ��־,'����' as סԺ," & _
        "           A.��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
        "           Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ, " & _
        "           A.���� * A.���� As ����, A.��׼����,nvl(A.Ӧ�ս��,0) as Ӧ�ս��,nvl(A.ʵ�ս��,0) as ʵ�ս��,Nvl(A.ʵ�ս��,0) as δ����, A.ͳ����," & _
        "           A.��������, A.�շ����,A.�ѱ�, A.ִ�в���id, A.������, A.���մ���id, A.�����־,A.ҽ����� " & _
        "   From ������ü�¼ A  " & _
        "   Where A.��¼״̬<>0 And A.���ʷ���=1" & strCond & strWherePage & _
        "   Union  all" & _
        "   SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,1 as ��־,'����' as סԺ," & _
        "          A.��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
        "          Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
        "          A.���� * A.���� As ����, A.��׼����,nvl(A.Ӧ�ս��,0) as Ӧ�ս��,nvl(A.ʵ�ս��,0) as ʵ�ս��,Nvl(A.ʵ�ս��,0) as δ����, A.ͳ����," & _
        "          A.��������, A.�շ����,A.�ѱ�, A.ִ�в���id,A.������, A.���մ���id, A.�����־,A.ҽ�����" & _
        "   From סԺ���ü�¼ A  " & _
        "   Where A.��¼״̬<>0 And (Mod(A.��¼����,10) = 5) And A.���ʷ���=1" & strCond
    Else
        strTable = "" & _
        " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,1 as ��־,'����' as סԺ," & _
        "       A.��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
        "       Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ, " & _
        "       A.���� * A.���� As ����, A.��׼����,nvl(A.Ӧ�ս��,0) as Ӧ�ս��,nvl(A.ʵ�ս��,0) as ʵ�ս��,Nvl(A.ʵ�ս��,0) as δ����, A.ͳ����," & _
        "       A.��������, A.�շ����,A.�ѱ�, A.ִ�в���id, A.������, A.���մ���id, A.�����־,A.ҽ�����" & _
        " From ������ü�¼ A," & _
        "      (    Select A.NO,A.���,A.��¼����,Nvl(Sum(A.ʵ�ս��),0) as ʵ�ս��" & _
        "           From  ������ü�¼ A " & _
        "           Where A.��¼״̬<>0  And A.���ʷ���=1  And Nvl(A.ʵ�ս��,0) <> 0  And A.����ID Is Null " & strConO & strWherePage & _
        "           Group by A.NO,A.���,A.��¼���� " & _
        "           Having Nvl(Sum(A.ʵ�ս��),0)<>0 ) B " & _
        " Where A.NO=B.NO And A.���=B.��� And A.��¼����=B.��¼���� And A.����ID Is Null" & strCond & strWherePage
        
        strTable = strTable & " Union ALL" & _
        " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬, 1 as ��־,'����' as סԺ," & _
        "        A.��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
        "        Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ, " & _
        "        A.���� * A.���� As ����, A.��׼����,nvl(A.Ӧ�ս��,0) as Ӧ�ս��,nvl(A.ʵ�ս��,0) as ʵ�ս��,Nvl(A.ʵ�ս��,0) as δ����, A.ͳ����," & _
        "        A.��������, A.�շ����,A.�ѱ�, A.ִ�в���id, A.������, A.���մ���id, A.�����־,A.ҽ�����" & _
        " From   סԺ���ü�¼ A ," & _
        "      (    Select A.NO,A.���,A.��¼����,Nvl(Sum(A.ʵ�ս��),0) as ʵ�ս��" & _
        "           From סԺ���ü�¼ A" & _
        "           Where A.��¼״̬<>0 And (Mod(A.��¼����,10) = 5)  And A.���ʷ���=1 And Nvl(A.ʵ�ս��,0)<>0 And A.����ID Is Null" & strConO & _
        "           Group by A.NO,A.���,A.��¼���� " & _
        "           Having Nvl(Sum(A.ʵ�ս��),0)<>0 ) B " & _
        " Where A.NO=B.NO And A.���=B.��� And A.��¼����=B.��¼���� And A.����ID Is Null" & strCond
    End If
     
    strSQL = IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼", 2, "", True, ""), "������ü�¼")
        
    'סԺ���ʴ��ʱȡ��(ԭ����,��ǰΪʲôҪ����,�Դ��Ժ��֤):
    '   And (Nvl(A.ʵ�ս��,0)<>Nvl(A.���ʽ��, 0) Or Nvl(A.���ʽ��, 0)=0)
        
        
    strSQL = "" & _
    "        SELECT 0 as ID,A.NO,A.���,Mod(A.��¼����,10) as ��¼����,Max(A.��¼״̬) As ��¼״̬,A.ִ��״̬," & _
    "              1 as ��־,'����' as סԺ,A.��ҳID," & _
    "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
    "               Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
    "               avg(A.���� * nvl(A.����,1)) As ����, avg(A.��׼����) as ��׼����,Sum(Nvl(A.Ӧ�ս��,0)) as Ӧ�ս��, Sum(Nvl(A.ʵ�ս��,0)) as ʵ�ս��, " & _
    "               Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ����, avg(A.ͳ����) as ͳ����," & _
    "               A.��������, max(A.�շ����) as �շ����, " & _
    "               max(A.�ѱ�) as �ѱ�, max(A.ִ�в���id) as ִ�в���id, max(A.������) as ������, " & _
    "               max( A.���մ���id) as ���մ���id, A.�����־,A.ҽ����� " & _
    "        FROM " & strSQL & " A" & _
    "        Where A.����id Is Not Null And A.��¼״̬<>0 And A.���ʷ���=1 " & _
    "              And Not Exists (Select 1 From ������ü�¼ C, ���˽��ʼ�¼ D Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id And Nvl(d.����״̬, 0) = 1) " & strCond2 & strWherePage & _
    "        Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0 " & _
    "                    Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and  Sum(Nvl(A.���ʽ��,0))=0  And Mod(Count(*),2)=0) " & _
    "                     Or Sum(Nvl(A.���ʽ��, 0))=0 And Sum(Nvl(A.Ӧ�ս��,0))<>0 And Mod(Count(*),2)=0" & _
    "        Group by A.NO,A.���,Mod(A.��¼����,10),A.ִ��״̬, " & _
    "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.�Ǽ�ʱ��,A.��ҳID, " & _
    "               Decode(Nvl(A.Ӥ����,0),0,'','��'),A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,A.��������,A.�����־,A.ҽ�����" & _
    ""
    
    strSQL = strSQL & " Union ALL " & _
    "        SELECT 0 as ID,A.NO,A.���,Mod(A.��¼����,10) as ��¼����,Max(A.��¼״̬) As ��¼״̬,A.ִ��״̬," & _
    "               1 as ��־,'����' as סԺ,A.��ҳID," & _
    "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
    "               Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
    "               avg(A.���� * nvl(A.����,1)) As ����, avg(A.��׼����) as ��׼����,Sum(Nvl(A.Ӧ�ս��,0)) as Ӧ�ս��, Sum(Nvl(A.ʵ�ս��,0)) as ʵ�ս��, " & _
    "               Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ����, avg(A.ͳ����) as ͳ����," & _
    "               A.��������, max(A.�շ����) as �շ����," & _
    "               max(A.�ѱ�) as �ѱ�, max(A.ִ�в���id) as ִ�в���id, max(A.������) as ������,max( A.���մ���id) as ���մ���id, A.�����־,A.ҽ����� " & _
    "        FROM " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼", 2, "", True, ""), "סԺ���ü�¼") & " A" & _
    "        Where A.����id Is Not Null And A.��¼״̬<>0 And A.���ʷ���=1 And (Mod(A.��¼����,10) = 5)  " & _
    "              And Not Exists (Select 1  From סԺ���ü�¼ C, ���˽��ʼ�¼ D Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id And Nvl(d.����״̬, 0) = 1) " & strCond2 & strWherePage & _
    "        Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0 " & _
    "                    Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and  Sum(Nvl(A.���ʽ��,0))=0  And Mod(Count(*),2)=0) " & _
    "                     Or Sum(Nvl(A.���ʽ��, 0))=0 And Sum(Nvl(A.Ӧ�ս��,0))<>0 And Mod(Count(*),2)=0" & _
    "        Group by A.NO,A.���,Mod(A.��¼����,10),A.ִ��״̬,A.��ҳID, " & _
    "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.�Ǽ�ʱ��,Decode(Nvl(A.Ӥ����,0),0,'','��'),A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,A.��������,A.�����־,A.ҽ�����" & _
    ""
 
    strTable = strTable & " Union ALL " & strSQL
    
    strSQL = _
        "Select A.��־,A.סԺ,Nvl(B.����,'δ֪') as ����,A.ʱ��,A.NO as ���ݺ� ,Nvl(E.����,C.����) as ��Ŀ,A.�վݷ�Ŀ as ��Ŀ, A.Ӥ����,A.ID,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,A.��ҳID,A.��������ID,A.�Ǽ�ʱ��," & _
        "      A.����, A.��׼���� as �۸�,nvl(A.Ӧ�ս��,0) as Ӧ�ս��,nvl(A.ʵ�ս��,0) as ʵ�ս��, " & _
        "       Nvl(A.δ����,0) δ����,Nvl(A.δ����,0) ���ʽ��, A.ͳ����," & _
        "       Nvl(A.��������,C.��������) as ����, A.�շ����,M.���� as �շ������," & _
        "       A.�ѱ�, A.ִ�в���id, A.������, A.���մ���id,A.�շ�ϸĿID,C.���㵥λ,A.�����־,Decode(a.��¼״̬, 2, 3, 3, 2, 1) As ����,Max(G.�������) As ���" & _
        " From (  " & strTable & ") A,���ű� B,�շ���ĿĿ¼ C,������Ŀ D,�շ���Ŀ���� E,�շ���Ŀ��� M,�������ҽ�� F,������ϼ�¼ G " & _
        " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID And a.ҽ����� = f.ҽ��id(+) And f.���id = g.Id(+) And A.������ĿID=D.ID" & strDiagCondition & _
                IIf(strClass = "", "", " And Instr([7],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
        "       And A.�շ����=M.����(+) And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        " Group By A.��־,A.סԺ,Nvl(B.����,'δ֪'),A.ʱ��,A.NO ,Nvl(E.����,C.����),A.�վݷ�Ŀ, A.Ӥ����,A.ID,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,A.��ҳID,A.��������ID,A.�Ǽ�ʱ��," & _
        "      A.����, A.��׼����,nvl(A.Ӧ�ս��,0),nvl(A.ʵ�ս��,0), " & _
        "       Nvl(A.δ����,0),Nvl(A.δ����,0), A.ͳ����," & _
        "       Nvl(A.��������,C.��������), A.�շ����,M.����," & _
        "       A.�ѱ�, A.ִ�в���id, A.������, A.���մ���id,A.�շ�ϸĿID,C.���㵥λ,A.�����־,Decode(a.��¼״̬, 2, 3, 3, 2, 1)" & _
        " Order by A.ʱ�� Desc,A.סԺ,A.NO Desc,A.��¼����,A.���"
    
    On Error GoTo errHandle
    Set GetMzBalanceData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����������", lng����ID, dtStartDate, dtEndDate, _
                    "," & strDeptIDs & ",", "," & strItem & ",", "," & strChargeType & ",", "," & strClass & ",", _
                    "," & strTime & ",", "," & strChargeTypeNot & ",", strDiag)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetZYBalanceData(lng����ID As Long, Optional strTime As String, _
    Optional strDeptIDs As String, Optional strClass As String, Optional DateBegin As Date, Optional DateEnd As Date, _
    Optional strBaby As String, Optional strItem As String, _
    Optional blnOnlyYbUpData As Boolean, Optional ByVal blnZero As Boolean, Optional blnDateMoved As Boolean, _
     Optional strChargeType As String = "", Optional bln����ʱ�� As Boolean, Optional strChargeTypeNot As String = "", Optional strDiag As String = "") As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ������δ���ʽ��ϸĿ(��ÿ������Ŀ��)
    '��Σ�lng����ID-����ID,
    '      strTime��סԺ������,"1,2,3"
    '      strDeptIds����������ID��,"1,2,3",�ձ�ʾ����
    '      strClass��""-���з���(��δ����),"'����1','����2',..."
    '      strItem���վݷ�Ŀ��,'��ҩ��','��ҩ��',...
    '      strBaby��0-���з���,1-���˷���,2�Լ���-��bytBaby-1��Ӥ������
    '      DateBegin,DateEnd�����ʷ����ڼ�,���Ǽ�ʱ�����ʱ��,ȱʡֵΪCDate("0:00:00")
    '      blnOnlyYbUpData���Ƿ�ֻ�������ϴ�����
    '      blnZero���Ƿ��ȡ�����
    '      blnDateMoved�����˵Ǽ�ʱ���Ƿ���ת������֮ǰ
    '      bln����ʱ��-�Ƿ񰴷���ʱ��ͳ��,true-����ʱ�䣬false-�Ǽ�ʱ��
    '      strChargeType:""��ʾ���з���,����Ϊָ���շ����ķ���;��:5,6,7��
    '���أ��ɹ�=��¼��,ʧ��=Nothing
    '���ƣ����˺�
    '���ڣ�2010-03-06 13:21:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, strCond As String, strCond2 As String, strConO As String
    Dim strTable As String, bytType As Byte '0-����,1-סԺ,2-�����סԺ
    Dim strWherePage As String 'סԺ��������
    Dim strWhereMzPage As String
    Dim strDiagCondition As String
        
    strCond = " And A.����ID=[1]"
    
    strWherePage = IIf(strTime = "", "", " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0")
    
    If Not DateBegin = CDate("0:00:00") Then
        strConO = strCond
        strCond = strCond & " And " & IIf(Not bln����ʱ��, "A.�Ǽ�ʱ��", "A.����ʱ��") & " Between [3] And [4]"
        DateBegin = CDate(Format(DateBegin, "yyyy-MM-dd 00:00:00"))
        DateEnd = CDate(Format(DateEnd, "yyyy-MM-dd 23:59:59"))
    End If
    
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([5],','||A.��������ID||',')>0")
    strCond = strCond & IIf(strBaby = "", "", " And Instr([6],','|| Nvl(A.Ӥ����,0) ||',')>0")
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.�վݷ�Ŀ||''',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([9],','''||A.�շ����||''',')>0")   '34260
    strCond = strCond & IIf(strChargeTypeNot = "", "", " And Instr([10],','||A.�շ����|| ',')=0")
    strCond = strCond & " And A.�����־ In (2,3) "
    
    strConO = strConO & IIf(strDeptIDs = "", "", " And Instr([5],','||A.��������ID||',')>0")
    strConO = strConO & IIf(strBaby = "", "", " And Instr([6],','|| Nvl(A.Ӥ����,0) ||',')>0")
    strConO = strConO & IIf(strItem = "", "", " And Instr([7],','''||A.�վݷ�Ŀ||''',')>0")
    strConO = strConO & IIf(strChargeType = "", "", " And Instr([9],','''||A.�շ����||''',')>0")   '34260
    strConO = strConO & IIf(strChargeTypeNot = "", "", " And Instr([10],','||A.�շ����|| ',')=0")
    strConO = strConO & " And A.�����־ In (2,3) "
    
    If Not (strDiag = "" Or strDiag = "�������") Then
        strDiagCondition = " And Exists (Select 1 From �������ҽ�� K,������ϼ�¼ L Where K.ҽ��ID = A.ҽ����� And K.���ID = L.ID And ������� = [11])"
    End If
    
    strCond2 = strCond   '�Ѿ�����ʵ�,�����Ƿ��ϴ���Ҫȡ,�����Ȱ����������¼����,�ڶ����Ӳ�ѯ��
    
    
    If blnOnlyYbUpData Then
        strCond = strCond & " And (A.�Ƿ��ϴ�=1 And A.����ID is Null) "
        strConO = strConO & " And A.�Ƿ��ϴ�=1"
    Else
        strCond = strCond & " And A.����ID Is Null "
    End If
  
    
    'סԺ,����,ʱ��,[���ݺ�],��Ŀ,��Ŀ,Ӥ����,[ID],[���],[��¼����],[��¼״̬],[ִ��״̬],[A.��ҳID],[A.��������ID],[�Ǽ�ʱ��],δ����,���ʽ��,[����]
    If blnZero Then
        strTable = "" & _
        "   SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,2 as ��־, '��'||NVL(A.��ҳID,0)||'��' as סԺ," & _
        "           A.��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
        "           Nvl(A.Ӥ����,0) as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
        "           A.���� * A.���� As ����, A.��׼����,nvl(A.Ӧ�ս��,0) as Ӧ�ս��,nvl(A.ʵ�ս��,0) as ʵ�ս��,Nvl(A.ʵ�ս��,0) as δ����, A.ͳ����," & _
        "           A.��������,A.�շ����,A.�ѱ�,A.ִ�в���id,A.������, A.���մ���id,A.ҽ�����" & _
        "   From סԺ���ü�¼ A " & _
        "   Where A.��¼״̬<>0 And A.���ʷ���=1" & strCond & strWherePage & _
        ""
    Else
        strTable = "" & _
            " SELECT  A.ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,2 as ��־,'��'||NVL(A.��ҳID,0)||'��'  as סԺ," & _
            "           A.��ҳID,A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
            "           Nvl(A.Ӥ����,0) as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
            "           A.���� * A.���� As ����,A.��׼����,nvl(A.Ӧ�ս��,0) as Ӧ�ս��,nvl(A.ʵ�ս��,0) as ʵ�ս��,Nvl(A.ʵ�ս��,0) as δ����,A.ͳ����," & _
            "           A.��������,A.�շ����,A.�ѱ�,A.ִ�в���id,A.������,A.���մ���id,A.ҽ�����" & _
            " From  סԺ���ü�¼ A," & _
            "      ( Select A.NO,A.���,A.��¼����, Nvl(Sum(A.ʵ�ս��),0) as ʵ�ս��" & _
            "        From  סԺ���ü�¼ A" & _
            "        Where A.��¼״̬<>0  And A.���ʷ���=1 And Nvl(A.ʵ�ս��,0)<>0  And A.����ID Is Null " & strConO & strWherePage & _
            "        Group by A.NO,A.���,A.��¼���� Having Nvl(Sum(A.ʵ�ս��),0)<>0  ) B " & _
            " Where A.NO=B.NO And A.���=B.��� And A.��¼����=B.��¼���� And A.����ID Is Null" & strCond & strWherePage
    End If
    
    'סԺ���ʴ��ʱȡ��(ԭ����,��ǰΪʲôҪ����,�Դ��Ժ��֤):
    '   And (Nvl(A.ʵ�ս��,0)<>Nvl(A.���ʽ��, 0) Or Nvl(A.���ʽ��, 0)=0)
    
    strSQL = "" & _
    "   SELECT 0 as ID,A.NO,A.���,Mod(A.��¼����,10) as ��¼����,Max(A.��¼״̬) As ��¼״̬,Nvl(A.ִ��״̬,0) As ִ��״̬,2 as ��־," & _
    "             '��'||NVL(A.��ҳID,0)||'��'  as סԺ,A.��ҳID," & _
    "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ʱ��,A.�Ǽ�ʱ��," & _
    "               Nvl(A.Ӥ����,0) as Ӥ����,A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ," & _
    "               avg(A.���� * nvl(A.����,1)) As ����,avg(A.��׼����) as ��׼����,Sum(Nvl(A.Ӧ�ս��,0)) as Ӧ�ս��,Sum(Nvl(A.ʵ�ս��,0)) as ʵ�ս��," & _
    "               Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as δ����,avg(A.ͳ����) as ͳ����," & _
    "               A.��������,max(A.�շ����) as �շ����,max(A.�ѱ�) as �ѱ�,max(A.ִ�в���id) as ִ�в���id,max(A.������) as ������," & _
    "               max( A.���մ���id) as ���մ���id,A.ҽ����� " & _
    "   FROM סԺ���ü�¼ A " & _
    "   Where A.����id Is Not Null And A.��¼״̬<>0 And A.���ʷ���=1 " & _
    "         And Not Exists (Select 1 From סԺ���ü�¼ C, ���˽��ʼ�¼ D Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id And Nvl(d.����״̬, 0) = 1) " & strCond2 & strWherePage & _
    "   Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0 " & _
    "                    Or (Sum(Nvl(A.ʵ�ս��, 0)) = 0 And Sum(Nvl(A.Ӧ�ս��, 0)) <> 0 and Sum(Nvl(A.���ʽ��,0)) =0 And Mod(Count(*),2)=0) " & _
    "                    Or Sum(Nvl(A.���ʽ��, 0))=0 And Sum(Nvl(A.Ӧ�ս��,0))<>0 And Mod(Count(*),2)=0" & _
    "   Group by A.NO,A.���,Mod(A.��¼����,10),Nvl(A.ִ��״̬,0), '��'||NVL(A.��ҳID,0)||'��' ,A.��ҳID," & _
    "               A.��������ID,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.�Ǽ�ʱ��,Nvl(A.Ӥ����,0),A.�շ�ϸĿID,A.������ĿID,A.�վݷ�Ŀ,A.��������,A.ҽ�����" & _
    ""
    
    strTable = strTable & " Union ALL " & strSQL
    strSQL = _
        "Select A.��־,A.סԺ,Nvl(B.����,'δ֪') as ����,A.ʱ��,A.NO as ���ݺ� ,Nvl(E.����,C.����) as ��Ŀ,A.�վݷ�Ŀ as ��Ŀ, A.Ӥ����, " & _
        "       A.ID,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,A.��ҳID,A.��������ID,A.�Ǽ�ʱ��," & _
        "       A.����, A.��׼���� as �۸�,nvl(A.Ӧ�ս��,0) as Ӧ�ս��,nvl(A.ʵ�ս��,0) as ʵ�ս��, " & _
        "       Nvl(A.δ����,0) δ����,Nvl(A.δ����,0) ���ʽ��, A.ͳ����," & _
        "       Nvl(A.��������,C.��������) as ����, A.�շ����,M.���� as �շ������," & _
        "       A.�ѱ�, A.ִ�в���id, A.������, A.���մ���id,A.�շ�ϸĿID,C.���㵥λ,Decode(a.��¼״̬, 2, 2, 3, 2, 1) As ����, Max(G.�������) As ��� " & _
        " From (  " & strTable & ") A,���ű� B,�շ���ĿĿ¼ C,������Ŀ D,�շ���Ŀ���� E,�շ���Ŀ��� M,�������ҽ�� F,������ϼ�¼ G " & _
        " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID And a.ҽ����� = f.ҽ��id(+) And f.���id = g.Id(+) And A.������ĿID=D.ID  And A.�շ����=M.����(+) " & strDiagCondition & _
        "        And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                IIf(strClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
        " Group By A.��־,A.סԺ,Nvl(B.����,'δ֪'),A.ʱ��,A.NO ,Nvl(E.����,C.����),A.�վݷ�Ŀ, A.Ӥ����, " & _
        "       A.ID,A.���,A.��¼����,A.��¼״̬,A.ִ��״̬,A.��ҳID,A.��������ID,A.�Ǽ�ʱ��," & _
        "       A.����, A.��׼����,nvl(A.Ӧ�ս��,0),nvl(A.ʵ�ս��,0), " & _
        "       Nvl(A.δ����,0),Nvl(A.δ����,0), A.ͳ����," & _
        "       Nvl(A.��������,C.��������), A.�շ����,M.����," & _
        "       A.�ѱ�, A.ִ�в���id, A.������, A.���մ���id,A.�շ�ϸĿID,C.���㵥λ,Decode(a.��¼״̬, 2, 2, 3, 2, 1) " & _
        " Order by A.ʱ�� Desc,A.סԺ,A.NO Desc,A.��¼����,A.���"
    
    'Mod(Count(*),2)=1��Ϊ��������ۺ�ʵ�ս��Ϊ��ķ����ڽ��ʺ��Ƿ����ϻ��ٴν���
    On Error GoTo errH
    Set GetZYBalanceData = zlDatabase.OpenSQLRecord(strSQL, "��ȡסԺ���ʼ�¼", lng����ID, "," & strTime & ",", DateBegin, DateEnd, _
                    "," & strDeptIDs & ",", "," & strBaby & ",", "," & strItem & ",", "," & strClass & ",", "," & strChargeType & ",", "," & strChargeTypeNot & ",", strDiag)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 



Public Function GetMzBalance_Insure(ByVal int���� As Integer, lng����ID As Long, _
     Optional dtBeginDate As Date, Optional dtEndDate As Date, Optional blnOnlyYbUpData As Boolean, _
     Optional blnDateMoved As Boolean, Optional blnOnly���� As Boolean, Optional bytKind As Byte, _
     Optional strItem As String, Optional strDeptIDs As String, Optional strClass As String, _
     Optional strChargeType As String = "", Optional bln����ʱ�� As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ������δ����ϸĿ��ϸ(���շ�ϸĿ)
    '��Σ�lng����ID-����ID,
    '      dtBeginDate,dtEndDate�� ���ʷ����ڼ�,���Ǽ�ʱ�����ʱ��,ȱʡֵΪCDate("0:00:00")
    '      blnOnlyYbUpData���Ƿ�ֻ�������ϴ�����
    '      blnDateMoved�����˵Ǽ�ʱ���Ƿ���ת������֮ǰ
    '      blnOnly�����������ʷ���
    '      bytKind��0-����ͨ����,1-��������,2-��ͨ���ú�������
    '      strItem:�վݷ�Ŀ��,'��ҩ��','��ҩ��',...
    '      strDeptIds����������ID��,"1,2,3",�ձ�ʾ����
    '      strClass�����з���(��δ����),"'����1','����2',..."
    '      bln����ʱ��-�Ƿ񰴷���ʱ��ͳ��,true-����ʱ�䣬false-�Ǽ�ʱ��
    '���أ��ɹ�=��¼��,ʧ��=Nothing
    '���ƣ����˺�
    '����:2015-01-06 17:28:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCond As String, blnRelation As Boolean
    Dim strTable As String
    
    On Error GoTo errH
     
    
    blnRelation = gclsInsure.GetCapability(support��������ҽ����Ŀ, lng����ID, int����)
     
    strCond = " And A.����ID=[1]"
    If dtBeginDate <> CDate("0:00:00") Then
        strCond = strCond & " And " & IIf(Not bln����ʱ��, "A.�Ǽ�ʱ��", "A.����ʱ��") & " Between [3] And [4]"
        dtBeginDate = CDate(Format(dtBeginDate, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(dtEndDate, "yyyy-MM-dd 23:59:59"))
    End If
 
    '���˺�:2010-03-06 11:23:52: Or A.����ID is Not NULL ����������ѽ��ʵ���ϸ,���ҵķ�������,�Ǵ��,����û��˵ҽ��������,����ݲ�����!
    If blnOnlyYbUpData Then strCond = strCond & " And (A.�Ƿ��ϴ�=1 And A.����ID is Null Or A.����ID is Not NULL)"
    
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.�վݷ�Ŀ||''',')>0")
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([8],','||A.��������ID||',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([10],','''||A.�շ����||''',')>0")   '34260
    
     '0-����ͨ����,1-��������,2-��ͨ���ú�������
    If bytKind = 1 Then '��������
        strCond = strCond & " And A.�����־=4"
    Else
        strCond = strCond & " And A.�����־<>2"
        If bytKind = 0 Then strCond = strCond & " And A.�����־<>4"
    End If
    '����Ҫ�󣺼�¼����,��¼״̬,NO����š��շ�����շ�ϸĿID,�շ����ơ����㵥λ���������š���񡢲��ء��������۸񡢽�ҽ��,
    '          ����ʱ��,�Ǽ�ʱ��,Ӥ����,ҽ����Ŀ���롢���մ���ID��������Ŀ���Ƿ��ϴ�,�Ƿ���
    'ע�⣺���ڽ���ֻ������б�����Ŀ�����,�������뱣��֧����Ŀ����ʱ����(+)
    '   ����Ϊ��ָ���ѱ����������Ĵ��۳����¼,���൥����ϸ���ϴ�
    
    '��ʱ���ģ��������Ϻ��SQL����,������ʱ��Ϊ"���/����"
    If blnOnly���� Then
        '�������
        strTable = IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A")
        strTable = "" & _
        "       Select NO, ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, Max(Decode(A.��¼����,2,A.ժҪ,Null)) As ժҪ, �Ƿ���, ��������id, ִ�в���id," & vbNewLine & _
        "              Avg(Nvl(����, 0) * ����) As ����, Sum(��׼����*decode(sign(A.��¼����-10),1,0,1)) As ����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ʵ�ս��," & vbNewLine & _
        "              Sum(ͳ����) As ͳ����,��������" & vbNewLine & _
        "       From " & strTable & vbNewLine & _
        "       Where ��¼״̬ <> 0 And ���ʷ��� = 1 And A.���� <> 0 And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼", , , , "C"), "������ü�¼ C") & ", ���˽��ʼ�¼ D Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id And Nvl(d.����״̬, 0) = 1) " & strCond & vbNewLine & _
        "       Group By NO, Mod(��¼����, 10),decode(��¼״̬,2,2,1), Nvl(�۸񸸺�, ���), ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, �Ƿ���," & vbNewLine & _
        "                ��������id, ִ�в���id,��������" & vbNewLine & _
        "       Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0 "
        
        strTable = strTable & " Union ALL " & _
        "       Select NO, ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, Max(Decode(A.��¼����,5,A.ժҪ,Null)) As ժҪ, �Ƿ���, ��������id, ִ�в���id," & vbNewLine & _
        "              Avg(Nvl(����, 0) * ����) As ����, Sum(��׼����*decode(sign(A.��¼����-10),1,0,1)) As ����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ʵ�ս��," & vbNewLine & _
        "              Sum(ͳ����) As ͳ����,��������" & vbNewLine & _
        "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & vbNewLine & _
        "       Where ��¼״̬ <> 0 And ���ʷ��� = 1 And A.���� <> 0 And  Mod(A.��¼����,10) = 5 And A.��ҳID Is Null And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼", , , , "C"), "סԺ���ü�¼ C") & ", ���˽��ʼ�¼ D Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id And Nvl(d.����״̬, 0) = 1) " & strCond & vbNewLine & _
        "       Group By NO, Mod(��¼����, 10), decode(��¼״̬,2,2,1), Nvl(�۸񸸺�, ���), ����id, �շ����, �վݷ�Ŀ, ���㵥λ, ������, �շ�ϸĿid, �Ƿ���," & vbNewLine & _
        "                ��������id, ִ�в���id,��������" & vbNewLine & _
        "       Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0 "
            
        strSQL = "" & _
        " Select Sysdate As ����ʱ��, A.����id, A.�շ����, A.�վݷ�Ŀ, A.���㵥λ, A.�շ�ϸĿid," & vbNewLine & _
        "       B.����id ����֧������id, B.�Ƿ�ҽ�� �Ƿ�ҽ��, B.��Ŀ���� ���ձ���, Sum(A.����) As ����, Avg(A.����) As ����," & vbNewLine & _
        "       Sum(A.ʵ�ս��) As ʵ�ս��, Sum(A.ͳ����) As ͳ����, Max(A.ժҪ) ժҪ, Max(A.�Ƿ���) �Ƿ���," & vbNewLine & _
        "       Max(A.��������id) ��������id, Max(A.ִ�в���id) ִ�в���id, Max(A.������) ������,Max(A.��������) ��������" & vbNewLine & _
        " From ( " & strTable & ") A, ����֧����Ŀ B, �շ���ĿĿ¼ C " & vbNewLine & _
        " Where A.�շ�ϸĿid = C.ID And A.�շ�ϸĿid = B.�շ�ϸĿid" & IIf(blnRelation, "(+)", "") & " And B.����" & IIf(blnRelation, "(+)", "") & " = [5] " & vbNewLine & _
                    IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
        " Group By A.�շ�ϸĿid, A.����id, A.�շ����, A.�վݷ�Ŀ, A.���㵥λ, B.����id, B.�Ƿ�ҽ��, B.��Ŀ����" & vbNewLine & _
        " Having Sum(A.ʵ�ս��) <> 0"
    Else
        'סԺ����
       strTable = IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A")
        strTable = "" & _
        "       Select Mod(A.��¼����, 10) As ��¼����, decode(A.��¼״̬,2,2,1) as ��־,max(A.��¼״̬) as ��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���) As ���, A.�����־, A.����id," & vbNewLine & _
        "               -1*NULL  as ��ҳid, Nvl(A.Ӥ����, 0) As Ӥ����, A.������ As ҽ��, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ," & vbNewLine & _
        "              A.���ձ���, Nvl(A.���մ���id, 0) As ���մ���id, Avg(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
        "              Sum(A.��׼����*decode(sign(A.��¼����-10),1,0,1)) As ��׼����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ���, A.����ʱ��," & vbNewLine & _
        "              A.�Ǽ�ʱ��, Min(Nvl(A.�Ƿ��ϴ�, 0)) As �Ƿ��ϴ�, Nvl(A.�Ƿ���, 0) As �Ƿ���,Nvl(A.������Ŀ��, 0) As ������Ŀ��, Max(Decode(A.��¼����,2,A.ժҪ,Null)) As ժҪ,A.��������" & vbNewLine & _
        "       From " & strTable & " , ������Ŀ B" & vbNewLine & _
        "       Where A.��¼״̬ <> 0 And A.���ʷ��� = 1 And A.������Ŀid = B.ID And A.���� <> 0 And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("������ü�¼", , , , "C"), "������ü�¼ C") & ", ���˽��ʼ�¼ D Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id And Nvl(d.����״̬, 0) = 1) " & strCond & vbNewLine & _
        "       Group By Mod(A.��¼����, 10), decode(A.��¼״̬,2,2,1), A.NO, Nvl(A.�۸񸸺�, ���), A.�����־, A.����id," & vbNewLine & _
        "                Nvl(A.Ӥ����, 0), A.������, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ, A.���ձ���," & vbNewLine & _
        "                Nvl(A.���մ���id, 0), A.����ʱ��, A.�Ǽ�ʱ��, Nvl(A.�Ƿ���, 0), Nvl(A.������Ŀ��, 0),A.��������" & vbNewLine & _
        "       Having Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) <> 0��"
            

        strTable = strTable & " Union ALL" & _
        "       Select Mod(A.��¼����, 10) As ��¼����,decode(A.��¼״̬,2,2,1) as ��־,max(A.��¼״̬) as ��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���) As ���, A.�����־, A.����id," & vbNewLine & _
        "               -1*NULL as ��ҳid, Nvl(A.Ӥ����, 0) As Ӥ����, A.������ As ҽ��, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ," & vbNewLine & _
        "              A.���ձ���, Nvl(A.���մ���id, 0) As ���մ���id, Avg(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
        "              Sum(A.��׼����*decode(sign(A.��¼����-10),1,0,1)) As ��׼����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ���, A.����ʱ��," & vbNewLine & _
        "              A.�Ǽ�ʱ��, Min(Nvl(A.�Ƿ��ϴ�, 0)) As �Ƿ��ϴ�, Nvl(A.�Ƿ���, 0) As �Ƿ���,Nvl(A.������Ŀ��, 0) As ������Ŀ��, Max(Decode(A.��¼����,5,A.ժҪ,Null)) As ժҪ,A.��������" & vbNewLine & _
        "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " , ������Ŀ B" & vbNewLine & _
        "       Where A.��¼״̬ <> 0 And A.���ʷ��� = 1 And  Mod(A.��¼����,10) = 5 And A.��ҳID Is Null  And A.������Ŀid = B.ID And A.���� <> 0 And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼", , , , "C"), "סԺ���ü�¼ C") & ", ���˽��ʼ�¼ D Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id And Nvl(d.����״̬, 0) = 1) " & strCond & vbNewLine & _
        "       Group By Mod(A.��¼����, 10), decode(A.��¼״̬,2,2,1), A.NO, Nvl(A.�۸񸸺�, ���), A.�����־, A.����id  ," & vbNewLine & _
        "                Nvl(A.Ӥ����, 0), A.������, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ, A.���ձ���," & vbNewLine & _
        "                Nvl(A.���մ���id, 0), A.����ʱ��, A.�Ǽ�ʱ��, Nvl(A.�Ƿ���, 0), Nvl(A.������Ŀ��, 0),A.��������" & vbNewLine & _
        "       Having Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) <> 0��"
        strSQL = "" & _
        "   Select A.��¼����, A.��¼״̬, A.NO, A.���, A.�����־, A.����id, A.��ҳid, A.Ӥ����, C.��Ŀ���� As ҽ����Ŀ����," & vbNewLine & _
        "       A.���ձ���, A.���մ���id, A.�շ����, A.�շ�ϸĿid, Nvl(E.����, B.����) As �շ�����, A.���㵥λ," & vbNewLine & _
        "       X.���� As ��������, B.���, B.����, A.����, A.��׼���� As �۸�, A.���," & vbNewLine & _
        "       A.ҽ��, A.����ʱ��, A.�Ǽ�ʱ��, A.�Ƿ��ϴ�, A.�Ƿ���, A.������Ŀ��, A.ժҪ,A.��������" & vbNewLine & _
        "   From ( " & strTable & ") A, �շ���ĿĿ¼ B, ����֧����Ŀ C, �շ���Ŀ���� E,���ű� X" & vbNewLine & _
        "   Where A.�շ�ϸĿid = B.ID And B.ID = C.�շ�ϸĿid" & IIf(blnRelation, "(+)", "") & " And C.����" & IIf(blnRelation, "(+)", "") & " = [5] And A.��������id = X.ID " & vbNewLine & _
                IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.��������,Nvl(B.��������,'��'))||''',')>0") & _
        "      And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = " & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1)
    End If
    Set GetMzBalance_Insure = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ������ѷ�����", lng����ID, "", dtBeginDate, dtEndDate, int����, 0, "," & strItem & ",", "," & strDeptIDs & ",", "," & strClass & ",", "," & strChargeType & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetZYBalance_Insure(ByVal int���� As Integer, lng����ID As Long, Optional strTime As String, _
     Optional dtBeginDate As Date, Optional dtEndDate As Date, Optional blnOnlyYbUpData As Boolean, _
     Optional blnDateMoved As Boolean, Optional strBaby As String, _
     Optional strItem As String, Optional strDeptIDs As String, Optional strClass As String, _
     Optional strChargeType As String = "", Optional bln����ʱ�� As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ������δ����ϸĿ��ϸ(���շ�ϸĿ)
    '��Σ�lng����ID-����ID,
    '      strTime�� ҽ������ֻ������סԺ�����ͷ����ڼ� [strTime=סԺ������,"0,1,2,3",0��ʾ����]
    '      dtBeginDate,dtEndDate�� ���ʷ����ڼ�,���Ǽ�ʱ�����ʱ��,ȱʡֵΪCDate("0:00:00")
    '      blnOnlyYbUpData���Ƿ�ֻ�������ϴ�����
    '      blnDateMoved�����˵Ǽ�ʱ���Ƿ���ת������֮ǰ
    '      strBaby��0-���з���,1-���˷���,2�Լ���-��bytBaby-1��Ӥ������]
    '      strItem:�վݷ�Ŀ��,'��ҩ��','��ҩ��',...
    '      strDeptIds����������ID��,"1,2,3",�ձ�ʾ����
    '      strClass�����з���(��δ����),"'����1','����2',..."
    '      bln����ʱ��-�Ƿ񰴷���ʱ��ͳ��,true-����ʱ�䣬false-�Ǽ�ʱ��
    '���Σ�
    '���أ��ɹ�=��¼��,ʧ��=Nothing
    '���ƣ����˺�
    '����:2015-01-06 17:31:52
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCond As String, blnRelation As Boolean
    Dim strTable As String, strWherePage As String 'סԺ��������
    On Error GoTo errH
    
    blnRelation = gclsInsure.GetCapability(support��������ҽ����Ŀ, lng����ID, int����)
    
    strCond = " And A.����ID=[1]"
    strWherePage = IIf(strTime = "", "", " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0")
    
    If dtBeginDate <> CDate("0:00:00") Then
        strCond = strCond & " And " & IIf(Not bln����ʱ��, "A.�Ǽ�ʱ��", "A.����ʱ��") & " Between [3] And [4]"
        dtBeginDate = CDate(Format(dtBeginDate, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(dtEndDate, "yyyy-MM-dd 23:59:59"))
    End If
 
    '���˺�:2010-03-06 11:23:52: Or A.����ID is Not NULL ����������ѽ��ʵ���ϸ,���ҵķ�������,�Ǵ��,����û��˵ҽ��������,����ݲ�����!
    If blnOnlyYbUpData Then strCond = strCond & " And (A.�Ƿ��ϴ�=1 And A.����ID is Null Or A.����ID is Not NULL)"

    strCond = strCond & IIf(strBaby = "", "", " And Instr([6],','|| Nvl(A.Ӥ����,0) ||',')>0")
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.�վݷ�Ŀ||''',')>0")
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([8],','||A.��������ID||',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([10],','''||A.�շ����||''',')>0")   '34260
    strCond = strCond & " And A.�����־ In (2,3) "
    '����Ҫ�󣺼�¼����,��¼״̬,NO����š��շ�����շ�ϸĿID,�շ����ơ����㵥λ���������š���񡢲��ء��������۸񡢽�ҽ��,
    '          ����ʱ��,�Ǽ�ʱ��,Ӥ����,ҽ����Ŀ���롢���մ���ID��������Ŀ���Ƿ��ϴ�,�Ƿ���
    'ע�⣺���ڽ���ֻ������б�����Ŀ�����,�������뱣��֧����Ŀ����ʱ����(+)
    '   ����Ϊ��ָ���ѱ����������Ĵ��۳����¼,���൥����ϸ���ϴ�
    
    '��ʱ���ģ��������Ϻ��SQL����,������ʱ��Ϊ"���/����"
 
    strTable = IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A")
    strTable = "" & _
    "       Select Mod(A.��¼����, 10) As ��¼����, decode(A.��¼״̬,2,2,1) as ��־, max( A.��¼״̬) as ��¼״̬, A.NO, Nvl(A.�۸񸸺�, ���) As ���, A.�����־, A.����id," & vbNewLine & _
    "               A.��ҳid as ��ҳid, Nvl(A.Ӥ����, 0) As Ӥ����, A.������ As ҽ��, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ," & vbNewLine & _
    "              A.���ձ���, Nvl(A.���մ���id, 0) As ���մ���id, Avg(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
    "              Sum(A.��׼����*decode(sign(A.��¼����-10),1,0,1)) As ��׼����, Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) As ���, A.����ʱ��," & vbNewLine & _
    "              A.�Ǽ�ʱ��, Min(Nvl(A.�Ƿ��ϴ�, 0)) As �Ƿ��ϴ�, Nvl(A.�Ƿ���, 0) As �Ƿ���,Nvl(A.������Ŀ��, 0) As ������Ŀ��, Max(Decode(A.��¼����,2,A.ժҪ,Null)) As ժҪ,A.��������" & vbNewLine & _
    "       From " & strTable & " , ������Ŀ B" & vbNewLine & _
    "       Where A.��¼״̬ <> 0 And A.���ʷ��� = 1  And A.��ҳID Is Not Null  And A.������Ŀid = B.ID And A.���� <> 0" & vbNewLine & _
    "             And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼", , , , "C"), "סԺ���ü�¼ C") & ", ���˽��ʼ�¼ D" & vbNewLine & _
    "                             Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id" & vbNewLine & _
    "                                   And Nvl(d.����״̬, 0) = 1) " & strCond & strWherePage & vbNewLine & _
    "       Group By Mod(A.��¼����, 10),decode(A.��¼״̬,2,2,1), A.NO, Nvl(A.�۸񸸺�, ���), A.�����־, A.����id ,A.��ҳid ," & vbNewLine & _
    "                Nvl(A.Ӥ����, 0), A.������, A.��������id, A.�շ����, A.�շ�ϸĿid, A.���㵥λ, A.���ձ���," & vbNewLine & _
    "                Nvl(A.���մ���id, 0), A.����ʱ��, A.�Ǽ�ʱ��, Nvl(A.�Ƿ���, 0), Nvl(A.������Ŀ��, 0),A.��������" & vbNewLine & _
    "       Having Sum(Nvl(A.ʵ�ս��, 0)) - Sum(Nvl(A.���ʽ��, 0)) <> 0��"
    
    strSQL = "" & _
    " Select A.��¼����, A.��¼״̬, A.NO, A.���, A.�����־, A.����id, A.��ҳid, A.Ӥ����, C.��Ŀ���� As ҽ����Ŀ����," & vbNewLine & _
    "       A.���ձ���, A.���մ���id, A.�շ����, A.�շ�ϸĿid, Nvl(E.����, B.����) As �շ�����, A.���㵥λ," & vbNewLine & _
    "       X.���� As ��������, B.���, B.����, A.����, A.��׼���� As �۸�, A.���," & vbNewLine & _
    "       A.ҽ��, A.����ʱ��, A.�Ǽ�ʱ��, A.�Ƿ��ϴ�, A.�Ƿ���, A.������Ŀ��, A.ժҪ,A.��������" & vbNewLine & _
    " From ( " & strTable & ") A, �շ���ĿĿ¼ B, ����֧����Ŀ C, �շ���Ŀ���� E,���ű� X" & vbNewLine & _
    " Where A.�շ�ϸĿid = B.ID And B.ID = C.�շ�ϸĿid" & IIf(blnRelation, "(+)", "") & " And C.����" & IIf(blnRelation, "(+)", "") & " = [5] And A.��������id = X.ID " & vbNewLine & _
        IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.��������,Nvl(B.��������,'��'))||''',')>0") & _
    "      And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = " & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1)
    Set GetZYBalance_Insure = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��סԺ������Ϣ��", lng����ID, "," & strTime & ",", dtBeginDate, dtEndDate, int����, "," & strBaby & ",", "," & strItem & ",", "," & strDeptIDs & ",", "," & strClass & ",", "," & strChargeType & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetFromIDToBalanceData(ByVal lng����ID As Long, ByVal blnNOMoved As Boolean, _
    ByRef rsOutBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID����ȡ��������
    '���:lng����ID-����ID
    '     blnNoMoved-�Ƿ��Ѿ�ת�Ƶ��󱸱���
    '����:rsOutBalance-��������
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 15:32:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String
    On Error GoTo errHandle
    
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�;6-����
     strSQL = "" & _
       "   Select  A.ID, " & _
       "        Case when Mod(A.��¼����,10)=1 then 1  " & _
       "             when nvl(M.����,0)=3 or nvl(M.����,0)=4  then 2 " & _
       "             when nvl(A.�����ID,0)<>0  then  3 " & _
       "             when J.���㷽ʽ is not null   then  4 " & _
       "             when nvl(M.����,0)=9 then 6 " & _
       "             else 0 end as ����, " & _
       "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��,A.ժҪ, " & _
       "        A.�����ID,A.���㿨���, " & _
       "        A.�������,A.����,A.������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
       "        nvl(C.�Ƿ�����,0) as �Ƿ�����,nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
       "        Decode(C.��������,NULL,0,1) as  �Ƿ�����," & _
       "        C.���� as ���������,A.����˵��,A.�������,A.У�Ա�־, " & _
       "        decode(nvl(M.����,0),3,1,4,1,0) as ҽ��,0 as ���ѿ�id,nvl(M.����,0) as ��������" & _
       "   From  ����Ԥ����¼ A ,ҽ�ƿ���� C,һ��ͨĿ¼ J,���㷽ʽ M" & _
       "   Where A.����ID= [1] And A.���㷽ʽ=M.����(+) " & _
       "         And A.���㷽ʽ=J.���㷽ʽ(+) And A.�����ID=C.ID(+) " & _
       "         And nvl(A.���㿨���,0)=0"
    
    strSQL = strSQL & " Union ALL " & _
        " Select a.Id," & vbNewLine & _
        "       Case" & vbNewLine & _
        "         When Mod(a.��¼����, 10) = 1 Then" & vbNewLine & _
        "          1" & vbNewLine & _
        "         When Nvl(m.����, 0) = 3 Or Nvl(m.����, 0) = 4 Then" & vbNewLine & _
        "          2" & vbNewLine & _
        "         When Nvl(a.�����id, 0) <> 0 Then" & vbNewLine & _
        "          3" & vbNewLine & _
        "         When j.���㷽ʽ Is Not Null Then" & vbNewLine & _
        "          4" & vbNewLine & _
        "         When Nvl(m.����, 0) = 9 Then" & vbNewLine & _
        "          6" & vbNewLine & _
        "         Else" & vbNewLine & _
        "          0" & vbNewLine & _
        "       End As ����, Mod(a.��¼����, 10) As ��¼����, a.���㷽ʽ, a.��Ԥ��, a.ժҪ, a.�����id, a.���㿨���, a.�������, a.����, a.������ˮ��," & vbNewLine & _
        "       Nvl(c.���ƿ�, 0) As ���ƿ�, Nvl(c.�Ƿ�����, 0) As �Ƿ�����, Nvl(c.�Ƿ�ȫ��, 0) As �Ƿ�ȫ��, Decode(c.�Ƿ�����, Null, 0, 1) As �Ƿ�����," & vbNewLine & _
        "       c.���� As ���������, a.����˵��, a.�������, a.У�Ա�־, Decode(Nvl(m.����, 0), 3, 1, 4, 1, 0) As ҽ��, 0 As ���ѿ�id," & vbNewLine & _
        "       Nvl(m.����, 0) As ��������" & vbNewLine & _
        "From ����Ԥ����¼ A, ���ѿ����Ŀ¼ C, һ��ͨĿ¼ J, ���㷽ʽ M" & vbNewLine & _
        "Where a.����id = [1] And a.���㷽ʽ = m.����(+) And a.���㷽ʽ = j.���㷽ʽ(+) And a.���㿨��� = c.��� And Nvl(a.�����id, 0) = 0 And Mod(a.��¼����,10) =1"

          
    strSQL = strSQL & " Union ALL " & _
       "   Select A.ID,5 as  ����,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,-1*nvl(b.Ӧ�ս��,0) as ��Ԥ��,A.ժҪ,A.�����ID,A.���㿨���," & _
       "        A.�������,B.����,B.������ˮ��,nvl( M.���ƿ�,0) as ���ƿ�, " & _
       "        nvl( M.�Ƿ�����,0) as �Ƿ�����,nvl(M.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
       "        nvl(M.�Ƿ�����,0) as  �Ƿ�����," & _
       "        M.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,0 as ҽ��,B.���ѿ�id,M1.���� as ��������" & _
       "   From ����Ԥ����¼ A ,���˿������¼ B, " & _
       "        ���ѿ����Ŀ¼ M,���㷽ʽ M1 " & _
       "   Where  a.Id = b.����Id " & _
       "        And a.���㿨��� = m.��� And A.���㷽ʽ=M1.����(+) " & _
       "        And A.����ID = [1] and Mod(A.��¼����,10)<>1 "
       
      strSQL = "" & _
      "   Select A.����,a.��¼����,a.���㷽ʽ,a.ժҪ,a.�����ID,a.���������,a.���ƿ�,a.���㿨���,a.�������,a.����,a.������ˮ��,a. ����˵��,a.�������,a.У�Ա�־,a.ҽ��,a.���ѿ�id," & _
      "         max(A.�Ƿ�����) as �Ƿ�����,max(A.�Ƿ�ȫ��) as �Ƿ�ȫ��,max(a.�Ƿ�����) as �Ƿ�����, nvl(sum(a.��Ԥ��),0) as ��Ԥ��,Max(A.��������) as ����" & _
      "   From (" & strSQL & ") A " & _
      "   Group by A.����,a.��¼����,a.���㷽ʽ,a.ժҪ,a.�����ID,a.���������,a.���ƿ�,a.���㿨���,a.�������,a.����,a.������ˮ��,a. ����˵��,a.�������,a. У�Ա�־,a.ҽ��,a.���ѿ�id" & _
      "   Order by ����"
    
    If blnNOMoved Then
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        strSQL = Replace(strSQL, "���˿������¼", "H���˿������¼")
    End If
    
    Set rsOutBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lng����ID)
    zlGetFromIDToBalanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function zlGetFormerBalanceID(strNO As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԭ���ʵ��ݵ�ID
    '����:��ȡ�ɹ�,����ԭ����ID,���򷵻�0
    '����:���˺�
    '����:2015-01-26 09:51:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ID From ���˽��ʼ�¼ Where ��¼״̬ IN(1,3) And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԭʼ����ID", strNO)
    If Not rsTmp.EOF Then zlGetFormerBalanceID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlFromIDGetChargeBalance(ByVal bytType As Byte, _
    ByVal strValue As String, Optional blnHistory As Boolean, _
    Optional ByRef blnDel As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ�շѽ�����Ϣ
    '���:bytType-��������:0-���ݽ���ID����;1-���ݵ��ݺ�����ȡ���㷽ʽ
    '     strValue-Ҫ���ҵ�ֵ(Ϊ0ʱ,����ID, 2ʱΪ���ʵ��ݺ�)
    '     blnDel-���Ͻ���:true-�����Ͻ���;false-�����Ͻ���
    '����:�շѽ���������Ϣ��
    '       �ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '       ����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    '����:���˺�
    '����:2014-06-24 16:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String, strWhere As String
    Dim strTable1 As String
    On Error GoTo errHandle
    
    strTable = IIf(blnHistory, "H", "") & "����Ԥ����¼"
    Select Case bytType
    Case 0  '0-���ݽ���ID����
        strWhere = " And  A.����ID= [1]"
    Case 1 '���ݵ��ݺ�����ȡ��������
        strTable1 = "" & _
        "   Select distinct ID  " & _
        "   From ���˽��ʼ�¼ M " & _
        "   Where m.no=[2]  And ��¼״̬ in (1,3) And nvl(M.����״̬,0)<>1"
        strTable1 = ",(" & strTable1 & ") Q1"
        
        If blnHistory Then strTable1 = Replace(strTable1, "���˽��ʼ�¼", "H���˽��ʼ�¼")
        strWhere = " And A.����ID=Q1.ID"
    Case Else
        Exit Function
    End Select
    
    If blnDel Then
        '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        strSQL = "" & _
        "   Select  A.ID,decode(A.��¼״̬,2,A.����ID,NULL) as ����ID," & _
        "        Case when Mod(A.��¼����,10)=1 then 1  " & _
        "             when B.���� is not null then  2 " & _
        "             when nvl(A.�����ID,0)<>0  then  3 " & _
        "             when J.���㷽ʽ is not null   then  4 " & _
        "             else 0 end as ����, " & _
        "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��," & _
        "        decode(A.��¼״̬,2,A.ժҪ,NULL) as ժҪ,decode(A.��¼״̬,2,1,0) as �˷�," & _
        "        A.�����ID,A.���㿨���, " & _
        "        decode(A.��¼״̬,2,A.�������,NULL) as �������,decode(A.��¼״̬,2,A.����,NULL) as ����, " & _
        "        decode(A.��¼״̬,2,A.������ˮ��,NULL) as ������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
        "        nvl(C.�Ƿ�����,0) as �Ƿ�����,nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
        "        Decode(C.��������,NULL,0,1) as  �Ƿ�����," & _
        "        C.���� as ���������,decode(A.��¼״̬,2,A.����˵��,NULL) as ����˵��,A.�������,decode(A.��¼״̬,2,A.У�Ա�־,0) as У�Ա�־, " & _
        "        decode(B.����,Null,0,1) as ҽ��,0 as ���ѿ�id,nvl(q.����,1) as ��������" & _
        "   From " & strTable & " A ,ҽ�ƿ���� C,һ��ͨĿ¼ J,���㷽ʽ q," & _
        "        (Select ���� From ���㷽ʽ where ���� in (3,4)) B " & strTable1 & _
        "   Where A.���㷽ʽ=J.���㷽ʽ(+) And A.�����ID=C.ID(+) " & _
        "         And A.���㷽ʽ=B.����(+) and A.���㷽ʽ=q.����(+) " & _
        "         And nvl(A.���㿨���,0)=0 " & strWhere
        strSQL = strSQL & " Union ALL " & _
        "   Select A.ID,decode(A.��¼״̬,2,A.����ID,NULL) as ����ID,5 as  ����,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,-1*nvl(b.Ӧ�ս��,0) as ��Ԥ��,A.ժҪ," & _
        "        decode(A.��¼״̬,2,1,0) as �˷�,A.�����ID,A.���㿨���," & _
        "        decode(A.��¼״̬,2,A.�������,NULL) as �������,decode(A.��¼״̬,2,B.����,NULL) as ����, " & _
        "        decode(A.��¼״̬,2,B.������ˮ��,NULL) as ������ˮ��,nvl(M.���ƿ�,0) as ���ƿ�, " & _
        "        nvl( M.�Ƿ�����,0) as �Ƿ�����,nvl(M.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
        "        nvl(M.�Ƿ�����,0) as  �Ƿ�����," & _
        "        M.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,0 as ҽ��,B.���ѿ�id,nvl(q.����,1) as ��������" & _
        "   From  " & strTable & " A ,���˿������¼ B, " & _
        "        ���ѿ����Ŀ¼ M ,���㷽ʽ q " & strTable1 & _
        "   Where  a.Id = b.����Id  And a.���㿨��� = m.���  " & _
        "         and Mod(A.��¼����,10)<>1 and A.���㷽ʽ=q.����(+) " & strWhere
        
        strSQL = "" & _
        "   Select /*+ Rule */ max(����id) as ����id,����,max(�˷�) as �˷�,��¼����,���㷽ʽ,Max(ժҪ) as ժҪ,�����ID,���������,max(���ƿ�) as ���ƿ�,���㿨���, " & _
        "         max(�������) as �������,max(����) as ����,max(������ˮ��) as ������ˮ��, max(����˵��) as ����˵��, " & _
        "         �������,max(У�Ա�־) as У�Ա�־,ҽ��,���ѿ�id,��������," & _
        "         max(�Ƿ�����) as �Ƿ�����,max(�Ƿ�ȫ��) as �Ƿ�ȫ��,max(�Ƿ�����) as �Ƿ����� , nvl(sum(��Ԥ��),0) as ��Ԥ��" & _
        "   From (" & strSQL & ") " & _
        "   Group by ����, ��¼����,���㷽ʽ,�����ID,���������,���㿨���,�������,ҽ��,���ѿ�id,�������� having  sum(��Ԥ��) <>0"
        Set zlFromIDGetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շѽ��㷽ʽ", Val(strValue), strValue)
        Exit Function
    End If
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    strSQL = "" & _
    "   Select /*+ Rule */ A.ID,A.����ID," & _
    "        Case when Mod(A.��¼����,10)=1 then 1  " & _
    "             when B.���� is not null then  2 " & _
    "             when nvl(A.�����ID,0)<>0  then  3 " & _
    "             when J.���㷽ʽ is not null   then  4 " & _
    "             else 0 end as ����, " & _
    "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��," & _
    "        A.ժҪ,decode(A.��¼״̬,2,1,0) as �˷�," & _
    "        A.�����ID,A.���㿨���, " & _
    "        A.�������,A.����,A.������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
    "        nvl(C.�Ƿ�����,0) as �Ƿ�����,nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
    "        Decode(C.��������,NULL,0,1) as  �Ƿ�����," & _
    "        C.���� as ���������,A.����˵��,A.�������,A.У�Ա�־, " & _
    "        decode(B.����,Null,0,1) as ҽ��,0 as ���ѿ�id,nvl(q.����,1) as ��������" & _
    "   From " & strTable & " A ,ҽ�ƿ���� C,һ��ͨĿ¼ J,���㷽ʽ q," & _
    "        (Select ���� From ���㷽ʽ where ���� in (3,4)) B " & strTable1 & _
    "   Where A.���㷽ʽ=J.���㷽ʽ(+) And A.�����ID=C.ID(+) " & _
    "         And A.���㷽ʽ=B.����(+) and A.���㷽ʽ=q.����(+) " & _
    "         And nvl(A.���㿨���,0)=0 " & strWhere
       
    strSQL = strSQL & " Union ALL " & _
    "   Select /*+ Rule */ A.ID,A.����ID,5 as  ����,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,-1*nvl(b.Ӧ�ս��,0) as ��Ԥ��,A.ժҪ," & _
    "        decode(A.��¼״̬,2,1,0) as �˷�,A.�����ID,A.���㿨���," & _
    "        A.�������,B.����,B.������ˮ��,nvl( M.���ƿ�,0) as ���ƿ�, " & _
    "        nvl( M.�Ƿ�����,0) as �Ƿ�����,nvl(M.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
    "        nvl(M.�Ƿ�����,0) as  �Ƿ�����," & _
    "        M.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,0 as ҽ��,B.���ѿ�id,nvl(q.����,1) as ��������" & _
    "   From  " & strTable & " A ,���˿������¼ B, " & _
    "        ���ѿ����Ŀ¼ M ,���㷽ʽ q " & strTable1 & _
    "   Where  a.Id = b.����Id  And a.���㿨��� = m.���  " & _
    "         and Mod(A.��¼����,10)<>1 and A.���㷽ʽ=q.����(+) " & strWhere
    gstrSQL = "" & _
    "   Select  ����ID,����,�˷�,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id,��������," & _
    "         max(�Ƿ�����) as �Ƿ�����,max(�Ƿ�ȫ��) as �Ƿ�ȫ��,max(�Ƿ�����) as �Ƿ����� , nvl(sum(��Ԥ��),0) as ��Ԥ��" & _
    "   From (" & gstrSQL & ") " & _
    "   Group by ����ID,����,�˷�,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id,��������"
    Set zlFromIDGetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շѽ��㷽ʽ", Val(strValue), strValue)
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetPatiRsByUnit(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    blnȡ���ò��� As Boolean, ByVal blnȡʣ��� As Boolean, _
    Optional ByVal bln��Ԥ��Ժ As Boolean = False, _
    Optional ByVal int��ȡ��Χ As Integer = -1, _
    Optional ByVal bln������Ժ���� As Boolean = False, _
    Optional ByVal strOutBeginDate As String = "", _
    Optional ByVal strOutEndDate As String = "") As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�����Ĳ�����Ϣ����
    '���:lng����ID-����ID
    '     lng����ID-�����˲���ID
    '     blnȡ���ò���-Trueʱ:���صĲ�����Ϣ�����Ƿ����"���ò���"��Ϣ,��ͨ��"zl_PatiWarnScheme"��������ֵ
    '                   Falseʱ:����NULL
    '     blnȡʣ���-trueʱ:���صĲ�����Ϣ�����Ƿ����"ʣ���",��:Ԥ�����-�������+Ԥ������)
    '                 Falseʱ:����NULL
    '     bln��Ԥ��Ժ-�Ƿ����Ԥ��Ժ����
    '     int��ȡ��Χ:-1:���в��˰���Ӥ��
    '                 0-ֻ��������
    '                 1.ֻ����Ӥ��
    '     bln������Ժ����-�Ƿ������Ժ����
    '     strOutBeginDate:��Ժ��ʼʱ��(bln������Ժ���ˣ�trueʱ��Ч),��ʽ:yyyy-mm-dd hh24:mi:ss
    '     strOutEndDate:��Ժ����ʱ��(bln������Ժ���ˣ�trueʱ��Ч),��ʽ:yyyy-mm-dd hh24:mi:ss
    '����:�ɹ�,���ز�����Ϣ������true,���򷵻�Nothing
    '����:���˺�
    '����:2015-07-08 10:59:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intBedLen As Integer, strTable As String
    Dim dtOutBegin As Date, dtOutEnd As Date
    Dim strWithTable As String
    Dim strFields As String '�ֶ�
    Dim strWhere As String  '����
    
    On Error GoTo errH
    
    intBedLen = GetMaxBedLen(lng����ID, False)
    strFields = ""
    
    
    If blnȡʣ��� Then
        strFields = strFields & "," & _
        "  Nvl(E.Ԥ�����,0)-Nvl(E.�������,0)+ " & _
        "  Decode(B.����,Null,0,(Select Nvl(Sum(���),0) From ����ģ����� F Where B.����ID=F.����ID And B.��ҳID=F.��ҳID)) as ʣ���"
    Else
        strFields = strFields & ",NULL as ʣ���"
    End If
    strFields = strFields & "," & IIf(blnȡ���ò���, "zl_PatiWarnScheme(A.����ID,B.��ҳID)", "NULL") & " as ���ò���"
    
    If int��ȡ��Χ = 1 Then 'ֻ����Ӥ��
        strWhere = " And Exists(select 1 from ������������¼ Z Where z.����id=b.����ID And z.��ҳid=b.��ҳID)"
    End If
    
    
    
    strTable = "" & _
    " Select a.����ID" & vbNewLine & _
    " From ������Ϣ A, ������ҳ B,��Ժ���� R" & vbNewLine & _
    " Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����ID=R.����ID " & _
    "       And A.��ǰ����ID=R.����ID And (R.����id = [1] or B.Ӥ������ID = [1]) " & _
            IIf(bln��Ԥ��Ժ, "", " And Nvl(b.״̬,0)<>3") & vbNewLine & _
    " Union" & vbNewLine & _
    " Select 0+[2] as ����ID From Dual"
    
    dtOutBegin = CDate("1991-01-01")
    dtOutEnd = CDate("1991-01-01")
    If bln������Ժ���� Then
         '��Ժ����ʱ�䷶Χ
        If strOutBeginDate = "" Then
            strOutBeginDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD 00:00:00")
            strOutEndDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD 23:59:59")
        End If
        dtOutBegin = CDate(strOutBeginDate)
        dtOutEnd = CDate(strOutEndDate)
        strTable = strTable & _
        " Union" & vbNewLine & _
        " Select a.����id" & vbNewLine & _
        " From ������Ϣ A, ������ҳ B" & vbNewLine & _
        " Where a.����id = b.����id And a.��ҳid = b.��ҳid  " & _
        "       And (b.��ǰ����id + 0 = [1] Or b.Ӥ������id + 0 = [1]) " & _
        "       And B.��Ժ���� Between [3] And [4]"
    End If
        
        
    strWithTable = "" & _
    " With T������Ϣ as ( " & _
    "       Select A.����ID,B.��ҳID,nvl(B.����,A.����) as ����,B.סԺ��,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����," & _
    "           Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������," & _
    "           zl_PatiDayCharge(A.����ID) as ���ն�," & _
    "           E.Ԥ�����,E.�������,B.סԺҽʦ,nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,D.���� as ����ȼ�,C.���� as ����," & _
    "           c.id as ����id,B.��Ժ����,B.��Ժ����,B.��������,nvl(B.�Ա�,A.�Ա�) as �Ա�," & _
    "           nvl(B.����,A.����) as ����,b.��˱�־,B.Ӥ������ID,B.Ӥ������ID,nvl(B.����,0) as ����,B.״̬," & _
    "           nvl(b.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,B.�������� as ��������,B.סԺҽʦ As ������,A.��ǰ����ID As ��������ID,M.���� as ������������" & strFields & _
    "       From ������Ϣ A,������ҳ B,���ű� C,���ű� M,�շ���ĿĿ¼ D,������� E,(" & strTable & ") F" & _
    "       Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.����ȼ�ID=D.ID(+)" & _
    "           And B.��Ժ����ID=C.ID And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
    "           And A.��ǰ����ID=M.id(+) And A.����ID=E.����ID(+) And E.����(+)=1 And E.����(+) = 2 And A.����ID=F.����ID " & strWhere & _
    "       Order by ����)"
    
    strSQL = "" & vbCrLf & _
    " Select A.����id, A.��ҳid, A.����, A.סԺ��, A.����, A.������,A.���ն�,A.Ԥ�����,A.�������, A.ʣ���,  A.���ò���, A.����,b.���� as �������, " & _
    "       A.סԺҽʦ, A.�ѱ�, A.����ȼ�, A.����,A.����id," & _
    "       A.��Ժ����, A.��Ժ����, A.��������, A.�Ա�,A.����, A.��˱�־,A.Ӥ������ID,A.Ӥ������ID,Null as Ӥ������,Null as Ӥ�����," & _
    "       A.״̬,A.ҽ�Ƹ��ʽ,A.��������,A.������,A.��������ID,A.������������" & _
    " From T������Ϣ A,������� B " & _
    " Where A.����=B.���(+)"
    
    If int��ȡ��Χ = 1 Or int��ȡ��Χ = -1 Then '����Ӥ��
        strSQL = IIf(int��ȡ��Χ = 1, "", strSQL & " Union  ALL ") & vbCrLf & _
        " Select a.����id,a.��ҳid,a.����,a.סԺ��,a.����,a.������,A.���ն�,a.Ԥ�����,a.�������,a.ʣ���, a.���ò���,a.����,C.���� as �������," & _
        "        a.סԺҽʦ,a.�ѱ�,a.����ȼ�,a.����,a.����id," & _
        "        a.��Ժ����, a.��Ժ����, a.��������, a.�Ա�,a.����, a.��˱�־,a.Ӥ������ID,a.Ӥ������ID,b.Ӥ������,B.��� AS Ӥ�����," & _
        "        A.״̬,A.ҽ�Ƹ��ʽ,A.��������,A.������,A.��������ID,A.������������" & _
        " From T������Ϣ A,������������¼ B,������� C" & _
        " Where A.����id=b.����id and A.��ҳID=b.��ҳid and  A.����=C.���(+)"
    End If
    strSQL = strWithTable & vbCrLf & strSQL & vbCrLf
    strSQL = "" & _
    " Select * " & vbCrLf & _
    " From (" & strSQL & ")" & vbCrLf & _
    " Order by ����,NVL(Ӥ�����,0)"
    Set GetPatiRsByUnit = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ�б�", lng����ID, lng����ID, dtOutBegin, dtOutEnd)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function WriteInforToCard(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String, _
        ByVal objSquareCard As Object, ByVal lngCardTypeID As Long, _
        ByVal bytFunc As Byte, ByVal lng����ID As Long, lng����ID As Long) As Boolean
    '����:��סԺ��Ϣд�뿨��
    '��Σ�
    '    frmMain - ���ô���
    '    lngModul - ģ���
    '    strPrivs - Ȩ�޴�
    '    objSquareCard - ҽ�ƿ�����
    '    bytFun - 0:���1:סԺ
    Dim strExpend As String, lng������� As Long
    
    If lng����ID = 0 Or lng����ID = 0 Then Exit Function
    Err = 0: On Error GoTo errH:
    '����:56615
    If bytFunc = 0 Then
        If InStr(strPrivs, ";������Ϣд��;") = 0 Then Exit Function
        'Public Function zlMzInforWriteToCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, _
            ByVal lng����ID As Long, _
            ByVal lngBalanceID As Long, _
            Optional ByRef strExpend As String = "") As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:д������Ϣ�ӿ�
            '    frmMain Object  In  ���õ�������
            '    lngModule   Long    In  ���õ�ģ���
            '    lngCardTypeID   Long    In  ����д�����ID:
            '           1)����ˢ�������ID
            '           2)������ʱ,��Ҫѡ��ĳ�������ID
            '    lng����ID   Long    In  ����ID
            '    lngBalanceID    Long    In  �������(ĳ�ν�������)
            '    strExpend   String  In/Out  XML,����,���Ժ���չ
            ' ��������    True:���óɹ�,False:����ʧ��
            '����ʱ��:
            '         ҽ�ƿ����.�Ƿ�д��=1�ŵ���
        Call objSquareCard.zlMzInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng����ID, lng����ID, strExpend)
        '�����������û�н�����ţ�����ֱ�Ӵ�����ID
    Else
        If InStr(strPrivs, ";סԺ��Ϣд��;") = 0 Then Exit Function
        'Public Function zlZyInforWriteToCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, _
            ByVal lng����ID As Long, _
            ByVal lngBalanceID As Long, _
            Optional ByRef strExpend As String = "") As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:дסԺ��Ϣ�ӿ�
            '    frmMain Object  In  ���õ�������
            '    lngModule   Long    In  ���õ�ģ���
            '    lngCardTypeID   Long    In  ����д�����ID:
            '           1)����ˢ�������ID
            '           2)������ʱ,��Ҫѡ��ĳ�������ID
            '    lng����ID   Long    In  ����ID
            '    lngBalanceID    Long    In  ����ID(���Բ�����)
            '    strExpend   String  In/Out  XML,����,���Ժ���չ
            ' ��������    True:���óɹ�,False:����ʧ��
            '����ʱ��:
            '        ҽ�ƿ����.�Ƿ�д��=1�ŵ���
        Call objSquareCard.zlZyInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng����ID, lng����ID, strExpend)
    End If
    
    WriteInforToCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function CreatePlugIn(ByVal lngModule As Long, _
    Optional ByVal int���� As Integer) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugIn = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    If gobjPlugIn Is Nothing Then
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    End If
    If gobjPlugIn Is Nothing Then Exit Function
    
    Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngModule, int����)
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "Initialize")
        Set gobjPlugIn = Nothing
        Exit Function
    End If
    
    CreatePlugIn = True
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & _
            objErr.Number & vbCrLf & _
            objErr.Description, vbInformation, gstrSysName
    End If
    objErr.Clear
End Sub

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
    ByVal intTYPE As Integer, ByVal intMode As Integer, _
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
    If objPlugIn.CheckChargeItem(lngSys, lngModule, intTYPE, intMode, rsDetail, strExpend) = False Then
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

Public Sub CreatePublicExpenseObject(ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ò���
    '���:
    '����:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
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



Public Function CreatePublicExpenseBillOperation() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ò���
    '���:
    '����:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjPublicExpenseBillOperation Is Nothing Then
        Set gobjPublicExpenseBillOperation = CreateObject("zlPublicExpense.clsBillOperation")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Function
        End If
    Else
        CreatePublicExpenseBillOperation = True
        Exit Function
    End If
    If gobjPublicExpenseBillOperation Is Nothing Then Exit Function
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If gobjPublicExpenseBillOperation.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Function
    End If
    CreatePublicExpenseBillOperation = True
End Function

Public Function zlShowMsgBox(ByVal frmMain As Object, ByVal strInfo As String, Optional ByVal blnNoAsk As Boolean, Optional ByVal intTYPE As Integer) As VbMsgBoxResult
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ϣ��
    '���:frmMain-���õ�������
    '     strInfo=��ʾ��Ϣ,��Ҫ���Ѵ�����,����"^"��ʾ�س�,">"��ʾ����
    '     intType=��Ϣ������=0(ȱʡ)=MsgBox����,1-Ƥ������
    '     blnNoAsk="intType=0"ʱ��Ч����ʾ�Ƿ�ֻ��ʾһ��ȷ����ť,����ѯ�ʷ�ʽ��ʾ�Ǻͷ�
    '����:
    '    intType=0��vbIgnore=���Ҳ�����ʾ,vbCancel=���Ҳ�����ʾ,vbYes=��,vbNo=��
    '    intType=1��vbYes=����,vbNo=����,vbCancel=ȡ��
    '����:���˺�
    '����:2017-11-08 11:17:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPublicExpense Is Nothing Then
        Call CreatePublicExpenseObject(glngModul)
    End If
    If gobjPublicExpense Is Nothing Then GoTo GoMsgbox:

    Err = 0: On Error Resume Next
    zlShowMsgBox = gobjPublicExpense.zlShowMsgBox(frmMain, strInfo, blnNoAsk, intTYPE)
    If Err.Number = 438 Then GoTo GoMsgbox
    If Err <> 0 Then zlShowMsgBox = vbCancel
    Err = 0: On Error GoTo 0
    Exit Function
GoMsgbox:
    'ֱ��ʹ��Msgbox���ѿ�
    If blnNoAsk Then
        MsgBox strInfo, vbInformation + vbOKOnly, gstrSysName
        zlShowMsgBox = vbOK: Exit Function
    End If
    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
         zlShowMsgBox = vbIgnore
    Else
         zlShowMsgBox = vbCancel
    End If
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

Public Sub zlShowThreeSwapErrInfor(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ת�˼�������ҵ�������ʾ
    '����:Ƚ����
    'ʱ��:2014-12-2
    '����:
    '   bytType:0-ת�˼��,1-ת�˽���
    '   strXMLErrMsg:��ʽ����
    '            <OUT>
    '               <ERRMSG>������Ϣ</ERRMSG >
    '            </OUT>
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    '����������Ϣ
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '��ʾ������Ϣ
    If Trim(strValue) = "" Then
        If bytType = 0 Then
            strValue = vbCrLf & "����ת�ʼ�齻��ʧ�ܣ�"
        Else
            strValue = vbCrLf & "����ת�ʽ���ʧ�ܣ�"
        End If
    End If
    MsgBox strValue, vbExclamation + vbOKOnly, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function zlGetTimeDataFromTimes(ByVal str��ҳIds As String, ByRef int��ҳID As Integer, intInsure As Integer, _
    Optional strInsureName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����סԺ������Ϣ����ȡ��ҳID,���༰��������
    '���:str��ҳIDs:��ʽ:��ҳID|����|��������
    '����:int��ҳID
    '     intInsure-����
    '     strInsureName-��������
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2017-11-13 11:49:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    On Error GoTo errHandle
    varTemp = Split(str��ҳIds & "||||", "|")
    int��ҳID = Val(varTemp(0))
    intInsure = Val(varTemp(1))
    strInsureName = Trim(varTemp(2))
    zlGetTimeDataFromTimes = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlGetAllTims(ByVal str��ҳIds As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ�����ҳIDs,����ȡֻ������ҳID�Ĵ�
    '���:str��ҳIDs:��ʽ:��ҳID|����|��������,��ҳID1|����1|��������1,....
    '����:
    '����:ֻ�������漰��סԺ����
    '����:���˺�
    '����:2017-11-13 11:49:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, int��ҳID As Integer, intInsure As Integer
    Dim strAllTims As String, i As Long
    
    On Error GoTo errHandle
    
    
    varTemp = Split(str��ҳIds, ",")
    For i = 0 To UBound(varTemp)
        Call zlGetTimeDataFromTimes(varTemp(i), int��ҳID, intInsure)
        strAllTims = strAllTims & "," & int��ҳID
    Next
    If strAllTims <> "" Then strAllTims = Mid(strAllTims, 2)
    zlGetAllTims = strAllTims
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function TruncStringEx(ByVal strValue As String, Optional blnReverse As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ַ����⴦��
    '���:strValue-�ַ�ֵ
    '     blnReverse-�ߵ�
    '����:��ʽ���Ĵ�
    '����:���˺�
    '����:2017-11-13 09:53:05
    '˵��:�˹���Ϊ��ʱ�������пպ󣬲�Ӧ����ô����
    '    blnReverse=False
    '         1.��","�滻��"������"
    '         2.��"|"�滻��"������"
    '    blnReverse=true
    '         1.��"������"�滻��","
    '         2.��"������"�滻��"|"
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If blnReverse Then
        strValue = Replace(strValue, "������", ",")
        strValue = Replace(strValue, "������", "|")
    Else
        strValue = Replace(strValue, ",", "������")
        strValue = Replace(strValue, "|", "������")
    End If
    TruncStringEx = strValue
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ZlShowBillFormat(ByVal bytInvoiceKind As Byte, lblFormat As Label, ByVal intFormat As Integer)
    '���ܣ���ʾƱ�ݸ�ʽ����
    '��Σ�
    '   bytInvoiceKind - 0-סԺҽ�Ʒ��վ�,1-����ҽ�Ʒ��վ�
    '   lblFormat - ��ʾƱ�ݸ�ʽ�ı�ǩ����
    '   intFormat - Ʊ�ݸ�ʽ���
    '���أ�Ʊ�ݸ�ʽ������
    Dim strFormatName As String
    
    On Error GoTo errHandler
    strFormatName = ZlGetBillFormat(bytInvoiceKind, intFormat)
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

Public Function ZlGetBillFormat(ByVal bytInvoiceKind As Byte, ByVal intFormat As Integer) As String
    '���ܣ���ȡƱ�ݸ�ʽ����
    '��Σ�
    '   bytInvoiceKind - 0-סԺҽ�Ʒ��վ�,1-����ҽ�Ʒ��վ�
    '   intFormat - Ʊ�ݸ�ʽ���
    '���أ�Ʊ�ݸ�ʽ������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strRptName As String
    
    On Error GoTo errHandler
    If bytInvoiceKind = 0 Then
        strRptName = "ZL" & glngSys \ 100 & "_BILL_1137"
    Else
        strRptName = "ZL" & glngSys \ 100 & "_BILL_1137_2"
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
    '     rsSaveItems=��ǰ�������Ŀ�����ֶ�(�ֶ� :����ID����ҳID,�������, ���,�۸񸸺�,�շ�ϸĿID��������Ŀid������ �����Σ���׼���ۣ�Ӧ�ս�� ��
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

Public Sub zlChargeSaveAfter_Plugin(ByVal lngModule As Long, ByVal lng����ID, ByVal lng��ҳID As Long, ByVal bln���� As Boolean, _
                                    ByVal int��¼���� As Integer, ByVal strNos As String)
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


Public Function zlGetSaveDataItems_Plugin(ByVal objBills As ExpenseBill, ByRef rsItems As ADODB.Recordset, Optional blnBill As Boolean) As Boolean
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
    Dim int��� As Integer
    
    On Error GoTo errHandle
    
    Set rsItems = Nothing
    
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
    int��� = 0
    For Each objBillDetail In objBills.Details
        If objBillDetail.���� <> 0 Then
            int�۸񸸺� = 0
            For Each objBillIncome In objBillDetail.InComes
              int��� = int��� + 1 '��ǰ��¼���
               rsItems.AddNew
               rsItems!����ID = IIf(blnBill, objBills.Details(int���).����ID, objBills.����ID)
               rsItems!��ҳID = IIf(blnBill, objBills.Details(int���).��ҳID, objBills.��ҳID)
               rsItems!������� = 1
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
               rsItems!��������ID = objBills.��������ID
               rsItems!������ = objBills.������
               rsItems.Update
              If int�۸񸸺� = 0 Then int�۸񸸺� = int���
            Next     'ÿһ���շ���Ŀ
        End If
    Next
    
    zlGetSaveDataItems_Plugin = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCallReturnCashCheckInterface(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal str���� As String, ByVal strBalances As String, ByVal dblMoney As Double, _
    ByVal str������ˮ�� As String, ByVal str����˵�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����"zlReturnCashCheck"�ӿ�
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-09 14:21:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String
    If gobjSquare Is Nothing Then Call CreateSquareCardObject(frmMain, lngModule)
    If gobjSquare Is Nothing Then Exit Function
    If gobjSquare.objSquareCard Is Nothing Then Exit Function
    Err = 0: On Error GoTo errHandle
    With gobjSquare.objSquareCard
        If .zlReturnCashCheck(frmMain, lngModule, lngCardTypeID, str����, strBalances, dblMoney, str������ˮ��, str����˵��, strXMLExpend) = False Then
            MsgBox "�ӿڼ������ʧ�ܣ��޷����֣�", vbInformation, gstrSysName
           Exit Function
        End If
    End With
    zlCallReturnCashCheckInterface = True
    Exit Function
errHandle:
    If Err.Number = 438 Then
        MsgBox "ȱʧһ��ͨ��zlReturnCashCheck���ӿڣ��������֣�����ϵͳ����Ա��ϵ!", vbOKOnly, gstrSysName
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 Public Sub zlCloseSquareCardObject()
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
         Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub
Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = blnChildren And (Forms.Count = 0)
End Function
Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ͷ���Դ
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '������Ϊ0ʱ���ŷ���Դ
    If glngInstanceCount > 0 Then Exit Function
    
    Call zlCloseSquareCardObject  '�ͷ�CardSquare����
    
    Call zlCloseWindows   '�رմ���
    
    Err = 0: On Error Resume Next
    If Not gclsInsure Is Nothing Then Set gclsInsure = Nothing
    If Not gobjBillPrint Is Nothing Then Set gobjBillPrint = Nothing
    If Not gobjTax Is Nothing Then Set gobjTax = Nothing
    If Not grsABCNum Is Nothing Then Set grsABCNum = Nothing
    If Not gobjPati Is Nothing Then Set gobjPati = Nothing
    If Not grs�շ���� Is Nothing Then Set grs�շ���� = Nothing
    If Not gobjPlugIn Is Nothing Then Set gobjPlugIn = Nothing
    If Not gobjPublicDrug Is Nothing Then Set gobjPublicDrug = Nothing
    If Not gobjPublicExpense Is Nothing Then Set gobjPublicExpense = Nothing
    If Not gobjPublicExpenseBillOperation Is Nothing Then Set gobjPublicExpenseBillOperation = Nothing
    If Not grsҽ�Ƹ��ʽ Is Nothing Then Set grsҽ�Ƹ��ʽ = Nothing
    If Not grsOneCard Is Nothing Then Set grsOneCard = Nothing
    If Not grsSquareType Is Nothing Then Set grsSquareType = Nothing
    If Not gobjXml Is Nothing Then Set gobjXml = Nothing
    If Not gobjKernel Is Nothing Then Set gobjKernel = Nothing
    If Not gfrmMain Is Nothing Then Set gfrmMain = Nothing
    zlReleaseResources = True
End Function

Public Function zlSelectChargePatiFromInputName(ByVal frmMain As Object, ByVal strPrivsOpt As String, ByRef strInput As String, ByVal bln���в��� As Boolean, ByVal strUnitIDs As String, _
    ByVal intOutDay As Integer, ByRef lng����ID_Out As Long, Optional strErrMsg_out As String, Optional lngHwnd As Long, Optional lngHeight As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ĳ�����Ϣ����ȡ���������Ĳ�����Ϣ
    '���:frmMain-���õ�������
    '     strPrivsOpt-���ʲ��������Ȩ��
    '     strInput-�����ֵ
    '     intOutDay-���ҳ�Ժ��������
    '     strUnitIDs-���ҵĲ���IDs
    '     bln���в���-�Ƿ�������в���,����������в���, strUnitIDs����������
    '����:lng����ID_Out-�ӿڷ���trueʱ�����ز���ID,���򷵻�0
    '     strErrMsg_out-�ӿڷ���Falseʱ�����صĴ�����Ϣ,���򷵻ؿ�
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-10-08 18:04:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String, rsOutSel As ADODB.Recordset
    Dim blnCancel As Boolean, vRect As RECT
    
    On Error GoTo errHandle
    
    strErrMsg_out = ""
    'a.�Ƿ����ǿ�Ƽ���Ȩ��
    strWhere = ""
    If InStr(strPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(strPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        If intOutDay = 0 Then
            strWhere = " And nvl(A.��Ժ,0)=1"
        Else
            strWhere = " And (nvl(A.��Ժ,0)=1 Or  B.��Ժ����>Trunc(Sysdate)-" & intOutDay & ")"
        End If
    ElseIf InStr(strPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        If intOutDay = 0 Then
            strWhere = " And ((nvl(A.��Ժ,0)=1 And B.״̬<>3) Or (Nvl(X.�������,0)<>0 And nvl(A.��Ժ,0)=1  And B.״̬=3) )"
        Else
            strWhere = " And ((nvl(A.��Ժ,0)=1 And B.״̬<>3) Or (Nvl(X.�������,0)<>0 And ((nvl(a.��Ժ,0)=1 And B.״̬=3) Or (B.��Ժ����>Trunc(Sysdate)-" & intOutDay & "))))"
        End If
    ElseIf InStr(strPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        If intOutDay = 0 Then
            strWhere = " And ((nvl(A.��Ժ,0)=1  And B.״̬<>3) Or (Nvl(X.�������,0)=0 And nvl(A.��Ժ,0)=1  And B.״̬=3))"
        Else
            strWhere = " And ((nvl(A.��Ժ,0)=1  And B.״̬<>3) Or (Nvl(X.�������,0)=0 And ((nvl(A.��Ժ,0)=1 And B.״̬=3) Or (B.��Ժ����>Trunc(Sysdate)-" & intOutDay & "))))"
        End If
    Else
        'û��Ȩ�޶Գ�Ժ��Ԥ��Ժ���˽���
        strWhere = " And Nvl(A.��Ժ,0)=1 And Nvl(B.״̬,0)<>3 "
    End If
    
    
    'b.�Ƿ���Լ����в�������
    If Not bln���в��� Then
        If InStr(1, strUnitIDs, ",") = 0 Then
            strWhere = strWhere & " And B.��ǰ����ID+0=[3]"
        Else
            strWhere = strWhere & " And B.��ǰ����ID+0 IN(Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
    
    'c.�Ƿ����۲��˼���Ȩ��
    If (InStr(strPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(strPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        strWhere = strWhere & " And Nvl(B.��������,0) IN(0,1,2)"
    ElseIf InStr(strPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        strWhere = strWhere & " And Nvl(B.��������,0) IN(0,1)"
    ElseIf InStr(strPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        strWhere = strWhere & " And Nvl(B.��������,0) IN(0,2)"
    Else
        strWhere = strWhere & " And Nvl(B.��������,0)=0"
    End If
        
    strSQL = _
    " Select Rownum as ID, A.����ID,Decode(nvl(A.��Ժ,0),1,'��','') as ��Ժ, nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,nvl(b.����,A.����) as ����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ, " & _
    "       to_Char(B.��Ժ����,'yyyy-mm-dd') as ��Ժ����,to_Char(B.��Ժ����,'yyyy-mm-dd') ��Ժ����," & _
    "       A.סԺ��,B.��Ժ���� as ����,X.�������,C.���� as ��ǰ����,B.��������,B.��ע" & _
    " From ������Ϣ A,������ҳ B,������� X,���ű� C" & _
    " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID and B.��ǰ����ID=C.ID(+)" & strWhere & _
    "       And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) and X.����(+)=1 and X.����(+)=2 And A.ͣ��ʱ�� is NULL  " & _
    "       And A.���� like [1] " & _
    " Order by ��Ժ Desc,��Ժ����"
    
    If lngHwnd = 0 Then
        Set rsOutSel = zlDatabase.ShowSQLSelect(frmMain, strSQL, 0, "����ѡ����", False, "", "��ѡ����", False, False, False, 0, 0, 0, blnCancel, False, True, strInput & "%", "", Val(strUnitIDs), strUnitIDs, "bytSize=1")
    Else
        vRect = zlControl.GetControlRect(lngHwnd)
        Set rsOutSel = zlDatabase.ShowSQLSelect(frmMain, strSQL, 0, "����ѡ����", False, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, lngHeight, blnCancel, False, True, strInput & "%", "", Val(strUnitIDs), strUnitIDs, "bytSize=1")
    End If
    If blnCancel Then Exit Function
    
    If Not rsOutSel Is Nothing Then
        If rsOutSel.State = 1 Then
            If rsOutSel.EOF = False Then
                lng����ID_Out = Val(rsOutSel!����ID)
                Set rsOutSel = Nothing
                zlSelectChargePatiFromInputName = True: Exit Function
            End If
        End If
    End If
    strErrMsg_out = "δ�ҵ��������ϡ�" & strInput & "���Ĳ���,�����Ƿ�������ȷ!"
    Set rsOutSel = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


