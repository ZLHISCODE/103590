Attribute VB_Name = "mdl�ϳ�����"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
'    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;3-������������GetNextNO();
'    99-���н������Ӹ��Ӳ���(���°�)

Public Enum ҵ������_�ϳ�����
    ����籣����_���� = 0
    ��òα���Ա����_����
    �����Ա����_ҽ����_����
    �����Ա����_����_����
     
    ��ȡ�ʻ����_����
    ���κ�����_����
    �����κ�����_����
    �Ͽ��κ�����_����
    �����ʻ�����_����
    �����ʻ�����_���_����
    ��ʼ��_����
    ���ѳ���_����
    ���ؽ��׼�¼_����
    ��ȡ��������_����
    ����Ԥ����_����
    �޸�����_����
    
    ����籣����_סԺ_����
    ��Ժ�Ǽ�_����
    ȡ����Ժ�Ǽ�_����
    ��ȡ������¼��_����
    ���Ӵ�������_����
    ������������_����
    ���Ӵ�����ϸ_����
    ��Ժ����_����
    ȡ����Ժ����_����
    ����סԺ�Ż�ȡ��¼��_����
    ��ӡ���㱨��_����
    סԺ���˿�������_����
End Enum
Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    �������� As String                      'Ĭ�ϵ��籣��������
    ��ϸʱʵ�ϴ� As Boolean
    ���ݲ��Ȳ��ɽ���  As Boolean
    ������뷽ʽ As Boolean                 'false��ʾ�ֹ�����,true��ʾ����׼��������
    
End Type
Public InitInfor_�ϳ����� As InitbaseInfor

Private Type �������
    ҽ������    As String
    ҽ��֤��    As String
    ���֤����  As String
    ��¼��      As String
    ����        As String
    �Ա�        As String
    ��������    As String
    ����        As Integer
    ��λ����    As String
    ��������    As String
    
    �ʻ����    As String
    �����ܶ�    As Double
    ����        As String
    �籣����    As Long
    ����ID      As Long
    ֱ����    As Boolean
    
    ����ID As String
    
    �μӹ������� As String
    �������� As String
    ְ�񼶱� As String
    ְ�Ƽ��� As String
    ��Ա���� As String
    ��ؾ�ס��־ As String
    ��λID As String
    ���� As String
    סԺ���� As String
    ����ҽ�Ʊ�־ As String
    ����ҽ�Ʊ�־ As String
    ����Ա��־  As String
    ��������״̬ As String
    �������״̬ As String
    ����Ա����״̬  As String
    ����סԺ����    As String
    �����ѱ������ As String
    �ɷ�����    As String
    ��ȡʱ��    As String
    סԺ��¼��  As String
    
    strסԺ�� As String
    strסԺ��Ϣ  As String
    �����ʻ�֧�� As Double '�������ʱ����
End Type

Private Type ��������
    ���� As String
    ����    As String
    ����ǰ�ʻ���� As Double
    �����ʻ�֧����� As Double
    �Էѽ�� As Double
    ���Ѻ��ʻ���� As Double
    ����ʱ��  As String
    ǰ�˵��ݺ�  As String
    ���ĵ��ݺ�  As String
    ������  As String
    ����Ա����  As String
    ǰ������  As String
    
    ����ID As Long
    �����־ As Byte    '0-����,1-סԺ
    ����������� As Double
    ���䱨����� As Double
    ����Ա������� As Double
End Type
Private g�������� As ��������
Public g�������_�ϳ����� As �������
Public gcnOracle_�ϳ����� As ADODB.Connection     '�м������

Private gbln������� As Boolean
Private gbln�Ѿ���ʼ As Boolean             '�Ѿ�����ʼ����.

'1.����籣����_���б�ź������б�
Private Declare Function GetSBJGLB Lib "CDGK_GRZH.dll" Alias "GETSBJGLB" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETSBJGLB:PCHAR
'����: ����籣����_���б�ź������б�
'��ڲ���: ��
'���ڲ���: ��
'����: A�籣�������+�м����+A�籣��������+�м����+B�籣�������+�м����+B�籣��������+����
'===============================================================================================================

'2����òα���Ա�Ļ�������
'   A.����
Private Declare Function GETKZL Lib "CDGK_GRZH.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETKZL:PCHAR
'����: ��òα���Ա�Ļ�������
'��ڲ���:
'���ڲ���: ��
'����: OK(�������Ϣ)@$ҽ������||ҽ��֤��||���˼�¼��||����||���֤����||��λ����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��
'===============================================================================================================

'   B.סԺ(��òα���Ա�Ļ�������(����ҽ��֤��))
Private Declare Function GETRYJBZL Lib "CDGK_YB.dll" (ByVal strҽ��֤�� As String, ByVal str�������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETRYJBZL(AYBZH, ABXJGBH:PCHAR):PCHAR;
'����:ͨ������ҽ��֤�Ŵ��籣����ҽ�����ݿ���ȡҽ�����˵Ļ������ϡ�
'��ڲ���:AYBZH   PCHAR   �α���Ա��ҽ��֤��
'         ABXJGBH PCHAR   �α���Ա���ڵı��ջ������
'���ڲ���: ��
'����:
'OK(�������Ϣ)@$���˼�¼��||���֤��||����||�Ա�||�������ڣ���ʽ��YYYY��MM��DD��||��������||��������||ְ�񼶱�||ְ�Ƽ���||��Ա��� (��ְ \ ����)||��ؾ�ס��־||��λID||��λ����||����||����||����||תԺ��ҽԺID||סԺ��¼��
'===============================================================================================================

'   C.סԺ(��òα���Ա�Ļ�������(ֱ�Ӷ�ҽ��IC��))
Private Declare Function GETICCARD Lib "CDGK_YB.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETKZL:PCHAR
'����: ��òα�IC������Ϣ
'��ڲ���:
'���ڲ���: ��
'����: OK(�������Ϣ)@$ҽ��֤��||����||���˼�¼��||����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��||���֤��||��λ����
'===============================================================================================================

'3.�����ʻ�����ѯ
Private Declare Function GETZHYE Lib "CDGK_GRZH.dll" (ByVal str�������� As String, ByVal strPassWord As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETZHYE(YBJGBH,CPASSWORD:PCHAR):PCHAR
'����: ��óֿ���Ա�����ʻ����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'         CPASSWORD   PCHAR   �ֿ��˿�����
'���ڲ���: ��
'����:  OK(�������Ϣ)@$�����ʻ����
'===============================================================================================================

'4.���κ������Ƿ����ӳɹ�
Private Declare Function CheckCon Lib "CDGK_GRZH.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION CHECKCON:PCHAR;
'����:���κ������Ƿ����ӳɹ�
'��ڲ���:
'����:OK�������Ϣ
'===============================================================================================================

'5.�����κ�����
Private Declare Function RasDial Lib "CDGK_GRZH.dll" (ByVal str�������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION RASDIAL(SBXJGBH:PCHAR):PCHAR
'����:�κ���ѡ����籣�֣����佨������
'��ڲ���:SBXJGBH PCHAR   ���ջ������
'����:  �ɹ�    ������HIS�κ���״̬����ʾ"����"
'       ʧ�� ������Ϣ
'===============================================================================================================

'6.�Ͽ����籣�ֵ�����
Private Declare Function DisDial Lib "CDGK_GRZH.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION DISDIAL:PCHAR
'����:�κ���ѡ����籣�֣����佨������
'��ڲ���:
'����:
'===============================================================================================================

'7.�����ʻ�����
Private Declare Function GRZHXF_CF Lib "CDGK_GRZH.dll" (ByVal str������� As String, ByVal str������ As String, _
            ByVal str��ϸ���� As String, ByVal strPassWord As String, ByVal str����Ա As String) As String
'===============================================================================================================
'ԭ��:Function GRZHXF_CF()(YBJGBH,CFH:PCHAR;CFMXDATA:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'����:���и����ʻ�����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'        CFH PCHAR   ������
'        CFMXDATA    PCHAR   ������ϸ����    ��ʽ˵��������1(ҽ��ҩƷ���+�м����+����+�м��������+)+�м����+        ����        ����N(ҽ��ҩƷ���+�м����+����+�м����+����
'        CPASSWORD   PCHAR   �ֿ��˿�����
'        CCZYXM  PCHAR   ����Ա����
'����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
'   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
'===============================================================================================================


'8.�����ʻ����ѣ�ֱ���������ѽ�

Private Declare Function GRZHXF_JE Lib "CDGK_GRZH.dll" (ByVal str������� As String, ByVal str��� As String, _
             ByVal strPassWord As String, ByVal str����Ա As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GRZHXF_JE(YBJGBH,XFJE:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'����:���и����ʻ�����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'    XFJE    PCHAR   ���ѽ��(��֤ΪС�������ұ�����λС��)
'    CPASSWORD   PCHAR   �ֿ��˿�����
'    CCZYXM  PCHAR   ����Ա����
'����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
'   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
'===============================================================================================================

'����Ԥ�ʻ�����
Private Declare Function GRZHXF_CFPRE Lib "CDGK_GRZH.dll" (ByVal str������� As String, ByVal str��ϸ���� As String, _
             ByVal strPassWord As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GRZHXF_CFPRE(YBJGBH,CFMXDATA,CPASSWORD:PCHAR):PCHAR
'����:����Ԥ�ʻ�����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'    CFMXDATA    PCHAR   ������ϸ����    ��ʽ˵��������1(ҽ��ҩƷ���+�м����+����+�м��������+)+�м����+
'    ����
'    ����N(ҽ��ҩƷ���+�м����+����+�м����+����
'    CPASSWORD   PCHAR   �ֿ��˿�����
'����:OK@$�����ʻ�֧�����@$�Ը����
'===============================================================================================================

'9.���ѳ���

Private Declare Function XFCZ Lib "CDGK_GRZH.dll" (ByVal str������� As String, ByVal str���ĵ��ݺ� As String, _
             ByVal strPassWord As String, ByVal str����Ա As String) As String
'===============================================================================================================
'ԭ��:FUNCTION XFCZ(YBJGBH ��CZXDJH:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'����:���Ѿ����ѵļ�¼���г�����
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'        cZXDJH  PCHAR   ���ĵ��ݺ�(����ʱ����)
'        CPASSWORD   PCHAR   �ֿ��˿�����
'        CCZYXM  PCHAR   ����Ա����
'����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
'   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
'===============================================================================================================




'10.�ֿ���Ա���п������޸�

Private Declare Function CHANGPASSWORD Lib "CDGK_GRZH.dll" (ByVal str������� As String, ByVal str������ As String, _
             ByVal str������ As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CHANGPASSWORD(YBJGBH ,COLDPASS,CNEWPASS:PCHAR):PCHAR
'����:�ֿ���Ա���п������޸�
'��ڲ���:YBJGBH  PCHAR   ���ջ������
'    COLDPASS    PCHAR   ������
'    CNEWPAS PCHAR   ������
'����:(OK�������Ϣ)
'===============================================================================================================



'11.ǰ�˳�ʼ��
Private Declare Function QDINIT Lib "CDGK_GRZH.dll" (ByVal str������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION QDINIT(AYBJGBH:STRING):PCHAR;
'����:��ǰ�˽��г�ʼ���������Ա�ǰ�˸����ʻ�������ˮ�������ı���һ�¡�
'��ڲ���:AYBJGBH PCHAR   ҽ���������
'����:(OK�������Ϣ)
'===============================================================================================================


'12.���ؽ��׼�¼
Private Declare Function DOWNJYJL Lib "CDGK_GRZH.dll" (ByVal str������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION DOWNJYJL(AYBJGBH:PCHAR):PCHAR
'����:������ҽ�����ݿ��ƻ��󣬴��������ر�����ǰ�����л�δ��˽�������Ѽ�¼��
'��ڲ���:AYBJGBH PCHAR   ҽ���������
'����:(OK�������Ϣ)
'===============================================================================================================


'*****************************************************************************************************************************************
'**סԺ����
'*****************************************************************************************************************************************
'1.����籣����_���б�ź������б�
Private Declare Function GetSBJGLB1 Lib "CDGK_YB.dll" Alias "GETSBJGLB" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETSBJGLB:PCHAR
'����: ����籣����_���б�ź������б�
'��ڲ���: ��
'���ڲ���: ��
'����: A�籣�������+�м����+A�籣��������+�м����+B�籣�������+�м����+B�籣��������+����
'===============================================================================================================



'1.��Ժ�Ǽ�
Private Declare Function RYDJ Lib "CDGK_YB.dll" (ByVal strҽ��֤�� As String, ByVal strסԺ�� As String, ByVal str�������� As String, ByVal str������� As String, ByVal str����Ա���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION RYDJ(AZBZH,AZYH,;ARYZL,ABXJGBH,ACZYXM:PCHAR):PCHAR;
'����: ���籣����ҽ�����ݿ��ҽԺ����ҽ�����ݿ��ж�סԺ��ҽ�����˽��еǼǡ�
'��ڲ���:
'         AYBZH  PCHAR ----ҽ��֤��
'         AZYH   PCHAR ----סԺ��
'         ARYZL  PCHAR ----�α���Ա�ĸ������ϣ�
'             �α���Ա�ĸ������ϸ�ʽ: ���м�����ָ�
'             ��Ժ���ڡ�����ʽ��YYYY��MM��DD��
'             ���ֱ��
'             ��Ժ��� (�ش�)
'             ��Ժָ�� (�ش�)
'             ����
'             ����
'             ����
'         ABXJGBH:   PCHAR �D�D�α���Ա���ڵ��籣�������
'         ACZYXM:    PCHAR --������Ա
'���ڲ���: ��
'˵��:
'   �α���Ա����Ժ����:ҽ��֤��||סԺ��||��Ժ���ڣ���ʽ��YYYY-MM-DD��||���ֱ��||��Ժ���||��Ժָ��||����||����||����||�籣��������||������Ա
'����:���� �ɹ�: 'OK'+�м����+ҽ������סԺ��¼��
'          ʧ��:     ������Ϣ
'===============================================================================================================


'2.ȡ��סԺ
Private Declare Function ZYQX Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION ZYQX(AZYH:PCHAR):PCHAR;
'����: ��û����ʽ����ǰ�����籣����ҽ�����ݿ��ҽԺ����ҽ�����ݿ���ɾ��ҽ������סԺ��¼��
'��ڲ���: AZYH    PCHAR   סԺ��
'���ڲ���: ��
'����:���ر�־
'     �ɹ�: 'OK'
'     ʧ��: ������Ϣ
'===============================================================================================================


'3.���һ���µĴ�����¼�ţ��Ա�֤������Ψһ�ԡ�
Private Declare Function GETNEWCFID Lib "CDGK_YB.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETNEWCFID:PCHAR
'����: ���һ���µĴ�����¼�ţ��Ա�֤������Ψһ�ԡ�
'��ڲ���:
'���ڲ���: ��
'����:���ر�־:OK(�������Ϣ)@$������¼��
'===============================================================================================================




'4.����һ����������
Private Declare Function AddCFJL Lib "CDGK_YB.dll" _
    Alias "ADDCFJL" (ByVal strסԺ�� As String, ByVal str������ As String, ByVal str��� As String, ByVal str�������� As _
    String, ByVal strҽ�� As String, ByVal str���� As String, ByVal strҩƷ As String, ByVal str���� As _
    String, ByVal str���� As String, ByVal str��Ʒ�� As String) As String
'===============================================================================================================
'ԭ��:function ADDCFJL(AZYH,ACFID,ACFMXID,ACFRQ,AYS,AKS,AYPBH,ASL,ADJ,ASPM:PCHAR):PCHAR
'����:����һ����������,���뱣֤ACFID��ACFMXIDΨһ
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'    ACFID   PCHAR   ��������(���������ݿ��б�֤Ψһ)
'    ACFMXID PCHAR   ��ϸ���(��һ�������б�֤Ψһ)
'    ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
'    AYS     PCHAR   ҽ��
'    AKS     PCHAR   ����
'    AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
'    ASL     PCHAR   ����(����Ϊ����)
'    ADJ     PCHAR   ����
'    ASPM    PCHAR   ��Ʒ����ҽԺ��ҩƷ�����ش���
'���ڲ���: ��
'����:''OK'@$�Էѱ���@$�Էѽ��
'===============================================================================================================

'5.������������
Private Declare Function CFCS Lib "CDGK_YB.dll" (ByVal strסԺ�� As String, ByVal str������¼�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CFCS(AZYH:PCHAR;ACFID:PCHAR):PCHAR
'����:���籣����ÿ��Ĵ���������籣���������ݿ⴫�䣨ͬһ���������Զ���ظ����䣬��һ�δ�������ݽ�����ǰһ�δ�������ݣ���
'��ڲ���:
'       AZYH    PCHAR   סԺ��
'       ACFID   PCHAR   ������¼�ţ�ͨ������ADDCFJL���ص�ֵ��
'���ڲ���: ��
'����:'OK�������Ϣ
'===============================================================================================================



'6.���Ӵ�����ϸ
Private Declare Function AddCFMX Lib "CDGK_YB.dll" (ByVal str������¼�� As String, ByVal strҽ������ As String, ByVal str���� As String, ByVal str���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION ADDCFMX(ACFID,AYPBH,ASL,ADJ:PCHAR):PCHAR;
'����:����һ��������ϸ��
'��ڲ���:
'    ACFID   PCHAR   ������¼��
'    AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
'    ASL PCHAR   ����(����Ϊ����)
'    ADJ PCHAR   ����
'���ڲ���: ��
'����:OK@$������ϸ��¼��@$�Էѱ���@$�Էѽ��
'===============================================================================================================



'7.��Ժ����
Private Declare Function CYJS Lib "CDGK_YB.dll" (ByVal strסԺ�� As String, ByVal str��Ժ���1 As String, ByVal str��Ժ���2 As String, ByVal str��Ժ���3 As String, ByVal str��Ժ���4 As String, ByVal str����Ч�� As String, ByVal str��Ժ���� As String, ByVal strԤ���־ As String) As String
'===============================================================================================================
'ԭ��:FNCTION CYJS(AZYH:PCHAR; ISPREV:INTEGER;ZLXG,CYZD,CYRQ:PCHAR):PCHAR
'����:סԺ�α����˳�Ժ��סԺ��Ԥ����,��ΪԤ����ʱ, ZLXG,CYZD,CYRQ������������Ҫ����
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'    ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
'    ZLXG    PCHAR   ����Ч��
'    CYZD    PCHAR   ��Ժ���1||��Ժ���2||��Ժ���3||��Ժ���4
'    CYRQ    PCHAR   ��Ժ���ڣ�YYYY-MM-DD��
'���ڲ���: ��
'����:OK@$סԺ���ý�����@$�����ֶ���ϸ
'   ˵��:
'       סԺ���ý�����:����ҽ�ƴ���״̬||�𸶽��||�����ⶥ���||������������||�����ѱ������||�����������||���䱨�����||����Ա�������||����ҽ�ƴ���״̬||����Ա����״̬||���䱨������||����Ա��������||����סԺ����||�������||����ҩƷ��||�������Ʒ�||��������||�������||����ҩƷ��||�������Ʒ�||����������||�����Ը�||�������||����ҩƷ��||�������Ʒ�||��������||�����ϼ�||����֧��
'       �����ֶ���ϸ(����):����||����||����ʼ���||����ֹ���||���λ���||���α�������||���α������||�����Ը����@$.......
'===============================================================================================================


'8.ȡ����Ժ����
Private Declare Function CYJSQX Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CYJSQX(AZYH:PCHAR):PCHAR
'����:ȡ���α����˳�Ժ����
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'9.����סԺ�ŵõ�סԺ��¼��

Private Declare Function GETZYIDBYZYBH Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETZYIDBYZYBH(AZYH:PCHAR):PCHAR
'����:����סԺ�ŵõ�סԺ��¼�š�
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'���ڲ���: ��
'����:'OK@$סԺ��¼��
'===============================================================================================================

'10.��ӡ��Ժ���㱨����
Private Declare Function JSReport Lib "CDGK_YB.dll" (ByVal str��ʼסԺ�� As String, ByVal str����סԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION JSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL
'����: ��ӡ���㱨���˱���ȽϺ�����Դ�������HISϵͳ����ӡ,��ʽ���ҹ�˾�ṩ��
'��ڲ���:
'    ASTARTZYH   PCHAR   ��ӡ��ʼסԺ��
'    AENDZYH PCHAR      ��ӡ����סԺ��
'   ע��:
'    1 ?����סԺ��֮�����е�סԺ��¼����Ϊͬһ���籣��?
'    2����ֻ��ӡһ��סԺ�ŵı���ʱ����������ֵһ����
'���ڲ���: ��
'����:����ע�ⷵ��ֵ
'===============================================================================================================


'11.סԺ���˿���������Ա��������
Private Declare Function CWJSREPORT Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CWJSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL;
'����:סԺ���˿���������Ա��������
'��ڲ���:
'   AZYH    PCHAR   סԺ��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'12 ��ȡ��������_����

Private Declare Function GETJCXX Lib "CDGK_YB.dll" (ByVal str�������� As String, ByVal str���ر�־ As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETJCXX(SBXJGBH:PCHAR;DOWNALL:INTEGER):PCHAR
'����:��ָ�����籣������ȡ�������ϡ�
'��ڲ���:
'    SBXJGBH PCHAR   ���ջ������
'    DOWNALL PCHAR   ��ֵΪ0ʱ��ʾ���ر���ҽ�����ݿ���û�еĻ������ϣ�Ϊ����ʱ��ʾȫ����������
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'20.סԺ���˿���������Ա��������
Private Declare Function GETNEWRYZL Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETNEWRYZL(AZYH:PCHAR):PCHAR;STDCALL;
'����:סԺ���˿���������Ա�������ϡ�
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'����:OK@$������Ϣ
'===============================================================================================================



Public Function ҽ����ʼ��_�ϳ�����() As Boolean
    Dim strReg As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_�ϳ�����.ģ������ = True
    Else
        InitInfor_�ϳ�����.ģ������ = False
    End If
    
   Call GetRegInFor(g����ģ��, "ҽ��", "�籣��������", strReg)
   
   InitInfor_�ϳ�����.�������� = strReg
   g�������_�ϳ�����.�������� = strReg
   
   
   
   If strReg = "" Then
        MsgBox "��δ����Ĭ�ϵ��籣�������룬�����������!"
        Exit Function
   End If
   
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_�ϳ�����)
    InitInfor_�ϳ�����.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
    
    #If gverControl >= 4 Then
        InitInfor_�ϳ�����.������뷽ʽ = Not (Val(zlDatabase.GetPara(65, glngSys, , 1)) = 1)
    #Else
        InitInfor_�ϳ�����.������뷽ʽ = Not (Val(GetPara(65, glngSys, , , 1)) = 1)
    #End If
'    If Val(GetPara(65, glngSys, , , 1)) = 1 Then
'        InitInfor_�ϳ�����.������뷽ʽ = Not (Val(GetPara(65, glngSys, , , 1)) = 1)
'    Else
'        InitInfor_�ϳ�����.������뷽ʽ = True
'    End If
    
    '�м������
    gstrSQL = "select ������,����ֵ from ���ղ��� where  ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�山ҽ��", TYPE_�ϳ�����)
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "����Աֱ�¸����ʻ�"
                 g�������_�ϳ�����.ֱ���� = Nvl(rsTemp("����ֵ"), 0) = 1
            Case "��ϸʱʵ�ϴ�"
                InitInfor_�ϳ�����.��ϸʱʵ�ϴ� = IIf(Nvl(rsTemp!����ֵ, 1) = 1, 1, 0)
           Case "�ȽϽ�������"
                 InitInfor_�ϳ�����.���ݲ��Ȳ��ɽ��� = IIf(Nvl(rsTemp!����ֵ, 1) = 1, 1, 0)
        End Select
        rsTemp.MoveNext
    Loop

    
    Set gcnOracle_�ϳ����� = New ADODB.Connection
    If OraDataOpen(gcnOracle_�ϳ�����, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
   '�����κ�����
   If gbln�Ѿ���ʼ = False And gbln������� Then
       If ������������() = False Then Exit Function
   End If
   
   If gbln������� Then
        '���κ�����
        If ҵ������_�ϳ�����(���κ�����_����, "", strOutput) = False Then
             Exit Function
        End If
    End If
    gbln�Ѿ���ʼ = True
    ҽ����ʼ��_�ϳ����� = True
End Function

Public Function ҽ����ֹ_�ϳ�����() As Boolean
    Dim strOutput As String
    
    If gcnOracle_�ϳ�����.State = 1 Then
        gcnOracle_�ϳ�����.Close
    End If
    If gbln������� Then
         '�����κ�����
        Call ҵ������_�ϳ�����(�Ͽ��κ�����_����, "", strOutput)
    End If
    Err = 0
    On Error Resume Next
    ҽ����ֹ_�ϳ����� = True
End Function

Public Function ��ݱ�ʶ_�ϳ�����(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Err = 0
    On Error GoTo errHand:
    '��������20050420   ����������Һŵ�ҽ��ҵ��
    If bytType = 0 Or bytType = 3 Then
    Exit Function
    End If
    ��ݱ�ʶ_�ϳ����� = frmIdentify�ϳ�����.GetPatient(bytType, lng����ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_�ϳ����� = ""
End Function


Public Function �������_�ϳ�����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_�ϳ�����)
    
    If rsTemp.EOF Then
        �������_�ϳ����� = 0
    Else
        �������_�ϳ����� = rsTemp("�ʻ����")
    End If
End Function
Public Function �����������_�ϳ�����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim str��ϸ As String
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    g�������_�ϳ�����.�����ܶ� = 0
    str��ϸ = ""
    With rs��ϸ
        Do While Not .EOF
            gstrSQL = "Select * From ҽ��֧����Ŀ where ����=[1] and ����=[2] and �շ�ϸĿid=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ��ҽ��֧����Ŀ", TYPE_�ϳ�����, g�������_�ϳ�����.�籣����, CLng(Nvl(!�շ�ϸĿID, 0)))
            
            If Val(Nvl(rs��ϸ("ʵ�ս��"), 0)) <> 0 Then
                If rsTemp.EOF Then
                    str��ϸ = str��ϸ & "@$" & ""
                Else
                    str��ϸ = str��ϸ & "@$" & Nvl(rsTemp!��Ŀ����)
                End If
                str��ϸ = str��ϸ & "||" & Nvl(!����, 0)
                str��ϸ = str��ϸ & "||" & Nvl(!����, 0)
            End If
'            BJGBH  PCHAR   ���ջ������
'            '    CFMXDATA    PCHAR   ������ϸ����    ��ʽ˵��������1(ҽ��ҩƷ���+�м����+����+�м��������+)+�м����+
'            '    ����
'            '    ����N(ҽ��ҩƷ���+�м����+����+�м����+����
'            '    CPASSWORD   PCHAR   �ֿ��˿�����

            g�������_�ϳ�����.�����ܶ� = g�������_�ϳ�����.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    If str��ϸ <> "" Then
        str��ϸ = Mid(str��ϸ, 3)
    End If
    g�������_�ϳ�����.�����ʻ�֧�� = 0
    If g�������_�ϳ�����.ֱ���� Then
        If g�������_�ϳ�����.�����ܶ� > g�������_�ϳ�����.�ʻ���� Then
            str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & g�������_�ϳ�����.�ʻ���� & ";1"
        Else
            str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & g�������_�ϳ�����.�����ܶ� & ";1"
        End If
    Else
         StrInput = g�������_�ϳ�����.��������
         StrInput = StrInput & vbTab & str��ϸ
         StrInput = StrInput & vbTab & g�������_�ϳ�����.����
         If ҵ������_�ϳ�����(����Ԥ����_����, StrInput, strOutput) = False Then
            Exit Function
         End If
         strArr = Split(strOutput, "||")
         '�����ʻ�֧�����||�Ը����
         
        str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & Val(strArr(0)) & ";0"
        g�������_�ϳ�����.�����ʻ�֧�� = Val(strArr(0))
    End If
    �����������_�ϳ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ������������() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Static str������� As String
    Dim StrInput As String, strOutput As String
    ������������ = False
    
    Err = 0: On Error GoTo errHand:
    If str������� <> g�������_�ϳ�����.�������� Then
        '��������Ƿ���������
        If str������� = "" Then
            '�����һ��Զ��,��Ͽ�
            If ҵ������_�ϳ�����(�����κ�����_����, g�������_�ϳ�����.��������, strOutput) = False Then
                Exit Function
            End If
        Else
            '��ʾ�������������ϵĲ���,����Ͽ�����
            Call ҵ������_�ϳ�����(�Ͽ��κ�����_����, "", strOutput)
            If ҵ������_�ϳ�����(�����κ�����_����, g�������_�ϳ�����.��������, strOutput) = False Then Exit Function
        End If
        If ҵ������_�ϳ�����(���κ�����_����, "", strOutput) = False Then Exit Function
    Else
        If ҵ������_�ϳ�����(���κ�����_����, "", strOutput) = False Then
            '�����½�����������
            If ҵ������_�ϳ�����(�����κ�����_����, g�������_�ϳ�����.��������, strOutput) = False Then
                Exit Function
            End If
        End If
    End If
    str������� = g�������_�ϳ�����.��������
    ������������ = True
    Exit Function
errHand:
        If ErrCenter = 1 Then Resume
End Function
Public Function �������_�ϳ�����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim StrInput As String, strOutput As String
    Dim lng����ID  As Long
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strArr As Variant
'    If ������������() = False Then Exit Function
'
    On Error GoTo errHandle
    
    Call DebugTool("�����������")
    
    gstrSQL = "" & _
        "   Select a.*,a.����*a.���� as ����,a.ʵ�ս��/(nvl(a.����,1)*nvl(a.����,1)) as ���� " & _
        "   From ������ü�¼ a " & _
        "   Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
        
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ϸ��¼", lng����ID)
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If

    lng����ID = rs��ϸ("����ID")
    
    If g�������_�ϳ�����.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣"
        Exit Function
    End If
    If ҵ������_�ϳ�����(��ʼ��_����, g�������_�ϳ�����.��������, strOutput) = False Then Exit Function
    
    g��������.����ID = lng����ID
    g��������.�����־ = 0
    'д����ϸ
    If ������ϸд��(rs��ϸ, False) = False Then Exit Function
    
    If g�������_�ϳ�����.ֱ���� = False Then
'        '��ʾ��ᴦ��ʽ
        Call ���㷽ʽ����
        DebugTool "�����Ѿ���ʾ���"
    End If
    DebugTool "��ʼ��������"
    
    
    Dim dbl�����ʻ� As Double
    dbl�����ʻ� = cur�����ʻ�
    If dbl�����ʻ� <> g��������.�����ʻ�֧����� Then
        If g�������_�ϳ�����.ֱ���� Then
            '���¸����ʻ�֧��
            '��:YBJGBH  PCHAR   ���ջ������
            '    XFJE    PCHAR   ���ѽ��(��֤ΪС�������ұ�����λС��)
            '    CPASSWORD   PCHAR   �ֿ��˿�����
            '    CCZYXM  PCHAR   ����Ա����
            '����:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
            StrInput = g�������_�ϳ�����.��������
            StrInput = StrInput & vbTab & Format(dbl�����ʻ�, "###0.00;-###0.00;0.00;0.00")
            StrInput = StrInput & vbTab & g�������_�ϳ�����.����
            StrInput = StrInput & vbTab & gstrUserName
            If ҵ������_�ϳ�����(�����ʻ�����_���_����, StrInput, strOutput) = False Then Exit Function
            If strOutput = "" Then Exit Function
            strArr = Split(strOutput, "||")
            
            With g��������
                .���� = strArr(0)
                .���� = strArr(1)
                .����ǰ�ʻ���� = Val(strArr(2))
                .�����ʻ�֧����� = Val(strArr(3))
                .�Էѽ�� = Val(strArr(4))
                .���Ѻ��ʻ���� = Val(strArr(5))
                .����ʱ�� = strArr(6)
                .ǰ�˵��ݺ� = strArr(7)
                .���ĵ��ݺ� = strArr(8)
                .������ = strArr(9)
                .����Ա���� = strArr(10)
                .ǰ������ = strArr(11)
            End With
        End If
    End If
       
    '��д�����
    Call DebugTool("��д�����¼")
    

    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(����ǰ�ʻ����),�ۼ�ͳ�ﱨ��_IN(���Ѻ��ʻ����),סԺ����_IN(��),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(�Էѽ��),
    '   ����ͳ����_IN(��),ͳ�ﱨ�����_IN(��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(���ĵ��ݺ�),��ҳID_IN(��),��;����_IN,��ע_IN(ǰ�˵��ݺ�|������|����Ա����|ǰ������)
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�ϳ����� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL," & g��������.����ǰ�ʻ���� & "," & g��������.���Ѻ��ʻ���� & ",null,0,0,0," & _
            g�������_�ϳ�����.�����ܶ� & ",0," & g��������.�Էѽ�� & "," & _
          "0,0,0,0," & g��������.�����ʻ�֧����� & ",'" & _
            g��������.���ĵ��ݺ� & " ',NULL,NULL,'" & g��������.ǰ�˵��ݺ� & "|" & g��������.������ & "|" & g��������.����Ա���� & "|" & g��������.ǰ�˵��ݺ� & "')"
            
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------
    �������_�ϳ����� = True
    Exit Function

'Err������:
'
''��ڲ���:YBJGBH  PCHAR   ���ջ������
''        cZXDJH  PCHAR   ���ĵ��ݺ�(����ʱ����)
''        CPASSWORD   PCHAR   �ֿ��˿�����
''        CCZYXM  PCHAR   ����Ա����
'    strInput = g�������_�ϳ�����.��������
'    strInput = strInput & vbTab & g��������.���ĵ��ݺ�
'    strInput = strInput & vbTab & g�������_�ϳ�����.����
'    strInput = strInput & vbTab & gstrUserName
'
'    If ҵ������_�ϳ�����(���ѳ���_����, strInput, strOutPut) = False Then Exit Function
''����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
''   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
'    If strOutPut = "" Then Exit Function
'     strArr = Split(strOutPut, "||")
'
'    With g��������
'        .���� = strArr(0)
'        .���� = strArr(1)
'        .����ǰ�ʻ���� = Val(strArr(2))
'        .�����ʻ�֧����� = Val(strArr(3))
'        .�Էѽ�� = Val(strArr(4))
'        .���Ѻ��ʻ���� = Val(strArr(5))
'        .����ʱ�� = strArr(6)
'        .ǰ�˵��ݺ� = strArr(7)
'        .���ĵ��ݺ� = strArr(8)
'        .������ = strArr(9)
'        .����Ա���� = strArr(10)
'        .ǰ������ = strArr(11)
'    End With

    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function ����������_�ϳ�����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Dim intMouse As Integer
    Dim lng����ID  As Long
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    Dim lng����id1 As Long
    ����������_�ϳ����� = False
    
    '�����֤
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If ��ݱ�ʶ_�ϳ�����(2, lng����id1) = "" Then
        If lng����id1 = 0 Then
            Err.Raise 9000, gstrSysName, "�㲻�ǵ�ǰ�ֿ���!"
            Screen.MousePointer = intMouse
            Exit Function
        End If
    End If
    Screen.MousePointer = intMouse
    If ҵ������_�ϳ�����(��ʼ��_����, g�������_�ϳ�����.��������, strOutput) = False Then Exit Function
    
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rsTemp("����ID")
    
    
    
    gstrSQL = "Select * From ������ü�¼ " & _
        " Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
        
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������¼", lng����ID)
    g�������_�ϳ�����.�����ܶ� = 0
    With rs��ϸ
        Do While Not .EOF
                'д�ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            g�������_�ϳ�����.�����ܶ� = g�������_�ϳ�����.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    
    '����:
    gstrSQL = "Select ֧��˳��� from ���ս����¼ where ����=1 and ��¼id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ĵ��ݺ�"
    If rsTemp.EOF Then
        ShowMsgbox "�����ڽ����¼,���ܳ���!"
        Exit Function
    End If
    
    '��ڲ���:YBJGBH  PCHAR   ���ջ������
    '        cZXDJH  PCHAR   ���ĵ��ݺ�(����ʱ����)
    '        CPASSWORD   PCHAR   �ֿ��˿�����
    '        CCZYXM  PCHAR   ����Ա����
    StrInput = g�������_�ϳ�����.��������
    StrInput = StrInput & vbTab & Nvl(rsTemp!֧��˳���)
    StrInput = StrInput & vbTab & g�������_�ϳ�����.����
    StrInput = StrInput & vbTab & gstrUserName
    
    If ҵ������_�ϳ�����(���ѳ���_����, StrInput, strOutput) = False Then Exit Function
    '����:�����ʻ�������Ϣ(OK@$�����ʻ�������Ϣ)
    '   ��ʽ:����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
    If strOutput = "" Then Exit Function
     strArr = Split(strOutput, "||")
    
    With g��������
        .���� = strArr(0)
        .���� = strArr(1)
        .����ǰ�ʻ���� = Val(strArr(2))
        .�����ʻ�֧����� = Val(strArr(3))
        .�Էѽ�� = Val(strArr(4))
        .���Ѻ��ʻ���� = Val(strArr(5))
        .����ʱ�� = strArr(6)
        .ǰ�˵��ݺ� = strArr(7)
        .���ĵ��ݺ� = strArr(8)
        .������ = strArr(9)
        .����Ա���� = strArr(10)
        .ǰ������ = strArr(11)
    End With
    ����������_�ϳ����� = True
        
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(����ǰ�ʻ����),�ۼ�ͳ�ﱨ��_IN(���Ѻ��ʻ����),סԺ����_IN(��),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(�Էѽ��),
    '   ����ͳ����_IN(��),ͳ�ﱨ�����_IN(��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(���ĵ��ݺ�),��ҳID_IN(��),��;����_IN,��ע_IN(ǰ�˵��ݺ�|������|����Ա����|ǰ������)
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�ϳ����� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL," & -1 * g��������.����ǰ�ʻ���� & "," & -1 * g��������.���Ѻ��ʻ���� & ",null,0,0,0," & _
           -1 * g�������_�ϳ�����.�����ܶ� & ",0," & -1 * g��������.�Էѽ�� & "," & _
          "0,0,0,0," & -1 * g��������.�����ʻ�֧����� & ",'" & _
            g��������.���ĵ��ݺ� & " ',NULL,NULL,'" & g��������.ǰ�˵��ݺ� & "|" & g��������.������ & "|" & g��������.����Ա���� & "|" & g��������.ǰ�˵��ݺ� & "')"
            
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
    ����������_�ϳ����� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function Get��������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    '    ����ID||�籣���||����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��||�μӹ�������||��������||ְ�񼶱�||
    '   ְ�Ƽ���||��Ա����||��ؾ�ס��־||��λID||��λ����||����||����||ҽ��֤��||סԺ����||����ҽ�Ʊ�־||
    '   ����ҽ�Ʊ�־||����Ա��־||����ҽ�ƴ���״̬||����ҽ�ƴ���״̬||����Ա����״̬||����סԺ����||
    '   �����ѱ������||�ɷ�����||��ȡʱ��||סԺ��¼��||��Ժ���ڣ���ʽ��YYYY-MM-DD��||
    '   ��Ժ���||��Ժָ��||����||����||����
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String
    
    gstrSQL = "" & _
        "   Select  to_char(a.��Ժ����,'yyyy-mm-dd') as ��Ժ����,a.��Ժ����,b.���� as ����,c.���� as ����,d.���ֱ���,d.��Ժ���,a.��Ժ����" & _
        "   From ������ҳ a,���ű� b,���ű� c, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����)) AS ���ֱ���,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ���1, max(DECODE(a.��ϴ���,2,b.����,'')) AS ��Ժ���2,max(DECODE(a.��ϴ���,3,b.����,'')) AS ��Ժ���3, max(DECODE(a.��ϴ���,4,b.����||'-'||b.����,'')) AS ��Ժ���4 From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� in (1,2)  and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by ����id,��ҳid) D" & _
        " Where  " & _
        "        A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & "  and A.��Ժ����ID=b.id(+) and a.��Ժ����ID =c.id(+) " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) "
        
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��ҳ��Ϣ"
    With g�������_�ϳ�����
        'strInput = .strסԺ��Ϣ
        'strInput = strInput & "||" & Nvl(rsTemp!��Ժ����, "")
        StrInput = Nvl(rsTemp!��Ժ����, "")
        StrInput = StrInput & "||" & Nvl(rsTemp!���ֱ���)
        StrInput = StrInput & "||" & Nvl(rsTemp!��Ժ���1)     '���е���ֻ�ܴ����ƣ���Ϊ���⼲�����ж�
        StrInput = StrInput & "||" & Nvl(rsTemp!��Ժ����)      '��Ժָ��,Ŀǰû�д�
        StrInput = StrInput & "||" & Nvl(rsTemp!����)
        StrInput = StrInput & "||" & Nvl(rsTemp!��Ժ����)
        StrInput = StrInput & "||" & Nvl(rsTemp!����)
    End With
    Get�������� = StrInput
End Function

Public Function ��Ժ�Ǽ�_�ϳ�����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
  '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
    Dim strArr
    Dim strסԺ�� As String
    
    Err = 0: On Error GoTo errHand:
    
    '��ȡסԺ��
    DebugTool "������Ժ�Ǽǽӿ�"
    
    
   If InitInfor_�ϳ�����.�������� <> g�������_�ϳ�����.�������� Then
        '�����κ�����
        If gbln�Ѿ���ʼ = False And gbln������� Then
             If ҵ������_�ϳ�����(�����κ�����_����, g�������_�ϳ�����.��������, strOutput) = False Then
                  Exit Function
             End If
        End If
        
        If gbln������� Then
             '���κ�����
             If ҵ������_�ϳ�����(���κ�����_����, "", strOutput) = False Then
                  Exit Function
             End If
         End If
    End If
    
'    gstrSQL = "Select ҽ��סԺ��_ID.nextval  as סԺ��  From dual "
'    OpenRecordset_�ϳ����� rsTemp, "��ȡסԺ��"
'
    '    AZYH    PCHAR   סԺ��
    '    ARYZL   PCHAR   �α���Ա����Ժ����
    '    ABXJGBH PCHAR   �α���Ա���ڵ��籣�������
    '    ACZYXM  PCHAR   ����Ա����
    If lng��ҳID > 9 Then
        strסԺ�� = "9" & Lpad(lng��ҳID, 3, "0") & lng����ID
    Else
        strסԺ�� = Rpad(lng��ҳID, 4, "0") & lng����ID
    End If

    StrInput = g�������_�ϳ�����.ҽ��֤��
    StrInput = StrInput & vbTab & strסԺ��
    StrInput = StrInput & vbTab & Get��������(lng����ID, lng��ҳID)
    StrInput = StrInput & vbTab & g�������_�ϳ�����.��������
    StrInput = StrInput & vbTab & gstrUserName
    
    If ҵ������_�ϳ�����(��Ժ�Ǽ�_����, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ϳ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_�ϳ����� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�ϳ����� = False
End Function
Private Function Get���״���(ByVal intType As ҵ������_�ϳ�����, Optional bln������ As Boolean = False) As String
    '������û��
    Select Case intType
        Case ����籣����_����
            Get���״��� = IIf(bln������, "����籣����", "01")
        Case ����籣����_סԺ_����
            Get���״��� = IIf(bln������, "����籣����_סԺ_����", "27")
        
        Case ��òα���Ա����_����
            Get���״��� = IIf(bln������, "��òα���Ա����", "02")
        Case ��ȡ�ʻ����_����
                Get���״��� = IIf(bln������, "��ȡ�ʻ����", "03")
        Case ���κ�����_����
            Get���״��� = IIf(bln������, "���κ�����", "04")
        Case �����κ�����_����
            Get���״��� = IIf(bln������, "�����κ�����", "05")
        Case �Ͽ��κ�����_����
            Get���״��� = IIf(bln������, "�Ͽ��κ�����", "06")
        Case �����ʻ�����_����
            Get���״��� = IIf(bln������, "�����ʻ�����", "07")
        Case �����ʻ�����_���_����
            Get���״��� = IIf(bln������, "�����ʻ�����_���", "08")
        Case ���ѳ���_����
            Get���״��� = IIf(bln������, "���ѳ���", "09")
        Case �޸�����_����
            Get���״��� = IIf(bln������, "�޸�����", "10")
        Case ��ʼ��_����
            Get���״��� = IIf(bln������, "��ʼ��", "11")
        Case ���ؽ��׼�¼_����
            Get���״��� = IIf(bln������, "���ؽ��׼�¼", "12")
        Case �����Ա����_ҽ����_����
            Get���״��� = IIf(bln������, "�����Ա����_ҽ����_����", "13")
        Case �����Ա����_����_����
            Get���״��� = IIf(bln������, "�����Ա����_����_����", "14")
        Case ��Ժ�Ǽ�_����
            Get���״��� = IIf(bln������, "��Ժ�Ǽ�_����", "15")
        Case ȡ����Ժ�Ǽ�_����
            Get���״��� = IIf(bln������, "ȡ����Ժ�Ǽ�_����", "16")
        Case ��ȡ������¼��_����
            Get���״��� = IIf(bln������, "��ȡ������¼��_����", "17")
        Case ���Ӵ�������_����
            Get���״��� = IIf(bln������, "���Ӵ�������_����", "18")
        Case ������������_����
            Get���״��� = IIf(bln������, "������������_����", "19")
        Case ���Ӵ�����ϸ_����
            Get���״��� = IIf(bln������, "���Ӵ�����ϸ_����", "20")
        Case ��Ժ����_����
            Get���״��� = IIf(bln������, "��Ժ����_����", "21")
        Case ȡ����Ժ����_����
            Get���״��� = IIf(bln������, "ȡ����Ժ����_����", "22")
        Case ����סԺ�Ż�ȡ��¼��_����
            Get���״��� = IIf(bln������, "����סԺ�Ż�ȡ��¼��_����", "23")
        Case ��ӡ���㱨��_����
            Get���״��� = IIf(bln������, "��ӡ���㱨��_����", "24")
        Case סԺ���˿�������_����
            Get���״��� = IIf(bln������, "סԺ���˿�������_����", "25")
        Case ��ȡ��������_����
            Get���״��� = IIf(bln������, "��ȡ��������_����", "26")
        Case ����Ԥ����_����
            Get���״��� = IIf(bln������, "����Ԥ����_����", "28")
        Case Else
            Get���״��� = IIf(bln������, "����Ľ��״���", "-1")
    End Select
End Function
Public Function ҵ������_�ϳ�����(ByVal intType As ҵ������_�ϳ�����, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strOutput As String, strReturn As String
    Dim strInValue(0 To 20) As String
    
    Dim str���״��� As String
    Dim i As Integer
    Dim strArr
    
    str���״��� = Get���״���(intType, True)
    StrInput = strInputString
    DebugTool "����ҵ��������(ҵ�����ʹ���Ϊ:" & intType & " ҵ�����ƣ�" & str���״��� & ")" & vbCrLf & "        �������Ϊ:" & strInputString
    
    ҵ������_�ϳ����� = False
    If InitInfor_�ϳ�����.ģ������ Then
        '��ȡģ������
        Readģ������ intType, StrInput, strOutPutstring
         ҵ������_�ϳ����� = True
        Exit Function
    End If
    strArr = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    Err = 0
    On Error GoTo errHand:
    
    Select Case intType
        Case ����籣����_����
            strOutput = GetSBJGLB()
            
            If strOutput = "" Then
                MsgBox "��ȡ�籣����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Replace(strOutput, "OK@$", "")
        Case ����籣����_סԺ_����
            strOutput = GetSBJGLB1()
            
            If strOutput = "" Then
                MsgBox "����籣����_סԺ_����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Replace(strOutput, "OK@$", "")
        
        Case ��òα���Ա����_����
            strOutput = GETKZL()
            If strOutput = "" Then
                MsgBox "��òα���Ա����_����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case �����Ա����_ҽ����_����
            strOutput = GETRYJBZL(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "��òα���Ա����_����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        
        Case �����Ա����_����_����
            strOutput = GETICCARD()
            If strOutput = "" Then
                MsgBox "��òα���Ա����_����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)

        Case ��ȡ�ʻ����_����
            strOutput = GETZHYE(strInValue(0), strInValue(1))
            ''OK'+�м����+�����ʻ����
            If strOutput = "" Then
                MsgBox "��ȡ�ʻ����_ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        
        Case ���κ�����_����
            strOutput = CheckCon()
            If strOutput = "" Then
                MsgBox "���κ�����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case �����κ�����_����
            strOutput = RasDial(strInValue(0))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case �Ͽ��κ�����_����
            strOutput = DisDial()
            strOutput = Split(strOutput, Chr(0))(0)
            strOutput = ""
        Case �����ʻ�����_����
            strOutput = GRZHXF_CF(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
            strOutput = strArr(1)
        Case ����Ԥ����_����
            strOutput = GRZHXF_CFPRE(strInValue(0), strInValue(1), strInValue(2))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            'OK@$�����ʻ�֧�����||�Ը����
            strOutput = strArr(1)
        
        Case �����ʻ�����_���_����
            strOutput = GRZHXF_JE(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
            strOutput = strArr(1)
        Case ���ѳ���_����
            strOutput = XFCZ(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
            strOutput = strArr(1)
        Case �޸�����_����
            strOutput = CHANGPASSWORD(strInValue(0), strInValue(1), strInValue(2))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If Left(strArr(0), 2) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
           
        Case ��ʼ��_����
            strOutput = QDINIT(strInValue(0))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If Left(strArr(0), 2) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ���ؽ��׼�¼_����
        
            strOutput = DOWNJYJL(strInValue(0))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If Left(strArr(0), 2) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            
            strOutput = ""
        Case ��Ժ�Ǽ�_����
            strOutput = RYDJ(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4))
            If strOutput = "" Then
                MsgBox "��Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ȡ����Ժ�Ǽ�_����
            strOutput = ZYQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "ȡ����Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ��ȡ������¼��_����
            strOutput = GETNEWCFID()
            If strOutput = "" Then
                MsgBox "��ȡ������¼��ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""

        Case ���Ӵ�������_����
            '���:
            '    AZYH    PCHAR   סԺ��
            '    ACFID   PCHAR   ��������(���������ݿ��б�֤Ψһ)
            '    ACFMXID PCHAR   ��ϸ���(��һ�������б�֤Ψһ)
            '    ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
            '    AYS PCHAR   ҽ��
            '    AKS PCHAR   ����
            '    AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
            '    ASL PCHAR   ����(����Ϊ����)
            '    ADJ PCHAR   ����
            '    ASPM        --��Ʒ����ҽԺ��ҩƷ�����ش���
            strOutput = AddCFJL(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4), strInValue(5), strInValue(6), strInValue(7), strInValue(8), strInValue(9))
            If strOutput = "" Then
                MsgBox "���Ӵ�������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            
            strOutput = ""
            For i = 1 To UBound(strArr)
                strOutput = "||" & strArr(i)
            Next
            If strOutput <> "" Then
                strOutput = Mid(strOutput, 3)
            End If
        Case ������������_����
            strOutput = CFCS(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "������������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ���Ӵ�����ϸ_����
            strOutput = AddCFMX(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutput = "" Then
                MsgBox "���Ӵ�����ϸʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
            For i = 1 To UBound(strArr)
                strOutput = "||" & strArr(i)
            Next
            If strOutput <> "" Then
                strOutput = Mid(strOutput, 3)
            End If
        Case ��Ժ����_����
            strOutput = CYJS(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4), strInValue(5), strInValue(6), strInValue(7))
            If strOutput = "" Then
                MsgBox "��Ժ����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Mid(strOutput, 5)
        
        Case ȡ����Ժ����_����
            strOutput = CYJSQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "ȡ����Ժ����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ����סԺ�Ż�ȡ��¼��_����
            
            strOutput = GETZYIDBYZYBH(strInValue(0))
            If strOutput = "" Then
                MsgBox "����סԺ�Ż�ȡ��¼��ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case ��ӡ���㱨��_����
            strOutput = JSReport(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "��ӡ���㱨��ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case סԺ���˿�������_����
            strOutput = GETNEWRYZL(strInValue(0))
            If strOutput = "" Then
                MsgBox "����������Ա��������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""

        Case ��ȡ��������_����
             strOutput = GETJCXX(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "��ȡ��������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
    End Select
    strOutPutstring = strOutput
    ҵ������_�ϳ����� = True
    DebugTool "    �������Ϊ:" & strOutPutstring
     Exit Function
    
errHand:
    DebugTool "    �������Ϊ:" & strOutPutstring
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�ϳ�����(lng����ID As Long, lng��ҳID As Long) As Boolean
  '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    '���˺�:20040923���ӵ�
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
     Err = 0
    On Error GoTo errHand
    
    DebugTool "������Ժ�ǳ����ӿ�"
    
    ��Ժ�Ǽǳ���_�ϳ����� = False
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ����ã����ܳ�����Ժ�Ǽ�"
        Exit Function
    End If
    
    If lng��ҳID > 9 Then
        StrInput = "9" & Lpad(lng��ҳID, 3, "0") & lng����ID
    Else
        StrInput = Rpad(lng��ҳID, 4, "0") & lng����ID
    End If
    
    If ҵ������_�ϳ�����(ȡ����Ժ�Ǽ�_����, StrInput, strOutput) = False Then Exit Function

    '����ҽ���ʻ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ϳ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    DebugTool "ȡ���ɹ�"
    ��Ժ�Ǽǳ���_�ϳ����� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_�ϳ�����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0:    On Error GoTo errHand:
    ��Ժ�Ǽ�_�ϳ����� = False
    
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "��ǰ���˲�����δ����ã�������Ժ��������"
        Exit Function
    End If
    Call frm�����Ϣ_�Ĵ�.ShowME(lng����ID, lng��ҳID, 3, 3, TYPE_�ϳ�����, InitInfor_�ϳ�����.������뷽ʽ)
    '�ı䵱ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ϳ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_�ϳ����� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ��Ժ�Ǽǳ���_�ϳ�����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
  '��Ժ�Ǽǳ���
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
    ��Ժ�Ǽǳ���_�ϳ����� = False
    
    Err = 0: On Error GoTo errHand:
     
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "�ò����Ѿ���Ժ������,������ȡ����Ժ!"
        Exit Function
    End If
    
    gstrSQL = "ZL_������������Ϣ_DELETE(3," & lng����ID & "," & lng��ҳID & ")"
    ExecuteProcedure_�ϳ����� "ɾ���м�����ϼ�¼"
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ϳ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�ϳ����� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�ϳ�����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
  '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
 '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String
    
    Dim rs������ As New ADODB.Recordset, str�������1 As String, str�������2 As String, str�������3 As String
    
    Dim lng��ҳID As Long
    Dim dbl�����ܶ� As Double
    Dim strArr As Variant, strTmpArr As Variant
    Dim str���㷽ʽ  As String, strסԺ�� As String
    Dim obj���� As ��������
    סԺ����_�ϳ����� = False
        
 
    Err = 0: On Error GoTo errHand:
    Call DebugTool("����סԺ����")
    
    
    If g�������_�ϳ�����.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣"
        Exit Function
    End If
        
    gstrSQL = "Select ��ǰ״̬ From �����ʻ�  where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�жϵ�ǰ��סԺ״̬!"
    
    If Nvl(rsTemp!��ǰ״̬, 0) = 1 Then
        Err.Raise 9000, gstrSysName, "��ǰ���˻�������Ժ״̬,���Ժ���ٽ���!"
        Exit Function
    End If
    
    
    With g��������
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
        If IsNull(rsTemp("��ҳID")) = True Then
            Err.Raise 9000, gstrSysName, "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣"
            Exit Function
        End If
        lng��ҳID = rsTemp("��ҳID")
    End With
    
    If lng��ҳID > 9 Then
        strסԺ�� = "9" & Lpad(lng��ҳID, 3, "0") & lng����ID
    Else
        strסԺ�� = Rpad(lng��ҳID, 4, "0") & lng����ID
    End If
    
    gstrSQL = " " & _
          " Select sum(nvl(���ʽ��,0)) as ʵ�ս�� " & _
          " From סԺ���ü�¼ " & _
          " Where ��¼״̬<>0 and ����ID=" & lng����ID & " and  Nvl(���ӱ�־,0)<>9"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�ܷ���"
    
    dbl�����ܶ� = Nvl(rsTemp!ʵ�ս��, 0)
    If dbl�����ܶ� <> g�������_�ϳ�����.�����ܶ� Then
        Err.Raise 9000, gstrSysName, "����������ݵķ����ܶ��뱾�ν���ķ����ܶ�ȣ��������д������ʵ��ϴ���!"
        Exit Function
    End If
    
    gstrSQL = "" & _
        "   Select C.סԺ��,C.��ǰ����id,A.��Ժ���� ,c.סԺ��,to_char(A.ȷ������,'yyyyMMdd') as ȷ������,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����ʱ��," & _
        "           to_char(A.��Ժ����,'yyyyMMdd') ��Ժ����, A.��Ժ��ʽ,to_char(a.��Ժ����,'yyyy-mm-dd') as ��Ժ���� ,a.��Ժ����,H.���� as ��Ժ����," & _
        "           g.�������,G.��Ժ���1,G.��Ժ���2,G.��Ժ���3,G.��Ժ���4" & _
        " From ������ҳ A,���ű� B,������Ϣ C,���ű� H, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,��Ժ���,'')) as �������,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ���1,max(DECODE(a.��ϴ���,2,b.����,'')) AS ��Ժ���2,max(DECODE(a.��ϴ���,3,b.����,'')) AS ��Ժ���3,max(DECODE(a.��ϴ���,4,b.����,'')) AS ��Ժ���4 From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� = 3 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by ����id,��ҳid)   G" & _
        " Where A.����id=C.����id and C.����id=" & lng����ID & _
        "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID and A.��Ժ����ID=H.id(+) " & _
        "       and A.��ҳid=G.��ҳid(+) and a.����id=G.����id(+) " & _
        ""
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ��Ϣ"
    
    gstrSQL = "select * from ������������Ϣ where ����=3 and ����ID= " & lng����ID & " And ��ҳID=" & lng��ҳID
    zlDatabase.OpenRecordset rs������, gstrSQL, "��ȡ�����Ժ���"
    Do Until rs������.EOF
        Select Case rs������("��ϴ���")
            Case "1"
                str�������1 = IIf(IsNull(rs������("������")), "", rs������("������"))
            Case "2"
                str�������2 = IIf(IsNull(rs������("������")), "", rs������("������"))
            Case "3"
                str�������3 = IIf(IsNull(rs������("������")), "", rs������("������"))
        End Select
        rs������.MoveNext
    Loop
    
    '�ٴν���
    StrInput = strסԺ��
    StrInput = StrInput & vbTab & Nvl(rsTemp!��Ժ���1) & vbTab & Nvl(str�������1, "��") & vbTab & Nvl(str�������2, "��") & vbTab & Nvl(str�������3, "��")
    StrInput = StrInput & vbTab & Get�������_����(lng����ID, lng��ҳID)
    StrInput = StrInput & vbTab & Nvl(rsTemp!��Ժ����)
    StrInput = StrInput & vbTab & "1"
    If ҵ������_�ϳ�����(��Ժ����_����, StrInput, strOutput) = False Then Exit Function
    '������(20060920):����޸ĺ�ķ���ֵ:
    '   סԺ���ý�����:
    '       �𸶽��||�˶��ɷѻ���||�����ⶥ���||����ⶥ���||����סԺ����||�����ۼ�סԺ����||�������ۼ�סԺ����||��������״̬
    '       �������״̬||����Ա����״̬||������������||���䱨������||����Ա��������||�������||����ҩƷ��||�������Ʒ�
    '       ��������||�������||����ҩƷ��||�������Ʒ�||����������||����ҩƷ�����Ը�||�������������Ը�||�������
    '       ����ҩƷ��||�������Ʒ�||���������ʩ��||����ҽ�Ʊ���||����ҽ�Ʊ���||����Աҽ�Ʊ���||ҽ�Ʊ����ܼ�||�Ը��ϼ�
    '   �����ֶ���ϸ(����)
    '       סԺ��������||���εı�������||�𸶽��||ȫ�Ը�ҩƷ��||ȫ�Ը����Ʒ�||ȫ�Ը�������ʩ��||����ҩƷ��||�������Ʒ�
    '       �Ը�����||�Ը����||�Ը�С��||���α������
    strArr = Split(strOutput, "@$")
    strTmpArr = Split(strArr(0), "||")
'    With obj����
'        .����������� = 0
'        .���䱨����� = 0
'        .����Ա������� = 0
'    End With
    
'    With obj����
'        .����������� = Val(strTmpArr(27))
'        .���䱨����� = Val(strTmpArr(28))
'        .����Ա������� = Val(strTmpArr(29))
'    End With
    
    gcnOracle_�ϳ�����.BeginTrans

    If InsertIntoҽ�������¼(strArr, lng����ID) = False Then Exit Function
    
    
    '��д�����
    Call DebugTool("��д�����¼")
    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(�����ѱ������),סԺ����_IN(��ҳID),����(�𸶽��),�ⶥ��_IN(�����ⶥ���),ʵ������_IN(������������),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(���䱨������),�����Ը����_IN(����Ա��������),
    '   ����ͳ����_IN(�����������),ͳ�ﱨ�����_IN(���䱨�����),���Ը����_IN(����Ա�������),�����Ը����_IN(��),�����ʻ�֧��_IN(),"
    '   ֧��˳���_IN(סԺ��),��ҳID_IN(��ҳID),��;����_IN,��ע_IN(����)
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
   '����ҽ�ƴ���״̬||�𸶽��||�����ⶥ���||������������||�����ѱ������||�����������||���䱨�����||����Ա�������||
   '����ҽ�ƴ���״̬||����Ա����״̬||���䱨������||����Ա��������||����סԺ����||�������||����ҩƷ��||�������Ʒ�||��������||�������||����ҩƷ��||�������Ʒ�||����������||�����Ը�||�������||����ҩƷ��||�������Ʒ�||��������||�����ϼ�||����֧��
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�ϳ����� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL," & Val(strTmpArr(6)) & "," & lng��ҳID & "," & Val(strTmpArr(0)) & "," & Val(strTmpArr(2)) & "," & Val(strTmpArr(0)) & "," & _
            dbl�����ܶ� & "," & Val(strTmpArr(23)) & "," & Val(strTmpArr(21)) + Val(strTmpArr(22)) & "," & _
            Val(strTmpArr(27)) & "," & Val(strTmpArr(28)) & "," & Val(strTmpArr(29)) & ",0,0,'" & _
            strסԺ�� & "'," & lng��ҳID & ",NULL,'" & g�������_�ϳ�����.�籣���� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
     gcnOracle_�ϳ�����.CommitTrans

      
    סԺ����_�ϳ����� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function סԺ�������_�ϳ�����(lng����ID As Long) As Boolean
     '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim rs�����¼ As New ADODB.Recordset
    
    Dim StrInput As String, strOutput  As String
    Dim lng����ID As Long, strסԺ�� As String
    Dim strArr
    Dim lng����ID As Long, intMouse As Integer
    Err = 0: On Error GoTo errHand:
    
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", TYPE_�ϳ�����, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    g�������_�ϳ�����.����ID = Nvl(rsTemp!����ID, 0)
    
    
    gstrSQL = "select * from ҽ�������¼ where ����=2  and ����ID=" & lng����ID
    Call OpenRecordset_�ϳ�����(rs�����¼, "�������")
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    lng����ID = g�������_�ϳ�����.����ID
    
    Screen.MousePointer = 1
    If ��ݱ�ʶ_�ϳ�����(88, g�������_�ϳ�����.����ID) = "" Then
        Screen.MousePointer = intMouse
        סԺ�������_�ϳ����� = False
        Exit Function
    End If
    Screen.MousePointer = intMouse
    If lng����ID <> g�������_�ϳ�����.����ID Then
        Err.Raise 9000, gstrSysName, "���ǵ�ǰҪ��������Ĳ���!"
        Exit Function
    End If
    
    
    '�жϲ��˵�סԺ���������Ƿ��������ϡ��жϱ�׼�Ǽ�鲡�����µ�סԺ��¼������У��Ͳ��ܽ�����
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    strסԺ�� = Nvl(rsTemp("֧��˳���"))
    
    StrInput = strסԺ��
    If ҵ������_�ϳ�����(ȡ����Ժ����_����, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    '��������
    '    ����_IN     IN ҽ�������¼.����%TYPE,
    '    ����ID_IN   IN ҽ�������¼.����ID%TYPE,
    '    ����ID_IN   IN ҽ�������¼.����ID%TYPE)
    
    gcnOracle_�ϳ�����.BeginTrans
    gstrSQL = "ZL_ҽ�������¼_����("
    gstrSQL = gstrSQL & "2"
    gstrSQL = gstrSQL & "," & lng����ID
    gstrSQL = gstrSQL & "," & lng����ID & ")"
    ExecuteProcedure_�ϳ����� "��������¼"
    
 
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(�����ѱ������),סԺ����_IN(��ҳID),����(�𸶽��),�ⶥ��_IN(�����ⶥ���),ʵ������_IN(������������),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(���䱨������),�����Ը����_IN(����Ա��������),
    '   ����ͳ����_IN(�����������),ͳ�ﱨ�����_IN(���䱨�����),���Ը����_IN(����Ա�������),�����Ը����_IN(��),�����ʻ�֧��_IN(),"
    '   ֧��˳���_IN(סԺ��),��ҳID_IN(��ҳID),��;����_IN,��ע_IN(����)
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
    '---------------------------------------------------------------------------------------------
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�ϳ����� & "," & rsTemp("����ID") & "," & Year(zlDatabase.Currentdate) & "," & _
        "NULL,NULL,NULL," & Nvl(rsTemp("�ۼ�ͳ�ﱨ��"), 0) * -1 & "," & Nvl(rsTemp!��ҳID, 0) & "," & Nvl(rsTemp("����"), 0) * -1 & "," & Nvl(rsTemp("�ⶥ��"), 0) * -1 & "," & Nvl(rsTemp("ʵ������"), 0) * -1 & "," & _
        Nvl(rsTemp("�������ý��"), 0) * -1 & "," & Nvl(rsTemp("ȫ�Ը����"), 0) * -1 & "," & Nvl(rsTemp("�����Ը����"), 0) * -1 & "," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & "," & Nvl(rsTemp!���Ը����, 0) * -1 & ",0,NULL,'" & _
        strסԺ�� & "'," & Nvl(rsTemp!��ҳID, 0) & ",NULL,'" & Nvl(rsTemp!��ע) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ�������¼")
    gcnOracle_�ϳ�����.CommitTrans
    
    סԺ�������_�ϳ����� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function �����Ǽ�_�ϳ�����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------
   '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim str������¼�� As String, strժҪ As String
    Dim strArr
    
    
    �����Ǽ�_�ϳ����� = False
    
    
   '�������ŵ��ݵķ�����ϸ
  gstrSQL = "Select A.ID,a.��ʶ�� סԺ��,a.���,a.NO,A.��¼����,A.��¼״̬,A.����ID,A.��ҳID,to_char(A.����ʱ��,'yyyy-mm-dd') as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,M.����,Q.���� as ��������,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,�����ʻ� M,���ű� Q,������Ϣ J" & _
              "  where A.NO=[1] and A.��¼����=[2] and A.��¼״̬ = [3] And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
              "        and A.�շ�ϸĿID=B.ID and A.����ID=J.����ID  and A.��ҳID=J.סԺ���� And M.����=[4]" & _
              "        and a.����id=m.����id and a.��������id=q.id(+)" & _
              "  Order by A.����ID,A.��¼����,a.��¼״̬,A.NO,A.���,A.����ʱ��"
        
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ�ϴ�", str���ݺ�, lng��¼����, lng��¼״̬, TYPE_�ϳ�����)
    If InitInfor_�ϳ�����.��ϸʱʵ�ϴ� = False Then
        �����Ǽ�_�ϳ����� = True
        Exit Function
    End If
    
    Err = 0
    On Error GoTo errHand:
    gcnOracle.BeginTrans
    
    Dim lng������ As Long
    Dim strסԺ�� As String
    
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            gstrSQL = "Select * From ҽ��֧����Ŀ where ����=[1] and ����=[2] and �շ�ϸĿid=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ��ҽ��֧����Ŀ", TYPE_�ϳ�����, CLng(Nvl(!����, 0)), CLng(Nvl(!�շ�ϸĿID, 0)))
            If rsTemp.EOF Then
                ShowMsgbox "ע�⣺" & vbCrLf & "   �շ�ϸĿΪ:[" & Nvl(!����) & "]" & Nvl(!����) & " ��δ����ҽ������!"
            End If
            
            '���Ӵ�����ϸ
            '    AZYH    PCHAR   סԺ��
            '    ACFID   PCHAR   ��������(���������ݿ��б�֤Ψһ)
            '    ACFMXID PCHAR   ��ϸ���(��һ�������б�֤Ψһ)
            '    ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
            '    AYS PCHAR   ҽ��
            '    AKS PCHAR   ����
            '    AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
            '    ASL PCHAR   ����(����Ϊ����)
            '    ADJ PCHAR   ����
            '    ASPM        --��Ʒ����ҽԺ��ҩƷ�����ش���
            If lng����ID <> Nvl(!����ID, 0) Then
                lng����ID = Nvl(!����ID, 0)
                lng������ = !ID
            End If
            lng��ҳID = Nvl(!��ҳID, 0)
            
            If lng��ҳID > 9 Then
                strסԺ�� = "9" & Lpad(lng��ҳID, 3, "0") & lng����ID
            Else
                strסԺ�� = Rpad(lng��ҳID, 4, "0") & lng����ID
            End If
    
            StrInput = strסԺ��
            StrInput = StrInput & vbTab & lng������ 'Val(Mid(Nvl(!�Ǽ�ʱ��), 3, 4)) & Val(Substr(Nvl(!�Ǽ�ʱ��, "05"), 3, 2)) & Mid(Nvl(!no), 2) & Lpad(!��¼����, 3, "0") & Lpad(!��¼״̬, 3, "0") '����Ƕಡ�˵������Կ��ǽ�����id��������
            StrInput = StrInput & vbTab & Nvl(!ID)
            StrInput = StrInput & vbTab & Nvl(!�Ǽ�ʱ��)
            StrInput = StrInput & vbTab & Nvl(!ҽ��)
            StrInput = StrInput & vbTab & Nvl(!��������)
            
            If rsTemp.EOF Then
                StrInput = StrInput & vbTab & ""
            Else
                StrInput = StrInput & vbTab & Nvl(rsTemp!��Ŀ����)
            End If
            '��������2005-07-12 ����С�����ȴ���2λ����Ŀ�����ۺ��������⴦��
            If Round(Nvl(!�۸�) * 100) <> Nvl(!�۸� * 100) Then
                If !ʵ�ս�� < 0 Then
                   StrInput = StrInput & vbTab & "-1"
                Else
                   StrInput = StrInput & vbTab & "1"
                End If
                StrInput = StrInput & vbTab & Abs(!ʵ�ս��)
            Else
                StrInput = StrInput & vbTab & Nvl(!����)
                StrInput = StrInput & vbTab & Nvl(!�۸�)
            End If
            
'            strInput = strInput & vbTab & Nvl(!����)
'            strInput = strInput & vbTab & Nvl(!�۸�)
            StrInput = StrInput & vbTab & Nvl(!����)
            If rsTemp.EOF Then
                '���������ϴ�
            Else
                If ҵ������_�ϳ�����(���Ӵ�������_����, StrInput, strOutput) = False Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
                strOutput = Replace(strOutput, "@$", "||")
                '������ϸ��¼��@$�Էѱ���@$�Էѽ��
                'ժҪ����ֵ:������||�Էѱ���||�Էѽ��||סԺ��
                strժҪ = lng������ & "||" & strOutput & "||" & strסԺ��
                '�����ϴ���־
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strժҪ & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            End If
            .MoveNext
        Loop
    End With
    gcnOracle.CommitTrans
    �����Ǽ�_�ϳ����� = True
    Exit Function
errHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
End Function
Private Function Get�������_����(lng����ID As Long, lng��ҳID As Long) As String
    '����:��ȡ���������ʶ

    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.��Ժ��� " & _
             " From ������ A,��������Ŀ¼ B " & _
             " Where A.����ID=[1] And A.����ID=B.ID(+) And A.��ҳID=[2]" & _
             "       And A.������� in (2,3)" & _
             " Order by A.������� Desc"
     
    Set rsInNote = zlDatabase.OpenSQLRecord(strTmp, "ҽ���ӿ�", lng����ID, lng��ҳID)
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = Nvl(rsInNote!��Ժ���)
    End If
    If strTmp = "" Then
        strTmp = "����"
        
    End If
   ' strTmp = Decode(strTmp, "����", "1", "��ת", "2", "δ��", "3", "����", "4", "����", "9", "1")
    Get�������_���� = strTmp
   Call WriteDebugInfor_����("Get�������_����", lng����ID)
End Function


Private Function Readģ������(ByVal intҵ������ As ҵ������_�ϳ�����, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ���ù��ܶ�ȡģ������,�Ա����
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim STRNAME As String
    
    If intҵ������ = ��ȡ�������� Then
        strFile = App.Path & "\������.txt"
    Else
        strFile = App.Path & "\ģ���ύ��.txt"
    End If
    
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    STRNAME = Get���״���(intҵ������, True)
    
    Dim blnStart As Boolean
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If intҵ������ = ��ȡ�������� Then
                    strArr = Split(strText, vbTab & "|")
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab & "|"
                            End If
                            strArr = Split(strText, vbTab & "|")
                            
                            If Val(strArr(0)) = 1 Then
                                str = strArr(1)
                                Exit Do
                            End If
                        Else
                             If "<" & STRNAME & ">" = strText Then
                                 blnStart = True
                             End If
                        End If
                        If "</" & STRNAME & ">" = strText Then
                            Exit Do
                        End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
'    If InStr(1, strOutPutstring, "@$") <> 0 Then
'        strOutPutstring = Split(strOutPutstring, "@$")(1)
'    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_�ϳ�����(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_�ϳ�����, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub
Private Function ����סԺ��ϸ��¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '���������ϸ��¼
    Dim cnTemp As New ADODB.Connection
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput  As String, strOutput As String
    Dim strArr
    Dim strסԺ�� As String, str������¼�� As String
    Dim strNO  As String
    Dim strSQL As String, strTmp As String
    Dim strժҪ As String
    Dim i As Integer
    
    Err = 0
    On Error GoTo errHand:
      
    
    Call DebugTool("��������")
    Set cnTemp = GetNewConnection
    Call DebugTool("�����ӳɹ�����ʼ�����ϸ���ݵĺϷ��ԡ�")
      
      
    ����סԺ��ϸ��¼ = False
    
    gstrSQL = "Select A.ID,a.��ʶ�� סԺ��,A.ժҪ,a.���,a.NO,A.��¼����,A.��¼״̬,A.����ID,A.��ҳID,to_char(A.����ʱ��,'yyyy-mm-dd') as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
                "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
                "         ,M.����,Q.���� as ��������,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,Nvl(A.�Ƿ��ϴ�,0) as �ϴ���־ " & _
                "  From סԺ���ü�¼ A,�շ�ϸĿ B,�����ʻ� M,���ű� Q,������Ϣ J" & _
                "  where Nvl(���ӱ�־,0)<>9 and nvl(a.ʵ�ս��,0)<>0 " & _
                "        and A.�շ�ϸĿID=B.ID and A.����ID=J.����ID   and A.����ID=[1] and A.��ҳID=[2] And M.����=[3]" & _
                "        and a.����id=m.����id and a.��������id=q.id(+)" & _
                "  Order by A.����ID,A.��¼����,a.��¼״̬,A.NO,A.���,A.����ʱ��"
                
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID, lng��ҳID, TYPE_�ϳ�����)

   With rs��ϸ
        i = 1
        strNO = ""
        Do While Not .EOF
            g�ɶ�������Ϣ = "���ڴ��������ϸ�����Ժ" & vbCrLf & _
                            "��" & i & "����ϸ����" & rs��ϸ.RecordCount & "����ϸ��"
            frm�ɶ�������ʾ.Show 1
            
            '������(2006-03-06):����ͬһ����δ��ȫ�ϴ������,��Ҫȡ����ǰ�Ĵ�����¼��
            strTmp = Nvl(!��¼����, 0) & "_" & Nvl(!��¼״̬, 0) & "_" & Nvl(!NO, 0)
            If strNO <> strTmp Then
                strNO = strTmp
                str������¼�� = Split(Nvl(!ժҪ) & "||||", "||")(0)
                If str������¼�� = "" Then
                   str������¼�� = Nvl(!ID)
                End If
            Else
                If str������¼�� = "" Then
                   str������¼�� = Nvl(!ID)
                End If
            End If
            
            If !�ϴ���־ = 0 Then
                gstrSQL = "Select * From ҽ��֧����Ŀ where ����=[1] and ����=[2] and �շ�ϸĿid=[3]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ��ҽ��֧����Ŀ", TYPE_�ϳ�����, CLng(Nvl(!����, 0)), CLng(Nvl(!�շ�ϸĿID, 0)))
                If rsTemp.EOF Then
                    ShowMsgbox "ע�⣺" & vbCrLf & "   �շ�ϸĿΪ:[" & Nvl(!����) & "]" & Nvl(!����) & " ��δ����ҽ������,����������!"
                    Exit Function
                End If
 
               
                '���Ӵ�����ϸ
                '    AZYH    PCHAR   סԺ��
                '    ACFID   PCHAR   ��������(���������ݿ��б�֤Ψһ)
                '    ACFMXID PCHAR   ��ϸ���(���������ݿ��б�֤Ψһ)
                '    ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
                '    AYS PCHAR   ҽ��
                '    AKS PCHAR   ����
                '    AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
                '    ASL PCHAR   ����(����Ϊ����)
                '    ADJ PCHAR   ����
                '    ASPM         ��Ʒ��
                
                If lng��ҳID > 9 Then
                    strסԺ�� = "9" & Lpad(lng��ҳID, 3, "0") & lng����ID
                Else
                    strסԺ�� = Rpad(lng��ҳID, 4, "0") & lng����ID
                End If
                
                StrInput = strסԺ��
                StrInput = StrInput & vbTab & str������¼�� ' Val(Mid(Nvl(!�Ǽ�ʱ��), 3, 4)) & Val(Substr(Nvl(!�Ǽ�ʱ��, "05"), 3, 2)) & Mid(Nvl(!no), 2) & Lpad(!��¼����, 3, "0") & Lpad(!��¼״̬, 3, "0") '����Ƕಡ�˵������Կ��ǽ�����id��������
                StrInput = StrInput & vbTab & Nvl(!ID)
                StrInput = StrInput & vbTab & Nvl(!�Ǽ�ʱ��)
                StrInput = StrInput & vbTab & Nvl(!ҽ��)
                StrInput = StrInput & vbTab & Nvl(!��������)
                If rsTemp.EOF Then
                    StrInput = StrInput & vbTab & ""
                Else
                    StrInput = StrInput & vbTab & Nvl(rsTemp!��Ŀ����)
                End If
                '������:2005-07-06 ������۾��ȳ���2λС������������1�����۴�ʵ�ս�
                If Round(Nvl(!�۸�) * 100) = Nvl(!�۸�) * 100 Then
                   StrInput = StrInput & vbTab & Nvl(!����)
                   StrInput = StrInput & vbTab & Nvl(!�۸�)
                Else
                   StrInput = StrInput & vbTab & "1"
                   StrInput = StrInput & vbTab & Nvl(!ʵ�ս��)
                End If
                StrInput = StrInput & vbTab & Nvl(!����)
                
                If ҵ������_�ϳ�����(���Ӵ�������_����, StrInput, strOutput) = False Then
                    Exit Function
                End If
                strOutput = Replace(strOutput, "@$", "||")
                '�Էѱ���@$�Էѽ��
                'ժҪ����ֵ:������||�Էѱ���||�Էѽ��||סԺ��
                strժҪ = str������¼�� & "||" & strOutput & "||" & strסԺ��
                '�����ϴ���־
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strժҪ & "')"
                cnTemp.Execute gstrSQL, , adCmdStoredProc
            End If
            i = i + 1
            .MoveNext
        Loop
    End With
    ����סԺ��ϸ��¼ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�ϳ�����(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    'rsExse:�ַ���
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng��ҳID As Long
    Dim StrInput As String, strOutput   As String
    Dim strArr As Variant
    Dim strסԺ�� As String, str���㷽ʽ As String
    Dim lng����id1 As Long
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo errHand:
    g�������_�ϳ�����.����ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    intMouse = Screen.MousePointer
    
    gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    
    If IsNull(rsTemp("��ҳID")) = True Then
        MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ҳID = rsTemp("��ҳID")
    
'    If bln���ʴ� Then
'        Screen.MousePointer = 1
'        If ��ݱ�ʶ_�ϳ�����(4, lng����id1) = "" Then
'            Screen.MousePointer = intMouse
'            סԺ�������_�ϳ����� = ""
'            Exit Function
'        End If
'        Screen.MousePointer = intMouse
'        If lng����id <> lng����id1 Then
'            ShowMsgbox "���ǵ�ǰҪ����Ĳ���!"
'            Exit Function
'        End If
'    End If
    
    gstrSQL = "Select b.סԺ��,a.���� From �����ʻ� a,������Ϣ b  where a.����id=b.����id and a.����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ��"
    If rsTemp.EOF Then
        ShowMsgbox "�ò��˲���ҽ������!"
        Exit Function
    End If
    
    If lng��ҳID > 9 Then
        strסԺ�� = "9" & Lpad(lng��ҳID, 3, "0") & lng����ID
    Else
        strסԺ�� = Rpad(lng��ҳID, 4, "0") & lng����ID
    End If

    g�������_�ϳ�����.�籣���� = Nvl(rsTemp!����, 0)
    
    
    Screen.MousePointer = vbHourglass
   
    With rsExse
        g�������_�ϳ�����.�����ܶ� = 0
        Do While Not .EOF
            g�������_�ϳ�����.�����ܶ� = g�������_�ϳ�����.�����ܶ� + Nvl(!���, 0)
            .MoveNext
        Loop
    End With
     
    If ����סԺ��ϸ��¼(lng����ID, lng��ҳID) = False Then Exit Function
    
    'AZYH    PCHAR   סԺ��
    'ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
    'ZLXG    PCHAR   ����Ч��
    'CYZD    PCHAR   ��Ժ���1+�м����+ ��Ժ���2+�м����+ ��Ժ���3+�м����+��Ժ���4
    'CYRQ    PCHAR   ��Ժ���ڣ�YYYY-MM-DD��

    StrInput = strסԺ��
    StrInput = StrInput & vbTab & "Ԥ����"
    StrInput = StrInput & vbTab & ""
    StrInput = StrInput & vbTab & ""
    StrInput = StrInput & vbTab & ""
    StrInput = StrInput & vbTab & "Ԥ����"
    StrInput = StrInput & vbTab & "2000-01-01"
    StrInput = StrInput & vbTab & "0"
    
    If ҵ������_�ϳ�����(��Ժ����_����, StrInput, strOutput) = False Then Exit Function
    '����:OK@$סԺ���ý�����@$�����ֶ���ϸ
    '   ˵��:
    '     סԺ���ý�����:
    '       ����ҽ�ƴ���״̬||�𸶽��||�����ⶥ���||������������||�����ѱ������||�����������||���䱨�����||����Ա�������||
    '       ����ҽ�ƴ���״̬||����Ա����״̬||���䱨������||����Ա��������||����סԺ����||�������||����ҩƷ��||�������Ʒ�
    '       ��������||�������||����ҩƷ��||�������Ʒ�||����������||�����Ը�||�������||����ҩƷ��||�������Ʒ�||��������||�����ϼ�||����֧��
    '     �����ֶ���ϸ(����):����||����||����ʼ���||����ֹ���||���λ���||���α�������||���α������||�����Ը����@$.......
    '������(20060920):����޸ĺ�ķ���ֵ:
    '   סԺ���ý�����:
    '       �𸶽��||�˶��ɷѻ���||�����ⶥ���||����ⶥ���||����סԺ����||�����ۼ�סԺ����||�������ۼ�סԺ����||��������״̬
    '       �������״̬||����Ա����״̬||������������||���䱨������||����Ա��������||�������||����ҩƷ��||�������Ʒ�
    '       ��������||�������||����ҩƷ��||�������Ʒ�||����������||����ҩƷ�����Ը�||�������������Ը�||�������
    '       ����ҩƷ��||�������Ʒ�||���������ʩ��||����ҽ�Ʊ���||����ҽ�Ʊ���||����Աҽ�Ʊ���||ҽ�Ʊ����ܼ�||�Ը��ϼ�
    '   �����ֶ���ϸ(����)
    '       סԺ��������||���εı�������||�𸶽��||ȫ�Ը�ҩƷ��||ȫ�Ը����Ʒ�||ȫ�Ը�������ʩ��||����ҩƷ��||�������Ʒ�
    '       �Ը�����||�Ը����||�Ը�С��||���α������

    strArr = Split(strOutput, "||")
    With g��������
        .����������� = Val(strArr(27))
        .���䱨����� = Val(strArr(28))
        .����Ա������� = Val(strArr(29))
    End With
     If Format(strArr(4), "####0.00;-####0.00;0;0") <> Format(g�������_�ϳ�����.�����ܶ�, "####0.00;-####0.00;0;0") Then
        ShowMsgbox "�������ݲ���!" & vbCrLf & "ҽ�����ķ����ܶ�:" & Format(strArr(4), "####0.00;-####0.00;0;0") & vbCrLf & " ҽԺ��Ϊ:" & Format(g�������_�ϳ�����.�����ܶ�, "####0.00;-####0.00;0;0")
        If InitInfor_�ϳ�����.���ݲ��Ȳ��ɽ��� Then
            Exit Function
        End If
    End If
    
    str���㷽ʽ = "����ҽ�Ʊ���;" & g��������.����������� & ";0"
    str���㷽ʽ = str���㷽ʽ & "|���䱨��;" & g��������.���䱨����� & ";0"
    str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & g��������.����Ա������� & ";0"
    סԺ�������_�ϳ����� = str���㷽ʽ
    g�������_�ϳ�����.����ID = lng����ID   '��ʾ�ò����Ѿ��������������
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function ҽ������_�ϳ�����(ByVal lng���� As Long, ByVal lngҽ������ As Integer) As Boolean
    ҽ������_�ϳ����� = frmSet�ϳ�����.��������
End Function
Public Sub ExecuteProcedure_�ϳ�����(ByVal strCaption As String)
'���ܣ�ִ��SQL���
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_�ϳ�����.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Private Function ������ϸд��(ByVal rs��ϸ As ADODB.Recordset, Optional ByVal bln���� As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ���ϸ��¼
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim str��ϸ As String
    
    Dim strArr
    
    ������ϸд�� = False
    g�������_�ϳ�����.�����ܶ� = 0
    
    Err = 0:    On Error GoTo errHand:
    'Ȼ����봦����ϸ
    With rs��ϸ
        If .RecordCount = 0 Then
            ShowMsgbox "��������ص���ϸ���ü�¼!"
            Exit Function
        End If
        'YBJGBH  PCHAR   ���ջ������
        'CFH PCHAR   ������
        'CFMXDATA    PCHAR   ������ϸ����    ��ʽ˵��������1(ҽ��ҩƷ���+�м����+����+�м��������+)+�м����+
        'CPASSWORD   PCHAR   �ֿ��˿�����
        'CCZYXM  PCHAR   ����Ա����
        StrInput = g�������_�ϳ�����.��������
        StrInput = StrInput & vbTab & Nvl(!NO)
        
        Do While Not rs��ϸ.EOF
            gstrSQL = "Select * From ҽ��֧����Ŀ where ����=[1] and ����=[2] and �շ�ϸĿid=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ��ҽ��֧����Ŀ", TYPE_�ϳ�����, g�������_�ϳ�����.�籣����, CLng(Nvl(!�շ�ϸĿID, 0)))
            If Val(Nvl(rs��ϸ("ʵ�ս��"), 0)) <> 0 Then
                If rsTemp.EOF Then
                    str��ϸ = str��ϸ & "@$" & ""
                Else
                    str��ϸ = str��ϸ & "@$" & Nvl(rsTemp!��Ŀ����)
                End If
                '��������2005-07-27 ����С�����ȴ���2λ����Ŀ�����ۺ��������⴦��
                If Round(Nvl(!����) * 100) <> Nvl(!���� * 100) Then
                    str��ϸ = str��ϸ & "||" & Abs(!ʵ�ս��)
                    If !ʵ�ս�� < 0 Then
                       str��ϸ = str��ϸ & "||" & "-1"
                    Else
                       str��ϸ = str��ϸ & "||" & "1"
                    End If
                Else
                    str��ϸ = str��ϸ & "||" & Nvl(!����, 0)
                    str��ϸ = str��ϸ & "||" & Nvl(!����, 0)
                End If
'
'                str��ϸ = str��ϸ & "||" & Nvl(!����, 0)
'                str��ϸ = str��ϸ & "||" & Nvl(!����, 0)
                
            End If
            g�������_�ϳ�����.�����ܶ� = g�������_�ϳ�����.�����ܶ� + Nvl(!ʵ�ս��, 0)
            
            'д�ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            
            rs��ϸ.MoveNext
        Loop
    End With
    If g�������_�ϳ�����.ֱ���� = False Then
        str��ϸ = Mid(str��ϸ, 3)
        StrInput = StrInput & vbTab & str��ϸ
        StrInput = StrInput & vbTab & g�������_�ϳ�����.����
        StrInput = StrInput & vbTab & gstrUserName
        
        If ҵ������_�ϳ�����(�����ʻ�����_����, StrInput, strOutput) = False Then Exit Function
        If strOutput = "" Then Exit Function
        strArr = Split(strOutput, "||")
        
        With g��������
            .���� = strArr(0)
            .���� = strArr(1)
            .����ǰ�ʻ���� = Val(strArr(2))
            .�����ʻ�֧����� = Val(strArr(3))
            .�Էѽ�� = Val(strArr(4))
            .���Ѻ��ʻ���� = Val(strArr(5))
            .����ʱ�� = strArr(6)
            .ǰ�˵��ݺ� = strArr(7)
            .���ĵ��ݺ� = strArr(8)
            .������ = strArr(9)
            .����Ա���� = strArr(10)
            .ǰ������ = strArr(11)
        End With
    End If
    ������ϸд�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ���㷽ʽ����() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������ʾ������
    '--�����:
    '--������:str���㷽ʽ
    '--��  ��:�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    Dim dbl�����ܶ� As Double
        
    '�����ܶ�=�����Էѽ��+����ͳ��֧�����+��ͳ����      �˽����������˺�������湫ʽת��������
    
    '�����Էѽ�� = �ܷ��ö� - ����ͳ��֧����� - �� / �߶�ͳ��֧�����
    '�Էѽ��ֽ�֧����ʻ�֧���� (��:��ѡ�����ֽ�����ʻ�֧��)
    '��ͳ����߶�ͳ��������ͬ
    'ͳ��֧��������ҽ���ڷ��ø��ݲ�ͬ���𸶱�׼�ͱ���������ҽ��������
    '��˵�����ݱ��������漼�������ɷ����޹�˾�������Ľ���
    ���㷽ʽ���� = False
    
    Err = 0:    On Error GoTo errHand:
    DebugTool "����(" & "Get���㷽ʽ" & ")"
    
    '����||����||����ǰ�ʻ����||�����ʻ�֧�����||�Էѽ��||���Ѻ��ʻ����||����ʱ��||ǰ�˵��ݺ�||���ĵ��ݺ�||������||����Ա����||ǰ������
    dbl�����ܶ� = g��������.�����ʻ�֧����� + g��������.�Էѽ��
    str���㷽ʽ = "||�����ʻ�|" & g��������.�����ʻ�֧�����
    
    If Format(g�������_�ϳ�����.�����ܶ�, "###0.00;-###0.00;0;0") <> Format(dbl�����ܶ�, "###0.00;-###0.00;0;0") Then
        Dim blnYes As Boolean
        '�����ܶ���ҽ�����ķ����ܶ��,���ܽ��н���
        ShowMsgbox "���ν����ܶ�(" & g�������_�ϳ�����.�����ܶ� & ") ��" & vbCrLf & _
                    "   ���ķ��ص��ܶ�(" & dbl�����ܶ� & ")���������ݷ���������������ҽ��������ϵ!"
        Exit Function
    End If
    If g�������_�ϳ�����.�����ʻ�֧�� <> g��������.�����ʻ�֧����� Then
        ShowMsgbox "���������������ʻ�֧��(" & g�������_�ϳ�����.�����ʻ�֧�� & ") ��" & vbCrLf & _
                    "   ����ĸ����ʻ�֧��(" & g��������.�����ʻ�֧����� & ")���ȣ��������ݷ���������������ҽ��������ϵ!"
        Exit Function
    End If
    ���㷽ʽ���� = True
'
'    Exit Function
'   '�������,�򱣴��Ԥ����¼��
'    If str���㷽ʽ <> "" Then
'        str���㷽ʽ = Mid(str���㷽ʽ, 3)
'        g�������_�ɶ��ڽ�.���㷽ʽ = str���㷽ʽ
'
'        If g��������.�����־ = 0 Then
'            gstrSQL = "zl_���˽����¼_Update(" & g��������.����ID & ",'" & str���㷽ʽ & "', 0)"
'            Call zldatabase.ExecuteProcedure(gstrsql,"����Ԥ����¼")
'        Else
'                gstrSQL = "zl_���˽����¼_Update(" & g��������.����ID & ",'" & str���㷽ʽ & "',1)"
'                Call zldatabase.ExecuteProcedure(gstrsql,"����Ԥ����¼")
'        End If
'    End If
'
'    DebugTool "��ʼ��ʾ���㷽ʽ"
'    '��ʾ������Ϣ
'    If frm������Ϣ.ShowME(g��������.����ID, False, "", IIf(g��������.�����־ = 0, 0, 1)) = False Then
'
'        DebugTool "���㷽ʽ��ʾʧ��"
'        ���㷽ʽ���� = False
'        Exit Function
'    End If
    DebugTool "���㷽ʽ��ʾ�ɹ�"
    ���㷽ʽ���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ��ȡ�����ʻ�֧��() As Double
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�����ʻ�ֵ(��Ԥ����¼�л�ȡ)
    '--�����:
    '--������:
    '--��  ��:�ɹ�,���ر��θ����ʻ�֧��,���򷵻�0
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ��� From ����Ԥ����¼ where ����ID=[1] and  ���㷽ʽ='�����ʻ�'"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ�֧��", g��������.����ID)
    If Not rsTemp.EOF Then
        ��ȡ�����ʻ�֧�� = Nvl(rsTemp!���, 0)
    End If
    
End Function
Private Function InsertIntoҽ�������¼(ByVal strArr As Variant, ByVal lng����ID As Long) As Boolean
    '����:���м�����ҽ�������¼
    '����:strarr��split(stroutput,"||")����������
    '����:strArr(0)-סԺ���ý�����,strArr(1-n)�����ֶ���ϸ
      '   ˵��:
    '������(20060920):����޸ĺ�ķ���ֵ:
    ' סԺ���ý�����:�𸶽��(0)||�˶��ɷѻ���(1)||�����ⶥ���(2)||����ⶥ���(3)||����סԺ����(4)||�����ۼ�סԺ����(5)||
    '                  �������ۼ�סԺ����(6)||��������״̬(7)||�������״̬(8)||����Ա����״̬(9)||������������(10)||
    '                  ���䱨������(11)||����Ա��������(12)||�������(13)||����ҩƷ��(14)||�������Ʒ�(15)||��������(16)||
    '                  �������(17)||����ҩƷ��(18)||�������Ʒ�(19)||����������(20)||����ҩƷ�����Ը�(21)||�������������Ը�(22)||
    '                  �������(23)||����ҩƷ��(24)||�������Ʒ�(25)||���������ʩ��(26)||����ҽ�Ʊ���(27)||����ҽ�Ʊ���(28)||����Աҽ�Ʊ���(29)||ҽ�Ʊ����ܼ�(30)||�Ը��ϼ�(31)
    ' �����ֶ���ϸ(����):סԺ��������(0)||���εı�������(1)||�𸶽��(2)||ȫ�Ը�ҩƷ��(3)||ȫ�Ը����Ʒ�(4)||ȫ�Ը�������ʩ��(5)||
    '                    ����ҩƷ��(6)||�������Ʒ�(7)||�Ը�����(8)||�Ը����(9)||�Ը�С��(10)||���α������(11)
    Dim tmpArr As Variant
    Dim i As Long
    Err = 0
    On Error GoTo errHand:
    InsertIntoҽ�������¼ = False
    
    '����סԺ��������
    tmpArr = Split(strArr(0), "||")
    
    DebugTool "����InsertIntoҽ�������¼"
       
    gstrSQL = "ZL_ҽ�������¼_INSERT(2"
    gstrSQL = gstrSQL & "," & lng����ID
    gstrSQL = gstrSQL & "," & Val(tmpArr(0)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(1)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(2)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(3)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(4)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(5)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(6)) & ""
    gstrSQL = gstrSQL & ",'" & tmpArr(7) & "'"
    gstrSQL = gstrSQL & ",'" & tmpArr(8) & "'"
    gstrSQL = gstrSQL & ",'" & tmpArr(9) & "'"
    gstrSQL = gstrSQL & "," & Val(tmpArr(10)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(11)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(12)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(13)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(14)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(15)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(16)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(17)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(18)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(19)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(20)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(21)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(22)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(23)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(24)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(25)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(26)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(27)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(28)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(29)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(30)) & ""
    gstrSQL = gstrSQL & "," & Val(tmpArr(31)) & ")"
    ExecuteProcedure_�ϳ����� "��������¼���м��"
        
    '������ϸ����
    '������(2005-07-27):������δ�����ߵ�ʱ�򣬻����η���Ϊ��ֵ����������жϡ�
    For i = 1 To UBound(strArr)
        '������ϸ����
         
         'סԺ��������(0)||���εı�������(1)||�𸶽��(2)||ȫ�Ը�ҩƷ��(3)||ȫ�Ը����Ʒ�(4)||ȫ�Ը�������ʩ��(5)||
         '����ҩƷ��(6)||�������Ʒ�(7)||�Ը�����(8)||�Ը����(9)||�Ը�С��(10)||���α������(11)

        tmpArr = Split(strArr(i), "||")
        If UBound(tmpArr) > 0 Then
            gstrSQL = "ZL_ҽ������ֶ���ϸ_INSERT("
            gstrSQL = gstrSQL & "2"
            gstrSQL = gstrSQL & "," & lng����ID & ""
            gstrSQL = gstrSQL & ",'" & IIf(tmpArr(0) = "", "������" & i, tmpArr(0)) & "'"
            gstrSQL = gstrSQL & "," & Val(tmpArr(1)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(2)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(3)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(4)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(5)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(6)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(7)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(8)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(9)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(10)) & ""
            gstrSQL = gstrSQL & "," & Val(tmpArr(11)) & ")"
            
            ExecuteProcedure_�ϳ����� "�������ֶ���Ϣ���м��"
        End If
    Next
    InsertIntoҽ�������¼ = True
    DebugTool "����ҽ�������¼�ɹ�"
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function ����ҽ����Ժ_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ����²��˵ĳ�Ժ����������������������ʱ���߻����
    Dim StrInput As String
    Dim strOutput As String
    Dim blnYes  As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    
    On Error GoTo errHandle
    '������(2006-2-17):���ܴ��ڳ�������,����ֱ������Ժ�Ǽ��н�����Ժ����
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        If MsgBox("�ò����Ѿ���Ժ���޷��÷���������ͨ����Ժ�Ǽǽ�����Ժ����!���п��ܴ��ڳ�������,���ֻȡ��ҽ���Ǽ�,�����˳�!", vbYesNo) = vbYes Then
           Exit Function
        End If
    End If
    
    
    gstrSQL = "Select * From סԺ���ü�¼ where nvl(�Ƿ��ϴ�,0)=1 and rownum<=1 and ����id=" & lng����ID & " and ��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�ж��Ƿ�����ϴ���¼"
        
    If Not rsTemp.EOF Then
        ShowMsgbox "�Ѿ����ϴ���������ϸ���ã��Ƿ����Ҫȡ��ҽ����Ժ?", True, blnYes
        If blnYes = False Then
            Exit Function
        End If
    End If
    
    
    If lng��ҳID > 9 Then
        StrInput = "9" & Lpad(lng��ҳID, 3, "0") & lng����ID
    Else
        StrInput = Rpad(lng��ҳID, 4, "0") & lng����ID
    End If
    
    If ҵ������_�ϳ�����(ȡ����Ժ�Ǽ�_����, StrInput, strOutput) = False Then Exit Function

    '����ҽ���ʻ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    
    gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
    
    '�����ϴ���־
    gstrSQL = "update סԺ���ü�¼ set �Ƿ��ϴ�=0 where ���ʽ�� is null and ����ID= " & lng����ID & " and ��ҳID= " & lng��ҳID
    gcnOracle.Execute gstrSQL
    
    DebugTool "ȡ���ɹ�"
    
    ����ҽ����Ժ_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
