Attribute VB_Name = "mdl������"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

'���������˺����20040924
Private mblnInit As Boolean

Public Enum ҵ������_����
    ��ʼ��������� = 0
    ����
    ȡҩƷ��Ϣ
    ȡ������Ϣ
    ȡ������Ϣ
    ���÷�Ʊ����
    ���������վ���ϸ
    �������������Ϣ
    ȡ�������վݼ�����Ϣ
    ȡ������
    ��Ժ�Ǽ�
    ȡ����Ժ�Ǽ�
    ��Ժ�Ǽ�
    ȡ����Ժ�Ǽ�
    ����סԺ�ʵ�
    ���ü��ʵ���ϸ����
    ���ý��㵥
    ����סԺ������Ϣ
    ȡ��סԺ���������Ϣ
    ���������ύ
    ���߷����ύ
    �����������
End Enum
Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    ��ϸʱʵ�ϴ� As Boolean
End Type
Public InitInfor_���� As InitbaseInfor

Private Type �������
        ���� As String
        ����  As String
        ���֤�� As String
        ����     As String
        �Ա�     As String
        ��������  As String
        ҽ����    As String
        ��λ����    As String
        �������    As String
        ����Ա��־  As String   '(0-�ǹ���Ա1-����Ա���������չ���Ա)
        ���䱣��    As String   '0-���μ�1-�μ�
        ��ҽ��    As String   '0-���μ�1-�μ�
        ������ϵ    As String   '1-��������0-����
        �չ˼���    As String   '0-��1-һ��2-����3-����
        ְ������    As String   '0����1��פ���2-��ذ���
        �Ƿ����Բ�  As String   '0-����1-��
        �ش󼲲�    As String   '0-����1-��
        סԺ��־    As String   '0-��סԺ 1-סԺ
        �𸶶�ҽ�Ʒ��ۼ�  As Double
        ͳ��֧���ۼ� As Double
        ����ͳ���ۼ�  As Double
        ���߽���ۼ� As Double
        ����ͳ��֧���ۼ� As Double
        ���ۼ�    As Double
        �ʻ����    As Double
        סԺ����    As Integer
        ֧������ As Integer
        
        �����ܶ�    As Double
        ��ϱ���    As String
        �������    As String
        ���ִ���    As String
        ��������  As Integer
End Type
Public g�������_���� As �������



'-----------------------------------------------------------------------------------------------------------------

Private str�������� As String * 1, str���ղ��� As String * 10, strTempArr(50, 1) As String

'===============================================================================================================
'����: ��ʼ�������
'��ڲ���: ��������(2)
'˵��: 10-�����շ�,11-�����˷�,20-��Ժ�Ǽ�,21-ҽ��¼��,22-סԺ����,23-סԺ����,24-��Ժ�Ǽ�,25-ȡ����Ժ�Ǽ�
'      26-����סԺ����,27-ȡ������,28-ȡ����Ժ�Ǽ�,29-ȡ��ҽ��¼��,����-ϵͳ����/�ֵ书��
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function InitCalc Lib "YhYbClient.dll" (ByVal str�������� As String) As Long

'===============================================================================================================
'����: �����������
'��ڲ���: ��
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function FinalCalc Lib "YhYbClient.dll" () As Long

'===============================================================================================================
'����: ȡҽ��������
'��ڲ���: ��
'���ڲ���: ��
'����: ϵͳ����ʹ�õ�Ӧ�÷���������
'===============================================================================================================
Public Declare Function GetAppServerName Lib "YhYbClient.dll" () As String

'===============================================================================================================
'����: ָ���ṩ����ķ�����
'��ڲ���: Ӧ�÷���������(windows���Ƴ���)
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetAppServerName Lib "YhYbClient.dll" (ByVal str�������� As String) As Long

'===============================================================================================================
'����: ���ÿ���д���˿�
'��ڲ���: �˿ں�
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetCardPort Lib "YhYbClient.dll" (ByVal str�˿ں� As String) As Long

'===============================================================================================================
'����: ��ȡ������Ϣ
'��ڲ���: ��
'���ڲ���: ��
'����: ������Ϣ
'===============================================================================================================
Public Declare Function GetErrMsg Lib "YhYbClient.dll" () As String

'===============================================================================================================
'����: ȡҽ�ƻ�����Ϣ
'��ڲ���: ��
'���ڲ���: ���Ĵ���(4λ),��������(4λ),��������(40λ),ҽԺ����(2λ),��������(20λ)
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function GetInfo_MediOrgan Lib "YhYbClient.dll" (ByVal str���Ĵ��� As String, _
    ByVal str�������� As String, ByVal str�������� As String, ByVal strҽԺ���� As String, _
    ByVal str�������� As String) As Long

'===============================================================================================================
'����: ȡҩƷ�ֵ���Ϣ
'��ڲ���: ��Ŀ����(10)
'���ڲ���: ��Ŀ����(40λ),���ô���(2λ),��Ŀ���(1λ),�Ƿ�ҽ��(1λ),�Ƿ��޼�(1λ),�Ը�����(4λ��3λС��)
'          ��׼����(8λ��2λС��)
'˵��: ��Ŀ���:0-�׻���ͨ,1-�һ�߾���,2-�Է�
'      �Ƿ�ҽ��:0-����,1-��
'      �Ƿ��޼�:0-����,1-��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function GetInfo_MediDic Lib "YhYbClient.dll" (ByVal str��Ŀ���� As String, _
    str��Ŀ���� As String, str���ô��� As String, str��Ŀ��� As String, _
    str�Ƿ�ҽ�� As String, str�Ƿ��޼� As String, dbl�Ը����� As Double, _
    dbl��׼���� As Double) As Long

'===============================================================================================================
'����: ȡ�����ֵ���Ϣ
'��ڲ���: ��Ŀ����(10)
'���ڲ���: ��Ŀ����(40λ),���ô���(2λ),��Ŀ���(1λ),�Ƿ�ҽ��(1λ),�Ƿ��޼�(1λ),�Ը�����(4λ��3λС��)
'          ��׼����(8λ��2λС��)
'˵��: ��Ŀ���:0-�׻���ͨ,1-�һ�߾���,2-�Է�
'      �Ƿ�ҽ��:0-����,1-��
'      �Ƿ��޼�:0-����,1-��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function GetInfo_ItemDic Lib "YhYbClient.dll" (ByVal str��Ŀ���� As String, _
    str��Ŀ���� As String, str���ô��� As String, str��Ŀ��� As String, _
    str�Ƿ�ҽ�� As String, str�Ƿ��޼� As String, dbl�Ը����� As Double, _
    ByRef dbl��׼���� As Double) As Long

'===============================================================================================================
'����: ȡ������ʩ�ֵ���Ϣ
'��ڲ���: ��Ŀ����(10)
'���ڲ���: ��Ŀ����(40λ),���ô���(2λ),��Ŀ���(1λ),�Ƿ�ҽ��(1λ),�Ƿ��޼�(1λ),�Ը�����(4λ��3λС��)
'          ��׼����(8λ��2λС��)
'˵��: ��Ŀ���:0-�׻���ͨ,1-�һ�߾���,2-�Է�
'      �Ƿ�ҽ��:0-����,1-��
'      �Ƿ��޼�:0-����,1-��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function GetInfo_ServerDic Lib "YhYbClient.dll" (ByVal str��Ŀ���� As String, _
    str��Ŀ���� As String, str���ô��� As String, str��Ŀ��� As String, _
    str�Ƿ�ҽ�� As String, str�Ƿ��޼� As String, dbl�Ը����� As Double, _
    dbl��׼���� As Double) As Long

'===============================================================================================================
'����: ȡ������Ϣ
'��ڲ���: ���ֱ���(10)
'���ڲ���: ��������(10λ),���ָ���(10λ),��������(40λ),����ͳ����(3λ),ע��(200λ),���ֱ�־(1λ),�������(2λ)
'          �Ʊ�(3λ)
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function GetInfo_SickDic Lib "YhYbClient.dll" (ByVal str���ֱ��� As String, _
     str�������� As String, str���ָ��� As String, str�������� As String, _
    str����ͳ���� As String, strע�� As String, str���ֱ�־ As String, _
    str������� As String, str�Ʊ� As String) As Long

'===============================================================================================================
'����: ȡ���ô�����Ϣ
'��ڲ���: ���ô������(2λ)
'���ڲ���: ���ô�������(20λ)
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function GetInfo_MediKind Lib "YhYbClient.dll" (ByVal str������� As String, _
     str�������� As String) As Long

'===============================================================================================================
'����: ����
'��ڲ���: ��
'���ڲ���: ���Ĵ���(4λ),����(16λ),���֤��(18λ),����(10λ),�Ա�(1λ),��������(8λ),ҽ����(8λ)
'          ���˵�λ����(5λ),�������,����Ա��־(1λ),�Ƿ�μӲ��䱣��(1λ),�μӴ�ҽ��(1λ),������ϵ(1λ)
'          �չ˼���(1λ),ְ������(1λ),�Ƿ����Բ�(1λ),�Ƿ��ش󼲲�(1λ),סԺ��־(1λ)
'          �𸶶�����ҽ�Ʒ��ۼ�(8λ,2λС��),����ͳ��֧���ۼ�(8λ,2λС��),�����߽���ۼ�(8λ,2λС��)
'          ���Բ�����ͳ��֧���ۼ�(8λ,2λС��),�ش󼲲�����ͳ��֧���ۼ�(8λ,2λС��),�ʻ����(8λ,2λС��)
'          ������ЧסԺ����(3λ)
'˵��: �Ա�:0-Ů,1-��
'      ��������:��ʽyyyymmdd
'      �������:0-��ְ,1-����
'      ����Ա��־:0-�ǹ���Ա,1-����Ա,�������չ���Ա
'      �Ƿ�μӲ��䱣��:0-���μ�,1-�μ�
'      �μӴ�ҽ��:0-���μ�,1-�μ�
'      ������ϵ:1-��������,0-����
'      �չ˼���:0-��,1-һ��,2-����,3-����
'      ְ������:0-����,1-��פ���,2-��ذ���
'      �Ƿ����Բ�:0-����,1-��
'      �Ƿ��ش󼲲�:0-����,1-��
'      סԺ��־:0-��סԺ,1-סԺ
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function ReadCard Lib "YhYbClient.dll" (str���Ĵ��� As String, str���� As String, _
     str���֤�� As String, STR���� As String, str�Ա� As String, _
     str�������� As String, strҽ���ʺ� As String, str���˵�λ���� As String, _
     str������� As String, str����Ա��־ As String, str���䱣�� As String, _
     str��ҽ�� As String, str������ϵ As String, str�Ƿ����Բ� As String, str�Ƿ�� As String, str�չ˼��� As String, _
     strְ������ As String, _
     strסԺ��־ As String, dbl����ͳ���ۼ� As Double, dbl����ͳ���ۼ� As Double, _
     dbl�������ۼ� As Double, dbl�ز������ۼ� As Double, dbl�󲡱����ۼ� As Double, _
     dbl�ʻ���� As Double, int����סԺ���� As Long, int֧�����к� As Long) As Long            '�ĵ���������˵����һ��
    
    

'===============================================================================================================
'����: ���������վ�
'��ڲ���: ����,�վݺ�(13λ),�����������(1λ),ҽ�����ִ���(10λ),���˵��(200λ),��������(20λ),ҽ������(10λ)
'          ҩʦ(10λ),��ҩ����(2λ),���(8λ,2λС��)
'˵��: �վݺ�:Not Null
'      �����������:Not Null,1-��ͨ,2-���Բ�,3-�ش󼲲�,4-�չ˶���,5-������,6-�ƻ�����,7-����
'      ���:>0
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetClinicBill Lib "YhYbClient.dll" (ByVal lng���� As Long, ByVal str�վݺ� As String, ByVal str����������� As String, ByVal strҽ�����ִ��� As String, _
    ByVal str���˵�� As String, ByVal str�������� As String, ByVal strҽ������ As String, _
    ByVal strҩʦ As String, ByVal dbl��ҩ���� As Double, ByVal dbl��� As Double) As Long

'===============================================================================================================
'����: ���������վ���ϸ
'��ڲ���: ����,ҽ����Ŀ���(10λ),ҽԺ��Ŀ����(40λ),��������(20λ),��λ����(14λ),�÷�����(40λ)
'          ���ô������(2λ),�������(1λ),�Ƿ�ҽ��(1λ),�Ƿ�ҩƷ(1λ),����(8λ,2λС��),����(8λ,2λС��)
'          ���(8λ,2λС��)
'˵��: ҽ����Ŀ���:Not Null
'      ҽԺ��Ŀ����:Not Null
'      ���ô������:Not Null
'      �������:0-�׻���ͨ,1-�һ�߾���,2-�Է�,Not Null
'      �Ƿ�ҽ��:0-����,1-��[Ϊ��ɽӿڼ��ݶ�����δ��]
'      �Ƿ�ҩƷ:0-��Ŀ,1-ҩƷ,2-������ʩ[��],Not Null
'      ����,����,���:>0
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetClinicBillDetail Lib "YhYbClient.dll" (ByVal lng���� As Long, _
    ByVal strҽ����Ŀ��� As String, ByVal strҽԺ��Ŀ���� As String, ByVal str�������� As String, _
    ByVal str��λ���� As String, ByVal str�÷����� As String, ByVal str���ô������ As String, _
    ByVal str������� As String, ByVal str�Ƿ�ҽ�� As String, ByVal str�Ƿ�ҩƷ As String, _
    ByVal dbl���� As Double, ByVal dbl���� As Double, ByVal dbl��� As Double) As Long

'===============================================================================================================
'����: �������������Ϣ
'��ڲ���: ����,ҽ�����ô������(2λ),��Ӧ���ô�����(8λ,2λС��)
'˵��: ҽ�����ô������:Not Null
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetClinicMediKind Lib "YhYbClient.dll" (ByVal lng���� As Long, _
    ByVal str������� As String, ByVal dbl������ As Double) As Long

'===============================================================================================================
'����: ��Ժ�Ǽ�
'��ڲ���: ���˾��(��ReadCard������÷���),סԺ��(13λ),��Ժ����(8λ),סԺ����(20λ),����(20λ),����(10λ)
'          ����(3λ),����ҽ��(10λ),��Ժ��ϴ���(10λ),str�������(200)
'˵��: סԺ��:Not Null
'      ��Ժ����:��ʽyyyymmdd,Not Null
'      ��Ժ��ϴ���:���ִ���,Not Null
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function InhosRegister Lib "YhYbClient.dll" (ByVal int���˾�� As Integer, _
    ByVal strסԺ�� As String, ByVal str��Ժ���� As String, ByVal strסԺ���� As String, _
    ByVal str���� As String, ByVal str���� As String, ByVal str���� As String, ByVal str����ҽ�� As String, _
    ByVal str��Ժ��ϴ��� As String, ByVal str������� As String) As Long

'===============================================================================================================
'����: ��Ժ�Ǽ�
'��ڲ���: סԺ��(13λ),��Ժ����(8λ),סԺ����(20λ),����(20λ),����(10λ),����(3λ),����ҽ��(10λ)
'          ��Ժ��ϴ���(10λ),סԺ����(3λ),��Ժ�������(1λ)
'˵��: סԺ��:Not Null
'      ��Ժ����:��ʽyyyymmdd,Not Null
'      ��Ժ��ϴ���:���ִ���,Not Null
'      ��Ժ�������:1-����,2-��ת,3-δ��,4-����,9-����,Not Null
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function OuthosRegister Lib "YhYbClient.dll" (ByVal strסԺ�� As String, _
    ByVal str��Ժ���� As String, ByVal strסԺ���� As String, ByVal str���� As String, _
    ByVal str���� As String, ByVal str���� As String, ByVal strҽ�� As String, _
    ByVal str��Ժ��ϴ��� As String, ByVal str��Ժ������� As String, ByVal lngסԺ���� As Long, ByVal str���� As String) As Long
    '�ĵ���������˵���Ĳ�����һ��

'===============================================================================================================
'����: ����ҽ��
'��ڲ���: סԺ��(13λ),ҽ����(13λ),����ҽ������(10λ),ͣ��ҽ������(10λ),ִ������(10λ),¼����(10λ)
'          �Ƿ���ҽ��(1λ),ҽ����ʼ����(8λ),ҽ����ʼʱ��(8λ),ҽ��ֹͣ����(8λ),ҽ��ֹͣʱ��(8λ)
'          ִ������(8λ),ִ��ʱ��(8λ),¼������(8λ),ҽ������(200λ)
'˵��: סԺ��:Not Null
'      ҽ����:Not Null
'      �Ƿ���ҽ��:0-����,1-��
'      ҽ����ʼ����,ҽ��ֹͣ����,ִ������,¼������:��ʽyyyymmdd,not null
'      ҽ����ʼʱ��,ҽ��ֹͣʱ��,ִ��ʱ��:��ʽhh:mi:ss,not null
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetDoctorAdvice Lib "YhYbClient.dll" (ByVal intסԺ�� As Integer, _
    ByVal strҽ���� As String, ByVal str����ҽ������ As String, ByVal strͣ��ҽ������ As String, _
    ByVal strִ������ As String, ByVal str¼���� As String, ByVal str�Ƿ���ҽ�� As String, _
    ByVal strҽ����ʼ���� As String, ByVal strҽ����ʼʱ�� As String, ByVal strҽ��ֹͣ���� As String, _
    ByVal strҽ��ֹͣʱ�� As String, ByVal strִ������ As String, ByVal strִ��ʱ�� As String, _
    ByVal str¼������ As String, ByVal strҽ������ As String) As Long

'===============================================================================================================
'����: ����סԺ�ʵ�
'��ڲ���: סԺ��(13),סԺ�ʵ���(13λ),���ִ���(10λ),����(20λ),ҽ��(10λ),�в�ҩ����(2λ),���(8λ,2λС��)
'˵��: סԺ��:Not Null
'      סԺ�ʵ���:Not Null
'      ���ִ���:Not Null
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetInHosBill Lib "YhYbClient.dll" (ByVal strסԺ�ǼǺ� As String, _
    ByVal str�ʵ��� As String, ByVal str���ִ��� As String, ByVal str���� As String, ByVal strҽ�� As String, _
    ByVal dbl�в�ҩ���� As Double, ByVal dbl��� As Double) As Long

'===============================================================================================================
'����: ����סԺ�ʵ���ϸ
'��ڲ���: ����,ҽ����Ŀ���(10λ),ҽԺ��Ŀ����(40λ),��������(20λ),��λ����(14λ),�÷�����(40λ)
'          ���ô������(2λ),�������(1λ),�Ƿ�ҽ��(1λ),�Ƿ�ҩƷ(1λ),����(8λ,2λС��),����(8λ,2λС��)
'          ���(8λ,2λС��)
'˵��: ҽ����Ŀ���:Not Null
'      ҽԺ��Ŀ����:Not Null
'      ���ô������:Not Null
'      �������:0-�׻���ͨ,1-�һ�߾���,2-�Է�,Not Null
'      �Ƿ�ҽ��:0-����,1-��[Ϊ��ɽӿڼ��ݶ�����δ��]
'      �Ƿ�ҩƷ:0-��Ŀ,1-ҩƷ,2-������ʩ[��],Not Null
'      ����,����,���:>0
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetInHosBillDetail Lib "YhYbClient.dll" (ByVal int���� As Integer, _
    ByVal strҽ����Ŀ��� As String, ByVal strҽԺ��Ŀ���� As String, ByVal str�������� As String, _
    ByVal str��λ���� As String, ByVal str�÷����� As String, ByVal str���ô������ As String, _
     ByVal str�Ƿ�ҽ�� As String, ByVal str������� As String, ByVal str�Ƿ�ҩƷ As String, _
    ByVal dbl���� As Double, ByVal dbl���� As Double, ByVal dbl��� As Double) As Long

'===============================================================================================================
'����: ���ý��㵥
'��ڲ���: ����,���㵥��(13λ),�����ܶ�(8λ,2λС��),��Ժ����(8λ),��Ժ����(8λ),����(20λ),ҽ��(10λ)
'          ��Ժ���ֱ���(10λ),����֢˵��(200λ),��Ժ���(1λ),סԺ����(3λ)
'˵��: ���㵥��:Not Null
'      ��Ժ����,��Ժ����:��ʽyyyymmdd,Not Null
'      ��Ժ���:1-����,2-��ת,3-δ��,4-����,9-����,Not Null
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetCheckOutBill Lib "YhYbClient.dll" (ByVal lng���� As Long, _
    ByVal str���㵥�� As String, ByVal str��Ժ���� As String, _
    ByVal str��Ժ���� As String, ByVal str���� As String, ByVal strҽ�� As String, _
    ByVal str���ֱ��� As String, ByVal str����֢ As String, ByVal STR��Ժ��� As String, _
    ByVal intסԺ���� As Long, ByVal dbl�����ܶ� As Double) As Long

'===============================================================================================================
'����: ���ý��㵥��ϸ
'��ڲ���: ����,ҽ����Ŀ���(10λ),ҽԺ��Ŀ����(40λ),��������(20λ),��λ����(14λ),�÷�����(40λ)
'          ���ô������(2λ),�������(1λ),�Ƿ�ҽ��(1λ),�Ƿ�ҩƷ(1λ),����(8λ,2λС��),����(8λ,2λС��)
'          ���(8λ,2λС��)
'˵��: ҽ����Ŀ���:Not Null
'      ҽԺ��Ŀ����:Not Null
'      ���ô������:Not Null
'      �������:0-�׻���ͨ,1-�һ�߾���,2-�Է�,Not Null
'      �Ƿ�ҽ��:0-����,1-��[Ϊ��ɽӿڼ��ݶ�����δ��]
'      �Ƿ�ҩƷ:0-��Ŀ,1-ҩƷ,2-������ʩ[��],Not Null
'      ����,����,���:>0
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetCheckOutBillDetailX Lib "YhYbClient.dll" (ByVal lng���� As Long, _
    ByVal strҽ����Ŀ��� As String, ByVal strҽԺ��Ŀ���� As String, ByVal str�������� As String, _
    ByVal str��λ���� As String, ByVal str�÷����� As String, ByVal str���ô������ As String, _
    ByVal str������� As String, ByVal str�Ƿ�ҽ�� As String, ByVal str�Ƿ�ҩƷ As String, _
    ByVal dbl���� As Double, ByVal dbl���� As Double, ByVal dbl��� As Double) As Long

'===============================================================================================================
'����: ����סԺ������Ϣ
'��ڲ���: ����,ҽ�����ô������(2λ),��Ӧ���ô�����(8λ,2λС��)
'˵��: ҽ�����ô������:Not Null
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function SetInHosMediKind Lib "YhYbClient.dll" (ByVal lng���� As Long, ByVal str������� As String, ByVal dbl������ As Double) As Long

'===============================================================================================================
'����: ȡ����Ժ�Ǽ�
'��ڲ���: ����
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function uInhosRegister Lib "YhYbClient.dll" (ByVal int���� As Integer) As Long

'===============================================================================================================
'����: ȡ����Ժ�Ǽ�
'��ڲ���: סԺ��(13λ)
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function uOuthosRegister Lib "YhYbClient.dll" (ByVal strסԺ�� As String) As Long

'===============================================================================================================
'����: ���ҩ���˷�,סԺ�˷�, סԺȡ������
'��ڲ���: ����,�˷��µ��ݺ�(13λ),Ҫ�˷ѵĵ��ݺ�(13λ)
'˵��: �˷��µ��ݺ�,Ҫ�˷ѵĵ��ݺ�:Not Null
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function ReturnCharge Lib "YhYbClient.dll" (ByVal int���� As Integer, ByVal str�µ��� As String, _
    ByVal strԭ���� As String) As Long

'===============================================================================================================
'����: ȡ��ҽ��¼��
'��ڲ���: ҽ����(13λ)
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function uDoctorAdvice Lib "YhYbClient.dll" (ByVal strҽ���� As String) As Long

'===============================================================================================================
'����: ȡ�������վݼ�����Ϣ(ҩ��)
'��ڲ���: ����
'���ڲ���: �ʻ����(8λ,2λС��),�����ʻ�֧��(8λ,2λС��),�����ֽ�֧��(8λ,2λС��),���˱�������(8λ,2λС��)
'          ͳ��֧��(8λ,2λС��),�չ�֧��(8λ,2λС��),�չ˵渶(8λ,2λС��),�Ը���֧��(8λ,2λС��)
'          �̱�֧��(8λ,2λС��),����ҩƷ(8λ,2λС��),�Է�ҩƷ(8λ,2λС��),����ҩƷ(8λ,2λС��)
'          ��������(��ͨ)(8λ,2λС��),�Է�����(8λ,2λС��),��������(�߾���)(8λ,2λС��),������ʩ(8λ,2λС��)
'          �Է���ʩ(8λ,2λС��),������ʩ(8λ,2λС��),�����Է�(8λ,2λС��),�Ը����ۼ�(8λ,2λС��)
'          ͳ��֧���ۼ�(8λ,2λС��),�ز�֧���ۼ�(8λ,2λС��),����֧���ۼ�(8λ,2λС��)
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function GetClinicBillData Lib "YhYbClient.dll" (ByVal lng���� As Long, _
    ByRef dbl�ʻ���� As Double, ByRef dbl�����ʻ�֧�� As Double, ByRef dbl�����ֽ�֧�� As Double, _
    ByRef dbl���˱������� As Double, ByRef dblͳ��֧�� As Double, ByRef dbl�չ�֧�� As Double, _
    ByRef dbl�չ˵渶 As Double, ByRef dbl�Ը���֧�� As Double, ByRef dbl�̱�֧�� As Double, _
    ByRef dbl����ҩƷ As Double, ByRef dbl�Է�ҩƷ As Double, ByRef dbl����ҩƷ As Double, _
    ByRef dbl�������� As Double, ByRef dbl�Է����� As Double, ByRef dbl�������� As Double, _
    ByRef dbl������ʩ As Double, ByRef dbl�Է���ʩ As Double, ByRef dbl������ʩ As Double, _
    ByRef dbl�����Է� As Double, ByRef dbl�Ը����ۼ� As Double, ByRef dblͳ��֧���ۼ� As Double, _
    ByRef dbl�ز�֧���ۼ� As Double, ByRef dbl����֧���ۼ� As Double, ByRef dbl�ǻ���ҽ�Ʒ� As Double) As Long

'===============================================================================================================
'����: ȡ��סԺ���������Ϣ
'��ڲ���: ����
'���ڲ���: �ʻ����(8λ,2λС��),�����ʻ�֧��(8λ,2λС��),�����ֽ�֧��(8λ,2λС��),���˱�������(8λ,2λС��)
'          ͳ��֧��(8λ,2λС��),�չ�֧��(8λ,2λС��),�չ˵渶(8λ,2λС��),�Ը���֧��(8λ,2λС��)
'          �̱�֧��(8λ,2λС��),����ҩƷ(8λ,2λС��),�Է�ҩƷ(8λ,2λС��),����ҩƷ(8λ,2λС��)
'          ��������(��ͨ)(8λ,2λС��),�Է�����(8λ,2λС��),��������(�߾���)(8λ,2λС��),������ʩ(8λ,2λС��)
'          �Է���ʩ(8λ,2λС��),������ʩ(8λ,2λС��),�����Է�(8λ,2λС��),�Ը����ۼ�(8λ,2λС��)
'          ͳ��֧���ۼ�(8λ,2λС��),�ز�֧���ۼ�(8λ,2λС��),����֧���ۼ�(8λ,2λС��)
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function GetCheckOutBillData Lib "YhYbClient.dll" (ByVal lng���� As Long, _
    ByRef dbl�ʻ���� As Double, ByRef dbl�����ʻ�֧�� As Double, ByRef dbl�����ֽ�֧�� As Double, _
    ByRef dbl���˱������� As Double, ByRef dblͳ��֧�� As Double, ByRef dbl�չ�֧�� As Double, _
    ByRef dbl�չ˵渶 As Double, ByRef dbl�Ը���֧�� As Double, ByRef dbl�̱�֧�� As Double, _
    ByRef dbl����ҩƷ As Double, ByRef dbl�Է�ҩƷ As Double, ByRef dbl����ҩƷ As Double, _
    ByRef dbl�������� As Double, ByRef dbl�Է����� As Double, ByRef dbl�������� As Double, _
    ByRef dbl������ʩ As Double, ByRef dbl�Է���ʩ As Double, ByRef dbl������ʩ As Double, _
    ByRef dbl�����Է� As Double, ByRef dbl�Ը����ۼ� As Double, ByRef dblͳ��֧���ۼ� As Double, _
    ByRef dbl�ز�֧���ۼ� As Double, ByRef dbl����֧���ۼ� As Double, ByRef dbl�ǻ���ҽ�Ʒ� As Double) As Long

'===============================================================================================================
'����: ���߷����ύ����(�����շѡ�סԺ������)
'��ڲ���: �����ʻ�֧��(8λ,2λС��),�ֽ�֧��(8λ,2λС��)
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function CommitDataX Lib "YhYbClient.dll" (ByVal dbl����֧�� As Double, _
    ByVal dbl�ֽ�֧�� As Double) As Long

'===============================================================================================================
'����: ���������ύ����
'��ڲ���: ��
'���ڲ���: ��
'����: 0�ɹ�,-1ʧ��
'===============================================================================================================
Public Declare Function CommitData Lib "YhYbClient.dll" () As Long

Private Function Get���״���(ByVal intType As ҵ������_����, Optional bln������ As Boolean = False) As String
    Select Case intType
        Case ��ʼ���������
            Get���״��� = IIf(bln������, "��ʼ���������", "01")
        Case ����
            Get���״��� = IIf(bln������, "����", "02")
        Case ȡҩƷ��Ϣ
            Get���״��� = IIf(bln������, "ȡҩƷ��Ϣ", "03")
        Case ȡ������Ϣ
            Get���״��� = IIf(bln������, "ȡ������Ϣ", "04")
        Case ȡ������Ϣ
            Get���״��� = IIf(bln������, "ȡ������Ϣ", "05")
        Case ���÷�Ʊ����
            Get���״��� = IIf(bln������, "���÷�Ʊ����", "06")
        Case ���������վ���ϸ
            Get���״��� = IIf(bln������, "���������վ���ϸ", "07")
        Case �������������Ϣ
            Get���״��� = IIf(bln������, "�������������Ϣ", "08")
        Case ȡ�������վݼ�����Ϣ
            Get���״��� = IIf(bln������, "ȡ�������վݼ�����Ϣ", "09")
        Case ȡ������
            Get���״��� = IIf(bln������, "ȡ������", "10")
        Case ��Ժ�Ǽ�
            Get���״��� = IIf(bln������, "��Ժ�Ǽ�", "11")
        Case ȡ����Ժ�Ǽ�
            Get���״��� = IIf(bln������, "ȡ����Ժ�Ǽ�", "12")
        Case ��Ժ�Ǽ�
            Get���״��� = IIf(bln������, "��Ժ�Ǽ�", "13")
        Case ȡ����Ժ�Ǽ�
            Get���״��� = IIf(bln������, "ȡ����Ժ�Ǽ�", "14")
        Case ����סԺ�ʵ�
            Get���״��� = IIf(bln������, "����סԺ�ʵ�", "15")
        Case ���ü��ʵ���ϸ����
            Get���״��� = IIf(bln������, "���ü��ʵ���ϸ����", "16")
        Case ���ý��㵥
            Get���״��� = IIf(bln������, "���ý��㵥", "17")
        Case ����סԺ������Ϣ
            Get���״��� = IIf(bln������, "����סԺ������Ϣ", "18")
        Case ȡ��סԺ���������Ϣ
            Get���״��� = IIf(bln������, "ȡ��סԺ���������Ϣ", "19")
        Case ���������ύ
            Get���״��� = IIf(bln������, "���������ύ", "20")
        Case ���߷����ύ
            Get���״��� = IIf(bln������, "���߷����ύ", "21")
        Case �����������
            Get���״��� = IIf(bln������, "�����������", "22")
        Case Else
            Get���״��� = IIf(bln������, "����Ľ��״���", "-1")
    End Select
End Function

Public Function CheckReturn����() As Boolean
    CheckReturn���� = True
    If glngReturn = -1 Then
        MsgBox "�ڽ���ҽ������ʱ��ҽ���������´���" & vbCrLf & "    " & GetErrMsg(), vbInformation, "�ӿڴ���"
        CheckReturn���� = False
    End If
End Function

Public Sub delArrar()
    Dim iLoop As Long
    For iLoop = 0 To 50
        strTempArr(iLoop, 0) = ""
        strTempArr(iLoop, 1) = "0"
    Next
End Sub

Public Sub setArrar(str���� As String, dbl���� As Double)
    Dim iLoop As Long
    For iLoop = 0 To 50
        If strTempArr(iLoop, 0) = str���� Then
            strTempArr(iLoop, 1) = CLng(strTempArr(iLoop, 1)) + dbl����
            Exit Sub
        ElseIf strTempArr(iLoop, 0) = "" Then
            strTempArr(iLoop, 0) = str����
            strTempArr(iLoop, 1) = dbl����
            Exit Sub
        End If
    Next
End Sub

Public Function ҽ����ʼ��_����() As Boolean
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    If mblnInit = True Then
        ҽ����ʼ��_���� = True
        Exit Function
    End If
    DebugTool "����ҽ����ʼ���ӿ�"
    
    '���˺�:20040923����
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_����.ģ������ = True
    Else
        InitInfor_����.ģ������ = False
    End If
    
    
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_����
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����"
    
    InitInfor_����.��ϸʱʵ�ϴ� = False
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "��ϸʱʵ�ϴ�"
                InitInfor_����.��ϸʱʵ�ϴ� = Nvl(!����ֵ, 1) = 1
            End Select
            .MoveNext
        Loop
    End With
    
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_����)
    InitInfor_����.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
    mblnInit = True
    ҽ����ʼ��_���� = True
    DebugTool "ҽ����ʼ���ӿڳɹ�"
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Err = 0
    On Error GoTo errHand:
    ��ݱ�ʶ_���� = frmIdentify����.GetPatient(bytType, lng����ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_���� = ""
End Function
Private Function IS�Ƿ�ˢ������(ByVal lng����ID As Long) As Boolean
    '�жϵ�ǰ�Ĳ����Ƿ�ˢ���Ĳ���
    Dim rsTemp As New ADODB.Recordset
    IS�Ƿ�ˢ������ = False
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select * from �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽ��������Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�����ڵ�ǰ��ҽ������"
        Exit Function
    End If
    If g�������_����.ҽ���� <> Trim(Nvl(rsTemp!ҽ����)) Then
        ShowMsgbox "���д���,���ǵ�ǰ���˵�.��ȷ������Ŀ��Ƿ��ȷ!"
        Exit Function
    End If
    IS�Ƿ�ˢ������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function �������_����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_����)
    
    If rsTemp.EOF Then
        �������_���� = 0
    Else
        �������_���� = rsTemp("�ʻ����")
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency, Optional ByRef strAdvance As String) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strArr
    Dim dbl��ҩ���� As Double, dbl��� As Double, strҽ�����ִ��� As String, str�Ƿ�ҩƷ As String, str���㷽ʽ As String
    Dim StrInput As String, strOutput As String
    Dim iLoop As Integer
    Dim blnOld As Boolean '�Ƿ���Ҫ��дУ���ֶ�
    On Error GoTo errHandle
    
    gstrSQL = " " & _
        "  Select Rownum ��ʶ��,A.ID,A.����ID,A.�շ�ϸĿid,A.NO,A.���,A.����Ա����,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ��,A.������ as ҽ��,H.��� as ҽ�����, " & _
        "      A.����,A.����*A.���� as ����,A.�Ƿ��ϴ�,A.���㵥λ,B.���,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ�� as ʵ�ս��, " & _
        "      A.ҽ�����,A.�շ����,B.���� as ��Ŀ����,B.���� as ��Ŀ����,decode(J.��ʶ��,null,B.��ʶ����||B.��ʶ����,nvl(J.��ʶ��,' ')) as ���ұ���, " & _
        "      D.��Ŀ���� ҽ������,D.��Ŀ���� as ҽ������,J.���� as ����,D.�Ƿ�ҽ��,C.���� ��������,E.���� �ܵ�����, " & _
        "      L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.��λ����,L.˳���,L.����֤��,L.�ʻ����,L.��ǰ״̬,L.����ID,L.��ְ,L.�����,L.�Ҷȼ�,L.����ʱ�� " & _
        "  From (Select * From ������ü�¼ Where nvl(ʵ�ս��,0)<>0 and  ��¼״̬<>0 and ����ID=[2] and  Nvl(���ӱ�־,0)<>9 ) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,���ű� E,  " & _
        "       (Select distinct Q.ҩƷid,Q.��ʶ��,T.���� From ҩƷĿ¼ Q,ҩƷ��Ϣ R,ҩƷ���� T  Where  Q.ҩ��id=R.ҩ��id and R.����=T.���� ) J, " & _
        "       ��Ա�� H,�����ʻ� L" & _
        "  Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID(+)  and A.����id=L.����id  and L.����=[1] and a.�շ�ϸĿid=J.ҩƷid(+) " & _
        "        and A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=[1] and a.������=H.����(+) " & _
        "  Order by A.ID"
                        
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If
    dbl��ҩ���� = 0
    dbl��� = 0
    While Not rs��ϸ.EOF
        dbl��� = dbl��� + rs��ϸ!ʵ�ս��
        If Nvl(rs��ϸ!�շ����) = "6" Or Nvl(rs��ϸ!�շ����) = "7" Then
            dbl��ҩ���� = dbl��ҩ���� + Nvl(rs��ϸ!����, 0)
        End If
        rs��ϸ.MoveNext
    Wend
    
    rs��ϸ.MoveFirst
    If dbl��� = 0 Then
        Err.Raise 9000, gstrSysName, "����û�з�������,���ܽ���ҽ������"
        Exit Function
    End If
    
    'ȡ����ID�Ͳ���Ա
    lng����ID = rs��ϸ!����ID
    
    gstrSQL = "Select * From ���ղ��� Where ID=" & Nvl(rs��ϸ!����ID, 0) & " And ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����)
        
    '���˺�:20040923����û�в���
    If rsTemp.EOF Then
        strҽ�����ִ��� = ""
    Else
        strҽ�����ִ��� = Substr(rsTemp!����, 1, 10)
    End If
    
    'If ҵ������_����(��ʼ���������, "10", strOutPut) = False Then Exit Function
    '���˺�:������ˢһ�ο�
    If ��ݼ���_����(0, "10") = False Then
        Exit Function
    End If
    If IS�Ƿ�ˢ������(lng����ID) = False Then
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
        Exit Function
    End If

    
    '���÷�Ʊ����
    StrInput = 1
    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!NO), 1, 13)
    StrInput = StrInput & vbTab & g�������_����.��������
    '�º�����20060512�޸�
    StrInput = StrInput & vbTab & Substr(g�������_����.���ִ���, 1, 10)
    StrInput = StrInput & vbTab & Substr(g�������_����.�������, 1, 200)
    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!��������), 1, 20)
    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!ҽ��), 1, 10)
    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!����Ա����), 1, 10)
    StrInput = StrInput & vbTab & Substr(dbl��ҩ����, 1, 2)
    StrInput = StrInput & vbTab & dbl���
    If ҵ������_����(���÷�Ʊ����, StrInput, strOutput) = False Then Exit Function
    
    delArrar            '������ô����¼����
    
    '���÷�Ʊ��ϸ���ݣ�SetClinicBillDetail
    Do While Not rs��ϸ.EOF
        If Nvl(rs��ϸ!ҽ������, "") = "" Then
            Err.Raise 9000, gstrSysName, "��Ŀ[" & Nvl(rs��ϸ!��Ŀ����) & "]δ���ö�Ӧ��ҽ����Ŀ,��������ҽ��"
            Exit Function
        End If
        
        If Nvl(rs��ϸ!�Ƿ��ϴ�, 0) = 0 And rs��ϸ!ʵ�ս�� <> 0 Then
            str�Ƿ�ҩƷ = "0"
            StrInput = Substr(Nvl(rs��ϸ!ҽ������), 1, 10)
            
            Select Case UCase(Nvl(rs��ϸ!�շ����))
                Case "5", "6", "7"
                    str�Ƿ�ҩƷ = "1"
                    'aItemName����Ŀ����(40λ)
                    'aMediKindCode�����ô���(2λ)
                    'aIsCityRich��(0-�׻���ͨ1-�һ�߾���)(1λ)[����2�Է�]
                    'aIsCityMedi���Ƿ�ҽ��(1λ)(0-����1-��)[Ϊ��ɽӿڼ��ݶ�����]
                    'aIsLimit���Ƿ�ҽ���޼���Ŀ(1λ)(0-����1-��)
                    'aCitySelfPayRate���Ը�����(4λ��3λС��)
                    'aPrice����׼����(8λ��2λС��)
                    
                    If ҵ������_����(ȡҩƷ��Ϣ, StrInput, strOutput) = False Then Exit Function
                Case "J", "H", "I"
                    If ҵ������_����(ȡ������Ϣ, StrInput, strOutput) = False Then Exit Function
                    str�Ƿ�ҩƷ = "2"
                Case Else
                    If ҵ������_����(ȡ������Ϣ, StrInput, strOutput) = False Then Exit Function
            End Select
            If strOutput = "" Then Exit Function
            strArr = Split(strOutput, vbTab)
            
            '��¼���ô�����
            setArrar CStr(strArr(1)), Nvl(rs��ϸ!ʵ�ս��, 0)
            
            '���ýӿ�,д����ϸ
            'aInvoiceHandle: [Ϊ��ɽӿڼ��ݶ�����δ��]
            StrInput = "1"
            'aCityMediCareNo��ҽ����Ŀ��š�(10λ)not null
            StrInput = StrInput & vbTab & Nvl(rs��ϸ!ҽ������, "")
            'aItemName��ҽԺ��Ŀ����(40λ)not null
            StrInput = StrInput & vbTab & Nvl(rs��ϸ!��Ŀ����, "")
            'aConformationName����������(20λ)
            StrInput = StrInput & vbTab & Nvl(rs��ϸ!����, "")
            'aUnitContent����λ����(14λ)
            StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!���, ""), 1, 14)
                '���˺�:��������
                'gstrSQL = "Select ����,Ƶ��,�÷� From ҩƷ�շ���¼ where ����id=" & Nvl(!ID, 0)
                'zlDataBase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵���Ƶ��"
                'If Not rsTemp.EOF Then
                'strInput = strInput & vbTab & Substr("����:" & Nvl(rsTemp!����, "") & " Ƶ��:" & Nvl(rsTemp!Ƶ��) & "�÷�:" & Nvl(rsTemp!�÷�), 1, 14)
                'Else
                'strInput = strInput & vbTab & ""
                'End If
            'aDosage���÷�����(40λ)
            StrInput = StrInput & vbTab & ""
            'aMediKindCode�����ô������(2λ)not null
            StrInput = StrInput & vbTab & strArr(1)
            'aIsRich��(0-�׻���ͨ1-�һ�߾���)(1λ)[����2�Է�]not null
            StrInput = StrInput & vbTab & strArr(2)
            'aIsCityMedi���Ƿ�ҽ��(1λ)(0-����1-��)[Ϊ��ɽӿڼ��ݶ�����δ��]
            StrInput = StrInput & vbTab & strArr(3)
            'aIsMedi:�Ƿ�ҩƷ(0-��Ŀ1-ҩƷ2������ʩ[��] )not null
            StrInput = StrInput & vbTab & str�Ƿ�ҩƷ
            'aPrice������(8λ��2λС��)>0
            StrInput = StrInput & vbTab & Nvl(rs��ϸ!ʵ�ʼ۸�, 0)
            'aQuantity������(8λ��2λС��)>0
            StrInput = StrInput & vbTab & Nvl(rs��ϸ!����, 0)
            'aAmount�����(8λ��2λС��)>0
            StrInput = StrInput & vbTab & Nvl(rs��ϸ!ʵ�ս��, 0)
            
            If ҵ������_����(���������վ���ϸ, StrInput, strOutput) = False Then Exit Function
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        End If
        rs��ϸ.MoveNext
    Loop
    
    For iLoop = 0 To 50
        If strTempArr(iLoop, 0) <> "" Then
            StrInput = "1"
            StrInput = StrInput & vbTab & strTempArr(iLoop, 0)
            StrInput = StrInput & vbTab & strTempArr(iLoop, 1)
            If ҵ������_����(�������������Ϣ, StrInput, strOutput) = False Then Exit Function
        Else
            Exit For
        End If
    Next
        
    If ҵ������_����(ȡ�������վݼ�����Ϣ, "1", strOutput) = False Then Exit Function
    strArr = Split(strOutput, vbTab)
    
    
    If Val(strArr(1)) <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & Format(Val(strArr(1)), "####0.00;-####0.00; ;")
    End If
    
    If Val(strArr(4)) <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||ͳ��֧��|" & Format(Val(strArr(4)), "####0.00;-####0.00; ;")
    End If
    
    '�������
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        #If gverControl < 2 Then
            blnOld = True
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
        #Else
            strAdvance = str���㷽ʽ
            gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    
    If blnOld Then
        If frm������Ϣ.ShowME(lng����ID, True) = False Then
            Exit Function
        End If
    End If

   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(�Ը����ۼ�),�ʻ��ۼ�֧��_IN(ͳ��֧���ۼ�),�ۼƽ���ͳ��_IN(�ز�֧���ۼ�),�ۼ�ͳ�ﱨ��_IN(����֧���ۼ�),סԺ����_IN,����(�ǻ���ҽ�Ʒ�),�ⶥ��_IN(�����ֽ�֧��),ʵ������_IN(���˱�������),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(�Ը���֧��),�����Ը����_IN(�Է�ҩƷ),
    '   ����ͳ����_IN(����ҩƷ),ͳ�ﱨ�����_IN(ͳ��֧��),    ���Ը����_IN(�չ�֧��),�����Ը����_IN(�չ˵渶),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(��������),��ҳID_IN,��;����_IN,��ע_IN
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
    
    With g�������_����
        gstrSQL = "zl_���ս����¼_insert( 1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          Val(strArr(19)) & "," & Val(strArr(20)) & "," & Val(strArr(21)) & "," & Val(strArr(22)) & ",NULL," & Val(strArr(23)) & "," & Val(strArr(2)) & "," & Val(strArr(3)) & "," & _
         dbl��� & "," & Val(strArr(7)) & "," & Val(strArr(10)) & "," & _
          Val(strArr(9)) & "," & Val(strArr(4)) & "," & Val(strArr(5)) & "," & Val(strArr(6)) & "," & Val(strArr(1)) & ",'" & _
          .�������� & "',Null,Null,NULl" & IIf(blnOld, "", ",1") & ")"
              
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
        '������Ϣ����:
            '����_IN ,��¼ID_IN,�ʻ����_IN ,�̱�֧��_IN ,����ҩƷ_IN ,��������_IN ,�Է�����_IN ,��������_IN ,������ʩ_IN ,�Է���ʩ_IN ,������ʩ_IN ,�����Է�_IN
        gstrSQL = "zl_���ս����¼_������Ϣ( 1," & lng����ID & "," & Val(strArr(0)) & "," & Val(strArr(8)) & "," & Val(strArr(11)) & "," & Val(strArr(12)) & "," & Val(strArr(13)) & "," & Val(strArr(14)) & "," & Val(strArr(15)) & "," & Val(strArr(16)) & "," & Val(strArr(17)) & "," & Val(strArr(18)) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼������Ϣ")
    End With

    StrInput = strArr(1)    '�����ʻ�֧��
    StrInput = StrInput & vbTab & strArr(2) '�ֽ�֧��
    
    If ҵ������_����(���߷����ύ, StrInput, strOutput) = False Then Exit Function
    If ҵ������_����(�����������, "", strOutput) = False Then Exit Function


    �������_���� = True
    DebugTool "�������ɹ�"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, rs��ϸ As New ADODB.Recordset
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curƱ���ܽ�� As Currency, int���� As Integer, datCurr As Date
    Dim StrInput As String, strOutput As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select *  From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rs��ϸ.EOF Then
        Err.Raise 9000, gstrSysName, "û���ҵ�ԭ���ݵ���ϸ���ݣ����ܽ��г���", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do While Not rs��ϸ.EOF
        curƱ���ܽ�� = curƱ���ܽ�� + Nvl(rs��ϸ("���ʽ��"), 0)
        rs��ϸ.MoveNext
    Loop
    
    rs��ϸ.MoveFirst
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp("����ID")
    

    '���ýӿ�������
    If ��ݼ���_����(0, "11") = False Then Exit Function
    If IS�Ƿ�ˢ������(lng����ID) = False Then
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
        Exit Function
    End If
    
    StrInput = "1"
    StrInput = StrInput & vbTab & Nvl(rs��ϸ!NO) & "R"
    StrInput = StrInput & vbTab & Nvl(rs��ϸ!NO)
    If ҵ������_����(ȡ������, StrInput, strOutput) = False Then Exit Function
    If ҵ������_����(���������ύ, "", strOutput) = False Then Exit Function
    If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
    
    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(�Ը����ۼ�),�ʻ��ۼ�֧��_IN(ͳ��֧���ۼ�),�ۼƽ���ͳ��_IN(�ز�֧���ۼ�),�ۼ�ͳ�ﱨ��_IN(����֧���ۼ�),סԺ����_IN,����(�ǻ���ҽ�Ʒ�),�ⶥ��_IN(�����ֽ�֧��),ʵ������_IN(���˱�������),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(�Ը���֧��),�����Ը����_IN(�Է�ҩƷ),
    '   ����ͳ����_IN(����ҩƷ),ͳ�ﱨ�����_IN(ͳ��֧��),    ���Ը����_IN(�չ�֧��),�����Ը����_IN(�չ˵渶),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(),��ҳID_IN,��;����_IN,��ע_IN
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
    
    gstrSQL = "Select * From ���ս����¼ Where ��¼ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "û���ҵ�ԭ�еı��ս����¼", vbInformation, gstrSysName
        Exit Function
    End If
        
    With g�������_����
        gstrSQL = "zl_���ս����¼_insert( 1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          -1 * Nvl(rsTemp!�ʻ��ۼ�����, 0) & "," & -1 * Nvl(rsTemp!�ʻ��ۼ�֧��, 0) & "," & -1 * Nvl(rsTemp!�ۼƽ���ͳ��, 0) & "," & -1 * Nvl(rsTemp!�ۼ�ͳ�ﱨ��, 0) & ",NULL," & -1 * Nvl(rsTemp!����, 0) & "," & -1 * Nvl(rsTemp!�ⶥ��, 0) & "," & -1 * Nvl(rsTemp!ʵ������, 0) & "," & _
         -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & _
          -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & "," & -1 * Nvl(rsTemp!���Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & _
          rsTemp!֧��˳��� & "',Null,Null,NULl)"
              
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
        '������Ϣ����:
        '����_IN ,��¼ID_IN,�ʻ����_IN ,�̱�֧��_IN ,����ҩƷ_IN ,��������_IN ,�Է�����_IN ,��������_IN ,������ʩ_IN ,�Է���ʩ_IN ,������ʩ_IN ,�����Է�_IN
        gstrSQL = "zl_���ս����¼_������Ϣ( 1," & lng����ID & "," & -1 * Nvl(rsTemp!�ʻ����, 0) & "," & -1 * Nvl(rsTemp!�̱�֧��, 0) & "," & -1 * Nvl(rsTemp!����ҩƷ, 0) & "," & -1 * Nvl(rsTemp!��������, 0) & "," & -1 * Nvl(rsTemp!�Է�����, 0) & "," & -1 * Nvl(rsTemp!��������, 0) & "," & -1 * Nvl(rsTemp!������ʩ, 0) & "," & -1 * Nvl(rsTemp!�Է���ʩ, 0) & "," & -1 * Nvl(rsTemp!������ʩ, 0) & "," & -1 * Nvl(rsTemp!�����Է�, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼������Ϣ")
    End With

    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
        
    
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ�����,���Ƚ��н���!"
        Exit Function
    End If
    
    
    On Error GoTo errHand:
    If ��ݼ���_����(1, "20") = False Then Exit Function
    
    '��ȡ��ز�����Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����,L.���� as ����,A.��ǰ����id,to_char(A.ȷ������,'yyyyMMdd') as ȷ������,A.����ҽʦ,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����ʱ��," & _
        " to_char(A.��Ժ����,'yyyyMMdd') ��Ժ����,to_char(A.��Ժ����,'ss') as ��� ,to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժʱ��,D.��Ժ��� " & _
        " From ������ҳ A,���ű� B,���ű� L,������Ϣ C, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ��� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by  ����id,��ҳid)   D" & _
        " Where A.����id=C.����id and a.��ǰ����ID=L.iD(+) and C.����id=" & lng����ID & _
        "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
        ""
    If g�������_����.������� = "" Then
        ShowMsgbox "û������������,������ݴ���������!"
        Exit Function
    End If
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ��Ϣ"
    
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    'aPrnHandle�����˾������ReadCard������÷��ء�
    StrInput = "1"
    'aInHosNo��סԺ��(13λ)not null
    StrInput = StrInput & vbTab & lng����ID & "-" & lng��ҳID & "-" & Nvl(rsTemp!���)
    'aInHosDate����Ժ����(8λ)(YYYYMMDD)not null
    StrInput = StrInput & vbTab & Nvl(rsTemp!��Ժ����)
    'aDepartmentName��סԺ����(20λ)
    StrInput = StrInput & vbTab & Nvl(rsTemp!��Ժ����)
    'aSickArea������(20λ)
    StrInput = StrInput & vbTab & Nvl(rsTemp!����)
    
    gstrSQL = "Select * From ��λ״����¼ D where ����ID=[1] And ����=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ��Ϣ", CLng(Nvl(rsTemp!��ǰ����ID, 0)), CLng(Nvl(rsTemp!��ǰ����)))
    If rsData.EOF Then
        'aRoom������(10λ)
        StrInput = StrInput & vbTab & ""
        'aBedNo������(3λ)
        StrInput = StrInput & vbTab & ""
    Else
        'aRoom������(10λ)
        StrInput = StrInput & vbTab & Nvl(rsData!�����)
        'aBedNo������(3λ)
        StrInput = StrInput & vbTab & Right(Nvl(rsData!����), 3)
    End If
    'aClinicDoctorCode������ҽ��(10λ)
    StrInput = StrInput & vbTab & Nvl(rsTemp!����ҽʦ)
    'aInHosDiagnoseCode����Ժ��ϴ���(����)(10λ)not null
    'strInput = strInput & vbTab & Substr(g�������_����.���ִ���, 1, 10)
    StrInput = StrInput & vbTab & Substr(g�������_����.���ִ���, 1, 10)
    StrInput = StrInput & vbTab & Substr(g�������_����.�������, 1, 200)
    
    If ҵ������_����(��Ժ�Ǽ�, StrInput, strOutput) = False Then Exit Function
    If ҵ������_����(���������ύ, "", strOutput) = False Then Exit Function
    If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
    
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function
Public Function ҵ������_����(ByVal intType As ҵ������_����, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strReturn As String
    Dim strOutput(0 To 20) As String, dblOutPut(0 To 25) As Double, intOutPut(0 To 5) As Integer, lngOutPut(0 To 5) As Long
    Dim strArr1
    Dim strArr(0 To 20) As String
    Dim strReg As String
    Dim str���� As String
    
    Dim i As Integer
    str���� = Get���״���(intType, True)
    DebugTool "����ҵ��������(ҵ������Ϊ:" & intType & " ҵ������:" & str���� & ")," & vbCrLf & "   �������Ϊ" & strInputString
    
    ҵ������_���� = False
    
    StrInput = strInputString
    
    If InitInfor_����.ģ������ Then
        '��ȡģ������
        Readģ������ intType, strInputString, strOutPutstring
         ҵ������_���� = True
        Exit Function
    End If
   
    strArr1 = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr1)
        strArr(i) = strArr1(i)
    Next
        
    
    Err = 0
    On Error GoTo errHand:
    
    Select Case intType
        Case ����
           lngReturn = ReadCard(strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), strOutput(5), strOutput(6), strOutput(7), strOutput(8), strOutput(9), strOutput(10), strOutput(11), strOutput(12), strOutput(13), strOutput(14), strOutput(15), strOutput(16), strOutput(17), dblOutPut(0), dblOutPut(1), dblOutPut(2), dblOutPut(3), dblOutPut(4), dblOutPut(5), lngOutPut(0), lngOutPut(1))
           If lngReturn < 0 Then
                ShowMsgbox "�ڽ���ҽ������ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                Exit Function
           End If
           '�������ش�
           strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & strOutput(5) & vbTab & strOutput(6) & vbTab & strOutput(7) & vbTab & strOutput(8) & vbTab & strOutput(9) & vbTab & strOutput(10) & vbTab & strOutput(11) & vbTab & strOutput(12) & vbTab & strOutput(13) & vbTab & strOutput(14) & vbTab & strOutput(15) & vbTab & strOutput(16) & vbTab & strOutput(17) & vbTab & dblOutPut(0) & vbTab & dblOutPut(1) & vbTab & dblOutPut(2) & vbTab & dblOutPut(3) & vbTab & dblOutPut(4) & vbTab & dblOutPut(5) & vbTab & lngOutPut(0) & vbTab & lngOutPut(1)
        Case ��ʼ���������
            '�����շ�(10),�����˷�(11),��Ժ�Ǽ�(20),ҽ��¼��(21),סԺ����(22),סԺ����(23),��Ժ�Ǽ�(24),ȡ����Ժ�Ǽ�(25),����סԺ����(26),ȡ������(27),ȡ����Ժ�Ǽ�(28),ȡ��ҽ��¼��(29),����
            lngReturn = InitCalc(strArr(0))
           If lngReturn < 0 Then
                ShowMsgbox "�ڽ���ҽ����ʼ������ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                Exit Function
           End If
        Case ��Ժ�Ǽ�
           lngReturn = InhosRegister(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9))
           If lngReturn < 0 Then
                ShowMsgbox "�ڽ�����Ժ�Ǽ�ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                Exit Function
           End If
        Case ȡ����Ժ�Ǽ�
           lngReturn = uInhosRegister(0)
           If lngReturn < 0 Then
                ShowMsgbox "�ڽ���ȡ����Ժ�Ǽ�ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                Exit Function
           End If
        Case ��Ժ�Ǽ�
           lngReturn = OuthosRegister(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), Val(strArr(9)), strArr(10))
           If lngReturn < 0 Then
                ShowMsgbox "�ڽ��г�Ժ�Ǽ�ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                Exit Function
           End If
            
        Case ���������ύ
           lngReturn = CommitData()
           If lngReturn < 0 Then
                ShowMsgbox "�ڽ��з��������ύʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                Exit Function
           End If
        Case �����������
           lngReturn = FinalCalc()
           If lngReturn < 0 Then
                ShowMsgbox "�ڽ��н����������ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                Exit Function
           End If
        Case ���÷�Ʊ����
            lngReturn = SetClinicBill(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), Val(strArr(8)), Val(strArr(9)))
            If lngReturn < 0 Then
                ShowMsgbox "�ڽ������÷�Ʊ����ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                Exit Function
           End If
        Case ȡҩƷ��Ϣ
            lngReturn = GetInfo_MediDic(strArr(0), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), dblOutPut(0), dblOutPut(1))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡҩƷ��Ϣʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
           strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & dblOutPut(0) & vbTab & dblOutPut(1)
           
        Case ȡ������Ϣ
            lngReturn = GetInfo_ServerDic(strArr(0), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), dblOutPut(0), dblOutPut(1))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡ������Ϣʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & dblOutPut(0) & vbTab & dblOutPut(1)
                   
        Case ȡ������Ϣ
            lngReturn = GetInfo_ItemDic(strArr(0), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), dblOutPut(0), dblOutPut(1))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡ������Ϣʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & dblOutPut(0) & vbTab & dblOutPut(1)
        Case ���������վ���ϸ
            lngReturn = SetClinicBillDetail(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), Val(strArr(10)), Val(strArr(11)), Val(strArr(12)))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡ������Ϣʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case �������������Ϣ
            lngReturn = SetClinicMediKind(Val(strArr(0)), strArr(1), Val(strArr(2)))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ����������������Ϣʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case ȡ�������վݼ�����Ϣ
            lngReturn = GetClinicBillData(1, dblOutPut(0), dblOutPut(1), dblOutPut(2), dblOutPut(3), dblOutPut(4), dblOutPut(5), dblOutPut(6), dblOutPut(7), dblOutPut(8), dblOutPut(9), dblOutPut(10), dblOutPut(11), dblOutPut(12), dblOutPut(13), dblOutPut(14), dblOutPut(15), dblOutPut(16), dblOutPut(17), dblOutPut(18), dblOutPut(19), dblOutPut(20), dblOutPut(21), dblOutPut(22), dblOutPut(23))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡ�������վݼ�����Ϣʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
            strReturn = ""
            For i = 0 To 23
                '�������ش�
                strReturn = strReturn & dblOutPut(i) & vbTab
            Next
        Case ���߷����ύ
            lngReturn = CommitDataX(Val(strArr(0)), Val(strArr(1)))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ������߷����ύʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case ȡ������
            lngReturn = ReturnCharge(Val(strArr(0)), strArr(1), strArr(2))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡ������ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case ����סԺ�ʵ�
            lngReturn = SetInHosBill(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), Val(strArr(5)), Val(strArr(6)))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡ������ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case ���ü��ʵ���ϸ����
            lngReturn = SetInHosBillDetail(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), Val(strArr(10)), Val(strArr(11)), Val(strArr(12)))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ������ü��ʵ���ϸ����ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case ȡ����Ժ�Ǽ�
            lngReturn = uOuthosRegister(strArr(0))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡ����Ժ�Ǽ�ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case ���ý��㵥
            lngReturn = SetCheckOutBill(Val(strArr(0)), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), Val(strArr(9)), Val(strArr(10)))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ������ý��㵥ʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case ����סԺ������Ϣ
            lngReturn = SetInHosMediKind(Val(strArr(0)), strArr(1), Val(strArr(2)))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ�������סԺ������Ϣʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
        Case ȡ��סԺ���������Ϣ
            lngReturn = GetCheckOutBillData(1, dblOutPut(0), dblOutPut(1), dblOutPut(2), dblOutPut(3), dblOutPut(4), dblOutPut(5), dblOutPut(6), dblOutPut(7), dblOutPut(8), dblOutPut(9), dblOutPut(10), dblOutPut(11), dblOutPut(12), dblOutPut(13), dblOutPut(14), dblOutPut(15), dblOutPut(16), dblOutPut(17), dblOutPut(18), dblOutPut(19), dblOutPut(20), dblOutPut(21), dblOutPut(22), dblOutPut(23))
            If lngReturn < 0 Then
                 ShowMsgbox "�ڽ���ȡ��סԺ���������Ϣʱ�������´���" & vbCrLf & "�����:" & lngReturn & vbCrLf & "��������:" & GetErrMsg()
                Call ҵ������_����(�����������, "", "")
                 Exit Function
            End If
            strReturn = ""
            For i = 0 To 23
                '�������ش�
                strReturn = strReturn & dblOutPut(i) & vbTab
            Next
    End Select
    strOutPutstring = strReturn
    ҵ������_���� = True
    DebugTool "     �������Ϊ:" & strReturn
    DebugTool "ҵ������ɹ�(ҵ������Ϊ:" & intType & " ҵ������:" & str���� & ")"
     Exit Function
errHand:
    DebugTool "ҵ������ʧ��(ҵ������Ϊ:" & intType & " ҵ������:" & str���� & ")"
    If ErrCenter = 1 Then
        Resume
    End If
End Function
    
    
    



Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    '���˺�:20040923���ӵ�
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strҽ����  As String
    
    Err = 0
    On Error GoTo errHand
    
    DebugTool "������Ժ�ǳ����ӿ�"
    
    ��Ժ�Ǽǳ���_���� = False
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ����ã����ܳ�����Ժ�Ǽ�"
        Exit Function
    End If
    
    '�ȳ�ʼ������
    If ҵ������_����(��ʼ���������, "25", strOutput) = False Then Exit Function
    
    '���������
    If ҵ������_����(����, "", strOutput) = False Then Exit Function
    If ҵ������_����(ȡ����Ժ�Ǽ�, "", strOutput) = False Then Exit Function
    If ҵ������_����(���������ύ, "", strOutput) = False Then Exit Function
    If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
    DebugTool "����ҽ����ȡ��ҵ��ɹ�,����ʼ���±����ʻ������״̬��"
    
    '����ҽ���ʻ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    
    DebugTool "ȡ���ɹ�"
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Dim datCurr As Date, bln����ó�Ժ As Boolean, strסԺ�� As String, _
        strInNote As String, str���ֱ��� As String, strҽ���� As String
    Dim StrInput As String, strOutput As String
    
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    '���˺�:20040924����
    bln����ó�Ժ = Not ����δ�����(lng����ID, lng��ҳID)
    
    
    
    If bln����ó�Ժ = True Then
        '���˺�:20040924����
       If ��Ժ�Ǽǳ���_����(lng����ID, lng��ҳID) = True Then
            ��Ժ�Ǽ�_���� = True
       End If
        Exit Function
    End If
        
    '��ȷ���Ƿ��Ѿ����ڳ�Ժ����
    Dim str��Ժ���� As String, str����֢ As String
    
Go����:
    If frm����ѡ��_����.ShowSelect(TYPE_����, lng����ID, lng��ҳID, str��Ժ����, str����֢) = False Then Exit Function
    
    gstrSQL = "" & _
        "   Select a.*,b.����,b.����,a.����֢ From �����ʻ� a,���ղ��� b where a.��Ժ����ID=b.ID(+) and a.����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ����"
    str��Ժ���� = Nvl(rsTemp!����)
    If Nvl(rsTemp!����) = "" Or Nvl(rsTemp!����֢) = "" Then
           If MsgBox("����û�в��ֻ򲢷�֢�����Բ��ܳ�Ժ�Ǽ�,�Ƿ�����¼��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                GoTo Go����:
            Else
                Exit Function
            End If
    End If
    
        
    If ҵ������_����(��ʼ���������, "24", strOutput) = False Then Exit Function
    
    '��ȡ��Ժ���
    'strInNote = ��ȡ���Ժ���(lng����id, lng��ҳID, False, False, True)
    
    '��ȡסԺҽʦ
    gstrSQL = "" & _
        "   Select A.��Ժ����,(sysdate-a.��Ժ����)/365 as סԺ����,b.��ǰ����,B.סԺ��," & _
        "           to_char(A.��Ժ����,'ss') as ���,A.��ǰ����id,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
        "           C.����,G.��Ժ����,D.���� As ���ұ���,J.���� as ����,A.��Ժ��ʽ " & _
        "   from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D,���ű� J, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ���� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� = 3 and a.��ҳid=[2] and a.����id=[1] Group by ����id,��ҳid)   G" & _
        "   Where   A.����ID = B.����ID And A.����ID = C.����ID And A.��ǰ����ID=J.id(+) and " & _
        "           A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]" & _
        "           and A.��ҳid=G.��ҳid(+) and a.����id=G.����id(+) " & _
        ""
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    
    If rsTemp.EOF Then
        MsgBox "����ȡ�ò��˵���Ժ�Ǽ���Ϣ", vbInformation, gstrSysName
        Exit Function
    End If
    gstrSQL = "Select * From ��λ״����¼ D where ����ID=[1] And ����=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ��Ϣ", CLng(Nvl(rsTemp!��ǰ����ID, 0)), CLng(Nvl(rsTemp!��ǰ����)))
    
    '���˺�:20040923��֤��סԺ��
    'strסԺ�� = Format(lng����ID, "0#########") & Format(lng��ҳID, "0##")
    'aInHosNo��סԺ��(13λ)not null
    StrInput = lng����ID & "-" & lng��ҳID & "-" & Nvl(rsTemp!���) & vbTab
    'aOutHosDate����Ժ����(8λ)(YYYYMMDD)not null
    StrInput = StrInput & Format(datCurr, "yyyymmdd") & vbTab
    'aDepartmentName��סԺ����(20λ)
    StrInput = StrInput & Substr(Nvl(rsTemp!סԺ����), 1, 20) & vbTab
    'aSickArea������(20λ)
    StrInput = StrInput & Substr(Nvl(rsTemp!����), 1, 20) & vbTab
    'aRoom������(10λ)
    'aBedNo������(3λ)
    If rsData.EOF Then
        StrInput = StrInput & "" & vbTab
        StrInput = StrInput & "" & vbTab
    Else
        StrInput = StrInput & Substr(Nvl(rsData!�����), 1, 10) & vbTab
        StrInput = StrInput & Right(Nvl(rsData!����), 3) & vbTab
    End If
    
    'aDoctorCode��ҽ��(����)(10λ)
    StrInput = StrInput & Substr(Nvl(rsTemp!סԺҽʦ), 1, 10) & vbTab
    'aoutHosDiagnoseCode����Ժ��ϴ���(����)(10λ)not null
    'strInput = strInput & Substr(g�������_����.���ִ���, 1, 10) & vbTab
    
    StrInput = StrInput & Substr(str��Ժ����, 1, 10) & vbTab
    '1-����2-��ת 3-δ�� 4-����9-����
    'aOutHosCure����Ժ�������(1-����2-��ת 3-δ�� 4-����9-����)(1λ)not null
     StrInput = StrInput & Substr(Get�������_����(lng����ID, lng��ҳID), 1, 1) & vbTab
    'aInHosDays��סԺ����(3λ)
    StrInput = StrInput & Substr(Int(Nvl(rsTemp!סԺ����, 0)), 1, 3) & vbTab
    'aoutHostYPE: ��Ժ����
    StrInput = StrInput & "1" & vbTab

    If ҵ������_����(��Ժ�Ǽ�, StrInput, strOutput) = False Then Exit Function
    If ҵ������_����(���������ύ, StrInput, strOutput) = False Then Exit Function
    If ҵ������_����(�����������, StrInput, strOutput) = False Then Exit Function
    
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
Private Function Get�������_����(lng����ID As Long, lng��ҳID As Long) As String
    '����:��ȡ���������ʶ
    '     A-������B-��ת��C-δ����D-������E-����
    '??49  ���������ʶ    CHAR    439 1   1������2��ת��3δ����4������5������סԺ���� Ժ��
    'A-������B-��ת��C-δ����D-������E-����
    
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.��Ժ���" & _
             " From ������ A,��������Ŀ¼ B " & _
             " Where A.����ID=[1] And A.����ID=B.ID(+) And A.��ҳID=[2]" & _
             "       And A.������� in (2,3)" & _
             " Order by A.������� Desc"
    Set rsInNote = zlDatabase.OpenSQLRecord(strTmp, "ҽ���ӿ�", lng����ID, lng��ҳID)
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = Nvl(rsInNote!��Ժ���)
    End If
    strTmp = Decode(strTmp, "����", "1", "��ת", "2", "δ��", "3", "����", "4", "����", "9", "1")
    Get�������_���� = strTmp
   Call WriteDebugInfor_����("Get�������_����", lng����ID)
End Function

Public Function ��Ժ�Ǽǳ���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '��Ժ�Ǽǳ���
    Dim StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    gstrSQL = "Select to_char(��Ժ����,'ss') as ��� From ������ҳ where ����id= " & lng����ID & " and ��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ�Ǽ�"
    If rsTemp.EOF Then Exit Function
    
    StrInput = lng����ID & "-" & lng��ҳID & "-" & rsTemp!���
    
    If ҵ������_����(��ʼ���������, "28", strOutput) = False Then Exit Function
    If ҵ������_����(ȡ����Ժ�Ǽ�, StrInput, strOutput) = False Then Exit Function
    If ҵ������_����(���������ύ, StrInput, strOutput) = False Then Exit Function
    If ҵ������_����(�����������, StrInput, strOutput) = False Then Exit Function
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, Optional ByRef strAdvance As String) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
    '        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
    '        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str���ִ��� As String, str���㷽ʽ As String, str����֤ As String
    Dim dbl��ҩ����  As Double, dbl��� As Double
    Dim lng��ҳID As Long
    Dim lng����ID As Long, iLoop As Integer
    Dim StrInput  As String, strOutput As String, str�Ƿ�ҩƷ As String
    Dim strArr
    Dim blnOld As Boolean '�Ƿ���Ҫ��дУ���ֶ�
    
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select ��ǰ״̬ from �����ʻ� where ����id=" & lng����ID
    

    gstrSQL = " " & _
        "        select a.ʵ�ս��,a.id,a.��¼����,a.��ҳid,a.��¼״̬,a.����ʱ��,a.�Ǽ�ʱ��,a.no,a.���˲���id,a.����,a.���,a.��ʶ�� as סԺ��,a.���˿���id,a.����id,a.�շ����,b.���,a.���㵥λ, " & _
        "               A.���㵥λ,A.����,A.����*Nvl(A.����,1) ����,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ�� ,a.������ as ҽ��,c.��� as ҽ�����, " & _
        "               a.ҽ�����,nvl(a.Ӥ����,0) as Ӥ����,to_char(F1.��Ժ����,'yyyyMMdd') as ��Ժ����,(F1.��Ժ����-F1.��Ժ����)/365 AS סԺ����,to_char(F1.��Ժ����,'yyyyMMDD') as ��Ժ����, A.ʵ�ս��,nvl(A.�Ƿ��ϴ�,0) as �Ƿ��ϴ�, " & _
        "               D.���� as ��Ŀ����,D.���� as ��Ŀ����,decode(J.��ʶ��,null,D.��ʶ����||D.��ʶ����,nvl(J.��ʶ��,' ')) as ���ұ���, " & _
        "               E.��Ŀ���� as ҽ������,E.��Ŀ���� as ҽ������,e.�Ƿ�ҽ��,e.����id,H.���� as ��������,J.���� as ����, " & _
        "               L.����,l.���� , l.����, l.ҽ����, l.��Ա���, l.��λ����, l.˳���, l.����֤��, l.�ʻ����, l.��ǰ״̬, l.����ID, l.��ְ, l.�����, l.�Ҷȼ�, l.����ʱ�� " & _
        "        from סԺ���ü�¼ a,�շ���� b,������ҳ F1,��Ա�� c,�շ�ϸĿ D,����֧����Ŀ E,�����ʻ� L,���ű� H, " & _
        "             (Select distinct Q.ҩƷid,Q.��ʶ��,T.���� From ҩƷĿ¼ Q,ҩƷ��Ϣ R,ҩƷ���� T  Where  Q.ҩ��id=R.ҩ��id and R.����=T.���� ) J " & _
        "        where  a.��¼״̬<>0 and  a.�շ����=b.���� and a.�շ�ϸĿid=J.ҩƷid(+)   and  Nvl(a.���ӱ�־,0)<>9 and a.�շ�ϸĿid=D.id and a.������=c.����(+)  and " & _
        "              a.�շ�ϸĿid=E.�շ�ϸĿID and A.����id=F1.����id and A.��ҳid=F1.��ҳID and a.����id=L.����ID and a.��������id=h.id  and " & _
        "              a.����ID = " & lng����ID & " And E.���� = " & TYPE_����
        
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡסԺ������ϸ"
    If rs��ϸ.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "û�н�����ϸ,���ܽ���"
        Exit Function
    End If
    
    dbl��ҩ���� = 0
    dbl��� = 0
    While Not rs��ϸ.EOF
        dbl��� = dbl��� + Nvl(rs��ϸ!���ʽ��, 0)
        If Nvl(rs��ϸ!�շ����) = "6" Or Nvl(rs��ϸ!�շ����) = "7" Then
            dbl��ҩ���� = dbl��ҩ���� + Nvl(rs��ϸ!����, 0)
        End If
        rs��ϸ.MoveNext
    Wend
    rs��ϸ.MoveFirst
    If dbl��� = 0 Then
        Err.Raise 9000, gstrSysName, "����û�з�������,���ܽ���ҽ������"
        Exit Function
    End If
    
    lng����ID = Nvl(rs��ϸ!����ID, 0)
    lng��ҳID = Nvl(rs��ϸ!��ҳID, 0)
   
   gstrSQL = "" & _
        "   Select a.����,b.����,b.����,a.����֢ From �����ʻ� a,���ղ��� b where a.��Ժ����ID=b.ID(+) and a.����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ����"
    str���ִ��� = Nvl(rsTemp!����)
    
    If Nvl(rsTemp!����) = "" Or Nvl(rsTemp!����֢) = "" Then
        Err.Raise 9000, gstrSysName, "�������벡�ִ���Ͳ���֢!"
        Exit Function
    End If
    g�������_����.�������� = Nvl(rsTemp!����, "1")
    
    'gstrSQL = "Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ���� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� = 3 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����id & " Group by ����id,��ҳid"
'    gstrSQL = "Select * From ���ղ��� where id=" & Nvl(rs��ϸ!����ID, 0)
    str���ִ��� = Nvl(rsTemp!����)
    str����֤ = Nvl(rsTemp!����֢)
    
    'aPrnHandle: [Ϊ��ɽӿڼ��ݶ�����δ��] ?
    StrInput = g�������_����.��������
    'aCheckOutBillNo�����㵥�ݺš�(13λ)not null
    StrInput = StrInput & vbTab & lng����ID
    'aInHosDate����Ժ����(8λ)(yyyymmdd)not null
    StrInput = StrInput & vbTab & Nvl(rs��ϸ!��Ժ����)
    'aOutHosDate����Ժ����(8λ)(yyyymmdd)not null
    StrInput = StrInput & vbTab & Nvl(rs��ϸ!��Ժ����)
    'aDepartmentName������(20λ)
    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!��������), 1, 20)
    'aDoctorName��ҽ��(10λ)
    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!ҽ��), 1, 10)
    'aoutHosDiagnoseCode����Ժ���ֱ���(10λ)
    StrInput = StrInput & vbTab & Substr(Nvl(str���ִ���), 1, 20)
    'aSubDiagnose������֢˵��(200λ)
    StrInput = StrInput & vbTab & Substr(str����֤, 1, 200)
    'aOutHosCure����Ժ���(1-����2-��ת 3-δ�� 4-����9-����)(1λ)not null
    StrInput = StrInput & vbTab & Substr(Get�������_����(lng����ID, lng��ҳID), 1, 1)
    'aInHosDays��סԺ����(3λ)
    StrInput = StrInput & vbTab & Substr(Int(Nvl(rs��ϸ!סԺ����, 0)), 1, 3)
    'aAmount����Ӧ��������ܽ�(8λ��2λС��)
    StrInput = StrInput & vbTab & dbl���
    
    If ��ݼ���_����(4, "23") = False Then Exit Function
    If IS�Ƿ�ˢ������(lng����ID) = False Then
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
        Exit Function
    End If
    
    If ҵ������_����(���ý��㵥, StrInput, strOutput) = False Then Exit Function
    
    delArrar            '������ô����¼����
    
    
    Do While Not rs��ϸ.EOF
        If Nvl(rs��ϸ!ҽ������, "") = "" Then
                Err.Raise 9000, gstrSysName, "��Ŀ[" & Nvl(rs��ϸ!��Ŀ����) & "]δ���ö�Ӧ��ҽ����Ŀ,��������ҽ��"
                Exit Function
        End If
        
        If Nvl(rs��ϸ!�Ƿ��ϴ�, 0) = 0 And Nvl(rs��ϸ!ʵ�ս��, 0) <> 0 Then
            Err.Raise 9000, gstrSysName, " ����δ�ϴ�����ϸ,������Ԥ��һ��!"
            Exit Function
        End If
        
        str�Ƿ�ҩƷ = "0"
        StrInput = Substr(Nvl(rs��ϸ!ҽ������), 1, 10)
        
        Select Case UCase(Nvl(rs��ϸ!�շ����))
            Case "5", "6", "7"
                str�Ƿ�ҩƷ = "1"
                'aItemName����Ŀ����(40λ)
                'aMediKindCode�����ô���(2λ)
                'aIsCityRich��(0-�׻���ͨ1-�һ�߾���)(1λ)[����2�Է�]
                'aIsCityMedi���Ƿ�ҽ��(1λ)(0-����1-��)[Ϊ��ɽӿڼ��ݶ�����]
                'aIsLimit���Ƿ�ҽ���޼���Ŀ(1λ)(0-����1-��)
                'aCitySelfPayRate���Ը�����(4λ��3λС��)
                'aPrice����׼����(8λ��2λС��)
                
                If ҵ������_����(ȡҩƷ��Ϣ, StrInput, strOutput) = False Then Exit Function
            Case "J", "H", "I"
                If ҵ������_����(ȡ������Ϣ, StrInput, strOutput) = False Then Exit Function
                str�Ƿ�ҩƷ = "2"
            Case Else
                If ҵ������_����(ȡ������Ϣ, StrInput, strOutput) = False Then Exit Function
        End Select
        If strOutput = "" Then
            Err.Raise 9000, gstrSysName, "�ڻ�ȡҩƷ����Ϣʱ���ַ���ֵΪ��,����ҽ���ṩ����ϵ!" & vbCrLf & " �������Ϊ:" & StrInput
            
            Exit Function
        End If
        strArr = Split(strOutput & vbTab & vbTab & vbTab, vbTab)
        
        '��¼���ô�����
        setArrar CStr(strArr(1)), Nvl(rs��ϸ!���ʽ��, 0)
        
        rs��ϸ.MoveNext
    Loop
    
    
    For iLoop = 0 To 50
        If strTempArr(iLoop, 0) <> "" Then
            StrInput = "1"
            StrInput = StrInput & vbTab & strTempArr(iLoop, 0)
            StrInput = StrInput & vbTab & strTempArr(iLoop, 1)
            If ҵ������_����(����סԺ������Ϣ, StrInput, strOutput) = False Then Exit Function
        Else
            Exit For
        End If
    Next
    
    If ҵ������_����(ȡ��סԺ���������Ϣ, "", strOutput) = False Then Exit Function
    'aAccRemain���ʻ����(8λ��2λС��)
    'aPayAcc�������ʻ�֧��(8λ��2λС��)
    'aPayCash�������ֽ�֧��(8λ��2λС��)
    'aPayPer�����˱�������(8λ��2λС��)
    'aPayPlan��ͳ��֧��(8λ��2λС��)
    'aPayCarePlan���չ�֧��(8λ��2λС��)
    'aPayCareSelf���չ˵渶(8λ��2λС��)
    'aPaySelfPart���Ը���֧��(8λ��2λС��)
    'aPayBusiness���̱�֧��(8λ��2λС��)
    'aCompMediFir������ҩƷ(8λ��2λС��)
    'aCompMediSelf���Է�ҩƷ(8λ��2λС��)
    'aCompMediSec������ҩƷ(8λ��2λС��)
    'aCompTreatFir����������(��ͨ)(8λ��2λС��)
    'aCompTreatSelf���Է�����(8λ��2λС��)
    'aCompTreatSec����������(�߾���)(8λ��2λС��)
    'aCompBedFir��������ʩ(8λ��2λС��)
    'aCompBedSelf���Է���ʩ(8λ��2λС��)
    'aCompBedSec��������ʩ(8λ��2λС��)
    'aCompOtherSelf�������Է�(8λ��2λС��)
    'aAccSelfPayPart���Ը����ۼ�(8λ��2λС��)
    'aAccPlanPay��ͳ��֧���ۼ�(8λ��2λС��)
    'aAccHeavyIll���ز�֧���ۼ�(8λ��2λС��)
    'aAccDeferIll������֧���ۼ�(8λ��2λС��)
    'aUBasePay  �ǻ���ҽ�Ʒѣ�8λ,2λС�������˺꣺����
    If strOutput = "" Then
        Err.Raise 9000, gstrSysName, "��ȡ��סԺ���������Ϣʱ�������˿�ֵ!"
        Exit Function
    End If
    strArr = Split(strOutput, vbTab)
    
    
    If Val(strArr(1)) <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & Format(Val(strArr(1)), "####0.00;-####0.00; ;")
    End If
    
    If Val(strArr(4)) <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||ͳ��֧��|" & Format(Val(strArr(4)), "####0.00;-####0.00; ;")
    End If
    
    '�������
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        #If gverControl < 2 Then
            blnOld = True
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
        #Else
            strAdvance = str���㷽ʽ
            gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    Dim intMouse As Integer
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If blnOld Then
        If frm������Ϣ.ShowME(lng����ID, True) = False Then
            Exit Function
        End If
    End If
    Screen.MousePointer = intMouse
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(�Ը����ۼ�),�ʻ��ۼ�֧��_IN(ͳ��֧���ۼ�),�ۼƽ���ͳ��_IN(�ز�֧���ۼ�),�ۼ�ͳ�ﱨ��_IN(����֧���ۼ�),סԺ����_IN,����(�ǻ���ҽ�Ʒ�),�ⶥ��_IN(�����ֽ�֧��),ʵ������_IN(���˱�������),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(�Ը���֧��),�����Ը����_IN(�Է�ҩƷ),
    '   ����ͳ����_IN(����ҩƷ),ͳ�ﱨ�����_IN(ͳ��֧��),    ���Ը����_IN(�չ�֧��),�����Ը����_IN(�չ˵渶),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(),��ҳID_IN,��;����_IN,��ע_IN
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
    
    With g�������_����
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          Val(strArr(19)) & "," & Val(strArr(20)) & "," & Val(strArr(21)) & "," & Val(strArr(22)) & ",NULL," & Val(strArr(23)) & "," & Val(strArr(2)) & "," & Val(strArr(3)) & "," & _
         dbl��� & "," & Val(strArr(7)) & "," & Val(strArr(10)) & "," & _
          Val(strArr(9)) & "," & Val(strArr(4)) & "," & Val(strArr(5)) & "," & Val(strArr(6)) & "," & Val(strArr(1)) & "," & _
          "'" & g�������_����.�������� & "',Null,Null,NULL" & IIf(blnOld, "", ",1") & ")"
              
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
        '������Ϣ����:
            '����_IN ,��¼ID_IN,�ʻ����_IN ,�̱�֧��_IN ,����ҩƷ_IN ,��������_IN ,�Է�����_IN ,��������_IN ,������ʩ_IN ,�Է���ʩ_IN ,������ʩ_IN ,�����Է�_IN
        gstrSQL = "zl_���ս����¼_������Ϣ( 2," & lng����ID & "," & Val(strArr(0)) & "," & Val(strArr(8)) & "," & Val(strArr(11)) & "," & Val(strArr(12)) & "," & Val(strArr(13)) & "," & Val(strArr(14)) & "," & Val(strArr(15)) & "," & Val(strArr(16)) & "," & Val(strArr(17)) & "," & Val(strArr(18)) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼������Ϣ")
    End With
    
    StrInput = strArr(1)    '�����ʻ�֧��
    StrInput = StrInput & vbTab & strArr(2) '�ֽ�֧��
    
    If ҵ������_����(���߷����ύ, StrInput, strOutput) = False Then Exit Function
    If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
    
    סԺ����_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

    Public Function סԺ�������_����(lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput  As String
    Dim strArr
    Dim lng����ID As Long
    
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    
    
    '�˷�
    gstrSQL = "select distinct A.����id, A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    'gstrSQL = "select distinct A.����id,A.����ID from ���˷��ü�¼ A,���˷��ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����ID"
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "û����ص���"
        Exit Function
    End If
    
    lng����ID = Nvl(rsTemp!ID)
    lng����ID = Nvl(rsTemp!����ID)
   
    StrInput = "1"
    StrInput = StrInput & vbTab & lng����ID & "R"
    StrInput = StrInput & vbTab & lng����ID
    If ��ݼ���_����(4, "27") = False Then Exit Function
    If IS�Ƿ�ˢ������(lng����ID) = False Then
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
        Exit Function
    End If
    
    If ҵ������_����(ȡ������, StrInput, strOutput) = False Then Exit Function
       
    gstrSQL = "Select * From ���ս����¼ Where ��¼ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "û���ҵ�ԭ�еı��ս����¼", vbInformation, gstrSysName
        Exit Function
    End If
        
    With g�������_����
        gstrSQL = "zl_���ս����¼_insert( 2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          -1 * Nvl(rsTemp!�ʻ��ۼ�����, 0) & "," & -1 * Nvl(rsTemp!�ʻ��ۼ�֧��, 0) & "," & -1 * Nvl(rsTemp!�ۼƽ���ͳ��, 0) & "," & -1 * Nvl(rsTemp!�ۼ�ͳ�ﱨ��, 0) & ",NULL," & -1 * Nvl(rsTemp!����, 0) & "," & -1 * Nvl(rsTemp!�ⶥ��, 0) & "," & -1 * Nvl(rsTemp!ʵ������, 0) & "," & _
         -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & _
          -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & "," & -1 * Nvl(rsTemp!���Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & "," & _
          "NULL,Null,Null,NULl)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
        '������Ϣ����:
        '����_IN ,��¼ID_IN,�ʻ����_IN ,�̱�֧��_IN ,����ҩƷ_IN ,��������_IN ,�Է�����_IN ,��������_IN ,������ʩ_IN ,�Է���ʩ_IN ,������ʩ_IN ,�����Է�_IN
        gstrSQL = "zl_���ս����¼_������Ϣ( 2," & lng����ID & "," & -1 * Nvl(rsTemp!�ʻ����, 0) & "," & -1 * Nvl(rsTemp!�̱�֧��, 0) & "," & -1 * Nvl(rsTemp!����ҩƷ, 0) & "," & -1 * Nvl(rsTemp!��������, 0) & "," & -1 * Nvl(rsTemp!�Է�����, 0) & "," & -1 * Nvl(rsTemp!��������, 0) & "," & -1 * Nvl(rsTemp!������ʩ, 0) & "," & -1 * Nvl(rsTemp!�Է���ʩ, 0) & "," & -1 * Nvl(rsTemp!������ʩ, 0) & "," & -1 * Nvl(rsTemp!�����Է�, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼������Ϣ")
    End With
    If ҵ������_����(���������ύ, "", strOutput) = False Then Exit Function
    If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
    סԺ�������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function �����Ǽ�_����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim str���ִ��� As String
    Dim dbl���� As Double, dbl��� As Double
    Dim StrInput As String, strOutput As String
    Dim str�Ƿ�ҩƷ  As String
    Dim strArr
    Dim collData  As Collection
    
    
    Err = 0
    On Error GoTo errHand:
    
    �����Ǽ�_���� = False
    DebugTool "���봦���Ǽ�:" & Time
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then

        gstrSQL = " " & _
            " Select A.id,A.����ID,F.סԺ��,A.NO,A.���,A.ҽ�����,A.��¼����,A.��¼״̬,A.�շ����,D.���,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') �Ǽ�ʱ��, " & _
            "        A.������ ҽ��,F1.סԺҽʦ סԺҽ��,to_char(f1.��Ժ����,'ss') as �Ǽ����,a.�Ƿ��ϴ�,V.��� AS ҽ�����,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,A.����,Round(A.ʵ�ս��/(A.����*A.����),2) as ʵ�ʼ۸�,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "        C.��Ŀ���� ҽ����Ŀ����,G.��� ,F.סԺ���� AS ��ҳid, " & _
            "        G.��ʶ����||G.��ʶ���� AS ���ұ���,G.���� AS ��Ŀ����,K.���� AS ����, " & _
            "        E.����,E.����,E.����,E.ҽ����,E.����,E.��Ա���,E.��λ����,E.˳���,E.����֤��,E.�ʻ����,E.��ǰ״̬, " & _
            "        E.����ID,E.��ְ,E.�����,E.�Ҷȼ�,to_char(E.����ʱ��,'yyyyMMddhh24miss') ����ʱ�� " & _
            " From סԺ���ü�¼ A,���ű� B,�շ���� D,�����ʻ� E,������Ϣ F,������ҳ F1,�շ�ϸĿ G,��Ա�� V," & _
            "       (Select J.����,O.ҩƷid From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K, " & _
            "       (Select M.��Ŀ����,M.�շ�ϸĿid From ����֧����Ŀ M Where M.����=[4]) C " & _
            " Where   a.��¼״̬<>0 and   a.����id=E.����ID AND a.����id=F.����ID AND A.����id=F1.����id and F1.����=" & TYPE_���� & " AND F.��ҳid= F1.��ҳid  AND  a.������=V.����(+) AND a.�շ�ϸĿid=k.ҩƷid(+) AND a.�շ�ϸĿid=G.id AND E.����=" & TYPE_���� & "   AND A.�շ����=D.���� AND  " & _
            "           A.��¼����=[1] and  A.��¼״̬=[2] And A.NO=[3]" & _
            "           And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
            " order by A.����ID,A.��¼״̬"
    Else
        gstrSQL = " " & _
            " Select A.id,A.����ID,F.סԺ��,A.NO,A.���,A.ҽ�����,A.��¼����,A.��¼״̬,A.�շ����,D.���,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') �Ǽ�ʱ��, " & _
            "        A.������ ҽ��,F1.סԺҽʦ סԺҽ��,to_char(f1.��Ժ����,'ss') as �Ǽ����,a.�Ƿ��ϴ�,V.��� AS ҽ�����,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,A.����,Round(A.ʵ�ս��/(A.����*A.����),2) as ʵ�ʼ۸�,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "        C.��Ŀ���� ҽ����Ŀ����,G.��� ,F.סԺ���� AS ��ҳid, " & _
            "        G.��ʶ����||G.��ʶ���� AS ���ұ���,G.���� AS ��Ŀ����,K.���� AS ����, " & _
            "        E.����,E.����,E.����,E.ҽ����,E.����,E.��Ա���,E.��λ����,E.˳���,E.����֤��,E.�ʻ����,E.��ǰ״̬, " & _
            "        E.����ID,E.��ְ,E.�����,E.�Ҷȼ�,to_char(E.����ʱ��,'yyyyMMddhh24miss') ����ʱ�� " & _
            " From סԺ���ü�¼ A,���ű� B,�շ���� D,�����ʻ� E,������Ϣ F,������ҳ F1,�շ�ϸĿ G,��Ա�� V," & _
            "       (Select J.����,O.ҩƷid From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K, " & _
            "       (Select M.��Ŀ����,M.�շ�ϸĿid From ����֧����Ŀ M Where M.����=[4]) C " & _
            " Where   a.��¼״̬<>0 and   a.����id=E.����ID AND a.����id=F.����ID AND A.����id=F1.����id and F1.����=" & TYPE_���� & " AND F.סԺ����= F1.��ҳid  AND  a.������=V.����(+) AND a.�շ�ϸĿid=k.ҩƷid(+) AND a.�շ�ϸĿid=G.id AND E.����=" & TYPE_���� & "   AND A.�շ����=D.���� AND  " & _
            "           A.��¼����=[1] and  A.��¼״̬=[2] And A.NO=[3]" & _
            "           And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
            " order by A.����ID,A.��¼״̬"

    End If
    
    '��һ��: ��ȡ������ϸ��¼
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", lng��¼����, lng��¼״̬, str���ݺ�, TYPE_����)
    
    If rs��ϸ.RecordCount = 0 Then
        ShowMsgbox "û����ϸ��¼!"
        Exit Function
    End If
    dbl���� = 0
    dbl��� = 0
    Dim lngCount As Long
    lngCount = 0
    lng����ID = 0
    Set collData = New Collection
    
    While Not rs��ϸ.EOF
        If Nvl(rs��ϸ!ҽ����Ŀ����, "") = "" Then
                ShowMsgbox "��Ŀ[" & Nvl(rs��ϸ!��Ŀ����) & "]δ���ö�Ӧ��ҽ����Ŀ,��������ҽ��"
                Exit Function
        End If
        If lng����ID <> Nvl(rs��ϸ!����ID, 0) Then
            lng����ID = Nvl(rs��ϸ!����ID, 0)
            dbl��� = 0: dbl���� = 0
            collData.Add Array(dbl����, dbl���), "K" & lng����ID
            lngCount = lngCount + 1
        End If
        collData.Remove "K" & lng����ID
        
        dbl��� = dbl��� + rs��ϸ!���
        If Nvl(rs��ϸ!�շ����) = "6" Or Nvl(rs��ϸ!�շ����) = "7" Then
            dbl���� = dbl���� + Nvl(rs��ϸ!����, 0)
        End If
        collData.Add Array(dbl����, dbl���), "K" & lng����ID
        
        rs��ϸ.MoveNext
    Wend
    
    If lngCount > 1 Then
        ShowMsgbox "����ͬʱ�Զ�����˽��м���,������Ϊ:" & lngCount
        Exit Function
    End If
    If InitInfor_����.��ϸʱʵ�ϴ� = False Then
        �����Ǽ�_���� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    
    
    lng����ID = Nvl(rs��ϸ!����ID, 0)
    lng��ҳID = Nvl(rs��ϸ!��ҳID, 0)
    
    gstrSQL = "Select * From ���ղ��� Where ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(Nvl(rs��ϸ!����ID, 0)), TYPE_����)
    str���ִ��� = ""
    If Not rsTemp.EOF Then
        str���ִ��� = Nvl(rsTemp!����)
    End If
    
    If lng��¼״̬ = 1 Then
        If ҵ������_����(��ʼ���������, "22", strOutput) = False Then Exit Function
    End If
    
    lng����ID = 0
   Do While Not rs��ϸ.EOF
        
        If Nvl(rs��ϸ!�Ƿ��ϴ�, 0) = 0 And rs��ϸ!��� <> 0 Then
            If lng��¼״̬ = 1 Then
                If lng����ID <> Nvl(rs��ϸ!����ID, 0) Then
                    lng����ID = Nvl(rs��ϸ!����ID, 0)
                    lng��ҳID = Nvl(rs��ϸ!��ҳID, 0)
                    
                    
                    DebugTool "�ϴ��������ʵ� ��ʼ:" & Time
                    'aInHosRegisterNo�����δ����סԺ�ǼǺš�
                    StrInput = lng����ID & "-" & lng��ҳID & "-" & Nvl(rs��ϸ!�Ǽ����)
                    'aSerialNo��סԺ�ʵ���(13λ)not null
                    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!NO) & "-" & Nvl(rs��ϸ!��¼����, 0), 1, 13)
                    'aDiagnoseCode�����ִ���(10λ)not null
                    StrInput = StrInput & vbTab & Substr(str���ִ���, 1, 10)
                    'aDepartmentName������(20λ)
                    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!��������), 1, 20)
                    'aDoctorName��ҽ��(10λ)
                    StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!ҽ��), 1, 10)
                    'aHerbalCopy���в�ҩ����(2λ)
                    StrInput = StrInput & vbTab & Substr(collData("K" & lng����ID)(0), 1, 10)
                            
                    'aAmount�����(8λ��2λС��)
                    StrInput = StrInput & vbTab & Substr(collData("K" & lng����ID)(1), 1, 10)
                    If ҵ������_����(����סԺ�ʵ�, StrInput, strOutput) = False Then Exit Function
                    DebugTool "�ϴ��������ʵ� ����:" & Time
                End If
                str�Ƿ�ҩƷ = "0"
                DebugTool "�ϴ�������ϸ ��ʼ:" & Time
                
                StrInput = Substr(Nvl(rs��ϸ!ҽ����Ŀ����), 1, 10)
                
                Select Case UCase(Nvl(rs��ϸ!�շ����))
                    Case "5", "6", "7"
                        str�Ƿ�ҩƷ = "1"
                        If ҵ������_����(ȡҩƷ��Ϣ, StrInput, strOutput) = False Then Exit Function
                        Case "J", "H", "I"
                        If ҵ������_����(ȡ������Ϣ, StrInput, strOutput) = False Then Exit Function
                        str�Ƿ�ҩƷ = "2"

                    Case Else
                        If ҵ������_����(ȡ������Ϣ, StrInput, strOutput) = False Then Exit Function
                End Select
                If strOutput = "" Then
                    strOutput = "" & vbTab & vbTab & vbTab & vbTab
'                    Exit Function
                End If
                strArr = Split(strOutput, vbTab)
                '���ýӿ�,д����ϸ
                'aBillHandle: [Ϊ��ɽӿڼ��ݶ�����δ��] ?
                StrInput = "1"
                'aCityMediCareNo��ҽ����Ŀ��š�(10λ)not null
                StrInput = StrInput & vbTab & Nvl(rs��ϸ!ҽ����Ŀ����, "")
                'aItemName��ҽԺ��Ŀ����(40λ)not null
                StrInput = StrInput & vbTab & Nvl(rs��ϸ!��Ŀ����, "")
                'aConformationName����������(20λ)
                StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!����, ""), 1, 20)
                'aUnitContent����λ����(14λ)
                StrInput = StrInput & vbTab & Substr(Nvl(rs��ϸ!���, ""), 1, 14)
                    '���˺�:��������
                    'gstrSQL = "Select ����,Ƶ��,�÷� From ҩƷ�շ���¼ where ����id=" & Nvl(!ID, 0)
                    'zlDataBase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵���Ƶ��"
                    'If Not rsTemp.EOF Then
                    'strInput = strInput & vbTab & Substr("����:" & Nvl(rsTemp!����, "") & " Ƶ��:" & Nvl(rsTemp!Ƶ��) & "�÷�:" & Nvl(rsTemp!�÷�), 1, 14)
                    'Else
                    'strInput = strInput & vbTab & ""
                    'End If
                'aDosage���÷�����(40λ)
                StrInput = StrInput & vbTab & ""
                'aMediKindCode�����ô������(2λ)not null
                StrInput = StrInput & vbTab & strArr(1)
                'aIsCityMedi���Ƿ�ҽ��(1λ)(0-����1-��)[Ϊ��ɽӿڼ��ݶ�����δ��]
                StrInput = StrInput & vbTab & strArr(3)
                'aIsRich��(0-�׻���ͨ1-�һ�߾���)(1λ)[����2�Է�]not null
                StrInput = StrInput & vbTab & strArr(2)
                'aIsMedi:�Ƿ�ҩƷ(0-��Ŀ1-ҩƷ2������ʩ[��] )not null
                StrInput = StrInput & vbTab & str�Ƿ�ҩƷ
                'aPrice������(8λ��2λС��)>0
                StrInput = StrInput & vbTab & Nvl(rs��ϸ!ʵ�ʼ۸�, 0)
                'aQuantity������(8λ��2λС��)>0
                StrInput = StrInput & vbTab & Nvl(rs��ϸ!����, 0)
                'aAmount�����(8λ��2λС��)>0
                StrInput = StrInput & vbTab & Nvl(rs��ϸ!���, 0)
                
                If ҵ������_����(���ü��ʵ���ϸ����, StrInput, strOutput) = False Then Exit Function
                DebugTool "�ϴ�������ϸ ����:" & Time
            End If
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        End If
        rs��ϸ.MoveNext
    Loop
    
    If lng��¼״̬ <> 1 Then
        '��������
        If ҵ������_����(��ʼ���������, "26", strOutput) = False Then Exit Function
        
'        If ��ݼ���_����(0, "26") = False Then Exit Function
        StrInput = "1"
        StrInput = StrInput & vbTab & str���ݺ� & "-" & lng��¼���� & "R"
        StrInput = StrInput & vbTab & str���ݺ� & "-" & lng��¼����
        If ҵ������_����(ȡ������, StrInput, strOutput) = False Then Exit Function
        If ҵ������_����(���������ύ, "", strOutput) = False Then Exit Function
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
    Else
        DebugTool "�ϴ�������ϸ �����ύ��ʼ:" & Time
        If ҵ������_����(���������ύ, "", strOutput) = False Then Exit Function
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
        DebugTool "�ϴ�������ϸ �����ύʧ��:" & Time
    End If
    
    
    �����Ǽ�_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Function Readģ������(ByVal intҵ������ As ҵ������_����, ByVal strInputString As String, ByRef strOutPutstring As String)
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
    
    strFile = App.Path & "\ģ���ύ��.txt"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Select Case intҵ������
    Case ����
        STRNAME = "����"
    Case ��ʼ���������
        Exit Function
    Case ȡ����Ժ�Ǽ�
        Exit Function
    Case ���������ύ
        Exit Function
    Case �����������
        Exit Function
    Case ��Ժ�Ǽ�
        Exit Function
    End Select
   
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
                    
                If blnStart Then
                    If strText = "" Then
                        strText = "" & vbTab
                    End If
                    strArr = Split(strText, "|")
                    
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
            Loop
            objText.Close
            strOutPutstring = str
    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function
Public Function ��ݼ���_����(ByVal bytType As Byte, Optional str���� As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:Զ����ݼ���
    '--�����:bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '--������:
    '--��  ��:�ɹ�true,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim blnReturn As Boolean
    Dim StrInput As String
    Dim strOutput As String
    Dim strArr
    
        
    Err = 0
    On Error GoTo errHand:
    
    ��ݼ���_���� = False
    Select Case bytType
    Case 0      '����
        StrInput = "10"
    Case 1      '��Ժ�Ǽ�
        StrInput = "20"
    Case Else
        StrInput = "23"
    End Select
    If str���� <> "" Then
        '��ȡ�������:
        StrInput = str����
    End If
    
    If ҵ������_����(��ʼ���������, StrInput, strOutput) = False Then Exit Function
    
    If ҵ������_����(����, "", strOutput) = False Then Exit Function
    
    If strOutput = "" Then
        '���˺� /*200408*/
        DebugTool "��ȡ������Ϣʱ�����˴�����Ϊ����!"
        Exit Function
    End If
    
    strArr = Split(strOutput, vbTab)
    
    
    '�����ñ�����ֵ
    With g�������_����
        .���� = strArr(0)
        .���� = strArr(1)
        .���֤�� = strArr(2)
        .���� = strArr(3)
        .�Ա� = IIf(Val(strArr(4)) = 0, "Ů", "��")
        .�������� = GetStringToDate(strArr(5))
        .ҽ���� = strArr(6)
        .��λ���� = strArr(7)
        .������� = strArr(8)
        .����Ա��־ = strArr(9)  '(0-�ǹ���Ա1-����Ա���������չ���Ա)
        .���䱣�� = strArr(10)    '0-���μ�1-�μ�
        .��ҽ�� = strArr(11)    '0-���μ�1-�μ�
        .������ϵ = strArr(12)     '1-��������0-����
        .�Ƿ����Բ� = strArr(13)   '0-����1-��
        .�ش󼲲� = strArr(14)     '0-����1-��
        .�չ˼��� = strArr(15)    '0-��1-һ��2-����3-����
        .ְ������ = strArr(16)     '0����1��פ���2-��ذ���
        .סԺ��־ = strArr(17)     '0-��סԺ 1-סԺ
        
        .�𸶶�ҽ�Ʒ��ۼ� = Val(strArr(18))
        .ͳ��֧���ۼ� = Val(strArr(19))
        .���߽���ۼ� = Val(strArr(20))
        
        .����ͳ��֧���ۼ� = Val(strArr(21))
        .���ۼ� = Val(strArr(22))
        .�ʻ���� = Val(strArr(23))
        .סԺ���� = Val(strArr(24))
        .֧������ = Val(strArr(25))
    End With
    ��ݼ���_���� = True
    DebugTool "��ݼ���ɹ�"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    ��ݼ���_���� = False
End Function

Private Function GetStringToDate(ByVal StrInput As String) As String
    '����:������"20040404"ת����"2004-04-04"
    Dim intLen As Integer
    Dim strTemp As String
    intLen = Len(StrInput)
    Select Case intLen
    Case 6
        strTemp = Left(StrInput, 4) & "-0" & Mid(StrInput, 5, 1) & "-0" & Mid(StrInput, 6, 1)
    Case 8
        strTemp = Left(StrInput, 4) & "-" & Mid(StrInput, 5, 2) & "-" & Mid(StrInput, 7, 2)
    Case Else
        strTemp = StrInput
    End Select
    GetStringToDate = strTemp
    
    
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    
    'rsExse:�ַ���
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    
    Dim rsTemp As New ADODB.Recordset
    Dim dbl�ܶ� As Double
    Dim lng��ҳID As Long
    Dim lng����id1 As Long
    Dim intMouse  As Integer
    Err = 0
    On Error GoTo errHand:
    סԺ�������_���� = ""
    DebugTool "�����������:" & Time
    
    gstrSQL = "Select ��ǰ״̬ From �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�ж��Ƿ��Ժû��"
    If Nvl(rsTemp!��ǰ״̬, 0) = 1 Then
        ShowMsgbox "�ò��˻������Ժ,���Բ��ܽ���!"
        DebugTool "�˳��������ʧ��,����δ��Ժ:" & Time
        Exit Function
    End If
    
    With rsExse
        dbl�ܶ� = 0
        lng��ҳID = Nvl(!��ҳID, 0)
        Do While Not .EOF
            g�������_����.�����ܶ� = g�������_����.�����ܶ� + Nvl(!���, 0)
            .MoveNext
        Loop
    End With
    
    If bln���ʴ� Then
        Screen.MousePointer = 1
        If ��ݱ�ʶ_����(4, lng����id1) = "" Then
            Screen.MousePointer = intMouse
            סԺ�������_���� = ""
            Exit Function
        End If
        Screen.MousePointer = intMouse
        If lng����ID <> lng����id1 Then
            ShowMsgbox "���ǵ�ǰҪ����Ĳ���!"
            סԺ�������_���� = ""
            Exit Function
        End If
    End If
    
    Call ������ϸ��¼(lng����ID, lng��ҳID)


    DebugTool "�˳��������ɹ�:" & Time
    סԺ�������_���� = "ͳ��֧��;" & 0 & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ������ϸ��¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '���������ϸ��¼
    Dim rsTemp As New ADODB.Recordset
    Dim strNO As String, str���ִ��� As String
    Dim dbl��ҩ���� As Double, dbl�ܶ� As Double
    Dim StrInput  As String, strOutput As String
    Dim strArr
    Dim str�Ƿ�ҩƷ  As String
    Dim bln�ύ As Boolean
    
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select B.���� From �����ʻ� A,���ղ��� B where a.����ID=B.ID and B.����=" & TYPE_���� & " and a.����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����"
    If Not rsTemp.EOF Then
        str���ִ��� = Nvl(rsTemp!����)
    End If
    ������ϸ��¼ = False
    DebugTool "���벹��ϸ�ӿ�:" & Time
    
    gstrSQL = "" & _
            "   Select A.ID,A.����ID,a.��ҳid,A.������ as ҽ��,A.�շ����,A.��¼����,B.���� as ��������,to_char(f.��Ժ����,'ss') as �Ǽ����,A.��¼״̬,A.NO,Decode(A.�շ����,'6',A.����,'7',A.����,0) as ��ҩ����," & _
            "           C.��Ŀ���� as ҽ����Ŀ����,G.���� as ��Ŀ����,G.���� as �շ���Ŀ,G.���,K.���� ����," & _
            "           A.����*A.���� as ����,Round(Nvl(A.ʵ�ս��,0)/(A.����*A.����),2) as ʵ�ʼ۸�,Nvl(A.ʵ�ս��,0) as ʵ�ս��" & _
            "   From סԺ���ü�¼ A,���ű� B,������ҳ F,�շ�ϸĿ G," & _
            "       (Select M.��Ŀ����,M.�շ�ϸĿid From ����֧����Ŀ M Where M.����=" & TYPE_���� & ") C," & _
            "       (Select J.����,O.ҩƷid From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K " & _
            "   where nvl(A.�Ƿ��ϴ�,0)=0 and mod(nvl(A.��¼״̬,0),3)<>2 and nvl(a.��¼״̬,0)<>0  and A.���ʷ���=1 and A.����ID is null and nvl(A.ʵ�ս��,0)<>0  and  " & _
            "       nvl(a.Ӥ����,0)=0 and  A.�շ�ϸĿid=K.ҩƷid(+) And A.��������id+0=B.ID and A.�շ�ϸĿid=G.id and A.�շ�ϸĿid=C.�շ�ϸĿid(+)  and A.����id=F.����id and A.��ҳid=F.��ҳid   and A.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & _
            "   ORDER BY a.��¼����,A.NO,A.���"
            
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�����ϴ���ϸ"
    
    If rsTemp.EOF Then
        DebugTool "�޲�����¼,����:" & Time
        GoTo go120:
    End If
    strNO = ""
    
    DebugTool "���ڱ�����¼,����:" & Time
    
    If ��ݼ���_����(1, "22") = False Then Exit Function
    
    If IS�Ƿ�ˢ������(Nvl(rsTemp!����ID, 0)) = False Then
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
        Exit Function
    End If
  
    With rsTemp
        DebugTool "��ʼ���ҽ����Ŀ�Ķ������,����:" & Time
        Do While Not .EOF
            If Nvl(!ҽ����Ŀ����) = "" Then
                ShowMsgbox "��Ŀ[" & Nvl(!��Ŀ����) & "]δ���ö�Ӧ��ҽ����Ŀ,��������ҽ��"
                Exit Function
            End If
            .MoveNext
        Loop
        DebugTool "�������ҽ����Ŀ�Ķ������,����:" & Time
        .MoveFirst
        '�ϴ�������ϸ
        bln�ύ = False
        DebugTool "��ʼ����������ϸ,����:" & Time
        strNO = ""
        Do While Not .EOF
            If strNO <> Nvl(!��¼����) & "-" & Nvl(!NO) & "-" & Nvl(!��¼״̬) Then
                strNO = Nvl(!��¼����) & "-" & Nvl(!NO) & "-" & Nvl(!��¼״̬)
                If GetSumJe(Nvl(!��¼����, 0), Nvl(!NO), Nvl(!��¼״̬, 1), dbl��ҩ����, dbl�ܶ�) = False Then Exit Function
                'aInHosRegisterNo�����δ����סԺ�ǼǺš�
                StrInput = lng����ID & "-" & lng��ҳID & "-" & Nvl(!�Ǽ����)
                'aSerialNo��סԺ�ʵ���(13λ)not null
                StrInput = StrInput & vbTab & Substr(Nvl(!NO) & "-" & Nvl(!��¼����, 0), 1, 13)
                'aDiagnoseCode�����ִ���(10λ)not null
                StrInput = StrInput & vbTab & Substr(str���ִ���, 1, 10)
                'aDepartmentName������(20λ)
                StrInput = StrInput & vbTab & Substr(Nvl(!��������), 1, 20)
                'aDoctorName��ҽ��(10λ)
                StrInput = StrInput & vbTab & Substr(Nvl(!ҽ��), 1, 10)
                'aHerbalCopy���в�ҩ����(2λ)
                StrInput = StrInput & vbTab & Substr(dbl��ҩ����, 1, 10)
                'aAmount�����(8λ��2λС��)
                StrInput = StrInput & vbTab & Substr(dbl�ܶ�, 1, 10)
                
                If ҵ������_����(����סԺ�ʵ�, StrInput, strOutput) = False Then Exit Function
            End If
            
            str�Ƿ�ҩƷ = "0"
            StrInput = Substr(Nvl(!ҽ����Ŀ����), 1, 10)
            
            Select Case UCase(Nvl(!�շ����))
                Case "5", "6", "7"
                    str�Ƿ�ҩƷ = "1"
                    If ҵ������_����(ȡҩƷ��Ϣ, StrInput, strOutput) = False Then Exit Function
                Case "J", "H", "I"
                    If ҵ������_����(ȡ������Ϣ, StrInput, strOutput) = False Then Exit Function
                    str�Ƿ�ҩƷ = "2"

                Case Else
                    If ҵ������_����(ȡ������Ϣ, StrInput, strOutput) = False Then Exit Function
            End Select
            
            If strOutput = "" Then
                ShowMsgbox "�ڻ�ȡҩƷ����Ϣʱ�����˿�ֵ������ҽ���ṩ����ϵ!" & vbCrLf & " �������Ϊ:" & StrInput
                Exit Function
            End If
            
            strArr = Split(strOutput, vbTab)
            
            '���ýӿ�,д����ϸ
            'aBillHandle: [Ϊ��ɽӿڼ��ݶ�����δ��] ?
            StrInput = "1"
            'aCityMediCareNo��ҽ����Ŀ��š�(10λ)not null
            StrInput = StrInput & vbTab & Nvl(!ҽ����Ŀ����, "")
            'aItemName��ҽԺ��Ŀ����(40λ)not null
            StrInput = StrInput & vbTab & Nvl(!�շ���Ŀ, "")
            'aConformationName����������(20λ)
            StrInput = StrInput & vbTab & Substr(Nvl(!����, ""), 1, 20)
            'aUnitContent����λ����(14λ)
            StrInput = StrInput & vbTab & Substr(Nvl(!���, ""), 1, 14)
                '���˺�:��������
                'gstrSQL = "Select ����,Ƶ��,�÷� From ҩƷ�շ���¼ where ����id=" & Nvl(!ID, 0)
                'zlDataBase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵���Ƶ��"
                'If Not rsTemp.EOF Then
                'strInput = strInput & vbTab & Substr("����:" & Nvl(rsTemp!����, "") & " Ƶ��:" & Nvl(rsTemp!Ƶ��) & "�÷�:" & Nvl(rsTemp!�÷�), 1, 14)
                'Else
                'strInput = strInput & vbTab & ""
                'End If
            'aDosage���÷�����(40λ)
            StrInput = StrInput & vbTab & ""
            'aMediKindCode�����ô������(2λ)not null
            StrInput = StrInput & vbTab & strArr(1)
            'aIsRich��(0-�׻���ͨ1-�һ�߾���)(1λ)[����2�Է�]not null
            StrInput = StrInput & vbTab & strArr(2)
            'aIsCityMedi���Ƿ�ҽ��(1λ)(0-����1-��)[Ϊ��ɽӿڼ��ݶ�����δ��]
            StrInput = StrInput & vbTab & strArr(3)
            'aIsMedi:�Ƿ�ҩƷ(0-��Ŀ1-ҩƷ2������ʩ[��] )not null
            StrInput = StrInput & vbTab & str�Ƿ�ҩƷ
            'aPrice������(8λ��2λС��)>0
            StrInput = StrInput & vbTab & Nvl(!ʵ�ʼ۸�, 0)
            'aQuantity������(8λ��2λС��)>0
            StrInput = StrInput & vbTab & Nvl(!����, 0)
            'aAmount�����(8λ��2λС��)>0
            StrInput = StrInput & vbTab & Nvl(!ʵ�ս��, 0)
            If ҵ������_����(���ü��ʵ���ϸ����, StrInput, strOutput) = False Then Exit Function
            bln�ύ = True
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & !ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            .MoveNext
        Loop
        DebugTool "��������������ϸ,����:" & Time
        
        If bln�ύ Then
            '������,���ύ
            If ҵ������_����(���������ύ, "", strOutput) = False Then Exit Function
        End If
        
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
    End With
    
go120:
    gstrSQL = "" & _
            "   Select A.ID,A.����ID,a.��ҳid,A.������ as ҽ��,A.�շ����,A.��¼����,B.���� as ��������,to_char(f.��Ժ����,'ss') as �Ǽ����,A.��¼״̬,A.NO,Decode(A.�շ����,'6',A.����,'7',A.����,0) as ��ҩ����," & _
            "           C.��Ŀ���� as ҽ����Ŀ����,G.���� as ��Ŀ����,G.���� as �շ���Ŀ,G.���,K.���� ����," & _
            "           A.����*A.���� as ����,Round(Nvl(A.ʵ�ս��,0)/(A.����*A.����),2) as ʵ�ʼ۸�,Nvl(A.ʵ�ս��,0) as ʵ�ս��" & _
            "   From סԺ���ü�¼ A,���ű� B,������ҳ F,�շ�ϸĿ G," & _
            "       (Select M.��Ŀ����,M.�շ�ϸĿid From ����֧����Ŀ M Where M.����=" & TYPE_���� & ") C," & _
            "       (Select J.����,O.ҩƷid From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K " & _
            "   where nvl(A.�Ƿ��ϴ�,0)=0 and mod(nvl(A.��¼״̬,0),3)=2 and A.���ʷ���=1 and A.����ID is null and nvl(A.ʵ�ս��,0)<>0  and  " & _
            "       nvl(a.Ӥ����,0)=0 and  A.�շ�ϸĿid=K.ҩƷid(+) And A.��������id+0=B.ID and A.�շ�ϸĿid=G.id and A.�շ�ϸĿid=C.�շ�ϸĿid(+)  and A.����id=F.����id and A.��ҳid=F.��ҳid   and A.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & _
            "   ORDER BY a.��¼����,A.NO,A.���"
            
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�����ϴ���ϸ"
    
    If rsTemp.EOF Then
        ������ϸ��¼ = True
        DebugTool "�޲������˷Ѽ�¼,����:" & Time
        Exit Function
    End If
    
    With rsTemp
        DebugTool "��ʼ���ҽ����Ŀ�Ķ������,����:" & Time
        Do While Not .EOF
            If Nvl(!ҽ����Ŀ����) = "" Then
                ShowMsgbox "��Ŀ[" & Nvl(!��Ŀ����) & "]δ���ö�Ӧ��ҽ����Ŀ,��������ҽ��"
                Exit Function
            End If
            .MoveNext
        Loop
        DebugTool "�������ҽ����Ŀ�Ķ������,����:" & Time
        .MoveFirst
        '������صĵ���
        strNO = ""
        If ��ݼ���_����(0, "26") = False Then Exit Function
        
        bln�ύ = False
        Do While Not .EOF
            If strNO <> Nvl(!��¼����) & "-" & Nvl(!NO) & "-" & Nvl(!��¼״̬) Then
                strNO = Nvl(!��¼����) & "-" & Nvl(!NO) & "-" & Nvl(!��¼״̬)
                '��������
                StrInput = "1"
                StrInput = StrInput & vbTab & Nvl(!NO) & "-" & Nvl(!��¼����, 0) & "R"
                StrInput = StrInput & vbTab & Nvl(!NO) & "-" & Nvl(!��¼����, 0)
                If ҵ������_����(ȡ������, StrInput, strOutput) = False Then Exit Function
                bln�ύ = True
            End If
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & !ID & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            .MoveNext
        Loop
        If bln�ύ Then
            If ҵ������_����(���������ύ, "", strOutput) = False Then Exit Function
        End If
        
        If ҵ������_����(�����������, "", strOutput) = False Then Exit Function
    End With
    
    DebugTool "������¼�ϴ��ɹ�,����:" & Time
    ������ϸ��¼ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetSumJe(ByVal lng��¼���� As Long, ByVal strNO As String, ByVal lng��¼״̬ As Long, dbl��ҩ���� As Double, dbl�ܶ� As Double) As Boolean
    '����:��ȡָ�����ݵĻ��ܶ�
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
            "   Select  SUM(Decode(A.�շ����,'6',NVL(A.����,0),'7',NVL(A.����,0),0)) as ��ҩ����," & _
            "           Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as ��� From סԺ���ü�¼ a " & _
            "   where nvl(A.�Ƿ��ϴ�,0)=0 and A.��¼״̬=" & lng��¼״̬ & " and A.���ʷ���=1 and " & _
            "       nvl(a.Ӥ����,0)=0  and a.��¼����=" & lng��¼���� & " and a.No='" & strNO & "'"
                
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ݻ��ܶ�"
    dbl��ҩ���� = Nvl(rsTemp!��ҩ����, 0)
    dbl�ܶ� = Nvl(rsTemp!���, 0)
    GetSumJe = True
    Exit Function
errHand:
    dbl��ҩ���� = 0
    dbl�ܶ� = 0
    GetSumJe = False
End Function
Public Function �ҺŽ���_����(ByVal lng����ID As Long) As Boolean
     �ҺŽ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function �Һų���_����(ByVal lng����ID As Long) As Boolean
    �Һų���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function ���²���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim str��Ժ���� As String, str����֢ As String
    
    Err = 0
    On Error GoTo errHand:
    
    ���²���_���� = frm����ѡ��_����.ShowSelect(TYPE_����, lng����ID, lng��ҳID, str��Ժ����, str����֢)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.��������
End Function
'
