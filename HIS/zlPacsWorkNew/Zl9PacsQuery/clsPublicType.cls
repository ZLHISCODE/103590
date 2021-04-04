VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'������PacsWork�б�ѡ����Ҫ������  TPacsInfo
Public Type TStudyInfo
     lngPatId As Long                 '1 ����ID
     lngPageID As Long                  '2 ��ҳID
     lngAdviceID As Long                '3 ҽ��ID
     lngSendNo As Long                  '4 ���ͺ�
     strPatientName As String           '5 ��������
     strPatientSex As String
     strPatientAge As String
     
     dtRefreshDate As Date
'Public strNO As String
'Public lngRecordKind As Long

'Public lngPatDept As Long                 '6 ������������
'Public strRegNo As String                 '7 �Һŵ�
'Public lngRegId As Long                   '8 �Һ�id
     intMoved As Integer                '9 �Ƿ�ת��
     intState As Integer                '10 ���״̬
     intStep As Integer                 '11 ������
'Public lngUnit As Long                    '12 ����ID
'Public strStudyUID As String              '13 ���UID
'Public blnCanPrint As Boolean             '14 �Ƿ��ܹ���ӡ
'Public blnIsInsidePatient As Boolean      '15 �Ƿ������סԺ����
'Public lngExeDepartmentId As Long         '16 ִ�в���ID
'Public strDoDoctor As String              '17 ��鼼ʦ
'Public strExeRoom As String               '18 ִ�м�
     lngPatientFrom As Long             '19 ������Դ
 strStudyNum As String              '20 ����
'     strBedNum As String                '21 ����
     lngMarkNum As Double               '22 ��־��
'Public lngBaby As Long                    '23 Ӥ��
'Public strPatientDepartment As String     '24 ���˿�������
'Public strReportDoctor As String          '25 ������
'Public strReportOperation As String       '26 �������
     lngLinkId As Long                  '27 ����ID
'Public strImgType As String               '28 Ӱ�����
'Public intImageLocation As Integer        '29 PACSӰ�����ڵ�λ�ã�0������PACS��1������PACS
'Public strPatientType As String           '30 ��������
'Public intFilmGiveOut As Integer          '��Ƭ����
'Public intReportGiveOut As Integer        '���淢��
'Public intDangerState As Integer          'Σ��״̬
'Public intEmergentTag As Integer          '����״̬
'Public intGreenChannel As Integer         '��ɫͨ��
'Public strAdviceContext As String             'ҽ������
'Public strAdviceDepartAndMethod As String     '��λ����
     strStuStateDesc As String          '���״̬����
'Public blnIsTechincalSure As Boolean      '�Ƿ�ʦȷ��
'Public strMoneyState As String            '����״̬����
'Public blnIsReported As Boolean           '�Ѿ��б���
End Type


'�ؼ�����
Public Enum TControlType
    ctText = 0  '�ı���
    ctDate = 1  '���ڿ�
    ctTime = 2  'ʱ���
    ctDateTime = 3  '����ʱ���
    ctCombobox = 4  '������
    ctList = 5      '��ѡ��
    ctAgeCbx = 6    '�����ѡ��
    ctMutxCbx = 7   '�����
    ctChk = 8       '��ѡ��
    ctQueryWay = 9  '����ѡ���
    ctFastDate = 10  '���ڿ�ѡ��
End Enum

'������Դ����
Public Enum TDBFromType
    dbftText = 0
    dbftSql = 1
End Enum

''���ٹ���ѡ������
'Public Enum TFilterType
'    ftSelect = 0
'    ftFast = 1
'End Enum

'���ٹ���ѡ������
Public Enum TSelectWay
    swSingle = 0
    swMulti = 1
End Enum

Public Enum TInputType
    itPopup = 0
    itFast = 1
    itBoth = 2
End Enum

Public Enum TLikeWay
    lwNormal = 0    '����
    lwLeft = 1      '��ƥ��
    lwRight = 2     '��ƥ��
    lwAll = 3       'ȫƥ��
End Enum

'���˴�����¼������������Ϣ�ṹ
Public Type TInputTag
    ParName As String   '������������
    DisplayName As String   '������ʾ���ƣ���[ϵͳ.��ʼ����]������ʾΪ[��ʼ����]
    FromType As TDBFromType    '������Դ����0-�ı���1-sql
    ExtProperty As String
    DataFrom As String  '������Դ��������
    Default As String   'Ĭ��ֵ
    ControlType As Long '�ؼ�����
    
    ParList() As String '��ʹ�ò����б�
    InputControl As Object  '��Ӧ��¼��ؼ�
    ReleationInputIndex() As Long    '����¼����������λ��
    Index As Long
    Tag As String
End Type

'��ѯ����������Ϣ�ṹ
Public Type TSchemeBaseInfo
    SchemeId As Double    '����ID
    Name As String      '��������
    Ver As Double       '�����汾
    Memo As String '����˵��
    IsDefault As Boolean '�Ƿ�Ĭ��
    IsOften As Boolean  '�Ƿ���
End Type

'�����ṹ
Public Type TSchemeItem
    IdKey As Double    '����ʶ��ؼ���
    BaseInfo As TSchemeBaseInfo '������Ϣ
    Index As Long           '����λ��
    IsLoaded As Boolean '�Ƿ��������
    SqlScheme As clsSqlScheme   '�������ö���
    SqlParse As clsSqlParse     'sql��������
    Startdate As Date
    EndDate As Date
    InitValues As clsSqlFilterValue
    FilterValues As clsSqlFilterValue
    FilterWindow As Object
End Type

'��ѯ�����Ĺ�������¼�����
Public Type FilterData
    Name As String
    Value As Variant
End Type
