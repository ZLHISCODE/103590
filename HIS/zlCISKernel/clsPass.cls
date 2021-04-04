VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------------
'̫Ԫͨ�ӿڽṹ
'--------------------------------------------------------------------------------------------------------------------------------------
'ҩ���ṹ
Public Type PatDrug
    drugID As String                'his ϵͳ��ҩƷ���벻Ϊ��
    DrugName As String              'his ϵͳ��ҩƷ���Ʋ�Ϊ��
    recMainNo As String             'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψһ��ֻ��Ϊ���ֱ��룬��Ϊ��
    recSubNo As String              'his ϵͳ��ҽ����ţ���һ�ξ���/סԺ��Ψһ��ֻ��Ϊ���ֱ���,��Ϊ��
    dosage As String                'his ϵͳ��ҽ��ҩƷʹ�ü�����Ϊ��
    doseUnits As String             'his ϵͳ��ҽ��ҩƷ������λ��Ϊ��
    administrationID As String      'his ϵͳ��ҽ��;�����벻Ϊ��
    performFreqDictID As String     'his ϵͳ��ҽ��Ƶ�δ��벻Ϊ��
    performFreqDictText As String   'his ϵͳ��ҽ��ִ��Ƶ��������Ϊ��
    startDateTime As String         'his ϵͳ��ҽ����ʼʱ��,��ʽ YYYY-MM-DD HH: MM: SS ��Ϊ��
    stopDateTime As String          'his ϵͳ��ҽ������ʱ��,��ʽ YYYY-MM-DD HH: MM: SS
    doctorDept As String            'his ϵͳ�Ŀ�ҽ��ҽ�����ڿ��Ҵ���
    DoctorID As String              'his ϵͳ�Ŀ�ҽ��ҽ������
    Doctor As String                'his ϵͳ�Ŀ�ҽ��ҽ������,
    isNew As String                 '����ҽ��ֵΪ1������Ϊ0
End Type
'�����Ϣ
Public Type PatDiagnosis
    diagnosisID As String          'his ϵͳ����ϱ���
    diagnosisName As String          'his ϵͳ���������
    diagnosisType As String          'his ϵͳ��������ͣ���������ϡ���Ժ��ϵ�
End Type
'����Դ
Public Type PatDrugSensitive
    patOrderDrugSensitiveID As String   'Ĭ��Ϊ0
    drugAllergenID As String        'his ϵͳ�Ĺ�������
End Type
'֢״���ݽṹ
Public Type PatSymptom
    symptomID As String           'his ϵͳ��֢״����
    symptomName As String           'his ϵͳ��֢״����
End Type

'��鴫����Ϣ
Public Type PatientOrder
    PatientID As String        '���˱�ʶ�Ų�Ϊ��
    Pname As String             '����������Ϊ��
    pSex As String              '�����Ա� 1-Ů, 0-��, 2-����
    pdateOfBirth As String      '���˳������ڣ���ʽ YYYY-MM-DD ��Ϊ�գ����
    pWeight As String           ' ��������
    pHeight As String           ' ��������

    isLact As String                '�Ƿ��飬��Ϊ1����Ϊ0 ��Ϊ��
    isPregnant As String            '�Ƿ��и�����Ϊ1 ����Ϊ0 ��Ϊ��
    isLiverWhole As String          '�Ƿ�ι��쳣 1-�쳣��0-������Ϊ��
    isKidneyWhole As String         '�Ƿ������쳣 1-�쳣��0-������Ϊ��
    PvisitID As String               '������ţ�Ψһ��ʶ���ߵ�һ�ξ���(�����\סԺ��)

    PatDrugs As Variant              '����ҩ������
    PatDiagnoses As Variant   '�������
    PatDrugSensitives As Variant   '������¼
    PatSymptoms As Variant       '֢״����

    DoctDeptID As String           '��½ҽ�����Ŵ��벻Ϊ��
    DoctDeptName As String         '��½ҽ���������Ʋ�Ϊ��
    DoctID As String               '��½ҽ�����벻Ϊ��
    DoctName As String             '��½ҽ��������Ϊ��
    DoctTitleID As String         '��½ҽ��ְ�����Ϊ��'C --�����ڣ����ڣ�������ҽʦ������ҽʦ��ר��
    'B������ҽʦ����ʦ  'A�������ϵ�����ְ��
    DoctTitleName As String          '��½ҽ��ְ�����Ʋ�Ϊ��
    SysFlag As String              '1�� ����ҽ������վ 2�� סԺҽ������
End Type

'�������<table></table>
Public Type AuditResult
    recMainNo As String         'ҽ�����
    recSubNo As String           'ҽ�����
    alertLevel As String         '��ʾ���� 1�� ���ɣ���ɫ��ʾ�ƣ�2�� ���ã���ɫ��ʾ��ʾ) 3�� ��ʾ����ɫ��ʾ�ƣ�
    strChecksum As String        '�����
End Type