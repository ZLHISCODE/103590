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


Public Enum TImageType
    mtTagImage = 0      '���ͼ
    mtReportImage = 1   '����ͼ
    mtStudyImage = 2    '���ͼ
End Enum


Public Enum TReportImgType
    ritReport = 0   '����ͼ
    ritMark = 1     '���ͼ
End Enum


Public Enum TOutlineType
    otNone = 0
    otDesc = 1      '����
    otOpin = 2      '���   '������
    otAdvi = 3      '����
End Enum


Public Enum TImgProcessType 'ͼ������������
    ptPreview = 0
    ptProcess = 1
    ptMark = 2
End Enum

Public Type TReportImgTag
    strKey As String        '����ؼ���
    lngFromAdvice As Long    'ͼ����Դ��ҽ��
    lngFileId As Double       '�ļ�ID
    lngTableId As Double      '����ID
    strImgFile As String    'ͼ���ļ�
    strObjectTag As String  '������
    strPros As String       '��������
    strImgMarks As String   'ͼ����
    lngStartVer As String   '��ʼ�汾
    lngImgType As TReportImgType    'ͼ������
End Type


Public Enum TImgMarkType
    imtNormal = 0       '�����б��
    imtAuto = 1         '�Զ����
    imtSpecify = 2      'ָ�����
End Enum

Public Type TReportSignInfo
    ID As Double
    �ļ�ID As Double
    ��ID As Double
    Key As String
    ������� As Double
    ���� As String
    ǩ����Ϣ As String
    ǰ������ As String
    ǩ������ As String
    �������� As String
    ǩ���汾 As Long
    ��ʼ�� As Long
    ��ֹ�� As Long
    ǩ����ʽ As Long
    ǩ������ As Long
    ʱ��� As String
    ʱ�����Ϣ As String
    ֤��ID As Long
    ǩ��ͼƬ As Boolean
    ǩ����ID As String
    ǩ��Ҫ�� As String
    ��ʾ��ǩ As Boolean
    ǩ��ʱ�� As String
    ��ʾʱ�� As String
End Type



Public Type TDicomBaseInfo
    lngAdviceId As Long         'ҽ��ID
    lngSendNo As Long           '���ͺ�
    '-----------------------------------------
    strModality As String       '�豸
    lngID As Long               '����ID
    strName As String           '��������
    strSex As String            '�Ա�
    strAge As String            '����
    strBirthDate As String      '��������
    
    strStudyUID As String       '���uid
    strSeriesUID As String      '����UID
    strInstanceUID As String    'ʵ��UID
    lngSeriesNo As Long         '���к�
    lngImgNo As Long            'ͼ���
    strReceiveFullTime As String    '��������
    
    strInstitution As String    '��������
    
    strSeriesDes As String      '��������
    strImageDes As String       'ͼ������
    
    strMediaEncode As String    '��ý�����
    lngMediaTag As Long         '��ý����
    lngMediaLen As Long         '��ý�峤��
    strDeviceNo As String       '�豸��
End Type