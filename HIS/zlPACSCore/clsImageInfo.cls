VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public InstanceUID As String        'ͼ���UID�����ؼ��֣�
Public ImageNo As String            'ͼ��� (0020,0013) : Image Number
Public ImageName As String          'ͼ���ļ���������ֱ�Ӵӱ���Ŀ¼�򿪵�ͼ����Ҫ��д��ʵ���ƣ����ڴӹ���վ�򿪵ģ���дͼ��UID��
Public blnSelected As Boolean       '�Ƿ�ѡ�񣬽���ZLShowSeriesInfos��ʹ�á�
Public blnDisplayed As Boolean      '�Ƿ�����ʾ������ZLShowSeriesInfos��ʹ�á�
Public AcquisitionTime As String    '�ɼ�ʱ�� (0008,0022) : Acquisition Date+ (0008,0032) : Acquisition Time
Public ImageTime As String          'ͼ��ʱ�� (0008,0023) : Image Date  + (0008,0033) : Image Time
Public SliceThickness As String     '��� (0018,0050): Slice Thickness
Public ImagePositionPatient As String       'ͼ��λ�ò��� (0020,0032) : Image Position (Patient)
Public ImageOrientationPatient As String    'ͼ������ (0020,0037) : Image Orientation (Patient)
Public FrameOfReferenceUID As String        '�ο�֡UID (0020,0052) : Frame of Reference UID
Public SliceLocation As String      '��Ƭλ�� (0020,1041) : Slice Location
Public Rows As String               '���� (0028,0010) : Rows
Public Columns As String            '���� (0028,0011) : Columns
Public PixelSpacing As String       '���ؾ��� (0028,0030): Pixel Spacing
Public int3DLabelIndex  As Integer     '��ά����ж�λ�ߵ�����
Public blnPrinted As Boolean        '�Ƿ��Ѿ���ӡ��Ƭ
Public ImgTag As Long               'ͼ�����ͣ� 0��ͼ��1����֡ͼ��2��¼��3��¼������ʹ�ã�
