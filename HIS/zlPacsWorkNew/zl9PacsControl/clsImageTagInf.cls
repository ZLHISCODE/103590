VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageTagInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public EncoderName As String '��������
Public CaptureTime As String '�ɼ�ʱ��
Public RecordTimeLen As Long '¼�Ƴ���
Public VideoFile As String   '��Ƶ�ļ�����
Public Tag As Variant        '��� VIDEOTAG: ��ʾ��Ƶ��IMGTAG:��ʾͼ��
Public FilePath As String    '��¼�ļ��Ĵ洢λ��
Public ReportImage As String   '����ͼ������ �� ���Ǳ���ͼ��0���½��������뱨��ͼ��1-n ������ͼ
Public IsChanged As Boolean   'ͼ���Ƿ񱻴���δ���棬����ͼ��������ʹ��