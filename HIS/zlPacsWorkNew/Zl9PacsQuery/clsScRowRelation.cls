VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScRowRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

  

Public TiggerData As String         '��������
Public Icon As String               'ͼ��
Public IconPerformCol As String     'ͼ�������
Public IsStateIcon As Boolean       '�Ƿ�״̬ͼ��
Public RowFontColor As Long         '��������ɫ
Public RowBackColor As Long         '�б�����ɫ
Public CellFontColor As Long        '��Ԫ��������ɫ
Public CellBackColor As Long        '��Ԫ�񱳾���ɫ
Public ColorPerformCol As String    '��ɫ������
Public FlickerTimeOut As Long       '��ɫ��ʱ��˸ʱ��
Public TimeOutReferCol As String    '��ʱʱ��ο���
