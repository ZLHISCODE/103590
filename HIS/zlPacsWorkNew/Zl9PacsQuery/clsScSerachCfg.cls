VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScSerachCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Name As String
Public InputType As TInputType  '¼������0-����¼�룬1-�б��Ϸ��������¼�룬2-���߶���
Public ControlType As Long  '����¼��Ŀؼ�����
Public ExtProperty As String  '��չ����
'Public SourceType As Long
Public Default As String    'Ĭ��ֵ
Public DataFrom As String   '������Դ����
'Public IsCustom As Boolean