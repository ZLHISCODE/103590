VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScFilterCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Name As String
'Public FilterType As TFilterType
Public SelectWay As TSelectWay  'ѡ��ʽ��0-��ѡ��1-��ѡ
Public ExtProperty As String  '��չ����
'Public SourceType As Long
Public DataFrom As String       '������Դ����
'Public IsCustom As Boolean
Public CustomScript As String