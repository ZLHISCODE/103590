VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function getReportEx(objClsEx As Object) As Object
'���ܣ���ȡReport��Ƭ�����¼�����
    Dim frmTmp As frmReport
    
    Set getReportEx = New frmReport
    Set getReportEx.mclsReport = objClsEx
    Set objClsEx.mfrmReportEx = getReportEx
End Function



Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿڳ�ʼ��
'������cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    Set gcnOracle = cnOracle
End Sub

