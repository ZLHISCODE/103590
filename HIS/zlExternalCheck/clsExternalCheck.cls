VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExternalCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, ByVal int���� As Integer)
'���ܣ��ӿڳ�ʼ��
'������cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    Set gcnOracle = cnOracle
End Sub


Public Function ExternalCheck(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int���� As Integer, ByVal lngID As Long, ByVal strName As String, ByRef strXml As String, ByRef strVerificationCode As String, ByRef blnGranted As Boolean) As Boolean
'���ܣ�����������Ȩ���
'��Σ� lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'       int���ϣ����ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'       lngID���������ù��ܵĲ˵�ID(��������Ŀ¼.ID)
'       strName���������ù��ܵĲ˵�����(��������Ŀ¼.�˵�����)
'       strXml:��չ����

'���Σ� strVerificationCode������ҽԺ���ƣ�����ʱ����Ƿ��뵱ǰҽԺ����һ�£�����������һ������ͳһ�ֲ�ͬ�û���Ȩ
'       blnGranted����Ȩ����ȷ

End Function