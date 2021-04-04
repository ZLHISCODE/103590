VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchiveMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'���ģ�����
Private mclsArchive As zlMedRecPage.clsArchive

'��ҳ�Ƿ��Ǵ�״̬��ֻ������
Public Property Get IsOpen() As Boolean
    IsOpen = mclsArchive.IsOpen
End Property
'��ҳ�Ƿ�ϲ���,ֻ������
Public Property Get IsMerged() As Boolean
    IsMerged = mclsArchive.IsMerged
End Property
'��Ժ��Ҫ����Ƿ����ı�,ֻ������
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsArchive.IsDiagChange
End Property
'IsDiagInput:ָ����ϵ���ҳ����Ƿ�����
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsArchive.IsDiagInput
End Property

Public Sub InitArchiveMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, Optional ByVal lngModule As Long, Optional ByRef objMip As Object)
'���ܣ���ʼ��
'������objMip=��Ϣ����
    Call mclsArchive.InitArchiveMedRec(cnMain, lngSys, lngModule, objMip)
End Sub

Private Sub Class_Initialize()
    Set mclsArchive = New zlMedRecPage.clsArchive
End Sub

Private Sub Class_Terminate()
    If Not mclsArchive Is Nothing Then
        Set mclsArchive = Nothing
    End If
End Sub

Public Function zlGetForm(ByVal intType As Integer) As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
'������intType 0-���ﲡ������ 1-סԺ��������
    Set zlGetForm = mclsArchive.zlGetForm(intType)
End Function

Public Function zlRefresh(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal blnMoved As Boolean) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
'������intType 0-���ﲡ������ 1-סԺ��������
'      lng��ҳID :סԺ��ҳ-��ҳID�� ������ҳ-�Һ�ID
'      blnMoved �Ƿ�ת��
    zlRefresh = mclsArchive.zlRefresh(intType, lng����ID, lng��ҳID, blnMoved)
End Function