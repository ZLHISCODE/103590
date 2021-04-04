VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceErrPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'Ԥ�������
'�����ӿ�:
'    1.zlInitCommon-��ʼ���ӿ�
'    2.zlRefreshData-����ˢ������
'    3.zlGetForm -��ȡ���������
'����:���˺�
'����:2019-11-27 16:52:50
'*********************************************************************************************************************************************
Private mlngfrmMainHwnd As Long '�����ڶ���
Private mfrmSendCardAndDepositErrPage As frmSendCardAndDepositErrPage
Attribute mfrmSendCardAndDepositErrPage.VB_VarHelpID = -1
Private mlngModule As Long

Public Function zlInitCommon(ByVal frmMain As Object, ByVal lngSys As Long, ByVal cnOracle As ADODB.Connection, _
    ByVal lngModule As Long, intӦ�ó��� As Byte, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '     frmMain-���õ�������
    '     lngModule-ģ���
    '     intӦ�ó��� -1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ �Ǽ�;4-ԤԼ�ҺŽ���
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2020-01-04 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser: mlngModule = lngModule
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    zlInitCommon = mfrmSendCardAndDepositErrPage.zlInit(frmMain, intӦ�ó���, lngModule)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetForm() As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    '����:��ȡ������Ԥ���������,�����������
    '����:���˺�
    '����:2020-01-04 12:00:14
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetForm = mfrmSendCardAndDepositErrPage
End Function
Private Sub Class_Initialize()
    Set mfrmSendCardAndDepositErrPage = New frmSendCardAndDepositErrPage
    glngInstanceCount = glngInstanceCount + 1
End Sub
Public Sub zlRefreshData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������
    '���:objPati-������Ϣ��
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-01-04 16:03:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmSendCardAndDepositErrPage.zlRefreshData
     
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmSendCardAndDepositErrPage Is Nothing Then Unload mfrmSendCardAndDepositErrPage
    Set mfrmSendCardAndDepositErrPage = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
 