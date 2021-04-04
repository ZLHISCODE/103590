VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendCardAndDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'Ԥ�������
'�����ӿ�:
'    1.zlInitCommon-��ʼ���ӿ�
'    2.zlSetFontSize-���ý��������С
'    3.zlGetForm -��ȡ���������
'    4.zlRecalcCardFee-�¼��㿨���ã����Ա�/����/ҽ�Ƹ��ʽ/�ѱ�/�Է�֤�Ÿı�����
'    5.zlSetCardNo-�ڲ��������ˢ����������¿�,����Ҫ���ÿ�����Ϊ�������ţ���Ҫ���øÿ���
'    6.zlGetSendCard-��ȡ��ǰ�ķ���������
'    7.zlSetUnitInfo-���ýɿλ��Ϣ(������λ���˺ŵ�������ɺ���Ҫ����)
'    8.zlSetInsureInfo:����ҽ����Ϣ(ҽ��������֤�����)
'    7.zlClearControlInfo -�������ؼ���Ϣ
'���ݱ���ӿ�
'    1.zlSaveDataBeforCheckIsValid-������ݵĺϷ���:�ڱ���ǰ��Ҫ����
'    2.zlSaveData-ִ�����ݱ������
'    3.zlSaveDataAfter-���ݱ����ִ��(��Ҫ������ˢ��Ʊ��,��տؼ���Ϣ��)
'�����¼�:
'    1.RequestRefreshPatiInf-�������¸���XML��ʽ��������ݣ�ˢ�²�����Ϣ
'    2.InputOver-��������¼�(��ʾ���һ��������ɣ��Ա�����ת����һ����������
'    3.Activate-�Ӵ��弤��
'    4.ExcuteQRCodePayment-Ҫ��ִ��ɨ�븶����(��Ҫ���ⲿֱ�ӵ���ȷ����ť0
'    5.zlSetFocus-��궨λ
'    6.ControlGotFocus-�������ؼ�ʱ����
'��������
'    1.RealName-���õ�ǰ�����Ƿ������ʵ����֤(ʵ����֤����Ҫ��ֵ)
'    2.Width-�������
'    3.Height-����߶�
'����:���˺�
'����:2019-11-27 16:52:50
'*********************************************************************************************************************************************
Private mlngfrmMainHwnd As Long '�����ڶ���
Private WithEvents mfrmSendCardAndDeposit As frmSendCardAndDeposit
Attribute mfrmSendCardAndDeposit.VB_VarHelpID = -1
Private mlngModule As Long
'------------------------------------------------------------------------------------------------------
'�����¼�
Public Event Activate() '�Ӵ��弤��
Public Event RequestRefreshPatiInf(ByVal strCardNo As String, ByVal strPatiInfoXML As String)
Public Event InputOver()    '�������
Public Event ExcuteQRCodePayment() 'ִ��ɨ�븶
Public Event ExcuteReadQRCode()
Public Event ControlGotFocus(objControl As Object)

Public Function zlInitCommon(ByVal frmMain As Object, ByVal lngSys As Long, ByVal cnOracle As ADODB.Connection, _
    ByVal lngModule As Long, ByVal bln����Ԥ�� As Boolean, ByVal blnסԺԤ�� As Boolean, _
    ByVal lngCardTypeID As Long, blnAllowSendCard As Boolean, ByVal blnAllowBoundCard As Boolean, ByVal blnAllowInsureAccDeposit As Boolean, _
    Optional btQRCodePay As Object, Optional objShowTotalMoneyControl As Object, Optional blnView As Boolean = False, _
    Optional ByVal blnCancel As Boolean, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '     frmMain-���õ�������
    '     lngModule-ģ���
    '     btQRCodePay-ɨ�븶��ť
    '     objShowTotalMoneyControl-��ʾ���ܶ�ؼ�:lable��Text
    '     bln����Ԥ��-�Ƿ������Ԥ��
    '     blnסԺԤ��-�Ƿ��סԺԤ��
    '     lngSendCardTypeID-��ǰ�������ID:����0ʱ���������blnAllowSendCard��blnAllowBoundCard-��Ч
    '     blnAllowSendCard-��������
    '     blnAllowBoundCard-�����󶨿�
    '     blnView-�Ƿ�鿴
    '     strPrivs-��ǰ����ģ��Ȩ��
    '     blnAllowInsureAccDeposit-�Ƿ�����ҽ���˻���Ԥ��
    '     blnCancel-��ǰ�Ƿ����ϲ���
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser: mlngModule = lngModule
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    Call gclsInsure.InitOracle(gcnOracle)
    zlInitCommon = mfrmSendCardAndDeposit.zlInit(frmMain, lngModule, bln����Ԥ��, blnסԺԤ��, lngCardTypeID, blnAllowSendCard, blnAllowBoundCard, blnAllowInsureAccDeposit, _
        btQRCodePay, objShowTotalMoneyControl, blnView, , blnCancel)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlSetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:���˺�
    '����:2014-04-09 11:46:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmSendCardAndDeposit Is Nothing Then Exit Sub
     mfrmSendCardAndDeposit.zlSetFontSize (bytSize)
End Sub

Public Sub zlSetFocus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��궨λ
    '���
    '����:���˺�
    '����:2020-01-13 17:53:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmSendCardAndDeposit Is Nothing Then Exit Sub
    mfrmSendCardAndDeposit.zlSetFocus
End Sub


Public Function zlRecalcCardFee(ByVal objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼��㿨����Ϣ
    '���:objPati-���˶���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-23 17:44:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlRecalcCardFee = mfrmSendCardAndDeposit.zlRecalcCardFee(objPati)
End Function

Public Function zlSetCardNo(ByVal strCardNo As String, objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿ��Ÿ������ı���
    '���:objPati-������Ϣ��
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 18:53:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSetCardNo = mfrmSendCardAndDeposit.zlSetCardNo(strCardNo, objPati)
End Function
Public Sub zlSetUnitInfo(ByVal str��λ�ʺ� As String, ByVal str�ɿλ As String, ByVal str��λ������ As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ�λ�˺�
    '����:���˺�
    '����:2019-11-26 13:37:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmSendCardAndDeposit.zlSetUnitInfo(str��λ�ʺ�, str�ɿλ, str��λ������)
End Sub
Public Sub zlSetInsueInfo(ByVal int���� As Integer, ByVal cur�˻���� As Currency, ByVal strҽ���� As String, ByVal str���� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ����Ϣ(ҽ��������֤�����)
    '����:���˺�
    '����:2019-11-26 13:37:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmSendCardAndDeposit.zlSetInsueInfo(int����, cur�˻����, strҽ����, str����)
End Sub
Public Sub zlClearControlInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ؼ���Ϣ
    '����:���˺�
    '����:2019-11-26 13:37:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmSendCardAndDeposit.zlClearControlInfo
End Sub


Public Function zlGetSendCard(ByRef objSendCard_Out As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�ķ�������
    '���:
    '����:objSendCard_Out-���ص�ǰ�����Ķ���
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 15:13:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetSendCard = mfrmSendCardAndDeposit.zlGetSendCard(objSendCard_Out)
End Function

Public Function zlSaveDataBeforCheckIsValid(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo, _
    Optional ByVal bln�Զ�ʶ�������֤ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '���:objPati-������Ϣ��
    '     blnNewPati-�Ƿ��²���
    '     bln�Զ�ʶ�������֤-�Ƿ��Զ�ʶ�������֤��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 13:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSaveDataBeforCheckIsValid = mfrmSendCardAndDeposit.zlSaveDataBeforCheckIsValid(blnNewPati, objPati, bln�Զ�ʶ�������֤)
End Function
 

Public Function zlSaveData(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:objPati-������Ϣ��
    '     blnNewPati-�Ƿ��²���
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 13:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSaveData = mfrmSendCardAndDeposit.zlSaveData(blnNewPati, objPati)
End Function

Public Function zlSaveDataAfter() As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ����ִ��(��Ҫ������ˢ��Ʊ��,��տؼ���Ϣ��)
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 15:08:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlSaveDataAfter = mfrmSendCardAndDeposit.zlSaveDataAfter
End Function

Private Sub Class_Initialize()
    Set mfrmSendCardAndDeposit = New frmSendCardAndDeposit
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmSendCardAndDeposit Is Nothing Then Unload mfrmSendCardAndDeposit
    Set mfrmSendCardAndDeposit = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
Private Sub mfrmSendCardAndDeposit_Activate()
    RaiseEvent Activate
End Sub
Public Function zlGetForm() As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    '����:��ȡ������Ԥ���������,�����������
    '����:���˺�
    '����:2014-04-09 12:00:14
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetForm = mfrmSendCardAndDeposit
End Function
Public Property Get RealName() As Boolean
       RealName = mfrmSendCardAndDeposit.RealName
End Property
Public Property Let RealName(ByVal vNewValue As Boolean)
     mfrmSendCardAndDeposit.RealName = vNewValue
End Property

Public Property Get Width() As Long
       Width = mfrmSendCardAndDeposit.GetWidth
End Property

Public Property Get Height() As Long
       Height = mfrmSendCardAndDeposit.GetHeight
End Property

Private Sub mfrmSendCardAndDeposit_ControlGotFocus(objControl As Object)
    RaiseEvent ControlGotFocus(objControl)
End Sub

Private Sub mfrmSendCardAndDeposit_ExcuteQRCodePayment()
    RaiseEvent ExcuteQRCodePayment
End Sub

Private Sub mfrmSendCardAndDeposit_ExcuteReadQRCode()
   RaiseEvent ExcuteReadQRCode
End Sub

Private Sub mfrmSendCardAndDeposit_InputOver()
    RaiseEvent InputOver
End Sub

Private Sub mfrmSendCardAndDeposit_RequestRefreshPatiInf(ByVal strCardNo As String, ByVal strPatiInfoXML As String)
    RaiseEvent RequestRefreshPatiInf(strCardNo, strPatiInfoXML)
End Sub

