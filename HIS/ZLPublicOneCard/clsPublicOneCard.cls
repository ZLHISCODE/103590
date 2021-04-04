VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicOneCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'����:һ��ͨ�漰��������ݶ���(������ȡһ��ͨ�ķ��������ݼ���)
'�漰�ķ�������:
'  1.��������
'  1.1  zlInitComponents-��ʼ������(��һ�������ȳ�ʼ�����������ʹ���лᷢ���쳣)
'  1.2  ErrCenter-����������
'  1.3  SaveErrLog-���������־
'  1.3  WritDebugLog-д������־
'  1.4  SetPara-��������
'  1.5  GetPara-��ȡ����
'  1.7  GetPrivFunc-��ȡȨ�޴�
'  2.�����������
'  2.1 zlGetPatiIDFromBedNumber:���ݴ�λ�ţ���ȡ����ID
'  2.2 zlGetPatiInforFromPatiID:���ݲ���ID��ȡ������Ϣ����
'  2.3 zlGetPatiInforFromXML:����XML������ȡ������Ϣ����
'  3. ���������
'  3.1 zlGetCards-��ȡ��Ч�Ŀ����
'�漰�Ĺ�������:
'  5.1 objThirdSwap-��ȡ�����ӿڶ���
'  5.2 objOneCardObject-һ��ͨ����
'  5.3 objDataBase-�漰���ݿ��������
'����:���˺�
'����:2018-12-03 11:06:14
'*********************************************************************************************************************************************
Private mcnOracle As ADODB.Connection
Private mstrDBUser As String, mlngModul As Long
Private mfrmMain As Object
Private mobjDataBase As clsDataBase
Private mobjOneCardObject As clsOneCardDataObject
Private mobjThirdSwap As clsThirdSwapInterface  '�������׽ӿ�
Private mblnIsObjRegisterAlone As Boolean

Public Property Get objThirdSwap() As clsThirdSwapInterface
    If mobjThirdSwap Is Nothing Then
        Set mobjThirdSwap = New clsThirdSwapInterface
        Call mobjThirdSwap.InitCommon(mcnOracle)
    End If
    Set objThirdSwap = mobjThirdSwap
End Property
Public Property Get objOneCardObject() As clsOneCardDataObject
    Call zlGetOneCardDataObject(mcnOracle, mobjOneCardObject)
    Set objOneCardObject = mobjOneCardObject
End Property
Public Function zlIsExistOraConnect() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ����Oracle����
    '����:���˺�
    '����:2018-12-06 18:17:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    zlIsExistOraConnect = mobjDataBase.zlIsExistOraConnect
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlInitComponents(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    ByVal cnOracle As ADODB.Connection, Optional blnDeviceSet As Boolean = False, Optional strExpand As String, _
    Optional ByVal blnIsObjRegisterAlone As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '���: frmMain-���õ�������
    '        lngModule-HIS����ģ���
    '       lngSys-�����ϵͳ��
    '       strDBUser-���ݿ��û���
    '       cnOracle -HIS/��������
    '       blnDeviceSet-�豸���õ��ó�ʼ��
    '       strExpand-��չ��Ϣ(��ѡ����:�����ID-����ʱ,��ʾȫ����ʼ��,����ʱ,ֻ��ʼ��ָ���Ľӿ�)
    '       blnIsObjRegisterAlone-�Ƿ�ʹ�ö�����ע�Ჿ��(True:ʹ��:zlRegisterAlone.DLL,����ʹ��zlRegister.dll)
    '����:��������True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2011-05-23 11:49:39
    '˵��:����ǵ������ӿ�������ȡ����,�����������ĵĲ����д�������.
    '    1.��������ӿڴ���ǩ��������������ڴ˽ӿ��н���ǩ��.
    '    2.����������ӿڶ�ȡ����,��Ҫ�ȳ�ʼ��.
    '������:
    '    1.  ���������շѡ��Һ�ʱ���ñ��ӿ�
    '    2.  ����ҽ��վ����ʿվ��ҽ��վ��ʱ���ñ��ӿ�
    '    3.  �������ӿ���Ҫ������ȡ����ʱ,�����ȳ�ʼ��.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    glngSys = lngSys: gstrDBUser = strDBUser: Set mfrmMain = frmMain
    mstrDBUser = strDBUser: mlngModul = lngModule: glngModul = lngModule
     
    Set mcnOracle = cnOracle
    If gcnOracle Is Nothing Then Set gcnOracle = cnOracle
    
    If mblnIsObjRegisterAlone <> blnIsObjRegisterAlone Then
        Set gobjOneDataBase = Nothing
    End If
    
    mblnIsObjRegisterAlone = blnIsObjRegisterAlone
    Call zlInitPublicVar '��ʼ��ȫ�ֱ���
    If zlGetOneDataBase(cnOracle, mobjDataBase, blnIsObjRegisterAlone) = False Then Exit Function
    If zlGetOneCardDataObject(cnOracle, mobjOneCardObject) = False Then Exit Function
    
    If mobjThirdSwap Is Nothing Then Set mobjThirdSwap = New clsThirdSwapInterface
    Call mobjThirdSwap.InitCommon(cnOracle, mblnIsObjRegisterAlone)   '��ʼ�������ӿڱ���
    gstrUnitName = mobjDataBase.GetUnitName
    zlInitComponents = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function zlIsMobileNo(ByVal strInput As String, Optional ByRef strRutType As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '����:�жϴ�����Ƿ�Ϊ�ֻ���
    '���:strInput-�ֻ���
    '����:strRutType-��ѯ���:0-�ɹ�;1-������Ч�Ŷ�;2-���볤�Ȳ���
    '����:True-�������Ϊ�ֻ���;False-������벻Ϊ�ֻ���
    '����:������
    '����:2017-1-25
    '---------------------------------------------------------------------------------------------
    zlIsMobileNo = mobjOneCardObject.zlIsMobileNo(strInput, strRutType)
End Function

Public Function zlGetPatiIDFromBedNumber(ByVal lng����ID As Long, _
    ByVal str���� As String, Optional ByRef lng��ҳID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݴ��Ż�ȡ����ID
    '����:lng��ҳID-���ص�ǰ���ŵ���ҳID
    '����:�ɹ����ز���ID,���򷵻�False
    '����:���˺�
    '����:2012-09-19 15:50:18
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlGetPatiIDFromBedNumber = mobjOneCardObject.zlGetPatiIDFromBedNumber(lng����ID, str����, lng��ҳID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlGetPatiID(ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng����ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����ҽ�����Ϳ���,��ȡ��Ӧ�Ĳ���ID
    '���:strCardType-�����,���Ϊ����,��Ϊ�����ID,���Ϊ�ַ�,��Ϊ�������
    '       strCardNo-����
    '       blnNotShowErrMsg-����ʾ�������ʾ��Ϣ
    '       frmMain-���õ�������
    '       objCtl-���õĿؼ�
    '       blnShowMergePati-�����ֶ�����������Ĳ���ʱ,�Ƿ���ʾ�ϲ����ܰ�ť
    '       blnOnlyContractPati-ǩԼ����
    '       blnUserCancel-ѡ�����У��û�ѡ����ȡ��
    '       lngShowCardNoTypeID-���˳���������Ϣʱ������ѡ��������ʾ�Ŀ��ŵĿ����ID,0-��ʾ����ʾ���ţ�>0��ʾ��ʾָ����������ID
    '       blnNotCheckValidDate-�Ƿ�Կ���ֹʹ��ʱ����м��,true-�������ֹʹ��ʱ��,false-���
    '����:strErrMsg-���صĴ�����Ϣ
    '       lng����ID-���صĲ���ID
    '       strCardPass-���ؿ��ŵ�����
    '       lngCardTypeID-���ؿ����ID(0��ʾ����ȷ�������ID)
    '����:��ȡ����ID�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-14 17:07:51
    '˵��:ֻ�д���ҽ�����Ĳŵ��ô˺���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    zlGetPatiID = mobjOneCardObject.zlGetPatiID(strCardType, strCardNo, blnNotShowErrMsg, lng����ID, _
      strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, _
      blnCertificate, blnUserCancel, lngShowCardNoTypeID, blnNotCheckValidDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCards(ByVal bytType As Byte) As Cards
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ч�Ŀ�����
    '���:bytType-0-����ҽ�ƿ�;
    '             1-���õ�ҽ�ƿ�,
    '             2-���д��������˻���������
    '             3-���õ������˻���ҽ�ƿ�
    '����:
    '����:�ɹ����ر���֧�ֵĿ�����
    '����:���˺�
    '����:2013-10-23 17:43:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set zlGetCards = mobjThirdSwap.zlGetCards(bytType)
End Function


Public Function zlGetPatiInforFromPatiID(ByVal lng����ID As Long, ByRef objPati As clsPatiInfor, _
    ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String) As Boolean
    
    zlGetPatiInforFromPatiID = mobjOneCardObject.zlGetPatiInforFromPatiID(lng����ID, objPati, strErrMsg, strOtherName, strOtherValue)
End Function

Public Function zlGetPatiInforFromXML(ByVal strPatiXml As String, ByRef strErrMsg_Out As String, ByRef objPatiInfor_Out As clsPatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��XML�л�ȡ������Ϣ
    '���:strPatiXml-������ϢXML
    '
    '����:objPatiInfor_Out-���ز�����Ϣ����
    '      strErrMsg_Out-���ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-05 14:29:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiInforFromXML = mobjOneCardObject.zlGetPatiInforFromXML(strPatiXml, strErrMsg_Out, objPatiInfor_Out)
End Function
    

Public Function ErrCenter() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '����:���˺�
    '����:2018-12-05 11:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
   ErrCenter = mobjDataBase.ErrCenter
End Function

Public Sub SaveErrLog()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:д��־
    '����:���˺�
    '����:2018-12-05 11:33:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Sub
    End If
    Call mobjDataBase.SaveErrLog
End Sub
Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True, Optional ByVal lngDeptID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngSys=ʹ�øò�����ϵͳ��ţ���100
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnSetup=����ģ���Ƿ��в�������Ȩ��
    '      lngDeptID=���ż�����
    '���أ������Ƿ�ɹ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    SetPara = mobjDataBase.SetPara(varPara, strValue, glngSys, lngModual, blnSetup, lngDeptID)
End Function
Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer, Optional ByVal lngDeptID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      lngSys=ʹ�øò�����ϵͳ��ţ���100
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
    '      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
    '      arrControl=�ؼ����飬��Array(Me.Text1, Me.CheckBox1)�����ں����ڲ��Զ�������Ӧ�ؼ�����ʾ��ɫ���Ƿ��ֹ���á�
    '      blnSetup=����ģ���Ƿ��в�������Ȩ��
    '      intType=���ز��������ز�������
    '      lngDeptID=���ż����������ݲ��Ŷ�ȡ
    '���أ�����ֵ���ַ�����ʽ
    '����:���˺�
    '����:2018-12-05 14:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    GetPara = mobjDataBase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType, lngDeptID)
End Function

Public Function GetPrivFunc(lngSys As Long, lngProgId As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ص�ǰ�û����е�ָ������Ĺ��ܴ�
    '������lngSys     ����ǹ̶�ģ�飬��Ϊ0
    '      lngProgId  �������
    '���أ��ֺż���Ĺ��ܴ�,Ϊ�ձ�ʾû��Ȩ��
    '����:���˺�
    '����:2018-12-05 15:01:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    GetPrivFunc = mobjDataBase.GetPrivFunc(lngSys, lngProgId)
End Function

Public Sub WritDebugLog(ByVal strDev As String, strInput As String, strOutPut As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��¼��־
    '����:���˺�
    '����:2018-12-05 11:35:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Sub
    End If
    Call mobjDataBase.WritLog(strDev, strInput, strOutPut)
End Sub


Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Call zlInitCommLib  '��ʼ��CommLib����
End Sub

Private Sub Class_Terminate()

    Err = 0: On Error Resume Next
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mobjOneCardObject Is Nothing Then Set mobjOneCardObject = Nothing
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If Not mobjThirdSwap Is Nothing Then Set mobjThirdSwap = Nothing
    Call zlReleaseResources '�ͷ���Դ
End Sub

 