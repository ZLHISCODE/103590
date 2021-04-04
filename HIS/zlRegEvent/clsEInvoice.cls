VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*********************************************************************************************************************************************
'����:����Ʊ�ݽ�����ؽӿ�
'һ��������ؽӿ�
'   1.zlReReadParaAndVar:���¶�ȡ����ֵ������ֵ
'   2.zlInitialize:��ʼ���ӿ�
'   3.zlTerminate:��ֹ�ӿ�
'   4.zlEInvoiceSet:���ӷ�Ʊ�豸��������ýӿ�
'����Ʊ���ж���ؽӿ�
'   1.zlIsStartEInvoice:�ж��Ƿ������˵���Ʊ�ݵ�
'   2.zlIsStartEInvoiceFromBalanceID:�ж�ָ��������Ϣ�Ƿ������˵��ӷ�Ʊ
'   3.zlIsHisManagerInvoice:�ж��Ƿ�HIS����ֽ�ʷ�Ʊ
'����ҵ������ؽӿ�
'   1.zlCreateEInvoice-���ߵ���Ʊ��
'   2.zlCancelEInvoice-���ϵ���Ʊ��
'     zlCancelEInvoiceFromBalanceInfor-���ݽ��ʶ���ִ�е���Ʊ��Ʊ�����ϲ���
'   3.zlPrintEInvoice-��ӡ����Ʊ��A4�ӿ�
'   4.zlTurnPaperInvoice-����ֽ��Ʊ��
'   5.zlCancelPaperInvoice-����ֽ��Ʊ��
'   6.zlGetNextInvoiceNo-��ȡ��һ����Ч�ķ�Ʊ��
'   7.zlGetTranPaperInvoiceModule-��ȡֽ�ʷ�Ʊ�Ļ�����ʽ
'   8.zlPrintNotice-��ӡ��֪��
'�ġ���ȡҵ������
'   1.zlGetEInvoiceInforFromBalanceID-���ݽ���ID,��ȡ����Ʊ����Ϣ(���ؼ�¼��)
'   4.zlGetEInvoiceSwapCollectFromBalanceID-���ݽ���ID����������Ʊ����ؽ�����Ϣ��
'����:���˺�
'����:2020-03-19 10:43:04
'*********************************************************************************************************************************************
Private mcnOracle As ADODB.Connection, mfrmMain As Object
Private mobjEInvoice As Object
Private mlngMoudle As Long, mstrPrivs As String
Private mblnUseChargeInvoice As Boolean
Private mlngCardInvoiceId As Long


Public Property Let lngMoudle(ByVal varValue As Long)
    mlngMoudle = varValue
End Property

Public Property Get lngMoudle() As Long
    lngMoudle = mlngMoudle
End Property

Public Property Set frmMain(ByVal varValue As Object)
    Set mfrmMain = varValue
End Property

Public Property Get frmMain() As Object
    Set mfrmMain = frmMain
End Property

Public Property Get blnPrintCardInvoice() As Long
    blnPrintCardInvoice = mblnUseChargeInvoice
End Property

Public Property Get lngCardInvoiceId() As Long
    lngCardInvoiceId = mlngCardInvoiceId
End Property


Private Function GetEInvoiceObject(ByRef objEInvoice As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ʊ�ݶ�Ӧ�Ӷ���
    '���:
    '����:objEinvoice-����Invoice����
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 14:32:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mobjEInvoice Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
        If Err <> 0 Then
            MsgBox "�����ڿ��õĵ���Ʊ�ݽӿڲ���(zlPublicExpense.clsPubEInvoice)������ϵͳ����Ա��ϵ,��ϸ�Ĵ�����ϢΪ:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrProductName
            Exit Function
        End If
    End If
    Set objEInvoice = mobjEInvoice
    GetEInvoiceObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetNextInvoiceNo(ByVal frmMain As Object, ByRef strInvoiceNO_Out As String, _
    Optional cllPati As Collection, Optional lng����ID As Long, Optional ByVal byt���� As Byte = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ���ŷ�Ʊ��
    '���:byt���ϣ�1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-���￨
    '     cllPati-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����):Ŀǰ���ã���Ҫ���Ժ���չHIS����Ʊ��
    '     lng����ID-Ʊ������ID,�Ժ�HIS����Ʊ�ݿ���Ҫ��,�����ô�
    '����:strInvoiceNO_Out-������һ����Ч��Ʊ��
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-23 16:53:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objEInvoice As Object
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlGetNextInvoiceNo = objEInvoice.zlGetNextInvoiceNo(frmMain, byt����, strInvoiceNO_Out, cllPati, lng����ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReReadParaAndVar() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¶�ȡ����������ֵ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-19 11:30:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlReReadParaAndVar = objEInvoice.zlReReadParaAndVar
End Function
Public Function zlInitialize(ByVal frmMain As Object, Optional ByVal byt���� As Byte = 4, Optional ByVal blnDeviceSet As Boolean, Optional ByVal lngModule As Long = 1111) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿڲ���
    '���:
    '    byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�
    '    frmMain -���õ�������
    '    cnOracle -���ݿ����Ӷ���
    '    lngSys -��ǰ����ϵͳ���
    '    lngModule-��ǰ����ģ���:
    '    blnDeviceSet -�豸���õ��õĳ�ʼ��
    '����:strExtend-��չ����,����,�ִ�Ϊ��
    '����:��ʼ���ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '    1.ʹ�ñ�����ǰ,�����ȵ��ñ��ӿڽ��г�ʼ��
    '    2.��ʼ���ӿ�,��HIS����ģ��ʱ����(���磺�����շѹ�������),���ڱ��ӿ�����ɵ���������Ʊ�ݴ�ӡ�ӿڵĳ�ʼ������¼�������ӵ�ʱʹ��
    '    3.ģ��ţ�1111=����Һ�,1121=�����շ�,1124=���ղ������,1137=סԺ����,1107=ҽ�ƿ����Ź���;
    '           1103=Ԥ�������;1504=Ʊ��������;1501=Ʊ��ʹ�ü��;2128=�����ʹ���
    '������:
    '   Ʊ������Ʊ������?�Һ�?�շ�?����?�������?�豸����?���������ʵ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    Dim strExtend As String
    Set mfrmMain = frmMain: mlngMoudle = lngModule
    
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlInitialize = objEInvoice.zlInitialize(frmMain, byt����, gcnOracle, glngSys, lngModule, blnDeviceSet, strExtend)
End Function

Public Function zlTerminate() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ֹ�ӿ�
    '����:��ֹ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   ��ֹ�ӿ�,��HIS�˳�ģ��ʱ����(���磺�˳��շѹ�������),���ڱ��ӿ�����ɵ���������Ʊ�ݽӿڵ���Դ�ͷš��Ͽ����ӵ�ʱʹ��
    '������:
    '    Ʊ������Ʊ������,�Һ�,�շ�,����,�������,�豸����,���������ʵ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlTerminate = objEInvoice.zlTerminate()
    Set objEInvoice = Nothing
End Function


Public Function zlIsStartEInvoice(ByVal byt���� As Byte, Optional ByVal intInsure As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ������˵���Ʊ�ݵ�
    '���:byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�
    '     intInsure-��ǰҽ��:0-������ͨ����;>0��ʾ����ҽ��
    '����:
    '����:���÷���true,���򷵻�False
    '����:���˺�
    '����:2020-03-19 14:32:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlIsStartEInvoice = objEInvoice.zlIsStartEInvoice(byt����, intInsure)
    Set objEInvoice = Nothing
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlEInvoiceSet(ByVal frmMain As Object, Optional ByRef strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ӷ�Ʊ�豸��������ýӿ�
    '���:
    '    frmMain -���õ�������
    '    strExtend-��չ����,����,�ִ�Ϊ��
    '����:
    '    strExtend-��չ����,����,�ִ�Ϊ��
    '����:���óɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '    һ����Ե���Ʊ���豸����Ҫ�Ǵ�ӡ�豸������Ʊ������������õ�)һЩ���ý�������,���磺��ӡ��,�˿ںš�IP���û����������ַ�ȵ�����
    '������:
    '   �ڸ�ҵ��ģ��(������շѡ����ʵ�)�еĲ�������  : �豸���á�>����Ʊ���豸, ��������á���ťʱ,���ñ��ӿ�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlEInvoiceSet = objEInvoice.zlEInvoiceSet(frmMain, strExtend)
End Function

Public Function zlIsHisManagerInvoice(Optional ByVal byt���� As Byte = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ�HIS����ֽ�ʷ�Ʊ
    '���:byt���ϣ�1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-���￨
    '����:
    '����:HIS����ֽ�ʵķ�Ʊ����true,���򷵻�False
    '����:���˺�
    '����:2020-03-19 14:32:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlIsHisManagerInvoice = objEInvoice.zlIsHisManagerInvoice(byt����)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsStartEInvoiceFromBalanceID(ByVal byt���� As Byte, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ�����Ƿ������˵���Ʊ��
    '���:byt����-1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-���￨
    '     lng����ID-byt����=2ʱ,Ԥ��ID,���߽���ID
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-23 17:17:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlIsStartEInvoiceFromBalanceID = objEInvoice.zlIsStartEInvoiceFromBalanceID(byt����, lng����ID)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 

Public Function zlCreateEInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, _
    Optional ByRef blnTurnPaperInvoice_Out As Boolean, Optional ByVal blnOlnyCreateEInv As Boolean = False, Optional ByVal byt���� As Byte = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ߵ���Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    byt����-1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-���￨
    '    blnOlnyCreateEInv-�Ƿ�ֻ���ߵ���Ʊ��,true-��ʾֻ���ߵ���Ʊ�ݣ�False-���ݲ����������Ƿ񻻿�ֽ�ʷ�Ʊ
    '    cllSwapData-�������ݼ�
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    '����:
    '    blnTurnPaperInvoice_Out-ֽ�ʷ�Ʊ�����Ƿ�ɹ���true-��ʾ�ɹ���false-��ʾʧ��
    '����:���ߵ���Ʊ�ݳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '����:
    '   1.
    '������:
    '   1.�����շ�(�����˷�)�����ʡ��Һš� �����ʡ���Ԥ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlCreateEInvoice = objEInvoice.zlCreateEInvoice(frmMain, byt����, blnOlnyCreateEInv, cllSwapData, blnTurnPaperInvoice_Out)
    
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
Public Function zlTurnPaperInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, ByVal lngEInvoiceID As Long, _
    ByVal strEInvoiceCode As String, ByVal strEInvoiceNO As String, ByRef strInvoiceNO_Out As String, _
        Optional ByVal int����״̬ As Integer = 0, Optional ByVal strUseDate As String = "", Optional ByVal int���� As Byte = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:
    '    frmMain-����������
    '    int����-1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-ҽ�Ʒ���
    '    strInvoiceNO-��Ʊ��
    '    lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '    strEInvoiceCode-����Ʊ�ݴ���
    '    strEInvoiceNO-����Ʊ�ݺ���
    '    strUseDate-Ʊ��ʹ��ʱ��,������ʱ��Ҳ��ǰʱ��Ϊ׼
    '    int����״̬:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
    '    cllSwapData-�������ݼ�
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)

    '����:
    '   strInvoiceNO_Out:���ر��εķ�Ʊ����Ϣ,����ö��ŷ���;���ؿ�ʱ�����Դ���ķ�Ʊ��Ϊ׼
    '����:�����ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.��Ҫ����Ե���Ʊ�ݻ���ֽ��Ʊ��ҵ��
    '������:
    '   1.�����շ�(�����˷�)�����ʡ��Һš� �����ʡ���Ԥ��
    '   2.���������»���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlTurnPaperInvoice = objEInvoice.zlTurnPaperInvoice(frmMain, int����, cllSwapData, lngEInvoiceID, strEInvoiceCode, strEInvoiceNO, strInvoiceNO_Out, int����״̬, strUseDate)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCancelEInvoice(ByVal lngEInvoiceID As Long, _
    ByVal cllPati As Collection, Optional ByVal str����ʱ�� As String, Optional byt���� As Byte = 4, _
    Optional ByVal str���� As String = "�˺�") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ϵ���Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    int����-1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-ҽ�Ʒ���
    '    lngEInvoiceID-��Ҫ���ϵĵ���Ʊ�ݼ�¼ID: ����Ʊ��ʹ�ü�¼.ID
    '    strExtend-��չ�������������ִ�Ϊ��
    '    cllPati-������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    '    str����ʱ��-yyyy-mm-dd hh24:mi:ss
    '����:
    '    strEInvoiceClientCode-����Ʊ�ݿ�Ʊ��
    '    strSysSouceName_Out-ϵͳ��Դ����
    '    strEInvoiceCode_out-����Ʊ�ݴ���(��Ʊ)
    '    strEInvoiceNO_out-����Ʊ�ݺ���(��Ʊ)
    '    strCheckCode_out-����Ʊ��У����(��Ʊ)
    '    strCreateTime_Out-����Ʊ������ʱ��(��Ʊ),��ʽ:yyyy-mm-dd HH:MM:SS
    '    strEInvQRCode_Out-����Ʊ�ݶ�ά��ͼƬ����(��Ʊ),base64���ܴ洢,ʹ��ʱ�������
    '    strEInvUrl_Out-����Ʊ��URL��ַ(��Ʊ)
    '    strEinvRemark_Out-����Ʊ�ݵ�������Ϣ(��Ʊ)
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:���ϳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.���Ѿ������ĵ���Ʊ�ݽ�������
    '   2.���صĵ���Ʊ����Ϣ�洢�ڡ�����Ʊ��ʹ�ü�¼����
    '������:
    '   1.�����˷ѡ��������ϡ��˺š� ���������ϡ�Ԥ���˿��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlCancelEInvoice = objEInvoice.zlCancelEInvoice(mfrmMain, byt����, lngEInvoiceID, cllPati, str����ʱ��, str����)
    Set objEInvoice = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlPrintEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô�ӡƱ�ݽӿ�
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-25 10:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function

    zlPrintEInvoice = objEInvoice.zlPrintEInvoice(frmMain, lngEInvoiceID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetEInvoiceInforFromBalanceID(ByVal lng����ID As Long, ByRef rsEInvoce_Out As ADODB.Recordset, _
    Optional ByVal byt���� As Byte = 4, Optional ByVal int��ѯ��Χ As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID,��ȡָ������Ʊ����Ϣ
    '���:byt����-1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-ҽ�Ʒ���
    '     int��ѯ��Χ-0-����;1-ֻ��ѯ��Ч�ĵ���Ʊ��
    '
    '����:rsEInvoce_Out-����ָ���ĵ���Ʊ����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-25 11:28:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlGetEInvoiceInforFromBalanceID = objEInvoice.zlGetEInvoiceInforFromBalanceID(byt����, lng����ID, rsEInvoce_Out, int��ѯ��Χ)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlCancelPaperInvoice(ByVal frmMain As Object, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, _
    ByVal lngEInvoiceID As Long, Optional ByVal str����ʱ�� As String = "", Optional ByVal byt���� As Byte = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ��Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    byt����-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '    strInvoiceNO-��Ʊ��
    '    lngBalanceID-"1.byt����=2(Ԥ��)ʱ��ԭԤ��ID"
    '    2.byt����<>2(Ԥ��)ʱ��ԭ����ID"
    '    lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '    strExtend -��չ�������������ִ�Ϊ��
    '����:
    '����:����ֽ��Ʊ�ݳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.��Դ�ӡ����ӡ��ֽ��˺ֽ����ɵ�ֽ��Ʊ���𻵣���ƽ̨�ֻ�ֽ�ɹ��ģ���Ҫ���ϻ�����ֽ��Ʊ��
    '������:
    '   1.�����շѡ����ʡ��Һš� �����ʡ���Ԥ���ȵ����Ϲ���
    '   2.���»���������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    
    zlCancelPaperInvoice = objEInvoice.zlCancelPaperInvoice(frmMain, byt����, strInvoiceNO, lngBalanceID, _
      lngEInvoiceID, UserInfo.���, UserInfo.����, str����ʱ��)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetTranPaperInvoiceModule(Optional ByVal byt���� As Byte = 4) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ֽ�ʷ�Ʊ��ʽ
    '���:byt���ϣ�1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-���￨
    '����:
    '����:'0-��������1-�Զ�������2-��ʾ����
    '����:���˺�
    '����:2020-03-27 11:31:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlGetTranPaperInvoiceModule = objEInvoice.zlGetTranPaperInvoiceModule(byt����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPrintNotice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, Optional ByVal byt���� As Byte = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ��֪��
    '���:frmMain-����������
    '     lngEInvoiceID-����Ʊ��ID
    '     byt����:��1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-25 10:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    On Error GoTo errHandle
    
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlPrintNotice = objEInvoice.zlPrintNotice(frmMain, byt����, lngEInvoiceID)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlGetEInvoiceSwapCollectFromBalanceID(ByVal lng����ID As Long, ByRef cllSwapData_Out As Collection, ByRef blnStartEInvoice_Out As Boolean, _
    Optional ByVal strInvoiceNO As String, Optional ByVal lng����ID As Long, Optional ByVal blnԤ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�Ƿ���ȷ
    '���:lng����ID-����ID(blnԤ��=trueʱ��ΪԤ��ID)
    '     strInvoiceNO-��ǰ��Ʊ��
    '     lng����ID-��Ʊ������ID
    '     blnԤ��-��ǰ��Ԥ����
    '����:
    '   blnStartEInvoice_Out-�Ƿ����õ��ӷ�Ʊ
    '   cllSwapData_Out-���ؽ�����Ϣ
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-27 17:21:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, cllBalanceInfo As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    
    On Error GoTo errHandle
    If blnԤ�� Then
        strSQL = " " & _
        "   Select a.Id, a.No, a.����id, a.��ҳid, Sum(A.���) As ���ʽ��, Max(A.Ԥ������Ʊ��) As �Ƿ����Ʊ��, " & _
        "          Max(Nvl(d.����, c.����)) As ����, " & _
        "          Max(Nvl(d.�Ա�, c.�Ա�)) As �Ա�, Max(Nvl(d.����, c.����)) As ����, Max(Nvl(d.סԺ��, c.סԺ��)) As סԺ��, Max(c.�����) As �����, " & _
        "          max(M.����) as ����,to_char(max(A.�տ�ʱ��),'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ��,max(a.Ԥ�����) as ��������" & _
        "   From  ����Ԥ����¼ A, ������Ϣ C, ������ҳ D,(Select ��¼ID, ���� From ���ս����¼ where ����=3  and ��¼ID=[1] ) M" & _
        "   Where a.����id = c.����id(+) And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And a.Id=[1]  And A.ID=M.��¼ID(+)" & _
        "   Group By a.Id, a.No, a.����id, a.��ҳid"
    Else
        
        strSQL = " " & _
        "   Select a.Id, a.No, a.����id, a.��ҳid, Sum(b.��Ԥ��) As ���ʽ��, Max(b.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, " & _
        "          Max(decode(nvl(A.����ID,0),0,A.ԭ��,Nvl(d.����, c.����))) As ����, " & _
        "          Max(Nvl(d.�Ա�, c.�Ա�)) As �Ա�, Max(Nvl(d.����, c.����)) As ����, Max(Nvl(d.סԺ��, c.סԺ��)) As סԺ��, Max(c.�����) As �����, " & _
        "          max(M.����) as ����,to_char(max(A.�շ�ʱ��),'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ��,max(A.��������) as ��������" & _
        "   From ���˽��ʼ�¼ A, ����Ԥ����¼ B, ������Ϣ C, ������ҳ D,(Select ��¼ID, ���� From ���ս����¼ where ����=2  and ��¼ID=[1] ) M" & _
        "   Where a.id=b.����ID and  a.����id = c.����id(+) And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And a.Id=[1]  And A.ID=M.��¼ID(+)" & _
        "   Group By a.Id, a.No, a.����id, a.��ҳid"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���ID��������Ʊ����Ϣ", lng����ID)
    If rsTemp.EOF Then Exit Function
    Set cllPati = New Collection
    
    blnStartEInvoice_Out = Val(Nvl(rsTemp!�Ƿ����Ʊ��)) = 1
     
    '1.����������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    Set cllPati = New Collection
    cllPati.Add Val(Nvl(rsTemp!����ID)), "_����ID"
    cllPati.Add Val(Nvl(rsTemp!��ҳID)), "_��ҳID"
    
    cllPati.Add Nvl(rsTemp!����), "_����"
    cllPati.Add Nvl(rsTemp!�Ա�), "_�Ա�"
    cllPati.Add Nvl(rsTemp!����), "_����"
    cllPati.Add Nvl(rsTemp!�����), "_�����"
    cllPati.Add Nvl(rsTemp!סԺ��), "_סԺ��"
    cllPati.Add Val(Nvl(rsTemp!����)), "_����"

    '2.����������Ϣ:(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID,��Լ��λ����,��������)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_��Ʊ��"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add 0, "_����ID"
    cllBalanceInfo.Add Nvl(rsTemp!NO), "_���ݺ�"
    cllBalanceInfo.Add Format(rsTemp!�շ�ʱ��, "yyyy-mm-dd HH:MM:SS"), "_�Ǽ�ʱ��"
    cllBalanceInfo.Add 0, "_�Ƿ񲹽���"
    cllBalanceInfo.Add 0, "_�Ƿ񲿷��˿�"
    cllBalanceInfo.Add UserInfo.���, "_����Ա���"
    cllBalanceInfo.Add UserInfo.����, "_����Ա����"
    cllBalanceInfo.Add RoundEx(Val(Nvl(rsTemp!���ʽ��)), 2), "_���ʽ��"
    cllBalanceInfo.Add lng����ID, "_����ID"
    If blnԤ�� Then
        cllBalanceInfo.Add Decode(Val(Nvl(rsTemp!��������)) = 0, 3, Val(Nvl(rsTemp!��������))), "_��������" 'Ԥ�����:1-����;2-סԺ ;3-�����סԺ;
        cllBalanceInfo.Add IIf(Val(Nvl(rsTemp!����ID)) = 0, 1, 0), "_��Լ��λ����"
    Else
        cllBalanceInfo.Add Decode(Val(Nvl(rsTemp!��������)) = 0, 3, Val(Nvl(rsTemp!��������))), "_��������"  '��������:1-����;2-סԺ ;3-�����סԺ;
        cllBalanceInfo.Add IIf(Val(Nvl(rsTemp!����ID)) = 0, 1, 0), "_��Լ��λ����"
    End If
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    zlGetEInvoiceSwapCollectFromBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mfrmMain = Nothing
    Set mobjEInvoice = Nothing
End Sub

Public Function zlGetEinvoiceSwapCollectFromBalanceInfo(ByVal objPati As clsPatientInfo, _
    ByVal strInvoiceNO As String, ByRef cllSwapData_Out As Collection, _
    ByVal strNO As String, ByVal str�Ǽ�ʱ�� As String, ByVal dbl���ʽ�� As Double, _
    Optional ByVal lng����ID As Long, Optional ByVal lng����ID As Long, _
    Optional ByVal blnTranPaper As Boolean, Optional lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ������Ϣ��
    '���:objPati-������Ϣ����
    '     objBalanceInfor-��ǰ�������
    '     strInvoiceNo-��ǰ��Ʊ��
    '     lng����ID-����ID
    '     bln��Լ��λ����
    '����:cllSwapData_Out-������Ϣ��
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-25 12:02:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiInfo As Collection, cllBalanceInfo As Collection, cllSwapData As Collection
    Dim i As Long, bytFunc As Byte
    
    On Error GoTo errHandle
    '1.����������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    Call zlGetPatiCollectFromPatiObject(objPati, cllPatiInfo)
    
    If blnTranPaper And strInvoiceNO = "" Then    '����ֽ�ʷ�Ʊ
        '����Ʊ�����ö�ȡ
        Call mobjEInvoice.zlGetNextInvoiceNo(Me, strInvoiceNO, cllPatiInfo, lng����ID)
        If strInvoiceNO = "" Then
            If frmInputBox.InputBox(mfrmMain, "��ʼ��Ʊ��", "�޷���ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                            vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
        Else
            If frmInputBox.InputBox(mfrmMain, "��ʼ��Ʊ��", "��ȷ�ϻ���Ҫʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
        End If
        If strInvoiceNO = "" Then Exit Function
    End If
    
    '2.����������Ϣ:(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID,��Լ��λ����,��������)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_��Ʊ��"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add strNO, "_���ݺ�"
    cllBalanceInfo.Add Format(str�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS"), "_�Ǽ�ʱ��"
    cllBalanceInfo.Add 0, "_�Ƿ񲹽���"
    cllBalanceInfo.Add 0, "_�Ƿ񲿷��˿�"
    cllBalanceInfo.Add UserInfo.���, "_����Ա���"
    cllBalanceInfo.Add UserInfo.����, "_����Ա����"
    cllBalanceInfo.Add dbl���ʽ��, "_������"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add 1, "_��������" '��������:1-����;2-סԺ ;3-�����סԺ;
    cllBalanceInfo.Add 0, "_��Լ��λ����"
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPatiInfo, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    zlGetEinvoiceSwapCollectFromBalanceInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiCollectFromPatiObject(ByVal objPati As clsPatientInfo, ByRef cllPatiInfo_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�����Ϣ���󣬹���������Ϣ��
    '���:objPati-������Ϣ��
    '     int����-����
    '����:cllPatiInfo_Out-������Ϣ��(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-27 11:44:46
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    
    If objPati Is Nothing Then Exit Function
    
    '1.����������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    Set cllPatiInfo_Out = New Collection
    cllPatiInfo_Out.Add objPati.����ID, "_����ID"
    cllPatiInfo_Out.Add 0, "_��ҳID"
    
    cllPatiInfo_Out.Add objPati.����, "_����"
    cllPatiInfo_Out.Add objPati.�Ա�, "_�Ա�"
    cllPatiInfo_Out.Add objPati.����, "_����"
    cllPatiInfo_Out.Add objPati.�����, "_�����"
    cllPatiInfo_Out.Add "", "_סԺ��"
    cllPatiInfo_Out.Add objPati.����, "_����"
    zlGetPatiCollectFromPatiObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCancelEInvoiceFromBalanceInfor(ByVal objRegInfo As clsRegEventInfor, Optional ByVal byt���� As Byte = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ�����Ϣ������Ե���Ʊ�ݳ���
    '���:objPati-������Ϣ��
    '     objBalanceInfor-������Ϣ��
    '����:objBalanceInfor(�Ƿ����Ʊ������)��Ϊtrue
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-27 12:06:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsEInvoice As ADODB.Recordset
    Dim cllPati As Collection, strDate As String, lng����ID As Long
    Dim strSQL As String, strNO As String
    
    On Error GoTo errHandle
    strNO = IIf(byt���� = 4, objRegInfo.objPayInfos.Reg_���ݺ�, objRegInfo.objPayInfos.Card_���ݺ�)
    If byt���� = 4 Then
        strSQL = "   Select Max(a.�Ƿ����Ʊ��), Max(b.Id) As ����Ʊ��id " & _
        "   From ����Ԥ����¼ a, ����Ʊ��ʹ�ü�¼ b, ������ü�¼ c " & _
        "   Where c.No = [1] And c.��¼���� = 4 and c.����ID = a.����id and a.��¼״̬ In (1, 3) And a.����id = b.����id(+) And b.Ʊ��(+) = 4 And b.��¼״̬(+) = 1 "
        Set rsEInvoice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Һ���Ϣ", strNO)
        If rsEInvoice.EOF Then
            MsgBox "δ�ҵ��Һŵ���" & strNO & "����ԭʼ�Һ����ݣ����飡", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    Else
        strSQL = "   Select Max(a.�Ƿ����Ʊ��), Max(b.Id) As ����Ʊ��id " & _
        "   From ����Ԥ����¼ a, ����Ʊ��ʹ�ü�¼ b, סԺ���ü�¼ c " & _
        "   Where c.No = [1] And c.��¼���� = 5 and c.����ID = a.����id and a.��¼״̬ In (1, 3) And a.����id = b.����id(+) And b.Ʊ��(+) = 5 And b.��¼״̬(+) = 1 "
        Set rsEInvoice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Һ���Ϣ", strNO)
        If rsEInvoice.EOF Then
            MsgBox "δ�ҵ���������" & strNO & "����ԭʼ�������ݣ����飡", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If Val(Nvl(rsEInvoice!����Ʊ��Id)) = 0 Then
        zlCancelEInvoiceFromBalanceInfor = True: Exit Function
    End If
    If zlGetPatiCollectFromPatiObject(objRegInfo.objPati, cllPati) = False Then Exit Function
    
    '���ϵ��ӷ�Ʊ
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    zlCancelEInvoiceFromBalanceInfor = zlCancelEInvoice(Val(Nvl(rsEInvoice!����Ʊ��Id)), cllPati, strDate, IIf(byt���� = 4, 4, 5), IIf(byt���� = 4, "�˺�", "�˿�"))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function