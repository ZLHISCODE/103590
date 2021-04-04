VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEInvoiceObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'   4.zlIsStartEinvoicFromNO-���ݷ��þݺţ��ж��Ƿ�ǰ���ʵ��Ƿ������˵���Ʊ�ݵ�
'����ҵ������ؽӿ�
'   1.zlCreateEInvoice-���ߵ���Ʊ��
'   2.zlCancelEInvoice-���ϵ���Ʊ��
'     zlCancelEInvoiceFromBalanceInfor-���ݽ��ʶ���ִ�е���Ʊ��Ʊ�����ϲ���
'   3.zlPrintEInvoice-��ӡ����Ʊ��A4�ӿ�
'   4.zlTurnPaperInvoice-����ֽ��Ʊ��
'   5.zlCancelPaperInvoice-����ֽ��Ʊ��
'   6.zlGetNextInvoiceNo-��ȡ��һ����Ч�ķ�Ʊ��
'   7.zlGetTranPaperInvoiceModule-��ȡֽ�ʷ�Ʊ�Ļ�����ʽ
'�ġ���ȡҵ������
'   1.zlGetEInvoiceInforFromBalanceID-���ݽ���ID,��ȡ����Ʊ����Ϣ(���ؼ�¼��)
'����:���˺�
'����:2020-03-19 10:43:04
'*********************************************************************************************************************************************
Private mlngMoudle As Long, mstrPrivs As String
Private mfrmMain As Object
Private mobjEInvoice As Object
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
        mobjEInvoice.IsExseRegion = False
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
    Optional cllPati As Collection, Optional lng����ID As Long, Optional ByVal byt���� As Byte = 5) As Boolean
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
Public Function zlInitialize(ByVal frmMain As Object, Optional ByVal byt���� As Byte = 5, Optional ByVal blnDeviceSet As Boolean, Optional ByVal lngModule As Long = 1107) As Boolean
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


Public Function zlIsStartEInvoice(Optional ByVal byt���� As Byte = 5, Optional ByVal intInsure As Integer = 0, Optional int��� As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ������˵���Ʊ�ݵ�
    '���:byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�
    '     intInsure-��ǰҽ��:0-������ͨ����;>0��ʾ����ҽ��
    '     int���-0-���������;1-����;2-סԺ
    '����:
    '����:���÷���true,���򷵻�False
    '����:���˺�
    '����:2020-03-19 14:32:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlIsStartEInvoice = objEInvoice.zlIsStartEInvoice(byt����, intInsure, int���)
    
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
    '    strExtend-1-��ǰ��֧�ֿ��ߵ���Ʊ�ݲ�֧�ֻ���
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

Public Function zlIsHisManagerInvoice(Optional ByVal byt���� As Byte = 5) As Boolean
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

Public Function zlIsStartEInvoiceFromBalanceID(ByVal lng����ID As Long, Optional ByVal byt���� As Byte = 5) As Boolean
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
    Optional ByRef blnTurnPaperInvoice_Out As Boolean, Optional ByVal blnOlnyCreateEInv As Boolean = False, Optional ByVal byt���� As Byte = 5) As Boolean
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
        Optional ByVal int����״̬ As Integer = 0, Optional ByVal strUseDate As String = "", Optional ByVal int���� As Byte = 5) As Boolean
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

Public Function zlCancelEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, _
    ByVal cllPati As Collection, Optional ByVal str����ʱ�� As String, Optional byt���� As Byte = 5, _
    Optional ByVal str���� As String = "��������") As Boolean
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
    zlCancelEInvoice = objEInvoice.zlCancelEInvoice(frmMain, byt����, lngEInvoiceID, cllPati, str����ʱ��, str����)
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
    Optional ByVal byt���� As Byte = 5, Optional ByVal int��ѯ��Χ As Integer = 1) As Boolean
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


Public Function zlCancelPaperInvoice(ByVal frmMain As Object, ByVal strInvoiceNo As String, ByVal lngBalanceID As Long, ByVal cllPati As Collection, ByVal strNO As String, _
    ByVal lngEInvoiceID As Long, Optional ByVal str����ʱ�� As String = "", Optional ByVal byt���� As Byte = 5) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ��Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    byt����-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '    strInvoiceNO-��Ʊ��
    '    lngBalanceID-"1.byt����=2(Ԥ��)ʱ��ԭԤ��ID,2.byt����<>2(Ԥ��)ʱ��ԭ����ID"
    '    lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '    strExtend -��չ�������������ִ�Ϊ��
    '    cllPati-������Ϣ��, �Ƿ���ش��˲���,��ʽ������ID,����,�Ա�,����,�����,סԺ�ţ�,Key("_����")
    '    strNO-���ݺ�
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
    Dim bln������ As Boolean, bln�����ϵ���Ʊ�� As Boolean, blnShowMsg As Boolean, strErrmsg_Out As String, intԤ����Ʊ As Integer
    On Error GoTo errHandle
    
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    
    blnShowMsg = True
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ��Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    byt����-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '    strInvoiceNO-��Ʊ��
    '    lngBalanceID-"1.byt����=2(Ԥ��)ʱ��ԭԤ��ID"
    '    2.byt����<>2(Ԥ��)ʱ��ԭ����ID"
    '    lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '    strOperatorCode-����Ա���
    '    strOperatorName-����Ա����
    '    strExtend -��չ�������������ִ�Ϊ��
    '    bln������-�Ƿ���ò�����
    '    cllPati-������Ϣ��, �Ƿ���ش��˲���,��ʽ������ID,����,�Ա�,����,�����,סԺ�ţ�,Key("_����")
    '    strNO-���ݺ�,�Ƿ���ش��˲���,�ֱ�Ϊ�շѡ�Ԥ�������ʻ���￨���õ��ݺ�
    '����: bln���ϵ���Ʊ��-����Ʊ���Ƿ�ͬʱ����
    zlCancelPaperInvoice = objEInvoice.zlCancelPaperInvoice(frmMain, byt����, strInvoiceNo, lngBalanceID, _
      lngEInvoiceID, UserInfo.���, UserInfo.����, str����ʱ��, bln������, bln�����ϵ���Ʊ��, blnShowMsg, strErrmsg_Out, intԤ����Ʊ, cllPati, strNO)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetTranPaperInvoiceModule(Optional ByVal byt���� As Byte = 5) As Integer
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

'Public Function zlGetEInvoiceSwapCollectFromBalanceID(ByVal lng����ID As Long, ByRef cllSwapData_Out As Collection, ByRef blnStartEInvoice_Out As Boolean, _
'    Optional ByVal strInvoiceNo As String, Optional ByVal lng����ID As Long, Optional ByVal blnԤ�� As Boolean) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ȡ�����������Ϣ
'    '���:lng����ID-����ID(blnԤ��=trueʱ��ΪԤ��ID)
'    '     strInvoiceNO-��ǰ��Ʊ��
'    '     lng����ID-��Ʊ������ID
'    '     blnԤ��-��ǰ��Ԥ����
'    '     strNo-���źŲ�Ϊ��ʱ�������ݺŴ���
'    '����:
'    '   blnStartEInvoice_Out-�Ƿ����õ��ӷ�Ʊ
'    '   cllSwapData_Out-���ؽ�����Ϣ
'    '      |-PatiInfo   Key="_PatiInfo"
'    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
'    '      |-BalanceInfo Key="_BalanceInfo"
'    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
'    '����:�ɹ�����true,���򷵻�False
'    '����:���˺�
'    '����:2020-03-27 17:21:08
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cllPati As Collection, cllBalanceInfo As Collection
'    Dim strSQL As String, rsTemp As ADODB.Recordset
'
'
'    On Error GoTo errHandle
'
'
'    If blnԤ�� Then
'        strSQL = " " & _
'        "   Select a.Id, a.No, a.����id, a.��ҳid, Sum(A.���) As ���ʽ��, Max(A.Ԥ������Ʊ��) As �Ƿ����Ʊ��, " & _
'        "          Max(Nvl(d.����, c.����)) As ����, " & _
'        "          Max(Nvl(d.�Ա�, c.�Ա�)) As �Ա�, Max(Nvl(d.����, c.����)) As ����, Max(Nvl(d.סԺ��, c.סԺ��)) As סԺ��, Max(c.�����) As �����, " & _
'        "          max(M.����) as ����,to_char(max(A.�տ�ʱ��),'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ��,max(a.Ԥ�����) as ��������" & _
'        "   From  ����Ԥ����¼ A, ������Ϣ C, ������ҳ D,(Select ��¼ID, ���� From ���ս����¼ where ����=3  and ��¼ID=[1] ) M" & _
'        "   Where a.����id = c.����id(+) And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And a.Id=[1]  And A.ID=M.��¼ID(+)" & _
'        "   Group By a.Id, a.No, a.����id, a.��ҳid"
'
'    Else
'    strSQL = " " & _
'        "   Select a.����id As ID, b.No, a.����id, a.��ҳid, Sum(a.��Ԥ��) As ���ʽ��, Max(a.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, " & _
'        "          Max(c.����) As ����, Max(c.סԺ��) As סԺ��, Max(c.�����) As �����, 0 As ����, " & _
'        "          To_Char(Max(a.�տ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �շ�ʱ��, 1 as �������� " & _
'        "   From ����Ԥ����¼ A, (Select  ����id, From סԺ���ü�¼ Where ����id = [1]) B, ������Ϣ C  " & _
'        "   Where a.����id = b.����id And a.����id = c.����id(+)  And a.Id = [1] " & _
'        "   Group By a.����id, b.No, a.����id, a.��ҳid"
'    End If
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���ID��������Ʊ����Ϣ", lng����ID)
'    If rsTemp.EOF Then Exit Function
'    Set cllPati = New Collection
'
'    blnStartEInvoice_Out = Val(nvl(rsTemp!�Ƿ����Ʊ��)) = 1
'
'    '1.����������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
'    Set cllPati = New Collection
'    cllPati.Add Val(nvl(rsTemp!����ID)), "_����ID"
'    cllPati.Add Val(nvl(rsTemp!��ҳID)), "_��ҳID"
'
'    cllPati.Add nvl(rsTemp!����), "_����"
'    cllPati.Add nvl(rsTemp!�Ա�), "_�Ա�"
'    cllPati.Add nvl(rsTemp!����), "_����"
'    cllPati.Add nvl(rsTemp!�����), "_�����"
'    cllPati.Add nvl(rsTemp!סԺ��), "_סԺ��"
'    cllPati.Add Val(nvl(rsTemp!����)), "_����"
'
'    '2.����������Ϣ:(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID,��Լ��λ����,��������)
'    Set cllBalanceInfo = New Collection
'    cllBalanceInfo.Add strInvoiceNo, "_��Ʊ��"
'    cllBalanceInfo.Add lng����ID, "_����ID"
'    cllBalanceInfo.Add 0, "_����ID"
'    cllBalanceInfo.Add nvl(rsTemp!NO), "_���ݺ�"
'    cllBalanceInfo.Add Format(rsTemp!�շ�ʱ��, "yyyy-mm-dd HH:MM:SS"), "_�Ǽ�ʱ��"
'    cllBalanceInfo.Add 0, "_�Ƿ񲹽���"
'    cllBalanceInfo.Add 0, "_�Ƿ񲿷��˿�"
'    cllBalanceInfo.Add UserInfo.���, "_����Ա���"
'    cllBalanceInfo.Add UserInfo.����, "_����Ա����"
'    cllBalanceInfo.Add RoundEx(Val(nvl(rsTemp!���ʽ��)), 5), "_������"
'    cllBalanceInfo.Add lng����ID, "_����ID"
'    cllBalanceInfo.Add IIf(Val(nvl(rsTemp!����ID)) = 0, 1, 0), "_��Լ��λ����"
'
'    If blnԤ�� Then
'        cllBalanceInfo.Add decode(Val(nvl(rsTemp!��������)) = 0, 3, Val(nvl(rsTemp!��������))), "_��������" 'Ԥ�����:1-����;2-סԺ ;3-�����סԺ;
'    Else
'        cllBalanceInfo.Add 1, "_��������"  '��������:1-����;2-סԺ ;3-�����סԺ;
'    End If
'    Set cllSwapData_Out = New Collection
'    cllSwapData_Out.Add cllPati, "_PatiInfo"
'    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
'    zlGetEInvoiceSwapCollectFromBalanceID = True
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function
  
Public Function zlGetEInvoiceSwapCollectFromNo(ByVal strNos As String, ByRef cllSwapData_Out As Collection, ByRef blnStartEInvoice_Out As Boolean, _
    Optional ByVal strInvoiceNo As String, Optional ByVal lng����ID As Long, Optional ByVal byt���� As Byte = 5, _
    Optional ByRef dblEInvoice_Out As Double, Optional ByRef lngԭ����ID_Out As Long, Optional str�Ǽ�ʱ�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺţ���ȡ��Ҫ�������Ϣ
    '���:lng����ID-����ID(blnԤ��=trueʱ��ΪԤ��ID)
    '     strInvoiceNO-��ǰ��Ʊ��
    '     lng����ID-��Ʊ������ID
    '     strNo-���źŲ�Ϊ��ʱ�������ݺŴ���
    '     byt����-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '����:
    '   dblEInvoice_Out-����Ʊ�ݿ��Կ��ߵĽ��
    '   lngԭ����ID_Out-ԭ����ID
    '   blnStartEInvoice_Out-�Ƿ����õ��ӷ�Ʊ
    '   cllSwapData_Out-���ؽ�����Ϣ
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID,�Ƿ����Ʊ��)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-27 17:21:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Zl_Exsesvr_Getbalanceinfo(strNos, cllSwapData_Out, blnStartEInvoice_Out, strInvoiceNo, lng����ID, byt����, dblEInvoice_Out, lngԭ����ID_Out, str�Ǽ�ʱ��) = False Then Exit Function
    zlGetEInvoiceSwapCollectFromNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsStartEinvoicFromNO(ByVal strNO As String, Optional ByRef lngԭ����ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��õ����ж��Ƿ������˵���Ʊ�ݵ�
    '���:strNo-���ʵ���
    '���Σ�lngԭ����ID_Out-���صĽ���ID
    '����:���÷���true,���򷵻�False
    '����:���˺�
    '����:2020-03-31 17:48:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnStartEinvice As Boolean, dblMoney As Double
    On Error GoTo errHandle
    
    If Not Zl_Exsesvr_GetbalanceinfoFromNos(strNO, 5, dblMoney, lngԭ����ID_Out, blnStartEinvice) Then Exit Function
 
    zlIsStartEinvoicFromNO = blnStartEinvice
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetBalanceInforFromNO(ByVal strNO As String, ByRef dblTotalMoney_Out As Double, ByRef lngԭ����ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��õ���ԭ����ID
    '���:strNo-���õ���
    '���Σ�lngԭ����ID_Out-���صĽ���ID
    '����:���÷���true,���򷵻�False
    '����:���˺�
    '����:2020-03-31 17:48:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnStartEinvice As Boolean
    On Error GoTo errHandle
    
    If Not Zl_Exsesvr_GetbalanceinfoFromNos(strNO, 5, dblTotalMoney_Out, lngԭ����ID_Out, blnStartEinvice) Then Exit Function
    zlGetBalanceInforFromNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCancelEInvoiceFromNO(ByVal frmMain As Object, ByVal strNO As String, Optional bln��ʾ����Ʊ�� As Boolean, Optional strInvoiceNO_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ�ƿ����ŷ���Ե���Ʊ�ݳ���
    '���:strNO-���ݺ�
    '     bln��ʾ����Ʊ��-�Ƿ���ʾ��Ҫ���յ�Ʊ����Ϣ
    '����:
    '     strInvoiceNo_Out-��ʾ���յķ�Ʊ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-27 12:06:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsEInvoice As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim cllPati As Collection, strDate As String, lng����ID As Long
    Dim strSQL As String, lngEInvoiceID As Long
    
    On Error GoTo errHandle
         
    If Zl_Exsesvr_Geteinvoicesinfo(Me, strNO, cllPati, lngEInvoiceID, strInvoiceNO_Out) = False Then Exit Function
    
    
    If lngEInvoiceID = 0 Then
        zlCancelEInvoiceFromNO = True: Exit Function
    End If
    If bln��ʾ����Ʊ�� And strInvoiceNO_Out <> "" Then
        MsgBox "��ע��������·�Ʊ�ţ�" & vbCrLf & strInvoiceNO_Out, vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    End If
    
    
    '�ȼ�������Ƿ��fc
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If zlCancelEInvoice(frmMain, lngEInvoiceID, cllPati, strDate) = False Then Exit Function
    zlCancelEInvoiceFromNO = True
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

Public Function zlPrintNotice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, Optional ByVal byt���� As Byte = 5) As Boolean
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

Public Function zlExcuteErrDataEInvoice(ByVal frmMain As Object, ByVal lng����ID As Long, Optional ByVal lng����ID As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���쳣�ĵ���Ʊ��
    '���:  lng����ID-����ID
    '       lng����ID-������ID��
    '����:
    '����:0-��ʾ���쳣����
    '     1-��ʾ�����쳣���ݣ��������ɹ�!
    '     2-��ʾ�����쳣���ݣ�������ʧ��!
    '����:���˺�
    '����:2020-03-27 17:21:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object, rsTemp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errHandle
    
    strWhere = ""
    If lng����ID <> 0 Then strWhere = strWhere & " And  ����ID=[1] "
    If lng����ID <> 0 Then strWhere = strWhere & " And ҵ��id=[2] "
    
    If strWhere = "" Then
        'ֻ�������˵�
        If lng����ID <> 0 Then strWhere = strWhere & " And ����Ա����=[3] "
    End If
     ' --��¼��־    number(2)       0-���ߵ���Ʊ��;1-������Ʊ��;2-ֽ��Ʊ��;3-����ֽ��Ʊ��
    
    strSQL = " " & _
    "  Select ID,  ҵ������, ��¼��־" & _
    "  From ����Ʊ���쳣��¼ " & _
    "  where ��������=1 " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���ID��������Ʊ����Ϣ", lng����ID, lng����ID, UserInfo.����)
    
    '0-��ʾ���쳣����
    If rsTemp.EOF Then zlExcuteErrDataEInvoice = 0: Exit Function
    If GetEInvoiceObject(objEInvoice) = False Then zlExcuteErrDataEInvoice = 2: Exit Function
    With rsTemp
        Do While Not .EOF
            Select Case Val(nvl(rsTemp!��¼��־))
            Case 0  '���ߵ���Ʊ��ʧ��
                If MsgBox("���ϴο��ߵ���Ʊ��ʧ�ܣ�����Ҫ���ȴ����쳣�ĵ���Ʊ����Ϣ���ܼ�������,���Ƿ���Ҫ����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then zlExcuteErrDataEInvoice = 2: Exit Function
            Case 1  '������Ʊ��;
                If MsgBox("���ϴγ�����Ʊ��ʧ�ܣ�����Ҫ���ȴ����쳣�ĳ�����Ʊ����Ϣ���ܼ�������,���Ƿ���Ҫ����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then zlExcuteErrDataEInvoice = 2: Exit Function
            Case 2 'ֽ��Ʊ��;
                If MsgBox("���ϴλ���Ʊ��ʧ�ܣ�����Ҫ���ȴ����쳣��Ʊ����Ϣ���ܼ�������,���Ƿ���Ҫ����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then zlExcuteErrDataEInvoice = 2: Exit Function
            Case 3 '����ֽ��Ʊ��
                If MsgBox("���ϴ�����ֽ��Ʊ��ʧ�ܣ�����Ҫ���ȴ����쳣��Ʊ����Ϣ���ܼ�������,���Ƿ���Ҫ����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then zlExcuteErrDataEInvoice = 2: Exit Function
            Case Else
                MsgBox "����ʶ����쳣���ݣ�����!", vbOKOnly + vbInformation, gstrSysName
                zlExcuteErrDataEInvoice = 2: Exit Function
            End Select
            '�����쳣
            If objEInvoice.zlExcuteErrDataFromErrID(Val(nvl(rsTemp!ID))) = False Then zlExcuteErrDataEInvoice = 2: Exit Function
            rsTemp.MoveNext
        Loop
    End With
    zlExcuteErrDataEInvoice = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
  

Public Function zlExcuteErrDataEInvoiceFromErrID(ByVal frmMain As Object, ByVal lng�쳣ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����쳣ID�������쳣����
    '���:  lng�쳣ID-�쳣ID
    '����:
    '����: �ɹ�����true, ���򷵻�False
    '����:���˺�
    '����:2020-03-27 17:21:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objEInvoice As Object
    
    On Error GoTo errHandle
    If GetEInvoiceObject(objEInvoice) = False Then Exit Function
    zlExcuteErrDataEInvoiceFromErrID = objEInvoice.zlExcuteErrDataFromErrID(lng�쳣ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
  
