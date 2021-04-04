VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEInvoice_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsEInvoiceModule
'*********************************************************************************************************************************************
'����Ʊ����ؽӿ�
'һ������Ʊ�ݹ����ӿ���:
'   1.zlInitialize :��ʼ������
'   2.zlTerminate:��ֹ�ӿ�
'   3.zlEInvoiceSet:���ӷ�Ʊ�豸��������ýӿ�
'   4.zlEInvoiceManager:����Ʊ�ݹ����ӿ�
'����Ʊ�ݿ�������ؽӿ�
'   1.zlGetStockInvoices:��ȡ��ǰ�����Ч�ķ�Ʊ�ӿ�
'   2.zlOutStockInvoices:�¿��Ʊ�ݽӿ�
'   3.zlReturnStockInvoices :Ʊ�����˽ӿ�
'   4.zlCancelBlankInvoice:����հ�Ʊ�ݽӿ�
'��������Ʊ����ؽӿ�
'   1.zlCreateEInvoice:���ߵ���Ʊ�ݽӿ�
'   2.zlCheckCancelEInvoiceValied:������ϵ���Ʊ�ݵĺϷ���
'   3.zlCancelEInvoice:���ϵ���Ʊ�ݽӿ�
'   4.zlPrintEInvoice:��ӡ����Ʊ��(A4ֽ)
'   5.zlSendEinvoiceMsg:���͵���Ʊ��֪ͨ��Ϣ
'�ġ�ֽ��Ʊ����ؽӿ�
'   1.zlGetNextInvoiceNo:��ȡ��һ����Ч��Ʊ��
'   2.zlTurnPaperInvoice:����ֽ�ʷ�Ʊ�ӿ�
'   3.zlPrintPaperInvoice:��ӡֽ�ʷ�Ʊ�ӿ�
'   4.zlCancelPaperInvoice:����ֽ��Ʊ�ݽӿ�
'����:���˺�
'����:2020-03-03 14:11:42
'*********************************************************************************************************************************************
Private mfrmMain As Object, mlngSys As Long, mlngModule As Long
Private mstrOperatorCode As String, mstrOperatorName As String, mstrDBUser As String
Private mstrEInvoiceNodeCode As String
Private mcnOracle As ADODB.Connection
Private Declare Function GetCurrentDirectory Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private mfrmEInvoiceManage As frmEInvoiceManage

Private Function clsEInvoiceModule_zlGetEInvoiceProductName(ByVal frmMain As Object, strProductName As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ʊ��ϵͳ����
    '���:
    '    frmMain-����������
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:
    '    strProductName-����Ʊ��ϵͳ����
    '    strExtend-��չ�������������ִ�Ϊ
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���ϴ�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.�ڿ��ߵ���Ʊ��ʱ����Ҫ��ȡ����Ʊ����Դ
    '������:
    '   1.�����շ�(�����˷�)�����ʡ��Һš� �����ʡ���Ԥ���ȿ��ߵ���Ʊ��ʱ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strProductName = "�������Կ�"
    gstrProductName = strProductName
    clsEInvoiceModule_zlGetEInvoiceProductName = True
End Function

Private Function clsEInvoiceModule_zlGetParaFrom() As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ý���
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '    1.��Ҫ���ڵ���̨�����е��õ���Ʊ�ݹ���,������ҽ�������еġ������˻�������ģ��
    '    2.�ýӿڹ��ܣ��Ե���Ʊ�ݵĶ��ˡ����ϡ���ѯ��
    '������:
    '   1.����ģ�飺"1144-����Ʊ�ݹ���"
    '   2.��ڳ����ɲ���zL9CashBill���ñ��ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set clsEInvoiceModule_zlGetParaFrom = Nothing
End Function

Private Function clsEInvoiceModule_zlInitialize(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModule As Long, ByVal strOperatorCode As String, ByVal strOperatorName As String, ByVal strDBUser As String, Optional ByVal blnDeviceSet As Boolean, Optional ByRef strEInvoiceNodeCode As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿڲ���
    '���:
    '    frmMain -���õ�������
    '    cnOracle -���ݿ����Ӷ���
    '    lngSys -��ǰ����ϵͳ���
    '    lngModule-��ǰ����ģ���:
    '    strOperatorCode -����Ա���
    '    strOperatorName -����Ա����
    '    strDBUser -���ݿ��û���
    '    blnDeviceSet -�豸���õ��õĳ�ʼ��
    '    strExtend-��չ����,����,�ִ�Ϊ��
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
    Set mcnOracle = cnOracle: Set mfrmMain = frmMain
    mlngSys = lngSys: mlngModule = lngModule: mstrDBUser = strDBUser
    mstrOperatorCode = strOperatorCode: mstrOperatorName = strOperatorName
    mstrEInvoiceNodeCode = strEInvoiceNodeCode
    clsEInvoiceModule_zlInitialize = True
End Function

Private Function clsEInvoiceModule_zlPrintEInvoice(ByVal frmMain As Object, ByVal lngID As Long, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ����Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    lngID-��Ҫ���ϵĵ���Ʊ�ݼ�¼ID: ����Ʊ��ʹ�ü�¼.ID
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:��ӡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.��Ҫ�Ǵ�ӡ����Ʊ��(��Ҫ��A4ֽƱ��)
    '������:
    '    1.�����շѣ��򲿷��˷ѣ���Ԥ�������ʡ������ʵȿ��ߵ���Ʊ�ݳɹ�����Ҫ���ñ��ӿ�֪ͨ����"
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If MsgBox("�Ƿ��ӡ����Ʊ�ݽӿڳɹ�?", vbQuestion + vbYesNo + vbDefaultButton1, "����Ʊ��") = vbNo Then
        strErrMsg_Out = "��ӡ����Ʊ�ݽӿ�ʱ��ѡ���ˡ���"
        Exit Function
    End If
    clsEInvoiceModule_zlPrintEInvoice = True
End Function

Private Function clsEInvoiceModule_zlTerminate() As Boolean
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
    clsEInvoiceModule_zlTerminate = True
End Function
Private Function clsEInvoiceModule_zlEInvoiceSet(ByVal frmMain As Object, Optional strExtend As String) As Boolean
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
    clsEInvoiceModule_zlEInvoiceSet = True
End Function
Private Function clsEInvoiceModule_zlGetStockInvoices(ByVal frmMain As Object, ByVal bytOccasion As Byte, ByVal bytInvoiceType As Byte, strBatchNo_out As String, strPrefix_out As String, strStartInvoice_out As String, strEndInvoice_out As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�����Ч�ķ�Ʊ�ӿ�
    '���:
    '    frmMain-����������
    '    bytOccasion-����:0-Ʊ����ⷢ��;1-Ʊ�����÷���
    '    bytInvoiceType-��Ʊ����:1-�շ�,2-Ԥ��,3-����,4-�Һ�
    '    strExtend-��չ����,����,�ִ�Ϊ��
    '����:
    '    strBatchNo_out-����
    '    strPrefix_out-ǰ׺�ı�
    '    strStartInvoice_out-��ʼ��Ʊ��
    '    strEndInvoice_out-��ֹ��Ʊ��
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend-��չ����,����,�ִ�Ϊ��
    '����:��ȡ��ǰ���ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.ֻ���ص�ǰ��Ч��Ʊ��Ϣ.
    '   2.�ӿڷ���trueʱ,���û�з��ؿ�ʼ��Ʊ�Ż���ֹ��Ʊ��,�������ֹ�¼�����ķ�Ʊ��Ϣ,�����ֹ��������
    '   3.���δ�д˽ӿ�,��ȱʡΪTrue;
    '   4.���������ƽ̨�ӿ��з��ض�����Ч��Ʊ��,���ڽӿ��е���ѡ����,ѡ��һ���Ϸ������Ʊ��
    '������:
    '   1.Ʊ��������:���
    '   2.Ʊ��ʹ�ü��:����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    clsEInvoiceModule_zlGetStockInvoices = True
End Function
Private Function clsEInvoiceModule_zlOutStockInvoices(ByVal frmMain As Object, ByVal lngID As Long, ByVal bytOccasion As Byte, ByVal bytInvoiceType As Byte, ByVal strUserName As String, ByVal strRemark As String, ByVal strBatchNo As String, ByVal strStartInvoice As String, ByVal strEndInvoice As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�¿��Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    lngID-����ID�����ID
    '    bytOccasion-����:0-Ʊ����ⷢ��;1-Ʊ�����÷���
    '    bytInvoiceType -��Ʊ����:1-�շ�,2-Ԥ��,3-����,4-�Һ�
    '    strUserName-ʹ����
    '    strRemark-��ע
    '    strBatchNo-����
    '    strStartInvoice-��ʼ��Ʊ��
    '    strEndInvoice-��ֹ��Ʊ��
    '    strExtend -��չ����,����,�ִ�Ϊ��
    '����:
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend -��չ����,����,�ִ�Ϊ��
    '����:�¿��ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.���ӿ���Ҫ���������ƽ̨Ʊ�ݿ����¿��,HIS��ʾ��������
    '������:
    '   1.Ʊ��������:���
    '   2.Ʊ��ʹ�ü��:����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_zlOutStockInvoices = True
End Function
Private Function clsEInvoiceModule_zlReturnStockInvoices(ByVal frmMain As Object, ByVal lngID As Long, ByVal bytOccasion As Byte, ByVal bytInvoiceType As Byte, ByVal strReturnerName As String, ByVal strRemark As String, ByVal strBatchNo As String, ByVal strStartInvoice As String, ByVal strEndInvoice As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ʊ�����˽ӿ�
    '���:
    '    frmMain-����������
    '    lngID-����ID�����ID
    '    bytOccasion-����:0-Ʊ����ⷢ��;1-Ʊ�����÷���
    '    bytInvoiceType-��Ʊ����:1-�շ�,2-Ԥ��,3-����,4-�Һ�
    '    strReturnerName-������
    '    strRemark-����ԭ��
    '    strBatchNo-����
    '    strStartInvoice-��ʼ��Ʊ��
    '    strEndInvoice-��ֹ��Ʊ��
    '    strExtend -��չ����,����,�ִ�Ϊ��
    '����:
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend -��չ����,����,�ִ�Ϊ��
    '����:Ʊ�����˳ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.Ʊ��������Ҫ�ṩ������������Ʊ���졢������,���˵�����ƽ̨�Ŀ��
    '������:
    '   1.Ʊ��������:����
    '   2.Ʊ��ʹ�ü��:����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_zlReturnStockInvoices = True
End Function

Private Function clsEInvoiceModule_zlCancelBlankInvoice(ByVal frmMain As Object, ByVal lngID As Long, ByVal bytInvoiceType As Byte, ByVal strAuthorName As String, ByVal strReason As String, ByVal strBatchNo As String, ByVal strStartInvoice As String, ByVal strEndInvoice As String, ByVal strHappenTime As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����հ�Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    lngID-����ID
    '    bytInvoiceType-��Ʊ����:1-�շ�,2-Ԥ��,3-����,4-�Һ�
    '    strAuthorName-������
    '    strReason-����ԭ��
    '    strBatchNo-����
    '    strStartInvoice-��ʼ��Ʊ��
    '    strEndInvoice-��ֹ��Ʊ��
    '    strHappenTime -����ʱ��:yyyy-mm-dd HH:MM:SS
    '    strExtend -��չ����,����,�ִ�Ϊ��
    '����:
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend -��չ����,����,�ִ�Ϊ��
    '����:����հ�Ʊ�ݳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.��Ҫ������շ�Ա��ͷ�Ŀհ�Ʊ��δʹ�õ�������ԭ����ɷ�Ʊ������Ĵ�������
    '������:
    '   1.Ʊ��������:����
    '   2.Ʊ��ʹ�ü��:����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_zlCancelBlankInvoice = True
End Function
Private Function clsEInvoiceModule_zlCreateEInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, _
    ByVal lngEInvoiceID As Long, ByVal lngBalanceID As Long, ByVal lngWriteOffID As Long, _
    ByRef strEInvoiceClientCode As String, ByRef strSysSouceName_Out As String, _
    ByRef strEInvoiceCode_out As String, ByRef strEInvoiceNo_Out As String, ByRef strCheckCode_out As String, ByRef strCreateTime_Out As String, _
    ByRef strEInvQRCode_Out As String, ByRef strEInvUrl_Out As String, ByRef strEInvUrl1_Out As String, ByRef strEinvRemark_Out As String, ByRef strErrMsg_Out As String, _
    Optional ByRef strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ߵ���Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    bytInvoiceType-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '    lngBalanceID-bytInvoiceType=2(Ԥ��)ʱ��ԭԤ��ID ;bytInvoiceType<>2(Ԥ��)ʱ��ԭ����ID
    '    lngWriteOffID-bytInvoiceType<>2(Ԥ��)ʱ����ǰ����ID,�����˷�ʱ��Ч
    '    strEInvoiceClientCode-����ͻ��˿�Ʊ����
    '    strExtend-��չ����,����,�ִ�Ϊ��
    '����:
    '    strEInvoiceClientCode-���صĿͻ��˿�Ʊ����
    '    strSysSouceName_Out-ϵͳ��Դ����
    '    strEInvoiceCode_out-����Ʊ�ݴ���
    '    strEInvoiceNO_out-����Ʊ�ݺ���
    '    strCheckCode_out-����Ʊ��У����
    '    strCreateTime_Out-����Ʊ������ʱ��,��ʽ:yyyy-mm-dd HH:MM:SS
    '    strEInvQRCode_Out-����Ʊ�ݶ�ά��ͼƬ����,base64���ܴ洢,ʹ��ʱ,�����
    '    strEInvUrl_Out-����Ʊ��URL��ַ
    '    strEInvUrl1_Out-����Ʊ��URL������ַ
    '    strEinvRemark_Out-����Ʊ�ݵ�������Ϣ
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend-��չ����,JSON��ʽ
    '    strExtend-��չ����,��Ϊ��,����Ԥ��ƾ֤/Json��ʽ
    '        voucherbatchcode    String  50  Ԥ����ƾ֤����
    '        voucherno   String  20  Ԥ����ƾ֤����
    '        voucherrandom   String  20  Ԥ����ƾ֤У����
    '����:���ߵ���Ʊ�ݳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.��ҽ�Ƶ���Ʊ�ݹ���ƽ̨�������Ʊ�ݿ�������,���ɵ���Ʊ��,�����ص���Ʊ�������Ϣ
    '   2.���صĵ���Ʊ����Ϣ�洢�ڡ�����Ʊ��ʹ�ü�¼����
    '������:
    '   1.�����շ�(�����˷�)�����ʡ��Һš� �����ʡ���Ԥ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAppPath As String
    Dim strƾ֤���� As String, strƾ֤���� As String, strƾ֤������ As String
    
    On Error GoTo errHandle
    
    strAppPath = App.Path & "\�ĵ�����\QRCode.jpg"
    zlStr.PrefixNO
    strSysSouceName_Out = "�������Կ�"
    strEInvoiceCode_out = "C" & Format(Now, "yyyymmddHHMMSS")
    strEInvoiceNo_Out = "E" & Format(Now, "yyyymmddHHMMSS") & zlStr.LPAD(Int(Rnd * 1000), 4, "0")
    strCheckCode_out = "X" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    strCreateTime_Out = Format(zlDatabase.Currentdate, "yyyyMMddHHmmss000")
    'base64ת��
    strEInvQRCode_Out = zlStr.EncodeBase64_File(strAppPath)
    
    '�����ά��
'    strAppPath = App.Path & "\�ĵ�����\QRCode1.jpg"
'    strAppPath = zlStr.DecodeBase64_File(strEInvQRCode_Out, strAppPath)
    
    strEInvUrl_Out = "http://192.168.0.1/h"
    strEInvUrl1_Out = "http://192.168.0.1/h1"
    strEinvRemark_Out = "����"
    strErrMsg_Out = ""
    strƾ֤���� = "A" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    strƾ֤���� = "B" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    strƾ֤������ = "C" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    strExtend = "{""voucherbatchcode"":""" & strƾ֤���� & """,""voucherNo"":""" & strƾ֤���� & """,""voucherrandom"":""" & strƾ֤������ & """}"
    If MsgBox("�Ƿ�Ʊ�ɹ�?", vbQuestion + vbYesNo + vbDefaultButton1, "����Ʊ��") = vbNo Then
        strErrMsg_Out = "��Ʊ����ʱ��ѡ���ˡ���"
        Exit Function
    End If
    
    clsEInvoiceModule_zlCreateEInvoice = True
     
    Exit Function
errHandle:
     strErrMsg_Out = Err.Description
End Function
Private Function clsEInvoiceModule_zlCheckCancelEInvoiceValied(ByVal frmMain As Object, ByVal lngID As Long, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϵ���Ʊ�ݵĺϷ���
    '���:
    '    frmMain-����������
    '    lngID-��Ҫ���ϵĵ���Ʊ�ݼ�¼ID:����Ʊ��ʹ�ü�¼.ID
    '    strExtend-��չ����,����,�ִ�Ϊ��
    '����:
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend-��չ����,����,�ִ�Ϊ��
    '����:���ϺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.����Ʊ������ǰ��һЩ�Ϸ��Լ��
    '������:
    '   1.�����˷ѡ��������ϡ��˺š� ���������ϡ�Ԥ���˿��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_zlCheckCancelEInvoiceValied = True
End Function



Private Function clsEInvoiceModule_zlCancelEInvoice(ByVal frmMain As Object, ByVal lngID As Long, ByRef strEInvoiceClientCode As String, ByRef strSysSouceName_Out As String, _
    ByRef strEInvoiceCode_out As String, ByRef strEInvoiceNo_Out As String, _
    ByRef strCheckCode_out As String, ByRef strCreateTime_Out As String, ByRef strEInvQRCode_Out As String, ByRef strEInvUrl_Out As String, ByRef strEInvUrl1_Out As String, _
    ByRef strEinvRemark_Out As String, ByRef strErrMsg_Out As String, Optional ByRef strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ϵ���Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    lngID-��Ҫ���ϵĵ���Ʊ�ݼ�¼ID: ����Ʊ��ʹ�ü�¼.ID
    '    strEInvoiceClientCode-����Ʊ�ݿ�Ʊ��
    '    strEInvoiceNode_Out-��Ʊ��
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:
    '    strEInvoiceClientCode-����Ʊ�ݿ�Ʊ��
    '    strSysSouceName_Out-ϵͳ��Դ����
    '    strEInvoiceCode_out-����Ʊ�ݴ���(��Ʊ)
    '    strEInvoiceNO_out-����Ʊ�ݺ���(��Ʊ)
    '    strCheckCode_out-����Ʊ��У����(��Ʊ)
    '    strCreateTime_Out-����Ʊ������ʱ��(��Ʊ),��ʽ:yyyy-mm-dd HH:MM:SS
    '    strEInvQRCode_Out-����Ʊ�ݶ�ά��ͼƬ����(��Ʊ),base64���ܴ洢,ʹ��ʱ�������
    '    strEInvUrl_Out-����Ʊ��URL��ַ(��Ʊ)
    '    strEInvUrl1_Out-����Ʊ������URL��ַ(��Ʊ)
    '    strEinvRemark_Out-����Ʊ�ݵ�������Ϣ(��Ʊ)
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend-��չ������JSON��ʽ����
    '        voucherbatchcode    String  50  ��ƱԤ����ƾ֤����
    '        voucherno   String  20  ��ƱԤ����ƾ֤����
    '        voucherrandom   String  20  ��ƱԤ����ƾ֤У����
    '����:���ϳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.���Ѿ������ĵ���Ʊ�ݽ�������
    '   2.���صĵ���Ʊ����Ϣ�洢�ڡ�����Ʊ��ʹ�ü�¼����
    '������:
    '   1.�����˷ѡ��������ϡ��˺š� ���������ϡ�Ԥ���˿��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAppPath As String
    On Error GoTo errHandle
    
    strAppPath = App.Path & "\�ĵ�����\QRCode.jpg"
 
    strEInvoiceCode_out = "CT" & Format(Now, "yyyymmddHHMMSS")
    strEInvoiceNo_Out = "ET" & Format(Now, "yyyymmddHHMMSS") & zlStr.LPAD(Int(Rnd * 1000), 4, "0")
    strCheckCode_out = "XT" & zlStr.LPAD(Int(Rnd * 1000000000), 6, "0")
    strCreateTime_Out = Format(zlDatabase.Currentdate, "yyyyMMddHHmmss000")
    
    'base64ת��
    strEInvQRCode_Out = zlStr.EncodeBase64_File(strAppPath)
    strSysSouceName_Out = "�������Կ�"
    '�����ά��
    'strAppPath = App.Path & "\�ĵ�����\QRCode1.jpg"
    'strAppPath = zlStr.DecodeBase64_File(strEInvQRCode_Out, strAppPath)

    
    strEInvUrl_Out = "http://192.168.0.1/h"
    strEInvUrl1_Out = "http://192.168.0.1/h1"
    strEinvRemark_Out = "����"
    strErrMsg_Out = ""
    strExtend = ""
    If MsgBox("�Ƿ����ϵ���Ʊ�ݳɹ�?", vbQuestion + vbYesNo + vbDefaultButton1, "����Ʊ��") = vbNo Then
        strErrMsg_Out = "���ϵ���Ʊ������ʱ��ѡ���ˡ���"
        Exit Function
    End If
    clsEInvoiceModule_zlCancelEInvoice = True
    Exit Function
errHandle:
     strErrMsg_Out = Err.Description
End Function
Private Function clsEInvoiceModule_zlSendEinvoiceMsg(ByVal frmMain As Object, ByVal lngID As Long, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���͵���Ʊ��֪ͨ��Ϣ
    '���:
    '    frmMain-����������
    '    lngID-����Ʊ�ݼ�¼ID
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:���ͳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.��Ҫ��֪ͨ���ߵ���Ʊ�ݵ������Ϣ:���� ��ӡ��֪������΢�š������ʼ���ƽ̨֪ͨ���߲鿴����Ʊ��
    '������:
    '   1.�����շѣ��򲿷��˷ѣ���Ԥ�������ʡ������ʵȿ��ߵ���Ʊ�ݳɹ�����Ҫ���ñ��ӿ�֪ͨ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If MsgBox("�Ƿ��͵���Ʊ��֪ͨ�ɹ�?", vbQuestion + vbYesNo + vbDefaultButton1, "����Ʊ��") = vbNo Then
        strErrMsg_Out = "���͵���Ʊ��֪ͨʱ��ѡ���ˡ���"
        Exit Function
    End If
    clsEInvoiceModule_zlSendEinvoiceMsg = True
End Function
Private Function clsEInvoiceModule_zlGetNextInvoiceNo(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, strInvoiceNO_Out As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��һ����Ч��Ʊ��
    '���:
    '    frmMain-����������
    '    bytInvoiceType-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:
    '    strInvoiceNO_Out-��һ����Ч��Ʊ��
    '    strErrMsg_Out-��ȡʧ��ʱ�����ص���ʾ��Ϣ
    '    strExtend-��չ�������������ִ�Ϊ
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.�ڻ���ֽ�ʷ�Ʊʱ����Ҫ�ȶ�ȡ��һ����Ч�ķ�Ʊ��,�Ա���ʾ����������
    '������:
    '   1.�����շ�(�����˷�)�����ʡ��Һš� �����ʡ���Ԥ���Ƚ������ʱ����Ҫ��ȡ��ʾ�ڽ�����
    '   2.���������»���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strInvoiceNO_Out = "F" & Format(Now, "yyyymmddHHMMSS") & zlStr.LPAD(Int(Rnd * 10000), 5, "0")
    If MsgBox("��ȡ��һ����Ч��Ʊ���Ƿ�ɹ�?" & vbCrLf & "���صķ�Ʊ��Ϊ" & strInvoiceNO_Out, vbQuestion + vbYesNo + vbDefaultButton1, "����Ʊ��") = vbNo Then
        strErrMsg_Out = "��ȡ��һ����Ч��Ʊ��ʱ��ѡ���ˡ���"
        strInvoiceNO_Out = ""
        Exit Function
    End If
    clsEInvoiceModule_zlGetNextInvoiceNo = True
End Function
Private Function clsEInvoiceModule_zlTurnPaperInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, ByVal lngEInvoiceID As Long, ByVal strEInvoiceCode As String, ByVal strEInvoiceNO As String, _
    ByVal strCreateTime As String, ByVal strOperatorCode As String, ByVal strOperatorName As String, ByRef strInvoiceNO_Out As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ�ʷ�Ʊ�ӿ�
    '���:
    '    frmMain-����������
    '    bytInvoiceType-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '    strInvoiceNO-��Ʊ��
    '    lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '    strEInvoiceCode-����Ʊ�ݴ���
    '    strEInvoiceNO-����Ʊ�ݺ���
    '    strCreateTime-����Ʊ������ʱ��,��ʽ:yyyy-mm-dd HH:MM:SS
    '    strOperatorCode-����Ա���
    '    strOperatorName-����Ա����
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:
    '    strInvoiceNO_Out:���ر��εķ�Ʊ����Ϣ,����ö��ŷ���;���ؿ�ʱ�����Դ���ķ�Ʊ��Ϊ׼
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend-��չ�������������ִ�Ϊ��
    '����:�����ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.��Ҫ����Ե���Ʊ�ݻ���ֽ��Ʊ��ҵ��
    '������:
    '   1.�����շ�(�����˷�)�����ʡ��Һš� �����ʡ���Ԥ��
    '   2.���������»���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If MsgBox("�Ƿ񻻿�ֽ�ʷ�Ʊ�ӿڳɹ�?" & vbCrLf & "��ǰ��Ʊ�ţ�" & strInvoiceNO, vbQuestion + vbYesNo + vbDefaultButton1, "����Ʊ��") = vbNo Then
        strErrMsg_Out = "��ӡ����Ʊ�ݽӿ�ʱ��ѡ���ˡ���"
        Exit Function
    End If
    strInvoiceNO_Out = strInvoiceNO
    clsEInvoiceModule_zlTurnPaperInvoice = True
    
End Function
Private Function clsEInvoiceModule_zlPrintPaperInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, ByVal lngEInvoiceID As Long, ByVal strOperatorCode As String, ByVal strOperatorName As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡֽ�ʷ�Ʊ�ӿ�
    '���:
    '    frmMain-����������
    '    bytInvoiceType-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '    strInvoiceNO-��Ʊ��
    '    lngBalanceID-bytInvoiceType=2(Ԥ��)ʱ��ԭԤ��ID,bytInvoiceType<>2(Ԥ��)ʱ��ԭ����ID
    '    lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '    strOperatorCode-����Ա���
    '    strOperatorName-����Ա����
    '����:
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend -��չ�������������ִ�Ϊ��
    '����:��ӡ����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.   ��ӡ�����֣�һ����HIS�Լ���ӡ;һ����ͨ�������ṩ��Ʊ�ݴ�ӡ   �������ַ�ʽ�������ڴ˽ӿ��б��ƴ�ӡ
    '������:
    '   1.�����շ�(�����˷�)�����ʡ��Һš� �����ʡ���Ԥ��
    '   2.���������»���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If MsgBox("��ӡֽ�ʷ�Ʊ�ӿڳɹ�?" & vbCrLf & "��ǰ��Ʊ�ţ�" & strInvoiceNO, vbQuestion + vbYesNo + vbDefaultButton1, "����Ʊ��") = vbNo Then
        strErrMsg_Out = "��ӡֽ�ʷ�Ʊ�ӿ�ʱ��ѡ���ˡ���"
        Exit Function
    End If
    clsEInvoiceModule_zlPrintPaperInvoice = True
End Function

Private Function clsEInvoiceModule_zlCancelPaperInvoice(ByVal frmMain As Object, ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, ByVal lngEInvoiceID As Long, ByVal strOperatorCode As String, ByVal strOperatorName As String, strErrMsg_Out As String, Optional strExtend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ��Ʊ�ݽӿ�
    '���:
    '    frmMain-����������
    '    bytInvoiceType-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    '    strInvoiceNO-��Ʊ��
    '    lngBalanceID-"1.bytInvoiceType=2(Ԥ��)ʱ��ԭԤ��ID"
    '    2.bytInvoiceType<>2(Ԥ��)ʱ��ԭ����ID"
    '    lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '    strOperatorCode-����Ա���
    '    strOperatorName-����Ա����
    '    strExtend -��չ�������������ִ�Ϊ��
    '����:
    '    strErrMsg_Out-ʧ��ԭ��
    '    strExtend -��չ������JSON��ʽ����ʽ����
    '            einvoiceinfo    String      ����Ʊ����Ϣ��������ֽ��Ʊ��ͬʱ�����˵���Ʊ��ʱͨ��strExpend����
    '                einvoicecode    String  20  ���Ӻ�Ʊ����
    '                einvoiceno      String  20  ���Ӻ�Ʊ����
    '                einvoicerandom  String  20  ���Ӻ�ƱУ����
    '                createtime      String  17  ���Ӻ�Ʊ����ʱ��
    '                billqrcode      String  ����    ���Ӻ�Ʊ��ά��ͼƬ����
    '                pictureurl      String  ����    ����Ʊ��H5ҳ��URL
    '                pictureneturl   String  ����    ����Ʊ������H5ҳ��URL��ַ
    '                syssoucename    String  ����    ��Ʊϵͳ����
    '                voucherbatchcode    String  50  ��ƱԤ����ƾ֤����,����Ԥ��Ʊ��ʱ���нڵ�
    '                voucherno       String  20  ��ƱԤ����ƾ֤����,ͬ��
    '                voucherrandom   String  10  ��ƱԤ����ƾ֤У����,ͬ��
    '����:����ֽ��Ʊ�ݳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '   1.��Դ�ӡ����ӡ��ֽ��˺ֽ����ɵ�ֽ��Ʊ���𻵣���ƽ̨�ֻ�ֽ�ɹ��ģ���Ҫ���ϻ�����ֽ��Ʊ��
    '������:
    '   1.�����շѡ����ʡ��Һš� �����ʡ���Ԥ���ȵ����Ϲ���
    '   2.���»���������
    '---------------------------------------------------------------------------------------------------------------------------------------------

    If MsgBox("����ֽ��Ʊ�ݽӿڳɹ�?" & vbCrLf & "��ǰ��Ʊ�ţ�" & strInvoiceNO, vbQuestion + vbYesNo + vbDefaultButton1, "����Ʊ��") = vbNo Then
        strErrMsg_Out = "����ֽ��Ʊ�ݽӿ�ʱ��ѡ���ˡ���"
        Exit Function
    End If
    
    clsEInvoiceModule_zlCancelPaperInvoice = True
End Function

Private Sub clsEInvoiceModule_zlEInvoiceManager(ByVal frmMain As Object, Optional strExtend As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݹ����ӿ�
    '���:
    '    frmMain-����������
    '    strExtend -��չ�������������ִ�Ϊ��
    '����:
    '    strExtend -��չ�������������ִ�Ϊ��
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '    1.��Ҫ���ڵ���̨�����е��õ���Ʊ�ݹ���,������ҽ�������еġ������˻�������ģ��
    '    2.�ýӿڹ��ܣ��Ե���Ʊ�ݵĶ��ˡ����ϡ���ѯ��
    '������:
    '   1.����ģ�飺"1144-����Ʊ�ݹ���"
    '   2.��ڳ����ɲ���zL9CashBill���ñ��ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmEInvoiceManage Is Nothing Then Set mfrmEInvoiceManage = New frmEInvoiceManage
    Call mfrmEInvoiceManage.ShowMe(frmMain, mlngSys, mlngModule, mstrDBUser, Me)
End Sub

Private Sub clsEInvoiceModule_zlBHEInvoiceManager(ByVal lngMain As Long, Optional ByRef strExtend As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݹ����ӿ�
    '���:
    '    lngMain:��������
    '    strExtend -��չ�������������ִ�Ϊ��
    '����:
    '    strExtend -��չ�������������ִ�Ϊ��
    '����:���˺�
    '����:2020-03-03 11:17:34
    '˵��:
    '    1.��Ҫ���ڵ���̨�����е��õ���Ʊ�ݹ���,������ҽ�������еġ������˻�������ģ��
    '    2.�ýӿڹ��ܣ��Ե���Ʊ�ݵĶ��ˡ����ϡ���ѯ��
    '������:
    '   1.����ģ�飺"1144-����Ʊ�ݹ���"
    '   2.��ڳ����ɲ���zL9CashBill���ñ��ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmEInvoiceManage Is Nothing Then Set mfrmEInvoiceManage = New frmEInvoiceManage
    Call mfrmEInvoiceManage.BHShowMe(lngMain, mlngSys, mlngModule, mstrDBUser, Me)
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mcnOracle = Nothing
    Set mfrmMain = Nothing
    If Not mfrmEInvoiceManage Is Nothing Then Unload mfrmEInvoiceManage: Set mfrmEInvoiceManage = Nothing
End Sub

Public Function clsEInvoiceModule_ZlGetTotalCheckData(ByVal dtBegin As Date, ByVal dtEnd As Date, ByRef cllDatas_Out As Collection, _
    Optional ByVal bytMode As Byte = 1, Optional ByVal str��Ʊ�� As String, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƽ̨����Ʊ�ݻ�������
    '���:
    '   dtBegin��dtEnd ҵ��ʼ���ڡ���������
    '   bytMode �������ͣ�1-��Ʊ����Ʊ��2-����Ʊ
    '   str��Ʊ�� ƽ̨��Ʊ�����
    '����:
    '   cllDatas �������ݣ�����Ԫ�أ�(ҵ������,��Ʊ��,��Ʊ���,�ܱ���,���ؽ��,����ԭ��),Key=_ҵ������
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
End Function

Public Function clsEInvoiceModule_ZlGetDetailCheckData(ByVal dtBegin As Date, ByVal dtEnd As Date, ByRef cllDatas_Out As Collection, _
    Optional ByVal bytMode As Byte = 1, Optional ByVal str��Ʊ�� As String, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƽ̨����Ʊ����ϸ����
    '���:
    '   dtBegin��dtEnd ҵ��ʼʱ�䡢����ʱ��
    '   bytMode �������ͣ�1-��Ʊ����Ʊ��2-����Ʊ
    '   str��Ʊ�� ƽ̨��Ʊ�����
    '����:
    '   cllDatas �������ݣ�����Ԫ��(ҵ������,ҵ������,ҵ����ˮ��,��Ʊ��,Ʊ����������,Ʊ�ݴ���,Ʊ�ݺ���,��Ʊ���,��Ʊʱ��,��������,����Ʊ�ݴ���,����Ʊ�ݺ���),Key=_ҵ����ˮ��
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
End Function

Public Property Get clsEInvoiceModule_ZLCheckTimeMode() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ����Ʊ�ݺ˶Ե�ʱ������
    '˵����0-Ʊ�ݿ���ʱ�䣬1-����ҵ����ʱ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    clsEInvoiceModule_ZLCheckTimeMode = 0
End Property
