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
Private mlngModule As Long, mfrmMain As Object
Private mobjEInvoice As zlPublicExpense.clsPubEInvoice

Public Function Initialize(ByVal frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal byt���� As Byte = 1, _
    Optional ByVal blnDeviceSet As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿڲ���
    '���:
    '   frmMain�����õ�������
    '   lngModule����ǰ����ģ���
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '   blnDeviceSet���豸���õ��õĳ�ʼ��
    '����:
    '����:��ʼ���ɹ�����true,���򷵻�False
    '˵��:
    '   1.ʹ�ñ�����ǰ,�����ȵ��ñ��ӿڽ��г�ʼ��
    '   2.��ʼ���ӿ�,��HIS����ģ��ʱ����(���磺�����շѹ�������)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExtend As String
    
    Set mfrmMain = frmMain: mlngModule = lngModule
    
    If mobjEInvoice Is Nothing Then
        On Error Resume Next
        Set mobjEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
        If Err <> 0 Then
            MsgBox "�����ڿ��õĵ���Ʊ�ݽӿڲ���(zlPublicExpense.clsPubEInvoice)������ϵͳ����Ա��ϵ����ϸ�Ĵ�����ϢΪ:" & vbCrLf & Err.Description, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mobjEInvoice Is Nothing Then Exit Function
    
    Initialize = mobjEInvoice.zlInitialize(frmMain, byt����, gcnOracle, glngSys, lngModule, blnDeviceSet, strExtend)
End Function

Public Function ReReadParaAndVar() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¶�ȡ����������ֵ
    '����:�ɹ�����true,���򷵻�False
    '˵����
    '   1.�ڲ������ú����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ReReadParaAndVar = mobjEInvoice.zlReReadParaAndVar
End Function

Public Function GetNextInvoiceNo(ByVal frmMain As Object, ByRef strInvoiceNO_Out As String, _
    Optional ByVal cllPati As Collection, Optional ByRef lng����ID As Long, Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��һ�ŷ�Ʊ��
    '���:
    '   frmMain -���õ�������
    '   cllPati-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����):Ŀǰ���ã���Ҫ���Ժ���չHIS����Ʊ��
    '   lng����ID-Ʊ������ID,�Ժ�HIS����Ʊ�ݿ���Ҫ��,�����ô�
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:
    '   strInvoiceNO_Out-������һ����Ч��Ʊ��
    '����:��ȡ�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetNextInvoiceNo = mobjEInvoice.zlGetNextInvoiceNo(frmMain, byt����, strInvoiceNO_Out, cllPati, lng����ID)
End Function

Public Function IsStartEInvoice(Optional ByVal intInsure As Integer, Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ������˵���Ʊ�ݵ�
    '���:
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '   intInsure����ǰҽ��:0-������ͨ����;>0��ʾ����ҽ��
    '����:
    '����:���÷���true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    IsStartEInvoice = mobjEInvoice.zlIsStartEInvoice(byt����, intInsure)
End Function

Public Function EInvoiceSet(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ӷ�Ʊ�豸��������ýӿ�
    '���:
    '   frmMain�����õ�������
    '����:
    '����:���óɹ�����true,���򷵻�False
    '˵��:
    '    һ����Ե���Ʊ���豸����Ҫ�Ǵ�ӡ�豸������Ʊ������������õ�)һЩ���ý�������,���磺��ӡ��,�˿ںš�IP���û����������ַ�ȵ�����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExtend As String
    
    EInvoiceSet = mobjEInvoice.zlEInvoiceSet(frmMain, strExtend)
End Function

Public Function IsHisManagerInvoice(Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ�HIS����ֽ�ʷ�Ʊ
    '���:
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:
    '����:HIS����ֽ�ʵķ�Ʊ����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    IsHisManagerInvoice = mobjEInvoice.zlIsHisManagerInvoice(byt����)
End Function

Public Function IsStartEInvoiceFromBalanceID(ByVal lng����ID As Long, Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ�����Ƿ������˵���Ʊ��
    '���:
    '   lng����ID-����ID
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    IsStartEInvoiceFromBalanceID = mobjEInvoice.zlIsStartEInvoiceFromBalanceID(byt����, lng����ID)
End Function

Public Function CreateEInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, _
    Optional ByRef blnTurnPaperInvoice_Out As Boolean, _
    Optional ByVal blnOlnyCreateEInv As Boolean, Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ߵ���Ʊ�ݽӿ�
    '���:
    '   frmMain-����������
    '    cllSwapData-�������ݼ�
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    '   blnOlnyCreateEInv-�Ƿ�ֻ���ߵ���Ʊ��,true-��ʾֻ���ߵ���Ʊ�ݣ�False-���ݲ����������Ƿ񻻿�ֽ�ʷ�Ʊ
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:
    '    blnTurnPaperInvoice_Out-ֽ�ʷ�Ʊ�����Ƿ�ɹ���true-��ʾ�ɹ���false-��ʾʧ��
    '����:���ߵ���Ʊ�ݳɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CreateEInvoice = mobjEInvoice.zlCreateEInvoice(frmMain, byt����, blnOlnyCreateEInv, cllSwapData, blnTurnPaperInvoice_Out)
End Function

Public Function GetNextPaperInvoice(ByVal frmMain As Object, ByVal cllPatiInfo As Collection, _
    ByRef lng����ID As Long, Optional ByVal byt���� As Byte = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ���ŷ�Ʊ��
    '���:
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:��Ʊ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceNO As String
    
    On Error GoTo errHandle
    '����Ʊ�����ö�ȡ
    If GetNextInvoiceNo(frmMain, strInvoiceNO, cllPatiInfo, lng����ID, byt����) = False Then Exit Function
    If strInvoiceNO = "" Then
        If frmInputBox.InputBox(frmMain, "��Ʊ��ȷ��", "�޷���ȡ��Ҫʹ�õķ�Ʊ�ţ�" & _
                        vbCrLf & "�������뻻����Ҫʹ�õķ�Ʊ���룺", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
    Else
        If frmInputBox.InputBox(frmMain, "��Ʊ��ȷ��", "��ȷ�ϻ�����Ҫʹ�õķ�Ʊ�ţ�", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
    End If
    GetNextPaperInvoice = strInvoiceNO
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function TurnPaperInvoice(ByVal frmMain As Object, ByVal cllSwapData As Collection, ByVal lngEInvoiceID As Long, _
    ByVal strEInvoiceCode As String, ByVal strEInvoiceNo As String, ByRef strInvoiceNO_Out As String, _
    Optional ByVal int����״̬ As Integer, Optional ByVal strUseDate As String, Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:
    '   frmMain-����������
    '   cllSwapData-�������ݼ�
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    '   lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '   strEInvoiceCode-����Ʊ�ݴ���
    '   strEInvoiceNO-����Ʊ�ݺ���
    '   int����״̬:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
    '   strUseDate-Ʊ��ʹ��ʱ��,������ʱ��Ҳ��ǰʱ��Ϊ׼
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:
    '   strInvoiceNO_Out:���ر��εķ�Ʊ����Ϣ,����ö��ŷ���;���ؿ�ʱ�����Դ���ķ�Ʊ��Ϊ׼
    '����:�����ɹ�����true,���򷵻�False
    '˵��:
    '   1.��Ҫ����Ե���Ʊ�ݻ���ֽ��Ʊ��ҵ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceNO As String, lng����ID As Long
    Dim cllPatiInfo As Collection, cllBalance As Collection
    
    If strInvoiceNO_Out <> "" Then
        If MsgBox("���ν�����������ѻ���ֽ��Ʊ�ݣ��Ƿ������ЩƱ�ݣ�" & _
            vbCrLf & strInvoiceNO_Out, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    Set cllPatiInfo = cllSwapData("_PatiInfo")
    Set cllBalance = cllSwapData("_BalanceInfo")
    
    lng����ID = cllBalance("_����ID")
    strInvoiceNO = GetNextPaperInvoice(frmMain, cllPatiInfo, lng����ID, byt����)
    If strInvoiceNO = "" Then Exit Function '�û�ȡ�����룬������
    
    cllBalance.Remove "_��Ʊ��"
    cllBalance.Add strInvoiceNO, "_��Ʊ��"
    cllBalance.Remove "_����ID"
    cllBalance.Add lng����ID, "_����ID"
    
    TurnPaperInvoice = mobjEInvoice.zlTurnPaperInvoice(frmMain, byt����, cllSwapData, _
        lngEInvoiceID, strEInvoiceCode, strEInvoiceNo, strInvoiceNO_Out, int����״̬, strUseDate)
End Function

Private Function CancelEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, _
    ByVal cllPatiInfo As Collection, Optional ByVal str����ʱ�� As String, _
    Optional ByVal str���� As String = "�����˷�", Optional ByVal byt���� As Byte = 1, _
    Optional ByVal bln������ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ϵ���Ʊ�ݽӿ�
    '���:
    '   frmMain-����������
    '   lngEInvoiceID-��Ҫ���ϵĵ���Ʊ�ݼ�¼ID: ����Ʊ��ʹ�ü�¼.ID
    '   cllPatiInfo - ������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����)
    '   str����ʱ��-yyyy-mm-dd hh24:mi:ss
    '   str�������������ܣ���ʾʹ��
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:
    '����:���ϳɹ�����true,���򷵻�False
    '˵��:
    '   1.���Ѿ������ĵ���Ʊ�ݽ�������
    '   2.���صĵ���Ʊ����Ϣ�洢�ڡ�����Ʊ��ʹ�ü�¼����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CancelEInvoice = mobjEInvoice.zlCancelEInvoice(frmMain, byt����, lngEInvoiceID, cllPatiInfo, str����ʱ��, str����, bln������)
End Function

Public Function PrintEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô�ӡƱ�ݽӿ�
    '���:
    '   lngEInvoiceID-��Ҫ���ϵĵ���Ʊ�ݼ�¼ID: ����Ʊ��ʹ�ü�¼.ID
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PrintEInvoice = mobjEInvoice.zlPrintEInvoice(frmMain, lngEInvoiceID)
End Function

Public Function GetEInvoiceInforFromBalanceID(ByVal lng����ID As Long, ByRef rsEInvoce_Out As ADODB.Recordset, _
    Optional ByVal int��ѯ��Χ As Integer = 1, Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID,��ȡָ������Ʊ����Ϣ
    '���:
    '   lng����ID-�շ�ʱΪԭʼ����ID��������ʱΪԭʼ����ID
    '   int��ѯ��Χ-0-����;1-ֻ��ѯ��Ч�ĵ���Ʊ��
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:rsEInvoce_Out-����ָ���ĵ���Ʊ����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetEInvoiceInforFromBalanceID = mobjEInvoice.zlGetEInvoiceInforFromBalanceID(byt����, lng����ID, rsEInvoce_Out, int��ѯ��Χ)
End Function

Public Function CancelPaperInvoice(ByVal frmMain As Object, ByVal strInvoiceNO As String, ByVal lngBalanceID As Long, _
    ByVal lngEInvoiceID As Long, Optional ByVal str����ʱ�� As String, Optional ByVal byt���� As Byte = 1, _
    Optional ByVal bln������ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ��Ʊ�ݽӿ�
    '���:
    '   frmMain-����������
    '   strInvoiceNO-��Ʊ��
    '   lngBalanceID-byt����=2(Ԥ��)ʱ��ԭԤ��ID��byt����<>2(Ԥ��)ʱ��ԭ����ID
    '   lngEInvoiceID-����Ʊ��ʹ�ü�¼ID
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:
    '����:����ֽ��Ʊ�ݳɹ�����true,���򷵻�False
    '˵��:
    '   1.��Դ�ӡ����ӡ��ֽ��˺ֽ����ɵ�ֽ��Ʊ���𻵣���ƽ̨�ֻ�ֽ�ɹ��ģ���Ҫ���ϻ�����ֽ��Ʊ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If MsgBox("���ν�����������ѻ���ֽ��Ʊ�ݣ��Ƿ������ЩƱ�ݣ�" & _
        vbCrLf & strInvoiceNO, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    CancelPaperInvoice = mobjEInvoice.zlCancelPaperInvoice(frmMain, byt����, strInvoiceNO, lngBalanceID, _
        lngEInvoiceID, UserInfo.���, UserInfo.����, str����ʱ��, bln������)
End Function

Public Function GetTranPaperInvoiceModule(Optional ByVal byt���� As Byte = 1) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ֽ�ʷ�Ʊ��ʽ
    '���:
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:
    '����:'0-��������1-�Զ�������2-��ʾ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
     GetTranPaperInvoiceModule = mobjEInvoice.zlGetTranPaperInvoiceModule(byt����)
End Function

Public Function GetSwapCollectFromBalanceID(ByVal lngԭ����ID As Long, _
    ByRef cllSwapData_Out As Collection, _
    Optional ByVal strInvoiceNO As String, Optional ByVal lng����ID As Long, _
    Optional ByVal lng����ID As Long, Optional ByVal bln�����˷� As Boolean, _
    Optional ByVal strUseDate As String, Optional ByVal bln������ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ���׽�����Ϣ
    '���:
    '   lng����ID-�շ�ʱΪԭʼ����ID��������ʱΪԭʼ����ID
    '   strInvoiceNO-��ǰ��Ʊ��
    '   lng����ID-��Ʊ������ID
    '����:
    '   cllSwapData_Out-���ؽ�����Ϣ
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����),key(_�ڵ�����)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, cllBalanceInfo As Collection
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, strInsureSql As String
    
    On Error GoTo ErrHandler
    If bln������ Then
        strWhere = " And b.����id In(Select �շѽ���ID From ���ò����¼ Where ����ID=[1])"
    Else
        strWhere = " And b.����id = [1]"
    End If

    strSql = _
        " Select Max(a.����id) As ����ID, Max(a.��ҳid) As ��ҳID, Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����," & _
        "        f_List2Str(Cast(Collect(a.No) As t_StrList)) As NO, Sum(a.���ʽ��) As ���ʽ��, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��" & _
        " From (Select a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.No, a.���, Sum(a.���ʽ��) As ���ʽ��, Max(b.�Ǽ�ʱ��) As �Ǽ�ʱ��" & _
        "        From ������ü�¼ A, ������ü�¼ B" & _
        "        Where Mod(a.��¼����, 10) = Mod(b.��¼����, 10) And a.No = b.No And a.��� = b.���" & strWhere & _
        "        Group By a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.No, a.���" & _
        "        Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0) A"
    
    strInsureSql = "Select Max(����) As ���� From ���ս����¼ Where ���� = 1 And ��¼id = [1]"
    
    strSql = _
        " Select a.����id, a.��ҳid, a.����, a.�Ա�, a.����, m.�����, Nvl(n.סԺ��, m.סԺ��) As סԺ��," & _
        "           a.No, a.���ʽ��, a.�Ǽ�ʱ��, b.����" & _
        " From (" & strSql & ") A, (" & strInsureSql & ") B, ������Ϣ M, ������ҳ N" & _
        " Where a.����id = m.����id(+) And a.����id = n.����id(+) And a.��ҳid = n.��ҳid(+) And a.No Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ݽ���ID��������Ʊ����Ϣ", lngԭ����ID)
    If rsTemp.EOF Then Exit Function

    '1.����������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����)
    Set cllPati = New Collection
    cllPati.Add Val(Nvl(rsTemp!����ID)), "_����ID"
    cllPati.Add Val(Nvl(rsTemp!��ҳID)), "_��ҳID"
    cllPati.Add Nvl(rsTemp!����), "_����"
    cllPati.Add Nvl(rsTemp!�Ա�), "_�Ա�"
    cllPati.Add Nvl(rsTemp!����), "_����"
    cllPati.Add Nvl(rsTemp!�����), "_�����"
    cllPati.Add Nvl(rsTemp!סԺ��), "_סԺ��"
    cllPati.Add Val(Nvl(rsTemp!����)), "_����"

    '2.����������Ϣ:(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_��Ʊ��"
    cllBalanceInfo.Add lngԭ����ID, "_����ID"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add Nvl(rsTemp!NO), "_���ݺ�"
    If strUseDate = "" Then
        cllBalanceInfo.Add Format(Nvl(rsTemp!�Ǽ�ʱ��), "yyyy-mm-dd HH:MM:SS"), "_�Ǽ�ʱ��"
    Else
        cllBalanceInfo.Add Format(strUseDate, "yyyy-mm-dd HH:MM:SS"), "_�Ǽ�ʱ��"
    End If
    cllBalanceInfo.Add IIf(bln������, 1, 0), "_�Ƿ񲹽���"
    cllBalanceInfo.Add IIf(bln�����˷�, 1, 0), "_�Ƿ񲿷��˿�"
    cllBalanceInfo.Add UserInfo.���, "_����Ա���"
    cllBalanceInfo.Add UserInfo.����, "_����Ա����"
    cllBalanceInfo.Add Val(Nvl(rsTemp!���ʽ��)), "_������"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add 1, "_��������"
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    
    GetSwapCollectFromBalanceID = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CancelEInvoiceFromBalance(ByVal frmMain As Object, _
    ByVal strNos As String, Optional ByVal intInsure As Integer, Optional ByVal bln������ As Boolean, _
    Optional ByVal str���� As String = "�����˷�", Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ�����Ϣ������Ե���Ʊ�ݳ���
    '���:
    '   cllPatiInfo��������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,����)
    '   strNos��������ʱΪ���㵥�ţ���ͨ�շ�ʱΪһ�ν����е���һ�ŵ��ݺ�;���ʱ���ŷָ�
    '   str�������������ܣ���ʾʹ��
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:objBalanceInfor(�Ƿ����Ʊ������)��Ϊtrue
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsEInvoice As ADODB.Recordset, strSql As String, strԭ����ID As String
    Dim cllPati As Collection, strDate As String, strInvoice As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strԭ����ID = GetBalanceIDByFeeNo(strNos, bln������, byt����)
    
    strSql = _
        "  Select a.ID, a.����id, a.����, a.�Ա�, a.����, a.�����, a.סԺ��, a.�Ƿ񻻿�, a.ֽ�ʷ�Ʊ��" & _
        "  From ����Ʊ��ʹ�ü�¼ A,Table(F_Num2List([1])) J" & _
        "  Where a.����ID=j.Column_Value And a.Ʊ��=[2] And a.��¼״̬=1" & _
        "   Order By a.�Ǽ�ʱ��"
    Set rsEInvoice = zlDatabase.OpenSQLRecord(strSql, "��ȡ����Ʊ��ʹ�����", strԭ����ID, byt����)
    If rsEInvoice.EOF Then CancelEInvoiceFromBalance = True: Exit Function
    
'    If IsStartEInvoice(intInsure, byt����) = False Then
'        MsgBox "���ν����ѿ����˵���Ʊ�ݣ�����ǰ���ڲ�֧�ֵ���Ʊ�ݡ�" & vbCrLf & _
'            "�뵽֧�ֵ���Ʊ�ݵĴ��ڽ���" & str���� & "��", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    Do While Not rsEInvoice.EOF
        If cllPati Is Nothing Then
            Set cllPati = New Collection
            cllPati.Add Val(Nvl(rsEInvoice!����ID)), "_����ID"
            cllPati.Add 0, "_��ҳID"
            cllPati.Add Nvl(rsEInvoice!����), "_����"
            cllPati.Add Nvl(rsEInvoice!�Ա�), "_�Ա�"
            cllPati.Add Nvl(rsEInvoice!����), "_����"
            cllPati.Add Nvl(rsEInvoice!�����), "_�����"
            cllPati.Add Nvl(rsEInvoice!סԺ��), "_סԺ��"
            cllPati.Add intInsure, "_����"
        End If
        
        If Val(Nvl(rsEInvoice!�Ƿ񻻿�)) = 1 Then
            strInvoice = strInvoice & vbCrLf & Nvl(rsEInvoice!ֽ�ʷ�Ʊ��)
        End If
        
        rsEInvoice.MoveNext
    Loop
    rsEInvoice.MoveFirst
    
    If strInvoice <> "" Then
        strInvoice = Mid(strInvoice, 3)
        If MsgBox("���ν�����������ѻ���ֽ��Ʊ�ݣ��Ƿ������ЩƱ�ݣ�" & _
            vbCrLf & strInvoice, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Do While Not rsEInvoice.EOF
        If CancelEInvoice(frmMain, Val(Nvl(rsEInvoice!ID)), cllPati, strDate, str����, byt����, bln������) = False Then Exit Function
        rsEInvoice.MoveNext
    Loop
    CancelEInvoiceFromBalance = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Terminate() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ֹ�ӿ�
    '����:��ֹ�ɹ�����true,���򷵻�False
    '˵��:
    '   ��ֹ�ӿ�,��HIS�˳�ģ��ʱ����(���磺�˳��շѹ�������),���ڱ��ӿ�����ɵ���������Ʊ�ݽӿڵ���Դ�ͷš��Ͽ����ӵ�ʱʹ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Terminate = mobjEInvoice.zlTerminate()
End Function

Public Function SplitFeeNoByEInvoice(ByVal strNos As String, ByVal int��¼���� As Integer, _
    ByRef strEInvoiceNos As String, ByRef strNoEInvoiceNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ֽⵥ�ݺţ������õ���Ʊ�ݺ�δ���õ���Ʊ�ݵķֿ�
    '���:
    '   strNos-���ݺţ�������ŷָ�
    '����:
    '   strNoEInvoiceNos-δ���õ���Ʊ�ݵĵ��ݣ���ʽ��NO1,NO2,...
    '   strEInvoiceNos-���õ���Ʊ�ݵĵ��ݣ���ʽ��NO1,NO2,...
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim varNos As Variant, str����IDs As String
    
    On Error GoTo ErrHandler
    strEInvoiceNos = "": strNoEInvoiceNos = ""
    
    strSql = "Select /*+Cardinality(j,10)*/ a.NO, Max(b.�Ƿ����Ʊ��) As ����Ʊ��" & _
                " From ������ü�¼ A, ����Ԥ����¼ B, Table(F_Str2List([1])) J" & _
                " Where a.NO=j.Column_Value And a.����ID = b.����ID And b.��¼���� <> 1" & _
                "            And a.��¼����=[2] And a.��¼״̬ in (1,3)" & _
                " Group By a.NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�����շѵ���ȡԭ����ID", strNos, int��¼����)
    
    Do While Not rsTemp.EOF
        If Val(Nvl(rsTemp!����Ʊ��)) = 0 Then
            strNoEInvoiceNos = strNoEInvoiceNos & "," & rsTemp!NO
        Else
            strEInvoiceNos = strEInvoiceNos & "," & rsTemp!NO
        End If
        rsTemp.MoveNext
    Loop
    If strEInvoiceNos <> "" Then strEInvoiceNos = Mid(strEInvoiceNos, 2)
    If strNoEInvoiceNos <> "" Then strNoEInvoiceNos = Mid(strNoEInvoiceNos, 2)
    
    SplitFeeNoByEInvoice = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBalanceIDByFeeNo(ByVal strNos As String, Optional ByVal bln������ As Boolean, _
    Optional ByVal byt���� As Byte = 1) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ���NO,��ȡԭʼ�Ľ���ID
    '���:
    '   strNos-���ݺţ�������ŷָ�
    '   byt���ϣ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨���ǲ�����ʱ����
    '����:�������һ�����ݺţ��򷵻�ԭʼ�Ľ���ID��������������ݺţ��򷵻ض������ID���Զ��ŷָ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim str����IDs As String
    
    On Error GoTo ErrHandler
    If InStr(strNos, ",") = 0 Then
        If bln������ Then
            '�������ؽ��¼��ԭ��¼����
            strSql = _
                " Select ����id As ����ID" & _
                " From (Select b.����id From ���ò����¼ B" & _
                "        Where b.No = [1] And b.��¼���� = 1 And b.��¼״̬ In (1, 3)" & _
                "        Order By b.�Ǽ�ʱ��)" & _
                " Where Rownum < 2"
        Else
            strSql = "Select ����ID From ������ü�¼ Where NO=[1] And ��¼����=[2] And ��¼״̬ in (1,3)"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�����շѵ���ȡԭ����ID", strNos, byt����)
        If rsTemp.EOF Then Exit Function
        
        GetBalanceIDByFeeNo = Val(Nvl(rsTemp!����ID))
        Exit Function
    End If
    
    If bln������ Then
        '�������ؽ��¼��ԭ��¼����
        strSql = _
            " Select ����id As ����id" & _
            " From (Select /*+Cardinality(j,10)*/" & _
            "         a.����id, Row_Number() Over(Partition By a.��¼����, a.No Order By a.�Ǽ�ʱ��) As ���" & _
            "        From ���ò����¼ A, Table(f_Str2List([1])) J" & _
            "        Where a.No = j.Column_Value And ��¼���� = 1 And ��¼״̬ In (1, 3))" & _
            " Where ��� = 1"
    Else
        strSql = _
            " Select /*+Cardinality(j,10)*/ a.����ID" & _
            " From ������ü�¼ A,Table(F_Str2List([1])) J" & _
            " Where a.NO=j.Column_Value And ��¼����=[2] And ��¼״̬ in (1,3)" & _
            " Group By a.����ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�����շѵ���ȡԭ����ID", strNos, byt����)
    
    str����IDs = ""
    Do While Not rsTemp.EOF
        str����IDs = str����IDs & "," & rsTemp!����ID
        rsTemp.MoveNext
    Loop
    GetBalanceIDByFeeNo = Mid(str����IDs, 2)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintNotice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, Optional ByVal byt���� As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ��֪��
    '���:frmMain-����������
    '     lngEInvoiceID-����Ʊ��ID
    '     byt����:��1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PrintNotice = mobjEInvoice.zlPrintNotice(frmMain, byt����, lngEInvoiceID)
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Set mfrmMain = Nothing
    Set mobjEInvoice = Nothing
End Sub