Attribute VB_Name = "mdlDrugAndStuffSvr"
Option Explicit
'*********************************************************************************************************************************************
'����:ҩƷ��������ش���
'����:
'    1.zlGetPublicDrugObjct:��ȡ����ҩƷ����
'    2.zlGetServiceObject:��ȡҩƷ�����������
'    3.zlGetStock:��ȡָ��ҩƷ������������ָ���ⷿ�еĿ��ÿ����
'    4.zlGetMultiStock:��ȡָ��ҩƷ�����������ڶ���ⷿ�еĿ��ÿ����
'    5.zlCheckWaitSendDrugAndSutff:���δ��ҩƷ������(������ʱ������true,���򷵻�false)
'    6.zlDrugSvr_RecipeAffirm:�շѡ����ʵȻ������ҩƷ����ȷ��
'    7.zlStuffSvr_BillAffirm:�շѡ����ʵȻ���������Ĵ���ȷ��
'    8.zlGetDrugSendWindows:��ȡ��ҩ����
'����:
'����:���˺�
'����:2019*08*08 19:20:59
'*********************************************************************************************************************************************
Private mobjService  As zlPublicExpense.clsService    'ҩƷ��������ط�����
Private mobjPublicDrug As Object 'ҩƷ��������,105875

Public Function zlGetPublicDrugObjct(Optional ByRef objPubDrug As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩƷ����
    '����:objPubDrug-���ع���ҩƷ��ض���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-08 21:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjPublicDrug Is Nothing Then Set objPubDrug = mobjPublicDrug: zlGetPublicDrugObjct = True: Exit Function

    Err = 0: On Error Resume Next
    Set mobjPublicDrug = CreateObject("zlPublicDrug.clsPublicDrug")
    If Err <> 0 Then
        MsgBox "ҩƷ����������zlPublicDrug������ʧ�ܣ�����ϵͳ��Ա��ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    Err = 0: On Error GoTo errHandle
    'Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If mobjPublicDrug.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
        MsgBox "ҩƷ����������zlPublicDrug����ʼ��ʧ�ܣ�����ϵͳ��Ա��ϵ��", vbInformation, gstrSysName
        Set mobjPublicDrug = Nothing: Exit Function
    End If
    zlGetPublicDrugObjct = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set mobjPublicDrug = Nothing
End Function

Public Function zlGetServiceObject(Optional ByRef objService As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������
    '����:objService-���ع����������
    '����:��ȡ����true,���򷵻�False
    '����:���˺�
    '����:2018-11-30 10:08:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If Not mobjService Is Nothing Then Set objService = mobjService: zlGetServiceObject = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjService = CreateObject("zlPublicExpense.clsService")
    If Err <> 0 Then
        MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense.clsService)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If mobjService.zlInitCommon(glngSys, glngModul, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Function
    End If
    Set objService = mobjService
    zlGetServiceObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set mobjService = Nothing
End Function

Public Function zlCheckPriceAdjustBySellFromBillDetails(ByVal objBillDetails As BillDetails) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�����ϸ����������ۼ��
    '���:objBill-���õ��ݶ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-08 21:18:15
    '˵��:���ۼ��,105875
    '
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubDrug As Object
    Dim i As Long
    
    On Error GoTo errHandle
    If zlGetPublicDrugObjct(objPubDrug) = False Then Exit Function
    If objBillDetails Is Nothing Then Exit Function
    

    'Private Function zlCheckPriceAdjustBySell(ByVal lngҩƷid As Long, ByVal lngҩ��id As Long) As Boolean
    '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
    '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    '���۳���ʱֻ�ж�ҩ��
    '���أ�True-�����������۳��⣻false-���ܽ������۳���
    For i = 1 To objBillDetails.Count
        With objBillDetails(i)
            If InStr(",5,6,7,", .�շ����) > 0 Then
                If objPubDrug.zlCheckPriceAdjustBySell(.�շ�ϸĿID, .ִ�в���ID) = False Then
                    Exit Function
                End If
            End If
        End With
    Next
    zlCheckPriceAdjustBySellFromBillDetails = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetStock(ByVal bln���� As Boolean, ByVal lng�շ�ϸĿID As Long, ByVal lng�ⷿID As Long, _
    Optional ByVal lng���� As Long = -1, Optional ByVal lngMoudle As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��ҩƷ������������ָ���ⷿ�еĿ��ÿ����(�����۵�λ)
    '���:bln����-�Ƿ���������
    '     lng�շ�ϸĿID-ҩƷID������ID
    '     lng�ⷿID-lng�ⷿID
    '     lng����-����(��ȡ���(�����������),ҩ��������(����=0,����Ϊҩ��)����Ч��)
    '     lngMoudle-��ǰ���õ�ģ���
    '����:
    '����:����ҩƷ�����Ŀ��
    '����:���˺�
    '����:2019-08-08 19:25:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    If bln���� Then
        If objService.zlStuffSvr_GetStock(lng�շ�ϸĿID, lng�ⷿID, lng����, dbl���, lngMoudle) = False Then Exit Function
    Else
        If objService.zlDrugSvr_GetStock(lng�շ�ϸĿID, lng�ⷿID, lng����, dbl���, lngMoudle) = False Then Exit Function
    End If
    zlGetStock = dbl���
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetMultiStock(ByVal lng�շ�ϸĿID As Long, ByVal str�ⷿIds As String, Optional ByVal bln���� As Boolean = False, Optional lngModule As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��ҩƷ�����������ڶ���ⷿ�еĿ��ÿ����(�����۵�λ)
    '���:lng�շ�ϸĿID-ҩƷID������ID
    '     str�ⷿIDs-�ⷿID:����ö���
    '     lngModule-ģ���
    '����:
    '����:���ؿ����
    '����:���˺�
    '����:2019-08-08 21:32:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    If zlGetServiceObject(objService) = False Then Exit Function
    If bln���� Then
        If objService.zlStuffSvr_GetMultiStock(lng�շ�ϸĿID, str�ⷿIds, dbl���, lngModule) = False Then Exit Function
    Else
        If objService.zlDrugSvr_GetMultiStock(lng�շ�ϸĿID, str�ⷿIds, dbl���, lngModule) = False Then Exit Function
    End If
    zlGetMultiStock = dbl���
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetStockInfo(lngҩƷID As Long, blnҩ�� As Boolean, blnҩ�� As Boolean, Optional ByVal dbl����ϵ�� As Double, Optional lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ�ڸ���ҩ����ҩ��Ŀ����Ϣ
    '���:objPati-������Ϣ��
    '     "blnҩ��/blnҩ��"����Ҫ��һ������Ϊ��
    '     dbl����ϵ��-����ϵ��
    '     lngModule-ģ���
    '����:
    '����:�ɹ����ؿ����Ϣ
    '����:���˺�
    '����:2019-08-08 22:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str�ⷿ���� As String, strStockInfor As String
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    
    If blnҩ�� And blnҩ�� Then
        str�ⷿ���� = "��ҩ��,��ҩ��,��ҩ��,��ҩ��,��ҩ��,��ҩ��"
    ElseIf blnҩ�� Then
        str�ⷿ���� = "��ҩ��,��ҩ��,��ҩ��"
    ElseIf blnҩ�� Then
        str�ⷿ���� = "��ҩ��,��ҩ��,��ҩ��"
    End If
  
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlDrugSvr_GetStockInfo(lngҩƷID, str�ⷿ����, dbl����ϵ��, strStockInfor, lngModule) = False Then Exit Function
    zlGetStockInfo = strStockInfor
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlCheckValidity(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, _
    Optional ByVal blnAsk As Boolean = True, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ϵ����Ч���Ƿ����
    '���:objPati-������Ϣ��
    '     blnAsk=��ʾ�Ƿ�ѯ���Ƿ����,����Ϊ����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-08 23:21:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
  
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlStuffSvr_CheckValidity(lng����ID, lng�ⷿID, dbl����, blnAsk, lngModule) = False Then Exit Function
    zlCheckValidity = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Public Function zlCheckWaitSendDrugAndSutff(ByVal str���� As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    Optional ByVal int�����Ժ��ҩ As Integer = 0, Optional int�����־ As Integer = 1, _
    Optional ByVal intӤ����� As Integer = -1, Optional ByVal strNos As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡����ҩ���Ƿ���δ��ҩ��ҩƷ������
    '���:lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     int�����Ժ��ҩ-��Ժ��ҩ
    '     int�����־-1-����;2-סԺ
    '     lngModule -����ģ���
    '����:������ʱ������true,���򷵻�False
    '����:���˺�
    '����:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotSendInfor As String
    Dim objExpenceSvr As clsExpenceSvr
    
    On Error GoTo errHandle
    
    If int�����־ = 1 Then
        If gTy_System_Para.TY_Balance.byt������δ��ҩ = 0 Then zlCheckWaitSendDrugAndSutff = True: Exit Function
    Else
        If gTy_System_Para.TY_Balance.byt���δ��ҩ = 0 Then zlCheckWaitSendDrugAndSutff = True: Exit Function
    End If
    
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then
        MsgBox "���ù�������(zlpubExpence)ʧ�ܣ��������Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objExpenceSvr.zlGetWaitExcuteDrugAndStuff(lng����ID, lng��ҳID, strNotSendInfor, int�����Ժ��ҩ, intӤ�����, int�����־, strNos, lngModule) = False Then Exit Function
    
    If strNotSendInfor = "" Then zlCheckWaitSendDrugAndSutff = True: Exit Function
    
    If gTy_System_Para.TY_Balance.byt���δ��ҩ = 1 And int�����־ <> 1 _
        Or gTy_System_Para.TY_Balance.byt������δ��ҩ = 1 And int�����־ = 1 Then
        If MsgBox("���ֲ���" & str���� & strNotSendInfor & vbCrLf & vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        zlCheckWaitSendDrugAndSutff = True: Exit Function
    End If
    MsgBox "���ֲ���" & str���� & strNotSendInfor & vbCrLf & vbCrLf & "������" & IIf(int�����־ <> 2, "����", "��Ժ") & "���ʡ�", vbInformation, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Public Function zlStuffSvr_AutoBatchSendStuff(ByVal strNO As String, ByVal str����IDs As String, ByVal strCurDate As String, ByVal str����Ա���� As String, ByVal str����Ա��� As String, _
    Optional intBillType As Integer = 2, Optional lngModule As Long, Optional strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Զ����Ϸ���
    '���:strNo-���ݺ�
    '     strCurDate-��ǰ����ʱ��
    '     str����Ids-����Ids
    '     str����Ա����-����Ա����
    '     str����Ա���-����Ա���
    '     lngModule -����ģ���
    '     intBilltype-��������(1-�շѣ�2-����,3-���ʱ�)
    '     intSendMode-��ҩ��ʽ( 1-������ҩ;2-������ҩ;3-���ŷ�ҩ)
    '����:strErrMsg_Out-��������ʱ�����ش�����Ϣ
     '����:�Զ����ϳɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-08-09 10:08:04
    '˵��:
    '   ���Ӵ��󲶻����ϼ����̲���(��������վ�ã�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objService As zlPublicExpense.clsService
    If zlGetServiceObject(objService) = False Then Exit Function
    zlStuffSvr_AutoBatchSendStuff = objService.zlStuffSvr_AutoSendStuffFromNo(strNO, str����IDs, _
        strCurDate, str����Ա����, str����Ա���, intBillType, lngModule, , strErrMsg_out)
    
End Function

Public Function zlDrugSvr_RecipeAffirm(ByVal cllRecipeData As Collection, ByVal str����ʱ�� As String, _
    ByVal str����˱�� As String, ByVal str��������� As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�շѡ����ʵȻ������ҩƷ����ȷ��
    '���:
    '   cllRecipeData-��������,ÿ��������:
    '       Array(��������,���ݺ�,����IDs,��ҩ����,�Ƿ��Զ�����,�Զ�������ϸIDs)
    '           ��������:1-�շѴ�����ҩ;2-���ʵ�������ҩ;3-���ʱ�����ҩ
    '           ��ҩ����:��ҩ����1:ҩ��ID1|��|��ҩ����n:ҩ��Idn
    '   str����ʱ��-��ǰ������ʱ��,���ʱ���������ʱ��
    '   lngModule-ģ���
    '����:
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService

    On Error GoTo errHandle
    If cllRecipeData.Count = 0 Then zlDrugSvr_RecipeAffirm = True: Exit Function '��ҩƷ��ش���ֱ�ӷ���true
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlDrugSvr_RecipeAffirm(cllRecipeData, 0, str����ʱ��, str����˱��, str���������, lngModule) = False Then Exit Function
    zlDrugSvr_RecipeAffirm = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlStuffSvr_BillAffirm(ByVal cllRecipeData As Collection, ByVal str����ʱ�� As String, _
    ByVal str����˱�� As String, ByVal str��������� As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�շѡ����ʵȻ���������Ĵ���ȷ��
    '���:
    '   cllRecipeData-��������,ÿ��������:
    '       Array(��������,���ݺ�,����IDs,�Ƿ��Զ�����,�Զ�������ϸIDs)
    '           ��������:1 -�շѴ�����ҩ;2-���ʵ�������ҩ;3-���ʱ�����ҩ
    '   str����ʱ��-��ǰ������ʱ��,���ʱ���������ʱ��
    '   lngModule-ģ���
    '����:
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    
    If cllRecipeData.Count = 0 Then zlStuffSvr_BillAffirm = True: Exit Function '��������ش���ֱ�ӷ���true
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlStuffSvr_BillAffirm(cllRecipeData, 0, str����ʱ��, str����˱��, str���������, lngModule) = False Then Exit Function
    zlStuffSvr_BillAffirm = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDrugSendWindows(ByVal lng����ID As Long, ByVal int�Һ���Ч���� As Integer, ByVal strҩ��Ids As String, _
    ByRef rsSendWindows_out As ADODB.Recordset, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ���ݵķ�ҩ����
    '���:objPati-������Ϣ��
    '    strҩ��Ids-ҩ��ID1,ȱʡ��ҩ����|ҩ��ID2,ȱʡ��ҩ����2|...
    '����:rsSendWindows_out-ҩ��ID,��ҩ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-16 11:40:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    On Error GoTo errHandle
        
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlDrugSvr_GetSendWindows(2, lng����ID, int�Һ���Ч����, strҩ��Ids, rsSendWindows_out, lngModule) = False Then Exit Function
    
    zlGetDrugSendWindows = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 


Public Function zlGetDrugAndStuff_ExcuteNum(ByVal strNos As String, ByRef rsRecipe As ADODB.Recordset, ByVal blnExistDrug As Boolean, _
    ByVal blnExistStuff As Boolean, Optional ByVal bytBillType As Byte = 2, Optional lngModule As Long, Optional ByVal str����IDs As String, _
    Optional blnAppend As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ׼������
    '���:strNos:���ݺţ������Ӣ�Ķ��ŷָ�
    '   bytBillType:1-�շѵ�,2-���ʵ�
    '   blnExistDrug:�Ƿ�ҩƷ
    '   blnExistStuff:�Ƿ�����
    '   str����IDs-����ö��ţ������ǰ��ĵ��ݺż��������;���Ч
    '   blnAppend-�Ƿ��Զ�׷������
    '����: rsRecipe:�����������ݣ���������,ҩƷID,������ϸID,�ѷ�����,��Ʒ����,�ڲ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-16 14:48:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objExpenceSvr As clsExpenceSvr
    On Error GoTo errHandle
        
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
    If objExpenceSvr.zlGetDrugStuff_ExecutedNum(strNos, bytBillType, rsRecipe, _
        blnExistDrug, blnExistStuff, blnAppend, str����IDs, , , lngModule) = False Then Exit Function
    zlGetDrugAndStuff_ExcuteNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

 
Public Function zlExcute_SendDrug(strNO As String, strTime As String, Optional ByVal intBillType As Integer = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ҩƷ���Ų���
    '���:strNO-���ݺ�
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-18 15:04:40
    '˵������ͨ��ҩʱΪ���˿��ң����ҽ����Ϊ�������ҡ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, dtCurdate As Date
    Dim objService As zlPublicExpense.clsService, objExpenceSvr As clsExpenceSvr
    Dim blnTrans As Boolean
    Dim str����IDs As String, cllFeeIds As Collection, str��ֹ��ҩ����IDs As String
    On Error GoTo errHandle
        
    If zlGetServiceObject(objService) = False Then Exit Function
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
  
    strSQL = _
    " Select  b.ID, b.ִ�в���ID,c.ͬ����־ As �Ƿ�ͬ����־, c1.ͬ����־ As ����ͬ����־ " & _
    " From סԺ���ü�¼ B, ���˷����쳣��¼ C, ���˷����쳣��¼ C1" & _
    " Where b.NO=[1] And b.��¼����=2 And  b.��¼״̬ in (0,1,3) And b.�۸񸸺� is null And nvl(b.ִ��״̬,0)<>1 And B.�Ǽ�ʱ��+0=[3]" & _
    "       And instr(',5,6,7,',','||b.�շ����||',')>0 " & _
    "       And b.ID = c.����ID(+) And c.��������(+) = 0" & _
    "       And b.ID = c1.����ID(+) And c1.��������(+) = 1"
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, 2, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, 2)
    End If
    
    If rsTmp.RecordCount > 0 Then '�ų�������Һ�������ĵ�ҩƷ
        If mobjService.zlPivasSvr_Getinfusion_Record(strNO, str��ֹ��ҩ����IDs) = False Then Exit Function
    End If
    
    With rsTmp
        Do While Not .EOF
            If Val(Nvl(!�Ƿ�ͬ����־)) <> 0 Or Val(Nvl(!����ͬ����־)) <> 0 Then
                MsgBox "���ݡ�" & strNO & "��Ϊ�쳣���ݣ�����������ռ�ã����Ժ�����!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If Nvl(!ִ�в���ID) = "" Then
                MsgBox "���ݡ�" & strNO & "���д���δȷ��ִ��ҩ�������������﷢ҩ��", vbInformation, gstrSysName
                Exit Function
            End If
            
            If InStr("," & str��ֹ��ҩ����IDs & ",", "," & Val(Nvl(!ID)) & ",") = 0 Then
                str����IDs = IIf(str����IDs = "", "", str����IDs & ",") & Val(Nvl(!ID))
            End If
            .MoveNext
        Loop
    End With
    
    If str����IDs = "" Then
        MsgBox "����""" & strNO & """��ǰ������û�п��Է��ŵ�ҩƷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    Set cllFeeIds = New Collection  'array(����ids,ִ��״̬)
    cllFeeIds.Add Array(str����IDs, 1)
    
    dtCurdate = zlDatabase.Currentdate
    
    gcnOracle.BeginTrans: blnTrans = True
    
    If objExpenceSvr.zlUpdateExcuteStatu(cllFeeIds) = False Then Exit Function
    If objService.zlDrugsvr_AutoSendDrugFromNo(strNO, Format(dtCurdate, "yyyy-mm-dd HH:MM:SS"), UserInfo.����, UserInfo.���, intBillType) = False Then Exit Function
   
    gcnOracle.CommitTrans: blnTrans = False
    
    zlExcute_SendDrug = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Function zlGetPrice(ByVal objBillDetail As BillDetail, ByVal dbl���� As Double, ByRef dbl�ɱ��� As Double, _
    Optional ByVal lngRow As Long, Optional ByVal lngModule As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�۸�
    '���:������ϸ����
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-18 16:02:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim dblPrice As Double, dbl��ȱ���� As Double
        
    On Error GoTo errHandle
    dbl�ɱ��� = 0
    If zlGetServiceObject(objService) = False Then Exit Function
  
  
    '��ȡҩƷ/���ļ۸�
    If objBillDetail Is Nothing Then Exit Function
    
    With objBillDetail
        If .�շ���� = "4" Then
            If objService.zlStuffSvr_GetPrice(.�շ�ϸĿID, .ִ�в���ID, dbl����, .Detail.����, "", _
                  dblPrice, dbl�ɱ���, dbl��ȱ����, lngModule) = False Then
                '��ȡ�۸�ʧ��
                MsgBox IIf(lngRow > 0, "��" & lngRow & "��", "") & "��������""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If objService.zlDrugSvr_GetPrice(.�շ�ϸĿID, .ִ�в���ID, dbl����, .Detail.����, "", _
                  dblPrice, dbl�ɱ���, dbl��ȱ����, lngModule) = False Then
                '��ȡ�۸�ʧ��
                MsgBox IIf(lngRow > 0, "��" & lngRow & "��", "") & "ҩƷ""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If dbl��ȱ���� <> 0 And .Detail.��� Then
            '����δ�ֽ����
            If InStr(",5,6,7,", .�շ����) > 0 Then
                MsgBox IIf(lngRow > 0, "��" & lngRow & "��", "") & "ʱ��ҩƷ""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
            Else
                MsgBox IIf(lngRow > 0, "��" & lngRow & "��", "") & "ʱ����������""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End With
    zlGetPrice = dblPrice
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlAutoSplitSpeci(ByVal lngҩ��ID As Long, ByVal byt��ҩ��̬ As Byte, _
    ByVal int���� As Integer, ByVal dbl���� As Double, ByVal lngҩ��ID As Long, _
    Optional ByVal byt���� As Byte = 1, Optional lngModule As Long) As String
    '����:����в�ҩ����ҩ����Ʒ�֣����Զ�����ҩƷ(�Զ��ֽ������)
    '���:
    '   byt��ҩ��̬ 0-ɢװ;1-��ҩ��Ƭ;2-����
    '   byt���� 1-���� ��2-סԺ
    '����:
    '����:��ʽ��ҩƷid,����;ҩƷid,����;...(ɢװֻѡ��һ�����)
    '               ������ȫ����ʱ����:����Ϊ6��10�������,17�˵ķ���=23755,6;23756,10|1
    '               ���ܷ���ʱ���ؿ�,����:����Ϊ6��10�������,3�˵ķ���
    '����:���˺�
    '����:2019-08-18 15:44:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
        
    If zlGetServiceObject(objService) = False Then Exit Function

    If objService.zlDrugSvr_AutoSplitSpeci(lngҩ��ID, byt��ҩ��̬, int����, dbl����, lngҩ��ID, strData, byt����, lngModule) = False Then Exit Function
    zlAutoSplitSpeci = strData
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ZlDrugsvr_GetStockByDrugName(ByVal lngҩ��ID As Long, ByVal lngҩ��ID As Long, int��ʾ��λ As Integer, _
ByRef cllStockData_Out As Collection, ByVal int���� As Integer, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ��,��ȡָ�������Ϣ
    '���:lngҩ��ID-
    '     lngҩ��ID-ҩ��ID,=0ʱ��ʾ������ҩ��
    '     int����-1-����;2-סԺ
    '     int��ʾ��λ-0-�ۼ۵�λ;1-���ﵥλ;2-סԺ��λ
    '����:cllStockData_Out-�������ݼ���ÿ����Ա��key(pharmacy_id,drug_id,stock)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-18 17:30:45
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
        
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlDrugsvr_GetStockByDrugName(lngҩ��ID, lngҩ��ID, int��ʾ��λ, cllStockData_Out, int����, lngModule) = False Then Exit Function
    ZlDrugsvr_GetStockByDrugName = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetStockCheck(ByVal bytType As Byte, Optional lngModule As Long) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ�����ĳ�����ļ���
    '���:bytType:0-ҩƷ��1-����
    '����:���ؿ���鷽ʽ
    '����:���˺�
    '����:2019-08-18 20:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllStockCheck As Collection, colStock As Collection
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
       
    If zlGetServiceObject(objService) = False Then GoTo ReInit
    If bytType = 1 Then
        If objService.zlStuffSvr_GetStockCheck(cllStockCheck, lngModule) = False Then GoTo ReInit
    Else
        If objService.zlDrugSvr_GetStockCheck(cllStockCheck, lngModule) = False Then GoTo ReInit
    End If
    
    Err = 0: On Error Resume Next
    cllStockCheck.Add 0, "_0"
    Set zlGetStockCheck = cllStockCheck
    On Error GoTo 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
ReInit:
    Set colStock = New Collection
    colStock.Add 0, "_0" '�������
    Set zlGetStockCheck = colStock
End Function


Public Function zlHaveNOAuditing(ByVal lng����ID As Long, Optional ByVal str��ҳIDS As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϲ���δ��������Ƿ����δ��˼��ʷ���
    '���:str��ҳIds-����ö��ŷ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-19 09:20:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, str����IDs As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim objService As zlPublicExpense.clsService
    Dim strNos As String, rsAdvice As ADODB.Recordset
    
    On Error GoTo errHandle
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlDrugSvr_GetRefuseSendList(lng����ID, str��ҳIDS, str����IDs, lngModule) = False Then Exit Function
 
    strSQL = _
        " Select Distinct a.ҽ�����,a.No,a.��¼���� From סԺ���ü�¼ A" & _
        " Where ���ʷ���=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And ����ID=[1]" & _
                IIf(str����IDs <> "", " And instr([3] ,','||ID||',')=0 ", "") & _
                IIf(str��ҳIDS <> "", " And a.��ҳID In(Select /*+cardinality(j,10)*/ Column_Value From Table(f_num2list([2])) J)", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, str��ҳIDS, str����IDs)
    If rsTmp.RecordCount = 0 Then Exit Function
    
    '�Ѿ����ڷ�ҽ��δ��˵ķ��þͲ����ټ��ҽ����
    strNos = ""
    Do While rsTmp.EOF
        If Val(Nvl(rsTmp!ҽ�����)) = 0 Then
            zlHaveNOAuditing = True: Exit Function
        Else
            strNos = strNos & "," & Nvl(rsTmp!ҽ�����) & ":" & Nvl(rsTmp!NO) & ":" & Nvl(rsTmp!��¼����)
        End If
        rsTmp.MoveNext
    Loop
    
    strNos = Mid(strNos, 2)
    If ZLGetAdviceSendInfo(1, strNos, rsAdvice) = False Then zlHaveNOAuditing = True: Exit Function   '���ز���
    
    rsAdvice.Filter = "ִ��״̬<>2" '0-δִ��;1-��ȫִ��;2-�ܾ�ִ��;3-����ִ��(�����ֽܷ�Ϊ����ʵ�ʲ���)
    zlHaveNOAuditing = Not rsAdvice.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlCheckChargeAudit(ByVal lng����ID As Long, ByVal bln��Ժ As Boolean, _
    Optional blnSaveCheck As Boolean = False, _
    Optional ByVal str��ҳIDS As String = "", Optional bln��ѡ��;����_out As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������˼��
    '���:
    '����:bln��ѡ��;����_out -����ѡ������;����
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 15:04:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'bytAuditing:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strWhere As String, str����IDs As String
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    bln��ѡ��;����_out = False
    If zlHaveNOAuditing(lng����ID, str��ҳIDS, lngModule) = False Then zlCheckChargeAudit = True: Exit Function
    
    Select Case gTy_System_Para.TY_Balance.bytAuditing
    Case 1
        '�ڶ�ȡ������Ϣʱ,�Ѿ���ʾ��
        If Not blnSaveCheck Then
            If MsgBox("�ò��˻�����δ��˵ļ��ʷ��ã�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    Case 2
        If blnSaveCheck Then
            If bln��Ժ Then
                MsgBox "�ò��˻�����δ��˵ļ��ʷ���,���ܳ�Ժ���ʣ�", vbInformation, gstrSysName
                Exit Function
            End If
            '�ڶ�ȡ������Ϣʱ,�Ѿ���ʾ��
        Else
            If MsgBox("�ò��˻�����δ��˵ļ��ʷ��ã����ܳ�Ժ���ʣ�" & vbCrLf & vbCrLf & _
                "�Ƿ�Ըò��˽�����;���ʣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                bln��ѡ��;����_out = True
        End If
    Case Else
    End Select
    zlCheckChargeAudit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function zlPivasSvr_Getinfusion_RecordFeeids(ByVal strNos As String, ByRef str����IDs_out As String, Optional ByVal lng����ID As Long, Optional ByVal str��ҳIDS As String, _
    Optional ByVal str����Ids_in As String = "", Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˵ľܷ�ҩ�嵥
    '���:lng����ID-������ID��(str��ҳIds-����ö���)
    '     strNos-�����ݲ�
    '     str����Ids_in-������ID��
    '     ��������������һ��
    '����:str����IDs_out-���������漰������Һ��ҩ��¼�еķ���IDs
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
     
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    zlPivasSvr_Getinfusion_RecordFeeids = objService.zlPivasSvr_Getinfusion_Record(strNos, str����IDs_out, lng����ID, str��ҳIDS, str����Ids_in, lngModule)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlPivasSvr_Isexsitinfusion(ByVal str����IDs As String, Optional blnIsExist_Out As Boolean, Optional lngModule As Long, _
    Optional lngҽ��ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҩƷ�Ƿ������Һ��ҩ��¼
    '���: str����Ids_in-������ID��
    '      lngҽ��ID-�������ҽ��id,��ҽ��ID��֤
    '����:blnIsExist_Out-�Ƿ���ڵģ�����true,���򷵻�False
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
     
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    zlPivasSvr_Isexsitinfusion = objService.zlPivasSvr_Isexsitinfusion(str����IDs, blnIsExist_Out, lngModule, lngҽ��ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ZlStuffSvr_AutoReturnStuff(ByVal strAutoStuffDatas As String, _
    Optional ByVal strTittle As String, Optional lngModule As Long, Optional rsSendDatas As ADODB.Recordset, _
    Optional ByVal bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ����ϲ���
    '���:strAutoStuffDatas-�Զ���������,��ʽ:����ID1:����,����ID2:����2,...
    '    rsSendDatas-�����Ѿ�ִ������
    '    blnReturnAll-�Ƿ����漰�ķ���ID����Ӧ��ʣ��δ�˵�ȫ�ˣ�ȫ��ʱ��strAutoStuffDatas������������)
    '    bln����������-�ڲ�����������
    '����:strErrMsg_out-ʧ��ʱ�����ش�����Ϣ
    '����:���ϳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService, strErrMsg_out As String
    Dim blnTrans As Boolean, varData As Variant, strFeeIds As String, lng����ID As Long
    Dim i As Long, strSQL As String, varParaValue()  As Variant, rsTemp As ADODB.Recordset
    Dim cllExcuteUpdate As Collection, intִ��״̬ As Integer, dblִ���� As Double
    Dim strSubTable As String, strFeeTable As String
    Dim strAutoStuff As String, dbl������ As Double
    
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    'Ϊ��ʱ��ֱ�ӷ���true
    If strAutoStuffDatas = "" Then ZlStuffSvr_AutoReturnStuff = True: Exit Function
    
    
    If rsSendDatas Is Nothing Then
        '���»�ȡ
        strFeeIds = ""
        varData = Split(strAutoStuffDatas, ",")
        For i = 0 To UBound(varData)
            lng����ID = Split(varData(i) & ":", ":")(0)
            If InStr(strFeeIds & ",", "," & lng����ID & ",") > 0 Then
                 strFeeIds = strFeeIds & "," & lng����ID
            End If
        Next
        If objService.zlStuffSvr_GetExecutedNum("", 2, rsSendDatas, , strFeeIds, lngModule, True, strErrMsg_out) = False Then Exit Function
    End If
    
    '���:
    '   bytType: 0-Num2List;1-Str2List;2-Num2List2;3-Str2List2
    '   strValues: bytType=0,1ʱ,�����","����
    '              bytType=2,3ʱ,��֮����":"����,��֮����","����:��:����:22,����:22
    '   lngStep: ����(���󶨱����Ӻö࿪ʼ)

    If zlGetVarBoundSQL(2, strAutoStuffDatas, strSubTable, varParaValue, 0) = False Then Exit Function
    
    strFeeTable = IIf(bln����, "������ü�¼", "סԺ���ü�¼")
    
    strSQL = "With ������Ϣ As (" & strSubTable & ")" & vbCrLf & _
    "   Select  A.No,a.���,a.�շ�ϸĿID,max(Decode(a.��¼״̬,2,0,a.ID)) as ����id, " & vbCrLf & _
    "           sum(decode(a.��¼״̬,2,0,1)*nvl(a.����,1)*nvl(a.����,0)) as ԭʼ����," & vbCrLf & _
    "           sum(nvl(a.����,1)*nvl(a.����,0)) as ʣ������," & vbCrLf & _
    "           max(B.��������) as �������� " & vbCrLf & _
    "   From " & strFeeTable & " a ," & _
    "         ( Select /*+CARDINALITY(M,10)*/ distinct J.No,J.���,C2 as �������� From " & strFeeTable & " J,������Ϣ  M  Where J.ID=M.C1 ) B " & vbCrLf & _
    "   Where a.NO=B.NO and a.���=b.��� And a.�۸񸸺� is null " & _
    "   Group by A.NO,A.���,a.�շ�ϸĿID "
    
    strSQL = "" & _
    "    Select A.No,A.���,A.����ID,A.ԭʼ����,A.ʣ������,A.��������,C.���� as �շ���Ŀ" & vbCrLf & _
    "    From (" & strSQL & ") A,�շ���ĿĿ¼ C" & vbCrLf & _
    "    Where  a.�շ�ϸĿID=C.ID" & vbCrLf & _
    "    Order by a.no,a.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, strTittle, varParaValue)
    
    '�ȼ��
    Set cllExcuteUpdate = New Collection
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        strAutoStuff = ""
        Do While Not .EOF
            
            If Val(Nvl(!ʣ������)) = 0 Then
                strErrMsg_out = "��" & !���� & "���Ѿ�û�п������ˣ���������Ϊ����ԭ���������ʣ���������!"
                MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If Val(Nvl(!ʣ������)) < Val(Nvl(!��������)) Then
                strErrMsg_out = "��" & !���� & "�����������ϵĿ���������С����������������������!" & vbCrLf & _
                       "ʣ������:" & FormatEx(Val(Nvl(!ʣ������)), 4) & vbCrLf & _
                       "��������:" & FormatEx(Val(Nvl(!��������)), 4) & vbCrLf
                MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            
            dblִ���� = 0
            rsSendDatas.Filter = "������ϸID=" & Val(Nvl(!����ID))
            If rsSendDatas.EOF = False Then dblִ���� = Val(Nvl(rsSendDatas!�ѷ�����))
            rsSendDatas.Filter = 0
            
            '���㱾���Զ���������,��������:
            '1.��������>(ʣ������-��ִ������) �򱾴��Զ�����=��������-(ʣ������-��ִ������)
            dbl������ = Val(Nvl(!��������)) - (Val(Nvl(!ʣ������)) - dblִ����)
        
            If dblִ���� < dbl������ And dbl������ <> 0 Then
                strErrMsg_out = "��" & !���� & "�����������ϵı������������������ѷ�������,���ܽ����Զ����ϲ���!" & vbCrLf & _
                       "�ѷ�������:" & FormatEx(dblִ����, 4) & vbCrLf & _
                       "������������:" & FormatEx(dbl������, 4) & vbCrLf
                MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If dbl������ > 0 Then
                If RoundEx(dblִ���� - dbl������, 5) > 0 Then
                    '����ִ��
                    intִ��״̬ = 2
                Else
                    intִ��״̬ = 0
                End If
                cllExcuteUpdate.Add Array(Val(Nvl(!����ID)), intִ��״̬)
                strAutoStuff = strAutoStuff & "," & Val(Nvl(!����ID)) & ":" & dbl������
            End If
            .MoveNext
        Loop
    End With
    If strAutoStuff = "" Then ZlStuffSvr_AutoReturnStuff = True: Exit Function
    strAutoStuff = Mid(strAutoStuff, 2)
 
    '�Զ�����
    gcnOracle.BeginTrans: blnTrans = True
    '���µ����ݼ�(array(����ids,ִ��״̬))
    If mdlExseSvr.zlUpdateExcuteStatu(cllExcuteUpdate, IIf(bln����, 1, 2)) = False Then gcnOracle.RollbackTrans: Exit Function
    If objService.ZlStuffSvr_AutoReturnStuff(strAutoStuff, strErrMsg_out, strTittle, lngModule) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
        If strErrMsg_out <> "" Then MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    ZlStuffSvr_AutoReturnStuff = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ZlStuffSvr_AutoReturnStuffFromFeeIds(ByVal strAutoStuffDatas As String, _
    Optional ByVal strTittle As String, Optional lngModule As Long, _
    Optional ByVal bln���� As Boolean, Optional strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ���ID�Զ����ϲ���
    '���:strAutoStuffDatas-�Զ���������,��ʽ:����ID1,����ID2,...
    '����:strErrMsg_out-ʧ��ʱ�����ش�����Ϣ
    '����:���ϳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim varData As Variant
    Dim i As Long, cllExcuteUpdate As Collection
    Dim strAutoStuff As String, blnTrans As Boolean
    
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    'Ϊ��ʱ��ֱ�ӷ���true
    If strAutoStuffDatas = "" Then ZlStuffSvr_AutoReturnStuffFromFeeIds = True: Exit Function
    
    varData = Split(strAutoStuffDatas, ",")
    Set cllExcuteUpdate = New Collection
    For i = 0 To UBound(varData)
        cllExcuteUpdate.Add Array(varData(i), 0)
        strAutoStuff = strAutoStuff & "," & varData(i) & ":"    '�ӿ�Ҫ�󣬱ش�:�ţ���ʾ��ʣ��ȫ��
    Next
    If strAutoStuff <> "" Then strAutoStuff = Mid(strAutoStuff, 2)
     
     '���µ����ݼ�(array(����ids,ִ��״̬))
    gcnOracle.BeginTrans: blnTrans = True
    If mdlExseSvr.zlUpdateExcuteStatu(cllExcuteUpdate, IIf(bln����, 1, 2)) = False Then gcnOracle.RollbackTrans: Exit Function
    If objService.ZlStuffSvr_AutoReturnStuff(strAutoStuff, strErrMsg_out, strTittle, lngModule) = False Then gcnOracle.RollbackTrans: Exit Function
    gcnOracle.CommitTrans: blnTrans = False
    ZlStuffSvr_AutoReturnStuffFromFeeIds = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    strErrMsg_out = Err.Description
End Function


Public Function zlExecuteUpdateSyncSymbol(ByVal str����IDs As String, ByVal byt��־���� As Byte, _
    ByVal byt�����־ As Byte, ByVal bytԭ��־ֵ As Byte, Optional ByVal byt�±�־ֵ As Byte = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���·��ü�¼�е�ҩƷ/����ͬ����־
    '���:
    '   str����IDs ��Ҫ���µķ���ID�������Ӣ�Ķ��ŷָ�
    '   byt��־���� ��־���ͣ�0-�Ƿ�ͬ����־,1-����ͬ����־
    '   byt�����־ �����־��1-���2-סԺ
    '����:
    '����:
    '  �Ƿ�ͬ����־��NULL��0-������1-δ���ɴ�����/���ϵ���2-δ����ҩƷ/�����շ�״̬
    '  ����ͬ����־��NULL��0-������1-ҩƷ/���������ϵ�����δ����(��ֹ����/�˻�)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objExpenceSvr As clsExpenceSvr
    On Error GoTo errHandle
    
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
     
    If objExpenceSvr.zlExecuteUpdateSyncSymbol(str����IDs, byt��־����, byt�����־, bytԭ��־ֵ, byt�±�־ֵ) = False Then Exit Function
    zlExecuteUpdateSyncSymbol = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

