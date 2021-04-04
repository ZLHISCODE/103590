Attribute VB_Name = "mdlExseSvr"
Option Explicit

'*********************************************************************************************************************************************
'����:������ط���ӿڴ���
'����:
' һ����������
'    1.GetJsonNodeString:���ݽڵ�����ڵ�ֵ��ȡJson��
'    2.GetNodeString:��ʽ���ڵ�����
'    3.zlGetPubExseSvrObject:��ȡҩƷ�����������
'    4.zlGetRecipe_ID:��ȡ����ID
' ����ҵ������
'    1.zlHospitalization_Charge_Verfiy_isValied:סԺ������˺Ϸ��Լ��
'    2.zlUpdateExcuteStatu:�޸Ĳ��˷��ü�¼��ִ��״̬
'    3.zlExcuteBillVerfiy:ִ�м��ʵ���˲���
'    4.SaveBill_NewRecipeBill-���ʲ���Ϸ��Լ��
'����:
'����:���˺�
'����:2019*08*08 19:20:59
'*********************************************************************************************************************************************
Private mobjPubExseSvr As clsExpenceSvr '������ط���ӿ�

Public Function zlGetPubExseSvrObject(ByRef objPubExseSvr As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ�����Ĺ����������
    '����:objPubExseSvr-����ҩƷ�����Ĺ����������
    '����:��ȡ����true,���򷵻�False
    '����:���˺�
    '����:2018-11-30 10:08:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If Not mobjPubExseSvr Is Nothing Then Set objPubExseSvr = mobjPubExseSvr: zlGetPubExseSvrObject = True: Exit Function
    
    
    Err = 0: On Error Resume Next
    Set mobjPubExseSvr = CreateObject("zlPublicExpense.clsExpenceSvr")
    If Err <> 0 Then
        MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense.clsExpenceSvr)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
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
    If mobjPubExseSvr.zlInitCommon(glngSys, glngModul, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Function
    End If
    Set objPubExseSvr = mobjPubExseSvr
    zlGetPubExseSvrObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set mobjPubExseSvr = Nothing
End Function

Public Function zlReadInPatientDelBillData(ByVal int���� As Integer, ByVal strNO As String, _
    Optional ByVal blnסԺ��λ As Boolean, Optional ByVal bln�ѽ��ֹ���� As Boolean, _
    Optional ByVal bln��ֹ�������� As Boolean, Optional ByVal str�Ǽ�ʱ�� As String, _
    Optional ByVal str�շ����s As String, Optional ByVal str�ų��շ����s As String, _
    Optional ByRef rsBillData_out As ADODB.Recordset, Optional ByRef rsIncome_out As ADODB.Recordset, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long, _
    Optional ByVal str����IDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ���ʵĵ�������
    '���:
    '   int����-1-�շѵ�;2-���ʵ�;3-�Զ����ʵ�;
    '   strNo-���ݺ�
    '   str�Ǽ�ʱ��-����ʱ��
    '   lngModule-ģ���
    '   bln��ֹ��������-
    '   str�շ����s-����ö��ŷ���,��"5,6,7"
    '   str�ų��շ����s-����ö��ŷ���,��"5,6,7"
    '   str����IDs-�ಡ�˵������ʵĲ���IDs,��"1,2,3"
    '����:rsBillData_out-��������
    '     rsIncome_out-������ܼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-16 20:04:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    If objPubExseSvr.zlReadInPatientDelBillData(int����, strNO, blnסԺ��λ, _
        bln�ѽ��ֹ����, bln��ֹ��������, str�Ǽ�ʱ��, str�շ����s, str�ų��շ����s, _
        rsBillData_out, rsIncome_out, strFrmCaption, lngModule, str����IDs) = False Then Exit Function
    
    zlReadInPatientDelBillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlMzTOZyExceptionUpdate(ByVal strNos As String, Optional ByVal lngModule As Long, _
    Optional ByVal bln�������� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ǰ�����Ƿ�����תסԺ�쳣����������תסԺ�쳣
    '���:
    '     strNos-���ݺţ���ʽ:A001,A002,...
    '     lngModule-ģ���
    '     bln��������-�Ƿ���������(true-��������,false-����)
    '����:strErrMsg_Out-���ش�����Ϣ(����blnShowMsg=falseʱ)
    '����:�������쳣�ҳɹ������쳣����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objPubExseSvr As clsExpenceSvr, strErrMsg As String
    Dim strSubTable As String, varPara() As Variant

    On Error GoTo errHandle
    If zlGetVarBoundSQL(1, strNos, strSubTable, varPara, 0) = False Then Exit Function
    strSQL = _
        " Select Distinct a.No" & _
        " From סԺ���ü�¼ A, (" & strSubTable & ") B" & _
        " Where NO = b.Column_Value And ��¼���� = 2 " & _
        "   And Not Exists(Select 1 From ���˷����쳣��¼ c Where c.����id = a.id And c.�������� = 0 And c.ͬ����־ = 1) " & _
        "   And Exists(Select 1 From ���˷����쳣��¼ c Where c.����id = a.id And c.�������� = 2)"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "���ת��ͬ����־", varPara)
    If rsTemp.EOF Then zlMzTOZyExceptionUpdate = True: Exit Function
    
    strNos = ""
    Do While Not rsTemp.EOF
        strNos = strNos & "," & Nvl(rsTemp!NO)
        rsTemp.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    '����ʱ������ͬ������
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    If objPubExseSvr.zlAdjustFeeData(strNos, True, False, strErrMsg) = False Then
        If strErrMsg = "" Then
            Call MsgBox("����[" & strNos & "]Ϊ�������תסԺ���ݣ�Ŀǰ���������쳣״̬����ֹ��������" & IIf(bln��������, "����", "") & "��", vbInformation + vbOKOnly, gstrSysName)
        Else
            Call MsgBox("����[" & strNos & "]Ϊ�������תסԺ���ݣ�Ŀǰ���������쳣״̬����ֹ��������" & IIf(bln��������, "����", "") & "����ϸ����ԭ������: " & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        End If
        Exit Function
    End If
    zlMzTOZyExceptionUpdate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExcuteBillVerfiy(ByVal strNO As String, ByVal str��� As String, ByVal dt���ʱ�� As Date, _
    ByVal lngModule As Long, Optional ByVal strInsure As String, Optional ByRef blnAutoSendDrug As Boolean, _
    Optional ByVal bln�Ƿ�ಡ�˵� As Boolean, Optional ByVal lng����ID As Long, _
    Optional ByVal byt������Դ As Byte = 2, Optional ByRef blnStuffSync As Boolean, Optional ByRef blnDrugSync As Boolean, _
    Optional ByRef blnAutoSendStuff As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��סԺ���ʵ���˲���
    '���:
    '   strNo-ִ�еĵ��ݺ�
    '   str���-��˵����
    '   blnAutoSendDrug-�Ƿ����Զ�����ҩƷ
    '   strInsure-���ʱ�ʱ�����ҽ��(����1,����2,...)
    '   bln�Ƿ�ಡ�˵�-�Ƿ�ಡ�˵�(���ʱ�)
    '   lng����ID-ֻ���ָ������,���ڰ�������˼��ʱ�
    '   byt������Դ-������Դ��1-���2-סԺ
    '   blnAutoSendStuff-������Դ=2ʱ��Ч���Ƿ��Զ���ҩ���ģ�True=���� Zl_סԺ���ʼ�¼_Verify_Check ���ؿ����Զ����ϣ�false=���Զ�����
    '����:
    '   blnAutoSendDrug-True:�Ѿ��Զ�����ҩƷ;false-δ�Զ����ųɹ�ҩƷ
    '   blnStuffSync-���������Ƿ���ͬ��
    '   blnDrugSync-ҩƷ�����Ƿ���ͬ��
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
    
    On Error GoTo ErrHandler
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    
    If byt������Դ = 2 Then
        If strInsure = "0" Then strInsure = ""
        If objPubExseSvr.zlExcute_InBillVerfiy(strNO, str���, dt���ʱ��, lngModule, _
            strInsure, blnAutoSendDrug, bln�Ƿ�ಡ�˵�, lng����ID, blnStuffSync, blnDrugSync, blnAutoSendStuff) = False Then Exit Function
    Else
        'strNos-������Ϣ, ��ʽ��NO1:���1,���2,...|NO1:���1,���2,...|...
        If objPubExseSvr.zlVerfyBillingPriceBill(1, strNO & ":" & str���, Format(dt���ʱ��, "yyyy-MM-dd HH:mm:ss")) = False Then Exit Function
        'ҩƷ���շ�״̬ȷ��
        blnDrugSync = objPubExseSvr.zlDrugOutRecipeAffirm(strNO, 1, 2)
        '�������շ�״̬ȷ��
        blnStuffSync = objPubExseSvr.zlStuffOutBillAffirm(strNO, 1, 2)
    End If
    zlExcuteBillVerfiy = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlUpdateExcuteStatu(ByVal cllData As Collection, _
    Optional ByVal int������Դ As Integer = 2, Optional ByVal blnAutoCalc As Boolean = False, Optional ByVal strִ����� As String, _
    Optional str����Ա���� As String, Optional str����ʱ�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ĳ��˷��ü�¼��ִ��״̬
    '���:cllData-���µ����ݼ�(array(����ids,ִ��״̬,�ѷ���)
    '     int������Դ-1-����;2-סԺ;0-�����������סԺ
    '     blnAutoCalc-�Ƿ��Զ�����ִ��״̬
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-10 16:42:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
    
    On Error GoTo ErrHandler
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    zlUpdateExcuteStatu = objPubExseSvr.zlUpdateExcuteStatu(cllData, int������Դ, blnAutoCalc, strִ�����, str����Ա����, str����ʱ��)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlUpdateExcuteStautsFromFeeIDs(ByVal strFeeIds As String, ByVal byt��� As Byte, Optional byt������Դ As Byte = 0, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ���ID,�Զ�����ִ��״̬
    '���:strFeeIds-���漰�ķ���ID
    '     byt���-���漰�����:0-ҩƷ;1-����,2-����
    '     byt������Դ-1-����;2-סԺ;0-�����������סԺ
    '����:�����ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-24 11:49:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnExistStuff As Boolean, blnExistDrug As Boolean
    Dim cllUpdate As Collection, varData As Variant, rsRecipe As ADODB.Recordset '��������,ҩƷID,������ϸID,�ѷ�����,��Ʒ����,�ڲ�����
    Dim dbl�ѷ����� As Double, i As Long
    
    On Error GoTo errHandle
    blnExistDrug = byt��� = 0 Or byt��� = 2
    blnExistStuff = byt��� = 1 Or byt��� = 2
    If strFeeIds = "" Then zlUpdateExcuteStautsFromFeeIDs = True: Exit Function
    If mdlDrugAndStuffSvr.zlGetDrugAndStuff_ExcuteNum("", rsRecipe, blnExistDrug, blnExistStuff, , lngModule, strFeeIds) = False Then Exit Function
    
    varData = Split(strFeeIds, ",")
    Set cllUpdate = New Collection
    'cllData-���µ����ݼ�(array(����ids,ִ��״̬,�ѷ���)
    For i = 0 To UBound(varData)
        rsRecipe.Filter = "������ϸID=" & Val(varData(i))
        dbl�ѷ����� = 0
        If Not rsRecipe.EOF Then dbl�ѷ����� = Val(Nvl(rsRecipe!�ѷ�����))
        
        cllUpdate.Add Array(Val(varData(i)), 0, dbl�ѷ�����)
    Next
    
    If zlUpdateExcuteStatu(cllUpdate, byt������Դ, True) = False Then Exit Function
    zlUpdateExcuteStautsFromFeeIDs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlSaveBill_NewRecipeBill(ByVal strNO As String, ByVal cllRcpBillData As Collection, _
    ByVal strFrmCaption As String, ByVal bln���� As Boolean, ByVal intBillType As Integer, _
    Optional ByVal lngModule As Long, Optional bln���� As Boolean, Optional ByRef blnSendMateria As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�м��ʱ������
    '���:
    ' cllRcpBillData(�ṹ)
    '   |-cllPati(Collect):������Ϣ����Ա��
    '     (����ID,��ҳID,����,�Ա�,����,�ѱ�,����,���˿���ID,����ID,����,
    '      ��������,������Դ,��ʶ��(����Ż�סԺ��),�Һŵ�ID(N),����������(N),���������֤��(N),
    '      ��������(N),�������ص�λ(N),��ϼ�¼ID(N),���ID(N),�������(N),
    '      ���,��������,���֤��,ҽ�Ƹ��ʽ����,ҽ�Ƹ��ʽ����,������˱�־(N),סԺ״̬(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):������Ϣ��=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):������Ϣ����Ա��
    '       (���ݺ�,�Ƿ�ಡ�˵�,ҽ��С��id,ҽ�����ٴ�����,�򵥼���,���ʵ�id,��¼����,�Ƿ񻮼�,�Ƿ���,�Ӱ��־,��������ID,
    '        ��ҩ����ID,������,������,����Ա����,����Ա���,����ʱ��,�Ǽ�ʱ��,���˿���ID,[cllBillDetails(collect)])=cllBillists(_���ݺ�)
    '       |-cllBillDetails(Collect):������ϸ��=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):ÿ����ϸ���ݼ�����Ա��
    '           (���,��������,ҩ��ID,�շ�ϸĿID,�۸񸸺�,������ĿID,�ѱ�,Ӥ�����,�շ����,���㵥λ,�Ƿ�����Ŀ,���մ���ID,
    '            ���ձ���,�վݷ�Ŀ,����,����,����,Ӧ�ս��,ʵ�ս��,ͳ����,���ӱ�־,ִ�в���ID,�Ƿ��Զ�����,����ժҪ(N),
    '            ҽ��ID(N),����ID(N),��������,��ҩ��̬,�巨,ִ������,�Ƿ񱸻�����(N),������������(N),�Ƿ��������,��ҩ����(N),
    '            �������(N),������ĿID(N),��ҩ;��ID(N),��ҩ;������(N),��ҩ;������(N),��ҩƵ��ID(N)����ҩƵ�����ƣ�N),
    '            ҽ��������־(N),ҽ����Ч(N),�Ƽ�����(N),Ƶ��(N),������N),�÷�(N),Ƥ�Խ��(N),����˵��(N),ʹ������(N),
    '            ��ҩ��ʽ(N),ҩƷ����(N),����ִ������(N),�巨(N)
    '            �����ʱ����ӡ�:(����ID,��ҳID,����,�Ա�,����,����,���˿���ID,����ID,����,
    '              ��������,��ʶ��(����Ż�סԺ��),�Һŵ�ID(N),����������(N),���������֤��(N),
    '              ������Դ,��������(N),�������ص�λ(N),��ϼ�¼ID(N),���ID(N),�������(N),
    '              ���,��������,���֤��,ҽ�Ƹ��ʽ����,ҽ�Ƹ��ʽ����,������˱�־(N),סԺ״̬(N),
    '              ҽ��С��id)=cllBillDetails(_���)
    '    ����Ԫ���У���(N)�ģ���ʾ��ѡ�ڵ�
    '  intBillType-��������(1-�շѵ�;2-���ʵ�;3-���ʱ�)
    '  bln����-�Ƿ񻮼�
    '����:blnSendMateria:���ʺ��Զ���ҩ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-13 22:52:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHave As Boolean, cllDrawDeptFeeIds As Collection
    Dim objExpenceSvr As clsExpenceSvr
    
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
    
    blnHave = False
    If Not cllRcpBillData Is Nothing Then blnHave = cllRcpBillData.Count <> 0
    If Not blnHave Then
         MsgBox "��������Ҫ����ļ������ݣ�����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '���²�����ҳ��Ϣ��������˱�־��סԺ״̬
    If UpdatePatiPageInfo(cllRcpBillData, intBillType, lngModule) = False Then Exit Function
    
    '1.�Ƚ������ݺϷ��Լ��
    If objExpenceSvr.zlExcute_SaveRecipeBill_Check(cllRcpBillData, intBillType, _
        cllDrawDeptFeeIds, bln����, bln����, lngModule) = False Then Exit Function
    
    '2.���ݱ���
    If objExpenceSvr.zlExcute_SaveRecipeBill(cllRcpBillData, strFrmCaption, intBillType, bln����, _
        cllDrawDeptFeeIds, bln����, , blnSendMateria, lngModule) = False Then Exit Function
    
    zlSaveBill_NewRecipeBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiPageInfo(ByVal byt��ѯ���� As Byte, _
    ByVal str������Ϣ As String, ByRef rsPatiPageInfo As ADODB.Recordset, _
    Optional ByVal bln����Ӥ����Ϣ As Boolean, _
    Optional ByRef rsBabyInfo As ADODB.Recordset, Optional ByVal lngModule As Long, _
    Optional ByVal blnȡ�����ҳID As Boolean = True, _
    Optional ByVal str�������� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ѯ������ҳ��Ϣ
    '���:
    '   byt��ѯ���� ��ѯ����:0-������Ϣ;1-������Ϣ����չ;2-��ȡ��ҳ
    '   str������Ϣ ������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
    '   bln����Ӥ����Ϣ �Ƿ����Ӥ����Ϣ
    '   blnȡ�����ҳID �Ƿ�ȡ���һ��סԺ
    '   str��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�������ŷָ�������Ϊ����
    '����:
    '   rsPatiPageInfo ���˲�����ҳ��Ϣ������id,��ҳid,����,�Ա�,����,�ѱ�,��������,��˱�־,סԺ״̬,��Ժʱ��,��Ժʱ��,סԺҽʦ,
    '                                   ҽ�Ƹ��ʽ����,ҽ�Ƹ��ʽ����,��ǰ����id,��ǰ��������,��ǰ����id,��ǰ��������,
    '                                   ����,ǰ����,��������,[ѧ��,ְҵ,����,����״��,��Ŀ����,���˱�ע]
    '   rsBabyInfo Ӥ����Ϣ������id,��ҳid,���,����,�Ա�,����ʱ��
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim cllBabyInfo As Collection
    
    On Error GoTo ErrHandler
    If str������Ϣ = "" Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    zlGetPatiPageInfo = objService.zlCIsSvr_GetPatiPageInfo(byt��ѯ����, str������Ϣ, _
        rsPatiPageInfo, bln����Ӥ����Ϣ, cllBabyInfo, lngModule, blnȡ�����ҳID, str��������)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function UpdatePatiPageInfo(cllRcpBillData As Collection, ByVal intBillType As Integer, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���²�����ҳ��Ϣ��������˱�־��סԺ״̬
    '���:
    ' cllRcpBillData(�ṹ)
    '   |-cllPati(Collect):������Ϣ����Ա��
    '     (����ID,��ҳID,����,�Ա�,����,�ѱ�,����,���˿���ID,����ID,����,
    '      ��������,������Դ,��ʶ��(����Ż�סԺ��),�Һŵ�ID(N),����������(N),���������֤��(N),
    '      ��������(N),�������ص�λ(N),��ϼ�¼ID(N),���ID(N),�������(N),
    '      ���,��������,���֤��,ҽ�Ƹ��ʽ����,ҽ�Ƹ��ʽ����,������˱�־(N),סԺ״̬(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):������Ϣ��=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):������Ϣ����Ա��
    '       (���ݺ�,�Ƿ�ಡ�˵�,ҽ��С��id,ҽ�����ٴ�����,�򵥼���,���ʵ�id,��¼����,�Ƿ񻮼�,�Ƿ���,�Ӱ��־,��������ID,
    '        ��ҩ����ID,������,������,����Ա����,����Ա���,����ʱ��,�Ǽ�ʱ��,���˿���ID,[cllBillDetails(collect)])=cllBillists(_���ݺ�)
    '       |-cllBillDetails(Collect):������ϸ��=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):ÿ����ϸ���ݼ�����Ա��
    '           (���,��������,ҩ��ID,�շ�ϸĿID,�۸񸸺�,������ĿID,�ѱ�,Ӥ�����,�շ����,���㵥λ,�Ƿ�����Ŀ,���մ���ID,
    '            ���ձ���,�վݷ�Ŀ,����,����,����,Ӧ�ս��,ʵ�ս��,ͳ����,���ӱ�־,ִ�в���ID,�Ƿ��Զ�����,����ժҪ(N),
    '            ҽ��ID(N),����ID(N),��������,��ҩ��̬,�巨,ִ������,�Ƿ񱸻�����(N),������������(N),�Ƿ��������,��ҩ����(N),
    '            �������(N),������ĿID(N),��ҩ;��ID(N),��ҩ;������(N),��ҩ;������(N),��ҩƵ��ID(N)����ҩƵ�����ƣ�N),
    '            ҽ��������־(N),ҽ����Ч(N),�Ƽ�����(N),Ƶ��(N),������N),�÷�(N),Ƥ�Խ��(N),����˵��(N),ʹ������(N),
    '            ��ҩ��ʽ(N),ҩƷ����(N),����ִ������(N),�巨(N)
    '            �����ʱ����ӡ�:(����ID,��ҳID,����,�Ա�,����,����,���˿���ID,����ID,����,
    '              ��������,��ʶ��(����Ż�סԺ��),�Һŵ�ID(N),����������(N),���������֤��(N),
    '              ������Դ,��������(N),�������ص�λ(N),��ϼ�¼ID(N),���ID(N),�������(N),
    '              ���,��������,���֤��,ҽ�Ƹ��ʽ����,ҽ�Ƹ��ʽ����,������˱�־(N),סԺ״̬(N),
    '              ҽ��С��id)=cllBillDetails(_���)
    '    ����Ԫ���У���(N)�ģ���ʾ��ѡ�ڵ�
    '  intBillType-��������(1-�շѵ�;2-���ʵ�;3-���ʱ�)
    '����:
    '����:���³ɹ�����True,ʧ�ܷ���False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiInfo As Collection, cllBillLists As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim rsPatiPageInfo As ADODB.Recordset
    Dim p As Long, i As Long, str����IDs As String, str������Ϣ As String

    On Error GoTo ErrHandler
    '�ռ�������Ϣ
    str������Ϣ = ""
    If intBillType = 3 Then '���ʱ�
        str����IDs = ""
        Set cllBillLists = cllRcpBillData("_cllBillLists")
        For p = 1 To cllBillLists.Count
            Set cllBillDetails = cllBillLists(p)("_cllBillDetails")
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                If InStr("," & str����IDs & ",", "," & cllBillDetail("����ID") & ",") = 0 Then
                    str������Ϣ = str������Ϣ & "," & cllBillDetail("����ID") & ":" & cllBillDetail("��ҳID")
                    str����IDs = str����IDs & "," & cllBillDetail("����ID")
                End If
            Next
        Next
    Else
        Set cllPatiInfo = cllRcpBillData("_patiinfor")
        str������Ϣ = str������Ϣ & "," & cllPatiInfo("����ID") & ":" & cllPatiInfo("��ҳID")
    End If
    If str������Ϣ <> "" Then str������Ϣ = Mid(str������Ϣ, 2)

    '��ȡ������ҳ��Ϣ
    If zlGetPatiPageInfo(0, str������Ϣ, rsPatiPageInfo, False, , lngModule) = False Then Exit Function

    '���²�����ҳ��Ϣ
    If intBillType = 3 Then '���ʱ�
        Set cllBillLists = cllRcpBillData("_cllBillLists")
        For p = 1 To cllBillLists.Count
            Set cllBillDetails = cllBillLists(p)("_cllBillDetails")
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                rsPatiPageInfo.Filter = "����ID=" & cllBillDetail("����ID")
                If rsPatiPageInfo.EOF Then
                    MsgBox "��ȡ���ˡ�" & cllBillDetail("����") & "���Ĳ�����Ϣʧ�ܣ�", vbInformation, gstrSysName
                    Exit Function
                End If
                If CollectionExitsValue(cllBillDetail, "������˱�־") Then cllBillDetail.Remove "������˱�־"
                cllBillDetail.Add Nvl(rsPatiPageInfo!��˱�־), "������˱�־"
                If CollectionExitsValue(cllBillDetail, "סԺ״̬") Then cllBillDetail.Remove "סԺ״̬"
                cllBillDetail.Add Nvl(rsPatiPageInfo!סԺ״̬), "סԺ״̬"
            Next
        Next
    Else
        Set cllPatiInfo = cllRcpBillData("_patiinfor")
        If rsPatiPageInfo.EOF Then
            MsgBox "��ȡ���ˡ�" & cllPatiInfo("����") & "���Ĳ�����Ϣʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        If CollectionExitsValue(cllPatiInfo, "������˱�־") Then cllPatiInfo.Remove "������˱�־"
        cllPatiInfo.Add Nvl(rsPatiPageInfo!��˱�־), "������˱�־"
        If CollectionExitsValue(cllPatiInfo, "סԺ״̬") Then cllPatiInfo.Remove "סԺ״̬"
        cllPatiInfo.Add Nvl(rsPatiPageInfo!סԺ״̬), "סԺ״̬"
    End If

    UpdatePatiPageInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Function zlExcute_DelRecipeBill(ByVal cllBillLists As Collection, _
    ByVal strFrmCaption As String, ByVal intBillType As Integer, _
    Optional ByVal bln���� As Boolean, Optional ByVal lngModule As Long, Optional bln���� As Boolean, _
    Optional ByVal cllPro As Collection, Optional ByVal cllPatients As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�д����������ϲ���
    '���:
    '     cllBillList(Collect):������Ϣ����Ա��
    '        (���ݺ�,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         ���ʲ���(����״̬(N),����Ա���,����Ա����,�Ǽ�ʱ��,�ѽ��ֹ����,��ֹ��������)
    '         �˷Ѳ���(����Ա���,����Ա����,�Ǽ�ʱ��,����ID,ժҪ,��������))=cllBillLists(_���ݺ�)
    '       |-cllBillDetails(Collect):������Ϣ��=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):ÿ����ϸ���ݼ�����Ա��
    '           (���,��������,��ҩIDs(N))=cllBillDetails(_���)
    '       |-cllAdviceUpdateDatas(collect):ҽ���������ݣ���ִ�м������=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)ÿ����ϸ���ݼ�����Ա��
    '           (ҽ��ID,���ͺ�(N),�Ʒ�״̬,ɾ������(N))=cllAdviceUpdateDatas(i)
    '    ����Ԫ���У���(N)�ģ���ʾ��ѡ�ڵ㡣
    '   intBillType-��������(1-�շѵ�;2-���ʵ�;3-�Զ����ʵ�)
    '   cllPro ��Ҫһ��ִ�е�SQL���
    '   cllPatients(Collect):-������Ϣ������סԺ���ʴ���
    '     |-cllPatient(Collect):-ÿ��������Ϣ����Ա��
    '       (����ID,��ҳID,����,��˱�־,סԺ״̬,��Ŀ����)=cllPatients(_����ID)
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-13 22:52:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHave As Boolean
    Dim objExpenceSvr As clsExpenceSvr
    
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
    
    blnHave = False
    If Not cllBillLists Is Nothing Then blnHave = cllBillLists.Count <> 0
    If Not blnHave Then
         MsgBox "��������Ҫ���ʵ����ݣ�����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '1.�Ƚ����������ݵ���Ч�Խ��м��
    If objExpenceSvr.zlExcute_DelRecipeBill_Check(cllBillLists, intBillType, lngModule, _
        bln����, bln����, cllPatients) = False Then Exit Function
    
    '2.�ٽ������ʴ���
    If objExpenceSvr.zlExcute_DelRecipeBill(cllBillLists, strFrmCaption, _
        intBillType, bln����, lngModule, bln����, cllPro, cllPatients) = False Then Exit Function
    zlExcute_DelRecipeBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl���˷�������_Insert(ByVal cllApplyDatas As Collection, ByRef str�������ids_Out As String, _
    Optional ByVal strFrmCaption As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˷�����������
    '���:
    '   cllApplyDatas-�����������ݼ�(RowData=collect:(����ID,���ݺ�,�շ�ϸĿID,�������ID,��˿���ID,��������,������,����ʱ��,�������,�շ����,����ԭ��))
    '����:
    '   str�������ids_Out-�����������漰�ķ���ID(��Ҫ�Ǻ�����Ҫ��)
    '����:���뷵��true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    zl���˷�������_Insert = objPubExseSvr.zl���˷�������_Insert(cllApplyDatas, str�������ids_Out, strFrmCaption, lngModule)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_���˷�������_Audit_Check(ByVal cllAuditDatas As Collection, _
    Optional ByVal strFrmCaption As String, Optional lngModule As Long, _
    Optional ByRef rsRecipe_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˷�����������
    '���:
    '   cllAuditDatas-����������ݼ�(RowData=collect:(����ID,���ݺ�,�շ����,�����Ƿ��Զ�����,����ʱ��,�������))
    '          ��Աֵ˵��:1.���������Ƿ��Զ�����:1-�Զ�����;0-���Զ�����
    '                     2.�������:0-δ��ҩ(��);1-�ѷ�ҩ(��);����Ϊ0
    '����:
    '   rsRecipe_Out-ҩƷ���������ϵ���ִ�����������ݼ�
    '����:��˳ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    Zl_���˷�������_Audit_Check = objPubExseSvr.Zl_���˷�������_Audit_Check(cllAuditDatas, strFrmCaption, lngModule, rsRecipe_Out)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_���˷�������_Cancel_Check(ByVal cllAuditDatas As Collection, _
    Optional ByVal strFrmCaption As String, Optional lngModule As Long, _
    Optional ByRef rsSendDatas As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ԡ�ȡ���ܷ���������ܷ����Ĺ��ܽ��кϷ��Լ��
    '���:
    '   cllAuditDatas-����������ݼ�(RowData=collect:(����״̬,����ID,���ݺ�(N),�շ����,�����Ƿ��Զ�����(N),����ʱ��,�������(N),����Ա����))
    '          ��Աֵ˵��:0.����״̬:0-��˾ܾ�������;1-ȡ���ܾ������루����״̬_In=1ʱ��Ч))
    '                     1.���������Ƿ��Զ�����:1-�Զ�����;0-���Զ����ϣ�����״̬_In=1ʱ��Ч))
    '                     2.�������:0-δ��ҩ(��);1-�ѷ�ҩ(��);����Ϊ0������״̬_In=1ʱ��Ч))
    '                     3.N�����ѡ
    '   rsSendDatas-not Nothing �����Ѿ����ⲿ��ȡ���ѷ�ҩ���ϵ����ݣ��ڲ�ֱ�ӻ��ã������ٵ��÷���
    '����:
    '   rsSendDatas-ҩƷ���������ϵ���ִ�����������ݼ�
    '����:�Ϸ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    Zl_���˷�������_Cancel_Check = objPubExseSvr.Zl_���˷�������_Cancel_Check(cllAuditDatas, strFrmCaption, lngModule, rsSendDatas)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDeptName(ByVal lngDeptID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '���:lngDeptID-����ID(����ID)
    '����:����ȡ��������
    '����:���˺�
    '����:2015-07-15 17:52:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select ���� From ���ű� where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lngDeptID)
    If rsTemp.EOF = False Then zlGetDeptName = Nvl(rsTemp!����)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsExistInfusion_FromNo(ByVal strNO As String, str���s As String, _
    Optional blnIsExist_Out As Boolean, Optional ByVal int��¼���� As Integer = 2, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҩƷ�Ƿ��������Һ��ҩ����
    '���:strNo-���ݺ�
    '     str���s-����Ϊ��,Ϊ��ʱ����ʾ��������
    '����:blnIsExist_Out-�Ƿ���ڣ����ڷ���true,���򷵻�False
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-22 16:30:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str����IDs As String
    Dim strWhere As String
    
    On Error GoTo errHandle
    blnIsExist_Out = False
    If str���s <> "" Then strWhere = strWhere & " And instr([3],','||���||',')>0 "
    strSQL = "Select ID" & _
            " From סԺ���ü�¼" & _
            " Where NO=[1] and ��¼���� =[2]  And ��¼״̬ in (0,1,3)" & _
            "       And �շ���� in ('5','6','7') " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ID", strNO, int��¼����, "," & str���s & ",")
    If rsTemp.EOF Then zlIsExistInfusion_FromNo = True: Exit Function
    
    Do While Not rsTemp.EOF
        str����IDs = str����IDs & "," & rsTemp!ID
        rsTemp.MoveNext
    Loop
    str����IDs = Mid(str����IDs, 2)
    If mdlDrugAndStuffSvr.zlPivasSvr_Isexsitinfusion(str����IDs, blnIsExist_Out, lngModule) = False Then Exit Function
    
    zlIsExistInfusion_FromNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlFeeExecute_Check(ByVal strNO As String, ByVal str��� As String, ByVal int������Դ As Integer, ByVal int�������� As Integer, _
    Optional ByRef strSendStuffFeeIDs_Out As String, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    '����: ִ�еǼǺϷ��Լ��
    '���:
    '   strNO-���ݺ�
    '   str���-���
    '   int������Դ-1-����;2-סԺ
    '   int��������-1-�շ�;2-����;3-�Զ�����
    '����:
    '   strSendStuffFeeIDs_Out-����ִ�����漰�Զ�������������Ӧ�ķ���Ids,����ö��ŷ���
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    
    'ִ�еǼ�ǰ���
    If objPubExseSvr.zlFeeExecute_Check(strNO, str���, int������Դ, int��������, _
        strSendStuffFeeIDs_Out, strFrmCaption, lngModule, 1) = False Then Exit Function
    zlFeeExecute_Check = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlFeeUnExecute_Check(ByVal strNO As String, ByVal str��� As String, ByVal int������Դ As Integer, _
    ByVal int�������� As Integer, ByRef strStuffFeeIDs_Out As String, _
    ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    '����: ȡ��ִ�еǼ�ǰ���
    '���:
    '   strNO-���ݺ�
    '   str���-���
    '   int������Դ-1-����;2-סԺ
    '   int��������-1-�շ�;2-����;3-�Զ�����
    '����:
    '   strStuffFeeIDs_Out-����ִ�����漰��������Ӧ�ķ���Ids,����ö��ŷ���
    '����:
    '---------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    
    'ִ�еǼ�ǰ���
    If objPubExseSvr.zlFeeUnExecute_Check(strNO, str���, int������Դ, int��������, strStuffFeeIDs_Out, _
         strFrmCaption, lngModule) = False Then Exit Function
    zlFeeUnExecute_Check = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetMaxBedLen(Optional ByVal lng����ID As Long, Optional ByVal bln���� As Boolean) As Integer
'���ܣ���ȡָ�����ŵĴ�λ�ŵ���󳤶�
'������lng����ID=����ID�����ID,Ϊ0��ʾ���в��������
    Dim objService As zlPublicExpense.clsService
    Dim lngBedNoMaxLen As Long, bln��������ѯ As Boolean
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    bln��������ѯ = (Not bln���� Or lng����ID = 0)
    If objService.ZlCissvr_GetMaxBedLen(lngBedNoMaxLen, bln��������ѯ, lng����ID, lng����ID) = False Then Exit Function
    
    zlGetMaxBedLen = lngBedNoMaxLen
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheckҽ���´��Ժҽ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ���Ƿ��´��˳�Ժҽ��
    '���:
    '����:
    '   blnExistOutAdvice=�Ƿ���ڳ�Ժҽ��
    '   lngOutAdviceId=�Ѿ����˳�Ժҽ���ģ�����ҽ����ID
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim blnExistOutAdvice  As Boolean
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then
        zlCheckҽ���´��Ժҽ�� = True: Exit Function '���ز���
    End If
    
    If objService.ZlCissvr_ExistOutAdvice(lng����ID, lng��ҳID, blnExistOutAdvice) = False Then
        zlCheckҽ���´��Ժҽ�� = True: Exit Function '���ز���
    End If
    
    zlCheckҽ���´��Ժҽ�� = blnExistOutAdvice
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetPatiInfoByPage(objPati As clsPatientInfo, _
    Optional ByVal lng��ҳID As Long, Optional ByVal bln����Ӥ����Ϣ As Boolean, _
    Optional ByRef rsBabyInfo As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ӳ�����ҳ�л�ȡ������Ϣ
    '���:
    '   objPati-���в�����Ϣ
    '   lng��ҳID-��ҳID��Ϊ0ʱ��ȡ���һ��סԺ��
    '   bln����Ӥ����Ϣ �Ƿ����Ӥ����Ϣ
    '����:
    '   objPati-���ز�����Ϣ����
    '   rsBabyInfo Ӥ����Ϣ������id,��ҳid,���,����,�Ա�,����ʱ��
    '����:�ɹ�����True�����򷵻�False
    '˵��:������� objPati ��ΪNothing���������Ϣ�ϲ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objService As zlPublicExpense.clsService
    Dim cllBabyInfo As Collection
    
    On Error GoTo errHandle
    If objPati Is Nothing Then Exit Function
    If objPati.����ID = 0 Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    If lng��ҳID = 0 Then lng��ҳID = objPati.��ҳID '������Ϣ�е���ҳIDΪ���һ����ҳID
    If objService.zlCIsSvr_GetPatiPageInfo(1, objPati.����ID & ":" & lng��ҳID, _
        rsTemp, bln����Ӥ����Ϣ, cllBabyInfo, lngModule) = False Then Exit Function
    If rsTemp Is Nothing Then zlGetPatiInfoByPage = True: Exit Function
    If rsTemp.EOF Then zlGetPatiInfoByPage = True: Exit Function
    
    If objPati Is Nothing Then Set objPati = New clsPatientInfo
    If bln����Ӥ����Ϣ And Not cllBabyInfo Is Nothing Then
        If cllBabyInfo.Count > 0 Then Set rsBabyInfo = cllBabyInfo("_" & objPati.����ID)
    End If
    
    With objPati
        .��ҳID = Nvl(rsTemp!��ҳID)
        .���� = Nvl(rsTemp!����)
        .�Ա� = Nvl(rsTemp!�Ա�)
        .���� = Nvl(rsTemp!����)
        .�ѱ� = Nvl(rsTemp!�ѱ�)
        .ҽ�Ƹ��ʽ = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
        .ҽ�Ƹ��ʽ���� = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
        .���� = Val(Nvl(rsTemp!����))
        .�������� = GetInsureName(Val(Nvl(rsTemp!����)))
        .�������� = Nvl(rsTemp!��������)
        .��ǰ����ID = Val(Nvl(rsTemp!��ǰ����ID))
        .��ǰ�������� = Nvl(rsTemp!��ǰ��������)
        .��ǰ����ID = Val(Nvl(rsTemp!��ǰ����ID))
        .��ǰ�������� = Nvl(rsTemp!��ǰ��������)
        .���� = Nvl(rsTemp!��ǰ����)
        .סԺ�� = Nvl(rsTemp!סԺ��)
        .�������� = Val(Nvl(rsTemp!��������))
        .��Ժ���� = Nvl(rsTemp!��Ժʱ��)
        .��Ժ���� = Nvl(rsTemp!��Ժʱ��)
        .סԺҽʦ = Nvl(rsTemp!סԺҽʦ)
        .���˱�ע = Nvl(rsTemp!���˱�ע)
        .סԺ״̬ = Val(Nvl(rsTemp!סԺ״̬))
        .��˱�־ = Val(Nvl(rsTemp!��˱�־))
        .��Ŀ���� = Nvl(rsTemp!��Ŀ����)
        .ҽ���� = Nvl(rsTemp!ҽ����)
    End With
    zlGetPatiInfoByPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiInfo(ByVal lng����ID As Long, _
    Optional ByVal lng��ҳID As Long, Optional ByVal lngModule As Long, _
    Optional ByVal bln����Ӥ����Ϣ As Boolean, _
    Optional ByRef rsBabyInfo As ADODB.Recordset, _
    Optional ByVal blnNotShowErrMsg As Boolean) As clsPatientInfo
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ���ȴӲ�����Ϣ�л�ȡ���ٴӲ�����ҳ�л�ȡ���кϲ�
    '���:
    '   objPati-���в�����Ϣ
    '   lng��ҳID-��ҳID��Ϊ0ʱ��ȡ���һ��סԺ�ģ�Ϊ-1ʱȡ���ﲡ����Ϣ
    '   bln����Ӥ����Ϣ �Ƿ����Ӥ����Ϣ
    '����:
    '   objPati-���ز�����Ϣ����
    '   rsBabyInfo Ӥ����Ϣ������id,��ҳid,���,����,�Ա�,����ʱ��
    '����:�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPati As clsPatientInfo
    Dim lngTmp As Long
    
    On Error GoTo errHandle
    '��ȡ������Ϣ
    If gobjSquare.objOneCardComLib.zlGetPatiInforFromPatiID(lng����ID, objPati, , , , , , , , , , , blnNotShowErrMsg) = False Then Exit Function
    If objPati Is Nothing Then Exit Function
    lngTmp = objPati.��ҳID
    '2.��ȡ������ҳ
    If lng��ҳID = 0 Then lng��ҳID = objPati.��ҳID
    If lng��ҳID > 0 Then
        If zlGetPatiInfoByPage(objPati, lng��ҳID, bln����Ӥ����Ϣ, rsBabyInfo, lngModule) = False Then Exit Function
    End If
    If lngTmp <> objPati.��ҳID Then objPati.��Ժ = False
    Set zlGetPatiInfo = objPati
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetMultiPatiInfo(ByVal str������Ϣ As String, _
    Optional ByRef cllBabyInfo As Collection, Optional ByVal lngModule As Long) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���������Ϣ���ȴӲ�����Ϣ�л�ȡ���ٴӲ�����ҳ�л�ȡ���кϲ�
    '���:
    '   str������Ϣ=����ID����ҳ��Ϣ����ʽ������ID1:��ҳID,����ID2:��ҳID,...������,��ҳID=0ʱ��ȡ���һ��סԺ�ģ�Ϊ-1ʱȡ���ﲡ����Ϣ
    '����:
    '   cllBabyInfo=Ӥ����Ϣ,��Ա:ADODB.Recordset=cllBabyInfo(_����ID)����Ա�ֶΣ�����id,��ҳid,���,����,�Ա�,����ʱ��
    '����:������Ϣ������Ա��clsPatinetInfo=cllPatis(_����ID)
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, varData As Variant, i As Long
    Dim str����IDs As String, lng����ID As Long, lng��ҳID As Long
    Dim rsTemp As ADODB.Recordset, objPati As clsPatientInfo
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    varData = Split(str������Ϣ, ",")
    For i = 0 To UBound(varData)
        lng����ID = Split(varData(i) & ":", ":")(0)
        str����IDs = str����IDs & "," & lng����ID
    Next
    If str����IDs = "" Then Exit Function
    
    str����IDs = Mid(str����IDs, 2)
    '��ȡ������Ϣ
    If gobjSquare.objOneCardComLib.zlGetMultiPatiInforFromPatiID(str����IDs, cllPati) = False Then Exit Function
    If cllPati Is Nothing Then Exit Function
    If cllPati.Count = 0 Then Exit Function
    
    Set zlGetMultiPatiInfo = cllPati
    '2.��ȡ������ҳ
    str������Ϣ = ""
    For i = 0 To UBound(varData)
        lng����ID = Split(varData(i) & ":", ":")(0)
        lng��ҳID = Val(Split(varData(i) & ":", ":")(1))
        If lng��ҳID = 0 Then '��ҳID=0ʱ��ȡ���һ��סԺ��
            Set objPati = cllPati("_" & lng����ID)
            lng��ҳID = objPati.��ҳID
        End If
        If lng��ҳID > 0 Then '��ҳID=-1ʱ����ȡ���ﲡ����Ϣ
            str������Ϣ = str������Ϣ & "," & lng����ID & ":" & lng��ҳID
        End If
    Next
    If str������Ϣ = "" Then Exit Function
    
    str������Ϣ = Mid(str������Ϣ, 2)
    If objService.zlCIsSvr_GetPatiPageInfo(1, str������Ϣ, rsTemp, True, cllBabyInfo, lngModule) = False Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.EOF Then Exit Function
    
    For i = 0 To UBound(varData)
        lng����ID = Split(varData(i) & ":", ":")(0)
        rsTemp.Filter = "����ID=" & lng����ID
        If Not rsTemp.EOF Then
            Set objPati = cllPati("_" & lng����ID)
            With objPati
                .��ҳID = Nvl(rsTemp!��ҳID)
                .���� = Nvl(rsTemp!����)
                .�Ա� = Nvl(rsTemp!�Ա�)
                .���� = Nvl(rsTemp!����)
                .�ѱ� = Nvl(rsTemp!�ѱ�)
                .ҽ�Ƹ��ʽ = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
                .ҽ�Ƹ��ʽ���� = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
                .���� = Val(Nvl(rsTemp!����))
                .�������� = GetInsureName(Val(Nvl(rsTemp!����)))
                .�������� = Nvl(rsTemp!��������)
                .��ǰ����ID = Val(Nvl(rsTemp!��ǰ����ID))
                .��ǰ�������� = Nvl(rsTemp!��ǰ��������)
                .��ǰ����ID = Val(Nvl(rsTemp!��ǰ����ID))
                .��ǰ�������� = Nvl(rsTemp!��ǰ��������)
                .���� = Nvl(rsTemp!��ǰ����)
                .סԺ�� = Nvl(rsTemp!סԺ��)
                .�������� = Val(Nvl(rsTemp!��������))
                .��Ժ���� = Nvl(rsTemp!��Ժʱ��)
                .��Ժ���� = Nvl(rsTemp!��Ժʱ��)
                .סԺҽʦ = Nvl(rsTemp!סԺҽʦ)
                .���˱�ע = Nvl(rsTemp!���˱�ע)
                .סԺ״̬ = Val(Nvl(rsTemp!סԺ״̬))
                .��˱�־ = Val(Nvl(rsTemp!��˱�־))
                .��Ŀ���� = Nvl(rsTemp!��Ŀ����)
                .ҽ���� = Nvl(rsTemp!ҽ����)
            End With
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlGetPatiPageInfByRange(ByVal cllFilter As Collection, _
    ByRef rsPatiPageInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long, Optional ByVal bln����Ӥ����Ϣ As Boolean, _
    Optional ByRef cllBabyInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ѯ������ҳ��Ϣ
    '���:
    '   cllFilter ��ѯ������:��Ա(Array(Key,Value),Array(Key,Value),,...)
    '       Key:����IDS,����IDS,����IDS,��ҳIDS,��Ժ��ʼʱ��,��Ժ����ʱ��,��Ժ��ʼʱ��,��Ժ����ʱ��,
    '           �ѱ�,סԺ״̬,��������,����,վ����,��ѯת�Ʋ���,���һ��סԺ,����,����վ����
    '       סԺ״̬:0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ
    '       �������ʣ�����ö��ŷ�0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�NULL-��ʾ������
    '       ����:���Դ�%�ֺű������ƥ��
    '       �ѳ�Ժ������סԺ״̬Ϊ1��2ʱ��Ч
    '       վ����:���Ҷ�Ӧ��վ����
    '       ����:>0:ָ������ҽ������,0:ҽ������ͨ����,-1:��ͨ����,-2:ҽ������
    '   bln����Ӥ����Ϣ �Ƿ����Ӥ����Ϣ
    '����:
    '   rsPatiPageInfo ���˲�����ҳ��Ϣ������ID,��ҳID,����,�Ա�,����,סԺ��,����,����,�ѱ�,��������,ҽ����,
    '                                   ��Ժʱ��,��Ժʱ��,סԺ״̬,��������,��ǰ����ID,��ǰ��������,��ǰ����ID,��ǰ��������,
    '                                   ҽ�Ƹ��ʽ����,ҽ�Ƹ��ʽ����,סԺҽʦ,���˱�ע,��Ŀ����,����ȼ�,
    '                                   ����ת��,��˱�־,�����,Ԥ��Ժʱ��,�ϴδ߿���
    '       סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)
    '       ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    '       ����ת��:0-δת����1-��ת��
    '       ��˱�־:0���-δ���,1-����˻�ʼ���;2-������
    '   cllBabyInfo Ӥ����Ϣ,��Ա��ADODB.Recordset=cllBabyInfo(_����ID_��ҳID)���ֶΣ�����id,��ҳid,���,����,�Ա�,����ʱ��
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    ZlGetPatiPageInfByRange = objService.ZlCissvr_GetPatiPageInfByRange(cllFilter, rsPatiPageInfo, lngModule, _
        bln����Ӥ����Ϣ, cllBabyInfo)
End Function

Public Function ZlGetBabyData(ByVal lng����ID As Long, _
    ByVal lng��ҳID As Long, ByRef rsBabyInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ѯ������ҳ��Ϣ
    '���:
    '����:
    '   cllBabyInfo Ӥ����Ϣ,�ֶΣ����,����
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If lng����ID = 0 Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    ZlGetBabyData = objService.ZlCissvr_GetBabyData(lng����ID, lng��ҳID, rsBabyInfo, lngModule)
End Function


Public Function ZLGetAdviceIDs(ByVal strҽ��ID As String) As String
    '��ȡһ��ҽ��������ҽ����¼ID��
    '���:
    '   strҽ��IDs ҽ��ID,���Ӣ�Ķ��ŷָ�
    Dim objService As zlPublicExpense.clsService
    Dim rsAdviceData As ADODB.Recordset, strҽ��IDs As String
    
    If strҽ��ID = "" Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAllGroupAdviceIDs(strҽ��ID, rsAdviceData) = False Then Exit Function
    If rsAdviceData.EOF Then Exit Function
    
    strҽ��IDs = ""
    Do While Not rsAdviceData.EOF
        strҽ��IDs = strҽ��IDs & "," & Nvl(rsAdviceData!ҽ��ID)
        rsAdviceData.MoveNext
    Loop
    ZLGetAdviceIDs = Mid(strҽ��IDs, 2)
End Function

Public Function ZlGetPatiIdFromPatiPage(ByRef lng����ID As Long, Optional ByVal lng����ID As Long, _
    Optional ByVal strסԺ�� As String, Optional ByVal str���� As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݴ��š�סԺ�Ż�ȡ����ID����ҳID
    '���:
    '   strסԺ�š�lng����id��str����-�������ٴ�һ��
    '����:
    '   str������Ϣ:����ID:��ҳID
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    ZlGetPatiIdFromPatiPage = objService.zlCisSvr_GetPatiID(lng����ID, lng����ID, strסԺ��, str����, lngModule)
End Function

Public Function ZlGetInDeptInfor(ByVal bytִ�й��� As Byte, ByVal bln��Ժ���� As Boolean, _
    Optional ByVal byt���ҷ�ʽ As Byte, Optional ByVal bln���в��� As Boolean, Optional ByVal str����IDs As String, _
    Optional ByVal str������� As String, Optional ByVal byt������Դ As Byte = 2, Optional ByVal lngModule As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����в��˵�סԺ����
    '���:
    '   bytִ�й���=0-��ȡ�����в��˵�סԺ���� 1-ͨ������id/����id�������в��˵���Ժ���һ��߲��� 2-����վ��
    '   bln��Ժ����=�Ƿ��ȡ������Ժ���˵Ŀ���\����
    '   byt���ҷ�ʽ=0-�����Ҳ��� 1-����������
    '   bln���в���=�Ƿ����в���
    '   str����ids=�����в���ʱ��Ч
    '   str�������=���ҷ�����󣬶�����ŷָ�,��:1,2,3��1-����,2-סԺ,3-�����סԺ��
    '����:
    '   rsDept ������Ϣ,�ֶΣ�ID,����,����,����
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim rsDept As ADODB.Recordset
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetInDeptInfor(bytִ�й���, byt������Դ, gstrNodeNo, _
        bln��Ժ����, rsDept, byt���ҷ�ʽ, bln���в���, str����IDs, str�������, lngModule) = False Then Exit Function
    Set ZlGetInDeptInfor = rsDept
End Function

Public Function zlCheckPatiIsDeath(ByVal lng����ID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���id��鲡���Ƿ��Ѿ�����
    '���:
    '����:
    '����:����������True,���򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDeath As Boolean
    Dim objService As zlPublicExpense.clsService
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_PatiIsDead(lng����ID, blnDeath, lngModule) = False Then Exit Function
    zlCheckPatiIsDeath = blnDeath
End Function

Public Function ZlGetPatiChangeStopInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str��ֹԭ�� As String, _
    ByRef str��ֹԭ��_Out As String, ByRef str��ֹʱ��_Out As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���id����ҳid�����Ҽ�����id����ȡ���˱䶯����ֹ��Ϣ(��ֹʱ�䡢��ֹԭ��ȣ�
    '���:
    '   str��ֹԭ��=��ֹԭ��:����ö��ŷ���,��:3,15,10,1
    '����:
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetPatiChangeStopInfo(lng����ID, lng��ҳID, _
        lng����ID, lng����ID, str��ֹԭ��, str��ֹԭ��_Out, str��ֹʱ��_Out, lngModule) = False Then Exit Function
    ZlGetPatiChangeStopInfo = True
End Function

Public Function zlCheckPatiIsMemo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���id����ҳid����Ƿ���ڱ�ע��Ϣ
    '���
    '����:
    '   blnIsExist_Out-���ڵģ�����true,���򷵻�False
    '����:�ɹ�����true,���򷵻�False
    '����:34763
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim blnIsExis As Boolean
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlCissvr_PatiExistMemo(lng����ID, lng��ҳID, blnIsExis, lngModule) = False Then Exit Function
    zlCheckPatiIsMemo = blnIsExis
End Function

Public Function zlGetPatiPageExtendInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal str��Ϣ�� As String, Optional ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���id����ҳid��ȡ������ҳ�ӱ���Ϣ
    '���
    '   str��Ϣ��=��Ϣ��������ö���
    '����:
    '����:������Ϣֵ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim cllValue As Collection '��Ϣֵ���ϣ���Ա:Array(��Ϣ��,��Ϣֵ)=cllValue(��Ϣ��)
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlCissvr_GetPatiPageExtendInfo(lng����ID, lng��ҳID, str��Ϣ��, cllValue, lngModule) = False Then Exit Function
    If cllValue Is Nothing Then Exit Function
    If cllValue.Count = 0 Then Exit Function
    zlGetPatiPageExtendInfo = cllValue(1)(str��Ϣ��)
End Function

Public Function ZlGetAdviceInfoByPati(ByVal lng����ID As Long, ByVal ln��ҳID As Long, _
    ByRef rsAdviceData As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ��ID��ȡҽ����Ϣ
    '���:
    '����:
    '   rsAdviceData ҽ����Ϣ��ҽ��ID,ҽ������,�Һŵ���,������ĿID,�������,��������,Ƥ�Խ��,ҽ����Ч,
    '               ���������ݻ�����:���ID,���,Ӥ�����,ҽ��״̬,����ҽ��,ҽ������,����ʱ��,��������ID,
    '                 �������,������־,����,����,����,���㵥λ,ִ��Ƶ��,�÷�,ִ��ʱ�䷽��,
    '                 ��ʼִ��ʱ��,ִ����ֹʱ��,ִ�п���ID,ִ�п�������,ִ������,�ϴ�ִ��ʱ��,ִ�б��,
    '                 У�Ի�ʿ,У��ʱ��,ͣ��ҽ��,ͣ��ʱ��,ͣ����ʿ,ȷ��ͣ��ʱ��,
    '                 �������,�����,���״̬,�Թܱ���,���δ�ӡ,�Ƿ�ǩ��,����ID,����״̬��
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim cllFilter As Collection
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    Set cllFilter = New Collection
    cllFilter.Add Array("����ID", lng����ID)
    cllFilter.Add Array("��ҳID", ln��ҳID)
    If objService.ZlCissvr_GetAdviceInfo(cllFilter, rsAdviceData, 1, lngModule) = False Then Exit Function
    ZlGetAdviceInfoByPati = True
End Function

Public Function ZLGetAdviceSendInfo(ByVal byt��ѯ���� As Byte, ByVal strValues As String, _
    ByRef rsAdviceSendData As ADODB.Recordset, Optional ByVal bln�������ҽ�� As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��������Ϣ
    '���:
    '   byt��ѯ����=0-��ҽ��id����advice_ids����ѯ;1-��ҽ��ID+ҽ�����ͺŲ�ѯ;2-��ҽ��ID+��¼����+NO��ѯ;3-����ҽ�����ͺŲ�ѯ
    '   strValues=��ѯֵ����������:
    '               byt��ѯ����=0:ҽ��ID��,��ʽ��ҽ��ID,ҽ��ID,...
    '               byt��ѯ����=1:������Ϣ,��ʽ��ҽ��ID:NO:��¼����,ҽ��ID:NO:��¼����,...
    '               byt��ѯ����=2:ҽ��������Ϣ,��ʽ:ҽ��ID:���ͺ�,ҽ��ID:���ͺ�,...
    '               byt��ѯ����=3:ҽ�����ͺ�,��ʽ:���ͺ�,���ͺ�,...
    '   bln�������ҽ��=�Ƿ�������ҽ��ID
    '����:
    '   rsAdviceSendData ҽ��������Ϣ��ҽ��ID,���ͺ�,�Һŵ���,����ID,��ҳID,���˿���ID,��������ID,
    '                                 ������Դ,�������,�Ƽ�����,���id,�������,ҽ������,��������,
    '                                 No,��¼����,�״�ʱ��,ĩ��ʱ��,ִ��״̬,����ʱ��,ҽ����Ч
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceSendInfo(byt��ѯ����, strValues, rsAdviceSendData, bln�������ҽ��, lngModule) = False Then Exit Function
    ZLGetAdviceSendInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlStopAutoAccount(ByVal lng����ID As Long, ByVal str��ҳIDS As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ֹͣ�Զ�����
    '���:
    '   str��ҳIDs=��ҳID,������ŷָ�
    '����:
    '����:ִ�гɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_UpddteAutoAccountSign(lng����ID, str��ҳIDS, True, lngModule) = False Then Exit Function
    ZlStopAutoAccount = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlRestoreAutoAccount(ByVal strNO As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ��Զ�����
    '���:
    '   strNO=���ʵ���
    '����:
    '����:ִ�гɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng����ID As Long, str��ҳIDS As String
    
    On Error GoTo ErrHandler
    strSQL = "Select ����id, סԺ����" & _
            " From ���˽��ʼ�¼" & _
            " Where NO = [1] And ��¼״̬ = 3 And �������� = 2 And ��;���� = 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����סԺ����", strNO)
    If rsTemp.EOF Then Exit Function
    
    lng����ID = Val(Nvl(rsTemp!����ID))
    str��ҳIDS = Nvl(rsTemp!סԺ����)
    If str��ҳIDS Then ZlRestoreAutoAccount = True: Exit Function
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_UpddteAutoAccountSign(lng����ID, str��ҳIDS, False, lngModule) = False Then Exit Function
    ZlRestoreAutoAccount = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetMedicalGroupID(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lng��������ID As Long, ByVal str������ As String, _
    ByVal dt����ʱ�� As Date, Optional ByVal lngModule As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ȡ��Ӧ��ҽ��С��ID
    '���:
    '   dt����ʱ��=���÷���ʱ��
    '����:
    '����:��ȡ����ҽ��С��ID
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim lng��id As Long
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetMedicalGroupID(lng����ID, lng��ҳID, _
        lng��������ID, str������, dt����ʱ��, lng��id, lngModule) = False Then Exit Function
        
    ZlGetMedicalGroupID = lng��id
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheckNotExcuteItem(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal bytӤ����� As Byte, ByVal byt������Դ As String, _
    ByRef strNotExcuteInfo As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�����Ϣ��ȡҽ��δִ�е���Ŀ
    '���:
    '   bytӤ�����=Ӥ�����:-1��ʾ������;0-ĸ�׵�;>0����Ӥ������
    '   byt������Դ=1-����;2-סԺ;4-���
    '����:
    '   strNotExcuteInfo=δִ�е���Ŀ��Ϣ
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim lng��id As Long
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_CheckNotExcuteItem(lng����ID, lng��ҳID, -1, byt������Դ, strNotExcuteInfo, lngModule) = False Then Exit Function
        
    zlCheckNotExcuteItem = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckNotExcuteItemValied(ByVal str���� As String, ByVal lng����ID As Long, _
    ByVal lng��ҳID As Long, ByVal int�����־ As Integer, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������δִ����Ŀ�Ƿ�Ϸ�
    '���:
    '   int�����־-1-����;2-סԺ
    '   lngModule -����ģ���
    '����:�Ϸ����ط���true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If int�����־ = 1 Then
        If gTy_System_Para.TY_Balance.byt������δִ�� = 0 Then CheckNotExcuteItemValied = True: Exit Function
    Else
        If gTy_System_Para.TY_Balance.byt���δִ�� = 0 Then CheckNotExcuteItemValied = True: Exit Function
    End If
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_CheckNotExcuteItem(lng����ID, lng��ҳID, -1, 2, strInfo) = False Then Exit Function
    If strInfo = "" Then CheckNotExcuteItemValied = True: Exit Function
        
    If gTy_System_Para.TY_Balance.byt���δִ�� = 1 And int�����־ <> 1 _
        Or gTy_System_Para.TY_Balance.byt������δִ�� = 1 And int�����־ = 1 Then
        If MsgBox("���ֲ���" & str���� & "������δִ����ɵ����ݣ�" & _
            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    Else
        MsgBox "���ֲ���" & str���� & "������δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & _
            vbCrLf & vbCrLf & "������" & IIf(int�����־ <> 2, "����", "��Ժ") & "����.", vbInformation, gstrSysName
        Exit Function
    End If
    CheckNotExcuteItemValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlAuditAdviceCharge(ByVal lngҽ��ID As Long, ByVal bln��� As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ��ҽ�����з���������
    '���:
    '   lngҽ��ID=ҽ��ID
    '   bln���=�Ƿ����:True-���;False-ȡ�����
    '����:
    '����:ִ�гɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_AuditAdviceCharge(lngҽ��ID, bln���, lngModule) = False Then Exit Function
    ZlAuditAdviceCharge = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceDefinedInfo(ByRef rsAdviceDefinedInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�����ݶ���������Ϣ
    '���:
    '����:
    '   rsAdviceDefinedInfo ҽ�����ݶ�����Ϣ���������,ҽ������
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceDefinedInfo(rsAdviceDefinedInfo, lngModule) = False Then Exit Function
    ZlGetAdviceDefinedInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceOperMaxTime(ByVal lngҽ��ID As Long, ByVal byt�������� As Byte, _
    Optional ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ���������һ�ε�ʱ��
    '���:
    '   lngҽ��ID=ҽ��ID
    '   byt��������=��������:1-�¿���2-У�����ʣ�3-У��ͨ����4-���ϣ�5-������
    '                       6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ����
    '                       11-���ͨ����12-���δͨ����13-ʵϰҽʦͣ�������ˣ�14-Ѫ����գ�15-Ѫ�����ͨ����
    '                       16-Ѫ����Ѫ�ܾ���17-Ѫ��ֹͣ��Ѫ��18-��Ѫ����ͨ����ǩ����9-��Ѫ������ˣ�20-��Ѫҽ�����δ��
    '����:
    '����:ҽ�����������һ��ʱ��
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim rsAdviceOper As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceOperInfo(lngҽ��ID, byt��������, rsAdviceOper, True, lngModule) = False Then Exit Function
    If Not rsAdviceOper.EOF Then ZlGetAdviceOperMaxTime = Nvl(rsAdviceOper!����ʱ��)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceOperLastNotes(ByVal lngҽ��ID As Long, ByVal byt�������� As Byte, _
    Optional ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ���������һ�ε�˵��
    '���:
    '   byt��������=��������:1-�¿���2-У�����ʣ�3-У��ͨ����4-���ϣ�5-������
    '                       6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ����
    '                       11-���ͨ����12-���δͨ����13-ʵϰҽʦͣ�������ˣ�14-Ѫ����գ�15-Ѫ�����ͨ����
    '                       16-Ѫ����Ѫ�ܾ���17-Ѫ��ֹͣ��Ѫ��18-��Ѫ����ͨ����ǩ����9-��Ѫ������ˣ�20-��Ѫҽ�����δ��
    '����:
    '����:ҽ���������һ�ε�˵��
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim rsAdviceOper As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If lngҽ��ID = 0 Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceOperInfo(lngҽ��ID, byt��������, rsAdviceOper, True, lngModule) = False Then Exit Function
    If Not rsAdviceOper.EOF Then ZlGetAdviceOperLastNotes = Nvl(rsAdviceOper!����˵��)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetMultiAdviceOperLastNotes(ByVal strҽ��IDs As String, ByVal byt�������� As Byte, _
    ByRef rsAdviceOper As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ȡҽ���������һ�ε�˵��
    '���:
    '   strҽ��IDs=ҽ��id,����ö��ŷָ�
    '   byt��������=��������:1-�¿���2-У�����ʣ�3-У��ͨ����4-���ϣ�5-������
    '                       6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ����
    '                       11-���ͨ����12-���δͨ����13-ʵϰҽʦͣ�������ˣ�14-Ѫ����գ�15-Ѫ�����ͨ����
    '                       16-Ѫ����Ѫ�ܾ���17-Ѫ��ֹͣ��Ѫ��18-��Ѫ����ͨ����ǩ����9-��Ѫ������ˣ�20-��Ѫҽ�����δ��
    '����:
    '   rsAdviceOper=ҽ��������Ϣ��ҽ��ID,����ʱ��,����˵��
    '����:ҽ���������һ�ε�˵��
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If strҽ��IDs = "" Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceOperInfo(strҽ��IDs, byt��������, rsAdviceOper, True, lngModule) = False Then Exit Function
    ZlGetMultiAdviceOperLastNotes = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlUpdatePatiAuditInfo(ByVal lng����ID As Long, _
    ByVal lng��ҳID As Long, ByVal byt��˱�� As Byte, ByVal blnȡ����� As Boolean, _
    Optional ByVal str����� As String, Optional ByVal str���˵�� As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���²��������Ϣ
    '���:
    '   byt��˱��=��˱�ǣ�0���-δ���,1-����˻�ʼ���;2-������
    '   blnȡ�����=�Ƿ�ȡ����ˣ�1-ȡ�����,0-���
    '����:
    '����:ִ�гɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_UpdatePatiAuditInfo(lng����ID, lng��ҳID, _
        byt��˱��, blnȡ�����, str�����, str���˵��, lngModule) = False Then Exit Function
    ZlUpdatePatiAuditInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlUpdateInpatientExtendInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal str������Ϣ As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ĳ�����ҳ�ӱ������Ϣ
    '���:
    '   str������Ϣ=��ʽ����Ϣ��:��Ϣֵ,��Ϣ��:��Ϣֵ,...
    '����:
    '����:ִ�гɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_UpdateInpatientExtendInfo(lng����ID, lng��ҳID, str������Ϣ, lngModule) = False Then Exit Function
    ZlUpdateInpatientExtendInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetGroupAdviceInfo(ByVal strҽ��IDs As String, ByRef rsAdvice As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ����ҩ��ҽ������
    '���:
    '   strҽ��IDs=ҽ��id,����ö��ŷָ�
    '����:
    '   rsAdvice=ҽ��������Ϣ��ҽ��ID,ҽ������
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetGroupAdviceInfo(strҽ��IDs, rsAdvice, lngModule) = False Then Exit Function
    ZlGetGroupAdviceInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlUpdateAdviceExeStatus(ByVal strNO As String, ByVal str���s As String, _
    ByVal byt��¼���� As Byte, ByVal byt������Դ As Byte, ByVal blnCancelExe As Boolean, _
    Optional ByVal strִ���� As String, Optional ByVal strִ��ʱ�� As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ�����͵�ִ��״̬
    '���:
    '   strNo�����õ��ݺ�
    '   str���s���������
    '   byt��¼����-1-�շ�;2-����;3-�Զ�����
    '   byt������Դ-������Դ��1-���2-סԺ
    '   blnCancelExe-�Ƿ���ȡ��ִ��
    '����:
    '����:ִ�гɹ�����True�����򷵻�False
    '˵��:
    '   ����ҽ����ִ��״̬,���ͬһ��ҽ�����͵���ִ������,�Ÿ���ִ��״̬Ϊ��ִ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim cllAdviceDatas As Collection, cllItem As Collection
    '   cllAdviceDatas(collect)-���ݼ�����ʽ����
    '     |-cllAdviceData(collect)ÿ����ϸ���ݼ�
    '        |-��Ա(ҽ��ID,���õ���,��������,ִ��״̬,ִ����,ִ��ʱ��,ԭִ��״̬)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    '����ҩƷ�͸������õ�����
    strSQL = _
        " Select 1 As ����,ҽ�����" & _
        " From סԺ���ü�¼ A" & _
        " Where a.No = [1] And a.��¼���� = [2] And a.��¼״̬ In (0, 1, 3) And a.ҽ����� Is Not Null" & _
        "       And (Instr(',' || [3] || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 Or [3] Is Null)" & _
        "       And Exists(Select 1 From סԺ���ü�¼" & _
        "                  Where NO = a.No And ��¼���� = a.��¼���� And ҽ����� = a.ҽ�����" & _
        "                        And ִ��״̬ = 0 And ��¼״̬ In (0, 1, 3))" & _
        "       And a.�շ���� Not In ('5', '6', '7')" & _
        "       And Not Exists(Select 1 From �������� B Where a.�շ�ϸĿid = b.����id And a.�շ���� = '4'  And Nvl(b.��������, 0) = 1)"

    strSQL = strSQL & " Union All" & _
        " Select 2 As ����,ҽ�����" & _
        " From סԺ���ü�¼ A" & _
        " Where a.No = [1] And a.��¼���� = [2] And a.��¼״̬ In (0, 1, 3) And a.ҽ����� Is Not Null" & _
        "       And (Instr(',' || [3] || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 Or [3] Is Null)" & _
        "       And Not Exists(Select 1 From סԺ���ü�¼" & _
        "                      Where NO = a.No And ��¼���� = a.��¼���� And ҽ����� = a.ҽ�����" & _
        "                            And ִ��״̬ = 0 And ��¼״̬ In (0, 1, 3))" & _
        "       And a.�շ���� Not In ('5', '6', '7')" & _
        "       And Not Exists(Select 1 From �������� B Where a.�շ�ϸĿid = b.����id And a.�շ���� = '4' And Nvl(b.��������, 0) = 1)"
    
    If byt������Դ <> 2 Then strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
    If blnCancelExe Then strSQL = Replace(strSQL, "And ִ��״̬ = 0", "And ִ��״̬ In(1,2)")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ������", strNO, byt��¼����, str���s)
    If rsTemp.EOF Then ZlUpdateAdviceExeStatus = True: Exit Function
    
    Set cllAdviceDatas = New Collection
    Do While Not rsTemp.EOF
        Set cllItem = New Collection
        cllItem.Add Val(Nvl(rsTemp!ҽ�����)), "ҽ��ID"
        cllItem.Add strNO, "���õ���"
        cllItem.Add byt��¼����, "��������"
        cllItem.Add strִ����, "ִ����"
        cllItem.Add strִ��ʱ��, "ִ��ʱ��"
        If blnCancelExe Then
            cllItem.Add IIf(Val(Nvl(rsTemp!����)) = 1, 3, 0), "ִ��״̬"
            cllItem.Add IIf(Val(Nvl(rsTemp!����)) = 1, "1", "1,3"), "ԭִ��״̬"
        Else
            cllItem.Add IIf(Val(Nvl(rsTemp!����)) = 1, 3, 1), "ִ��״̬"
            cllItem.Add IIf(Val(Nvl(rsTemp!����)) = 1, "0", "0,3"), "ԭִ��״̬"
        End If
        cllAdviceDatas.Add cllItem
        rsTemp.MoveNext
    Loop
    If objService.ZlCisSvr_UpdateAdviceExeStatus(cllAdviceDatas, lngModule) = False Then Exit Function
    ZlUpdateAdviceExeStatus = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
